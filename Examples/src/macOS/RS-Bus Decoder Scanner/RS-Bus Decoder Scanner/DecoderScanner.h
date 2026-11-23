//
//  DecoderScanner.h
//  RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 22-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Scans up to 5 command stations (Lenz LAN/USB interfaces) for decoders with RS-Bus feedback, and reads some CVs of
//  each decoder via PoM. All command stations are equal: each may have feedback, loco and accessory decoders.
//  The scan has three steps:
//  1) Connect: a XpressNet object is created for every IP address, and the connection is opened.
//  2) RS-Bus data: per command station, the feedback information of all 128 RS-Bus addresses is requested (as Get
//     RS-Bus data of the RS-Bus Monitor, but in parts of 32 addresses). The TT bits of the answers give the decoder
//     type: TT=01: accessory decoder with feedback, TT=10: feedback decoder (GBM). TT=00 and TT=11 are skipped.
//  3) PoM: for every decoder found, CV7 (version) is read via PoM. The PoM address is 6999 + RS-Bus address for an
//     accessory decoder, and 6000 + RS-Bus address for a GBM. The PoM message is sent via each command station in
//     turn (the last successful one first), since we do not know which command station the decoder listens to.
//     The decoder answers with two RS-Bus feedback messages (nibbles) via RS-Bus address 128, or (old GBM firmware)
//     via its own RS-Bus address. These answers are only accepted from the command station on whose RS-Bus the
//     decoder was found. After CV7, the other CVs are read via the command station that delivered the PoM message:
//     first the CVs needed for the table (RSFEC, DecType, SkipUnEven, SwitchCount, SwitchType, myAddrL, myAddrH),
//     then the CVs chosen in the preferences. Decoders that do not answer are omitted.
//     The lift decoder (PoM address given by the user) is read first. Its RS-Bus address comes from CV10. Its rows
//     (Level, and IR for the RS-Bus addresses of the IR sensors) are shown at the end of the table.
//     An accessory decoder may use more than one RS-Bus address (see rsAddressCount in DecoderScanner.m). Its next
//     RS-Bus address(es) belong to the same decoder, and are not scanned.
//  Error correction during step 3:
//  - After a timeout, all answers are ignored for some time (QUIET_TIME), before the next PoM message is sent. A late
//    answer can therefore not be taken as the answer to the next PoM message.
//  - After CV7, and after each answer of a decoder with RSFEC (CV20) > 0, the repeated answers are ignored for some
//    time (FEC_GUARD_TIME).
//  - A value that is not possible (see isPlausible in DecoderScanner.m) is read again (MAX_REREADS).
//

#import <Foundation/Foundation.h>
#import "XpressNet.h"

#define MAX_STATIONS 5                          // Maximum number of command stations
#define MAX_CV_NUMBER 1024                      // The highest CV number that can be read via PoM

// The decoder types, as derived from the TT bits of the RS-Bus feedback message
typedef enum
{ ACCESSORY_DECODER = 1,                        // TT=01: accessory decoder with feedback
  FEEDBACK_DECODER = 2,                         // TT=10: feedback decoder (GBM)
} rsBusType_t;


// ****************************************************************************************************************
// A decoder that was found on the RS-Bus, with the values of the CVs that have been read (-1 = not received)
@interface ScannedDecoderClass : NSObject
@property (assign) int rsStation;               // The command station (1..5) on whose RS-Bus the decoder was found
@property (assign) int rsAddress;               // 1..127
@property (assign) rsBusType_t rsBusType;
@property (assign) int pomStation;              // The command station (1..5) that delivered the PoM messages
@property (copy)   NSString *pomIpAddress;      // The IP address of that command station
@property (assign) BOOL skip;                   // This RS-Bus address belongs to the decoder at a lower RS-Bus address
@property (assign) int givenPomAddress;         // PoM address given by the user (the lift decoder); 0 = none
@property (assign) BOOL liftIr;                 // The row of the IR sensors of the lift (own RS-Bus addresses)
@property (assign) ScannedDecoderClass *liftLevel;  // For liftIr: the lift decoder, for the PoM address and CVs
- (int)pomAddress;
- (void)setCv:(int)cvNumber value:(int)value;
- (int)getCv:(int)cvNumber;                     // -1 if not received
- (int)numberOfSwitches;                        // The number of switches of an accessory decoder (-1 = unknown)
- (int)rsAddressCount;                          // The number of RS-Bus addresses of an accessory decoder
// The texts for the table and the CSV file
- (NSString *)rsAddressText;                    // For example "1-2"
- (NSString *)accessoryAddressText;             // For example "1-8"
- (NSString *)typeText;
- (NSString *)cmdStationText;                   // CV19: "Lenz", "Roco" ...
- (NSString *)cvText:(int)cvNumber;
@end


// ****************************************************************************************************************
@class DecoderScannerClass;

// One command station: the XpressNet object and its delegate. The events of the XpressNet library do not contain
// the object that raised them, therefore each XpressNet object has its own delegate, that adds the station number
@interface StationClass : NSObject <XpressNetDelegate>
@property (assign) int number;                  // 1..5, as in the preferences
@property (copy)   NSString *ipAddress;
@property (retain) XpressNetClass *XpressNet;
@property (assign) BOOL connected;
@property (assign) BOOL connectDone;            // The connection has been established, or has failed
@property (assign) DecoderScannerClass *scanner;
@end


// ****************************************************************************************************************
// The object that receives the results of the scan (the AppDelegate)
@protocol DecoderScannerDelegate <NSObject>
- (void)scannerStatus:(NSString *)text progress:(double)progress;   // progress: 0..100
- (void)scannerFoundDecoder:(ScannedDecoderClass *)decoder;         // CV7 received: the decoder is added to the table
- (void)scannerUpdatedDecoder:(ScannedDecoderClass *)decoder;       // Another CV received (or given up)
- (void)scannerFinished:(NSString *)text;
@end


@interface DecoderScannerClass : NSObject
@property (assign) id <DecoderScannerDelegate> delegate;
@property (readonly) NSMutableArray *decoders;  // The decoders that answered (ScannedDecoderClass)
@property (readonly) BOOL active;

// ipAddresses: NSString per command station; an empty string = not used.
// cvNumbers: NSNumber per CV that should be read for the table, in addition to the CVs the scanner needs itself.
// pomAddress: the PoM address of the lift decoder, whose PoM address does not follow from its RS-Bus address
// (0 = no lift). irAddress: the first of the two RS-Bus addresses of the IR sensors of the lift (0 = none)
- (void)startWithIpAddresses:(NSArray *)ipAddresses cvNumbers:(NSArray *)cvNumbers
                liftPomAddress:(int)pomAddress liftIrAddress:(int)irAddress;
- (void)stop;

// Called by StationClass
- (void)station:(StationClass *)station connected:(BOOL)ok text:(NSString *)text;
- (void)station:(StationClass *)station rsAddress:(int)rsAddress nibble:(int)nibble data:(int)data ttBits:(int)ttBits;
- (void)station:(StationClass *)station sendError:(NSString *)text;
@end
