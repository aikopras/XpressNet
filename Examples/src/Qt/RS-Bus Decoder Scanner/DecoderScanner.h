//
//  DecoderScanner.h
//  XpressNetQt examples: RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 23-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Qt version of DecoderScanner.m of the macOS version; the scan works the same.
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
//     An accessory decoder may use more than one RS-Bus address (see rsAddressCount in DecoderScanner.cpp). Its next
//     RS-Bus address(es) belong to the same decoder, and are not scanned.
//  Error correction during step 3:
//  - After a timeout, all answers are ignored for some time (QUIET_TIME), before the next PoM message is sent. A late
//    answer can therefore not be taken as the answer to the next PoM message.
//  - After CV7, and after each answer of a decoder with RSFEC (CV20) > 0, the repeated answers are ignored for some
//    time (FEC_GUARD_TIME).
//  - A value that is not possible (see isPlausible in DecoderScanner.cpp) is read again (MAX_REREADS).
//  Differences with the macOS version:
//  - The events of the XpressNet objects are Qt signals. They are connected to lambda functions, that add the
//    station (macOS: one delegate object per XpressNet object).
//  - One QTimer for all timeouts (macOS: NSTimer).
//

#ifndef DECODERSCANNER_H
#define DECODERSCANNER_H

#include <QObject>
#include <QList>
#include <QStringList>
#include <QTimer>

class XpressNetClass;

#define MAX_STATIONS 5                          // Maximum number of command stations
#define MAX_CV_NUMBER 1024                      // The highest CV number that can be read via PoM

// The decoder types, as derived from the TT bits of the RS-Bus feedback message
enum rsBusType_t
{ UNKNOWN_DECODER = 0,
  ACCESSORY_DECODER = 1,                        // TT=01: accessory decoder with feedback
  FEEDBACK_DECODER = 2,                         // TT=10: feedback decoder (GBM)
};


// ****************************************************************************************************************
// A decoder that was found on the RS-Bus, with the values of the CVs that have been read (-1 = not received)
class ScannedDecoder
{
public:
  ScannedDecoder();
  int rsStation = 0;                            // The command station (1..5) on whose RS-Bus the decoder was found
  int rsAddress = 0;                            // 1..127
  rsBusType_t rsBusType = UNKNOWN_DECODER;
  int pomStation = 0;                           // The command station (1..5) that delivered the PoM messages
  QString pomIpAddress;                         // The IP address of that command station
  bool skip = false;                            // This RS-Bus address belongs to the decoder at a lower RS-Bus address
  int givenPomAddress = 0;                      // PoM address given by the user (the lift decoder); 0 = none
  bool liftIr = false;                          // The row of the IR sensors of the lift (own RS-Bus addresses)
  ScannedDecoder *liftLevel = nullptr;          // For liftIr: the lift decoder, for the PoM address and CVs

  int pomAddress() const;
  void setCv(int cvNumber, int value);
  int getCv(int cvNumber) const;                // -1 if not received
  int numberOfSwitches() const;                 // The number of switches of an accessory decoder (-1 = unknown)
  int rsAddressCount() const;                   // The number of RS-Bus addresses of an accessory decoder
  // The texts for the table and the CSV file
  QString rsAddressText() const;                // For example "1-2"
  QString accessoryAddressText() const;         // For example "1-8"
  QString typeText() const;
  QString cmdStationText() const;               // CV19: "Lenz", "Roco" ...
  QString cvText(int cvNumber) const;
  static QString typeName(int value);           // The name of DecType (CV27); empty if unknown

private:
  QList<int> cvStore;
};


// ****************************************************************************************************************
// One command station: its XpressNet object
struct Station
{
  int number = 0;                               // 1..5, as in the preferences
  QString ipAddress;
  XpressNetClass *XpressNet = nullptr;
  bool connected = false;
  bool connectDone = false;                     // The connection has been established, or has failed
};


// ****************************************************************************************************************
class DecoderScanner : public QObject
{
  Q_OBJECT

public:
  explicit DecoderScanner(QObject *parent = nullptr);
  ~DecoderScanner() override;

  // ipAddresses: per command station; an empty string = not used.
  // cvNumbers: the CVs that should be read for the table, in addition to the CVs the scanner needs itself.
  // liftPomAddress: the PoM address of the lift decoder, whose PoM address does not follow from its RS-Bus address
  // (0 = no lift). liftIrAddress: the first of the two RS-Bus addresses of the IR sensors of the lift (0 = none)
  void start(const QStringList &ipAddresses, const QList<int> &cvNumbers, int liftPomAddress, int liftIrAddress);
  void stop();
  bool active() const { return phase != IDLE; }
  const QList<ScannedDecoder *> &decoders() const { return decoderList; }   // The decoders that answered

signals:
  void scannerStatus(QString text, double progress);           // progress: 0..100
  void scannerFoundDecoder(ScannedDecoder *decoder);            // CV7 received: the decoder is added to the table
  void scannerUpdatedDecoder(ScannedDecoder *decoder);          // Another CV received (or given up)
  void scannerFinished(QString text);

private:
  enum scanPhase_t { IDLE, CONNECTING, RSBUS_DATA, POM_READ };
  enum timerAction_t { CONNECT_TIMEOUT, RSBUS_TIMEOUT, POM_TIMEOUT, RESEND, NEXT_CV, NEXT_DECODER };

  scanPhase_t phase = IDLE;
  QList<Station *> stations;
  QList<Station *> retired;                     // The stations of the previous scan (deleted at the next scan)
  QList<ScannedDecoder *> owned;                // All ScannedDecoder objects of this scan (deleted at the next scan)
  QList<ScannedDecoder *> found;                // All decoders found on the RS-Bus
  QList<ScannedDecoder *> decoderList;          // The decoders that answered
  QStringList problems;                         // Texts for the final status line
  QStringList noAnswer;                         // RS-Bus addresses without PoM answer
  QTimer timer;                                 // The running timeout (one at a time)
  timerAction_t timerAction = CONNECT_TIMEOUT;
  int numberConnected = 0;
  // Step 2
  int rsStationIndex = 0;                       // Index in stations
  int rsStationsDone = 0;
  int partLastAddress = 0;                      // The last RS-Bus address of the part that has been requested
  int answerCount = 0;                          // The answers for partLastAddress
  // Step 3
  QList<int> userCvs;                           // The CVs chosen by the user
  int liftPomAddress = 0;                       // The lift decoder (0 = not included)
  int liftIrAddress = 0;                        // The first RS-Bus address of the IR sensors of the lift (0 = none)
  QList<ScannedDecoder *> liftRows;             // The rows of the lift decoder, added at the end of the table
  QList<int> cvList;                            // The CVs to read for the current decoder
  int decoderIndex = 0;                         // Index in found
  int cvIndex = 0;                              // Index in cvList
  QList<Station *> candidates;                  // The stations to try for CV7
  int candidateIndex = 0;
  int attempt = 0;
  int rereads = 0;                              // Extra reads because of a value that is not possible
  bool waitingForAnswer = false;
  bool haveLowNibble = false;
  int lowNibble = 0;
  int lastSuccessfulStation = 0;                // Number of the station that delivered the last PoM messages

  void startTimeout(double seconds, timerAction_t action);
  void timerFired();
  void finish(const QString &text);
  ScannedDecoder *newDecoder();
  Station *stationWithNumber(int number);
  // Events of the XpressNet objects
  void stationConnected(Station *station, bool ok, const QString &text);
  void stationAnswer(Station *station, int rsAddress, int nibble, int data, int ttBits);
  void stationSendError(Station *station, const QString &text);
  // Step 1
  void connectTimeout();
  // Step 2
  void startRsBusData();
  void nextRsStation();
  void requestFeedbacks(Station *station, int firstAddress, int lastAddress);
  void showRsBusProgress(int rsAddress);
  void rsBusTimeout();
  void rsBusAnswer(Station *station, int rsAddress, int ttBits);
  // Step 3
  void startPom();
  void nextDecoder();
  void buildCvList(ScannedDecoder *decoder);
  void sendPomRead();
  void pomAnswer(Station *station, int rsAddress, int nibble, int data);
  void nextCv();
  void addLiftIr(ScannedDecoder *lift);
  void pomTimeout();
};

#endif // DECODERSCANNER_H
