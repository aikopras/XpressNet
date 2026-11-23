//
//  UDPSocket.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  2026: Reusable UDP client "connection", based on Apple's Network framework. It is the UDP counterpart of
//  TCPSocketClass, and is used for the Z21 LAN protocol. The class has no knowledge of DCC, XpressNet or Z21;
//  it only sends and receives datagrams to / from a single remote host and port.
//
//  Usage:
//  - create the object, set its delegate and call openToHost:port:
//  - the delegate is informed of every state change and of every received datagram
//  - all delegate methods are called on the main thread
//  - call close to close the connection (the delegate receives UDP_SOCKET_CLOSED)
//  - data passed to sendBytes:length: while the connection is being prepared is kept, and sent as soon as the
//    connection is ready. In all other states than preparing or ready, data passed to sendBytes:length: is discarded
//  Note: UDP has no connection setup with the remote side. "Ready" only means that macOS can send datagrams
//  to the remote host; it does not mean that the remote host exists. The project should link Network.framework.
//

#import <Foundation/Foundation.h>

@class UDPSocketClass;

// The UDP socket may be in any of the following states
typedef enum
{ UDP_SOCKET_CLOSED = 0,      // No connection (initial state, or after close)
  UDP_SOCKET_PREPARING = 1,   // The connection is being prepared
  UDP_SOCKET_READY = 2,       // Datagrams can be sent and received
  UDP_SOCKET_FAILED = 3,      // The connection could not be prepared, or failed afterwards
} udpSocketState_t;


// The object that uses the socket should implement the following two methods
@protocol UDPSocketDelegate <NSObject>
// Called after every state change. For UDP_SOCKET_FAILED, reason contains a description of the error
- (void)udpSocket:(UDPSocketClass *)socket didChangeState:(udpSocketState_t)state reason:(NSString *)reason;
// Called for every datagram received. The bytes are only valid during this call
- (void)udpSocket:(UDPSocketClass *)socket didReceiveBytes:(const uint8_t *)bytes length:(NSUInteger)length;
@end


@interface UDPSocketClass : NSObject

// The following methods should be used
- (void)openToHost:(NSString *)host port:(NSString *)port;
- (void)sendBytes:(const uint8_t *)bytes length:(NSUInteger)length;
- (void)close;
- (BOOL)isOpen;                                        // YES while preparing or ready
- (int)localPort;                                      // 05-11-2026: the local port; 0 if not ready

@property (assign) id <UDPSocketDelegate> delegate;    // Not retained
@property (readonly) udpSocketState_t state;

@end
