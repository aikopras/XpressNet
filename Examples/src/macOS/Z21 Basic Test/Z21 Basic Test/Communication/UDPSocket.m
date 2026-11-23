//
//  UDPSocket.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  2026: Reusable UDP client "connection", based on Apple's Network framework. See UDPSocket.h.
//
//  Notes on the implementation (same approach as TCPSocketClass):
//  - All Network framework events are handled on the main queue.
//  - Manual memory management (MRC): the handlers do not retain self (__block variables). Instead, self
//    is retained when a connection is opened, and released when the final event of that connection
//    ("cancelled") has been delivered.
//

#import "UDPSocket.h"
#import <Network/Network.h>
#include <string.h>

@implementation UDPSocketClass {
  nw_connection_t _connection;           // The current connection; nil if there is no connection
  NSMutableData *_pendingData;           // Datagrams passed to sendBytes:length: while preparing; nil if none
  NSMutableArray *_pendingLengths;       // Length of each datagram in _pendingData
}

@synthesize delegate = _delegate;
@synthesize state = _state;


// ******************************************************************************************
// ************************************** PUBLIC METHODS ************************************
// ******************************************************************************************
- (void)openToHost:(NSString *)host port:(NSString *)port {
  // Only one connection at a time: close a previous connection first
  [self close];
  nw_endpoint_t endpoint = NULL;
  if ((host != nil) && (port != nil)) endpoint = nw_endpoint_create_host([host UTF8String], [port UTF8String]);
  if (endpoint == NULL) {[self setState:UDP_SOCKET_FAILED reason:@"invalid host or port"]; return;}
  nw_parameters_t parameters = nw_parameters_create_secure_udp(NW_PARAMETERS_DISABLE_PROTOCOL, NW_PARAMETERS_DEFAULT_CONFIGURATION);
  nw_connection_t connection = nw_connection_create(endpoint, parameters);
  nw_release(endpoint);
  nw_release(parameters);
  if (connection == NULL) {[self setState:UDP_SOCKET_FAILED reason:@"could not create connection"]; return;}
  _connection = connection;
  // Keep this object alive until the final event of this connection has been delivered (see top of file)
  [self retain];
  __block UDPSocketClass *blockSelf = self;
  __block nw_connection_t blockConnection = connection;
  nw_connection_set_queue(connection, dispatch_get_main_queue());
  nw_connection_set_state_changed_handler(connection, ^(nw_connection_state_t nwState, nw_error_t error) {
    [blockSelf connection:blockConnection changedState:nwState error:error];
  });
  nw_connection_start(connection);
  [self setState:UDP_SOCKET_PREPARING reason:nil];
}


- (void)sendBytes:(const uint8_t *)bytes length:(NSUInteger)length {
  if ((_connection == nil) || (length == 0)) return;
  // Datagrams offered while the connection is being prepared are kept, and sent as soon as the connection is ready
  if (_state == UDP_SOCKET_PREPARING) {
    if (_pendingData == nil) {
      _pendingData = [[NSMutableData alloc] init];
      _pendingLengths = [[NSMutableArray alloc] init];
    }
    [_pendingData appendBytes:bytes length:length];
    [_pendingLengths addObject:[NSNumber numberWithUnsignedInteger:length]];
    return;
  }
  if (_state != UDP_SOCKET_READY) return;
  // dispatch_data_create with the default destructor copies the bytes, so the caller may reuse its buffer
  dispatch_data_t content = dispatch_data_create(bytes, length, NULL, DISPATCH_DATA_DESTRUCTOR_DEFAULT);
  __block UDPSocketClass *blockSelf = self;
  __block nw_connection_t blockConnection = _connection;
  nw_connection_send(_connection, content, NW_CONNECTION_DEFAULT_MESSAGE_CONTEXT, true, ^(nw_error_t error) {
    if ((error != nil) && (blockConnection == blockSelf->_connection)) [blockSelf connectionFailed:blockConnection error:error];
  });
  dispatch_release(content);
}


- (void)close {
  if (_connection == nil) return;               // Nothing to close
  nw_connection_t connection = _connection;
  _connection = nil;                            // From now on, events of this connection are ignored
  nw_connection_cancel(connection);             // The last event will be "cancelled"
  [self discardPendingData];
  [self setState:UDP_SOCKET_CLOSED reason:nil];
}


- (BOOL)isOpen {
  return ((_state == UDP_SOCKET_PREPARING) || (_state == UDP_SOCKET_READY));
}


// 05-11-2026: the local port of the connection, as selected by macOS; 0 if not ready
- (int)localPort {
  if ((_connection == nil) || (_state != UDP_SOCKET_READY)) return 0;
  nw_path_t path = nw_connection_copy_current_path(_connection);
  if (path == NULL) return 0;
  nw_endpoint_t endpoint = nw_path_copy_effective_local_endpoint(path);
  nw_release(path);
  if (endpoint == NULL) return 0;
  int port = (int)nw_endpoint_get_port(endpoint);
  nw_release(endpoint);
  return port;
}


// ******************************************************************************************
// ************************************** INTERNAL METHODS **********************************
// ******************************************************************************************
- (void)setState:(udpSocketState_t)state reason:(NSString *)reason {
  _state = state;
  [_delegate udpSocket:self didChangeState:state reason:reason];
}


- (void)connection:(nw_connection_t)connection changedState:(nw_connection_state_t)nwState error:(nw_error_t)error {
  if (nwState == nw_connection_state_cancelled) {
    // Final event of this connection: release the connection and this object (see openToHost:port:)
    nw_release(connection);
    [self release];
    return;
  }
  if (connection != _connection) return;        // Event of a connection that has already been closed
  switch (nwState) {
    case nw_connection_state_ready:
      _state = UDP_SOCKET_READY;                // Set first, so the datagrams kept while preparing can be sent
      [self sendPendingData];
      [self setState:UDP_SOCKET_READY reason:nil];
      if (connection == _connection) [self receiveOn:connection];   // Unless the delegate closed the connection
      break;
    case nw_connection_state_waiting:           // No route to the host; the framework would retry later, we give up
    case nw_connection_state_failed:
      [self connectionFailed:connection error:error];
      break;
    default:                                    // invalid, preparing: nothing to do
      break;
  }
}


- (void)receiveOn:(nw_connection_t)connection {
  __block UDPSocketClass *blockSelf = self;
  __block nw_connection_t blockConnection = connection;
  nw_connection_receive_message(connection,
    ^(dispatch_data_t content, nw_content_context_t context, bool isComplete, nw_error_t error) {
      // Pass the received datagram to the delegate as one block of bytes
      if ((content != nil) && (blockConnection == blockSelf->_connection)) {
        NSUInteger size = dispatch_data_get_size(content);
        uint8_t *buffer = malloc(size);
        if (buffer != NULL) {
          __block NSUInteger offset = 0;
          dispatch_data_apply(content, ^bool(dispatch_data_t region, size_t regionOffset, const void *regionBuffer, size_t regionSize) {
            memcpy(buffer + offset, regionBuffer, regionSize);
            offset += regionSize;
            return true;
          });
          [blockSelf->_delegate udpSocket:blockSelf didReceiveBytes:buffer length:size];
          free(buffer);
        }
      }
      if (blockConnection != blockSelf->_connection) return;    // Connection closed in the meantime
      if (error != nil) {[blockSelf connectionFailed:blockConnection error:error]; return;}
      [blockSelf receiveOn:blockConnection];                     // Wait for the next datagram
    });
}


// The connection could not be prepared, or failed: cancel it and inform the delegate
- (void)connectionFailed:(nw_connection_t)connection error:(nw_error_t)error {
  NSString *reason = @"unknown reason";
  if (error != nil) {
    int code = nw_error_get_error_code(error);
    if (nw_error_get_error_domain(error) == nw_error_domain_posix) reason = [NSString stringWithUTF8String:strerror(code)];
    else reason = [NSString stringWithFormat:@"network error %d", code];
  }
  _connection = nil;
  nw_connection_cancel(connection);
  [self discardPendingData];
  [self setState:UDP_SOCKET_FAILED reason:reason];
}


// Send the datagrams that were kept while the connection was being prepared (see sendBytes:length:)
- (void)sendPendingData {
  if (_pendingData == nil) return;
  NSMutableData *data = _pendingData;
  NSMutableArray *lengths = _pendingLengths;
  _pendingData = nil;
  _pendingLengths = nil;
  const uint8_t *bytes = [data bytes];
  NSUInteger offset = 0;
  for (NSNumber *length in lengths) {
    [self sendBytes:(bytes + offset) length:[length unsignedIntegerValue]];
    offset += [length unsignedIntegerValue];
  }
  [data release];
  [lengths release];
}


- (void)discardPendingData {
  [_pendingData release];
  _pendingData = nil;
  [_pendingLengths release];
  _pendingLengths = nil;
}


- (void)dealloc {
  [_pendingData release];
  [_pendingLengths release];
  [super dealloc];
}

@end
