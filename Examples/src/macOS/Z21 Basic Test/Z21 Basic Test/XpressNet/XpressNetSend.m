//
//  XpressNetSend.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the SendClass of the VB6 XpressNet DLL: implementation of the message and nextMessage methods
//
//  message is called by the main XpressNet code to send (prepare the sending of) an XpressNet message.
//  The parameters to the message method are the number of bytes, followed by the individual bytes of the
//  XpressNet message, separated by comma's. These bytes should NOT include the two Framing bytes at the
//  beginning (FF and FE), nor the XOR byte at the end.
//
//  nextMessage is called from the ReceiveClass as soon as a previous message has been acknowledged by the
//  command station. nextMessage doesn't need any parameters, but instead takes (pop) the oldest XpressNet
//  message that is still waiting from the FIFO queue.
//
//  A FIFO queue is needed, since XpressNet implements flow control between PC and the Command Station.
//  The Command Station acknowledges the reception of every individual message to the PC.
//  The PC can send the next message only after the acknowledgement of the previous message is received,
//  or a time-out (of 5 seconds during normal operation, or 90 seconds during programming) has occured.
//  The FIFO allows the main program to "send" multiple XpressNet messages at once, and thus relieves
//  the main program from any concerns regarding flow control.
//
//  Version 2.0 (November 2026). The functionality below is the same in the VB6, .NET and macOS versions:
//  - Retransmission. If the interface answers a message with an error that may disappear if the message is sent
//    again, the message is sent again: at most 3 times, after 100, 200 and 400 msec. These answers are:
//    61 80 (Ubertragungsfehler), 61 81 (Zentrale Busy), 01 01, 01 02, 01 03, 01 06 and 01 0A (V3.6: 1.5 and 3.1.5;
//    for 01 0A the specification advises to repeat the last command). If the last retransmission fails as well, the
//    message is given up: logfile and event SendError, and the next message from the FIFO is sent.
//    The answers 01 09 (Fehler in den Befehlsparametern) and 61 82 (Befehl in Zentrale nicht vorhanden) are not
//    solved by a retransmission: logfile and event SendError, and the next message is sent.
//  - Pause. After 01 05 (Zentrale adressiert das Interface nicht mehr) or 01 08 (Derzeit konnen keine Befehle an
//    die Zentrale gesendet werden) the interface can not pass messages to the command station. Sending is paused
//    (event InterfaceNichtAdressiert); the FIFO is kept. With 01 08 the last message did not reach the command
//    station: it is held, and sent first after the pause. After 01 07 (Zentrale adressiert das Interface wieder)
//    sending continues (event InterfaceWiederAdressiert). If 01 07 does not come within 5 sec (90 sec in
//    programming mode), both queues are emptied (event SendError) and the pause ends.
//    During the pause, messages for the interface itself (header F0, F1, F2; for example the TCP keep alive
//    message) are still sent: these are answered by the interface, and not by the command station.
//  - Nothing is dropped silently. A message that is not sent gives a line in the logfile and the event SendError.
//  - Retransmission and pause are not used for the Z21 (UDP): without flow control it is not known to which message
//    an answer belongs.
//
//  Differences with the VB6 version:
//  - The response timeout is used for TCP, USB and serial connections (VB6: TCP only)
//  - The Z21 (UDP) has no flow control: the Z21 does not answer every command (Z21 spec, section 1.1),
//    therefore Z21 messages are sent immediately
//  - Header bytes for XpressNet / Z21 are added here, in transmit (VB6: in the classes Tcp, Udp and Usb)
//  - 2.0: a byte value outside 0..255 is not sent, but written to the logfile (event SendError)
//

#import "XpressNetSend.h"
#import "XpressNet.h"
#import "XpressNetLogFile.h"
#import "XpressNetAccAuto.h"
#import "TCPSocket.h"
#import "UDPSocket.h"
#import "SerialPort.h"
#include <stdarg.h>

@implementation SendClass {
  NSMutableArray *_fifo;                  // The FIFO queue: each element is an NSData with one XpressNet message (incl. XOR)
  NSTimer *_lanKeepAliveTimer;            // Sends a keep alive message, if nothing else has been sent for some time
  // 2.0: Retransmission (see the top of this file)
  NSData *_lastMessage;                   // The last transmitted message (including the XOR byte); nil if none
  int _retryCount;                        // The number of retransmissions of _lastMessage
  BOOL _retryPending;                     // A retransmission of _lastMessage waits for the retry timer
  // 2.0: Pause (see the top of this file)
  NSData *_heldMessage;                   // The message that is sent first after the pause (01 08); nil if none
  NSTimeInterval _responseTimeOut;        // The timeout of the running response timer, for the event SendError
  NSTimeInterval _pauseTimeOut;           // The timeout of the running pause timer, for the event SendError
}

@synthesize waitForAck                  = _waitForAck;
@synthesize waitForlanKeepAliveResponse = _waitForlanKeepAliveResponse;
@synthesize interfaceNotAddressed       = _interfaceNotAddressed;
@synthesize Parent                      = _Parent;

// ***************************************************************************************************************************************************
// ******************************************************* Local C type definitions and declarations *************************************************
// ***************************************************************************************************************************************************
#define MAX_FIFO_ENTRIES 128              // The number of XpressNet messages the queue is able to hold
#define MAX_MESSAGE_SIZE 15               // Maximum number of bytes of an XpressNet message from PC to Command Station (without XOR)
#define LAN_KEEPALIVE_INTERVAL 30.0       // Time (sec) between LAN Keep Alive messages
#define TCP_RESPONSE_TIMEOUT 5.0          // A response message should be received within 5 sec
#define TCP_PROGRAMMING_TIMEOUT 90.0      // In programming mode a response may take 90 sec
#define MAX_RETRIES 3                     // 2.0: at most 3 retransmissions
#define FIRST_RETRY_DELAY 0.1             // 2.0: the delays are 100, 200 and 400 msec
#define PAUSE_TIMEOUT 5.0                 // 2.0: 01 07 should be received within 5 sec
#define PAUSE_PROGRAMMING_TIMEOUT 90.0    // 2.0: in programming mode within 90 sec


- (id)init {
  self = [super init];
  if (self) {
    _waitForAck = NO;                     // The start value is that sending data is allowed
    _waitForlanKeepAliveResponse = NO;    // We did not yet send a LAN Keep Alive message
    _fifo = [[NSMutableArray alloc] init];
    [self restartKeepAliveTimer];
  }
  return self;
}


- (void)dealloc {
  [_fifo release];
  [_lastMessage release];
  [_heldMessage release];
  [super dealloc];
}


// Must be called before the object is released, since the timers retain this object.
// 2.0: also used if a connection is closed: the retransmission and the pause end
- (void)stop {
  [_lanKeepAliveTimer invalidate];
  _lanKeepAliveTimer = nil;
  [self stoptcpResponseTimer];
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(retryTimerFired) object:nil];
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(pauseTimerFired) object:nil];
  _retryPending = NO;
  _retryCount = 0;
  _interfaceNotAddressed = NO;
  [self setLastMessage:nil];
  [self setHeldMessage:nil];
}


// ----------------------------------------------------------------------------------------------------
// message is called from the XpressNetClass code. It either sends a message, or it stores the message
// in the FIFO. If a message is sent, it calls transmit, which in turn activates the flow control
// by setting the waitForAck Flag to YES.
// 2.0: a message that is not sent gives the event SendError (besides the line in the logfile). The checks of
// the message are done first, so the description of the event can include the XOR byte
- (void)message:(int)numberOfBytes, ... {
  // STEP 1: Perform 3 checks on the size and values of the received message.
  // Check 1: is the number of bytes within range?
  if ((numberOfBytes < 1) || (numberOfBytes > MAX_MESSAGE_SIZE)) {
    [self notSent:FehlerInDenBefehlsparametern text:@"Error in the message method: number of bytes out of range" message:nil];
    return;
  }
  int values[MAX_MESSAGE_SIZE];
  va_list arguments;
  va_start(arguments, numberOfBytes);
  for (int i = 0; i < numberOfBytes; i++) values[i] = va_arg(arguments, int);
  va_end(arguments);
  uint8_t buffer[MAX_MESSAGE_SIZE + 1] = {0};
  for (int i = 0; i < numberOfBytes; i++) buffer[i] = (uint8_t)values[i];
  // Check 2: Does the received number of bytes correspond to the XpressNet message header?
  // The number of bytes must match with the length as contained in (the lower 4 bits) of the first (=header) byte.
  // Note: the Z21 Specification sometimes violates this, and must be excluded from this check
  if ((![_Parent Z21Protocol]) && ((numberOfBytes - 1) != (buffer[0] & 0x0F))) {
    [self notSent:FehlerInDenBefehlsparametern text:@"Error in the message method: size doesn't match size in header"
          message:[NSData dataWithBytes:buffer length:(NSUInteger)numberOfBytes]];
    return;
  }
  // Check 3 (2.0): every byte must have a value between 0 and 255
  for (int i = 0; i < numberOfBytes; i++) {
    if ((values[i] < 0) || (values[i] > 255)) {
      [self notSent:FehlerInDenBefehlsparametern
               text:[NSString stringWithFormat:@"Error in the message method: value %d does not fit in a byte", values[i]]
            message:[NSData dataWithBytes:buffer length:(NSUInteger)numberOfBytes]];
      return;
    }
  }
  // STEP 2: Calculate the XOR value
  uint8_t xorValue = 0;
  for (int i = 0; i < numberOfBytes; i++) xorValue ^= buffer[i];
  buffer[numberOfBytes] = xorValue;
  NSData *xpnMessage = [NSData dataWithBytes:buffer length:(NSUInteger)(numberOfBytes + 1)];
  // STEP 3: Check if we have a connection. If not: drop the packet
  // 2.0: with a line in the logfile and the event SendError (1.2: silently)
  if (![_Parent isConnected]) {
    [self notSent:KeineVerbindung text:@"Message not sent: there is no connection" message:xpnMessage];
    return;
  }
  // STEP 4: Check if we can send immediately, or if the message should be stored in the FIFO.
  // The Z21 has no flow control.
  // 2.0: during a pause (01 05 / 01 08) only messages for the interface itself are sent (header F0, F1, F2)
  BOOL interfaceCommand = (buffer[0] >= 0xF0);
  if (([_Parent connectionType] == lanUdpPort) || ((!_waitForAck) && ((!_interfaceNotAddressed) || interfaceCommand))) {
    [self transmit:xpnMessage retransmission:NO];
  }
  else {
    // Check if the FIFO queue has free slots left to store the message
    // If not, increase MAX_FIFO_ENTRIES, or modify the calling program
    if ([_fifo count] >= MAX_FIFO_ENTRIES) {               // FIFO overflow :-(
      [self notSent:FifoVoll text:@"Error: FIFO queue is full" message:xpnMessage];
      return;
    }
    [_fifo addObject:xpnMessage];
  }
}


// ----------------------------------------------------------------------------------------------------
// nextMessage will be called from the ReceiveClass code. Provided that we are allowed to send now,
// it sends a message from the FIFO queue, if there are still messages waiting.
// 2.0: not during a pause. After a pause the held message (01 08) is sent first
- (void)nextMessage {
  if (_waitForAck || _interfaceNotAddressed) return;     // flow control, or pause
  if (_heldMessage != nil) {
    NSData *held = [_heldMessage retain];
    [self setHeldMessage:nil];
    [self transmit:held retransmission:YES];              // the held message keeps its number of retransmissions
    [held release];
    return;
  }
  if ([_fifo count] == 0) return;
  NSData *xpnMessage = [[_fifo objectAtIndex:0] retain];
  [_fifo removeObjectAtIndex:0];
  [self transmit:xpnMessage retransmission:NO];
  [xpnMessage release];
}


// transmit adds the header bytes that are needed for the connection type, and sends the message
// 2.0: retransmission is YES if xpnMessage is sent again (the number of retransmissions is kept)
- (void)transmit:(NSData *)xpnMessage retransmission:(BOOL)retransmission {
  connectionTypeEnum type = [_Parent connectionType];
  int length = (int)[xpnMessage length];
  uint8_t outBuffer[MAX_MESSAGE_SIZE + 6];
  int outLength;
  // 2.0: keep the message, for a retransmission
  if (!retransmission) _retryCount = 0;
  [self setLastMessage:xpnMessage];
  if (type != lanUdpPort) _waitForAck = YES;              // We may not send a next message immediately
  switch (type) {
    case lanTcpPort:
    case usbPort:
      // Add two Framing Bytes at the start of the message
      outBuffer[0] = 0xFF;
      outBuffer[1] = 0xFE;
      memcpy(&outBuffer[2], [xpnMessage bytes], (size_t)length);
      outLength = length + 2;
      break;
    case lanUdpPort:
      // Add the four bytes for the Z21 XpressNet tunnel at the start of the message: DataLen (2 bytes), Header 0x40 (2 bytes)
      outBuffer[0] = (uint8_t)(length + 4);
      outBuffer[1] = 0x00;
      outBuffer[2] = 0x40;
      outBuffer[3] = 0x00;
      memcpy(&outBuffer[4], [xpnMessage bytes], (size_t)length);
      outLength = length + 4;
      break;
    case serialPort:
      // Serial interfaces (LI100, LI101) use no Framing Bytes
      memcpy(outBuffer, [xpnMessage bytes], (size_t)length);
      outLength = length;
      break;
    default: return;
  }
  // Do we still have the TCP connection?
  if ((type == lanTcpPort) && (![[_Parent tcp] isOpen])) {
    [_Parent eventTcpClosedByRemoteHost];                 // Clean up, and inform the application
    return;
  }
  if ([_Parent LogHexMessages]) [[_Parent Logfile] SendHex:outBuffer length:outLength];
  switch (type) {
    case lanTcpPort: [[_Parent tcp] sendBytes:outBuffer length:(NSUInteger)outLength]; break;
    case lanUdpPort: [[_Parent udp] sendBytes:outBuffer length:(NSUInteger)outLength]; break;
    default:         [[_Parent usb] sendBytes:outBuffer length:(NSUInteger)outLength]; break;   // USB and serial use the same object
  }
  if ((type == lanTcpPort) || (type == lanUdpPort)) [self restartKeepAliveTimer];
  if (type != lanUdpPort) [self starttcpResponseTimer];  // This timer triggers if there is no response = Error
  [_Parent setNumberOfExpressNetMessagesTransmitted:[_Parent NumberOfExpressNetMessagesTransmitted] + 1];
}


// 2.0: sends a Z21 LAN message that is not an XpressNet message (no tunnel header 0x40, no XOR), such as
// LAN_SET_BROADCASTFLAGS and LAN_LOGOFF. z21Message contains the complete message, including DataLen and Header.
// As for the XpressNet messages to the Z21: no flow control and no FIFO. The message is not counted in
// NumberOfExpressNetMessagesTransmitted, since it is not an XpressNet message
- (void)transmitZ21:(const uint8_t *)z21Message length:(int)length {
  if (([_Parent connectionType] != lanUdpPort) || (![[_Parent udp] isOpen])) {
    [self notSent:KeineVerbindung text:@"Message not sent: there is no Z21 (UDP) connection"
          message:[NSData dataWithBytes:z21Message length:(NSUInteger)length]];
    return;
  }
  if ([_Parent LogHexMessages]) [[_Parent Logfile] SendHex:z21Message length:length];
  [[_Parent udp] sendBytes:z21Message length:(NSUInteger)length];
  [self restartKeepAliveTimer];
}


// ----------------------------------------------------------------------------------------------------
// LAN Keep Alive (V3.6: 2.1 - Interface Status Befehl). For the LAN interface, send a keep alive message
// to inform the interface we are still listening. This is required for the TCP (XpressNet) as well as
// the UDP (Z21) connection. The timer is restarted after every transmitted message.
// Note: NSTimer retains this object until the timer is invalidated (see stop).
- (void)restartKeepAliveTimer {
  [_lanKeepAliveTimer invalidate];
  _lanKeepAliveTimer = [NSTimer scheduledTimerWithTimeInterval:LAN_KEEPALIVE_INTERVAL target:self
                                                      selector:@selector(lanKeepAliveTimerFired:) userInfo:nil repeats:YES];
}


- (void)lanKeepAliveTimerFired:(NSTimer *)timer {
  // 2.0: not without a connection (the message would give the event SendError)
  if (![_Parent isConnected]) return;
  if ([_Parent connectionType] == lanTcpPort) {
    _waitForlanKeepAliveResponse = YES;
    [_Parent CmdInterfaceStatus];                         // This short message is used as XpressNet "Keep Alive"
    [[_Parent Logfile] DetailsSent:@"TCP Keep Alive Request"];
  }
  if ([_Parent connectionType] == lanUdpPort) {
    _waitForlanKeepAliveResponse = YES;
    [_Parent CmdStatusZentraleAnfordern];                 // This short message seems usable as Z21 "Keep Alive"
    [[_Parent Logfile] DetailsSent:@"UDP Keep Alive Request"];
  }
}


// ----------------------------------------------------------------------------------------------------
// Response timer. Note: performSelector:afterDelay: retains this object until it is cancelled or has fired.
- (void)starttcpResponseTimer {
  [self stoptcpResponseTimer];
  // Responses in programming mode may take 90 sec. In normal mode the response must be within 5 sec
  _responseTimeOut = [_Parent ZentraleProgrammierMode] ? TCP_PROGRAMMING_TIMEOUT : TCP_RESPONSE_TIMEOUT;
  [self performSelector:@selector(tcpResponseTimerFired) withObject:nil afterDelay:_responseTimeOut];
}


// Will be called from the receive class as soon as a response is received
- (void)stoptcpResponseTimer {
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(tcpResponseTimerFired) object:nil];
}


- (void)tcpResponseTimerFired {
  // If the timer expires, something went wrong. Let's go to a well-defined state
  // 2.0: the description for the event SendError (before _lastMessage changes)
  NSString *description = [NSString stringWithFormat:@"No answer within %d sec (message: %@)", (int)_responseTimeOut, [self hex:_lastMessage]];
  int removed = (int)[_fifo count];
  if (_heldMessage != nil) removed = removed + 1;
  [self emptyFifo];                                       // Remove all XpressNet messages still waiting
  [self setHeldMessage:nil];
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(retryTimerFired) object:nil];
  _retryPending = NO;                                     // 2.0: no retransmission is waiting
  _retryCount = 0;
  _waitForAck = NO;                                       // Clear the message acknowledgement flag
  // Remove all switch commands still waiting. 2.0: the activated switch is switched off first. _waitForAck is
  // cleared before, so this message is sent immediately
  int removedAccessories = [[_Parent AccAuto] emptyFifo];
  [[_Parent Logfile] Errors:@"Error: TCP Response Timer has timed out. Message queues emptied"];
  // Keep alive messages exist for the LAN interface only (V3.6: chapter 2)
  if ([_Parent connectionType] == lanTcpPort) {
    _waitForlanKeepAliveResponse = YES;                   // We will send a keep alive message
    [_Parent CmdInterfaceStatus];
  }
  description = [description stringByAppendingFormat:@". Message queues emptied (%d messages, %d CmdAccessoryAuto commands)", removed, removedAccessories];
  [_Parent eventSendError:Timeout description:description];
}


- (int)size {
  return (int)[_fifo count];
}


- (void)emptyFifo {
  [_fifo removeAllObjects];
}


// ----------------------------------------------------------------------------------------------------
// 2.0: RETRANSMISSION (see the top of this file)
// Called from the ReceiveClass if the interface answers the last message with an error that may disappear if the
// message is sent again: 61 80, 61 81, 01 01, 01 02, 01 03, 01 06 and 01 0A. answer is the text for the logfile
- (void)answerError:(NSString *)answer {
  if ([_Parent connectionType] == lanUdpPort) return;    // Z21: no flow control (see the top of this file)
  if (_lastMessage == nil) return;                        // No message has been sent yet
  if (_retryCount < MAX_RETRIES) {
    NSTimeInterval delay = FIRST_RETRY_DELAY * (1 << _retryCount);   // 100, 200, 400 msec
    _retryCount = _retryCount + 1;
    _retryPending = YES;
    _waitForAck = YES;                                    // The FIFO waits for the retransmission
    [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(retryTimerFired) object:nil];
    [self performSelector:@selector(retryTimerFired) withObject:nil afterDelay:delay];
    [[_Parent Logfile] Warnings:[NSString stringWithFormat:@"Answer %@: retransmission %d of %d after %d msec (message: %@)",
                                 answer, _retryCount, MAX_RETRIES, (int)(delay * 1000 + 0.5), [self hex:_lastMessage]]];
  }
  else {
    _retryCount = 0;
    NSString *description = [NSString stringWithFormat:@"Message given up after %d retransmissions. Answer: %@ (message: %@)",
                             MAX_RETRIES, answer, [self hex:_lastMessage]];
    [[_Parent Logfile] Errors:description];
    [_Parent eventSendError:WiederholungErfolglos description:description];
    // _waitForAck is NO: the next message is sent by nextMessage (ReceiveClass DecodeXpressNetMessage, STEP 4)
  }
}


// Called from the ReceiveClass if the interface answers the last message with 01 09 or 61 82. A retransmission
// does not help: the message is given up, and the next message is sent. reason: FehlerInDenBefehlsparametern (01 09)
// or BefehlNichtVorhanden (61 82)
- (void)answerRejected:(NSString *)answer reason:(SendErrorReasonEnum)reason {
  if ([_Parent connectionType] == lanUdpPort) return;    // Z21: no flow control (see the top of this file)
  if (_lastMessage == nil) return;                        // No message has been sent yet
  _retryCount = 0;
  NSString *description = [NSString stringWithFormat:@"Message rejected. Answer: %@ (message: %@)", answer, [self hex:_lastMessage]];
  [[_Parent Logfile] Errors:description];
  [_Parent eventSendError:reason description:description];
}


- (void)retryTimerFired {
  if (!_retryPending) return;
  _retryPending = NO;
  if (_interfaceNotAddressed) {                           // A pause started in the meantime: send after the pause
    [self setHeldMessage:_lastMessage];
    _waitForAck = NO;                                     // No message is underway
    return;
  }
  NSData *message = [_lastMessage retain];                // transmit replaces _lastMessage
  [self transmit:message retransmission:YES];
  [message release];
}


// ----------------------------------------------------------------------------------------------------
// 2.0: PAUSE (see the top of this file)
// Called from the ReceiveClass after 01 05 (answer = NO) or 01 08 (answer = YES). With 01 08 the last
// message did not reach the command station: it is held, and sent first after the pause
- (void)pause:(BOOL)answer {
  if ([_Parent connectionType] == lanUdpPort) return;    // Z21: no flow control (see the top of this file)
  if (answer && (_lastMessage != nil)) [self setHeldMessage:_lastMessage];
  if (_interfaceNotAddressed) return;                     // The pause has started already
  _interfaceNotAddressed = YES;
  // In programming mode the pause may take 90 sec
  _pauseTimeOut = [_Parent ZentraleProgrammierMode] ? PAUSE_PROGRAMMING_TIMEOUT : PAUSE_TIMEOUT;
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(pauseTimerFired) object:nil];
  [self performSelector:@selector(pauseTimerFired) withObject:nil afterDelay:_pauseTimeOut];
  [[_Parent Logfile] Warnings:@"Sending paused, until the command station addresses the interface again"];
  [_Parent eventInterfaceNichtAdressiert];
}


// Called from the ReceiveClass after 01 07. The held message is sent first, then the FIFO and CmdAccessoryAuto.
// The event is raised last, so messages the application sends from the event handler come after these
- (void)resume {
  if ([_Parent connectionType] == lanUdpPort) return;    // Z21: no flow control (see the top of this file)
  if (!_interfaceNotAddressed) return;                    // 01 07 without a pause (for example after power on): ignored
  _interfaceNotAddressed = NO;
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(pauseTimerFired) object:nil];
  [[_Parent Logfile] Warnings:@"Sending resumed: the command station addresses the interface again"];
  [self nextMessage];                                     // First the held message, otherwise the FIFO
  [[_Parent AccAuto] resume];                             // CmdAccessoryAuto continues
  [_Parent eventInterfaceWiederAdressiert];
}


- (void)pauseTimerFired {
  if (!_interfaceNotAddressed) return;
  _interfaceNotAddressed = NO;                            // The pause ends
  int removed = (int)[_fifo count];
  if (_heldMessage != nil) removed = removed + 1;
  [self emptyFifo];                                       // Remove all XpressNet messages still waiting
  [self setHeldMessage:nil];
  _retryCount = 0;
  if (_retryPending) {                                    // No retransmission anymore
    [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(retryTimerFired) object:nil];
    _retryPending = NO;
    _waitForAck = NO;
  }
  // Remove all switch commands still waiting. The activated switch (if any) is switched off first
  int removedAccessories = [[_Parent AccAuto] emptyFifo];
  NSString *description = [NSString stringWithFormat:@"The command station did not address the interface again within %d sec. Message queues emptied (%d messages, %d CmdAccessoryAuto commands)",
                           (int)_pauseTimeOut, removed, removedAccessories];
  [[_Parent Logfile] Errors:description];
  [_Parent eventSendError:ZentraleAdressiertDasInterfaceNichtMehr description:description];
}


// ----------------------------------------------------------------------------------------------------
// 2.0: A message that is not sent: logfile and event SendError. The description contains the bytes of the message.
// The same log level as in version 1.2: a missing connection is a warning, the rest an error
- (void)notSent:(SendErrorReasonEnum)reason text:(NSString *)text message:(NSData *)message {
  if (reason == KeineVerbindung) [[_Parent Logfile] Warnings:text];
  else [[_Parent Logfile] Errors:text];
  [_Parent eventSendError:reason description:[NSString stringWithFormat:@"%@ (message: %@)", text, [self hex:message]]];
}


// The bytes of a message as text, for example "52 00 88 DA"
- (NSString *)hex:(NSData *)message {
  NSMutableString *s = [NSMutableString string];
  const uint8_t *bytes = (const uint8_t *)[message bytes];
  for (NSUInteger i = 0; i < [message length]; i++) {
    if (i > 0) [s appendString:@" "];
    [s appendFormat:@"%02X", bytes[i]];
  }
  return s;
}


- (void)setLastMessage:(NSData *)message {
  if (message == _lastMessage) return;
  [_lastMessage release];
  _lastMessage = [message retain];
}


- (void)setHeldMessage:(NSData *)message {
  if (message == _heldMessage) return;
  [_heldMessage release];
  _heldMessage = [message retain];
}

@end
