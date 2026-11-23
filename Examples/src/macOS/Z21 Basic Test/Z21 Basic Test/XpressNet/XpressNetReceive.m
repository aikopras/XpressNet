//
//  XpressNetReceive.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the ReceiveClass of the VB6 XpressNet DLL.
//  The various methods contain references to the V3.6 and V4 XpressNet specifications:
//  - Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151
//    Kommunikation auf den Schnittstellen, Doku Version 1.2 05/2019
//  - Dokumentation XpressNet Version 4.0, 02/2022
//
//  Every time bytes are received from a physical (LAN/USB/Serial) interface, addToIndata is called to store the
//  received bytes in the inData buffer. Next, the buffer is checked for complete XpressNet messages.
//  Differences with the VB6 version:
//  - If the input buffer is full, the complete XpressNet messages are processed first (VB6: the remaining bytes
//    were dropped and the buffer stayed full, so nothing could be received anymore)
//  - The messages are decoded from a copy that is at least 32 bytes long and filled with zeros, so that
//    a message that is shorter than expected can not cause reading outside the message
//
//  Version 2.0 (November 2026), the same in the VB6, .NET and macOS versions:
//  - The answers of the interface to a message (V3.6: 1.5 and 3.1.5) are passed to the SendClass, which sends the
//    message again (retransmission), gives up the message, or pauses sending (see XpressNetSend.m).
//  - 01 05 and 01 07 are always handled as broadcast (unrequested), also for a serial interface without framing bytes
//  - After a message with an XOR error, the next framing bytes (FF FE or FF FD) are searched (TCP and USB). Without
//    this, a lost byte could also disturb the next messages. Serial interfaces have no framing bytes.
//

#import "XpressNetReceive.h"
#import "XpressNet.h"
#import "XpressNetSend.h"
#import "XpressNetLogFile.h"

@implementation ReceiveClass {
  uint8_t _inData[1024];                  // Buffer to store received bytes
  int _indexNextInByte;                   // Index into the inData Array. Empty space to store subsequent byte
  BOOL _inDataEmptied;                    // Set by emptyIndata (the application may open a new connection)
}

@synthesize Parent = _Parent;

#define IN_DATA_SIZE 1024
#define MESSAGE_BUFFER_SIZE 32            // Larger than the largest XpressNet message


// ================================================================================================
// Called by the XpressNet object with the bytes received from the TCP, USB or serial connection
- (void)addToIndata:(const uint8_t *)inBytes length:(int)bytesReceived {
  _inDataEmptied = NO;
  for (int i = 0; i < bytesReceived; i++) {
    // Some command stations (like the YD7001) may generate quite some data.
    // If the buffer is full, first process the complete XpressNet messages, and then continue. An XpressNet message
    // is at most 19 bytes (2 framing bytes, header, 15 data bytes, XOR), so afterwards the buffer contains at most 18 bytes
    if (_indexNextInByte >= IN_DATA_SIZE) {
      [self checkInDataForXpressNetmessages];
      // If the application opened a new connection meanwhile, the remaining bytes belong to the old connection.
      // The second condition can not be true (see above); it only guarantees that we never write outside _inData
      if (_inDataEmptied || (_indexNextInByte >= IN_DATA_SIZE)) return;
    }
    _inData[_indexNextInByte] = inBytes[i];
    _indexNextInByte++;
  }
  // Now that new data is stored, check if the inData buffer contains one or more
  // complete XpressNet messages. If yes, process this / these message(s)
  [self checkInDataForXpressNetmessages];
}


- (void)emptyIndata {
  _indexNextInByte = 0;
  _inDataEmptied = YES;
}


- (void)checkInDataForXpressNetmessages {
  // Since we may have multiple XpressNet messages within the inData buffer, we have a loop that doesn't end.
  // We use return to leave the loop, whenever the inData buffer is empty or contains an incomplete XpressNet message
  while (YES) {
    // STEP 1: Determine where the XpressNet header is located
    int indexHeader = (_inData[0] == 0xFF) ? 2 : 0;
    // Exit, if the header byte has not yet been received
    if (_indexNextInByte <= indexHeader) return;
    // STEP 2: Retrieve the length of the XpressNet message from the last four header bits
    int xpnLength = _inData[indexHeader] & 0x0F;
    int indexXor = indexHeader + xpnLength + 1;          // length excludes the XOR
    // Exit, if we have not yet received the last byte of an XpressNet message
    if (_indexNextInByte <= indexXor) return;
    // STEP 3: We have a complete XpressNet packet. Should we log it?
    if ([_Parent LogHexMessages]) [[_Parent Logfile] ReceiveHexRange:_inData from:0 to:indexXor];
    // STEP 4: Check if the XOR value is correct. If we include the XOR byte in the XOR calculation, the result should become 0
    uint8_t xorValue = 0;
    for (int i = indexHeader; i <= indexXor; i++) xorValue ^= _inData[i];
    if (xorValue == 0) {
      // STEP 5: Is this a broadcast message?
      BOOL broadcast = ((_inData[0] == 0xFF) && (_inData[1] == 0xFD));
      // STEP 6: Copy the XpressNet message into the xpnMessage buffer
      uint8_t xpnMessage[MESSAGE_BUFFER_SIZE];
      memset(xpnMessage, 0, MESSAGE_BUFFER_SIZE);
      memcpy(xpnMessage, &_inData[indexHeader], (size_t)(xpnLength + 1));
      // STEP 7: Call the method that analyses the individual XpressNet message
      [self DecodeXpressNetMessage:(xpnLength + 1) message:xpnMessage broadcast:broadcast];
    }
    else {
      [[_Parent Logfile] Errors:@"XOR Error in received XpressNet message"];
      if ([_Parent LogErrors]) [[_Parent Logfile] ReceiveHexRange:_inData from:0 to:indexXor];
      // 2.0: TCP and USB: remove the bytes up to the next framing bytes, and continue with that message
      if (([_Parent connectionType] == lanTcpPort) || ([_Parent connectionType] == usbPort)) {
        [self removeUntilNextFrame];
        continue;
      }
    }
    // STEP 8: Remove the old message from the inData buffer, and shift the remainder to the begin.
    // The buffer may have been emptied in the meantime (a new connection was opened by the application)
    int remaining = _indexNextInByte - (indexXor + 1);
    if (remaining <= 0) {_indexNextInByte = 0; return;}
    memmove(_inData, &_inData[indexXor + 1], (size_t)remaining);
    _indexNextInByte = remaining;
  }
}


// 2.0: After an XOR error: remove the bytes before the next framing bytes (FF FE or FF FD). The search starts at
// the second byte, since the first byte belongs to the message with the XOR error. Without framing bytes the buffer
// is emptied, except a last byte FF (this may be the first framing byte of the next message)
- (void)removeUntilNextFrame {
  int next = _indexNextInByte;                            // By default: remove everything
  for (int j = 1; j < _indexNextInByte; j++) {
    if (_inData[j] == 0xFF) {
      if (j == (_indexNextInByte - 1)) {next = j; break;} // A last byte FF: keep it
      if ((_inData[j + 1] == 0xFE) || (_inData[j + 1] == 0xFD)) {next = j; break;}
    }
  }
  int remaining = _indexNextInByte - next;
  if (remaining > 0) memmove(_inData, &_inData[next], (size_t)remaining);
  _indexNextInByte = remaining;
}


// ================================================================================================
// This method is called to analyse an individual XpressNet message.
// The first byte should be the XpressNet header; framing bytes should already be removed.
// XOR has already been checked and the XOR byte is removed.
- (void)DecodeXpressNetMessage:(int)length message:(const uint8_t *)message broadcast:(BOOL)broadcast {
  // A copy that is filled with zeros, see top of file
  uint8_t m[MESSAGE_BUFFER_SIZE];
  memset(m, 0, MESSAGE_BUFFER_SIZE);
  if (length > MESSAGE_BUFFER_SIZE) length = MESSAGE_BUFFER_SIZE;
  if (length > 0) memcpy(m, message, (size_t)length);
  // STEP 1: Extract from the XpressNet header byte the main command type
  uint8_t header = m[0];
  int msgCmd = header >> 4;                               // the main command type in the 4 highest order bits
  // STEP 2: Set the flow-control flag, except when the message was a broadcast
  // 2.0: 01 05 and 01 07 are always unrequested (V3.6: 1.5, framing bytes FF FD). Also for a serial interface,
  // which has no framing bytes: without this check, these would release the flow control
  if ((header == 0x01) && ((m[1] == 0x05) || (m[1] == 0x07))) broadcast = YES;
  if (!broadcast) {
    [[_Parent send] setWaitForAck:NO];
    [[_Parent send] stoptcpResponseTimer];
  }
  // STEP 3: Call the procedure that is specific for the message received
  // Check first for feedback messages, which have as Command = 4 but varying length
  if (msgCmd == 4) {
    [self DecodeFeedback:length message:m];
  }
  else {
    // All other message have a fixed length, so we can do a switch statement on the entire header
    switch (header) {
      case 0x01: [self AnswerFromInterface:m[1] broadcast:broadcast]; break;
      case 0x02: [_Parent eventInterfaceVersion:m[1] daten2:m[2]]; break;
      case 0x61:
        switch (m[1]) {
          case 0x00: [_Parent eventNotAus]; break;
          case 0x01: [_Parent eventAllesAn]; break;
          case 0x02: [_Parent eventProgrammierMode]; break;
          case 0x08: [_Parent eventProgrammierinfoKurzschluss]; break;       // Z21 LAN_X_BC_TRACK_SHORT_CIRCUIT
          case 0x11: [_Parent eventProgrammierinfoZentraleBereit]; break;
          case 0x12: [_Parent eventProgrammierinfoKurzschluss]; break;
          case 0x13: [_Parent eventProgrammierinfoDatenNichtGefunden]; break;
          case 0x1F: [_Parent eventProgrammierinfoZentraleBusy]; break;
          // 2.0: the answer is passed to the SendClass (retransmission or SendError)
          case 0x80:
            [[_Parent Logfile] Errors:@"<= Ubertragungsfehler"];
            if (!broadcast) [[_Parent send] answerError:@"61 80 (Ubertragungsfehler)"];
            break;
          case 0x81:
            [[_Parent Logfile] Errors:@"<= Zentrale Busy"];
            if (!broadcast) [[_Parent send] answerError:@"61 81 (Zentrale Busy)"];
            break;
          case 0x82:
            [[_Parent Logfile] Warnings:@"<= Befehl in Zentrale nicht vorhanden"];
            if (!broadcast) [[_Parent send] answerRejected:@"61 82 (Befehl in Zentrale nicht vorhanden)" reason:BefehlNichtVorhanden];
            break;
          default: break;
        }
        break;
      case 0x62:
        if (m[1] == 0x22) [_Parent eventStatusZentrale:m[2]];
        break;
      case 0x63:
        switch (m[1]) {
          case 0x03: [_Parent eventModellZeit:m[2] daten2:m[3] factor:255]; break;
          case 0x10: [_Parent eventProgInfoPageRegister:m[2] value:m[3]]; break;
          case 0x14: [_Parent eventProgInfo:1 daten1:m[2] value:m[3]]; break;     // CV1-255 u. 1024
          case 0x15: [_Parent eventProgInfo:2 daten1:m[2] value:m[3]]; break;     // CV256 - 511
          case 0x16: [_Parent eventProgInfo:3 daten1:m[2] value:m[3]]; break;     // CV512 - 767
          case 0x17: [_Parent eventProgInfo:4 daten1:m[2] value:m[3]]; break;     // CV768 - 1023
          case 0x20: [_Parent eventServiceVariableMelden:m[2] SvValue:m[3]]; break;
          case 0x21: [_Parent eventSoftwareVersionZentrale:m[2] daten3:m[3]]; break;
          default: break;
        }
        break;
      case 0x64:
        switch (m[1]) {
          case 0x14: [_Parent eventProgInfoZ21:m[2] CVAdr_LSB:m[3] value:m[4]]; break;
          case 0x24: [_Parent eventPomErgebnisMelden:m[2] AdrL:m[3] CvValue:m[4]]; break;
          case 0x25: [_Parent eventModellZeit:m[2] daten2:m[3] factor:m[4]]; break;
          default: break;
        }
        break;
      case 0x67:
        if (m[1] == 0x23) [_Parent eventErweiterteVersionZentrale:m[2] ZBldL:m[3] RMVer:m[4] RMBldH:m[5] RMBldL:m[6] BIvers:m[7]];
        break;
      case 0x81:
        if (m[1] == 0x00) [_Parent eventNotHalt];
        break;
      case 0xE1: [_Parent eventFehlermeldung:m[1]]; break;
      case 0xE2: break;                                   // MTR-Basisadresse
      case 0xE3:
        switch (m[1]) {
          case 0x40: break;                               // Lok besetzt
          case 0x50: [_Parent eventFunctionStatusF0F12:m[2] F5_F12:m[3]]; break;
          case 0x52: [_Parent eventFunctionsZustandF13F20:m[2] F21_F28:m[3]]; break;
          default: break;
        }
        break;
      case 0xE4:
        if (m[1] == 0x51) [_Parent eventFunctionStatusF13F20:m[2] F21_F28:m[3] RefreshMode:m[4]];
        else [_Parent eventNormaleLokinfo:m[1] Speed:m[2] F0_F4:m[3] F5_F12:m[4]];
        break;
      case 0xE5: break;                                   // MTR-Mitglied
      case 0xE6:
        switch (m[1]) {
          case 0x53: [_Parent eventFunctionsZustandF29F68:m[2] F37_F44:m[3] F45_F52:m[4] F53_F60:m[5] F61_F68:m[6]]; break;
          case 0x54: [_Parent eventFunctionStatusF29F68:m[2] F37_F44:m[3] F45_F52:m[4] F53_F60:m[5] F61_F68:m[6]]; break;
          default: break;                                 // Lok ist in DTR
        }
        break;
      case 0xEF:
        [_Parent eventZ21Lokinfo:m[1] Adr_LSB:m[2] Kennung:m[3] Speed:m[4] F0_F4:m[5] F5_F12:m[6] F13_F20:m[7] F21_F28:m[8]];
        break;
      case 0xF2:
        switch (m[1]) {
          case 1: [_Parent eventXpressNetAddress:m[2]]; break;
          case 2: [_Parent eventXpressNetVersion:m[2]]; break;
          case 3: [_Parent eventInterfaceFreieVerbindungen:m[2]]; break;
          default: break;
        }
        break;
      case 0xF3:
        if (m[1] == 0x0A) [_Parent eventz21FirmwareVersion:m[2] V_LSB:m[3]];
        break;
      default: break;
    }
  }
  // STEP 4: Try to send the next message that is waiting in the queue
  // Flow control (check waitForAck) is performed within nextMessage
  [[_Parent send] nextMessage];
  // STEP 5: Update the counter for the number of received messages
  [_Parent setNumberOfExpressNetMessagesReceived:[_Parent NumberOfExpressNetMessagesReceived] + 1];
}


// ======================================================================================================
// V3.6: Section 1.5 - Antworten des Interface 23151
// 2.0: the answers are passed to the SendClass: retransmission (01 01, 02, 03, 06, 0A), SendError (01 09),
// pause (01 05, 01 08) or end of the pause (01 07). An error answer that comes as broadcast is only logged
- (void)AnswerFromInterface:(int)data broadcast:(BOOL)broadcast {
  LogFileClass *logfile = [_Parent Logfile];
  SendClass *send = [_Parent send];
  switch (data) {
    case 1:
      [logfile Errors:@"<= Interface: Die im Header angegebene Anzahl Bytes stimmt nicht mit der Anzahl der empfangenen Bytes überein"];
      if (!broadcast) [send answerError:@"01 01 (Anzahl Bytes stimmt nicht)"];
      break;
    case 2:
      [logfile Errors:@"<= Interface: Fehler zwischen Interface und Zentrale (Timeout bei Datenubertragung des Interface an die Zentrale)"];
      if (!broadcast) [send answerError:@"01 02 (Fehler zwischen Interface und Zentrale)"];
      break;
    case 3:
      [logfile Errors:@"<= Interface: Unbekannter Fehler (Zentrale adressierte das Interface mit Quittierungsaufforderung)"];
      if (!broadcast) [send answerError:@"01 03 (Zentrale adressierte das Interface mit Quittierungsaufforderung)"];
      break;
    case 4:                                               // Befehl ist an Zentrale geschickt
      break;
    case 5:
      [logfile Warnings:@"<= Interface: Zentrale adressiert das Interface nicht mehr"];
      [send pause:NO];
      break;
    case 6:
      [logfile Errors:@"<= Interface: Puffer-Uberlauf im Interface 23151"];
      if (!broadcast) [send answerError:@"01 06 (Puffer-Uberlauf im Interface)"];
      break;
    case 7:
      [logfile Warnings:@"<= Interface: Zentrale adressiert das Interface wieder"];
      [send resume];
      break;
    case 8:
      [logfile Warnings:@"<= Interface: Derzeit konnen keine Befehle an die Zentrale gesendet werden"];
      [send pause:!broadcast];                            // as an answer: the last message is sent again after the pause
      break;
    case 9:
      [logfile Errors:@"<= Interface: Fehler in den Befehlsparametern (z.B. Lokadresse falsch)"];
      if (!broadcast) [send answerRejected:@"01 09 (Fehler in den Befehlsparametern)" reason:FehlerInDenBefehlsparametern];
      break;
    case 10:
      [logfile Errors:@"<= Interface: Unbekannter Fehler (Zentrale lieferte nicht die erwartete Antwort)"];
      if (!broadcast) [send answerError:@"01 0A (Zentrale lieferte nicht die erwartete Antwort)"];
      break;
    default: break;
  }
}


// ----------------------------------------------------------------------------------------------------
// Ruckmeldung / Schaltinformation (V3.6: 3.1.1.5 and 3.1.8, V3.8: 2.4.5, 2.15, 2.16)
// Z21: 5.3: LAN_X_TURNOUT_INFO
// DecodeFeedback is called whenever an XpressNet Feedback message is received.
// It may carry a single feedback message (Schaltinformation) or multiple messages (Ruckmeldung).
// The message may be a point-to-point, or a broadcast message.
// The input is the XpressNet message, without the Framing bytes and XOR byte.
// The addresses in this method are RS-Bus / decoder adresses.
- (void)DecodeFeedback:(int)length message:(const uint8_t *)message {
  // The Lenz feedback messages and the Z21 Extended Accessory Info messages have the same starting bytes.
  // Therefore we first check if we have a Z21 command station (0x12).
  // If yes, we call the associated method and return. If no, we handle the bytes as normal XpressNet messages
  if (([_Parent ZentraleKennung] == 0x12) && (message[0] == 0x44)) {
    [_Parent eventExtendedAccessoryInfo:message[1] Adr_LSB:message[2] DDDDDDDD:message[3] Status:message[4]];
    return;
  }
  // If we are here, we do not have a Z21 command station, but a normal (Lenz) XpressNet command station
  // STEP 2: Determine the length of the feedback message and the number of feedbacks
  int msgLength = message[0] & 0x0F;                      // Isolate the lowest order 4 bits
  if ((msgLength + 1) != length) {                        // Extra test: does the length coded in the message match the received size
    [[_Parent Logfile] Errors:@"Error decoding feedback message: Length doesn't match received Length"];
    return;
  }
  if (msgLength == 3) {
    // STEP 3: V3.8: Schaltinformation ab W1025 / Z21: LAN_X_TURNOUT_INFO
    // The Z21 does not return the Nibble, ttBits and the iBit. Only 2 data bits are used (a single switch)
    // Two bytes for the address, plus one byte for the value
    [self decodeFeedbackAddress:((message[1] * 256) + message[2]) data:message[3]];
  }
  else {
    // STEP 4: V3.6: Schaltinformation / BC Ruckmeldung. Analyse each feedback individually.
    // The RS-Bus address is stored at the indices 1, 3, 5, 7 etc. The RS-Bus data is stored in the byte immediately after
    int numberOfFeedbacks = msgLength / 2;
    for (int i = 0; i < numberOfFeedbacks; i++) {
      int index = 1 + (i * 2);
      [self decodeFeedbackAddress:message[index] data:message[index + 1]];
    }
  }
}


- (void)decodeFeedbackAddress:(int)address data:(int)data {
  int feedbackBits = data & 0x0F;                         // The "real feedback bits" are in bits 0..3
  int nibble = (data & 0x10) >> 4;                        // The nibble is coded in bit 4
  int ttBits = (data & 0x60) >> 5;                        // The TT bits are coded in bits 5 and 6
  int iBit = (data & 0x80) >> 7;                          // The iBit is coded in bit 7
  // Call feedback to raise a feedback event
  [_Parent eventFeedback:address Nibble:nibble ttBits:ttBits iBit:iBit data:feedbackBits];
}

@end
