//
//  XpressNetReceiveZ21.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the ReceiveClassZ21 of the VB6 XpressNet DLL.
//  This class implements the Z21 receive code for XpressNet messages that are transferred "within" Z21 messages
//  (the Z21 XpressNet tunnel). The first 4 bytes of these messages are: DataLen (2 bytes) and Header 0x40 0x00.
//  The remaining bytes are the XpressNet message.
//  Note that the last four bits of the XpressNet header do not always include the correct XpressNet message length.
//  Therefore we must use the Z21 DataLen.
//  For Z21 details, see the Z21 LAN Protocol Specification (Version 1.13, 06.11.2023).
//  According to section 1.3 of the Z21 protocol spec, multiple Z21 messages may be combined within a single UDP packet.
//  Z21 messages other than the XpressNet tunnel (Header 0x40) are ignored, as in the VB6 version.
//
//  Differences with the VB6 version:
//  - The XpressNet message is taken from the received Z21 message (VB6: from the input buffer, which at that moment
//    has already been shifted to the next Z21 message)
//  - DataLen is 16 bits (little endian) (VB6: only the first byte)
//  - A message with a wrong XOR is ignored (VB6: logged, but still decoded)
//  - Each UDP datagram is analysed on its own; a Z21 message is never split over two datagrams
//  - 2.0: every received Z21 message is logged (at most the first 256 bytes), also the messages that are not decoded
//    (such as other headers than 0x40). 1.2: only the XpressNet messages
//

#import "XpressNetReceiveZ21.h"
#import "XpressNet.h"
#import "XpressNetReceive.h"
#import "XpressNetLogFile.h"

@implementation ReceiveClassZ21

@synthesize Parent = _Parent;


- (void)addToIndata:(const uint8_t *)inBytes length:(int)bytesReceived {
  // Since we may have multiple Z21 messages within the datagram, we have a loop that ends at the end of the datagram
  int index = 0;
  while (index < bytesReceived) {
    // STEP 1: Determine the length of the Z21 message
    if ((bytesReceived - index) < 4) {
      [[_Parent Logfile] Errors:@"Error: incomplete Z21 message header"];
      return;
    }
    int z21Length = inBytes[index] + (inBytes[index + 1] * 256);
    // STEP 2: Check if we have at least this number of bytes. If not, the datagram is inconsistent
    if ((z21Length < 4) || (z21Length > (bytesReceived - index))) {
      [[_Parent Logfile] Errors:@"Error: Z21 messagesize exceeds the number of received UDP bytes"];
      return;
    }
    // 2.0: log every received Z21 message (see the top of this file)
    if ([_Parent LogHexMessages]) [[_Parent Logfile] ReceiveHexRange:&inBytes[index] from:0 to:MIN(z21Length - 1, 255)];
    // STEP 3: Analyse the individual Z21 message
    [self DecodeZ21Message:&inBytes[index] length:z21Length];
    index += z21Length;
  }
}


- (void)DecodeZ21Message:(const uint8_t *)z21Message length:(int)z21Length {
  // STEP 1: Return if the Z21Message does not contain an XpressNet message
  if ((z21Message[2] != 0x40) || (z21Message[3] != 0x00)) return;
  // STEP 2: The XpressNet message starts at index 4. It consists of the X-Header, the data bytes and the XOR byte
  int indexXor = z21Length - 1;
  if (indexXor < 5) return;                               // At least an X-Header and a XOR byte
  // STEP 3: Should we log the Z21 message?
  // 2.0: the message has already been logged in addToIndata:length:
  // if ([_Parent LogHexMessages]) [[_Parent Logfile] ReceiveHexRange:z21Message from:0 to:indexXor];
  // STEP 4: Check if the XOR value is correct. If we include the XOR byte in the XOR calculation, the result should become 0
  uint8_t xorValue = 0;
  for (int i = 4; i <= indexXor; i++) xorValue ^= z21Message[i];
  if (xorValue) {
    [[_Parent Logfile] Errors:@"XOR Error in received XpressNet message"];
    if ([_Parent LogErrors] && ![_Parent LogHexMessages]) [[_Parent Logfile] ReceiveHexRange:z21Message from:0 to:indexXor];
    return;
  }
  // STEP 5: Call the method that analyses the individual XpressNet message (length without the XOR byte).
  // Since there is no flow control, knowing if something is a broadcast message doesn't bring anything
  [[_Parent receive] DecodeXpressNetMessage:(z21Length - 5) message:&z21Message[4] broadcast:NO];
}

@end
