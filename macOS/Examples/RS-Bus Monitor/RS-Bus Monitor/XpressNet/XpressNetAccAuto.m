//
//  XpressNetAccAuto.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the AccAutoClass of the VB6 XpressNet DLL: implementation of the AccessoryAuto method.
//
//  This method can be used to set multiple switches at once, without worrying about deactivation and scheduling.
//  The standard XpressNet command for accesories include an activation and deactivation bit, which allows
//  activation of the decoder's output port and, after a short interval, deactivation. Such deactivation may,
//  for example, be necessary for coil or motor based switch engines. In practice, coil based switch engines
//  need to be deactivated after several hundreds of ms. Another concern with switch decoders is that, after
//  activation, a single output may consume 1 A of current. If all switches would activate their outputs at
//  exactly the same moment, the amount of current that is needed may easily exceed the amount that is available.
//  Therefore it is important to schedule switch commands after each other.
//
//  For the interval between activation and deactivation, as well as for the interval between two activations,
//  the "AccessoryInterval" attribute (declared in the XpressNet class) will be used.
//  After a deactivation there is a pause of AccessoryInterval / 2, before the next accessory is activated
//  (as in the example of the Z21 specification, 5.2.1).
//
//  Like the "message" method, this method implements a FIFO queue to store accessory commands. As soon as an
//  accessory command is received, the accessory address and position is saved in the FIFO. Next, the associated
//  XpressNet accessory command, with the activation bit set, is issued, and a timer is started. Once the timer
//  fires, the XpressNet command is send again, but this time the activation bit is released. If, in the time
//  between activation and deactivation, additional accessory commands are received, these are stored on the
//  FIFO queue as well.
//
//  Differences with the VB6 version (1.2):
//  - emptyFifo also stops the deactivation timer (VB6: the timer kept running on an empty FIFO).
//  - the pause after a deactivation (VB6: the next accessory was activated immediately).
//
//  Version 2.0 (November 2026), the same in the VB6, .NET and macOS versions:
//  - emptyFifo (after a response timeout, or after a pause without 01 07) first switches off the output of the
//    switch that is activated at that moment. In version 1.2 this switch got no deactivation.
//  - During a pause (see XpressNetSend.m: the command station does not address the interface) no new switch is
//    activated. The activation of the next switch waits for 01 07 (resume).
//  - A full queue gives the event SendError (besides the line in the logfile).
//  - stop (closing the connection, dealloc) removes all commands without a deactivation: there is no connection
//    to send it.
//

#import "XpressNetAccAuto.h"
#import "XpressNet.h"
#import "XpressNetLogFile.h"
#import "XpressNetSend.h"

@implementation AccAutoClass {
  NSMutableArray *_fifo;                  // Each element is an NSArray with two NSNumbers: address and position
  BOOL _pausing;                          // Is the timer used for the pause after a deactivation?
  BOOL _switchActivated;                  // 2.0: is the output of the switch at the head of the FIFO active?
  BOOL _waitingForInterface;              // 2.0: does an activation wait for the end of a pause (01 07)?
}

@synthesize Parent = _Parent;

#define MAX_ACC_FIFO_ENTRIES 32           // The number of Accessory messages the queue is able to hold


- (id)init {
  self = [super init];
  if (self) _fifo = [[NSMutableArray alloc] init];
  return self;
}


- (void)dealloc {
  [_fifo release];
  [super dealloc];
}


- (void)handle:(int)address position:(int)position {
  // Check if the FIFO queue has free slots left to store the message
  // If not, increase MAX_ACC_FIFO_ENTRIES, or modify the calling program
  if ([_fifo count] >= MAX_ACC_FIFO_ENTRIES) {           // FIFO overflow :-(
    [[_Parent Logfile] Errors:@"Error: Auto Accessory queue is full"];
    // 2.0: event SendError
    [_Parent eventSendError:FifoVoll description:[NSString stringWithFormat:@"Error: Auto Accessory queue is full (CmdAccessoryAuto: switch %d, position %d)", address, position]];
    return;
  }
  // Store address and position in the FIFO
  [_fifo addObject:[NSArray arrayWithObjects:[NSNumber numberWithInt:address], [NSNumber numberWithInt:position], nil]];
  // If this is the only accessory command in the FIFO, we may send it right now.
  // But not during the pause after a deactivation: the accessory is then activated after the pause
  if (([_fifo count] == 1) && (!_pausing)) [self activateHead];
}


// Activate the accessory at the head of the FIFO, and start the deactivation timer.
// Note: sending a command may empty the FIFO and stop the timer, if the connection turns out to be closed
// (see eventTcpClosedByRemoteHost). Therefore the FIFO is updated and the timer is started before sending.
// 2.0: not during a pause of the SendClass (01 05 / 01 08): the activation then waits for resume (01 07)
- (void)activateHead {
  if ([[_Parent send] interfaceNotAddressed]) {
    _waitingForInterface = YES;
    return;
  }
  NSArray *accessory = [_fifo objectAtIndex:0];
  int address = [[accessory objectAtIndex:0] intValue];
  int position = [[accessory objectAtIndex:1] intValue];
  // Note: performSelector:afterDelay: retains this object until it is cancelled or has fired (see stop)
  [self performSelector:@selector(deactTimerFired) withObject:nil afterDelay:([_Parent AccessoryInterval] / 1000.0)];
  _switchActivated = YES;
  [_Parent CmdAccessory:address position:position Activate:1];
}


- (void)deactTimerFired {
  // If the pause after a deactivation is over (see STEP 2), activate the next accessory (if the FIFO is not empty)
  if (_pausing) {
    _pausing = NO;
    if ([_fifo count] > 0) [self activateHead];
    return;
  }
  // STEP 1: If the timer expires, we need to remove this accessory from the FIFO
  if ([_fifo count] == 0) return;                        // should not occur: the timer stops if the FIFO is emptied
  NSArray *accessory = [_fifo objectAtIndex:0];
  int address = [[accessory objectAtIndex:0] intValue];
  int position = [[accessory objectAtIndex:1] intValue];
  [_fifo removeObjectAtIndex:0];
  _switchActivated = NO;
  // STEP 2: start the timer for the pause of AccessoryInterval / 2, after which the next accessory is activated
  // (the timer is started before sending, see activateHead)
  _pausing = YES;
  [self performSelector:@selector(deactTimerFired) withObject:nil afterDelay:(([_Parent AccessoryInterval] / 2) / 1000.0)];
  // STEP 3: and send the accessory command to deactivate the output port
  [_Parent CmdAccessory:address position:position Activate:0];
}


- (int)size {
  return (int)[_fifo count];
}


// 2.0: called by the SendClass at the end of a pause (01 07). An activation that waited, is done now
- (void)resume {
  if (!_waitingForInterface) return;
  _waitingForInterface = NO;
  if (([_fifo count] > 0) && (!_pausing) && (!_switchActivated)) [self activateHead];
}


// Removes all accessory commands and stops the deactivation timer
// 2.0: the output of an activated switch is switched off first (1.2: this switch got no deactivation).
// Returns the number of removed commands (without the switch that has been switched off)
- (int)emptyFifo {
  int removed = (int)[_fifo count];
  if (_switchActivated && ([_fifo count] > 0)) {
    NSArray *accessory = [[[_fifo objectAtIndex:0] retain] autorelease];
    [self stop];                                          // First: sending may close the connection (see activateHead)
    [_Parent CmdAccessory:[[accessory objectAtIndex:0] intValue] position:[[accessory objectAtIndex:1] intValue] Activate:0];
    return removed - 1;
  }
  [self stop];
  return removed;
}


- (void)stop {
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(deactTimerFired) object:nil];
  [_fifo removeAllObjects];
  _pausing = NO;
  _switchActivated = NO;                                  // 2.0
  _waitingForInterface = NO;
}

@end
