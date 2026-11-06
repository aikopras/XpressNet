//
//  XpressNetAccAuto.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  .NET version of the AccAutoClass of the VB6 XpressNet DLL: implementation of CmdAccessoryAuto.
//
//  This method can be used to set multiple switches at once, without worrying about deactivation
//  and scheduling. The standard XpressNet command for accesories include an activation and deactivation
//  bit, which allows activation of the decoder's output port and, after a short interval, deactivation.
//  Such deactivation may, for example, be necessary for coil or motor based switch engines.
//  In practice, coild based switch engines need to be deactivated after several hundreds of ms.
//  Another concern with switch decoders is that, after activation, a single output may consume 1 A
//  of current. If all switches would activate their outputs at exactly the same moment, the amount
//  of current that is needed may easily exceed the amount that is available. Therefore it is importent
//  to schedule switch commands after each other.
//
//  For the interval between activation and deactivation the "AccessoryInterval" variable (declared in the
//  XpressNet class) will be used. After a deactivation there is a pause of AccessoryInterval / 2, before the
//  next accessory is activated (as in the example of the Z21 specification, 5.2.1).
//
//  Like the "message" subroutine, this method implements a FIFO queue to store accessory commands
//  As soon as an accessory command is received, the accessory address and position is saved in the FIFO.
//  Next, the associated XpressNet accessory command, with the activation bit set, is issued, and a
//  timer is started. Once the timer fires, the XpressNet command is send again, but this time the
//  activation bit is released.
//  If, in the time between activation and deactivation, additional accessory commands are received,
//  these are stored on the FIFO queue as well.
//
//  This method is implemented in a separate class, since it is, strictly speaking, not part of the
//  XpressNet interface, but rather a layer on top of the XpressNet DLL. Still it should be instantiated
//  in, and called from the XpressNet class object.
//
//  Version 2.0 (November 2026), the same in the VB6, .NET and macOS versions:
//  - emptyFifo (after a response timeout, or after a pause without 01 07) first switches off the output of the
//    switch that is activated at that moment. In version 1.3 this switch got no deactivation.
//  - During a pause (see XpressNetSend.cs: the command station does not address the interface) no new switch is
//    activated. The activation of the next switch waits for 01 07 (resume).
//  - A full queue gives the event SendError (besides the line in the logfile).
//  - The interval of AccessoryInterval still starts at the moment the activation is offered to the SendClass
//    (as in version 1.3). If the activation waits in the FIFO of the SendClass, the output is active shorter.
//

using System.Collections.Generic;

namespace XpressNet
{
  internal class AccAutoClass
  {
    private readonly XpressNetClass Parent;

    // --------------------------------------------------------------------------------------------------------------
    // We store accessory commands in a FIFO. Each element holds two fields: the address and the position of the switch.
    // The maximum number of commands is set by maxFifoEntries.
    private const int maxFifoEntries = 32;              // The number of Accessory messages the queue is able to hold
    private struct AccessoryDataType
    {
      internal int address;
      internal byte position;
    }
    private readonly Queue<AccessoryDataType> fifo = new Queue<AccessoryDataType>();   // this is the queue
    private bool pausing;                               // is the timer used for the pause after a deactivation?
    private bool switchActivated;                       // 2.0: is the output of the switch at the head of the FIFO active?
    private bool waitingForInterface;                   // 2.0: does an activation wait for the end of a pause (01 07)?

    // For the timer we use the XTimerClass (see XpressNetMainThread.cs)
    private readonly XTimerClass deactTimer;


    internal AccAutoClass(XpressNetClass parent)
    {
      Parent = parent;
      // The timer is created first, since emptyFifo stops the timer
      deactTimer = new XTimerClass(Parent.mainThread);  // Create a new timer object
      deactTimer.Timer += deactTimer_Timer;
      emptyFifo();
    }

    // Stops the timer (Dispose)
    internal void close()
    {
      deactTimer.close();
    }

    internal void handle(int address, byte position)
    {
      if (Parent.otherThread(() => handle(address, position))) return;
      // Check if the FIFO queue has free slots left to store the message
      // If not, increase "maxFifoEntries", or modify the calling program
      if (fifo.Count >= maxFifoEntries) {               // FIFO overflow :-(
        Parent.Logfile.Errors("Error: Auto Accessory queue is full");
        // 2.0: event SendError
        Parent.eventSendError(SendErrorReasonEnum.FifoVoll, "Error: Auto Accessory queue is full (CmdAccessoryAuto: switch " + address + ", position " + position + ")");
        return;
      }
      AccessoryDataType accessory;                      // store address and position to the FIFO
      accessory.address = address;
      accessory.position = position;
      fifo.Enqueue(accessory);
      // If this is the only accessory command in the FIFO, we may send it right now
      // But not during the pause after a deactivation (the accessory is activated after the pause)
      if ((fifo.Count == 1) && (!pausing)) {
        activateHead();                                 // activate the accessory, and start the timer
      }
    }

    // 2.0: activates the output of the switch at the head of the FIFO, and starts the deactivation timer
    // (1.3: the same code in handle and deactTimer_Timer). Not during a pause of the SendClass (01 05 / 01 08):
    // the activation then waits for resume (01 07)
    private void activateHead()
    {
      if (Parent.send.interfaceNotAddressed) {
        waitingForInterface = true;
        return;
      }
      AccessoryDataType head = fifo.Peek();
      Parent.CmdAccessory(head.address, head.position, 1);  // activate the accessory
      switchActivated = true;
      deactTimer.Interval = Parent.AccessoryInterval;   // set the deactivation timer's interval
      deactTimer.Enabled = true;                        // start the timer
    }

    private void deactTimer_Timer()
    {
      // STEP 1: If the timer expires, we need to send the accessory command to deactivate the output port
      deactTimer.Enabled = false;
      // If the pause after a deactivation is over (see STEP 3), activate the port of the next
      // accessory (if the FIFO is not empty), and start the timer
      if (pausing) {
        pausing = false;
        if (fifo.Count > 0) activateHead();
        return;
      }
      if (fifo.Count == 0) return;                      // should not occur: the timer stops if the FIFO is emptied
      AccessoryDataType head = fifo.Peek();
      switchActivated = false;
      Parent.CmdAccessory(head.address, head.position, 0);
      // STEP 2: After deactivation, we should remove this accesory from the FIFO
      if (fifo.Count > 0) fifo.Dequeue();
      // STEP 3: First a pause of AccessoryInterval / 2. The next accessory is activated after the pause
      pausing = true;
      deactTimer.Interval = Parent.AccessoryInterval / 2;  // set the timer's interval for the pause
      deactTimer.Enabled = true;                           // start the timer
    }

    // 2.0: called by the SendClass at the end of a pause (01 07). An activation that waited, is done now
    internal void resume()
    {
      if (!waitingForInterface) return;
      waitingForInterface = false;
      if ((fifo.Count > 0) && (!pausing) && (!switchActivated)) activateHead();
    }

    internal int size()
    {
      return fifo.Count;
    }

    // 2.0: the output of an activated switch is switched off first (1.3: this switch got no deactivation).
    // Returns the number of removed commands (without the switch that has been switched off)
    internal int emptyFifo()
    {
      int removed = fifo.Count;
      if (switchActivated && (fifo.Count > 0)) {
        switchActivated = false;
        AccessoryDataType head = fifo.Peek();
        Parent.CmdAccessory(head.address, head.position, 0);
        removed = removed - 1;
      }
      switchActivated = false;
      waitingForInterface = false;
      fifo.Clear();
      // Stop the deactivation timer as well
      deactTimer.Enabled = false;
      pausing = false;
      return removed;
    }
  }
}
