//
//  XpressNetSend.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  .NET version of the SendClass of the VB6 XpressNet DLL: implementation of the message and nextMessage methods
//
//  message is called by the main XpressNet code to send (prepare the sending of) an XpressNet message.
//  The parameters to the message method are the individual bytes of the XpressNet message, separated
//  by comma's. These bytes should NOT include the two Framing bytes at the beginning (FF and FE), nor
//  the XOR byte at the end.
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
//  The Z21 (UDP) has no flow control: the Z21 does not answer every command (Z21 spec, section 1.1),
//  therefore Z21 messages are sent immediately.
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
//  - Header bytes for XpressNet / Z21 are added here, in transmit (VB6: in the classes Tcp, Udp and Usb)
//  - As long as the connection is not (yet) established, messages are dropped (VB6: only without connection)
//  - A byte value outside 0..255 is not sent, but written to the logfile (VB6: run time error "Overflow")
//

using System.Collections.Generic;
using System.Text;

namespace XpressNet
{
  internal class SendClass
  {
    private readonly XpressNetClass Parent;

    internal bool waitForAck;                           // Can we send the message now, or should it be stored in the FIFO?
    internal bool waitForlanKeepAliveResponse;          // Flag that helps to distinguish where certain responses belong to

    // --------------------------------------------------------------------------------------------------------------
    // If needed, we store XpressNet messages in a FIFO. Each element holds 1 XpressNet message, including the XOR
    private const int maxFifoEntries = 128;             // The number of XpressNet messages the queue is able to hold
    private const int maxMessageSize = 11;              // Maximum number of bytes of an XpressNet message from PC to
                                                        // Command Station (without XOR). As in VB6: maxFifoMsgSize - 1
    private readonly Queue<byte[]> fifo = new Queue<byte[]>();

    // Timers (see XTimerClass in XpressNetMainThread.cs)
    private const int lanKeepAliveInterval = 30000;     // Time between LAN Keep Alive messages
    private readonly XTimerClass lanKeepAliveTimer;
    private const int tcpResponseTimeOut = 5000;        // A response message should be received within 5 sec
    private const int tcpProgrammingTimeOut = 90000;    // In programming mode a response may take 90 sec
    private readonly XTimerClass tcpResponseTimer;

    // --------------------------------------------------------------------------------------------------------------
    // 2.0: Retransmission (see the header of this file)
    private byte[] lastMessage;                         // The last transmitted message (including the XOR byte)
    private int retryCount;                             // The number of retransmissions of lastMessage
    private bool retryPending;                          // A retransmission of lastMessage waits for the retryTimer
    private const int maxRetries = 3;                   // At most 3 retransmissions
    private const int firstRetryDelay = 100;            // The delays are 100, 200 and 400 msec
    private readonly XTimerClass retryTimer;

    // 2.0: Pause (see the header of this file)
    internal bool interfaceNotAddressed;                // Is sending paused (01 05 / 01 08)? Also read by AccAutoClass
    private byte[] heldMessage;                         // The message that is sent first after the pause (01 08)
    private const int pauseTimeOut = 5000;              // 01 07 should be received within 5 sec
    private const int pauseProgrammingTimeOut = 90000;  // In programming mode within 90 sec
    private readonly XTimerClass pauseTimer;


    internal SendClass(XpressNetClass parent)
    {
      Parent = parent;
      waitForAck = false;                               // The start value is that sending data is allowed
      emptyFifo();
      lanKeepAliveTimer = new XTimerClass(Parent.mainThread);  // Create a new timer object for TCP keep alive messages
      lanKeepAliveTimer.Timer += lanKeepAliveTimer_Timer;
      lanKeepAliveTimer.Interval = lanKeepAliveInterval;
      lanKeepAliveTimer.Enabled = true;
      waitForlanKeepAliveResponse = false;              // We did not yet send a LAN Keep Alive message
      tcpResponseTimer = new XTimerClass(Parent.mainThread);   // Create a new timer object for TCP response time outs
      tcpResponseTimer.Timer += tcpResponseTimer_Timer;
      tcpResponseTimer.Enabled = false;                 // The timer will be initialised as soon as a message is send
      // 2.0: retransmission and pause
      retryTimer = new XTimerClass(Parent.mainThread);
      retryTimer.Timer += retryTimer_Timer;
      retryTimer.Enabled = false;                       // The timer will be started after an error answer
      pauseTimer = new XTimerClass(Parent.mainThread);
      pauseTimer.Timer += pauseTimer_Timer;
      pauseTimer.Enabled = false;                       // The timer will be started at the begin of a pause
      retryCount = 0;
      retryPending = false;
      interfaceNotAddressed = false;
      heldMessage = null;
    }

    // Stops the timers (Dispose)
    internal void close()
    {
      lanKeepAliveTimer.close();
      tcpResponseTimer.close();
      retryTimer.close();
      pauseTimer.close();
    }


    // --------------------------------------------------------------------------------------------------------------
    // message is called from the XpressNetClass code. It either sends a message, or it stores the message
    // in the FIFO. If a message is send, it calles transmit, which in turn activates the flow control
    // by setting the waitForAck Flag to True.
    // Note: Header bytes for Xpressnet / Z21 are, if needed, added in transmit
    // 2.0: a message that is not sent gives the event SendError (besides the line in the logfile). The checks of
    // the message are done first, so the description of the event can include the XOR byte
    internal void message(params int[] byteSequence)
    {
      if (Parent.otherThread(() => message(byteSequence))) return;
      // STEP 1: Perform 3 checks on the size of the received message.
      // Check 1: did we receive any input??
      if ((byteSequence == null) || (byteSequence.Length == 0)) {
        notSent(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Error in the message subroutine: ParamArray is empty", byteSequence);
        return;
      }
      // Check 2: Is the input not too long?
      if (byteSequence.Length > maxMessageSize) {
        notSent(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Error in the message subroutine: ParamArray exceeds maxFifoMsgSize", byteSequence);
        return;
      }
      // Check 3: Does the received number of bytes correspond to the XpressNet message header?
      // The number of bytes received must match with the length as contained in (the lower 4 bits) of the first
      // (=header) byte. Note: the Z21 Specification sometimes violates this, and must be excluded from this check
      if ((!Parent.Z21Protocol) && ((byteSequence.Length - 1) != (byteSequence[0] % 16))) {
        notSent(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Error in the message subroutine: size ParamArray doesn't match size in header", byteSequence);
        return;
      }
      // Extra check: every byte must have a value between 0 and 255 (VB6: run time error "Overflow")
      foreach (int value in byteSequence) {
        if ((value < 0) || (value > 255)) {
          notSent(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Error in the message subroutine: value " + value + " does not fit in a byte", byteSequence);
          return;
        }
      }
      // STEP 2: Create the XpressNet message, including the XOR value
      byte[] outBuffer = new byte[byteSequence.Length + 1];
      byte xorValue = 0;
      for (int i = 0; i < byteSequence.Length; i++) {
        outBuffer[i] = (byte)byteSequence[i];
        xorValue ^= outBuffer[i];
      }
      outBuffer[byteSequence.Length] = xorValue;
      // STEP 3: Check if we have already a connection. If not: drop the packet
      // 2.0: also without any connection (connectionType none) the logfile and the event SendError (1.3: silently)
      if (Parent.connectionType == connectionTypeEnum.none) {
        notSent(SendErrorReasonEnum.KeineVerbindung, "Message not sent: there is no connection", outBuffer);
        return;
      }
      if (!Parent.isConnected()) {
        notSent(SendErrorReasonEnum.KeineVerbindung, "Message not sent: the connection has not (yet) been established", outBuffer);
        return;
      }
      // STEP 4: Check if we can send immediately, or if the message should be stored in the FIFO
      // 2.0: during a pause (01 05 / 01 08) only messages for the interface itself are sent (header F0, F1, F2)
      bool interfaceCommand = (outBuffer[0] >= 0xF0);
      if ((waitForAck == false) && ((!interfaceNotAddressed) || interfaceCommand)) {
        transmit(outBuffer);
      }
      else {                                            // Store the message in the FIFO
        // Check if the FIFO queue has free slots left to store the message
        // If not, increase "maxFifoEntries", or modify the calling program
        if (fifo.Count >= maxFifoEntries) {             // FIFO overflow :-(
          notSent(SendErrorReasonEnum.FifoVoll, "Error: FIFO queue is full", outBuffer);
          return;
        }
        fifo.Enqueue(outBuffer);
      }
    }


    // --------------------------------------------------------------------------------------------------------------
    // nextMessage will be called from the ReceiveClass code. Provided that we are allowed to send now,
    // it sends a message from the FIFO queue, if there are still messages waiting.
    // 2.0: not during a pause. After a pause the held message (01 08) is sent first
    internal void nextMessage()
    {
      if ((waitForAck == false) && (!interfaceNotAddressed)) {   // we may send (no flow control, no pause)
        if (heldMessage != null) {
          byte[] held = heldMessage;
          heldMessage = null;
          transmit(held, true);                         // the held message keeps its number of retransmissions
        }
        // Check if the FIFO has messages waiting. If yes, send the oldest message from the FIFO
        else if (fifo.Count > 0) {
          transmit(fifo.Dequeue());
        }
      }
    }

    // 2.0: retransmission is true if byteArray is sent again (the number of retransmissions is kept)
    internal void transmit(byte[] byteArray, bool retransmission = false)
    {
      // 2.0: keep the message, for a retransmission
      if (!retransmission) retryCount = 0;
      lastMessage = byteArray;
      // No flow control for the Z21 (UDP). The Z21 does not answer every command (Z21 spec 1.1: communication is
      // asynchronous; many commands have no reply), so waiting for an answer could block the FIFO.
      // Messages for the Z21 are therefore sent immediately; the FIFO is not used for UDP
      if (Parent.connectionType != connectionTypeEnum.lanUdpPort) waitForAck = true;  // We may not send a next message immediately
      byte[] outBuffer;
      switch (Parent.connectionType) {                  // Select output
        case connectionTypeEnum.lanTcpPort:
          outBuffer = addHeader(byteArray, 0xFF, 0xFE);    // Two Framing bytes
          if (Parent.LogHexMessages) Parent.Logfile.SendHex(outBuffer);
          lanKeepAliveTimer.Interval = lanKeepAliveInterval;  // Restart the timer, by resetting the Interval
          if (Parent.tcp != null) Parent.tcp.SendData(outBuffer);
          break;
        case connectionTypeEnum.lanUdpPort:
          // The four bytes for the XpressNet tunnel: DataLen (2 bytes, little endian), Header 0x40 0x00
          int dataLen = byteArray.Length + 4;
          outBuffer = addHeader(byteArray, (byte)(dataLen % 256), (byte)(dataLen / 256), 0x40, 0x00);
          if (Parent.LogHexMessages) Parent.Logfile.SendHex(outBuffer);
          lanKeepAliveTimer.Interval = lanKeepAliveInterval;  // Restart the timer, by resetting the Interval
          if (Parent.udp != null) Parent.udp.SendData(outBuffer);
          break;
        case connectionTypeEnum.usbPort:
          outBuffer = addHeader(byteArray, 0xFF, 0xFE);    // Two Framing bytes
          if (Parent.LogHexMessages) Parent.Logfile.SendHex(outBuffer);
          if (Parent.usb != null) Parent.usb.SendData(outBuffer);
          break;
        case connectionTypeEnum.serialPort:             // Serial interfaces have no Framing bytes
          if (Parent.LogHexMessages) Parent.Logfile.SendHex(byteArray);
          if (Parent.usb != null) Parent.usb.SendData(byteArray);
          break;
      }
      // The response timer is used for TCP, USB and serial. The USB interface has the same rule as LAN: no next
      // command before the response, or before the timeout (5 sec, 1.5 min in programming mode)
      // (LI-USB: 1.1, Interface 23151: 1.3). Serial interfaces may omit the confirmation 1/4/5 (LI101F: 1.4).
      // Without this timer, a missing response would block the FIFO.
      if (Parent.connectionType != connectionTypeEnum.lanUdpPort) {
        tcpResponseTimer.Enabled = true;                // This timer triggers if there is no response = Error
        if (Parent.ZentraleProgrammierMode) {           // Responses in programming mode may take 90 sec
          tcpResponseTimer.Interval = tcpProgrammingTimeOut;
        }
        else {
          tcpResponseTimer.Interval = tcpResponseTimeOut;  // In normal mode the response must be within 5 sec
        }
      }
      Parent.NumberOfExpressNetMessagesTransmitted = Parent.NumberOfExpressNetMessagesTransmitted + 1;
    }

    // 2.0: sends a Z21 LAN message that is not an XpressNet message (no tunnel header 0x40, no XOR), such as
    // LAN_SET_BROADCASTFLAGS and LAN_LOGOFF. z21Message contains the complete message, including DataLen and Header.
    // As for the XpressNet messages to the Z21: no flow control and no FIFO. The message is not counted in
    // NumberOfExpressNetMessagesTransmitted, since it is not an XpressNet message
    internal void transmitZ21(byte[] z21Message)
    {
      if ((Parent.connectionType != connectionTypeEnum.lanUdpPort) || (Parent.udp == null)) {
        notSent(SendErrorReasonEnum.KeineVerbindung, "Message not sent: there is no Z21 (UDP) connection", z21Message);
        return;
      }
      if (Parent.LogHexMessages) Parent.Logfile.SendHex(z21Message);
      lanKeepAliveTimer.Interval = lanKeepAliveInterval;  // Restart the timer, by resetting the Interval
      Parent.udp.SendData(z21Message);
    }

    private static byte[] addHeader(byte[] message, params byte[] header)
    {
      byte[] result = new byte[header.Length + message.Length];
      header.CopyTo(result, 0);
      message.CopyTo(result, header.Length);
      return result;
    }

    private void lanKeepAliveTimer_Timer()
    {
      // V3.6: 2.1 (Interface Status Befehl - Keep Alive)
      // For LAN interface, send keep alive message, to inform the interface we are still listening
      // This is required for the TCP (XpressNet) as well as the UDP (Z21) connection
      // 2.0: not while the TCP connection is being established (the message would give the event SendError)
      if (!Parent.isConnected()) return;
      if (Parent.connectionType == connectionTypeEnum.lanTcpPort) {
        waitForlanKeepAliveResponse = true;
        Parent.CmdInterfaceStatus();                    // This short message is used as XpressNet "Keep Alive"
        Parent.Logfile.DetailsSent("TCP Keep Alive Request");
      }
      if (Parent.connectionType == connectionTypeEnum.lanUdpPort) {
        waitForlanKeepAliveResponse = true;
        Parent.CmdStatusZentraleAnfordern();            // This short message seems usable as Z21 "Keep Alive"
        Parent.Logfile.DetailsSent("UDP Keep Alive Request");
      }
    }

    private void tcpResponseTimer_Timer()
    {
      // For TCP, USB and serial: if the timer expires, something went wrong. Let's go to a well-defined state
      // The keep alive message is for TCP only (Interface 23151: chapter 2 is for the LAN interface only)
      if (Parent.connectionType != connectionTypeEnum.lanUdpPort) {
        tcpResponseTimer.Enabled = false;               // Stop the response timer
        // 2.0: the description for the event SendError (before lastMessage changes)
        string description = "No answer within " + (tcpResponseTimer.Interval / 1000) + " sec (message: " + hex(lastMessage) + ")";
        int removed = fifo.Count;
        if (heldMessage != null) removed = removed + 1;
        emptyFifo();                                    // Remove all XpressNet messages still waiting
        heldMessage = null;
        retryTimer.Enabled = false;                     // 2.0: no retransmission is waiting
        retryPending = false;
        retryCount = 0;
        waitForAck = false;                             // Clear the message acknowledgement flag
        // Remove all switch commands still waiting. 2.0: the activated switch is switched off first. waitForAck is
        // cleared before, so this message is sent immediately
        int removedAccessories = Parent.AccAuto.emptyFifo();
        if (Parent.connectionType == connectionTypeEnum.lanTcpPort) {
          waitForlanKeepAliveResponse = true;           // We will send a keep alive message
          Parent.CmdInterfaceStatus();                  // We will send a keep alive message
        }
        Parent.Logfile.Errors("Error: TCP Response Timer has timed out. Message queues emptied");
        description = description + ". Message queues emptied (" + removed + " messages, " + removedAccessories + " CmdAccessoryAuto commands)";
        Parent.eventSendError(SendErrorReasonEnum.Timeout, description);
      }
    }

    internal void stoptcpResponseTimer()
    {
      // Will be called from the receive class as soon as a response is received
      tcpResponseTimer.Enabled = false;
    }

    internal int size()
    {
      return fifo.Count;
    }

    internal void emptyFifo()
    {
      fifo.Clear();
    }


    // --------------------------------------------------------------------------------------------------------------
    // 2.0: RETRANSMISSION (see the header of this file)
    // Called from the ReceiveClass if the interface answers the last message with an error that may disappear if the
    // message is sent again: 61 80, 61 81, 01 01, 01 02, 01 03, 01 06 and 01 0A. answer is the text for the logfile
    internal void answerError(string answer)
    {
      if (Parent.connectionType == connectionTypeEnum.lanUdpPort) return;   // Z21: no flow control (see the header)
      if (lastMessage == null) return;                  // No message has been sent yet
      if (retryCount < maxRetries) {
        int delay = firstRetryDelay << retryCount;      // 100, 200, 400 msec
        retryCount = retryCount + 1;
        retryPending = true;
        waitForAck = true;                              // The FIFO waits for the retransmission
        retryTimer.Interval = delay;
        retryTimer.Enabled = true;
        Parent.Logfile.Warnings("Answer " + answer + ": retransmission " + retryCount + " of " + maxRetries + " after " + delay + " msec (message: " + hex(lastMessage) + ")");
      }
      else {
        retryCount = 0;
        string description = "Message given up after " + maxRetries + " retransmissions. Answer: " + answer + " (message: " + hex(lastMessage) + ")";
        Parent.Logfile.Errors(description);
        Parent.eventSendError(SendErrorReasonEnum.WiederholungErfolglos, description);
        // waitForAck is false: the next message is sent by nextMessage (ReceiveClass.DecodeXpressNetMessage, STEP 4)
      }
    }

    // Called from the ReceiveClass if the interface answers the last message with 01 09 or 61 82. A retransmission
    // does not help: the message is given up, and the next message is sent. reason: FehlerInDenBefehlsparametern (01 09)
    // or BefehlNichtVorhanden (61 82)
    internal void answerRejected(string answer, SendErrorReasonEnum reason)
    {
      if (Parent.connectionType == connectionTypeEnum.lanUdpPort) return;   // Z21: no flow control (see the header)
      if (lastMessage == null) return;                  // No message has been sent yet
      retryCount = 0;
      string description = "Message rejected. Answer: " + answer + " (message: " + hex(lastMessage) + ")";
      Parent.Logfile.Errors(description);
      Parent.eventSendError(reason, description);
    }

    private void retryTimer_Timer()
    {
      retryTimer.Enabled = false;                       // The XTimerClass is periodic: we need a single tick
      if (!retryPending) return;
      retryPending = false;
      if (interfaceNotAddressed) {                      // A pause started in the meantime: send after the pause
        heldMessage = lastMessage;
        waitForAck = false;                             // No message is underway
        return;
      }
      transmit(lastMessage, true);
    }


    // --------------------------------------------------------------------------------------------------------------
    // 2.0: PAUSE (see the header of this file)
    // Called from the ReceiveClass after 01 05 (answer = false) or 01 08 (answer = true). With 01 08 the last
    // message did not reach the command station: it is held, and sent first after the pause
    internal void pause(bool answer)
    {
      if (Parent.connectionType == connectionTypeEnum.lanUdpPort) return;   // Z21: no flow control (see the header)
      if (answer && (lastMessage != null)) heldMessage = lastMessage;
      if (interfaceNotAddressed) return;                // The pause has started already
      interfaceNotAddressed = true;
      if (Parent.ZentraleProgrammierMode) {             // In programming mode the pause may take 90 sec
        pauseTimer.Interval = pauseProgrammingTimeOut;
      }
      else {
        pauseTimer.Interval = pauseTimeOut;
      }
      pauseTimer.Enabled = true;
      Parent.Logfile.Warnings("Sending paused, until the command station addresses the interface again");
      Parent.eventInterfaceNichtAdressiert();
    }

    // Called from the ReceiveClass after 01 07. The held message is sent first, then the FIFO and CmdAccessoryAuto.
    // The event is raised last, so messages the application sends from the event handler come after these
    internal void resume()
    {
      if (Parent.connectionType == connectionTypeEnum.lanUdpPort) return;   // Z21: no flow control (see the header)
      if (!interfaceNotAddressed) return;               // 01 07 without a pause (for example after power on): ignored
      interfaceNotAddressed = false;
      pauseTimer.Enabled = false;
      Parent.Logfile.Warnings("Sending resumed: the command station addresses the interface again");
      nextMessage();                                    // First the held message, otherwise the FIFO
      Parent.AccAuto.resume();                          // CmdAccessoryAuto continues
      Parent.eventInterfaceWiederAdressiert();
    }

    private void pauseTimer_Timer()
    {
      pauseTimer.Enabled = false;                       // The XTimerClass is periodic: we need a single tick
      if (!interfaceNotAddressed) return;
      int seconds = pauseTimer.Interval / 1000;
      interfaceNotAddressed = false;                    // The pause ends
      int removed = fifo.Count;
      if (heldMessage != null) removed = removed + 1;
      emptyFifo();                                      // Remove all XpressNet messages still waiting
      heldMessage = null;
      retryCount = 0;
      if (retryPending) {                               // No retransmission anymore
        retryTimer.Enabled = false;
        retryPending = false;
        waitForAck = false;
      }
      // Remove all switch commands still waiting. The activated switch (if any) is switched off first
      int removedAccessories = Parent.AccAuto.emptyFifo();
      string description = "The command station did not address the interface again within " + seconds + " sec. Message queues emptied (" + removed + " messages, " + removedAccessories + " CmdAccessoryAuto commands)";
      Parent.Logfile.Errors(description);
      Parent.eventSendError(SendErrorReasonEnum.ZentraleAdressiertDasInterfaceNichtMehr, description);
    }


    // --------------------------------------------------------------------------------------------------------------
    // 2.0: A message that is not sent: logfile and event SendError. The description contains the bytes of the message
    private void notSent(SendErrorReasonEnum reason, string text, byte[] message)
    {
      writeLog(reason, text);
      Parent.eventSendError(reason, text + " (message: " + hex(message) + ")");
    }

    private void notSent(SendErrorReasonEnum reason, string text, int[] byteSequence)
    {
      writeLog(reason, text);
      StringBuilder s = new StringBuilder();
      if (byteSequence != null) {
        foreach (int value in byteSequence) {
          if (s.Length > 0) s.Append(" ");
          s.Append(value.ToString("X2"));
        }
      }
      Parent.eventSendError(reason, text + " (message: " + s + ")");
    }

    // The same log level as in version 1.3: a missing connection is a warning, the rest an error
    private void writeLog(SendErrorReasonEnum reason, string text)
    {
      if (reason == SendErrorReasonEnum.KeineVerbindung) Parent.Logfile.Warnings(text);
      else Parent.Logfile.Errors(text);
    }

    // The bytes of a message as text, for example "52 00 88 DA"
    internal static string hex(byte[] message)
    {
      if (message == null) return "";
      StringBuilder s = new StringBuilder();
      foreach (byte value in message) {
        if (s.Length > 0) s.Append(" ");
        s.Append(value.ToString("X2"));
      }
      return s.ToString();
    }
  }
}
