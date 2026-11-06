VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Implementation of the Message and SendNextMessage methods
'
' message is called by the main XpressNet code to send (prepare the sending of) an XpressNet message
' The parameters to the message method are the individual bytes of the XpressNet message, separated
' by comma's. These bytes should NOT include the two Framing bytes at the beginning (FF and FE), nor
' the XOR byte at the end.
'
' nextMessage is called from the ReceiveClass as soon as a previous message has been acknowledged
' by the command station. nextMessage doesn't need any parameters, but instead takes (pop) the oldest
' XpressNet message that is still waiting from the FIFO queue
'
' A FIFO queue is needed, since XpressNet implements flow control between PC and the Command Station.
' The Command Station acknowledges the reception of every individual message to the PC.
' The PC can send the next message only after the acknowledgement of the previous message is received,
' or a time-out (of 5 seconds during normal operation, or 90 seconds during programming) has occured.
' The FIFO allows the main program to "send" multiple XpressNet messages at once, and thus relieves
' the main program from any concerns regarding flow control
'
' Version 2.0 (November 2026). The functionality below is the same in the VB6, .NET and macOS versions:
' - Retransmission. If the interface answers a message with an error that may disappear if the message is sent
'   again, the message is sent again: at most 3 times, after 100, 200 and 400 msec. These answers are:
'   61 80 (Ubertragungsfehler), 61 81 (Zentrale Busy), 01 01, 01 02, 01 03, 01 06 and 01 0A (V3.6: 1.5 and 3.1.5;
'   for 01 0A the specification advises to repeat the last command). If the last retransmission fails as well, the
'   message is given up: logfile and event SendError, and the next message from the FIFO is sent.
'   The answers 01 09 (Fehler in den Befehlsparametern) and 61 82 (Befehl in Zentrale nicht vorhanden) are not
'   solved by a retransmission: logfile and event SendError, and the next message is sent.
' - Pause. After 01 05 (Zentrale adressiert das Interface nicht mehr) or 01 08 (Derzeit konnen keine Befehle an
'   die Zentrale gesendet werden) the interface can not pass messages to the command station. Sending is paused
'   (event InterfaceNichtAdressiert); the FIFO is kept. With 01 08 the last message did not reach the command
'   station: it is held, and sent first after the pause. After 01 07 (Zentrale adressiert das Interface wieder)
'   sending continues (event InterfaceWiederAdressiert). If 01 07 does not come within 5 sec (90 sec in
'   programming mode), both queues are emptied (event SendError) and the pause ends.
'   During the pause, messages for the interface itself (header F0, F1, F2; for example the TCP keep alive
'   message) are still sent: these are answered by the interface, and not by the command station.
' - Nothing is dropped silently. A message that is not sent gives a line in the logfile and the event SendError.
'   Also without a connection, or as long as the connection has not (yet) been established (1.3: silently).
' - Retransmission and pause are not used for the Z21 (UDP): without flow control it is not known to which message
'   an answer belongs.
' Note: "resume" is a reserved word in VB6. The methods are therefore called pauseSending and resumeSending
' ----------------------------------------------------------------------------------------------------

Option Explicit

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long

Public waitForAck As Boolean                      ' Can we send the message now, or should it be stored in the FIFO?
Public waitForlanKeepAliveResponse As Boolean     ' Flag that helps to distinguish where certain responses belong to

' ----------------------------------------------------------------------------------------------------
' If needed, we store XpressNet messages in a FIFO. The FIFO is a two-dimensional Array of bytes
' Each Array row holds 1 XpressNet message. The maximum number of messages is set by maxFifoEntries.
' maxFifoMsgSize defines the maximum size of individual FIFO entries. This size should at least be the
' maximum XpressNet message size (including Framing and XOR bytes), plus one byte to store the actual
' XpressNet message size. This actual messages size is stored as first element of an Array row
Private Const maxFifoEntries = 128                 ' The number of XpressNet messages the queue is able to hold
Private Const maxFifoMsgSize = 12                 ' For message from PC to Command Station
                                                  ' Should be higher than the maximum XpressNet message size
Private head As Byte                              ' we take elements from the head of the queue
Private tail As Byte                              ' tail points to the position of the most recent element
Private numberOfFifoMessages As Byte
Private fifo(maxFifoEntries, maxFifoMsgSize) As Byte

' For timers we instantiate objects of the XTimerClass, which was copied from the Internet
Private Const lanKeepAliveInterval = 30000        ' Time between LAN Keep Alive messages
Private WithEvents lanKeepAliveTimer As XTimerClass
Attribute lanKeepAliveTimer.VB_VarHelpID = -1
Private Const tcpResponseTimeOut = 5000           ' A response message should be received within 5 sec
Private Const tcpProgrammingTimeOut = 90000       ' In programming mode a response may take 90 sec
Private WithEvents tcpResponseTimer As XTimerClass
Attribute tcpResponseTimer.VB_VarHelpID = -1

' 2026-11 modified: retransmission (see the header of this file)
Private lastMessage() As Byte                     ' The last transmitted message (including the XOR byte)
Private lastMessageValid As Boolean               ' Has a message been transmitted (is lastMessage filled)?
Private retryCount As Integer                     ' The number of retransmissions of lastMessage
Private retryPending As Boolean                   ' A retransmission of lastMessage waits for the retryTimer
Private Const maxRetries = 3                      ' At most 3 retransmissions
Private Const firstRetryDelay = 100               ' The delays are 100, 200 and 400 msec
Private WithEvents retryTimer As XTimerClass
Attribute retryTimer.VB_VarHelpID = -1

' 2026-11 modified: pause (see the header of this file)
Public interfaceNotAddressed As Boolean           ' Is sending paused (01 05 / 01 08)? Also read by AccAutoClass
Private heldMessage() As Byte                     ' The message that is sent first after the pause (01 08)
Private heldMessageValid As Boolean               ' Is there a held message?
Private Const pauseTimeOut = 5000                 ' 01 07 should be received within 5 sec
Private Const pauseProgrammingTimeOut = 90000     ' In programming mode within 90 sec
Private WithEvents pauseTimer As XTimerClass
Attribute pauseTimer.VB_VarHelpID = -1



Private Sub Class_Initialize()
  waitForAck = False                              ' The start value is that sending data is allowed
  Call emptyFifo
  Set lanKeepAliveTimer = New XTimerClass         ' Create a new timer object for TCP keep alive messages
  lanKeepAliveTimer.Interval = lanKeepAliveInterval
  lanKeepAliveTimer.Enabled = True
  waitForlanKeepAliveResponse = False             ' We did not yet send a LAN Keep Alive message
  Set tcpResponseTimer = New XTimerClass          ' Create a new timer object for TCP response time outs
  tcpResponseTimer.Enabled = False                ' The timer will be initialised as soon as a message is send
  ' 2026-11 modified: retransmission and pause
  Set retryTimer = New XTimerClass
  retryTimer.Enabled = False                      ' The timer will be started after an error answer
  Set pauseTimer = New XTimerClass
  pauseTimer.Enabled = False                      ' The timer will be started at the begin of a pause
  retryCount = 0
  retryPending = False
  lastMessageValid = False
  interfaceNotAddressed = False
  heldMessageValid = False
End Sub


' ----------------------------------------------------------------------------------------------------
' message is called from the XpressNetClass code. It either sends a message, or it stores the message
' in the FIFO. If a message is send, it calles transmit, which in turn activates the flow control
' by setting the waitForAck Flag to True.
' It was decided to have a ParamArray as input, and not an Array of Bytes.
' The advantage is, that the calling code becomes easier to understand, since there is
' no need to define an Array of Bytes and fill the individual elements of that array.
' The disadvantage is, that the individual variables must be of type "Variant",
' which involves some overhead during the procedure call.
' Note: Header bytes for Xpressnet / Z21 are, if needed, added in the classes: Tcp, Udp / Usb
' 2026-11 modified: a message that is not sent gives the event SendError (besides the line in the logfile). The checks
' of the message are done first, so the description of the event can include the XOR byte
Public Sub message(ParamArray byteSequence() As Variant)
  Dim lengthParamArray As Byte    ' Number of bytes we received from the calling procedure
  Dim i As Byte                   ' will be used in For loops
  Dim params As Variant           ' 2026-11 modified: a copy of the ParamArray, to pass it to paramsHex
  params = byteSequence
  ' STEP 1: Check if we have already a connection
  ' 2026-11 modified: the connection is checked after the checks of the message (see STEP 3A)
  'If Parent.connectionType = none Then Exit Sub        ' Ignore everything. Drop the packet
  ' STEP 2: Perform 3 checks on the size of the received message.
  ' Check 1: did we receive any input??
  If IsMissing(byteSequence) Then
    'Parent.Logfile.Errors "Error in the message subroutine: ParamArray is empty"
    Call notSentText(FehlerInDenBefehlsparametern, "Error in the message subroutine: ParamArray is empty", "")
    Exit Sub
  End If
  ' Check 2: Is the input not too long?
  ' The check is to protect the FIFO against out of range (runtime) errors. The check is whether the
  ' received bytes fit in an FIFO row; the check is NOT if the received bytes fit in an XpressNet message.
  ' Note that UBound does not return the number of bytes received, but 1 less, since the
  ' ParamArray is indexed from 0 .. number - 1
  ' 2026-11 modified: the check is done before lengthParamArray (a Byte) gets the value (Overflow)
  'lengthParamArray = UBound(byteSequence)                 ' length = # parameters - 1!!
  'If lengthParamArray >= (maxFifoMsgSize - 1) Then        ' -1, to have space for the XOR byte
  '  Parent.Logfile.Errors "Error in the message subroutine: ParamArray exceeds maxFifoMsgSize"
  '  Exit Sub
  'End If
  If UBound(byteSequence) >= (maxFifoMsgSize - 1) Then    ' -1, to have space for the XOR byte
    Call notSentText(FehlerInDenBefehlsparametern, "Error in the message subroutine: ParamArray exceeds maxFifoMsgSize", paramsHex(params))
    Exit Sub
  End If
  lengthParamArray = UBound(byteSequence)                 ' length = # parameters - 1!!
  ' Check 3: Does the received number of bytes correspond to the XpressNet message header?
  ' The number of bytes received must match with the length as contained in (the lower 4 bits) of the first
  ' (=header) byte. Note: the Z21 Specification sometimes violates this, and must be excluded from this check
  If (Not Parent.Z21Protocol) And (lengthParamArray <> (byteSequence(0) Mod 16)) Then
    'Parent.Logfile.Errors "Error in the message subroutine: size ParamArray doesn't match size in header"
    Call notSentText(FehlerInDenBefehlsparametern, "Error in the message subroutine: size ParamArray doesn't match size in header", paramsHex(params))
    Exit Sub
  End If
  ' 2026-11 modified: extra check: every byte must have a value between 0 and 255 (was: run time error "Overflow")
  For i = 0 To lengthParamArray
    If (CLng(byteSequence(i)) < 0) Or (CLng(byteSequence(i)) > 255) Then
      Call notSentText(FehlerInDenBefehlsparametern, "Error in the message subroutine: value " & byteSequence(i) & " does not fit in a byte", paramsHex(params))
      Exit Sub
    End If
  Next
  ' STEP 3: Calculate the XOR value
  ' 2026-11 modified: the message (including the XOR byte) is created here, for the connection check, the FIFO and
  ' the event SendError
  Dim xorValue As Byte                       ' The result will be stored in this byte
  Dim msg() As Byte                          ' The complete XpressNet message, including the XOR byte
  ReDim msg(lengthParamArray + 1)
  xorValue = 0
  For i = 0 To lengthParamArray              ' include all bytes that we received as input
    msg(i) = byteSequence(i)
    xorValue = xorValue Xor msg(i)
  Next
  msg(lengthParamArray + 1) = xorValue
  ' STEP 3A: 2026-11 modified: check if we have already a connection. If not: drop the packet, with the logfile and
  ' the event SendError (1.3: silently)
  If Parent.connectionType = none Then
    Call notSent(KeineVerbindung, "Message not sent: there is no connection", msg)
    Exit Sub
  End If
  If Not Parent.isConnected Then
    Call notSent(KeineVerbindung, "Message not sent: the connection has not (yet) been established", msg)
    Exit Sub
  End If
  ' STEP 4: Check if we can send immediately, or if the message should be stored in the FIFO
  ' 2026-11 modified: during a pause (01 05 / 01 08) only messages for the interface itself are sent (header F0, F1, F2)
  Dim interfaceCommand As Boolean
  interfaceCommand = (msg(0) >= &HF0)
  'If waitForAck = False Then
  If (waitForAck = False) And ((Not interfaceNotAddressed) Or interfaceCommand) Then
    ' STEP A1: Create an array that holds all XpressNet bytes. The Array size should match the number of bytes
    ' 2026-11 modified: the array (msg) is created in STEP 3
    'Dim outMsgSize As Byte                    ' msgSize will hold the number of bytes in the message
    'outMsgSize = lengthParamArray + 1         ' Number of bytes for the complete XpressNet message: add 1 for the XOR
    'Dim outBuffer() As Byte                   ' Create a new array, to hold all XpressNet bytes
    'ReDim outBuffer(outMsgSize)               ' Since this array starts at 0, its size should be msgSize - 1
    'For i = 0 To lengthParamArray
    '  outBuffer(i) = byteSequence(i)        ' Copy all bytes this routine received to the outBuffer
    'Next
    'outBuffer(lengthParamArray + 1) = xorValue
    ' STEP A2 : transmit the message over the wire
    'transmit outBuffer
    transmit msg
  Else                                          ' Store the message in aFIFIO
    ' STEP B1: Check if the FIFO queue has free slots left to store the message
    ' If not, increase "maxFifoEntries", or modify the calling program
    If numberOfFifoMessages >= maxFifoEntries Then  ' FIFO overflow :-(
      'Parent.Logfile.Errors "Error: FIFO queue is full"
      Call notSent(FifoVoll, "Error: FIFO queue is full", msg)
      Exit Sub
    End If
    ' STEP B2: Determine the index (=tail) where this new message should be stored within the FIFO queue
    numberOfFifoMessages = numberOfFifoMessages + 1  ' the FIFO has space to store the new element
    If numberOfFifoMessages > 1 Then           ' In the queue is empty, tail does not need to be increased
      tail = tail + 1                          ' If the queue is not empty, point to the next empty slot
      tail = tail Mod maxFifoEntries           ' And of course modulo the number of fifo slots
    End If                                     ' tail now points to the slot where the new element will be saved
    ' STEP B3: Fill the FIFO at position tail
    ' Store the length of the saved XpressNet message at index 0 (length starts at 0)
    ' Add Framing bytes, if needed. Add the XOR at the End
    fifo(tail, 0) = lengthParamArray + 2     ' We add 2 bytes: a length byte and the XOR
    ' 2026-11 modified: copied from msg (including the XOR byte)
    'For i = 0 To lengthParamArray            ' Copy the bytes that this subroutine received as input
    '  fifo(tail, i + 1) = byteSequence(i)
    'Next i
    'fifo(tail, lengthParamArray + 2) = xorValue
    For i = 0 To lengthParamArray + 1        ' Copy the bytes of the message, including the XOR byte
      fifo(tail, i + 1) = msg(i)
    Next i
  End If
End Sub


' ----------------------------------------------------------------------------------------------------
' nextMessage will be called from the ReceiveClass code. Provided that we are allowed to send now,
' it sends a message from the FIFO queue, if there are still messages waiting.
' 2026-11 modified: not during a pause. After a pause the held message (01 08) is sent first
Public Sub nextMessage()
  Dim i As Byte                               ' will be used in For loops
  'If waitForAck = False Then                  ' we may send (no flow control)
  If (waitForAck = False) And (Not interfaceNotAddressed) Then   ' we may send (no flow control, no pause)
    If heldMessageValid Then
      Dim held() As Byte
      held = heldMessage
      heldMessageValid = False
      transmit held, True                     ' the held message keeps its number of retransmissions
    ' Check if the FIFO has messages waiting. If yes, send the oldest message from the FIFO
    'If numberOfFifoMessages > 0 Then            ' there are still XpressNet messages waiting in the FIFO
    ElseIf numberOfFifoMessages > 0 Then        ' there are still XpressNet messages waiting in the FIFO
      ' STEP 1: copy the contents of the FIFO to an Array of Bytes
      Dim outMsgSize As Byte                    ' msgSize will hold the number of bytes in the message
      outMsgSize = fifo(head, 0)                ' Number of bytes of the complete XpressNet message
      Dim outBuffer() As Byte                   ' Create a new array, to hold all XpressNet bytes
      ReDim outBuffer(outMsgSize - 1)           ' Since this array starts at 0, its size should be msgSize - 1
      For i = 1 To outMsgSize                   ' Copy from the FIFO to the outBuffer
        outBuffer(i - 1) = fifo(head, i)        ' the XPressNet bytes
      Next
      ' STEP 2: transmit the message over the wire
      transmit outBuffer
      ' STEP 3: update the variables that maintain the FIFO
      numberOfFifoMessages = numberOfFifoMessages - 1
      If numberOfFifoMessages > 0 Then          ' If the queue became empty, head does not need to be increased
        head = head + 1                         ' point the head to the next (occupied) slot
        head = head Mod maxFifoEntries          ' And of course modulo the number of fifo slots
      End If
    End If
  End If
End Sub

' 2026-11 modified: retransmission is True if byteArray is sent again (the number of retransmissions is kept)
'Public Sub transmit(ByRef byteArray() As Byte)
Public Sub transmit(ByRef byteArray() As Byte, Optional ByVal retransmission As Boolean = False)
  ' 2026-11 modified: keep the message, for a retransmission
  If Not retransmission Then retryCount = 0
  lastMessage = byteArray
  lastMessageValid = True
  ' 2026-10 modified: no flow control for the Z21 (UDP). The Z21 does not answer every command (Z21 spec 1.1:
  ' communication is asynchronous; many commands have no reply), so waiting for an answer could block the FIFO.
  ' Messages for the Z21 are therefore sent immediately; the FIFO is not used for UDP
  'waitForAck = True                                      ' We may not send a next message immediately
  If Parent.connectionType <> lanUdpPort Then waitForAck = True   ' We may not send a next message immediately
  Select Case Parent.connectionType                      ' Select output
    Case lanTcpPort:
      Parent.tcp.SendData byteArray
      lanKeepAliveTimer.Interval = lanKeepAliveInterval  ' Restart the timer, by resetting the Interval
      ' 2026-10 modified: the response timer is started below End Select, for TCP, USB and serial
      'tcpResponseTimer.Enabled = True                    ' This timer triggers if there is no response = Error
      'If Parent.ZentraleProgrammierMode Then             ' Responses in programming mode may take 90 sec
      '  tcpResponseTimer.Interval = tcpProgrammingTimeOut
      'Else
      '  tcpResponseTimer.Interval = tcpResponseTimeOut   ' In normal mode the response must be within 5 sec
      'End If
    Case lanUdpPort:
      Parent.udp.SendData byteArray
      lanKeepAliveTimer.Interval = lanKeepAliveInterval  ' Restart the timer, by resetting the Interval
    Case usbPort: Parent.usb.SendUsbData byteArray
    Case serialPort: Parent.usb.SendSerialData byteArray ' Serial and USB use the same class and object
  End Select
  ' 2026-10 modified: the response timer is used for TCP, USB and serial (was: TCP only). The USB interface has
  ' the same rule as LAN: no next command before the response, or before the timeout (5 sec, 1.5 min in
  ' programming mode) (LI-USB: 1.1, Interface 23151: 1.3). Serial interfaces may omit the confirmation 1/4/5
  ' (LI101F: 1.4). Without this timer, a missing response blocked the FIFO for USB and serial.
  If Parent.connectionType <> lanUdpPort Then
    tcpResponseTimer.Enabled = True                      ' This timer triggers if there is no response = Error
    If Parent.ZentraleProgrammierMode Then             ' Responses in programming mode may take 90 sec
      tcpResponseTimer.Interval = tcpProgrammingTimeOut
    Else
      tcpResponseTimer.Interval = tcpResponseTimeOut   ' In normal mode the response must be within 5 sec
    End If
  End If
'  If Parent.LogHexMessages Then Parent.Logfile.SendHex byteArray
  Parent.NumberOfExpressNetMessagesTransmitted = Parent.NumberOfExpressNetMessagesTransmitted + 1
End Sub

Private Sub lanKeepAliveTimer_Timer(ByVal Seconds As Currency)
  ' V3.6: 2.1 (Interface Status Befehl - Keep Alive)
  ' For LAN interface, send keep alive message, to inform the interface we are still listening
  ' This is required for the TCP (XpressNet) as well as the UDP (Z21) connection
  ' 2026-11 modified: not while the TCP connection is being established (the message would give the event SendError)
  If Not Parent.isConnected Then Exit Sub
 If Parent.connectionType = lanTcpPort Then
   waitForlanKeepAliveResponse = True
   Call Parent.CmdInterfaceStatus                       ' This short message is used as XpressNet "Keep Alive"
   'Parent.Logfile.Details "TCP Keep Alive Request"
   Parent.Logfile.DetailsSent "TCP Keep Alive Request"
  End If
 If Parent.connectionType = lanUdpPort Then
   waitForlanKeepAliveResponse = True
   Call Parent.CmdStatusZentraleAnfordern               ' This short message seems usable as Z21 "Keep Alive"
   'Parent.Logfile.Details "UDP Keep Alive Request"
   Parent.Logfile.DetailsSent "UDP Keep Alive Request"
  End If
End Sub

Private Sub tcpResponseTimer_Timer(ByVal Seconds As Currency)
  ' For the TCP interface: if the timer expires, something went wrong. Let's go to a well-defined state
  ' 2026-10 modified: for TCP, USB and serial (see transmit). The keep alive message is for TCP only
  ' (Interface 23151: chapter 2 is for the LAN interface only)
  'If Parent.connectionType = lanTcpPort Then
  If Parent.connectionType <> lanUdpPort Then
    ' 2026-11 modified: the description for the event SendError (before lastMessage changes), the number of removed
    ' messages, no retransmission anymore, and the activated switch of CmdAccessoryAuto is switched off first
    Dim description As String
    Dim removed As Integer
    Dim removedAccessories As Integer
    description = "No answer within " & (tcpResponseTimer.Interval \ 1000) & " sec (message: " & lastHex() & ")"
    removed = numberOfFifoMessages
    If heldMessageValid Then removed = removed + 1
    tcpResponseTimer.Enabled = False                      ' Stop the response timer
    Call emptyFifo                                        ' Remove all XpressNet messages still waiting
    heldMessageValid = False
    retryTimer.Enabled = False                            ' No retransmission is waiting
    retryPending = False
    retryCount = 0
    waitForAck = False                                    ' Clear the message acknowledgement flag
    ' Remove all switch commands still waiting. waitForAck is cleared before, so the deactivation of the activated
    ' switch is sent immediately
    'Call Parent.AccAuto.emptyFifo                         ' Remove all switch commands still waiting
    removedAccessories = Parent.AccAuto.emptyFifo
    'tcpResponseTimer.Enabled = False                      ' Stop the response timer
    'waitForlanKeepAliveResponse = True                    ' We will send a keep alive message
    'Call Parent.CmdInterfaceStatus                        ' We will send a keep alive message
    If Parent.connectionType = lanTcpPort Then
      waitForlanKeepAliveResponse = True                  ' We will send a keep alive message
      Call Parent.CmdInterfaceStatus                      ' We will send a keep alive message
    End If
    Parent.Logfile.Errors "Error: TCP Response Timer has timed out. Message queues emptied"
    description = description & ". Message queues emptied (" & removed & " messages, " & removedAccessories & " CmdAccessoryAuto commands)"
    Call Parent.eventSendError(Timeout, description)
  End If
End Sub

Public Sub stoptcpResponseTimer()
  ' Will be called from the receive class as soon as a response is received
  tcpResponseTimer.Enabled = False
End Sub

Public Function size() As Byte
  size = numberOfFifoMessages
End Function

Public Sub emptyFifo()
  head = 0
  tail = 0
  numberOfFifoMessages = 0
End Sub


' ----------------------------------------------------------------------------------------------------
' 2026-11 modified: new. Sends a Z21 LAN message that is not an XpressNet message (no tunnel header 0x40, no XOR),
' such as LAN_SET_BROADCASTFLAGS and LAN_LOGOFF. z21Message contains the complete message, including DataLen and
' Header. As for the XpressNet messages to the Z21: no flow control and no FIFO. The message is not counted in
' NumberOfExpressNetMessagesTransmitted, since it is not an XpressNet message
Public Sub transmitZ21(ByRef z21Message() As Byte)
  If (Parent.connectionType <> lanUdpPort) Or (Parent.udp Is Nothing) Then
    Call notSent(KeineVerbindung, "Message not sent: there is no Z21 (UDP) connection", z21Message)
    Exit Sub
  End If
  lanKeepAliveTimer.Interval = lanKeepAliveInterval      ' Restart the timer, by resetting the Interval
  Parent.udp.SendZ21 z21Message                          ' Logs the hex values as well
End Sub


' ----------------------------------------------------------------------------------------------------
' 2026-11 modified: new. RETRANSMISSION (see the header of this file)
' Called from the ReceiveClass if the interface answers the last message with an error that may disappear if the
' message is sent again: 61 80, 61 81, 01 01, 01 02, 01 03, 01 06 and 01 0A. answer is the text for the logfile
Public Sub answerError(ByVal answer As String)
  Dim delay As Long
  Dim description As String
  If Parent.connectionType = lanUdpPort Then Exit Sub    ' Z21: no flow control (see the header)
  If Not lastMessageValid Then Exit Sub                  ' No message has been sent yet
  If retryCount < maxRetries Then
    delay = firstRetryDelay * (2 ^ retryCount)           ' 100, 200, 400 msec
    retryCount = retryCount + 1
    retryPending = True
    waitForAck = True                                    ' The FIFO waits for the retransmission
    retryTimer.Interval = delay
    retryTimer.Enabled = True
    Parent.Logfile.Warnings "Answer " & answer & ": retransmission " & retryCount & " of " & maxRetries & " after " & delay & " msec (message: " & lastHex() & ")"
  Else
    retryCount = 0
    description = "Message given up after " & maxRetries & " retransmissions. Answer: " & answer & " (message: " & lastHex() & ")"
    Parent.Logfile.Errors description
    Call Parent.eventSendError(WiederholungErfolglos, description)
    ' waitForAck is False: the next message is sent by nextMessage (ReceiveClass.DecodeXpressNetMessage, STEP 4)
  End If
End Sub

' Called from the ReceiveClass if the interface answers the last message with 01 09 or 61 82. A retransmission
' does not help: the message is given up, and the next message is sent. reason: FehlerInDenBefehlsparametern (01 09)
' or BefehlNichtVorhanden (61 82)
Public Sub answerRejected(ByVal answer As String, ByVal reason As SendErrorReasonEnum)
  Dim description As String
  If Parent.connectionType = lanUdpPort Then Exit Sub    ' Z21: no flow control (see the header)
  If Not lastMessageValid Then Exit Sub                  ' No message has been sent yet
  retryCount = 0
  description = "Message rejected. Answer: " & answer & " (message: " & lastHex() & ")"
  Parent.Logfile.Errors description
  Call Parent.eventSendError(reason, description)
End Sub

Private Sub retryTimer_Timer(ByVal Seconds As Currency)
  retryTimer.Enabled = False                             ' The XTimerClass is periodic: we need a single tick
  If Not retryPending Then Exit Sub
  retryPending = False
  If interfaceNotAddressed Then                          ' A pause started in the meantime: send after the pause
    heldMessage = lastMessage
    heldMessageValid = True
    waitForAck = False                                   ' No message is underway
    Exit Sub
  End If
  Dim retry() As Byte
  retry = lastMessage                                    ' A copy: transmit assigns lastMessage, and VB6 does not allow
  transmit retry, True                                   ' that for an array that is passed ByRef (error 10)
End Sub


' ----------------------------------------------------------------------------------------------------
' 2026-11 modified: new. PAUSE (see the header of this file)
' Called from the ReceiveClass after 01 05 (answer = False) or 01 08 (answer = True). With 01 08 the last
' message did not reach the command station: it is held, and sent first after the pause
Public Sub pauseSending(ByVal answer As Boolean)
  If Parent.connectionType = lanUdpPort Then Exit Sub    ' Z21: no flow control (see the header)
  If answer And lastMessageValid Then
    heldMessage = lastMessage
    heldMessageValid = True
  End If
  If interfaceNotAddressed Then Exit Sub                 ' The pause has started already
  interfaceNotAddressed = True
  If Parent.ZentraleProgrammierMode Then                 ' In programming mode the pause may take 90 sec
    pauseTimer.Interval = pauseProgrammingTimeOut
  Else
    pauseTimer.Interval = pauseTimeOut
  End If
  pauseTimer.Enabled = True
  Parent.Logfile.Warnings "Sending paused, until the command station addresses the interface again"
  Call Parent.eventInterfaceNichtAdressiert
End Sub

' Called from the ReceiveClass after 01 07. The held message is sent first, then the FIFO and CmdAccessoryAuto.
' The event is raised last, so messages the application sends from the event handler come after these
Public Sub resumeSending()
  If Parent.connectionType = lanUdpPort Then Exit Sub    ' Z21: no flow control (see the header)
  If Not interfaceNotAddressed Then Exit Sub             ' 01 07 without a pause (for example after power on): ignored
  interfaceNotAddressed = False
  pauseTimer.Enabled = False
  Parent.Logfile.Warnings "Sending resumed: the command station addresses the interface again"
  Call nextMessage                                       ' First the held message, otherwise the FIFO
  Call Parent.AccAuto.resumeAfterPause                   ' CmdAccessoryAuto continues
  Call Parent.eventInterfaceWiederAdressiert
End Sub

Private Sub pauseTimer_Timer(ByVal Seconds As Currency)
  Dim seconds_ As Long
  Dim removed As Integer
  Dim removedAccessories As Integer
  Dim description As String
  pauseTimer.Enabled = False                             ' The XTimerClass is periodic: we need a single tick
  If Not interfaceNotAddressed Then Exit Sub
  seconds_ = pauseTimer.Interval \ 1000
  interfaceNotAddressed = False                          ' The pause ends
  removed = numberOfFifoMessages
  If heldMessageValid Then removed = removed + 1
  Call emptyFifo                                         ' Remove all XpressNet messages still waiting
  heldMessageValid = False
  retryCount = 0
  If retryPending Then                                   ' No retransmission anymore
    retryTimer.Enabled = False
    retryPending = False
    waitForAck = False
  End If
  ' Remove all switch commands still waiting. The activated switch (if any) is switched off first
  removedAccessories = Parent.AccAuto.emptyFifo
  description = "The command station did not address the interface again within " & seconds_ & " sec. Message queues emptied (" & removed & " messages, " & removedAccessories & " CmdAccessoryAuto commands)"
  Parent.Logfile.Errors description
  Call Parent.eventSendError(ZentraleAdressiertDasInterfaceNichtMehr, description)
End Sub


' ----------------------------------------------------------------------------------------------------
' 2026-11 modified: new. A message that is not sent: logfile and event SendError. The description contains the bytes
' of the message
Private Sub notSent(ByVal reason As SendErrorReasonEnum, ByVal text As String, ByRef msg() As Byte)
  Call notSentText(reason, text, bytesHex(msg))
End Sub

Private Sub notSentText(ByVal reason As SendErrorReasonEnum, ByVal text As String, ByVal hexText As String)
  ' The same log level as in version 1.3: a missing connection is a warning, the rest an error
  If reason = KeineVerbindung Then
    Parent.Logfile.Warnings text
  Else
    Parent.Logfile.Errors text
  End If
  Call Parent.eventSendError(reason, text & " (message: " & hexText & ")")
End Sub

' The bytes of a message as text, for example "52 00 88 DA"
Private Function bytesHex(ByRef msg() As Byte) As String
  Dim i As Integer
  Dim s As String
  s = ""
  For i = LBound(msg) To UBound(msg)
    If Len(s) > 0 Then s = s & " "
    s = s & Right$("0" & Hex$(msg(i)), 2)
  Next
  bytesHex = s
End Function

' The same for the parameters of message (a Variant that holds the ParamArray)
Private Function paramsHex(ByRef params As Variant) As String
  Dim i As Integer
  Dim s As String
  s = ""
  For i = LBound(params) To UBound(params)
    If Len(s) > 0 Then s = s & " "
    If (CLng(params(i)) >= 0) And (CLng(params(i)) <= 255) Then
      s = s & Right$("0" & Hex$(params(i)), 2)
    Else
      s = s & params(i)
    End If
  Next
  paramsHex = s
End Function

Private Function lastHex() As String
  If lastMessageValid Then
    lastHex = bytesHex(lastMessage)
  Else
    lastHex = ""
  End If
End Function










' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property



