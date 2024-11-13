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
Private Const maxFifoEntries = 32                 ' The number of XpressNet messages the queue is able to hold
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



Private Sub Class_Initialize()
  waitForAck = False                              ' The start value is that sending data is allowed
  Call emptyFifo
  Set lanKeepAliveTimer = New XTimerClass         ' Create a new timer object for TCP keep alive messages
  lanKeepAliveTimer.Interval = lanKeepAliveInterval
  lanKeepAliveTimer.Enabled = True
  waitForlanKeepAliveResponse = False             ' We did not yet send a LAN Keep Alive message
  Set tcpResponseTimer = New XTimerClass          ' Create a new timer object for TCP response time outs
  tcpResponseTimer.Enabled = False                ' The timer will be initialised as soon as a message is send
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
Public Sub message(ParamArray byteSequence() As Variant)
  Dim lengthParamArray As Byte    ' Number of bytes we received from the calling procedure
  Dim i As Byte                   ' will be used in For loops
  ' STEP 1: Check if we have already a connection
  If Parent.connectionType = none Then Exit Sub        ' Ignore everything. Drop the packet
  ' STEP 2: Perform 3 checks on the size of the received message.
  ' Check 1: did we receive any input??
  If IsMissing(byteSequence) Then
    Parent.Logfile.Errors "Error in the message subroutine: ParamArray is empty"
    Exit Sub
  End If
  ' Check 2: Is the input not too long?
  ' The check is to protect the FIFO against out of range (runtime) errors. The check is whether the
  ' received bytes fit in an FIFO row; the check is NOT if the received bytes fit in an XpressNet message.
  ' Note that UBound does not return the number of bytes received, but 1 less, since the
  ' ParamArray is indexed from 0 .. number - 1
  lengthParamArray = UBound(byteSequence)                 ' length = # parameters - 1!!
  If lengthParamArray >= (maxFifoMsgSize - 1) Then        ' -1, to have space for the XOR byte
    Parent.Logfile.Errors "Error in the message subroutine: ParamArray exceeds maxFifoMsgSize"
    Exit Sub
  End If
  ' Check 3: Does the received number of bytes correspond to the XpressNet message header?
  ' The number of bytes received must match with the length as contained in (the lower 4 bits) of the first
  ' (=header) byte. Note: the Z21 Specification sometimes violates this, and must be excluded from this check
  If (Not Parent.Z21Protocol) And (lengthParamArray <> (byteSequence(0) Mod 16)) Then
    Parent.Logfile.Errors "Error in the message subroutine: size ParamArray doesn't match size in header"
    Exit Sub
  End If
  ' STEP 3: Calculate the XOR value
  Dim xorValue As Byte                       ' The result will be stored in this byte
  xorValue = 0
  For i = 0 To lengthParamArray              ' include all bytes that we received as input
    xorValue = xorValue Xor byteSequence(i)
  Next
  ' STEP 4: Check if we can send immediately, or if the message should be stored in the FIFO
  If waitForAck = False Then
    ' STEP A1: Create an array that holds all XpressNet bytes. The Array size should match the number of bytes
    Dim outMsgSize As Byte                    ' msgSize will hold the number of bytes in the message
    outMsgSize = lengthParamArray + 1         ' Number of bytes for the complete XpressNet message: add 1 for the XOR
    Dim outBuffer() As Byte                   ' Create a new array, to hold all XpressNet bytes
    ReDim outBuffer(outMsgSize)               ' Since this array starts at 0, its size should be msgSize - 1
    For i = 0 To lengthParamArray
      outBuffer(i) = byteSequence(i)        ' Copy all bytes this routine received to the outBuffer
    Next
    outBuffer(lengthParamArray + 1) = xorValue
    ' STEP A2 : transmit the message over the wire
    transmit outBuffer
  Else                                          ' Store the message in aFIFIO
    ' STEP B1: Check if the FIFO queue has free slots left to store the message
    ' If not, increase "maxFifoEntries", or modify the calling program
    If numberOfFifoMessages >= maxFifoEntries Then  ' FIFO overflow :-(
      Parent.Logfile.Errors "Error: FIFO queue is full"
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
    For i = 0 To lengthParamArray            ' Copy the bytes that this subroutine received as input
      fifo(tail, i + 1) = byteSequence(i)
    Next i
    fifo(tail, lengthParamArray + 2) = xorValue
  End If
End Sub


' ----------------------------------------------------------------------------------------------------
' nextMessage will be called from the ReceiveClass code. Provided that we are allowed to send now,
' it sends a message from the FIFO queue, if there are still messages waiting.
Public Sub nextMessage()
  Dim i As Byte                               ' will be used in For loops
  If waitForAck = False Then                  ' we may send (no flow control)
    ' Check if the FIFO has messages waiting. If yes, send the oldest message from the FIFO
    If numberOfFifoMessages > 0 Then            ' there are still XpressNet messages waiting in the FIFO
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

Public Sub transmit(ByRef byteArray() As Byte)
  waitForAck = True                                      ' We may not send a next message immediately
  Select Case Parent.connectionType                      ' Select output
    Case lanTcpPort:
      Parent.tcp.SendData byteArray
      lanKeepAliveTimer.Interval = lanKeepAliveInterval  ' Restart the timer, by resetting the Interval
      tcpResponseTimer.Enabled = True                    ' This timer triggers if there is no response = Error
      If Parent.ZentraleProgrammierMode Then             ' Responses in programming mode may take 90 sec
        tcpResponseTimer.Interval = tcpProgrammingTimeOut
      Else
        tcpResponseTimer.Interval = tcpResponseTimeOut   ' In normal mode the response must be within 5 sec
      End If
    Case lanUdpPort:
      Parent.udp.SendData byteArray
      lanKeepAliveTimer.Interval = lanKeepAliveInterval  ' Restart the timer, by resetting the Interval
    Case usbPort: Parent.usb.SendUsbData byteArray
    Case serialPort: Parent.usb.SendSerialData byteArray ' Serial and USB use the same class and object
  End Select
'  If Parent.LogHexMessages Then Parent.Logfile.SendHex byteArray
  Parent.NumberOfExpressNetMessagesTransmitted = Parent.NumberOfExpressNetMessagesTransmitted + 1
End Sub

Private Sub lanKeepAliveTimer_Timer(ByVal Seconds As Currency)
  ' V3.6: 2.1 (Interface Status Befehl - Keep Alive)
  ' For LAN interface, send keep alive message, to inform the interface we are still listening
  ' This is required for the TCP (XpressNet) as well as the UDP (Z21) connection
 If Parent.connectionType = lanTcpPort Then
   waitForlanKeepAliveResponse = True
   Call Parent.CmdInterfaceStatus                       ' This short message is used as XpressNet "Keep Alive"
   Parent.Logfile.Details "TCP Keep Alive Request"
  End If
 If Parent.connectionType = lanUdpPort Then
   waitForlanKeepAliveResponse = True
   Call Parent.CmdStatusZentraleAnfordern               ' This short message seems usable as Z21 "Keep Alive"
   Parent.Logfile.Details "UDP Keep Alive Request"
  End If
End Sub

Private Sub tcpResponseTimer_Timer(ByVal Seconds As Currency)
  ' For the TCP interface: if the timer expires, something went wrong. Let's go to a well-defined state
  If Parent.connectionType = lanTcpPort Then
    Call emptyFifo                                        ' Remove all XpressNet messages still waiting
    Call Parent.AccAuto.emptyFifo                         ' Remove all switch commands still waiting
    waitForAck = False                                    ' Clear the message acknowledgement flag
    tcpResponseTimer.Enabled = False                      ' Stop the response timer
    waitForlanKeepAliveResponse = True                    ' We will send a keep alive message
    Call Parent.CmdInterfaceStatus                        ' We will send a keep alive message
    Parent.Logfile.Errors "Error: TCP Response Timer has timed out. Message queues emptied"
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










' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property



