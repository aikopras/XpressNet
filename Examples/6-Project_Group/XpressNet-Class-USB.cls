VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsbClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' The purpose of this code is to create a wrapper around the USB/MSCOMM  control Form.
' Such wrapper is needed, since Forms must be declared as Private.
' This means that, besides the parent, other objects would be unable to use it.
' Therefore we need a wrapper class, which may be instantiated as Public, and is accessable

' Note that the USB Class can also be used for traditional RS-232 (Serial) interfaces.

' Info regarding Lenz interfaces:
' There are two types of interfaces that can use the MSCOMM control:
'
' 1) Serial:
' - LI100: First design. Metal case. Baudrate: 9600.
' - LI100F: Improved design. Metal case. Baudrates: 9600 (default) and 19200 baud.
' - LI101F: Last version of the serial interface. Plastic case. Galvanic isolation
'           Baudrates: 19200 (default), 38400, 57600, en 115200 baud.
' All serial interfaces use:
' - 8 Databits, 1 Start bit, 1 Stop bit, No parity => .Settings = "baudrate,N,8,1"
' - Hardware-Handshake: the interface uses CTS (Clear To Send) to tell the PC whether or not
'   it is able to receive data. The interface ignores the PC's RTS (Request to Send) signal.
'   => .Handshaking = comRTS
'   Note that the Interface's DSR (Data Set Ready) and the PC's DTR (Data terminal Ready)
'   are not used.
'
' 2) USB
' - LI-USB (23150)
' - LI-USB-Ethernet (23151)
' - LZV200
' - Yamorc YD7001
' All Lenz USB interfaces use:
' - 57600 Baud, 8 Databits, 1 Start bit, 1 Stop bit, No parity => .Settings = "57600,N,8,1"
' - No hardware handshake +> .Handshaking = ComNon
' The Yamorc YD7001 USB interface seems to work with 57600, as well as 115200. Somewhere on the
' Yamorc website it is written that the YD7001 works with 400Kbps. Maybe the baudrate setting is ignored.
'
' Receiving XpressNet messages via USB
' ====================================
' After a number of tests, which included the use of Wireshark to monitor the USB interface, the following
' conclusions were drown:
' - multiple XpressNet messages may be received when the MSComm control raises the receive event.
' - If a MSComm receive event is raised, only parts of the XpressNet message may have been received.
' => We therefore can not rely on a single "get all USB receive buffer data" call to receive the complete
'    XpressNet message
' => We need to read individual bytes from the USB Interface, determine the XpressNet header byte,
'    retrieve the length of the XpressNet message from (the lower 4 bits) from this header byte, and
'    continue reading bytes, until the number of received bytes match the length as indicated in the header.
'==========================================================================================================
Option Explicit

Private usbForm As FrmUsb        ' Declare a new Form that includes a MSComm control (tcpClient)

' The lines below are needed to pass a pointer of the main object to this subclass.
' This is needed, since we need to call some of the methods of the master and other objects
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long


Public Sub Class_Initialize()
   Set usbForm = New FrmUsb                ' Instantiate the form that holds the MSComm Control (usbClient)
   Set usbForm.Parent = Me                 ' Provide the form with a (weak) reference to me
   usbForm.Visible = False
   With usbForm.usbClient                  ' Initialise the MSComm control with settings that never change
     .InputMode = comInputModeBinary       ' Xpressnet messages are binary encoded
     .EOFEnable = False                    ' Therefore don't check for EOF (^z?) characters
     .NullDiscard = False                  ' And don't discard the binary value 0
     .InBufferSize = 1024                  ' Since we have application level flow control, the default is OK
     .OutBufferSize = 512                  ' Same here
     .DTREnable = False                    ' DSR/DTR is not implemented by Lenz. Keep the default
     .RTSEnable = False                    ' The Lenz XpressNet spec says it ignores RTS
     .InputLen = 0                         ' Read the entire contents of the receive buffer
     .SThreshold = 0                       ' Disable generating the OnComm event for data transmission events
     ' Setting the RThreshold property to 0 (the default)
     ' disables generating the OnComm event when characters are received.
     ' Setting RThreshold to 1, for example, causes the MSComm control to
     ' generate the OnComm event every time a single character is placed in the receive buffer.
     .RThreshold = 1                       ' Default = 0
   End With
End Sub

Private Sub Class_Terminate()
  If usbForm.usbClient.PortOpen Then usbForm.usbClient.PortOpen = False
  Set usbForm = Nothing
End Sub

Public Sub connectUsb(ByVal Port As Byte, ByRef Settings As String)
  ' Lenz: 57600,N,8,1 / Yamorc: 115200,N,8,1
  Call openPort(Port, Settings, False)
End Sub

Public Sub connectSerial(ByVal Port As Byte, ByRef Settings As String, ByVal Handshake As Boolean)
  Call openPort(Port, Settings, Handshake)
End Sub

Public Sub openPort(ByVal portNumber As Byte, ByRef portSettings As String, ByVal Handshake As Boolean)
  With usbForm.usbClient
    On Error GoTo PortError
    .CommPort = portNumber
    .Settings = portSettings
    If Handshake Then
      .Handshaking = comRTS               ' Serial uses CTS for hardware handshaking
    Else
      .Handshaking = comNone              ' No hardware handshaking
    End If
    ' Try to open the USB port
    If Not .PortOpen Then .PortOpen = True
    If .PortOpen = True Then Parent.eventInterfaceConnected
  End With
  Exit Sub
' Catch errors if the port does not exist
PortError:
  MsgBox "COM port:" & usbForm.usbClient.CommPort & vbCrLf & " Error: " & Err.number & " : " & Err.Description
  Resume Next
End Sub

Public Sub closePort()
  If usbForm.usbClient.PortOpen Then usbForm.usbClient.PortOpen = False
End Sub

Public Sub SendSerialData(ByRef inBuffer() As Byte)
  ' STEP 4: Log the message in HEX format, if requested
  If Parent.LogHexMessages Then Parent.Logfile.SendHex inBuffer
  ' STEP 5: Send the outbuffer
  usbForm.usbClient.Output = inBuffer
End Sub

Public Sub SendUsbData(ByRef inBuffer() As Byte)
  Dim LengthInBuffer As Byte
  Dim LengthOutBuffer As Byte
  ' STEP 1: Determine the length of the in/outbuffer
  LengthInBuffer = UBound(inBuffer)
  LengthOutBuffer = LengthInBuffer + 2      ' +2, for the two Framing bytes
  Dim outBuffer() As Byte                   ' Create a new array, to hold all XpressNet bytes
  ReDim outBuffer(LengthOutBuffer)          ' This buffer has leading Framing bytes
  ' STEP 2: add two Framing Bytes at the start of the message
  outBuffer(0) = &HFF
  outBuffer(1) = &HFE                       ' Z21 Header byte
  ' STEP 3: copy the message from inBuffer to Outbuffer
  Dim i As Byte
  For i = 0 To LengthInBuffer
    outBuffer(i + 2) = inBuffer(i)          ' Copy all bytes received to the outBuffer
  Next
  ' STEP 4: Log the message in HEX format, if requested
  If Parent.LogHexMessages Then Parent.Logfile.SendHex outBuffer
  ' STEP 5: Send the outbuffer
  usbForm.usbClient.Output = outBuffer
End Sub

Private Sub saveUsbBytesToInDataBuffer()
  Dim bytesReceived As Long                  ' number of received bytes in one event
  Dim usbData() As Byte                      ' local array to temporary copy the USB input buffer
  ' STEP 1: Get the input bytes from the MSComm driver
  bytesReceived = usbForm.usbClient.InBufferCount
  usbData = usbForm.usbClient.Input
  ' STEP 2: What we have received may be a part of an XpressNet message, a complete
  ' XpressNet message but also (parts of) multiple XpressNet messages.
  ' We call "addToIndata", which stores all received bytes into the inBuffer
  Call Parent.receive.addToIndata(bytesReceived, usbData)
End Sub

Public Sub handleUsbEvent()
  ' This routine is called whenever a USB Event is received via the MSComm interface.
  ' If we received USB data, fetch and save that data
  If usbForm.usbClient.CommEvent = comEvReceive Then Call saveUsbBytesToInDataBuffer
  ' The other events are not foreseen. We therefore save them in the logfile, for later debugging
  With usbForm.usbClient
    Select Case .CommEvent
        'Events
        Case comEvSend
          ' Parent.Logfile.AppendWithHost "COM Event: comEvSend - There are SThreshold number of characters in the transmit buffer"
        Case comEvCTS
          ' Parent.Logfile.AppendWithHost "COM Event: comEvCTS - Change in the CTS line"
        Case comEvCD
          ' Parent.Logfile.AppendWithHost "COM Event: comEvCD - Change in the CD line"
        Case comEvDSR
          ' Parent.Logfile.AppendWithHost "COM Event: comEvDSR - Change in the DSR line"
        Case comEvRing
          ' Parent.Logfile.AppendWithHost "COM Event: comEvRing - Change in the Ring Indicator"
        Case comEvEOF
          ' Parent.Logfile.AppendWithHost "COM Event: comEvEOF - An EOF charater was found in the input stream"
        'Errors
        Case comEventBreak
           Parent.Logfile.AppendWithHost "COM Error: comEventBreak - A Break was received"
        Case comEventCDTO
          Parent.Logfile.AppendWithHost "COM Error: comEventCDTO"
        Case comEventCTSTO
          Parent.Logfile.AppendWithHost "COM Error: CTS Timeout"
        Case comEventDSRTO
          Parent.Logfile.AppendWithHost "COM Error: DSR Timeout"
        Case comEventFrame
          Parent.Logfile.AppendWithHost "COM Error: Framing Error"
        Case comEventOverrun
          Parent.Logfile.AppendWithHost "COM Error: Overrun: Data Lost"
        Case comEventRxOver
          Parent.Logfile.AppendWithHost "COM Error: RX Overflow - Receive buffer overflow"
        Case comEventRxParity
          Parent.Logfile.AppendWithHost "COM Error: Parity Error"
        Case comEventTxFull
          Parent.Logfile.AppendWithHost "COM Error: TX buffer full"
        Case comEventDCB
          Parent.Logfile.AppendWithHost "COM Error: Unexpected error retrieving DCB"
    End Select
  End With
End Sub



' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property





