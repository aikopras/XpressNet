VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TcpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' The purpose of this class is to create a wrapper around the TCP/Winsock control Form.
' Such wrapper is needed, since Forms must be declared as Private.
' This means that, besides the parent, other objects would be unable to use it.
' Therefore we need a wrapper class, which may be instantiated as Public, and is accessable
'
' This class additionally receives the TCP segments, and split these segments into
' individual XpressNet messages, and calls the XpressNet receive object to further
' handle the received XpressNet messages
'
' 2026-11 modified (version 2.0): a connection attempt is stopped after 5 seconds (Windows itself gives up after about
' 21 seconds, for example after a typing error in the IP address). The events of the Winsock control are passed via
' this class (onConnect, onError), so the timer can be stopped, and an error after the time-out is ignored.

Option Explicit

Private tcpForm As FrmTcp        'Declare a new Form that includes a Socket control (tcpClient)

' 2026-11 modified: the time-out of a connection attempt
Private Const connectTimeOut = 5000             ' Maximum time for a connection attempt (msec)
Private WithEvents connectTimer As XTimerClass
Attribute connectTimer.VB_VarHelpID = -1
Private closed As Boolean                       ' closeConnection has been called, or the attempt has timed out

' The lines below are needed to pass a pointer of the main object to this subclass.
' This is needed, since we need to call some of the methods of the master and other objects
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long


Public Sub Class_Initialize()
   Set tcpForm = New FrmTcp                            ' Instantiate the form that holds the WinSock Control (tcpClient)
   Set tcpForm.Parent = Me                             ' Provide the form with a (weak) reference to me
   tcpForm.Visible = False
   Set connectTimer = New XTimerClass                  ' 2026-11 modified
   connectTimer.Enabled = False
   closed = False
End Sub

Private Sub Class_Terminate()
  connectTimer.Enabled = False                        ' 2026-11 modified
  tcpForm.tcpClient.Close
  Set tcpForm = Nothing
End Sub

Public Sub connect(ByVal Host As String, ByVal Port As String)
  ' 2026-11 modified: an error in the port number (for example: text) gives the event TcpError (was: run time error)
  On Error GoTo connectError
  tcpForm.tcpClient.RemoteHost = Host
  tcpForm.tcpClient.RemotePort = Port
  tcpForm.tcpClient.connect
  On Error GoTo 0
  connectTimer.Interval = connectTimeOut              ' 2026-11 modified: start the time-out
  connectTimer.Enabled = True
  Exit Sub
connectError:
  Dim description As String
  description = Err.Description
  Err.Clear
  Call Parent.eventTcpError(description)
End Sub

Public Sub closeConnection()
  closed = True                                       ' 2026-11 modified
  connectTimer.Enabled = False
  tcpForm.tcpClient.Close
End Sub

' 2026-11 modified: new. Is the connection established?
Public Function connected() As Boolean
  connected = (tcpForm.tcpClient.State = sckConnected)
End Function

' 2026-11 modified: new. The connection attempt takes too long
Private Sub connectTimer_Timer(ByVal Seconds As Currency)
  Dim description As String
  connectTimer.Enabled = False
  If closed Or (tcpForm.tcpClient.State = sckConnected) Then Exit Sub
  closed = True                                       ' The result of the connection attempt is ignored
  tcpForm.tcpClient.Close
  description = "Connection timed out"
  Call Parent.eventTcpError(description, "Connection attempt timed out after " & (connectTimeOut \ 1000) & " seconds")
End Sub

' 2026-11 modified: new. Called by FrmTcp (tcpClient_Connect): the connection is established
Public Sub onConnect()
  connectTimer.Enabled = False
  If closed Then Exit Sub
  ' Which connection belongs to the letter of this XpressNet object
  With tcpForm.tcpClient
    Parent.Logfile.Connection .RemoteHost & ", port " & .RemotePort & ", local port " & .LocalPort
  End With
  Parent.eventInterfaceConnected
End Sub

' 2026-11 modified: new. Called by FrmTcp (tcpClient_Error)
Public Sub onError(ByVal description As String)
  connectTimer.Enabled = False
  If closed Then Exit Sub
  Call Parent.eventTcpError(description)
End Sub

Public Sub SendData(ByRef inBuffer() As Byte)
  Dim LengthInBuffer As Byte
  Dim LengthOutBuffer As Byte
  ' STEP 1: Do we still have the TCP connection?
  If tcpForm.tcpClient.State <> sckConnected Then      ' Do we still have the connection??
    Parent.eventTcpClosedByRemoteHost                  ' Clean up, and inform the application
    Exit Sub
  End If
  ' STEP 2: Determine the length of the in/outbuffer
  LengthInBuffer = UBound(inBuffer)
  LengthOutBuffer = LengthInBuffer + 2      ' +2, for the two Framing bytes
  Dim outBuffer() As Byte                   ' Create a new array, to hold all XpressNet bytes
  ReDim outBuffer(LengthOutBuffer)          ' This buffer has leading Framing bytes
  ' STEP 3: add two Framing Bytes at the start of the message
  outBuffer(0) = &HFF
  outBuffer(1) = &HFE                       ' Z21 Header byte
  ' STEP 4: copy the message from inBuffer to Outbuffer
  Dim i As Byte
  For i = 0 To LengthInBuffer
    outBuffer(i + 2) = inBuffer(i)          ' Copy all bytes received to the outBuffer
  Next
  ' STEP 5: Log the message in HEX format, if requested
  If Parent.LogHexMessages Then Parent.Logfile.SendHex outBuffer
  ' STEP 6: Send the outbuffer
  tcpForm.tcpClient.SendData outBuffer
End Sub

Public Sub receiveTcpSegment(ByVal bytesReceived As Long)
  Dim tcpData() As Byte                                ' Array to store the TCP Segment (data)
  tcpForm.tcpClient.GetData tcpData                    ' Get the data from the TCP socket
  'Call Parent.receive.DecodeInMessage(bytesTotal, tcpData)
  Call Parent.receive.addToIndata(bytesReceived, tcpData)
End Sub



' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property

