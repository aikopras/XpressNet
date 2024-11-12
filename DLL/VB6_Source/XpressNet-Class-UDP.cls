VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UdpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' The purpose of this class is to create a wrapper around the UDP/Winsock control Form.
' Such wrapper is needed, since Forms must be declared as Private.
' This means that, besides the parent, other objects would be unable to use it.
' Therefore we need a wrapper class, which may be instantiated as Public, and is accessable
'
' This class additionally receives the UDP packets, and calls the XpressNet receive object
' to further handle the received XpressNet messages

Option Explicit

Private udpForm As FrmUdp        ' Declare a new Form that includes a Socket control (udpClient)

' The lines below are needed to pass a pointer of the main object to this subclass.
' This is needed, since we need to call some of the methods of the master and other objects
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long

Public Sub Class_Initialize()
  Set udpForm = New FrmUdp                            ' Instantiate the form that holds the WinSock Control (udpClient)
  Set udpForm.Parent = Me                             ' Provide the form with a (weak) reference to me
  udpForm.Visible = False
End Sub

Private Sub Class_Terminate()
  Set udpForm = Nothing
End Sub

Public Sub init(ByVal Host As String, ByVal Port As String)
  udpForm.udpClient.Close        ' Not sure if this is useful
  udpForm.udpClient.RemoteHost = Host
  udpForm.udpClient.RemotePort = Port
  ' udpForm.udpClient.Bind 21106   ' Preferred port, according to the Z21 Spec
  ' The problem with a fixed local port, is that we can only instantiate
  ' a single instance of the UDP class. The spec seems ambigious, however.
  udpForm.udpClient.Bind 0       ' Let Windows select a random free port
End Sub
                         
Public Sub SendData(ByRef inBuffer() As Byte)
  Dim LengthInBuffer As Byte
  Dim LengthOutBuffer As Byte
  ' STEP 1: Determine the length of the in/outbuffer
  LengthInBuffer = UBound(inBuffer)
  LengthOutBuffer = LengthInBuffer + 4
  Dim outBuffer() As Byte                   ' Create a new array, to hold all Z21 bytes
  ReDim outBuffer(LengthOutBuffer)          ' This buffer has leading Z21 bytes
  ' STEP 2: add the four bytes for the XpressNet tunnel at the start of the message
  outBuffer(0) = LengthOutBuffer + 1        ' DataLen
  outBuffer(1) = 0
  outBuffer(2) = &H40                       ' Z21 Header byte
  outBuffer(3) = 0
  ' STEP 3: copy the message from inBuffer to Outbuffer
  Dim i As Byte
  For i = 0 To LengthInBuffer               ' Neglect the Xpressnet
    outBuffer(i + 4) = inBuffer(i)          ' Copy all bytes received to the outBuffer
  Next
  ' STEP 4: Log the message in HEX format, if requested
  If Parent.LogHexMessages Then Parent.Logfile.SendHex outBuffer
  ' STEP 5: Send the outbuffer
  udpForm.udpClient.SendData outBuffer
End Sub

' Note: the Yamorc YD7001D retransmits broadcast messages three times
' For each instance of the UDP class we therefore receive three messages
Public Sub receiveUdpPacket()
  Dim udpData() As Byte                                ' Array to store the UDP message (data)
  udpForm.udpClient.GetData udpData                    ' Get the data from the UDP socket
  Call Parent.receiveZ21.addToIndata(udpData)
End Sub




' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property


