VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReceiveClassZ21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This class implements the Z21 receive code for XpressNet messages that are transferred "within"
' Z21 messages (the Z21 Xpressnet tunnel).
' The first 4 bytes of these massages are: DataLen 0x00 0x40 0x00.
' The remaining bytes are the XpressNet message.
' Note that the last four bits of the XpressNet header do not always include the correct
' XpressNet message length. Therefore we must use the Z21 DataLen
' For Z21 details, see the Z21 LAN Protocol Specification (Version 1.13, 06.11.2023)

' Although this code is very similar to the "normal" receive class code, there are too many
' subtle differences to add it to that "normal" code.

' In principle it wouldn't be too difficult to write a complete Z21 Interface, but since the main
' goal of this library is to make a VB6 interface for Accessory commands and RS-Bus feedbacks,
' we refrain from that.

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long

' ================================================================================================
' The buffer below is used to store the input data received from the UDP socket.
' According to section 1.3 of the Z21 protocol spec, multiple Z21 messages may be combined
' within a single UDP packet.
' Although the Z21 specification doesn't say anything regarding retransmission, it seems
' that at least the Yamorc YD7001D command station retransmits broadcast messages.
' Probably that is done since UDP is, by nature, unreliable.

Private Const inDataSize = 1478      ' Maximum size of the UDP Data (Section 1.3 Z21 spec)
Private inData(inDataSize) As Byte   ' Global Array that is used as buffer to store received bytes
Private indexNextInByte As Long      ' Index into the indata Array. Empty space to store subsequent byte

Public Sub addToIndata(ByRef inBytes() As Byte)
  Dim i As Long                              ' used in For loops
  Dim bytesReceived As Long
  bytesReceived = UBound(inBytes)
  For i = 0 To UBound(inBytes)             ' include all bytes that we received as input
    If indexNextInByte > inDataSize Then   ' avoid buffer overflow
      Parent.Logfile.Errors "Error: overflow Z21 receive input buffer"
      Exit Sub
    End If
    inData(indexNextInByte) = inBytes(i)
    indexNextInByte = indexNextInByte + 1
  Next
  ' Now that new data is stored, check if the inBuffer contains one or more
  ' Z21 messages. If yes, process this / these message(s)
  Call checkInDataForZ21messages
End Sub

Private Sub checkInDataForZ21messages()
  ' Check if inData contains one or more complete Z21 messages
  Dim Z21length As Long                      ' the length of the Z21 message (from the first byte)
  Dim Z21Message(22) As Byte                 ' array to store the Z21 Message (plus a 4 byte header)
  Dim i As Long                              ' used in For loops
  Dim j As Long                              ' used in For loops
  ' Since we may have multiple Z21/XpressNet message within the inData buffer, we have a While loop
  ' that ends whenever the inData buffer is empty (indexnextInByte points to 0).
  ' We use "Exit Sub" to leave the While loop, whenever the inData buffer contains "garbage")
  While indexNextInByte > 0
    ' Step 1: Determine the length of the Z21 and XpressNet messages
    Z21length = inData(0)
    ' Step 2: Check if we have at least this number of bytes in the indata buffer
    ' If not, we have something inconsistent in the inData buffer
    ' In such case we raise an error, clear the indata buffer and exit the subroutine
    If Z21length > indexNextInByte Then
      Parent.Logfile.Errors "Error: Z21 messagesize exceeds the number of received UDP bytes"
      indexNextInByte = 0                    ' clear the inBuffer
      Exit Sub                               ' and return
    End If
   ' STEP 3: Copy the Z21 message in a new buffer
    For i = 0 To (Z21length - 1)
      Z21Message(i) = inData(i)
    Next
    ' STEP 4: Remove the z21 message from the inData buffer, and shift the remainder to the begin
    i = 0
    For j = Z21length To (indexNextInByte - 1)
      inData(i) = inData(j)
      i = i + 1
    Next
    indexNextInByte = i                       ' points to the next empty byte
    ' STEP 5: Call the subroutine that analyses the individual Z21 message
    Call DecodeZ21Message(Z21length, Z21Message)
  Wend
End Sub

Private Sub DecodeZ21Message(ByVal Z21length As Long, ByRef Z21Message() As Byte)
  Dim indexHeader As Long                    ' points to the XpressNet header
  Dim indexXor As Long                       ' points to the last byte of the XpressNet message (=XOR byte)
  Dim Xpnlength As Byte                      ' the length of the XpressNet message (from the header byte)
  Dim XpnMessage(18) As Byte                 ' array to store the XpressNet Message
  Dim xorValue As Byte                       ' to calculate the XOR over all bytes in the XpressNet message
  Dim broadcast As Boolean                   ' is the received message a broadcast? (FF FD)
  Dim i As Long                              ' used in For loops
  Dim j As Long                              ' used in For loops
  ' STEP 1: Exit if the Z21Message does not contain an XpressNet message
  If Z21Message(2) <> &H40 Then Exit Sub
  ' STEP2: Retrieve the length of the XpressNet message from the first Z21 byte
  Xpnlength = inData(0) - 5                  ' Xpnlength is without the XOR byte
  indexXor = inData(0) - 1                   ' Since the first index is 0, we should subtract 1
  ' STEP 3: Should we log the XpressNet packet?
  If Parent.LogHexMessages Then Parent.Logfile.ReceiveHexRange Z21Message(), 0, indexXor
  ' STEP 4: Check if the XOR value is correct
  xorValue = 0                               ' Initialise
  For i = 4 To indexXor                      ' If we include the XOR byte in the XOR calculation
    xorValue = xorValue Xor inData(i)        ' the result should become 0
  Next
  If xorValue Then                           ' XOR error. Exit sub
    Parent.Logfile.Errors "XOR Error in received XpressNet message"
    If Parent.LogErrors Then Parent.Logfile.ReceiveHexRange Z21Message(), 0, indexXor
  End If
  ' NOTE: in the "normal receive routine we check if the received message is a broadcast message
  ' We could do that here as well, but since there is no flow control, knowing if something is a
  ' broadcast message doesn;t bring anything
  ' STEP 5: Copy the XpressNet message into the xpnMessage buffer
   For i = 0 To (indexHeader + Xpnlength)
     XpnMessage(i) = inData(i + 4)
   Next
  ' STEP 6: Call the subroutine that analyses the individual XpressNet message
  Call Parent.receive.DecodeXpressNetMessage(Xpnlength, XpnMessage, 0)
End Sub


Private Sub Class_Initialize()
   indexNextInByte = 0                    ' the first empty space in the inData array.
End Sub




'--------------------------------------------------------------------------------------------
' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property








