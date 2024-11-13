VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReceiveClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The various subroutines contain references to the V3.6 and V4 XpressNet specifications:
' - Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151
'   Kommunikation auf den Schnittstellen, Doku Version 1.2 05/2019
'   https://www.lenz-elektronik.de/src/pdf/23151_Info.zip
' - Dokumentation XpressNet Version 4.0, 02/2022:
'   https://www.lenz-elektronik.de/src/pdf/Lenz_XpressNet_Doku.pdf
' For details, see these documents

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long

' ================================================================================================
' The buffer below is used to store "raw" input data from the LAN/USB/Serial interfaces
' Everytime bytes are received from a physical (LAN/USB/Serial) interface, the event handler
' for that interface calls addToInData to store the received bytes within inBuffer
Private Const inDataSize = 1024
Private inData(inDataSize) As Byte       ' Global Array that is used as buffer to store received bytes
Private indexNextInByte As Long   ' Index into the indata Array. Empty space to store subsequent byte

Public Sub addToIndata(ByVal bytesReceived As Long, ByRef inBytes() As Byte)
  Dim i As Long                              ' used in For loops
  For i = 0 To bytesReceived - 1             ' include all bytes that we received as input
    ' Some command stations (like the YD7001) may generate quite some data
    If indexNextInByte > inDataSize Then
      Parent.Logfile.Errors "Error: overflow receive input buffer"
      Exit Sub
    End If
    inData(indexNextInByte) = inBytes(i)
    indexNextInByte = indexNextInByte + 1
  Next
  ' Now that new data is stored, check if the inBuffer contains one or more
  ' complete XpressNet messages. If yes, process this / these message(s)
  Call checkInDataForXpressNetmessages
End Sub

Private Sub checkInDataForXpressNetmessages()
  ' Check if inData contains one or more complete XpressNet messages
  Dim indexHeader As Long                    ' points to the XpressNet header
  Dim indexXor As Long                       ' points to the last byte of the XpressNet message (=XOR byte)
  Dim Xpnlength As Byte                      ' the length of the XpressNet message (from the header byte)
  Dim XpnMessage(18) As Byte                 ' array to store the XpressNet Message
  Dim xorValue As Byte                       ' to calculate the XOR over all bytes in the XpressNet message
  Dim broadcast As Boolean                   ' is the received message a broadcast? (FF FD)
  Dim i As Long                              ' used in For loops
  Dim j As Long                              ' used in For loops
  ' Since we may have multiple XpressNet message within the inData buffer, we have a While loop
  ' that doesn't end. We use "Exit Sub" to leave the While loop, whenever the inData buffer is
  ' empty or contains an incomplete XpressNet message
  While True
    ' Step 1: Determine where the XpressNet header is located
    If inData(0) = &HFF Then
      indexHeader = 2
    Else
      indexHeader = 0
    End If
    ' Exit, if the header byte has not yet been received
    If indexNextInByte <= indexHeader Then Exit Sub
    ' STEP2: Retrieve the length of the XpressNet message from the last four header bits
    Xpnlength = inData(indexHeader) Mod 16         ' Mod 16 gives us the last four bits
    indexXor = indexHeader + Xpnlength + 1         ' length excludes the XOR
    ' Exit, if we have not yet received the last byte of an XpressNet message
    If indexNextInByte <= indexXor Then Exit Sub
    ' STEP 3: We have a complete XpressNet packet. Should we log it?
    If Parent.LogHexMessages Then Parent.Logfile.ReceiveHexRange inData(), 0, indexXor
    ' STEP 4: Check if the XOR value is correct
    xorValue = 0                               ' Initialise
    For i = indexHeader To indexXor            ' If we include the XOR byte in the XOR calculation
      xorValue = xorValue Xor inData(i)        ' the result should become 0
    Next
    If xorValue = 0 Then                       ' No XOR error. XpressNet message is OK
      ' STEP 5: Is this a broadcast message?
      broadcast = False                        ' Is this a broadcast message??
      If (inData(0) = &HFF) And (inData(1) = &HFD) Then broadcast = True
      ' STEP 6: Copy the XpressNet message into the xpnMessage buffer
      i = 0
      For j = indexHeader To (indexHeader + Xpnlength)
        XpnMessage(i) = inData(j)
        i = i + 1
      Next                                     ' At the end, i holds the length of the XpressNet message
      ' STEP 7: Call the subroutine that analyses the individual XpressNet message
      Call DecodeXpressNetMessage(i, XpnMessage, broadcast)
    Else                                       ' XOR Error
      Parent.Logfile.Errors "XOR Error in received XpressNet message"
      If Parent.LogErrors Then Parent.Logfile.ReceiveHexRange inData, 0, (Xpnlength + 1)
    End If
    ' STEP 8 Remove the old message from the inData buffer, and shift the remainder to the begin
    ' Note that if, for some reason, the inData buffer contains garbage, the buffer will be emptied
    ' if the Acknowledgement times out.
    i = 0
    For j = (indexXor + 1) To (indexNextInByte - 1)
      inData(i) = inData(j)
      i = i + 1
    Next
    indexNextInByte = i                       ' points to the next empty byte
  Wend
End Sub


Public Sub DecodeXpressNetMessage(ByVal length As Byte, ByRef message() As Byte, ByVal broadcast As Boolean)
  ' This routine is called to analyse an individual XpressNet message
  ' The first byte should be the XpressNet header; framing bytes should already be removed
  ' XOR has already been checked and the XOR byte is removed.
  ' STEP 1: Extract from the XpressNet header byte the message length as well as the main command type
  Dim msgLength As Byte
  Dim msgCmd As Byte
  Dim header As Byte
  header = message(0)
  msgLength = header Mod 16       ' The length is contained in the last 4 bits
  msgCmd = header \ 16            ' And the main command type in the 4 highest order bits (\ is Integer division)
  ' STEP 2: Set the flow-control flag, except when the message was a broadcast
  If broadcast = False Then
    Parent.send.waitForAck = False
    Parent.send.stoptcpResponseTimer
  End If
  ' STEP 3: Call the procedure that is specific for the message received
  ' Check first for feedback messages, which have as Command = 4 but varying length
  If msgCmd = 4 Then
    Call DecodeFeedback(length, message())
  Else
    ' All other message have a fixed length, so we can do a Case statement on the entire header
    Select Case header
      Case 1: Call AnswerFromInterface(message(1))
      Case 2: Call Parent.eventInterfaceVersion(message(1), message(2))
      Case &H61:
        Select Case message(1)
          Case 0: Call Parent.eventNotAus
          Case 1: Call Parent.eventAllesAn
          Case 2: Call Parent.eventProgrammierMode
          Case &H8: Call Parent.eventProgrammierinfoKurzschluss  ' Z21 LAN_X_BC_TRACK_SHORT_CIRCUIT
          Case &H11: Call Parent.eventProgrammierinfoZentraleBereit
          Case &H12: Call Parent.eventProgrammierinfoKurzschluss
          Case &H13: Call Parent.eventProgrammierinfoDatenNichtGefunden
          Case &H1F: Call Parent.eventProgrammierinfoZentraleBusy
          Case &H80: Parent.Logfile.Errors "Ubertragungsfehler"
          Case &H81: Parent.Logfile.Errors "Zentrale Busy"
          Case &H82: Parent.Logfile.Warnings "Befehl in Zentrale nicht vorhanden"
        End Select
      Case &H62:
        If message(1) = &H22 Then Parent.eventStatusZentrale (message(2))
      Case &H63:
        Select Case message(1)
          Case 3:    Call Parent.eventModellZeit(message(2), message(3))
          Case &H10: Call Parent.eventProgInfoPageRegister(message(2), message(3))
          Case &H14: Call Parent.eventProgInfo(1, message(2), message(3)) ' CV1-255 u. 1024
          Case &H15: Call Parent.eventProgInfo(2, message(2), message(3)) ' CV256 - 511
          Case &H16: Call Parent.eventProgInfo(3, message(2), message(3)) ' CV512 - 767
          Case &H17: Call Parent.eventProgInfo(4, message(2), message(3)) ' CV768 - 1023
          Case &H20: Call Parent.eventServiceVariableMelden(message(2), message(3))
          Case &H21: Call Parent.eventSoftwareVersionZentrale(message(2), message(3))
        End Select
      Case &H64:
        Select Case message(1)
          Case &H14: Call Parent.eventProgInfoZ21(message(2), message(3), message(4))
          Case &H24: Call Parent.eventPomErgebnisMelden(message(2), message(3), message(4))
          Case &H25: Call Parent.eventModellZeit(message(2), message(3), message(4))
        End Select
      Case &H67:
        If message(1) = &H23 Then
          Call Parent.eventErweiterteVersionZentrale(message(2), message(3), message(4), message(5), message(6), message(7))
        End If
      Case &H81:
        If message(1) = &H0 Then Call Parent.eventNotHalt
      Case &HE1: Call Parent.eventFehlermeldung(message(1))
      Case &HE2:     ' MTR-Basisadresse
      Case &HE3:
        'If msgCmd = 3 Then ' Lok-Suchergebnis
        Select Case message(1)
          Case &H40: ' Lok besetzt
          Case &H50: Call Parent.eventFunctionStatusF0F12(message(2), message(3))
          Case &H52: Call Parent.eventFunctionsZustandF13F20(message(2), message(3))
        End Select
      Case &HE4:
        If message(1) = &H51 Then
          Call Parent.eventFunctionStatusF13F20(message(2), message(3), message(4))
        Else
          Call Parent.eventNormaleLokinfo(message(1), message(2), message(3), message(4))
        End If
      Case &HE5:     ' MTR-Mitglied
      Case &HE6:
        Select Case message(1)
          Case &H53: Call Parent.eventFunctionsZustandF29F68(message(2), message(3), message(4), message(5), message(6))
          Case &H54: Call Parent.eventFunctionStatusF29F68(message(2), message(3), message(4), message(5), message(6))
          Case Else: ' Lok ist in DTR
        End Select
      Case &HEF:
        Call Parent.eventZ21Lokinfo(message(1), message(2), message(3), message(4), message(5), message(6), message(7), message(8))
      Case &HF2:
        Select Case message(1)
          Case 1: Call Parent.eventXpressNetAddress(message(2))
          Case 2: Call Parent.eventXpressNetVersion(message(2))
          Case 3: Call Parent.eventInterfaceFreieVerbindungen(message(2))
        End Select
      Case &HF3:
        Select Case message(1)
          Case &HA: Call Parent.eventz21FirmwareVersion(message(2), message(3))
        End Select
    End Select
  End If
  ' STEP 4: Try to send the next message that is waiting in the queue
  ' Flow control (check WaitForAck) is performed within nextmessage()
  Parent.send.nextMessage
  ' STEP 5: Update the counter for the number of received messages
  Parent.NumberOfExpressNetMessagesReceived = Parent.NumberOfExpressNetMessagesReceived + 1
End Sub

'======================================================================================================
' V3.6: Section 1.5 - Antworten des Interface 23151
Private Sub AnswerFromInterface(ByVal data As Byte)
  Select Case data
    Case 1:
      Parent.Logfile.Errors "Interface: Die im Header angegebene Anzahl Bytes stimmt nicht mit der Anzahl der empfangenen Bytes überein"
    Case 2:
      Parent.Logfile.Errors "Interface: Fehler zwischen Interface und Zentrale (Timeout bei Datenubertragung des Interface an die Zentrale)"
    Case 3:
      Parent.Logfile.Errors "Interface: Nnbekannter Fehler (Zentrale adressierte das Interface mit Quittierungsaufforderung)"
    Case 4:                       ' Befehl ist an Zentrale geschickt
    Case 5:
      Parent.Logfile.Warnings "Interface: Zentrale adressiert das Interface nicht mehr"
    Case 6:
      Parent.Logfile.Errors "Interface: Puffer-Uberlauf im Interface 23151"
    Case 7:
      Parent.Logfile.Warnings "Interface: Zentrale adressiert das Interface wieder"
    Case 8:
      Parent.Logfile.Warnings "Interface: Derzeit konnen keine Befehle an die Zentrale gesendet werden"
    Case 9:
      Parent.Logfile.Errors "Interface: Fehler in den Befehlsparametern (z.B. Lokadresse falsch)"
    Case 10:
      Parent.Logfile.Errors "Interface: Unbekannter Fehler (Zentrale lieferte nicht die erwartete Antwort)"
  End Select
End Sub



' ----------------------------------------------------------------------------------------------------
' Ruckmeldung / Schaltinformation (V3.6: 3.1.1.5 and 3.1.8, V3.8: 2.4.5, 2.15, 2.16)
' Z21: 5.3: LAN_X_TURNOUT_INFO
' DecodeFeedback is called whenever an XpressNet Feedback message is received.
' It may carry a single feedback message (Schaltinformation) or multiple messages (Ruckmeldung).
' The message may be a point-to-point, or a broadcast message
' The input is the XpressNet message, without the Framing bytes and XOR byte
' The addresses in this routine are RS-Bus / decoder adresses.
' TODO: V3.8 TESTEN!!!
Private Sub DecodeFeedback(ByVal length As Byte, ByRef message() As Byte)
  ' STEP 1: Declare the variables that we need
  Dim msgLength As Byte                 ' The length of this XpressNet Feedback message, as indicated in the header
  Dim numberOfFeedbacks As Byte         ' A maximum of 8 Feedbacks may be contained within a single XpressNet message
  Dim feedbackNumber As Byte            ' A number between 1 and 8 (the maximum number of feedbacks allowed in a message
  Dim index As Byte                     ' Array index where the next Feedback starts
  Dim address As Integer                ' The RS-Bus / decoder address of this specific Feedback
  Dim data As Byte                      ' The RS-Bus / decoder data of this specific feedback
  Dim feedbackBits As Byte              ' The four feedback bits that were carried in this message
  Dim ttBits As Byte                    ' The TT Bits inform the kind of feedback decoder that generated the message
  Dim iBit As Byte                      ' The I-Bit tells if the switch is still moving
  Dim Nibble As Byte                    ' 0 = low nibble (feedback bits 1..4). 1 = high nibble (feedbaack bits 5..8)
  ' The Lenz feedback messages and the Z21 Extended Accessory Info messages have the same starting bytes
  ' Therefore we first check if we have a Z21 command station (&H12).
  ' If yes, we call the associated sub and exit this subroutine
  ' If no, we hanlde the bytes as normal XpressNet messages
  ' eventExtendedAccessoryInfo(ByVal Adr_MSB As Byte, ByVal Adr_LSB As Byte, ByVal DDDDDDDD As Byte, ByVal Status As Byte)
  If (Parent.ZentraleKennung = &H12) And (message(0) = &H44) Then
    Call Parent.eventExtendedAccessoryInfo(message(1), message(2), message(3), message(4))
    Exit Sub
  End If
  ' If we are here, we do not have a Z21 command station, but normal (Lenz) XpressNet command station
  ' STEP 2: Determine the length of the feedback message and the number of feedbacks
  msgLength = message(0) Mod 16         ' Isolated the lowest order 4 bits
  If (msgLength + 1) <> length Then     ' Extra test: if the length coded in the message matches the received size
    Parent.Logfile.Errors "Error decoding feedback message: Length doesn't match received Length"
    Exit Sub
  End If
  If msgLength = 3 Then
    ' STEP 3: V3.8: Schaltinformation ab W1025 / Z21: LAN_X_TURNOUT_INFO
    ' The Z21 does not return the Nibble, ttBits and the iBit. Only 2 data bits are used (a single switch)
    ' Two bytes for the address, plus one byte for the value
    address = (message(1) * 256) + message(2)
    ' The RS-Bus data is stored in the byte immediately after
    data = message(3)
    ' The "real feedback bits" are in bits 1..4 (0..3)
    feedbackBits = (data Mod 16)
    ' The nibble is coded in bit 5 (4). We can check this bit by performing an AND operation with the value 16 (0001 0000)
    Nibble = (data And 16) \ 16
    ' The TT bits are in coded in bits 6 and 7 (5 and 6)
    ttBits = (data And 96) \ 32
    ' The iBit is coded in bit 8 (7).
    iBit = data \ 128
    ' Call feedback to raise a feedback event
    Call Parent.eventFeedback(address, Nibble, ttBits, iBit, feedbackBits)
  Else
    ' STEP 4: V3.6: Schaltinformation / BC Ruckmeldung
    ' Analyse each feedback individually.
    Dim i As Byte
    numberOfFeedbacks = msgLength \ 2
    For i = 0 To (numberOfFeedbacks - 1)
      ' The RS-Bus address is stored at the Array indices 1, 3, 5, 7 etc => (i * 2 + 1)
      index = 1 + (i * 2)
      address = message(index)
      ' The RS-Bus data is stored in the byte immediately after
      data = message(index + 1)
      ' The "real feedback bits" are in bits 1..4
      feedbackBits = (data Mod 16)
      ' The nibble is coded in bit 5. We can check this bit by performing an AND operation with the value 16 (0001 0000)
      Nibble = (data And 16) \ 16
      ' The TT bits are in coded in bits 6 and 7
      ttBits = (data And 96) \ 32
      ' The iBit is coded in bit 8.
      iBit = data \ 128
      ' Call feedback to raise a feedback event
      Call Parent.eventFeedback(address, Nibble, ttBits, iBit, feedbackBits)
      ' TODO: IBit TOEVOEGEN
    Next
  End If
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







