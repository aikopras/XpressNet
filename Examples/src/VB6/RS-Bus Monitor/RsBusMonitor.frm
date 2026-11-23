VERSION 5.00
Begin VB.Form RsBusMonitor 
   Caption         =   "RsBusMonitor"
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17760
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   17760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Switch2RS 
      Caption         =   "<"
      Height          =   225
      Left            =   1680
      TabIndex        =   16
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton RS2Switch 
      Caption         =   ">"
      Height          =   225
      Left            =   1680
      TabIndex        =   15
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox TextSwitch2RS 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox TextRS2Switch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton SchaltInfo 
      Caption         =   "Get RS-Bus data"
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton CmdAllesAus 
      Caption         =   "Alles Aus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3120
      Width           =   1125
   End
   Begin VB.CommandButton CmdAllesAn 
      Caption         =   "Alles An"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   1125
   End
   Begin VB.CommandButton cmdTcpConnect 
      Caption         =   "Connect TCP/IP"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1680
      TabIndex        =   3
      Text            =   "192.168.0.200"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox picPanel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   3720
      ScaleHeight     =   6195
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   240
      Width           =   9492
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      Height          =   1335
      Left            =   360
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label LabelAddressConversion 
      Alignment       =   2  'Center
      Caption         =   "Address Conversions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label LabelSwitchAddress 
      Alignment       =   2  'Center
      Caption         =   "Switch"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label LabelRSAddress 
      Alignment       =   2  'Center
      Caption         =   "RS"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Shape ShapeAnAus 
      BorderColor     =   &H80000006&
      Height          =   2055
      Left            =   360
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Shape ledInterfaceStatus 
      BackColor       =   &H008080FF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape ledNotHalt 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape ledNotAus 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label LabelNotHalt 
      AutoSize        =   -1  'True
      Caption         =   "NotHalt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   855
      TabIndex        =   9
      Top             =   4710
      Width           =   675
   End
   Begin VB.Label LabelNotAus 
      AutoSize        =   -1  'True
      Caption         =   "NotAus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   855
      TabIndex        =   8
      Top             =   4230
      Width           =   660
   End
   Begin VB.Label LabelConnected 
      AutoSize        =   -1  'True
      Caption         =   "Connected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   855
      TabIndex        =   7
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label LabelStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label LabelIpAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IP Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   525
      TabIndex        =   5
      Top             =   1320
      Width           =   990
   End
   Begin VB.Shape ShapeTcp 
      BorderColor     =   &H80000006&
      Height          =   1335
      Left            =   360
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label_Header 
      Alignment       =   2  'Center
      Caption         =   "RS-Bus Matrix display"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7680
      TabIndex        =   1
      Top             =   4680
      Width           =   972
   End
End
Attribute VB_Name = "RsBusMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 2026-11 modified: for version 2.0 of the library. The program now uses the XpressNet DLL (XpressNet.dll), instead of a
' copy of the classes of the library. New: the IP address is kept for the next start (SaveSetting / GetSetting), the
' logfile directory TMC-Logfiles, the events SendError, InterfaceNichtAdressiert and InterfaceWiederAdressiert (shown
' in LabelStatus), and "Get RS-Bus data" in two parts of 128 requests (see SchaltInfo_Click). Get RS-Bus data stops
' after SendError, a new connection, TcpClosedByRemoteHost and TcpError
Option Explicit

' Use the events from the XpressNet object
Public WithEvents XpressNet As XpressNetClass
Attribute XpressNet.VB_VarHelpID = -1

' Referentie naar de PanelManager
Private m_PanelManager As PanelManager


Private Const LEFT_OFFSET As Long = 240

Private GetSchaltInfo As Boolean
Private AnswerCount As Integer                     ' 2026-11 modified: new. The answers for RS-Bus address 64 and 128



Private Sub Form_Load()
    Set XpressNet = New XpressNetClass
    ' Set the directory for logfiles
    ' 2026-11 modified: TMC-Logfiles, the default directory of version 2.0 (as in the other examples)
    'XpressNet.Logfile.LogDirectory = "C:\Users\Public\RSbusMatrix-Logfiles"
    XpressNet.Logfile.LogDirectory = "C:\Users\Public\TMC-Logfiles"
    ' 2026-11 modified: new. The IP address of the previous connection (the first time: the text of txtIP)
    txtIP.Text = GetSetting("RS-Bus Monitor", "Settings", "IpAddress", txtIP.Text)
    Me.ScaleMode = vbPixels
    picPanel.ScaleMode = vbPixels
    picPanel.AutoRedraw = True
    picPanel.BorderStyle = 0
    ' Initialiseer de PanelManager
    Set m_PanelManager = New PanelManager
    m_PanelManager.init picPanel
    picPanel.Move LEFT_OFFSET, 0, Me.ScaleWidth - LEFT_OFFSET, Me.ScaleHeight
    GetSchaltInfo = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        picPanel.Move LEFT_OFFSET, 0, Me.ScaleWidth - LEFT_OFFSET, Me.ScaleHeight
        m_PanelManager.Resize
    End If
End Sub


'===============================================================================================================
' XPressNet Commando's
Private Sub cmdTcpConnect_Click()
  LabelStatus = "TCP connecting ....."
  ' 2026-11 modified: new. The IP address is kept for the next start of the program
  SaveSetting "RS-Bus Monitor", "Settings", "IpAddress", txtIP.Text
  ' 2026-11 modified: new. A new connection stops Get RS-Bus data: the answers on its requests will not come
  GetSchaltInfo = False
  Call XpressNet.CmdConnectTcp(txtIP.Text, 5550)
End Sub

Private Sub CmdAllesAn_Click()
  Call XpressNet.CmdAllesAn
End Sub

Private Sub CmdAllesAus_Click()
  Call XpressNet.CmdAllesAus
End Sub

'===========================================================================================
' ALGEMENE EVENTS
Private Sub XpressNet_InterfaceConnected()
  ledInterfaceStatus.FillColor = vbGreen
  LabelStatus = "Connection established"
End Sub

Private Sub XpressNet_TcpClosedByRemoteHost()
  ledInterfaceStatus.FillColor = vbRed
  GetSchaltInfo = False                            ' 2026-11 modified: new. Get RS-Bus data stops
  LabelStatus = "TCP connection closed by remote host"
End Sub

Private Sub XpressNet_TcpError(ByRef ErrorDescription As String)
  ledInterfaceStatus.FillColor = vbRed
  GetSchaltInfo = False                            ' 2026-11 modified: new. Get RS-Bus data stops
  LabelStatus = ErrorDescription
End Sub

Private Sub XpressNet_AllesAn()
  ledNotAus.FillColor = vbGreen
  ledNotHalt.FillColor = vbGreen
End Sub

Private Sub XpressNet_NotAus()
  ledNotAus.FillColor = vbRed
End Sub

Private Sub XpressNet_NotHalt()
Debug.Print "NotHalt"
  ledNotHalt.FillColor = vbRed
End Sub

' 2026-11 modified: new (version 2.0). A message that the library could not send. Reason: see SendErrorReasonEnum
' Get RS-Bus data is stopped (for example: no answer, the message queues of the library have been emptied)
Private Sub XpressNet_SendError(ByVal Reason As Byte, ByVal Description As String)
  GetSchaltInfo = False
  LabelStatus = sendErrorName(Reason) & ": " & Description
End Sub

' 2026-11 modified: new (version 2.0). The command station does not address the interface: sending is paused
Private Sub XpressNet_InterfaceNichtAdressiert()
  LabelStatus = "Interface nicht adressiert"
End Sub

' 2026-11 modified: new (version 2.0). The command station addresses the interface again
Private Sub XpressNet_InterfaceWiederAdressiert()
  LabelStatus = "Interface wieder adressiert"
End Sub

' 2026-11 modified: new. The name of the reason of the event SendError (in VB6 the name of an Enum value is not
' available at run time). The names are the names of SendErrorReasonEnum
Private Function sendErrorName(ByVal Reason As Byte) As String
  Select Case Reason
    Case WiederholungErfolglos: sendErrorName = "WiederholungErfolglos"
    Case BefehlNichtVorhanden: sendErrorName = "BefehlNichtVorhanden"
    Case Timeout: sendErrorName = "Timeout"
    Case FifoVoll: sendErrorName = "FifoVoll"
    Case KeineVerbindung: sendErrorName = "KeineVerbindung"
    Case FehlerInDenBefehlsparametern: sendErrorName = "FehlerInDenBefehlsparametern"
    Case ZentraleAdressiertDasInterfaceNichtMehr: sendErrorName = "ZentraleAdressiertDasInterfaceNichtMehr"
    Case Else: sendErrorName = "SendError " & Reason
  End Select
End Function


'===========================================================================================
' FEEDBACK EVENTS

'' We hebben in de DLL 2 events voor RS-Bus input: Bezetmeld en Wissel feedbacks
'' Voor beide roepen we SetFeedbackNibble aan
''
'' Bezetmeld: Private Sub XpressNet_Feedback(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte)
'' - address = 1..128
'' - nibble = 0 / 1
'' - data = 0..15
Private Sub XpressNet_Feedback(ByVal address As Integer, ByVal Nibble As Byte, ByVal data As Byte)
  m_PanelManager.SetFeedbackNibble address, Nibble, data, BEZETMELDING
  ' 2026-11 modified: new. Get RS-Bus data: is this the answer on the last request of a part?
  CheckSchaltInfo address
End Sub
''
'' Wissel: Private Sub XpressNet_SwitchFeedback(switchAddress, ttBits, iBit, data)
'' - switchAddress =  ((address * 4) + (Nibble * 2) + 1)
'' - ttBits =
'' - iBit =
'' - data = 0..15
Private Sub XpressNet_SwitchFeedback(ByVal switchAddress As Integer, ByVal ttBits As Byte, ByVal iBit As Byte, ByVal data As Byte)
  Dim RsAddress As Integer
  Dim Nibble As Byte
  RsAddress = Int((switchAddress - 1) / 4) + 1
  Nibble = ((switchAddress - 1) \ 2) Mod 2
  If ttBits = 0 Then     ' Feedback comes from Command Station
    If Not GetSchaltInfo Then
      m_PanelManager.SetFeedbackNibble RsAddress, Nibble, data, CENTRALE
    End If
  Else                   ' Feedback comes from Decoder
    m_PanelManager.SetFeedbackNibble RsAddress, Nibble, data, WISSEL
  End If
  ' 2026-11 modified: new. Get RS-Bus data: is this the answer on the last request of a part?
  CheckSchaltInfo RsAddress
End Sub


'===========================================================================================
' VERTALING TUSSE ADDRESSEN
Private Sub RS2Switch_Click()
    Dim rsAddr As Integer
    Dim firstSw As Integer
    Dim lastSw As Integer
    If IsNumeric(TextRS2Switch.Text) Then
        rsAddr = CInt(TextRS2Switch.Text)
        If rsAddr >= 1 Then
            firstSw = (rsAddr - 1) * 4 + 1
            lastSw = firstSw + 3
            TextSwitch2RS.Text = firstSw & " .. " & lastSw
        End If
    End If
    ' Ongeldige input? Doe niets
End Sub

Private Sub Switch2RS_Click()
    Dim swAddr As Integer
    Dim rsAddr As Integer
    If IsNumeric(TextSwitch2RS.Text) Then
        swAddr = CInt(TextSwitch2RS.Text)
        If swAddr >= 1 Then
            rsAddr = (swAddr - 1) \ 4 + 1
            TextRS2Switch.Text = rsAddr
        End If
    End If
    ' Ongeldige input? Doe niets
End Sub


'===========================================================================================
' Alles ophalen
' 2026-11 modified: Get RS-Bus data in two parts of 128 requests (64 RS-Bus addresses, two nibbles each). The FIFO of
' version 2.0 of the library holds 128 messages: more requests at once would give the event SendError (FifoVoll).
' The second part is requested as soon as both requests for RS-Bus address 64 have been answered; Get RS-Bus data
' ends when both requests for RS-Bus address 128 have been answered (see CheckSchaltInfo). This does not depend on
' the speed of the interface. The event SendError stops Get RS-Bus data.
' (Old code: four parts of 64 requests, with a pause of 500 ms after each part)
Private Sub SchaltInfo_Click()
  If GetSchaltInfo Then Exit Sub                   ' Get RS-Bus data is already active
  GetSchaltInfo = True
  AnswerCount = 0
  LabelStatus = "Requesting RS-Bus data ....."
  RequestFeedbacks 1, 64
End Sub

' 2026-11 modified: new. Requests both nibbles of the RS-Bus addresses firstAddress .. lastAddress
Private Sub RequestFeedbacks(ByVal firstAddress As Integer, ByVal lastAddress As Integer)
  Dim address As Integer
  For address = firstAddress To lastAddress
    Call XpressNet.CmdFeedbackRequest(address, 0)  ' Nibble = 0
    Call XpressNet.CmdFeedbackRequest(address, 1)  ' Nibble = 1
  Next
End Sub

' 2026-11 modified: new. Called for every feedback message, after the matrix has been updated. The answers are
' counted per RS-Bus address, and not per nibble: for an RS-Bus address without feedback information, the command
' station (LZ100) answers both requests with nibble 0 (42 Address 00).
' Note: an unrequested feedback message (broadcast) for RS-Bus address 64 or 128 is counted as well
Private Sub CheckSchaltInfo(ByVal RsAddress As Integer)
  If Not GetSchaltInfo Then Exit Sub
  If (RsAddress <> 64) And (RsAddress <> 128) Then Exit Sub
  AnswerCount = AnswerCount + 1
  If AnswerCount = 2 Then RequestFeedbacks 65, 128 ' The first part has been answered
  If AnswerCount = 4 Then                          ' The second part has been answered
    GetSchaltInfo = False
    LabelStatus = "RS-Bus data received"
  End If
End Sub

'Private Sub SchaltInfo_Click()
'  GetSchaltInfo = True
'  Dim delayTime As Long
'  delayTime = 500                                  ' in ms
'  Dim address As Integer
'  For address = 1 To 32
'    Call XpressNet.CmdFeedbackRequest(address, 0)  ' Nibble = 0
'    Call XpressNet.CmdFeedbackRequest(address, 1)  ' Nibble = 1
'  Next
'  WaitMilliseconds delayTime                       ' Pauze 200 ms zonder UI blokkering
'  For address = 33 To 64
'    Call XpressNet.CmdFeedbackRequest(address, 0)  ' Nibble = 0
'    Call XpressNet.CmdFeedbackRequest(address, 1)  ' Nibble = 1
'  Next
'  WaitMilliseconds delayTime                       ' Pauze 200 ms zonder UI blokkering
'  For address = 65 To 96
'    Call XpressNet.CmdFeedbackRequest(address, 0)  ' Nibble = 0
'    Call XpressNet.CmdFeedbackRequest(address, 1)  ' Nibble = 1
'  Next
'  WaitMilliseconds delayTime                       ' Pauze 200 ms zonder UI blokkering
'  For address = 97 To 128
'    Call XpressNet.CmdFeedbackRequest(address, 0)  ' Nibble = 0
'    Call XpressNet.CmdFeedbackRequest(address, 1)  ' Nibble = 1
'  Next
'  WaitMilliseconds delayTime                       ' Pauze 200 ms zonder UI blokkering
'  GetSchaltInfo = False
'End Sub

'Private Sub WaitMilliseconds(ByVal ms As Long)
'    Dim endTime As Single
'    endTime = Timer + ms / 1000
'    Do While Timer < endTime
'        DoEvents
'    Loop
'End Sub



'===========================================================================================
' Release the XpressNet object (which in turn closes the TCP connection)
' If we don't do this, some garbage will remain somewhere within the Window system
' causing all kinds of problems, such as "subscript out of range" messages
' after the program is restarted.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set XpressNet = Nothing
  ' Laat de IDE weten dat het programma stopt
  End
End Sub

