VERSION 5.00
Begin VB.Form FormXpressNet 
   AutoRedraw      =   -1  'True
   Caption         =   " XPressNet basic test"
   ClientHeight    =   9645
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboBaudUsb 
      Height          =   360
      Left            =   4440
      TabIndex        =   82
      Text            =   "57600"
      ToolTipText     =   "Lenz: 57600 / Yamorc: 115200"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdLokAusStack 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   81
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton CmdLokEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   9600
      TabIndex        =   80
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton CmdGetF0_F12 
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   79
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton CmdFunctionOff 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   78
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdFunctionOn 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11640
      TabIndex        =   77
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdGetF29_68 
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   76
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdGetF13_28 
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   75
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox TxtLokFNumber 
      Height          =   375
      Left            =   11160
      TabIndex        =   65
      Text            =   "1"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdLokForward 
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   63
      Top             =   7440
      Width           =   615
   End
   Begin VB.CommandButton CmdLokReverse 
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   62
      Top             =   7440
      Width           =   615
   End
   Begin VB.ComboBox ComboLokSteps 
      Height          =   360
      Left            =   9600
      TabIndex        =   61
      Text            =   "28"
      Top             =   8520
      Width           =   735
   End
   Begin VB.TextBox TxtLokSpeed 
      Height          =   375
      Left            =   9600
      TabIndex        =   60
      Text            =   "0"
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton CmdLanInfo 
      Caption         =   "LAN Info"
      Height          =   600
      Left            =   11640
      TabIndex        =   55
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton CmdRsFeedback 
      Caption         =   "Feedback"
      Height          =   495
      Left            =   6360
      TabIndex        =   53
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtRsBusAddress 
      Height          =   375
      Left            =   6360
      TabIndex        =   48
      ToolTipText     =   "1..128"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox TxtBoxRsNibble 
      Height          =   3975
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   47
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox TxtBoxRsData 
      Height          =   3975
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox TxBoxtRsAddress 
      Height          =   3975
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox TxtBoxSwitchIBit 
      Height          =   3975
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox TxtBoxSwitchTtBits 
      Height          =   3975
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   39
      ToolTipText     =   "C=Command Station, D=Decoder"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox TxtBoxSwitchData 
      Height          =   3975
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox TxtBoxSwitchAddress 
      Height          =   3975
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton CmdUsbConnect 
      Caption         =   "Connect USB"
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox TxtUsbPort 
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
      Left            =   4440
      TabIndex        =   34
      Text            =   "3"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton CmdInterfaceInfo 
      Caption         =   "Interface Info"
      Height          =   600
      Left            =   11640
      TabIndex        =   27
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLokAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Text            =   "3"
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton CmdEineLokAnhalten 
      Caption         =   "Anhalten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   25
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton CmdAlleLoksAnhalten 
      Caption         =   "Alle Loks anhalten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   24
      Top             =   8520
      Width           =   1600
   End
   Begin VB.CommandButton CmdAllesAus 
      Caption         =   "Alles Aus"
      Height          =   495
      Left            =   11040
      TabIndex        =   23
      Top             =   3600
      Width           =   1600
   End
   Begin VB.CommandButton CmdAllesAn 
      Caption         =   "Alles An"
      Height          =   495
      Left            =   9240
      TabIndex        =   22
      Top             =   3600
      Width           =   1600
   End
   Begin VB.CommandButton cmdSwitchFeedback 
      Caption         =   "Feedback"
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtWissel 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "1..2048"
      Top             =   3120
      Width           =   1212
   End
   Begin VB.CommandButton cmdAf 
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdRecht 
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtTcpPort 
      Height          =   372
      Left            =   1440
      TabIndex        =   3
      Text            =   "5550"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Height          =   372
      Left            =   1440
      TabIndex        =   2
      Text            =   "192.168.0.200"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdTcpConnect 
      Caption         =   "Connect TCP/IP"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Schakel wissel 1 om"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label LabelBaudrateUSB 
      AutoSize        =   -1  'True
      Caption         =   "Baudrate"
      Height          =   240
      Index           =   1
      Left            =   3480
      TabIndex        =   83
      ToolTipText     =   "Lenz: 57600 / Yamorc: 115200"
      Top             =   960
      Width           =   825
   End
   Begin VB.Shape ShapeAnAus 
      Height          =   4215
      Left            =   8520
      Top             =   5280
      Width           =   4815
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   68
      Left            =   12480
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   67
      Left            =   12360
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   66
      Left            =   12240
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   65
      Left            =   12120
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   64
      Left            =   11880
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   63
      Left            =   11760
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   62
      Left            =   11640
      Top             =   8160
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   61
      Left            =   11520
      Top             =   8160
      Width           =   135
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F61-F68"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10800
      TabIndex        =   74
      Top             =   8160
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F53-F60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10800
      TabIndex        =   73
      Top             =   7920
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F45-F52"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   10800
      TabIndex        =   72
      Top             =   7680
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F37-F44"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   10800
      TabIndex        =   71
      Top             =   7440
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F29-F36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   10800
      TabIndex        =   70
      Top             =   7200
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F21-F28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   10800
      TabIndex        =   69
      Top             =   6960
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F13-F20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   10830
      TabIndex        =   68
      Top             =   6720
      Width           =   585
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F5-F12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   10950
      TabIndex        =   67
      Top             =   6480
      Width           =   495
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   60
      Left            =   12480
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   59
      Left            =   12360
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   58
      Left            =   12240
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   57
      Left            =   12120
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   56
      Left            =   11880
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   55
      Left            =   11760
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   54
      Left            =   11640
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   53
      Left            =   11520
      Top             =   7920
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   52
      Left            =   12480
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   51
      Left            =   12360
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   50
      Left            =   12240
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   49
      Left            =   12120
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   48
      Left            =   11880
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   47
      Left            =   11760
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   46
      Left            =   11640
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   45
      Left            =   11520
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   44
      Left            =   12480
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   43
      Left            =   12360
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   42
      Left            =   12240
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   41
      Left            =   12120
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   40
      Left            =   11880
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   39
      Left            =   11760
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   38
      Left            =   11640
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   37
      Left            =   11520
      Top             =   7440
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   36
      Left            =   12480
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   35
      Left            =   12360
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   34
      Left            =   12240
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   33
      Left            =   12120
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   32
      Left            =   11880
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   31
      Left            =   11760
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   30
      Left            =   11640
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   29
      Left            =   11520
      Top             =   7200
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   28
      Left            =   12480
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   27
      Left            =   12360
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   26
      Left            =   12240
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   25
      Left            =   12120
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   24
      Left            =   11880
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   23
      Left            =   11760
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   22
      Left            =   11640
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   21
      Left            =   11520
      Top             =   6960
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   20
      Left            =   12480
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   19
      Left            =   12360
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   18
      Left            =   12240
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   17
      Left            =   12120
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   16
      Left            =   11880
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   15
      Left            =   11760
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   14
      Left            =   11640
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   13
      Left            =   11520
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   12
      Left            =   12480
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   11
      Left            =   12360
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   10
      Left            =   12240
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   12120
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   11880
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   11760
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   11640
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   11520
      Top             =   6480
      Width           =   135
   End
   Begin VB.Label LabelFunctions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F0-F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   11040
      TabIndex        =   66
      Top             =   6225
      Width           =   405
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   12480
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   12360
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   12240
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   12120
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape LEDF0F68 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   11880
      Top             =   6240
      Width           =   135
   End
   Begin VB.Label LabelFunktion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "F:"
      Height          =   240
      Left            =   11000
      TabIndex        =   64
      Top             =   5805
      Width           =   165
   End
   Begin VB.Label LabelLoks 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Loks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9360
      TabIndex        =   59
      Top             =   5280
      Width           =   585
   End
   Begin VB.Label LabelLokSteps 
      Alignment       =   1  'Right Justify
      Caption         =   "Steps"
      Height          =   255
      Left            =   8760
      TabIndex        =   58
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label LabelLokSpeed 
      Alignment       =   1  'Right Justify
      Caption         =   "Speed"
      Height          =   255
      Left            =   8760
      TabIndex        =   57
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label LabelLokAddress 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   255
      Left            =   8760
      TabIndex        =   56
      Top             =   5760
      Width           =   735
   End
   Begin VB.Line Line6 
      X1              =   9120
      X2              =   12840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   9000
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label LabelInfo2 
      AutoSize        =   -1  'True
      Caption         =   "latest message at top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   54
      Top             =   8880
      Width           =   2025
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Nibble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   52
      Top             =   4560
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   240
      Left            =   6240
      TabIndex        =   51
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   240
      Left            =   4800
      TabIndex        =   50
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Label LabelRsBusAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RSBus Address"
      Height          =   240
      Left            =   4800
      TabIndex        =   49
      Top             =   3120
      Width           =   1440
   End
   Begin VB.Shape ShapeFeedback 
      BorderWidth     =   2
      Height          =   6735
      Left            =   4440
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label LabelSwitchIBit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "I"
      Height          =   240
      Left            =   3600
      TabIndex        =   44
      Top             =   4560
      Width           =   195
   End
   Begin VB.Label LabelSwitchTtBits 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "TT"
      Height          =   240
      Left            =   3000
      TabIndex        =   43
      Top             =   4560
      Width           =   405
   End
   Begin VB.Label LabelSwitchdata 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   240
      Left            =   1920
      TabIndex        =   42
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label LabelSwitchAddress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   240
      Left            =   480
      TabIndex        =   41
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Shape ShapeSwitch 
      BorderWidth     =   2
      Height          =   6735
      Left            =   120
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Shape ShapeInfo 
      Height          =   2415
      Left            =   9000
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label LabelConnected 
      AutoSize        =   -1  'True
      Caption         =   "Connected"
      Height          =   240
      Left            =   6615
      TabIndex        =   36
      Top             =   510
      Width           =   975
   End
   Begin VB.Shape ShapeUsb 
      Height          =   1815
      Left            =   3360
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape ShapeTcp 
      BorderColor     =   &H80000006&
      Height          =   1815
      Left            =   120
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label LabelUSBPort 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "USB Port"
      Height          =   240
      Left            =   3435
      TabIndex        =   33
      Top             =   480
      Width           =   825
   End
   Begin VB.Label LabelProgrammiermode 
      AutoSize        =   -1  'True
      Caption         =   "Programmier Mode"
      Height          =   240
      Left            =   6615
      TabIndex        =   32
      Top             =   1950
      Width           =   1740
   End
   Begin VB.Label LabelNotAus 
      AutoSize        =   -1  'True
      Caption         =   "NotAus"
      Height          =   240
      Left            =   6615
      TabIndex        =   31
      Top             =   990
      Width           =   660
   End
   Begin VB.Label LabelNotHalt 
      AutoSize        =   -1  'True
      Caption         =   "NotHalt"
      Height          =   240
      Left            =   6615
      TabIndex        =   30
      Top             =   1470
      Width           =   675
   End
   Begin VB.Shape ledProgrammiermode 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape ledNotAus 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape ledNotHalt 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label txtMasterVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10635
      TabIndex        =   29
      Top             =   240
      Width           =   855
   End
   Begin VB.Label LabelMasterVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Master Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9480
      TabIndex        =   28
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8400
      X2              =   13440
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line4 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   8280
      X2              =   8280
      Y1              =   2760
      Y2              =   9480
   End
   Begin VB.Label txtTcpFree 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10635
      TabIndex        =   21
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label txtXVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10635
      TabIndex        =   20
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label txtXAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10635
      TabIndex        =   19
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label txtInterfaceCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10635
      TabIndex        =   18
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label txtInterfaceVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10635
      TabIndex        =   17
      Top             =   720
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   9120
      X2              =   12840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   13440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label labelTcpFree 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Free TCP sockets:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9120
      TabIndex        =   15
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Label labelXVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "XpressNet Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   14
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label labelXAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "XpressNet Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   13
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label labelCode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Interface Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9480
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label labelInterfaceVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Interface Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9360
      TabIndex        =   11
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label LabelSwitch 
      AutoSize        =   -1  'True
      Caption         =   "Switch:"
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5880
      X2              =   12360
      Y1              =   2640
      Y2              =   2640
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
      Left            =   120
      TabIndex        =   6
      Top             =   2205
      Width           =   4575
   End
   Begin VB.Shape ledInterfaceStatus 
      BackColor       =   &H008080FF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Label LabelTcpPort 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TCP Port"
      Height          =   240
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   810
   End
   Begin VB.Label LabelIpAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IP Address"
      Height          =   240
      Left            =   285
      TabIndex        =   4
      Top             =   480
      Width           =   990
   End
   Begin VB.Label LabeLabelInfo1 
      AutoSize        =   -1  'True
      Caption         =   "latest message at top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   1
      Top             =   8880
      Width           =   2025
   End
End
Attribute VB_Name = "FormXpressNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents XpressNet As XpressNetClass
Attribute XpressNet.VB_VarHelpID = -1


' ----------------------------------------------------------------------------------------------------------------------
' Function Groups:
' For each Lok Address, the values must be fethed again
Private FB(10) As Byte                 ' FB: Function Befehl (on / off)
Private FS(10) As Byte                 ' Function Status (tastend oder nicht tastend)


Public Sub Form_Load()
  Set XpressNet = New XpressNetClass
  ' Set the directory for logfiles
  XpressNet.Logfile.LogDirectory = "C:\Users\Public\XPN-Logfiles"
  ' Possible Baudrates for the USB interface
  ComboBaudUsb.AddItem "57600"
  ComboBaudUsb.AddItem "115200"
  ' Possible lok decoder steps
  ComboLokSteps.AddItem "14"
  ComboLokSteps.AddItem "27"
  ComboLokSteps.AddItem "28"
  ComboLokSteps.AddItem "128"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set XpressNet = Nothing
End Sub


'===============================================================================================================
Private Sub cmdTcpConnect_Click()
  LabelStatus = "TCP connecting ....."
  Call XpressNet.CmdConnectTcp(txtIP.Text, txtTcpPort.Text)
End Sub

Private Sub CmdUsbConnect_Click()
  LabelStatus = "USB connecting ....."
  Dim Settings As String
  Settings = ComboBaudUsb.Text & ",N,8,1"
  Call XpressNet.CmdConnectUsb(TxtUsbPort.Text, Settings)
End Sub


Private Sub LabelMasterVersion_Click()
  Call XpressNet.CmdSoftwareversionZentraleAnfordern
End Sub

Private Sub CmdInterfaceInfo_Click()
  Call XpressNet.CmdInterfaceVersion
End Sub

Private Sub CmdLanInfo_Click()
  Call XpressNet.CmdXpressNetAddress
  Call XpressNet.CmdXpressNetVersion
  Call XpressNet.CmdInterfaceFreieVerbindungen
End Sub

Private Sub CmdErweiterteVersion_Click()
  Call XpressNet.CmdErweiterteZentralenVersionsinformation
End Sub

Private Sub CmdStartMan_Click()
  Call XpressNet.CmdZentralenStartmodeSetzen(False)
End Sub

Private Sub cmdStartAuto_Click()
  Call XpressNet.CmdZentralenStartmodeSetzen(True)
End Sub

'===============================================================================================================
Private Sub cmdRecht_Click()
  If Not IsNumeric(txtWissel.Text) Then Exit Sub
  Call XpressNet.CmdAccessoryAuto(txtWissel.Text, 1)
End Sub

Private Sub cmdAf_Click()
  If Not IsNumeric(txtWissel.Text) Then Exit Sub
  Call XpressNet.CmdAccessoryAuto(txtWissel.Text, 0)
End Sub

Private Sub cmdSwitchFeedback_Click()
  If IsNumeric(txtWissel.Text) Then
    Call XpressNet.CmdSwitchInfoRequest(txtWissel.Text)
  End If
End Sub

Private Sub CmdRsFeedback_Click()
  If IsNumeric(TxtRsBusAddress.Text) Then
    Call XpressNet.CmdFeedbackRequest(TxtRsBusAddress.Text, 0)  ' Nibble 0
    Call XpressNet.CmdFeedbackRequest(TxtRsBusAddress.Text, 1)  ' Nibble 1
  End If
End Sub

Private Sub CmdAllesAn_Click()
  Call XpressNet.CmdAllesAn
End Sub

Private Sub CmdAllesAus_Click()
  Call XpressNet.CmdAllesAus
End Sub

Private Sub CmdAlleLoksAnhalten_Click()
  Call XpressNet.CmdAlleLoksAnhalten
End Sub

Private Sub CmdEineLokAnhalten_Click()
  If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
  Call XpressNet.CmdEineLokAnhalten(txtLokAddress.Text)
End Sub

Private Sub CmdLokEnter_Click()
  ' A new lok address is entered. Get the lok info
  If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
  Call ClearLedsForFunctions
  Call XpressNet.CmdLokInfoAnfordern(txtLokAddress.Text)
End Sub

Private Sub CmdLokAusStack_Click()
  If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
  Call XpressNet.CmdLokAusStack(txtLokAddress.Text)
End Sub

Private Sub CmdLokForward_Click()
  If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
  If Not IsNumeric(TxtLokSpeed.Text) Then Exit Sub
  Call XpressNet.CmdFahrbefehl(txtLokAddress.Text, TxtLokSpeed.Text, ComboLokSteps, 1)
End Sub

Private Sub CmdLokReverse_Click()
  If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
  If Not IsNumeric(TxtLokSpeed.Text) Then Exit Sub
  Call XpressNet.CmdFahrbefehl(txtLokAddress.Text, TxtLokSpeed.Text, ComboLokSteps, 0)
End Sub

Private Sub CmdFunctionOn_Click()
  Call handleFunctionBefehl(1)
End Sub
  
Private Sub CmdFunctionOff_Click()
  Call handleFunctionBefehl(0)
End Sub

Private Sub CmdGetF0_F12_Click()
  Call XpressNet.CmdLokInfoAnfordern(txtLokAddress.Text)
End Sub

Private Sub CmdGetF13_28_Click()
  Call XpressNet.CmdFunktionsZustandF13F28Anfordern(txtLokAddress.Text)
End Sub

Private Sub CmdGetF29_68_Click()
  Call XpressNet.CmdFunktionsZustandF29F68Anfordern(txtLokAddress.Text)
End Sub





'===========================================================================================
' EVENTS
Private Sub XpressNet_InterfaceConnected()
  ledInterfaceStatus.FillColor = vbGreen
  LabelStatus = "Connection established"
End Sub

Private Sub XpressNet_TcpClosedByRemoteHost()
  ledInterfaceStatus.FillColor = vbRed
  LabelStatus = "TCP connection closed by remote host"
End Sub

Private Sub XpressNet_TcpError(ByRef ErrorDescription As String)
  ledInterfaceStatus.FillColor = vbRed
  LabelStatus = ErrorDescription
End Sub

Private Sub XpressNet_Feedback(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte)
  TxBoxtRsAddress = address & vbCrLf & TxBoxtRsAddress
  TxtBoxRsData = printBits(data) & vbCrLf & TxtBoxRsData
  TxtBoxRsNibble = nibble & vbCrLf & TxtBoxRsNibble
End Sub

Private Sub XpressNet_SwitchFeedback(ByVal address As Integer, ByVal ttBits As Byte, ByVal iBit As Byte, ByVal data As Byte)
  Dim ttText As String
  If ttBits = 0 Then
    ttText = "C"  ' Feedback comes from Command Station
  Else
    ttText = "D"  ' Feedback comes from Decoder
  End If
  TxtBoxSwitchAddress = address & "/" & (address + 1) & vbCrLf & TxtBoxSwitchAddress
  TxtBoxSwitchData = printBits(data) & vbCrLf & TxtBoxSwitchData
  TxtBoxSwitchTtBits = ttText & vbCrLf & TxtBoxSwitchTtBits
  TxtBoxSwitchIBit = iBit & vbCrLf & TxtBoxSwitchIBit
  printBits (data)
End Sub

Private Sub XpressNet_AllesAn()
  ledNotAus.FillColor = vbGreen
  ledNotHalt.FillColor = vbGreen
  ledProgrammiermode.FillColor = vbGreen
End Sub

Private Sub XpressNet_NotAus()
  ledNotAus.FillColor = vbRed
End Sub

Private Sub XpressNet_NotHalt()
Debug.Print "NotHalt"
  ledNotHalt.FillColor = vbRed
End Sub

Private Sub XpressNet_ProgrammierMode()
  ledProgrammiermode.FillColor = vbRed
End Sub

Private Sub XpressNet_ZentraleVersion(ByVal version As String)
  txtMasterVersion = version
End Sub

Private Sub XpressNet_Fehlermeldung(ByVal Value As Byte)
  LabelStatus = "Fehlermeldung: " & Value & " (V3.6: 3.1.14)"
End Sub

Private Sub XpressNet_XpressNetAddress(ByVal address As Byte)
  txtXAddress = address
End Sub

Private Sub XpressNet_XpressNetVersion(ByVal version As String)
  txtXVersion = version
End Sub

Private Sub XpressNet_InterfaceVersion(ByVal version As String, ByVal code As String)
  txtInterfaceVersion = version
  txtInterfaceCode = code
End Sub

Private Sub XpressNet_InterfaceFreieVerbindungen(ByVal number As Byte)
  txtTcpFree = number
End Sub

Private Sub XpressNet_NormaleLokinfo(ByVal Besetzt As Boolean, ByVal Speedsteps As Integer, ByVal Speed As Integer, ByVal Forward As Boolean, ByVal F0_F4 As Byte, ByVal F5_F8 As Byte, ByVal F9_F12 As Byte)
  ' Debug.Print "Lok ist besetzt: " & Besetzt
  ' Debug.Print "Vorwarts: " & Forward
  ' Set the combobox and speed, based on the received input
  ComboLokSteps = Speedsteps
  TxtLokSpeed = Speed
  ' Store the values of F0..F12, to allow changing a single Function later
  FB(1) = F0_F4                  ' Gruppe 1
  FB(2) = F5_F8                  ' Gruppe 2
  FB(3) = F9_F12                 ' Gruppe 3
  ' Show the LEDs for F0 .. F12
  Call ShowFbGroup(F0_F4, 1)     ' Led FB bits for Gruppe 1
  Call ShowFbGroup(F5_F8, 2)     ' Led FB bits for Gruppe 1
  Call ShowFbGroup(F9_F12, 3)    ' Led FB bits for Gruppe 1
End Sub

Private Sub XpressNet_FunctionsZustandF13F20(ByVal F13_F20 As Byte, ByVal F21_F28 As Byte)
  FB(4) = F13_F20                ' Store Gruppe 4
  FB(5) = F21_F28                ' Store Gruppe 5
  Call ShowFbGroup(F13_F20, 4)   ' Led FB bits for Gruppe 4
  Call ShowFbGroup(F21_F28, 5)   ' Led FB bits for Gruppe 5
End Sub
Private Sub XpressNet_FunctionsZustandF29F68(ByVal F29_F36 As Byte, ByVal F37_F44 As Byte, ByVal F45_F52 As Byte, ByVal F53_F60 As Byte, ByVal F61_F68 As Byte)
  FB(4) = F29_F36                ' Store Gruppe 6
  FB(4) = F37_F44                ' Store Gruppe 7
  FB(4) = F45_F52                ' Store Gruppe 8
  FB(4) = F53_F60                ' Store Gruppe 9
  FB(5) = F61_F68                ' Store Gruppe 10
  Call ShowFbGroup(F29_F36, 6)   ' Led FB bits for Gruppe 6
  Call ShowFbGroup(F37_F44, 7)   ' Led FB bits for Gruppe 7
  Call ShowFbGroup(F45_F52, 8)   ' Led FB bits for Gruppe 8
  Call ShowFbGroup(F53_F60, 9)   ' Led FB bits for Gruppe 9
  Call ShowFbGroup(F61_F68, 10)  ' Led FB bits for Gruppe 10
End Sub




'=====================================================================================================================
' Support routines below
'---------------------------------------------------------------------------------------------------------------------
' Lok Functions
Private Sub handleFunctionBefehl(ByVal turnOn As Boolean)
  If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
  If Not IsNumeric(TxtLokFNumber.Text) Then Exit Sub
  Dim FbNumber As Byte
  FbNumber = TxtLokFNumber.Text
  If FbNumber < 0 Then Exit Sub
  If FbNumber > 68 Then Exit Sub
  Dim group As Byte
  Dim bitposition As Byte
  ' F0 needs to be consideren as speciaal case
  If FbNumber = 0 Then group = 1: bitposition = 4
  ' The other functions are coded in a consistent way
  Select Case FbNumber
    Case 1 To 4: group = 1: bitposition = FbNumber - 1
    Case 5 To 8: group = 2: bitposition = FbNumber - 5
    Case 9 To 12: group = 3: bitposition = FbNumber - 9
    Case 13 To 20: group = 4: bitposition = FbNumber - 13
    Case 21 To 28: group = 5: bitposition = FbNumber - 21
    Case 29 To 36: group = 6: bitposition = FbNumber - 29
    Case 37 To 44: group = 7: bitposition = FbNumber - 37
    Case 45 To 52: group = 8: bitposition = FbNumber - 45
    Case 53 To 60: group = 9: bitposition = FbNumber - 53
    Case 61 To 68: group = 10: bitposition = FbNumber - 61
  End Select
  FB(group) = bitUpdate(FB(group), bitposition, turnOn)
  Call XpressNet.CmdLokFunktionBefehl(txtLokAddress.Text, group, FB(group))
End Sub

Private Sub ClearLedsForFunctions()
  Dim i As Integer
  For i = 0 To 68
    LEDF0F68(i).FillColor = vbWhite
  Next
End Sub

Private Function FbColor(ByVal Value As Boolean) As ColorConstants
  If Value Then
    FbColor = vbRed
  Else
    FbColor = vbGreen
  End If
End Function

'---------------------------------------------------------------------------------------------------------------------
' Functions to show the Feedbacks on this form
Private Sub ShowFbGroup(ByVal data As Byte, ByVal group As Byte)
  Dim bitsInGroup As Byte               ' Groups 1..3 contain 4 FB bits, the others contain 8 FB bits
  Dim offset As Byte                    ' Offset in the VB Array for a certain group
  If group < 4 Then
    bitsInGroup = 4
    offset = group * bitsInGroup - 4
  Else
    bitsInGroup = 8
    offset = group * bitsInGroup - 20
  End If
  Dim i As Byte
  For i = 1 To bitsInGroup
    LEDF0F68(i + offset).FillColor = FbColor(bitValue(data, i - 1))
  Next
  ' FB0 in Group 1 is a special case
  If group = 1 Then LEDF0F68(0).FillColor = FbColor(bitValue(data, 4))
End Sub

'---------------------------------------------------------------------------------------------------------------------
' Some bit manipulation functions
Private Function bitValue(ByVal data As Byte, ByVal index As Byte) As Boolean
  index = 2 ^ index
  bitValue = (data And index) / index
End Function

Private Function bitSet(ByVal data As Byte, ByVal index As Byte) As Byte
  index = 2 ^ index
  bitSet = data Or index
End Function

Private Function bitClear(ByVal data As Byte, ByVal index As Byte) As Byte
  index = 2 ^ index              ' Set the bit at position index (other bist are 0)
  index = 255 - index            ' Flip all bits
  bitClear = data And index      ' mask the data
End Function

Private Function bitUpdate(ByVal data As Byte, ByVal index As Byte, ByVal bitIsSet As Byte) As Byte
  If bitIsSet Then
    bitUpdate = bitSet(data, index)
  Else
    bitUpdate = bitClear(data, index)
  End If
End Function

'---------------------------------------------------------------------------------------------------------------------
' Transforms numbers form 0..15 (4 bits) in a bitstring
Private Function printBits(ByVal data As Byte) As String
  Dim outString As String
  Dim i As Integer
  For i = 0 To 3
    outString = (data Mod 2) & outString
    data = data \ 2
  Next
  printBits = outString
End Function

