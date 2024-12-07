VERSION 5.00
Begin VB.Form FormZ21 
   Caption         =   "TestZ21"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameLoks 
      Caption         =   "Loks"
      Height          =   3855
      Left            =   4440
      TabIndex        =   44
      Top             =   4080
      Width           =   4455
      Begin VB.CommandButton CmdAlleLoksAnhalten 
         Caption         =   "Alle Loks anhalten"
         Height          =   495
         Left            =   2160
         TabIndex        =   57
         Top             =   3120
         Width           =   1600
      End
      Begin VB.CommandButton CmdEineLokAnhalten 
         Caption         =   "Anhalten"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtLokAddress 
         Height          =   375
         Left            =   840
         TabIndex        =   55
         Text            =   "3"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtLokSpeed 
         Height          =   375
         Left            =   840
         TabIndex        =   54
         Text            =   "0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox ComboLokSteps 
         Height          =   315
         Left            =   840
         TabIndex        =   53
         Text            =   "28"
         Top             =   3120
         Width           =   735
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
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   615
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
         Left            =   960
         TabIndex        =   51
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox TxtLokFNumber 
         Height          =   375
         Left            =   2400
         TabIndex        =   50
         Text            =   "1"
         Top             =   360
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
         Left            =   2880
         TabIndex        =   49
         Top             =   360
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
         Left            =   3480
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdGetLokInfo 
         Caption         =   "Get"
         Height          =   375
         Left            =   3960
         TabIndex        =   47
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton CmdLokEnter 
         Caption         =   "Enter"
         Height          =   375
         Left            =   840
         TabIndex        =   46
         Top             =   840
         Width           =   735
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
         Left            =   1680
         TabIndex        =   45
         Top             =   360
         Width           =   255
      End
      Begin VB.Label LabelLokAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   165
         TabIndex        =   65
         Top             =   360
         Width           =   570
      End
      Begin VB.Label LabelLokSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Speed"
         Height          =   195
         Left            =   270
         TabIndex        =   64
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label LabelLokSteps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Steps"
         Height          =   195
         Left            =   210
         TabIndex        =   63
         Top             =   3120
         Width           =   405
      End
      Begin VB.Label LabelFunktion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F:"
         Height          =   240
         Left            =   2235
         TabIndex        =   62
         Top             =   405
         Width           =   165
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   3120
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   3360
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   3480
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   3600
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   3720
         Top             =   840
         Width           =   135
      End
      Begin VB.Label LabelFunctions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F0-F4"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   61
         Top             =   825
         Width           =   405
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   2760
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   2880
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   3000
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   8
         Left            =   3120
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   9
         Left            =   3360
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   10
         Left            =   3480
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   11
         Left            =   3600
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   12
         Left            =   3720
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   13
         Left            =   2760
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   14
         Left            =   2880
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   15
         Left            =   3000
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   16
         Left            =   3120
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   17
         Left            =   3360
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   18
         Left            =   3480
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   19
         Left            =   3600
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   20
         Left            =   3720
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   21
         Left            =   2760
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   22
         Left            =   2880
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   23
         Left            =   3000
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   24
         Left            =   3120
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   25
         Left            =   3360
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   26
         Left            =   3480
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   27
         Left            =   3600
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape LEDF0F68 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   28
         Left            =   3720
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label LabelFunctions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F5-F12"
         Height          =   195
         Index           =   1
         Left            =   2190
         TabIndex        =   60
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label LabelFunctions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F13-F20"
         Height          =   195
         Index           =   2
         Left            =   2070
         TabIndex        =   59
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label LabelFunctions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F21-F28"
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   58
         Top             =   1560
         Width           =   585
      End
   End
   Begin VB.Frame FrameVersion 
      Height          =   855
      Left            =   7920
      TabIndex        =   39
      Top             =   120
      Width           =   2775
      Begin VB.Label labeFirmwareVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Z21 Firmware Version:"
         Height          =   195
         Left            =   75
         TabIndex        =   43
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label txtXZ21Firmware 
         Height          =   255
         Left            =   1700
         TabIndex        =   42
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LabelMasterVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Master Version:"
         Height          =   195
         Left            =   600
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label txtMasterVersion 
         Height          =   255
         Left            =   1700
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameAnAus 
      Caption         =   "An / Aus"
      Height          =   1815
      Left            =   5640
      TabIndex        =   36
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton CmdAllesAn 
         Caption         =   "Alles An"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1600
      End
      Begin VB.CommandButton CmdAllesAus 
         Caption         =   "Alles Aus"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1600
      End
   End
   Begin VB.Frame FrameNumberOfmessages 
      Height          =   615
      Left            =   7920
      TabIndex        =   32
      Top             =   1080
      Width           =   2775
      Begin VB.Label LabelZ21tMessageCounter 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Z21Messages:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label txtNmbPckts 
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FrameExtAcc 
      Caption         =   "Extended Accessory"
      Height          =   1935
      Left            =   4440
      TabIndex        =   26
      Top             =   2040
      Width           =   2175
      Begin VB.TextBox TxtExtAccSwitch 
         Height          =   375
         Left            =   1080
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtExtAccValue 
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CmdExtAccSet 
         Caption         =   "Set"
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LabelExtSwitch 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   240
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   765
      End
      Begin VB.Label LabelExtValue 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   240
         Left            =   405
         TabIndex        =   30
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.Frame FrameAcc 
      Caption         =   "Accessory"
      Height          =   6495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4215
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
         TabIndex        =   19
         Top             =   360
         Width           =   735
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
         TabIndex        =   18
         Top             =   360
         Width           =   735
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
         TabIndex        =   17
         ToolTipText     =   "1..2048"
         Top             =   360
         Width           =   1212
      End
      Begin VB.CommandButton cmdSwitchFeedback 
         Caption         =   "Feedback"
         Height          =   495
         Left            =   960
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtBoxSwitchAddress 
         Height          =   3975
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TxtBoxSwitchData 
         Height          =   3975
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtBoxSwitchTtBits 
         Height          =   3975
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "C=Command Station, D=Decoder"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox TxtBoxSwitchIBit 
         Height          =   3975
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox CheckActivate 
         Caption         =   "activate"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox CheckAutoDeactivate 
         Caption         =   "auto deactivate"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1695
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
         TabIndex        =   25
         Top             =   6120
         Width           =   2025
      End
      Begin VB.Label LabelSwitch 
         AutoSize        =   -1  'True
         Caption         =   "Switch:"
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LabelSwitchAddress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   240
         Left            =   480
         TabIndex        =   23
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label LabelSwitchdata 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   240
         Left            =   1920
         TabIndex        =   22
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label LabelSwitchTtBits 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "TT"
         Height          =   240
         Left            =   3000
         TabIndex        =   21
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label LabelSwitchIBit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "I"
         Height          =   240
         Left            =   3600
         TabIndex        =   20
         Top             =   1800
         Width           =   195
      End
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   1815
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   2175
      Begin VB.Shape ledNotHalt 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Shape           =   3  'Circle
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape ledNotAus 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape ledProgrammiermode 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label LabelNotHalt 
         AutoSize        =   -1  'True
         Caption         =   "NotHalt"
         Height          =   240
         Left            =   495
         TabIndex        =   8
         Top             =   870
         Width           =   675
      End
      Begin VB.Label LabelNotAus 
         AutoSize        =   -1  'True
         Caption         =   "NotAus"
         Height          =   240
         Left            =   495
         TabIndex        =   7
         Top             =   390
         Width           =   660
      End
      Begin VB.Label LabelProgrammiermode 
         AutoSize        =   -1  'True
         Caption         =   "Programmier Mode"
         Height          =   240
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1380
      End
   End
   Begin VB.Frame FrameUDP 
      Caption         =   "Z21 Network Settings"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmdUdpSet 
         Caption         =   "Set"
         Height          =   375
         Left            =   1800
         TabIndex        =   35
         ToolTipText     =   "Set IP address and Port"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtIP 
         Height          =   372
         Left            =   1275
         TabIndex        =   2
         Text            =   "192.168.0.111"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtUdpPort 
         Height          =   372
         Left            =   1275
         TabIndex        =   1
         Text            =   "21105"
         ToolTipText     =   "Yamorc: 21105 / Z21: 21106"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label LabelIpAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   990
      End
      Begin VB.Label LabelUdpPort 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UDP Port"
         Height          =   195
         Left            =   450
         TabIndex        =   3
         ToolTipText     =   "Yamorc: 21105 / Z21: 21106"
         Top             =   840
         Width           =   675
      End
   End
End
Attribute VB_Name = "FormZ21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents XpressNet As XpressNetClass
Attribute XpressNet.VB_VarHelpID = -1

Private Deactivate As Boolean          ' Used for switches
Private AutoDeactivate As Boolean      ' Used for switches

' ----------------------------------------------------------------------------------------------------------------------
' Function Groups:
' For each Lok Address, the values must be fethed again
Private FB(10) As Byte                 ' FB: Function Befehl (on / off)
Private FS(10) As Byte                 ' Function Status (tastend oder nicht tastend)



Public Sub Form_Load()
  ' Normal Set command, needed in cases the DLL was registered, or included as second group
  ' Set XpressNet = New XpressNetClass
  '
  ' Set command that is needed if the DLL is distributed next to the exe.
  Set XpressNet = cRegFree.CreateObj("XpressNetClass", "XpressNet.dll")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set XpressNet = Nothing
End Sub


'===============================================================================================================
Private Sub cmdUdpSet_Click()
  Call XpressNet.CmdInitUdp(txtIP.Text, txtUdpPort.Text)
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

Private Sub CmdGetLokInfo_Click()
  Call XpressNet.CmdLokInfoAnfordern(txtLokAddress.Text)
End Sub


'===============================================================================================================
Private Sub cmdRecht_Click()
  If Not IsNumeric(txtWissel.Text) Then Exit Sub
  If CheckAutoDeactivate.Value Then
    Call XpressNet.CmdAccessoryAuto(txtWissel.Text, 0)
  Else
    Call XpressNet.CmdAccessory(txtWissel.Text, 0, CheckActivate.Value)
    If CheckActivate.Value Then
      CheckActivate.Value = 0
    Else:
      CheckActivate.Value = 1
    End If
  End If
End Sub

Private Sub cmdAf_Click()
  If Not IsNumeric(txtWissel.Text) Then Exit Sub
  If CheckAutoDeactivate.Value Then
    Call XpressNet.CmdAccessoryAuto(txtWissel.Text, 1)
  Else
    Call XpressNet.CmdAccessory(txtWissel.Text, 1, CheckActivate.Value)
    If CheckActivate.Value Then
      CheckActivate.Value = 0
    Else:
      CheckActivate.Value = 1
    End If
  End If
End Sub

Private Sub cmdSwitchFeedback_Click()
  If IsNumeric(txtWissel.Text) Then
    Call XpressNet.CmdFeedbackRequest(txtWissel.Text, 0)
  End If
End Sub

Private Sub CmdExtAccSet_Click()
  If IsNumeric(TxtExtAccSwitch.Text) Then
    'Call XpressNet.CmdOpenDCCSetExtAccessory(TxtExtAccSwitch.Text, TxtExtAccValue.Text)
    Call XpressNet.CmdZ21SetExtAccessory(TxtExtAccSwitch.Text, TxtExtAccValue.Text)
  End If
End Sub




'===========================================================================================
' EVENTS
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
  ledNotHalt.FillColor = vbRed
End Sub

Private Sub XpressNet_ProgrammierMode()
  ledProgrammiermode.FillColor = vbRed
End Sub

Private Sub XpressNet_ZentraleVersion(ByVal version As String)
  txtMasterVersion = version
End Sub

Private Sub XpressNet_Z21FirmwareVersion(ByVal version As String)
  txtXZ21Firmware = version
End Sub

Private Sub XpressNet_Z21Lokinfo(ByVal LokAddress As Integer, ByVal Besetzt As Boolean, ByVal Speedsteps As Integer, ByVal Speed As Integer, ByVal Forward As Boolean, ByVal F0_F4 As Byte, ByVal F5_F8 As Byte, ByVal F9_F12 As Byte, ByVal F13_F20 As Byte, ByVal F21_F28 As Byte)
  ' Set the combobox and speed, based on the received input
  ComboLokSteps = Speedsteps
  TxtLokSpeed = Speed
  ' Store the values of F0..F12, to allow changing a single Function later
  FB(1) = F0_F4                  ' Store Gruppe 1
  FB(2) = F5_F8                  ' Store Gruppe 2
  FB(3) = F9_F12                 ' Store Gruppe 3
  FB(4) = F13_F20                ' Store Gruppe 4
  FB(5) = F21_F28                ' Store Gruppe 5
  ' Show the LEDs for F0 .. F28
  Call ShowFbGroup(F0_F4, 1)     ' Led FB bits for Gruppe 1
  Call ShowFbGroup(F5_F8, 2)     ' Led FB bits for Gruppe 1
  Call ShowFbGroup(F9_F12, 3)    ' Led FB bits for Gruppe 1
  Call ShowFbGroup(F13_F20, 4)   ' Led FB bits for Gruppe 4
  Call ShowFbGroup(F21_F28, 5)   ' Led FB bits for Gruppe 5
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
  For i = 0 To 28
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


