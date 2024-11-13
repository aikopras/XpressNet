VERSION 5.00
Begin VB.Form FormStart 
   Caption         =   "Start"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   ScaleHeight     =   1740
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewZ21 
      Caption         =   "Add Z21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdNewXpressNet 
      Caption         =   "Add XpressNet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Schakel wissel 1 om"
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private XpnConnections() As FormXpressNet
Private Z21Connections() As FormZ21
Private XpnIndex As Integer
Private Z21Index As Integer


Private Sub Form_Load()
  XpnIndex = 1
  Z21Index = 1
End Sub

Private Sub cmdNewXpressNet_Click()
  XpnIndex = XpnIndex + 1
  ReDim Preserve XpnConnections(XpnIndex)
  Set XpnConnections(XpnIndex) = New FormXpressNet
  XpnConnections(XpnIndex).Show
End Sub

Private Sub cmdNewZ21_Click()
  Z21Index = Z21Index + 1
  ReDim Preserve Z21Connections(Z21Index)
  Set Z21Connections(Z21Index) = New FormZ21
  Z21Connections(Z21Index).Show
End Sub

