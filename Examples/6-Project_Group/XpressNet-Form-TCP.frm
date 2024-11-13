VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmTcp 
   Caption         =   "Form TCP"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long

' RESPONSES TO CONNECTION RELATED EVENTS
' WE CALL THE EVENT FORWARDING (AND LOGGING) SUBROUTINES IN THE MAIN XpressNet OBJECT
Private Sub tcpClient_Connect()
  Parent.Parent.eventInterfaceConnected
End Sub

Private Sub tcpClient_Close()
  Parent.Parent.eventTcpClosedByRemoteHost
End Sub

Private Sub tcpClient_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Parent.Parent.eventTcpError (Description)
End Sub
  
  
' RESPONSES TO DATA RECEPTION RELATED EVENTS
' WE CALL A ROUTINE IN THE TCP CLASS
Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
  Parent.receiveTcpSegment (bytesTotal)
End Sub


' The parent properety below is needed to pass a pointer of the parent TcpClass object to this object.
Public Property Get Parent() As TcpClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As TcpClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property

