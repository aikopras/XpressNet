VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmUdp 
   Caption         =   "Form UDP"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock udpClient 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "FrmUdp"
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
''Private Sub udpClient_Connect()
''  Parent.Parent.eventInterfaceConnected
''End Sub

''Private Sub udClient_Close()
''  Parent.Parent.eventTcpClosedByRemoteHost
''End Sub

Private Sub udpClient_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  'Parent.Parent.eventTcpError (Description)
End Sub
  
  
' RESPONSES TO DATA RECEPTION RELATED EVENTS
' WE CALL A ROUTINE IN THE UDP CLASS
Private Sub udpClient_DataArrival(ByVal bytesTotal As Long)
  ' It seems the bytesReceived parameter is not always correct
  ' Therefore we ignore this parameter. The called subroutine should use UBound instead
   Call Parent.receiveUdpPacket
End Sub


' The parent properety below is needed to pass a pointer of the parent TcpClass object to this object.
Public Property Get Parent() As UdpClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As UdpClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property


