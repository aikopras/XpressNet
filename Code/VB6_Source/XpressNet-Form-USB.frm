VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmUsb 
   Caption         =   "Form USB"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm usbClient 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputMode       =   1
   End
End
Attribute VB_Name = "FrmUsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long


' RESPONSES TO DATA RECEPTION RELATED EVENTS
' WE CALL A ROUTINE IN THE USB CLASS
Private Sub usbClient_OnComm()
  Parent.handleUsbEvent
End Sub


' The parent properety below is needed to pass a pointer of the parent TcpClass object to this object.
Public Property Get Parent() As UsbClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As UsbClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property


