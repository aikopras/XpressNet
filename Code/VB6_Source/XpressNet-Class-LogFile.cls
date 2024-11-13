VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogFileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long


' The maximum Logfile size as well as the names of the Logfile and backups may be modifief
Const maxLogFileSize As Long = 100000                      ' Size is in bytes
Const logFileName As String = "C:\Users\Public\TMC-LogFile.txt"
Const logFileNameOld1 As String = "C:\TMC-LogFile-1.txt"
Const logFileNameOld2 As String = "C:\TMC-LogFile-2.txt"
Const logFileNameOld3 As String = "C:\TMC-LogFile-3.txt"



' THE FOLLOWING METHODS ARE DEFINED:
' OpenNew                            ' Open a new Logfile / overwrite an existing Logfile
' Append          some Text          ' Append "some text" to the Logfile
' AppendWithHost  some text          ' Append "some text", followed by the Host name / IP address (myHost)
' SendHex         byteArray          ' Append "=>" followed by the hex values of byteArray
' ReveiveHex      length, byeArray   ' Append "<=" followed by the hex values upto length
' ReveiveHexRange byeArray, from, to ' Append "<=" followed by the hex values between from and to
'
' Details         some Text          ' Same as AppendWithHost, but checks Parent.LogDetails first
' Errors          some Text          ' Same as AppendWithHost, but checks Parent.LogErrors first
' Warnings        some Text          ' Same as AppendWithHost, but checks Parent.LogWarnings first

'Declarations to use the Windows Time function. We need that, in order to be able to have ms precision
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public myHost As String              ' Allows the Logfile Hex dumps to include the IP address / USB port



Public Sub OpenLogfile()
  ' The goal of this subroutine is to check the size of the existing Logfile.
  ' If that size exceeds maxLogFileSize, the Logfiles will be rotated and renamed.
  Dim fileSize As Long
  If Dir$(logFileName) <> vbNullString Then
    fileSize = FileLen(logFileName)
    If fileSize > maxLogFileSize Then
      ' Delete the oldest file and rename the others
      If Dir$(logFileNameOld3) <> vbNullString Then Kill logFileNameOld3
      If Dir$(logFileNameOld2) <> vbNullString Then Name logFileNameOld2 As logFileNameOld3
      If Dir$(logFileNameOld1) <> vbNullString Then Name logFileNameOld1 As logFileNameOld2
      Name logFileName As logFileNameOld1
    End If
  End If
  ' Write some empty lines, to indicate a restart
  Dim iFileNo As Integer
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  Print #iFileNo, ""
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

Public Sub Append(ByVal data As String)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  ' OLD - TODO: KAN NA ENIGE TIJD WEG
  'Print #iFileNo, Format(Now, "yyyy/mm/dd hh:mm:ss") & ": " & data
  ' NEW (with ms)
  Print #iFileNo, TimeText & ": " & data
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

Public Sub Details(ByVal data As String)
  If Parent.LogDetails Then AppendWithHost (data)
End Sub

Public Sub Errors(ByVal data As String)
  If Parent.LogErrors Then AppendWithHost (data)
End Sub

Public Sub Warnings(ByVal data As String)
  If Parent.LogWarnings Then AppendWithHost (data)
End Sub

Public Sub AppendWithHost(ByVal data As String)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  Print #iFileNo, TimeText & ": " & data & " (" & myHost & ")"
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

Public Sub SendHex(ByRef byteArray() As Byte)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  'create a string with Hex values
  Dim hexString As String
  Dim i As Integer
  For i = 0 To UBound(byteArray)
    hexString = hexString & Hex(byteArray(i)) & " "
  Next
  'print it
  Print #iFileNo, TimeText & ": => " & hexString & "(" & myHost & ")"
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub


Public Sub ReceiveHexRange(ByRef byteArray() As Byte, ByVal msgStart As Byte, ByVal msgEnd As Byte)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  'create a string with Hex values
  Dim hexString As String
  Dim i As Integer
  For i = msgStart To msgEnd
    hexString = hexString & Hex(byteArray(i)) & " "
  Next
  'print it
  Print #iFileNo, TimeText & ": <= " & hexString & "(" & myHost & ")"
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

'Public Sub ReceiveHex(ByVal length As Byte, ByRef byteArray() As Byte)
Public Sub ReceiveHex(ByRef byteArray() As Byte)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  'create a string with Hex values
  Dim hexString As String
  Dim i As Integer
  For i = 0 To UBound(byteArray) - 1
    hexString = hexString & Hex(byteArray(i)) & " "
  Next
  'print it
  Print #iFileNo, TimeText & ": <= " & hexString & "(" & myHost & ")"
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub





' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property



