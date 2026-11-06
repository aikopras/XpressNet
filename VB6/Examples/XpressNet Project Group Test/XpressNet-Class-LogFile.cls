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

' Version 1.1.0: Logfile directory can be changed by the application program
' Version 2.0:
' - the default directory is %PUBLIC%\TMC-Logfiles (1.3: %PUBLIC%). On some Windows 11 systems writing directly
'   in %PUBLIC% gives an access error; a subdirectory can be written. The library creates the directory.
' - if the logfile can not be written (for example: no access rights), the text is lost without an error
'   (1.3: run time error)
' - every line of an XpressNet object starts with its letter (A, B, ...; see XpressNetModule); when a connection is
'   made, a line "A = 192.168.0.200, port 5550, local port 54321" tells which connection the letter belongs to.
'   1.3: every line ended with the IP address or COM port (myHost, which is still set)
' - the description of a sent message starts with "=>", of a received message with "<=" (as the hex values)

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long


' The maximum Logfile size as well as the names of the Logfile and backups may be modifief
Const maxLogFileSize As Long = 100000                      ' Size is in bytes

Private m_LogDir As String
Public logFileName As String
Public logFileNameOld1 As String
Public logFileNameOld2 As String
Public logFileNameOld3 As String



' THE FOLLOWING METHODS ARE DEFINED:
' OpenNew                            ' Open a new Logfile / overwrite an existing Logfile
' Append          some Text          ' Append "some text" to the Logfile
' AppendWithHost  some text          ' Append "some text", preceded by the letter of the XpressNet object
' Connection      some text          ' 2.0: Append "= some text" (the connection of the XpressNet object)
' SendHex         byteArray          ' Append "=>" followed by the hex values of byteArray
' ReveiveHex      length, byeArray   ' Append "<=" followed by the hex values upto length
' ReveiveHexRange byeArray, from, to ' Append "<=" followed by the hex values between from and to
'
' Details         some Text          ' Same as AppendWithHost, but checks Parent.LogDetails first
' DetailsSent     some Text          ' 2.0: as Details, preceded by "=>" (description of a sent message)
' DetailsReceived some Text          ' 2.0: as Details, preceded by "<=" (description of a received message)
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

Private Sub Class_Initialize()
    ' In case the application doesn't provide a path, use the following defaults
    m_LogDir = Environ$("PUBLIC")     ' Use the Windows Environ variable
    If Len(m_LogDir) = 0 Then
        m_LogDir = "C:\Users\Public"  ' Works on several systems, but may require admin rights
    End If
    ' 2026-11 modified: a subdirectory of PUBLIC, created if it does not exist yet (see the header of this file)
    m_LogDir = m_LogDir & "\TMC-Logfiles"
    On Error Resume Next
    If Dir$(m_LogDir, vbDirectory) = vbNullString Then MkDir m_LogDir
    Err.Clear
    On Error GoTo 0
    logFileName = m_LogDir & "\TMC-LogFile.txt"
    logFileNameOld1 = m_LogDir & "\TMC-LogFile-1.txt"
    logFileNameOld2 = m_LogDir & "\TMC-LogFile-2.txt"
    logFileNameOld3 = m_LogDir & "\TMC-LogFile-3.txt"
End Sub


Public Sub OpenLogfile()
  ' The goal of this subroutine is to check the size of the existing Logfile.
  ' If that size exceeds maxLogFileSize, the Logfiles will be rotated and renamed.
  Dim fileSize As Long
  ' 2026-11 modified: errors (for example: no access rights) are ignored; the text is lost (was: run time error)
  On Error Resume Next
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
  ' 2026-11 modified: errors (for example: no access rights) are ignored; the text is lost (was: run time error)
  On Error Resume Next
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  Print #iFileNo, TimeText & ": " & data
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

Public Sub Details(ByVal data As String)
  If Parent.LogDetails Then AppendWithHost (data)
End Sub

' 2026-11 modified: new. The description of a message gets the same direction marker as its hex values
Public Sub DetailsSent(ByVal data As String)
  Details "=> " & data
End Sub

Public Sub DetailsReceived(ByVal data As String)
  Details "<= " & data
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
  ' 2026-11 modified: errors (for example: no access rights) are ignored; the text is lost (was: run time error)
  On Error Resume Next
  GetLocalTime MyTime
  TimeText = MyTime.wYear & "/" & Format(MyTime.wMonth, "00") & "/" & Format(MyTime.wDay, "00") & " " & Format(MyTime.wHour, "00") & ":" & Format(MyTime.wMinute, "00") & ":" & Format(MyTime.wSecond, "00") & "." & Format(MyTime.wMilliseconds, "000")
  iFileNo = FreeFile
  Open logFileName For Append As #iFileNo
  ' 2026-11 modified: preceded by the letter of the XpressNet object (was: followed by the host)
  'Print #iFileNo, TimeText & ": " & data & " (" & myHost & ")"
  Print #iFileNo, TimeText & ": " & Parent.instanceId & " " & data
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

' 2026-11 modified: new. The connection of this XpressNet object
Public Sub Connection(ByVal data As String)
  AppendWithHost "= " & data
End Sub

Public Sub SendHex(ByRef byteArray() As Byte)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  ' 2026-11 modified: errors (for example: no access rights) are ignored; the text is lost (was: run time error)
  On Error Resume Next
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
  ' 2026-11 modified: preceded by the letter of the XpressNet object (was: followed by the host)
  'Print #iFileNo, TimeText & ": => " & hexString & "(" & myHost & ")"
  Print #iFileNo, TimeText & ": " & Parent.instanceId & " => " & RTrim$(hexString)
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub


Public Sub ReceiveHexRange(ByRef byteArray() As Byte, ByVal msgStart As Byte, ByVal msgEnd As Byte)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  ' 2026-11 modified: errors (for example: no access rights) are ignored; the text is lost (was: run time error)
  On Error Resume Next
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
  ' 2026-11 modified: preceded by the letter of the XpressNet object (was: followed by the host)
  'Print #iFileNo, TimeText & ": <= " & hexString & "(" & myHost & ")"
  Print #iFileNo, TimeText & ": " & Parent.instanceId & " <= " & RTrim$(hexString)
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub

'Public Sub ReceiveHex(ByVal length As Byte, ByRef byteArray() As Byte)
Public Sub ReceiveHex(ByRef byteArray() As Byte)
  Dim iFileNo As Integer
  Dim MyTime As SYSTEMTIME
  Dim TimeText As String
  ' 2026-11 modified: errors (for example: no access rights) are ignored; the text is lost (was: run time error)
  On Error Resume Next
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
  ' 2026-11 modified: preceded by the letter of the XpressNet object (was: followed by the host)
  'Print #iFileNo, TimeText & ": <= " & hexString & "(" & myHost & ")"
  Print #iFileNo, TimeText & ": " & Parent.instanceId & " <= " & RTrim$(hexString)
  'close the file (if you dont do this, you wont be able to open it again!)
  Close #iFileNo
End Sub


' Added in Version 1.1.0
Public Property Let LogDirectory(ByVal value As String)
    Dim d As String
    d = Trim$(value)
    If Len(d) = 0 Then Exit Property
    ' trailing backslash weghalen
    If Right$(d, 1) = "\" Then d = Left$(d, Len(d) - 1)
    ' check folder exists
    If Dir$(d, vbDirectory) = vbNullString Then
        On Error Resume Next
        MkDir d
        If Err.number <> 0 Then
            ' Folder creation failed -> keep old m_LogDir and defaults
            Err.Clear
            On Error GoTo 0
            Exit Property
        End If
        On Error GoTo 0
    End If
    ' Once we are here, the folder already exists or has been created
    m_LogDir = d
    logFileName = m_LogDir & "\TMC-LogFile.txt"
    logFileNameOld1 = m_LogDir & "\TMC-LogFile-1.txt"
    logFileNameOld2 = m_LogDir & "\TMC-LogFile-2.txt"
    logFileNameOld3 = m_LogDir & "\TMC-LogFile-3.txt"
End Property



' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property
