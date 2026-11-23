Attribute VB_Name = "XpressNetModule"
' 2026-11 modified: new module (version 2.0).
' The letter of an XpressNet object in the logfile: A for the first XpressNet object of the program, B for the
' second, and so on (after Z: AA, AB, ...). A class module can not have a variable that is shared by all objects,
' a standard module can. Two programs that write in the same logfile both start with A
Option Explicit

Private instanceCounter As Long

Public Function nextInstanceId() As String
  Dim n As Long
  Dim s As String
  instanceCounter = instanceCounter + 1
  n = instanceCounter
  s = ""
  While n > 0
    n = n - 1
    s = Chr$(Asc("A") + (n Mod 26)) & s
    n = n \ 26
  Wend
  nextInstanceId = s
End Function
