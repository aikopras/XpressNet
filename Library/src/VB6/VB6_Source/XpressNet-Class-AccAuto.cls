VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccAutoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Implementation of the AccessoryAuto method
'
' This method can be used to set multiple switches at once, without worrying about deactivation
' and scheduling. The standard XpressNet command for accesories include an activation and deactivation
' bit, which allows activation of the decoder's output port and, after a short interval, deactivation.
' Such deactivation may, for example, be necessary for coil or motor based switch engines.
' In practice, coild based switch engines need to be deactivated after several hundreds of ms.
' Another concern with switch decoders is that, after activation, a single output may consume 1 A
' of current. If all switches would activate their outputs at exactly the same moment, the amount
' of current that is needed may easily exceed the amount that is available. Therefore it is importent
' to schedule switch commands after each other.
'
' For the interval between activation and deactivation, as well as for the interval between two
' activations, the "AccessoryInterval" variable (declared in the XpressNet class) will be used.
' 2026-10 modified: after a deactivation there is a pause of AccessoryInterval / 2, before the next accessory
' is activated (as in the example of the Z21 specification, 5.2.1). Without flow control (Z21), the next
' activation followed the deactivation immediately.
'
' Like the "message" subroutine, this method implements a FIFO queue to store accessory commands
' As soon as an accessory command is received, the accessory address and position is saved in the FIFO.
' Next, the associated XpressNet accessory command, with the activation bit set, is issued, and a
' timer is started. Once the timer fires, the XpressNet command is send again, but this time the
' activation bit is released.
' If, in the time between activation and deactivation, additional accessory commands are received,
' these are stored on the FIFO queue as well.
'
' The FIFO code that is used in this method is similar to the FIFO code as used in the "message"
' and "nextMessage" methods.
'
' This method is implemented in a separate class, since it is, strictly speaking, not part of the
' XpressNet interface, but rather a layer on top of the XpressNet DLL. Still it should be instantiated
' in, and called from the XpressNet class object.
'
' Version 2.0 (November 2026), the same in the VB6, .NET and macOS versions:
' - emptyFifo (after a response timeout, or after a pause without 01 07) first switches off the output of the
'   switch that is activated at that moment. In version 1.3 this switch got no deactivation.
' - During a pause (see XpressNet-Class-Send.cls: the command station does not address the interface) no new switch
'   is activated. The activation of the next switch waits for 01 07 (resumeAfterPause).
' - A full queue gives the event SendError (besides the line in the logfile).
' - The interval of AccessoryInterval still starts at the moment the activation is offered to the SendClass
'   (as in version 1.3). If the activation waits in the FIFO of the SendClass, the output is active shorter.
'
' ----------------------------------------------------------------------------------------------------
Option Explicit

' The lines below are needed to pass a pointer of the main object to this subclass.
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private m_lParentWeakRef    As Long


' ----------------------------------------------------------------------------------------------------
' We store accessory commands in a FIFO. The FIFO is a two-dimensional Array of bytes
' Each Array row holds two fields: the address and the position of the switch.
' The maximum number of commands is set by maxFifoEntries.
' maxFifoMsgSize defines the maximum size of individual FIFO entries. This size should at least be the
' maximum XpressNet message size (including Framing and XOR bytes), plus one byte to store the actual
' XpressNet message size. This actual messages size is stored as first element of an Array row
Private Const maxFifoEntries = 32                 ' The number of Accessory messages the queue is able to hold
Private Type AccessoryDataType
    address As Integer
    position As Byte
End Type
Private head As Byte                              ' we take elements from the head of the queue
Private pausing As Boolean                        ' 2026-10 modified: is the timer used for the pause after a deactivation?
Private switchActivated As Boolean                ' 2026-11 modified: is the output of the switch at the head active?
Private waitingForInterface As Boolean            ' 2026-11 modified: does an activation wait for the end of a pause?
Private tail As Byte                              ' tail points to the position of the most recent element
Private numberOfFifoMessages As Byte              ' number of messages the queue currently holds
Private fifo(maxFifoEntries) As AccessoryDataType ' this is the queue

' For timers we instantiate objects of the XTimerClass, which was copied from the Internet
Private WithEvents deactTimer As XTimerClass
Attribute deactTimer.VB_VarHelpID = -1

Public Sub handle(ByVal address As Integer, ByVal position As Byte)
  ' Check if the FIFO queue has free slots left to store the message
  ' If not, increase "maxFifoEntries", or modify the calling program
  If numberOfFifoMessages >= maxFifoEntries Then  ' FIFO overflow :-(
    Parent.Logfile.Errors "Error: Auto Accessory queue is full"
    ' 2026-11 modified: event SendError
    Call Parent.eventSendError(FifoVoll, "Error: Auto Accessory queue is full (CmdAccessoryAuto: switch " & address & ", position " & position & ")")
    Exit Sub
  End If
  numberOfFifoMessages = numberOfFifoMessages + 1
  If numberOfFifoMessages > 1 Then                ' update tail = position the data is stored
    tail = tail + 1
    tail = tail Mod maxFifoEntries
  End If
  fifo(tail).address = address                    ' store address and position to the FIFO
  fifo(tail).position = position
  ' If this is the only accessory command in the FIFO, we may send it right now
  ' 2026-10 modified: but not during the pause after a deactivation (the accessory is activated after the pause)
  'If numberOfFifoMessages = 1 Then
  If (numberOfFifoMessages = 1) And (Not pausing) Then
    ' 2026-11 modified: activateHead (not during a pause of the SendClass)
    'Call Parent.CmdAccessory(address, position, 1)  ' activate the accessory
    'deactTimer.Interval = Parent.AccessoryInterval  ' set the deactivation timer's interval
    'deactTimer.Enabled = True                       ' start the timer
    Call activateHead                               ' activate the accessory, and start the timer
  End If
End Sub

' 2026-11 modified: new. Activates the output of the switch at the head of the FIFO, and starts the deactivation timer
' (1.3: the same code in handle and deactTimer_Timer). Not during a pause of the SendClass (01 05 / 01 08):
' the activation then waits for resumeAfterPause (01 07)
Private Sub activateHead()
  If Parent.send.interfaceNotAddressed Then
    waitingForInterface = True
    Exit Sub
  End If
  Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 1)  ' activate the accessory
  switchActivated = True
  deactTimer.Interval = Parent.AccessoryInterval  ' set the deactivation timer's interval
  deactTimer.Enabled = True                       ' start the timer
End Sub

' 2026-11 modified: new. Called by the SendClass at the end of a pause (01 07). An activation that waited, is done now
Public Sub resumeAfterPause()
  If Not waitingForInterface Then Exit Sub
  waitingForInterface = False
  If (numberOfFifoMessages > 0) And (Not pausing) And (Not switchActivated) Then Call activateHead
End Sub

Private Sub Class_Initialize()
  ' 2026-10 modified: the timer is created first, since emptyFifo stops the timer
  'Call emptyFifo
  'Set deactTimer = New XTimerClass                ' Create a new timer object
  Set deactTimer = New XTimerClass                ' Create a new timer object
  Call emptyFifo
End Sub

Private Sub deactTimer_Timer(ByVal Seconds As Currency)
  ' STEP 1: If the timer expires, we need to send the accessory command to deactivate the output port
  deactTimer.Enabled = False
  ' 2026-10 modified: if the pause after a deactivation is over (see STEP 3), activate the port of the next
  ' accessory (if the FIFO is not empty), and start the timer
  If pausing Then
    pausing = False
    If numberOfFifoMessages > 0 Then
      ' 2026-11 modified: activateHead (not during a pause of the SendClass)
      'Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 1)
      'deactTimer.Interval = Parent.AccessoryInterval  ' set the deactivation timer's interval
      'deactTimer.Enabled = True                       ' start the timer
      Call activateHead
    End If
    Exit Sub
  End If
  If numberOfFifoMessages = 0 Then Exit Sub       ' 2026-11 modified: should not occur (emptyFifo stops the timer)
  switchActivated = False                         ' 2026-11 modified
  Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 0)
  ' STEP 2: After deactivation, we should remove this accesory from the FIFO
  numberOfFifoMessages = numberOfFifoMessages - 1
  If numberOfFifoMessages > 0 Then
    head = head + 1
    head = head Mod maxFifoEntries
  End If
  ' STEP 3: If the FIFO is not empty, Activate the port of the next accessory, and start the timer
  ' 2026-10 modified: first a pause of AccessoryInterval / 2. The next accessory is activated after the pause
  'If numberOfFifoMessages > 0 Then
  '  Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 1)
  '  deactTimer.Interval = Parent.AccessoryInterval  ' set the deactivation timer's interval
  '  deactTimer.Enabled = True                       ' start the timer
  'End If
  pausing = True
  deactTimer.Interval = Parent.AccessoryInterval \ 2 ' set the timer's interval for the pause
  deactTimer.Enabled = True                         ' start the timer
End Sub

Public Function size() As Byte
  size = numberOfFifoMessages
End Function

' 2026-11 modified: the output of an activated switch is switched off first (1.3: this switch got no deactivation).
' Returns the number of removed commands (without the switch that has been switched off). Was a Sub
'Public Sub emptyFifo()
Public Function emptyFifo() As Integer
  Dim removed As Integer
  removed = numberOfFifoMessages
  If switchActivated And (numberOfFifoMessages > 0) Then
    switchActivated = False
    Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 0)
    removed = removed - 1
  End If
  switchActivated = False
  waitingForInterface = False
  emptyFifo = removed
  head = 0
  tail = 0
  numberOfFifoMessages = 0
  ' 2026-10 modified: stop the deactivation timer as well. Before, the timer kept running with an empty FIFO,
  ' which resulted in an overflow error (numberOfFifoMessages is a Byte) and a deactivation of fifo(0)
  deactTimer.Enabled = False
  pausing = False
End Function






' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property

