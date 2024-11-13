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
  If numberOfFifoMessages = 1 Then
    Call Parent.CmdAccessory(address, position, 1)  ' activate the accessory
    deactTimer.Interval = Parent.AccessoryInterval  ' set the deactivation timer's interval
    deactTimer.Enabled = True                       ' start the timer
  End If
End Sub

Private Sub Class_Initialize()
  Call emptyFifo
  Set deactTimer = New XTimerClass                ' Create a new timer object
End Sub

Private Sub deactTimer_Timer(ByVal Seconds As Currency)
  ' STEP 1: If the timer expires, we need to send the accessory command to deactivate the output port
  deactTimer.Enabled = False
  Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 0)
  ' STEP 2: After deactivation, we should remove this accesory from the FIFO
  numberOfFifoMessages = numberOfFifoMessages - 1
  If numberOfFifoMessages > 0 Then
    head = head + 1
    head = head Mod maxFifoEntries
  End If
  ' STEP 3: If the FIFO is not empty, Activate the port of the next accessory, and start the timer
  If numberOfFifoMessages > 0 Then
    Call Parent.CmdAccessory(fifo(head).address, fifo(head).position, 1)
    deactTimer.Interval = Parent.AccessoryInterval  ' set the deactivation timer's interval
    deactTimer.Enabled = True                       ' start the timer
  End If
End Sub

Public Function size() As Byte
  size = numberOfFifoMessages
End Function

Public Sub emptyFifo()
  head = 0
  tail = 0
  numberOfFifoMessages = 0
End Sub






' The parent properety below is needed to pass a pointer of the main object to this subclass.
Public Property Get Parent() As XpressNetClass
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

Friend Property Set Parent(oValue As XpressNetClass)
    m_lParentWeakRef = ObjPtr(oValue)
End Property

