Option Explicit On
Imports XpressNet
' 2026-11: VB.NET version of the VB6 program RS-Bus Monitor (RsBusMonitor.frm), for XpressNet.NET 2.0.
' The layout is converted from the VB6 form; the code follows the VB6 code. Differences with VB6:
' - The events are connected with Handles (VB6: the name XpressNet_EventName).
' - The VB6 Shape controls are Led controls (Led.vb). The PanelManager draws in the Paint event of picPanel.
' - Default properties are written in full (for example LabelStatus.Text instead of LabelStatus).
' - Form_QueryUnload / Set XpressNet = Nothing is FormClosed / XpressNet.Dispose().
' As in VB6, the IP address is kept with SaveSetting / GetSetting (the same entry in the registry as the VB6 program).
' Get RS-Bus data stops after SendError, a new connection, TcpClosedByRemoteHost and TcpError.

Public Class FrmRsBusMonitor
    Private WithEvents XpressNet As XpressNetClass

    ' Referentie naar de PanelManager
    Private m_PanelManager As PanelManager


    Private Const LEFT_OFFSET As Integer = 240

    Private GetSchaltInfo As Boolean
    Private AnswerCount As Integer                   ' The answers for RS-Bus address 64 and 128



    Private Sub FrmRsBusMonitor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        XpressNet = New XpressNetClass
        ' Set the directory for logfiles
        XpressNet.Logfile.LogDirectory = "C:\Users\Public\TMC-Logfiles"
        ' The IP address of the previous connection (the first time: the text of txtIP)
        txtIP.Text = GetSetting("RS-Bus Monitor", "Settings", "IpAddress", txtIP.Text)
        picPanel.BorderStyle = BorderStyle.None
        picPanel.SetBounds(LEFT_OFFSET, 0, Me.ClientSize.Width - LEFT_OFFSET, Me.ClientSize.Height)
        ' Initialiseer de PanelManager
        m_PanelManager = New PanelManager
        m_PanelManager.Init(picPanel)
        GetSchaltInfo = False
    End Sub

    Private Sub FrmRsBusMonitor_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If m_PanelManager Is Nothing Then Exit Sub             ' The form is not yet loaded
        If Me.WindowState <> FormWindowState.Minimized Then
            picPanel.SetBounds(LEFT_OFFSET, 0, Me.ClientSize.Width - LEFT_OFFSET, Me.ClientSize.Height)
            m_PanelManager.Resize()
        End If
    End Sub


    '===============================================================================================================
    ' XPressNet Commando's
    Private Sub cmdTcpConnect_Click(sender As Object, e As EventArgs) Handles cmdTcpConnect.Click
        LabelStatus.Text = "TCP connecting ....."
        ' The IP address is kept for the next start of the program
        SaveSetting("RS-Bus Monitor", "Settings", "IpAddress", txtIP.Text)
        ' A new connection stops Get RS-Bus data: the answers on its requests will not come
        GetSchaltInfo = False
        Call XpressNet.CmdConnectTcp(txtIP.Text, "5550")
    End Sub

    Private Sub CmdAllesAn_Click(sender As Object, e As EventArgs) Handles CmdAllesAn.Click
        Call XpressNet.CmdAllesAn()
    End Sub

    Private Sub CmdAllesAus_Click(sender As Object, e As EventArgs) Handles CmdAllesAus.Click
        Call XpressNet.CmdAllesAus()
    End Sub

    '===========================================================================================
    ' ALGEMENE EVENTS
    Private Sub XpressNet_InterfaceConnected() Handles XpressNet.InterfaceConnected
        ledInterfaceStatus.FillColor = Color.Lime
        LabelStatus.Text = "Connection established"
    End Sub

    Private Sub XpressNet_TcpClosedByRemoteHost() Handles XpressNet.TcpClosedByRemoteHost
        ledInterfaceStatus.FillColor = Color.Red
        GetSchaltInfo = False                            ' Get RS-Bus data stops
        LabelStatus.Text = "TCP connection closed by remote host"
    End Sub

    Private Sub XpressNet_TcpError(ByRef ErrorDescription As String) Handles XpressNet.TcpError
        ledInterfaceStatus.FillColor = Color.Red
        GetSchaltInfo = False                            ' Get RS-Bus data stops
        LabelStatus.Text = ErrorDescription
    End Sub

    Private Sub XpressNet_AllesAn() Handles XpressNet.AllesAn
        ledNotAus.FillColor = Color.Lime
        ledNotHalt.FillColor = Color.Lime
    End Sub

    Private Sub XpressNet_NotAus() Handles XpressNet.NotAus
        ledNotAus.FillColor = Color.Red
    End Sub

    Private Sub XpressNet_NotHalt() Handles XpressNet.NotHalt
        ledNotHalt.FillColor = Color.Red
    End Sub

    ' Version 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum
    ' Get RS-Bus data is stopped (for example: no answer, the message queues of the library have been emptied)
    Private Sub XpressNet_SendError(ByVal Reason As Byte, ByVal Description As String) Handles XpressNet.SendError
        GetSchaltInfo = False
        LabelStatus.Text = CType(Reason, SendErrorReasonEnum).ToString() & ": " & Description
    End Sub

    ' Version 2.0: the command station does not address the interface: sending is paused
    Private Sub XpressNet_InterfaceNichtAdressiert() Handles XpressNet.InterfaceNichtAdressiert
        LabelStatus.Text = "Interface nicht adressiert"
    End Sub

    ' Version 2.0: the command station addresses the interface again
    Private Sub XpressNet_InterfaceWiederAdressiert() Handles XpressNet.InterfaceWiederAdressiert
        LabelStatus.Text = "Interface wieder adressiert"
    End Sub


    '===========================================================================================
    ' FEEDBACK EVENTS

    '' We hebben in de DLL 2 events voor RS-Bus input: Bezetmeld en Wissel feedbacks
    '' Voor beide roepen we SetFeedbackNibble aan
    ''
    '' Bezetmeld: Private Sub XpressNet_Feedback(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte)
    '' - address = 1..128
    '' - nibble = 0 / 1
    '' - data = 0..15
    Private Sub XpressNet_Feedback(ByVal address As Integer, ByVal Nibble As Byte, ByVal data As Byte) Handles XpressNet.Feedback
        m_PanelManager.SetFeedbackNibble(address, Nibble, data, BlockStateType.BEZETMELDING)
        ' Get RS-Bus data: is this the answer on the last request of a part?
        CheckSchaltInfo(address)
    End Sub
    ''
    '' Wissel: Private Sub XpressNet_SwitchFeedback(switchAddress, ttBits, iBit, data)
    '' - switchAddress =  ((address * 4) + (Nibble * 2) + 1)
    '' - ttBits =
    '' - iBit =
    '' - data = 0..15
    Private Sub XpressNet_SwitchFeedback(ByVal switchAddress As Integer, ByVal ttBits As Byte, ByVal iBit As Byte, ByVal data As Byte) Handles XpressNet.SwitchFeedback
        Dim RsAddress As Integer
        Dim Nibble As Byte
        RsAddress = Int((switchAddress - 1) / 4) + 1
        Nibble = ((switchAddress - 1) \ 2) Mod 2
        If ttBits = 0 Then     ' Feedback comes from Command Station
            If Not GetSchaltInfo Then
                m_PanelManager.SetFeedbackNibble(RsAddress, Nibble, data, BlockStateType.CENTRALE)
            End If
        Else                   ' Feedback comes from Decoder
            m_PanelManager.SetFeedbackNibble(RsAddress, Nibble, data, BlockStateType.WISSEL)
        End If
        ' Get RS-Bus data: is this the answer on the last request of a part?
        CheckSchaltInfo(RsAddress)
    End Sub


    '===========================================================================================
    ' VERTALING TUSSE ADDRESSEN
    Private Sub RS2Switch_Click(sender As Object, e As EventArgs) Handles RS2Switch.Click
        Dim rsAddr As Integer
        Dim firstSw As Integer
        Dim lastSw As Integer
        If IsNumeric(TextRS2Switch.Text) Then
            rsAddr = CInt(TextRS2Switch.Text)
            If rsAddr >= 1 Then
                firstSw = (rsAddr - 1) * 4 + 1
                lastSw = firstSw + 3
                TextSwitch2RS.Text = firstSw & " .. " & lastSw
            End If
        End If
        ' Ongeldige input? Doe niets
    End Sub

    Private Sub Switch2RS_Click(sender As Object, e As EventArgs) Handles Switch2RS.Click
        Dim swAddr As Integer
        Dim rsAddr As Integer
        If IsNumeric(TextSwitch2RS.Text) Then
            swAddr = CInt(TextSwitch2RS.Text)
            If swAddr >= 1 Then
                rsAddr = (swAddr - 1) \ 4 + 1
                TextRS2Switch.Text = CStr(rsAddr)
            End If
        End If
        ' Ongeldige input? Doe niets
    End Sub


    '===========================================================================================
    ' Alles ophalen
    ' Get RS-Bus data in two parts of 128 requests (64 RS-Bus addresses, two nibbles each). The FIFO of version 2.0 of
    ' the library holds 128 messages: more requests at once would give the event SendError (FifoVoll).
    ' The second part is requested as soon as both requests for RS-Bus address 64 have been answered; Get RS-Bus data
    ' ends when both requests for RS-Bus address 128 have been answered (see CheckSchaltInfo). This does not depend on
    ' the speed of the interface. The event SendError stops Get RS-Bus data.
    Private Sub SchaltInfo_Click(sender As Object, e As EventArgs) Handles SchaltInfo.Click
        If GetSchaltInfo Then Exit Sub                   ' Get RS-Bus data is already active
        GetSchaltInfo = True
        AnswerCount = 0
        LabelStatus.Text = "Requesting RS-Bus data ....."
        RequestFeedbacks(1, 64)
    End Sub

    ' Requests both nibbles of the RS-Bus addresses firstAddress .. lastAddress
    Private Sub RequestFeedbacks(ByVal firstAddress As Integer, ByVal lastAddress As Integer)
        Dim address As Integer
        For address = firstAddress To lastAddress
            Call XpressNet.CmdFeedbackRequest(address, 0)  ' Nibble = 0
            Call XpressNet.CmdFeedbackRequest(address, 1)  ' Nibble = 1
        Next
    End Sub

    ' Called for every feedback message, after the matrix has been updated. The answers are counted per RS-Bus
    ' address, and not per nibble: for an RS-Bus address without feedback information, the command station (LZ100)
    ' answers both requests with nibble 0 (42 Address 00).
    ' Note: an unrequested feedback message (broadcast) for RS-Bus address 64 or 128 is counted as well
    Private Sub CheckSchaltInfo(ByVal RsAddress As Integer)
        If Not GetSchaltInfo Then Exit Sub
        If (RsAddress <> 64) And (RsAddress <> 128) Then Exit Sub
        AnswerCount = AnswerCount + 1
        If AnswerCount = 2 Then RequestFeedbacks(65, 128) ' The first part has been answered
        If AnswerCount = 4 Then                          ' The second part has been answered
            GetSchaltInfo = False
            LabelStatus.Text = "RS-Bus data received"
        End If
    End Sub


    '===========================================================================================
    ' Release the XpressNet object (which in turn closes the TCP connection)
    Private Sub FrmRsBusMonitor_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        XpressNet.Dispose()
    End Sub
End Class
