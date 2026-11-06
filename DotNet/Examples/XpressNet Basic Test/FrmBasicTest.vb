Option Explicit On
Imports XpressNet
' 2026-11: VB.NET version of the VB6 example XpressNet Basic Test (B_UserInterfaceFormConnect.frm), for XpressNet.NET 2.0.
' The layout is converted from the VB6 form; the code follows the VB6 code. Differences with VB6:
' - The events are connected with Handles (VB6: the name XpressNet_EventName).
' - The VB6 Shape controls are Led controls (Led.vb); the control array LEDF0F68 is filled in InitControlArrays.
' - Default properties are written in full (for example LabelStatus.Text instead of LabelStatus).
' - Form_Unload / Set XpressNet = Nothing is FormClosed / XpressNet.Dispose().
' - The handlers CmdErweiterteVersion_Click, CmdStartMan_Click and cmdStartAuto_Click are left out: these buttons
'   are not on the form (in VB6 the code was never called).
' New in version 2.0: the logfile directory TMC-Logfiles, the option Handshake for USB, and the events SendError,
' InterfaceNichtAdressiert and InterfaceWiederAdressiert (shown in LabelStatus).

Public Class FrmBasicTest
    Private WithEvents XpressNet As XpressNetClass

    ' ----------------------------------------------------------------------------------------------------------------------
    ' Function Groups:
    ' For each Lok Address, the values must be fethed again
    Private FB(10) As Byte                 ' FB: Function Befehl (on / off)
    Private FS(10) As Byte                 ' Function Status (tastend oder nicht tastend)

    ' The control array of the VB6 form (see InitControlArrays)
    Private LEDF0F68(68) As Led


    Private Sub FrmBasicTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call InitControlArrays()
        XpressNet = New XpressNetClass
        ' Set the directory for logfiles
        ' 2026-11 modified: TMC-Logfiles, the default directory of version 2.0
        'XpressNet.Logfile.LogDirectory = "C:\Users\Public\XPN-Logfiles"
        XpressNet.Logfile.LogDirectory = "C:\Users\Public\TMC-Logfiles"
        ' Possible Baudrates for the USB interface
        ComboBaudUsb.Items.Add("57600")
        ComboBaudUsb.Items.Add("115200")
        ' Possible lok decoder steps
        ComboLokSteps.Items.Add("14")
        ComboLokSteps.Items.Add("27")
        ComboLokSteps.Items.Add("28")
        ComboLokSteps.Items.Add("128")
    End Sub

    Private Sub FrmBasicTest_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        XpressNet.Dispose()
    End Sub


    '===============================================================================================================
    Private Sub cmdTcpConnect_Click(sender As Object, e As EventArgs) Handles cmdTcpConnect.Click
        LabelStatus.Text = "TCP connecting ....."
        Call XpressNet.CmdConnectTcp(txtIP.Text, txtTcpPort.Text)
    End Sub

    Private Sub CmdUsbConnect_Click(sender As Object, e As EventArgs) Handles CmdUsbConnect.Click
        LabelStatus.Text = "USB connecting ....."
        Dim Settings As String
        Settings = ComboBaudUsb.Text & ",N,8,1"
        ' 2026-11 modified: version 2.0, with the optional parameter Handshake
        'Call XpressNet.CmdConnectUsb(TxtUsbPort.Text, Settings)
        Call XpressNet.CmdConnectUsb(TxtUsbPort.Text, Settings, CheckHandshake.Checked)
    End Sub


    Private Sub LabelMasterVersion_Click(sender As Object, e As EventArgs) Handles LabelMasterVersion.Click
        Call XpressNet.CmdSoftwareversionZentraleAnfordern()
    End Sub

    Private Sub CmdInterfaceInfo_Click(sender As Object, e As EventArgs) Handles CmdInterfaceInfo.Click
        Call XpressNet.CmdInterfaceVersion()
    End Sub

    Private Sub CmdLanInfo_Click(sender As Object, e As EventArgs) Handles CmdLanInfo.Click
        Call XpressNet.CmdXpressNetAddress()
        Call XpressNet.CmdXpressNetVersion()
        Call XpressNet.CmdInterfaceFreieVerbindungen()
    End Sub

    '===============================================================================================================
    Private Sub cmdRecht_Click(sender As Object, e As EventArgs) Handles cmdRecht.Click
        If Not IsNumeric(txtWissel.Text) Then Exit Sub
        Call XpressNet.CmdAccessoryAuto(txtWissel.Text, 1)
    End Sub

    Private Sub cmdAf_Click(sender As Object, e As EventArgs) Handles cmdAf.Click
        If Not IsNumeric(txtWissel.Text) Then Exit Sub
        Call XpressNet.CmdAccessoryAuto(txtWissel.Text, 0)
    End Sub

    Private Sub cmdSwitchFeedback_Click(sender As Object, e As EventArgs) Handles cmdSwitchFeedback.Click
        If IsNumeric(txtWissel.Text) Then
            Call XpressNet.CmdSwitchInfoRequest(txtWissel.Text)
        End If
    End Sub

    Private Sub CmdRsFeedback_Click(sender As Object, e As EventArgs) Handles CmdRsFeedback.Click
        If IsNumeric(TxtRsBusAddress.Text) Then
            Call XpressNet.CmdFeedbackRequest(TxtRsBusAddress.Text, 0)  ' Nibble 0
            Call XpressNet.CmdFeedbackRequest(TxtRsBusAddress.Text, 1)  ' Nibble 1
        End If
    End Sub

    Private Sub CmdAllesAn_Click(sender As Object, e As EventArgs) Handles CmdAllesAn.Click
        Call XpressNet.CmdAllesAn()
    End Sub

    Private Sub CmdAllesAus_Click(sender As Object, e As EventArgs) Handles CmdAllesAus.Click
        Call XpressNet.CmdAllesAus()
    End Sub

    Private Sub CmdAlleLoksAnhalten_Click(sender As Object, e As EventArgs) Handles CmdAlleLoksAnhalten.Click
        Call XpressNet.CmdAlleLoksAnhalten()
    End Sub

    Private Sub CmdEineLokAnhalten_Click(sender As Object, e As EventArgs) Handles CmdEineLokAnhalten.Click
        If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
        Call XpressNet.CmdEineLokAnhalten(txtLokAddress.Text)
    End Sub

    Private Sub CmdLokEnter_Click(sender As Object, e As EventArgs) Handles CmdLokEnter.Click
        ' A new lok address is entered. Get the lok info
        If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
        Call ClearLedsForFunctions()
        Call XpressNet.CmdLokInfoAnfordern(txtLokAddress.Text)
    End Sub

    Private Sub CmdLokAusStack_Click(sender As Object, e As EventArgs) Handles CmdLokAusStack.Click
        If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
        Call XpressNet.CmdLokAusStack(txtLokAddress.Text)
    End Sub

    Private Sub CmdLokForward_Click(sender As Object, e As EventArgs) Handles CmdLokForward.Click
        If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
        If Not IsNumeric(TxtLokSpeed.Text) Then Exit Sub
        Call XpressNet.CmdFahrbefehl(txtLokAddress.Text, TxtLokSpeed.Text, ComboLokSteps.Text, True)
    End Sub

    Private Sub CmdLokReverse_Click(sender As Object, e As EventArgs) Handles CmdLokReverse.Click
        If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
        If Not IsNumeric(TxtLokSpeed.Text) Then Exit Sub
        Call XpressNet.CmdFahrbefehl(txtLokAddress.Text, TxtLokSpeed.Text, ComboLokSteps.Text, False)
    End Sub

    Private Sub CmdFunctionOn_Click(sender As Object, e As EventArgs) Handles CmdFunctionOn.Click
        Call handleFunctionBefehl(True)
    End Sub

    Private Sub CmdFunctionOff_Click(sender As Object, e As EventArgs) Handles CmdFunctionOff.Click
        Call handleFunctionBefehl(False)
    End Sub

    Private Sub CmdGetF0_F12_Click(sender As Object, e As EventArgs) Handles CmdGetF0_F12.Click
        Call XpressNet.CmdLokInfoAnfordern(txtLokAddress.Text)
    End Sub

    Private Sub CmdGetF13_28_Click(sender As Object, e As EventArgs) Handles CmdGetF13_28.Click
        Call XpressNet.CmdFunktionsZustandF13F28Anfordern(txtLokAddress.Text)
    End Sub

    Private Sub CmdGetF29_68_Click(sender As Object, e As EventArgs) Handles CmdGetF29_68.Click
        Call XpressNet.CmdFunktionsZustandF29F68Anfordern(txtLokAddress.Text)
    End Sub





    '===========================================================================================
    ' EVENTS
    Private Sub XpressNet_InterfaceConnected() Handles XpressNet.InterfaceConnected
        ledInterfaceStatus.FillColor = Color.Lime
        LabelStatus.Text = "Connection established"
    End Sub

    Private Sub XpressNet_TcpClosedByRemoteHost() Handles XpressNet.TcpClosedByRemoteHost
        ledInterfaceStatus.FillColor = Color.Red
        LabelStatus.Text = "TCP connection closed by remote host"
    End Sub

    Private Sub XpressNet_TcpError(ByRef ErrorDescription As String) Handles XpressNet.TcpError
        ledInterfaceStatus.FillColor = Color.Red
        LabelStatus.Text = ErrorDescription
    End Sub

    Private Sub XpressNet_Feedback(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte) Handles XpressNet.Feedback
        TxBoxtRsAddress.Text = address & vbCrLf & TxBoxtRsAddress.Text
        TxtBoxRsData.Text = printBits(data) & vbCrLf & TxtBoxRsData.Text
        TxtBoxRsNibble.Text = nibble & vbCrLf & TxtBoxRsNibble.Text
    End Sub

    Private Sub XpressNet_SwitchFeedback(ByVal address As Integer, ByVal ttBits As Byte, ByVal iBit As Byte, ByVal data As Byte) Handles XpressNet.SwitchFeedback
        Dim ttText As String
        If ttBits = 0 Then
            ttText = "C"  ' Feedback comes from Command Station
        Else
            ttText = "D"  ' Feedback comes from Decoder
        End If
        TxtBoxSwitchAddress.Text = address & "/" & (address + 1) & vbCrLf & TxtBoxSwitchAddress.Text
        TxtBoxSwitchData.Text = printBits(data) & vbCrLf & TxtBoxSwitchData.Text
        TxtBoxSwitchTtBits.Text = ttText & vbCrLf & TxtBoxSwitchTtBits.Text
        TxtBoxSwitchIBit.Text = iBit & vbCrLf & TxtBoxSwitchIBit.Text
    End Sub

    Private Sub XpressNet_AllesAn() Handles XpressNet.AllesAn
        ledNotAus.FillColor = Color.Lime
        ledNotHalt.FillColor = Color.Lime
        ledProgrammiermode.FillColor = Color.Lime
    End Sub

    Private Sub XpressNet_NotAus() Handles XpressNet.NotAus
        ledNotAus.FillColor = Color.Red
    End Sub

    Private Sub XpressNet_NotHalt() Handles XpressNet.NotHalt
        Debug.Print("NotHalt")
        ledNotHalt.FillColor = Color.Red
    End Sub

    Private Sub XpressNet_ProgrammierMode() Handles XpressNet.ProgrammierMode
        ledProgrammiermode.FillColor = Color.Red
    End Sub

    Private Sub XpressNet_ZentraleVersion(ByVal version As String) Handles XpressNet.ZentraleVersion
        txtMasterVersion.Text = version
    End Sub

    Private Sub XpressNet_Fehlermeldung(ByVal Value As Byte) Handles XpressNet.Fehlermeldung
        LabelStatus.Text = "Fehlermeldung: " & Value & " (V3.6: 3.1.14)"
    End Sub

    Private Sub XpressNet_XpressNetAddress(ByVal address As Byte) Handles XpressNet.XpressNetAddress
        txtXAddress.Text = address
    End Sub

    Private Sub XpressNet_XpressNetVersion(ByVal version As String) Handles XpressNet.XpressNetVersion
        txtXVersion.Text = version
    End Sub

    Private Sub XpressNet_InterfaceVersion(ByVal version As String, ByVal code As String) Handles XpressNet.InterfaceVersion
        txtInterfaceVersion.Text = version
        txtInterfaceCode.Text = code
    End Sub

    Private Sub XpressNet_InterfaceFreieVerbindungen(ByVal number As Byte) Handles XpressNet.InterfaceFreieVerbindungen
        txtTcpFree.Text = number
    End Sub

    Private Sub XpressNet_NormaleLokinfo(ByVal Besetzt As Boolean, ByVal Speedsteps As Integer, ByVal Speed As Integer, ByVal Forward As Boolean, ByVal F0_F4 As Byte, ByVal F5_F8 As Byte, ByVal F9_F12 As Byte) Handles XpressNet.NormaleLokinfo
        ' Debug.Print "Lok ist besetzt: " & Besetzt
        ' Debug.Print "Vorwarts: " & Forward
        ' Set the combobox and speed, based on the received input
        ComboLokSteps.Text = Speedsteps
        TxtLokSpeed.Text = Speed
        ' Store the values of F0..F12, to allow changing a single Function later
        FB(1) = F0_F4                  ' Gruppe 1
        FB(2) = F5_F8                  ' Gruppe 2
        FB(3) = F9_F12                 ' Gruppe 3
        ' Show the LEDs for F0 .. F12
        Call ShowFbGroup(F0_F4, 1)     ' Led FB bits for Gruppe 1
        Call ShowFbGroup(F5_F8, 2)     ' Led FB bits for Gruppe 1
        Call ShowFbGroup(F9_F12, 3)    ' Led FB bits for Gruppe 1
    End Sub

    Private Sub XpressNet_FunctionsZustandF13F20(ByVal F13_F20 As Byte, ByVal F21_F28 As Byte) Handles XpressNet.FunctionsZustandF13F20
        FB(4) = F13_F20                ' Store Gruppe 4
        FB(5) = F21_F28                ' Store Gruppe 5
        Call ShowFbGroup(F13_F20, 4)   ' Led FB bits for Gruppe 4
        Call ShowFbGroup(F21_F28, 5)   ' Led FB bits for Gruppe 5
    End Sub
    Private Sub XpressNet_FunctionsZustandF29F68(ByVal F29_F36 As Byte, ByVal F37_F44 As Byte, ByVal F45_F52 As Byte, ByVal F53_F60 As Byte, ByVal F61_F68 As Byte) Handles XpressNet.FunctionsZustandF29F68
        FB(6) = F29_F36                ' Store Gruppe 6
        FB(7) = F37_F44                ' Store Gruppe 7
        FB(8) = F45_F52                ' Store Gruppe 8
        FB(9) = F53_F60                ' Store Gruppe 9
        FB(10) = F61_F68               ' Store Gruppe 10
        Call ShowFbGroup(F29_F36, 6)   ' Led FB bits for Gruppe 6
        Call ShowFbGroup(F37_F44, 7)   ' Led FB bits for Gruppe 7
        Call ShowFbGroup(F45_F52, 8)   ' Led FB bits for Gruppe 8
        Call ShowFbGroup(F53_F60, 9)   ' Led FB bits for Gruppe 9
        Call ShowFbGroup(F61_F68, 10)  ' Led FB bits for Gruppe 10
    End Sub

    ' 2026-11 modified: new (version 2.0). A message that the library could not send. Reason: see SendErrorReasonEnum
    Private Sub XpressNet_SendError(ByVal Reason As Byte, ByVal Description As String) Handles XpressNet.SendError
        LabelStatus.Text = CType(Reason, SendErrorReasonEnum).ToString() & ": " & Description
    End Sub

    ' 2026-11 modified: new (version 2.0). The command station does not address the interface: sending is paused
    Private Sub XpressNet_InterfaceNichtAdressiert() Handles XpressNet.InterfaceNichtAdressiert
        LabelStatus.Text = "Interface nicht adressiert"
    End Sub

    ' 2026-11 modified: new (version 2.0). The command station addresses the interface again
    Private Sub XpressNet_InterfaceWiederAdressiert() Handles XpressNet.InterfaceWiederAdressiert
        LabelStatus.Text = "Interface wieder adressiert"
    End Sub




    '=====================================================================================================================
    ' Support routines below
    '---------------------------------------------------------------------------------------------------------------------
    ' The VB6 control array LEDF0F68: the Led controls LEDF0F68_0 .. LEDF0F68_68 of the form
    Private Sub InitControlArrays()
        LEDF0F68(0) = LEDF0F68_0
        LEDF0F68(1) = LEDF0F68_1
        LEDF0F68(2) = LEDF0F68_2
        LEDF0F68(3) = LEDF0F68_3
        LEDF0F68(4) = LEDF0F68_4
        LEDF0F68(5) = LEDF0F68_5
        LEDF0F68(6) = LEDF0F68_6
        LEDF0F68(7) = LEDF0F68_7
        LEDF0F68(8) = LEDF0F68_8
        LEDF0F68(9) = LEDF0F68_9
        LEDF0F68(10) = LEDF0F68_10
        LEDF0F68(11) = LEDF0F68_11
        LEDF0F68(12) = LEDF0F68_12
        LEDF0F68(13) = LEDF0F68_13
        LEDF0F68(14) = LEDF0F68_14
        LEDF0F68(15) = LEDF0F68_15
        LEDF0F68(16) = LEDF0F68_16
        LEDF0F68(17) = LEDF0F68_17
        LEDF0F68(18) = LEDF0F68_18
        LEDF0F68(19) = LEDF0F68_19
        LEDF0F68(20) = LEDF0F68_20
        LEDF0F68(21) = LEDF0F68_21
        LEDF0F68(22) = LEDF0F68_22
        LEDF0F68(23) = LEDF0F68_23
        LEDF0F68(24) = LEDF0F68_24
        LEDF0F68(25) = LEDF0F68_25
        LEDF0F68(26) = LEDF0F68_26
        LEDF0F68(27) = LEDF0F68_27
        LEDF0F68(28) = LEDF0F68_28
        LEDF0F68(29) = LEDF0F68_29
        LEDF0F68(30) = LEDF0F68_30
        LEDF0F68(31) = LEDF0F68_31
        LEDF0F68(32) = LEDF0F68_32
        LEDF0F68(33) = LEDF0F68_33
        LEDF0F68(34) = LEDF0F68_34
        LEDF0F68(35) = LEDF0F68_35
        LEDF0F68(36) = LEDF0F68_36
        LEDF0F68(37) = LEDF0F68_37
        LEDF0F68(38) = LEDF0F68_38
        LEDF0F68(39) = LEDF0F68_39
        LEDF0F68(40) = LEDF0F68_40
        LEDF0F68(41) = LEDF0F68_41
        LEDF0F68(42) = LEDF0F68_42
        LEDF0F68(43) = LEDF0F68_43
        LEDF0F68(44) = LEDF0F68_44
        LEDF0F68(45) = LEDF0F68_45
        LEDF0F68(46) = LEDF0F68_46
        LEDF0F68(47) = LEDF0F68_47
        LEDF0F68(48) = LEDF0F68_48
        LEDF0F68(49) = LEDF0F68_49
        LEDF0F68(50) = LEDF0F68_50
        LEDF0F68(51) = LEDF0F68_51
        LEDF0F68(52) = LEDF0F68_52
        LEDF0F68(53) = LEDF0F68_53
        LEDF0F68(54) = LEDF0F68_54
        LEDF0F68(55) = LEDF0F68_55
        LEDF0F68(56) = LEDF0F68_56
        LEDF0F68(57) = LEDF0F68_57
        LEDF0F68(58) = LEDF0F68_58
        LEDF0F68(59) = LEDF0F68_59
        LEDF0F68(60) = LEDF0F68_60
        LEDF0F68(61) = LEDF0F68_61
        LEDF0F68(62) = LEDF0F68_62
        LEDF0F68(63) = LEDF0F68_63
        LEDF0F68(64) = LEDF0F68_64
        LEDF0F68(65) = LEDF0F68_65
        LEDF0F68(66) = LEDF0F68_66
        LEDF0F68(67) = LEDF0F68_67
        LEDF0F68(68) = LEDF0F68_68
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    ' Lok Functions
    Private Sub handleFunctionBefehl(ByVal turnOn As Boolean)
        If Not IsNumeric(txtLokAddress.Text) Then Exit Sub
        If Not IsNumeric(TxtLokFNumber.Text) Then Exit Sub
        Dim FbNumber As Byte
        FbNumber = TxtLokFNumber.Text
        If FbNumber < 0 Then Exit Sub
        If FbNumber > 68 Then Exit Sub
        Dim group As Byte
        Dim bitposition As Byte
        ' F0 needs to be consideren as speciaal case
        If FbNumber = 0 Then group = 1 : bitposition = 4
        ' The other functions are coded in a consistent way
        Select Case FbNumber
            Case 1 To 4 : group = 1 : bitposition = FbNumber - 1
            Case 5 To 8 : group = 2 : bitposition = FbNumber - 5
            Case 9 To 12 : group = 3 : bitposition = FbNumber - 9
            Case 13 To 20 : group = 4 : bitposition = FbNumber - 13
            Case 21 To 28 : group = 5 : bitposition = FbNumber - 21
            Case 29 To 36 : group = 6 : bitposition = FbNumber - 29
            Case 37 To 44 : group = 7 : bitposition = FbNumber - 37
            Case 45 To 52 : group = 8 : bitposition = FbNumber - 45
            Case 53 To 60 : group = 9 : bitposition = FbNumber - 53
            Case 61 To 68 : group = 10 : bitposition = FbNumber - 61
        End Select
        FB(group) = bitUpdate(FB(group), bitposition, turnOn)
        Call XpressNet.CmdLokFunktionBefehl(txtLokAddress.Text, group, FB(group))
    End Sub

    Private Sub ClearLedsForFunctions()
        Dim i As Integer
        For i = 0 To 68
            LEDF0F68(i).FillColor = Color.White
        Next
    End Sub

    Private Function FbColor(ByVal Value As Boolean) As Color
        If Value Then
            FbColor = Color.Red
        Else
            FbColor = Color.Lime
        End If
    End Function

    '---------------------------------------------------------------------------------------------------------------------
    ' Functions to show the Feedbacks on this form
    Private Sub ShowFbGroup(ByVal data As Byte, ByVal group As Byte)
        Dim bitsInGroup As Byte               ' Groups 1..3 contain 4 FB bits, the others contain 8 FB bits
        Dim offset As Byte                    ' Offset in the VB Array for a certain group
        If group < 4 Then
            bitsInGroup = 4
            offset = group * bitsInGroup - 4
        Else
            bitsInGroup = 8
            offset = group * bitsInGroup - 20
        End If
        Dim i As Byte
        For i = 1 To bitsInGroup
            LEDF0F68(i + offset).FillColor = FbColor(bitValue(data, i - 1))
        Next
        ' FB0 in Group 1 is a special case
        If group = 1 Then LEDF0F68(0).FillColor = FbColor(bitValue(data, 4))
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    ' Some bit manipulation functions
    Private Function bitValue(ByVal data As Byte, ByVal index As Byte) As Boolean
        index = 2 ^ index
        bitValue = (data And index) / index
    End Function

    Private Function bitSet(ByVal data As Byte, ByVal index As Byte) As Byte
        index = 2 ^ index
        bitSet = data Or index
    End Function

    Private Function bitClear(ByVal data As Byte, ByVal index As Byte) As Byte
        index = 2 ^ index              ' Set the bit at position index (other bist are 0)
        index = 255 - index            ' Flip all bits
        bitClear = data And index      ' mask the data
    End Function

    Private Function bitUpdate(ByVal data As Byte, ByVal index As Byte, ByVal bitIsSet As Boolean) As Byte
        ' 2026-11 modified: bitIsSet As Boolean (was As Byte; the callers pass a Boolean)
        If bitIsSet Then
            bitUpdate = bitSet(data, index)
        Else
            bitUpdate = bitClear(data, index)
        End If
    End Function

    '---------------------------------------------------------------------------------------------------------------------
    ' Transforms numbers form 0..15 (4 bits) in a bitstring
    Private Function printBits(ByVal data As Byte) As String
        Dim outString As String
        Dim i As Integer
        outString = ""
        For i = 0 To 3
            outString = (data Mod 2) & outString
            data = data \ 2
        Next
        printBits = outString
    End Function

End Class
