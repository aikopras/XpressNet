Option Explicit On
Imports XpressNet
' Aangepast 14/08/2025 - 21/08/25: Handles veranderd.

Public Class FrmBasicTest
    'XpressNet object instantieren en Handlers declareren
    Private XpressNet As New XpressNetClass()

    Private Sub FrmBasicTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'XpressNet object instantieren en Handlers declareren
        XpressNet = New XpressNetClass()
        AddHandler XpressNet.TcpError, AddressOf TcpError
        AddHandler XpressNet.InterfaceConnected, AddressOf TcpConnected
        AddHandler XpressNet.TcpClosedByRemoteHost, AddressOf TcpClosed
        AddHandler XpressNet.Feedback, AddressOf Feedback
        AddHandler XpressNet.SwitchFeedback, AddressOf SwitchFeedback
        AddHandler XpressNet.AllesAn, AddressOf AllesAn
        AddHandler XpressNet.NotAus, AddressOf NotAus
        AddHandler XpressNet.NotHalt, AddressOf NotHalt
        AddHandler XpressNet.ZentraleVersion, AddressOf ZentraleVersion
        AddHandler XpressNet.XpressNetAddress, AddressOf XpressNetAddress
        AddHandler XpressNet.XpressNetVersion, AddressOf XpressNetVersion
        AddHandler XpressNet.InterfaceVersion, AddressOf InterfaceVersion
        AddHandler XpressNet.InterfaceFreieVerbindungen, AddressOf InterfaceFreieVerbindungen


        'Vullen van de ComboBox Baudrate voor de USB Connnectie
        With Me.CbBaudrate.Items
            .Add(9600)
            .Add(57600)
            .Add(115200)
        End With
        CbBaudrate.SelectedItem = (57600)

    End Sub
    Private Sub BtnConnectTcpIp_Click(sender As Object, e As EventArgs) Handles BtnConnectTcpIp.Click
        Call XpressNet.CmdConnectTcp(TxtIpAddress.Text, TxtTcpPort.Text)
    End Sub
    Private Sub BtnConnectUsb_Click(sender As Object, e As EventArgs) Handles BtnConnectUsb.Click
        Dim Settings As String
        Settings = CbBaudrate.Text & ",N,8,1"
        Call XpressNet.CmdConnectUsb(TxtUsbPort.Text, Settings)
    End Sub

    Private Sub BtnInterfaceInfo_Click(sender As Object, e As EventArgs) Handles BtnInterfaceInfo.Click
        Call XpressNet.CmdSoftwareversionZentraleAnfordern()
        Call XpressNet.CmdInterfaceVersion()
        Call XpressNet.CmdInterfaceStatus()
    End Sub
    Private Sub BtnLanInfo_Click(sender As Object, e As EventArgs) Handles BtnLanInfo.Click
        Call XpressNet.CmdXpressNetAddress()
        Call XpressNet.CmdXpressNetVersion()
        Call XpressNet.CmdInterfaceFreieVerbindungen()
    End Sub
    Private Sub BtnRecht_Click(sender As Object, e As EventArgs) Handles BtnRecht.Click
        If Not IsNumeric(TxtSwitch.Text) Then Exit Sub
        Call XpressNet.CmdAccessoryAuto(TxtSwitch.Text, 1)
    End Sub
    Private Sub BtnAf_Click(sender As Object, e As EventArgs) Handles BtnAf.Click
        If Not IsNumeric(TxtSwitch.Text) Then Exit Sub
        Call XpressNet.CmdAccessoryAuto(TxtSwitch.Text, 0)
    End Sub
    Private Sub BtnAllesAan_Click(sender As Object, e As EventArgs) Handles BtnAllesAan.Click
        Call XpressNet.CmdAllesAn()
    End Sub
    Private Sub BtnAllesUit_Click(sender As Object, e As EventArgs) Handles BtnAllesUit.Click
        Call XpressNet.CmdAllesAus()
    End Sub

    '===========================================================================================
    ' EVENTS
    Public Sub TcpConnected()
        ShConnected.BackColor = Color.Green
    End Sub


    Public Sub TcpClosed()
        ShConnected.BackColor = Color.Red
    End Sub

    Public Sub TcpError()
        ShConnected.BackColor = Color.Red
    End Sub

    'Private Sub myXpressNet_SwitchFeedback(address As Integer, ttBits As Byte, iBit As Byte, data As Byte) Handles myXpressNet.SwitchFeedback
    '    MessageBox.Show($"Address: {address}, ttBits: {ttBits}, iBit: {iBit}, data: {data}")


    Public Sub Feedback(address As Integer, nibble As Byte, data As Byte)
        TxtRSBusAddress.Text = address & vbCrLf & TxtRSBusAddress.Text
        TxtRSBusData.Text = PrintBits(data) & vbCrLf & TxtRSBusData.Text
        TxtRSBusNibble.Text = nibble & vbCrLf & TxtRSBusNibble.Text
    End Sub

    Public Sub SwitchFeedback(address As Integer, ttBits As Byte, iBit As Byte, data As Byte)
        Dim ttText As String
        If ttBits = 0 Then
            ttText = "C"  ' Feedback comes from Command Station
        Else
            ttText = "D"  ' Feedback comes from Decoder
        End If
        TxtSwitchAddress.Text = address & "/" & (address + 1) & vbCrLf & TxtSwitchAddress.Text
        TxtSwitchData.Text = PrintBits(data) & vbCrLf & TxtSwitchData.Text
        TxtSwitchTT.Text = ttText & vbCrLf & TxtSwitchTT.Text
        TxtSwitchI.Text = iBit & vbCrLf & TxtSwitchI.Text
        PrintBits(data)
    End Sub

    Public Sub AllesAn()
        ShNotAus.BackColor = Color.Green
        ShNotHalt.BackColor = Color.Green
    End Sub

    Public Sub NotAus()
        ShNotAus.BackColor = Color.Red
    End Sub

    Public Sub NotHalt()
        ShNotHalt.BackColor = Color.Red
    End Sub

    Public Sub ZentraleVersion(Version As String)
        TxtMasterVersion.Text = Version
    End Sub

    Public Sub XpressNetAddress(address As Byte)
        TxtXpressNetAddress.Text = address
    End Sub

    Public Sub XpressNetVersion(Version As String)
        TxtXpressNetVersion.Text = Version
    End Sub

    Public Sub InterfaceVersion(Version As String, code As String)
        TxtInterfaceVersion.Text = Version
        TxtInterfaceCode.Text = code
    End Sub

    Public Sub InterfaceFreieVerbindungen(number As Byte)
        TxtFreeTcpSockets.Text = number
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    ' Some bit manipulation functions
    Public Function BitValue(data As Byte, index As Byte) As Boolean
        index = 2 ^ index
        BitValue = (data And index) / index
    End Function

    Public Function BitSet(data As Byte, index As Byte) As Byte
        index = 2 ^ index
        BitSet = data Or index
    End Function

    Public Function BitClear(data As Byte, index As Byte) As Byte
        index = 2 ^ index              ' Set the bit at position index (other bist are 0)
        index = 255 - index            ' Flip all bits
        BitClear = data And index      ' mask the data
    End Function

    Public Function BitUpdate(data As Byte, index As Byte, bitIsSet As Byte) As Byte
        If bitIsSet Then
            BitUpdate = BitSet(data, index)
        Else
            BitUpdate = BitClear(data, index)
        End If
    End Function

    '---------------------------------------------------------------------------------------------------------------------
    ' Transforms numbers form 0..15 (4 bits) in a bitstring
    Public Function PrintBits(data As Byte) As String
        Dim outString As String
        outString = ""
        Dim i As Integer
        For i = 0 To 3
            outString = (data Mod 2) & outString
            data \= 2
        Next
        PrintBits = outString
    End Function

End Class
