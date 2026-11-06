<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmZ21Test
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.LabelStatus = New System.Windows.Forms.Label()
        Me.FrameLoks = New System.Windows.Forms.GroupBox()
        Me.CmdAlleLoksAnhalten = New System.Windows.Forms.Button()
        Me.CmdEineLokAnhalten = New System.Windows.Forms.Button()
        Me.txtLokAddress = New System.Windows.Forms.TextBox()
        Me.TxtLokSpeed = New System.Windows.Forms.TextBox()
        Me.ComboLokSteps = New System.Windows.Forms.ComboBox()
        Me.CmdLokReverse = New System.Windows.Forms.Button()
        Me.CmdLokForward = New System.Windows.Forms.Button()
        Me.TxtLokFNumber = New System.Windows.Forms.TextBox()
        Me.CmdFunctionOn = New System.Windows.Forms.Button()
        Me.CmdFunctionOff = New System.Windows.Forms.Button()
        Me.CmdGetLokInfo = New System.Windows.Forms.Button()
        Me.CmdLokEnter = New System.Windows.Forms.Button()
        Me.CmdLokAusStack = New System.Windows.Forms.Button()
        Me.LabelLokAddress = New System.Windows.Forms.Label()
        Me.LabelLokSpeed = New System.Windows.Forms.Label()
        Me.LabelLokSteps = New System.Windows.Forms.Label()
        Me.LabelFunktion = New System.Windows.Forms.Label()
        Me.LEDF0F68_0 = New Led()
        Me.LEDF0F68_1 = New Led()
        Me.LEDF0F68_2 = New Led()
        Me.LEDF0F68_3 = New Led()
        Me.LEDF0F68_4 = New Led()
        Me.LabelFunctions_0 = New System.Windows.Forms.Label()
        Me.LEDF0F68_5 = New Led()
        Me.LEDF0F68_6 = New Led()
        Me.LEDF0F68_7 = New Led()
        Me.LEDF0F68_8 = New Led()
        Me.LEDF0F68_9 = New Led()
        Me.LEDF0F68_10 = New Led()
        Me.LEDF0F68_11 = New Led()
        Me.LEDF0F68_12 = New Led()
        Me.LEDF0F68_13 = New Led()
        Me.LEDF0F68_14 = New Led()
        Me.LEDF0F68_15 = New Led()
        Me.LEDF0F68_16 = New Led()
        Me.LEDF0F68_17 = New Led()
        Me.LEDF0F68_18 = New Led()
        Me.LEDF0F68_19 = New Led()
        Me.LEDF0F68_20 = New Led()
        Me.LEDF0F68_21 = New Led()
        Me.LEDF0F68_22 = New Led()
        Me.LEDF0F68_23 = New Led()
        Me.LEDF0F68_24 = New Led()
        Me.LEDF0F68_25 = New Led()
        Me.LEDF0F68_26 = New Led()
        Me.LEDF0F68_27 = New Led()
        Me.LEDF0F68_28 = New Led()
        Me.LabelFunctions_1 = New System.Windows.Forms.Label()
        Me.LabelFunctions_2 = New System.Windows.Forms.Label()
        Me.LabelFunctions_3 = New System.Windows.Forms.Label()
        Me.FrameVersion = New System.Windows.Forms.GroupBox()
        Me.labeFirmwareVersion = New System.Windows.Forms.Label()
        Me.txtXZ21Firmware = New System.Windows.Forms.Label()
        Me.LabelMasterVersion = New System.Windows.Forms.Label()
        Me.txtMasterVersion = New System.Windows.Forms.Label()
        Me.FrameAnAus = New System.Windows.Forms.GroupBox()
        Me.CmdAllesAn = New System.Windows.Forms.Button()
        Me.CmdAllesAus = New System.Windows.Forms.Button()
        Me.TimerMessageCounter = New System.Windows.Forms.Timer(Me.components)
        Me.FrameNumberOfmessages = New System.Windows.Forms.GroupBox()
        Me.LabelZ21tMessageCounter = New System.Windows.Forms.Label()
        Me.txtNmbPckts = New System.Windows.Forms.Label()
        Me.FrameExtAcc = New System.Windows.Forms.GroupBox()
        Me.TxtExtAccSwitch = New System.Windows.Forms.TextBox()
        Me.TxtExtAccValue = New System.Windows.Forms.TextBox()
        Me.CmdExtAccSet = New System.Windows.Forms.Button()
        Me.LabelExtSwitch = New System.Windows.Forms.Label()
        Me.LabelExtValue = New System.Windows.Forms.Label()
        Me.FrameAcc = New System.Windows.Forms.GroupBox()
        Me.cmdRecht = New System.Windows.Forms.Button()
        Me.cmdAf = New System.Windows.Forms.Button()
        Me.txtWissel = New System.Windows.Forms.TextBox()
        Me.cmdSwitchFeedback = New System.Windows.Forms.Button()
        Me.TxtBoxSwitchAddress = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchData = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchTtBits = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchIBit = New System.Windows.Forms.TextBox()
        Me.CheckActivate = New System.Windows.Forms.CheckBox()
        Me.CheckAutoDeactivate = New System.Windows.Forms.CheckBox()
        Me.LabeLabelInfo1 = New System.Windows.Forms.Label()
        Me.LabelSwitch = New System.Windows.Forms.Label()
        Me.LabelSwitchAddress = New System.Windows.Forms.Label()
        Me.LabelSwitchdata = New System.Windows.Forms.Label()
        Me.LabelSwitchTtBits = New System.Windows.Forms.Label()
        Me.LabelSwitchIBit = New System.Windows.Forms.Label()
        Me.FrameStatus = New System.Windows.Forms.GroupBox()
        Me.ledNotHalt = New Led()
        Me.ledNotAus = New Led()
        Me.ledProgrammiermode = New Led()
        Me.LabelNotHalt = New System.Windows.Forms.Label()
        Me.LabelNotAus = New System.Windows.Forms.Label()
        Me.LabelProgrammiermode = New System.Windows.Forms.Label()
        Me.FrameUDP = New System.Windows.Forms.GroupBox()
        Me.cmdUdpSet = New System.Windows.Forms.Button()
        Me.txtIP = New System.Windows.Forms.TextBox()
        Me.txtUdpPort = New System.Windows.Forms.TextBox()
        Me.LabelIpAddress = New System.Windows.Forms.Label()
        Me.LabelUdpPort = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.FrameLoks.SuspendLayout()
        Me.FrameVersion.SuspendLayout()
        Me.FrameAnAus.SuspendLayout()
        Me.FrameNumberOfmessages.SuspendLayout()
        Me.FrameExtAcc.SuspendLayout()
        Me.FrameAcc.SuspendLayout()
        Me.FrameStatus.SuspendLayout()
        Me.FrameUDP.SuspendLayout()
        Me.SuspendLayout()
        '
        'LabelStatus
        '
        Me.LabelStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStatus.Location = New System.Drawing.Point(296, 536)
        Me.LabelStatus.Name = "LabelStatus"
        Me.LabelStatus.Size = New System.Drawing.Size(417, 49)
        Me.LabelStatus.TabIndex = 66
        '
        'CmdAlleLoksAnhalten
        '
        Me.CmdAlleLoksAnhalten.Location = New System.Drawing.Point(144, 208)
        Me.CmdAlleLoksAnhalten.Name = "CmdAlleLoksAnhalten"
        Me.CmdAlleLoksAnhalten.Size = New System.Drawing.Size(107, 33)
        Me.CmdAlleLoksAnhalten.TabIndex = 57
        Me.CmdAlleLoksAnhalten.Text = "Alle Loks anhalten"
        Me.CmdAlleLoksAnhalten.UseVisualStyleBackColor = True
        '
        'CmdEineLokAnhalten
        '
        Me.CmdEineLokAnhalten.Location = New System.Drawing.Point(8, 168)
        Me.CmdEineLokAnhalten.Name = "CmdEineLokAnhalten"
        Me.CmdEineLokAnhalten.Size = New System.Drawing.Size(97, 25)
        Me.CmdEineLokAnhalten.TabIndex = 56
        Me.CmdEineLokAnhalten.Text = "Anhalten"
        Me.CmdEineLokAnhalten.UseVisualStyleBackColor = True
        '
        'txtLokAddress
        '
        Me.txtLokAddress.Location = New System.Drawing.Point(56, 24)
        Me.txtLokAddress.Name = "txtLokAddress"
        Me.txtLokAddress.Size = New System.Drawing.Size(49, 25)
        Me.txtLokAddress.TabIndex = 55
        Me.txtLokAddress.Text = "3"
        '
        'TxtLokSpeed
        '
        Me.TxtLokSpeed.Location = New System.Drawing.Point(56, 104)
        Me.TxtLokSpeed.Name = "TxtLokSpeed"
        Me.TxtLokSpeed.Size = New System.Drawing.Size(49, 25)
        Me.TxtLokSpeed.TabIndex = 54
        Me.TxtLokSpeed.Text = "0"
        '
        'ComboLokSteps
        '
        Me.ComboLokSteps.Location = New System.Drawing.Point(56, 208)
        Me.ComboLokSteps.Name = "ComboLokSteps"
        Me.ComboLokSteps.Size = New System.Drawing.Size(49, 21)
        Me.ComboLokSteps.TabIndex = 53
        Me.ComboLokSteps.Text = "28"
        '
        'CmdLokReverse
        '
        Me.CmdLokReverse.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdLokReverse.Location = New System.Drawing.Point(6, 136)
        Me.CmdLokReverse.Name = "CmdLokReverse"
        Me.CmdLokReverse.Size = New System.Drawing.Size(46, 25)
        Me.CmdLokReverse.TabIndex = 52
        Me.CmdLokReverse.Text = "<="
        Me.CmdLokReverse.UseVisualStyleBackColor = True
        '
        'CmdLokForward
        '
        Me.CmdLokForward.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdLokForward.Location = New System.Drawing.Point(59, 136)
        Me.CmdLokForward.Name = "CmdLokForward"
        Me.CmdLokForward.Size = New System.Drawing.Size(46, 25)
        Me.CmdLokForward.TabIndex = 51
        Me.CmdLokForward.Text = "=>"
        Me.CmdLokForward.UseVisualStyleBackColor = True
        '
        'TxtLokFNumber
        '
        Me.TxtLokFNumber.Location = New System.Drawing.Point(160, 24)
        Me.TxtLokFNumber.Name = "TxtLokFNumber"
        Me.TxtLokFNumber.Size = New System.Drawing.Size(25, 25)
        Me.TxtLokFNumber.TabIndex = 50
        Me.TxtLokFNumber.Text = "1"
        '
        'CmdFunctionOn
        '
        Me.CmdFunctionOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdFunctionOn.Location = New System.Drawing.Point(192, 24)
        Me.CmdFunctionOn.Name = "CmdFunctionOn"
        Me.CmdFunctionOn.Size = New System.Drawing.Size(25, 25)
        Me.CmdFunctionOn.TabIndex = 49
        Me.CmdFunctionOn.Text = "+"
        Me.CmdFunctionOn.UseVisualStyleBackColor = True
        '
        'CmdFunctionOff
        '
        Me.CmdFunctionOff.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdFunctionOff.Location = New System.Drawing.Point(232, 24)
        Me.CmdFunctionOff.Name = "CmdFunctionOff"
        Me.CmdFunctionOff.Size = New System.Drawing.Size(25, 25)
        Me.CmdFunctionOff.TabIndex = 48
        Me.CmdFunctionOff.Text = "-"
        Me.CmdFunctionOff.UseVisualStyleBackColor = True
        '
        'CmdGetLokInfo
        '
        Me.CmdGetLokInfo.Location = New System.Drawing.Point(264, 48)
        Me.CmdGetLokInfo.Name = "CmdGetLokInfo"
        Me.CmdGetLokInfo.Size = New System.Drawing.Size(40, 25)
        Me.CmdGetLokInfo.TabIndex = 47
        Me.CmdGetLokInfo.Text = "Get"
        Me.CmdGetLokInfo.UseVisualStyleBackColor = True
        '
        'CmdLokEnter
        '
        Me.CmdLokEnter.Location = New System.Drawing.Point(56, 56)
        Me.CmdLokEnter.Name = "CmdLokEnter"
        Me.CmdLokEnter.Size = New System.Drawing.Size(49, 25)
        Me.CmdLokEnter.TabIndex = 46
        Me.CmdLokEnter.Text = "Enter"
        Me.CmdLokEnter.UseVisualStyleBackColor = True
        '
        'CmdLokAusStack
        '
        Me.CmdLokAusStack.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdLokAusStack.Location = New System.Drawing.Point(111, 24)
        Me.CmdLokAusStack.Name = "CmdLokAusStack"
        Me.CmdLokAusStack.Size = New System.Drawing.Size(20, 22)
        Me.CmdLokAusStack.TabIndex = 45
        Me.CmdLokAusStack.Text = "-"
        Me.CmdLokAusStack.UseVisualStyleBackColor = True
        '
        'LabelLokAddress
        '
        Me.LabelLokAddress.Location = New System.Drawing.Point(2, 24)
        Me.LabelLokAddress.Name = "LabelLokAddress"
        Me.LabelLokAddress.AutoSize = False
        Me.LabelLokAddress.Size = New System.Drawing.Size(47, 15)
        Me.LabelLokAddress.TabIndex = 65
        Me.LabelLokAddress.Text = "Address"
        Me.LabelLokAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelLokSpeed
        '
        Me.LabelLokSpeed.Location = New System.Drawing.Point(10, 104)
        Me.LabelLokSpeed.Name = "LabelLokSpeed"
        Me.LabelLokSpeed.AutoSize = False
        Me.LabelLokSpeed.Size = New System.Drawing.Size(39, 15)
        Me.LabelLokSpeed.TabIndex = 64
        Me.LabelLokSpeed.Text = "Speed"
        Me.LabelLokSpeed.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelLokSteps
        '
        Me.LabelLokSteps.Location = New System.Drawing.Point(6, 208)
        Me.LabelLokSteps.Name = "LabelLokSteps"
        Me.LabelLokSteps.AutoSize = False
        Me.LabelLokSteps.Size = New System.Drawing.Size(35, 15)
        Me.LabelLokSteps.TabIndex = 63
        Me.LabelLokSteps.Text = "Steps"
        Me.LabelLokSteps.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunktion
        '
        Me.LabelFunktion.Location = New System.Drawing.Point(144, 27)
        Me.LabelFunktion.Name = "LabelFunktion"
        Me.LabelFunktion.AutoSize = False
        Me.LabelFunktion.Size = New System.Drawing.Size(16, 16)
        Me.LabelFunktion.TabIndex = 62
        Me.LabelFunktion.Text = "F:"
        Me.LabelFunktion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LEDF0F68_0
        '
        Me.LEDF0F68_0.Location = New System.Drawing.Point(208, 56)
        Me.LEDF0F68_0.Name = "LEDF0F68_0"
        Me.LEDF0F68_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_0.Round = False
        Me.LEDF0F68_0.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_1
        '
        Me.LEDF0F68_1.Location = New System.Drawing.Point(224, 56)
        Me.LEDF0F68_1.Name = "LEDF0F68_1"
        Me.LEDF0F68_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_1.Round = False
        Me.LEDF0F68_1.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_2
        '
        Me.LEDF0F68_2.Location = New System.Drawing.Point(232, 56)
        Me.LEDF0F68_2.Name = "LEDF0F68_2"
        Me.LEDF0F68_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_2.Round = False
        Me.LEDF0F68_2.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_3
        '
        Me.LEDF0F68_3.Location = New System.Drawing.Point(240, 56)
        Me.LEDF0F68_3.Name = "LEDF0F68_3"
        Me.LEDF0F68_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_3.Round = False
        Me.LEDF0F68_3.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_4
        '
        Me.LEDF0F68_4.Location = New System.Drawing.Point(248, 56)
        Me.LEDF0F68_4.Name = "LEDF0F68_4"
        Me.LEDF0F68_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_4.Round = False
        Me.LEDF0F68_4.Size = New System.Drawing.Size(9, 9)
        '
        'LabelFunctions_0
        '
        Me.LabelFunctions_0.Location = New System.Drawing.Point(143, 55)
        Me.LabelFunctions_0.Name = "LabelFunctions_0"
        Me.LabelFunctions_0.AutoSize = False
        Me.LabelFunctions_0.Size = New System.Drawing.Size(36, 15)
        Me.LabelFunctions_0.TabIndex = 61
        Me.LabelFunctions_0.Text = "F0-F4"
        Me.LabelFunctions_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LEDF0F68_5
        '
        Me.LEDF0F68_5.Location = New System.Drawing.Point(184, 72)
        Me.LEDF0F68_5.Name = "LEDF0F68_5"
        Me.LEDF0F68_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_5.Round = False
        Me.LEDF0F68_5.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_6
        '
        Me.LEDF0F68_6.Location = New System.Drawing.Point(192, 72)
        Me.LEDF0F68_6.Name = "LEDF0F68_6"
        Me.LEDF0F68_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_6.Round = False
        Me.LEDF0F68_6.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_7
        '
        Me.LEDF0F68_7.Location = New System.Drawing.Point(200, 72)
        Me.LEDF0F68_7.Name = "LEDF0F68_7"
        Me.LEDF0F68_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_7.Round = False
        Me.LEDF0F68_7.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_8
        '
        Me.LEDF0F68_8.Location = New System.Drawing.Point(208, 72)
        Me.LEDF0F68_8.Name = "LEDF0F68_8"
        Me.LEDF0F68_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_8.Round = False
        Me.LEDF0F68_8.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_9
        '
        Me.LEDF0F68_9.Location = New System.Drawing.Point(224, 72)
        Me.LEDF0F68_9.Name = "LEDF0F68_9"
        Me.LEDF0F68_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_9.Round = False
        Me.LEDF0F68_9.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_10
        '
        Me.LEDF0F68_10.Location = New System.Drawing.Point(232, 72)
        Me.LEDF0F68_10.Name = "LEDF0F68_10"
        Me.LEDF0F68_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_10.Round = False
        Me.LEDF0F68_10.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_11
        '
        Me.LEDF0F68_11.Location = New System.Drawing.Point(240, 72)
        Me.LEDF0F68_11.Name = "LEDF0F68_11"
        Me.LEDF0F68_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_11.Round = False
        Me.LEDF0F68_11.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_12
        '
        Me.LEDF0F68_12.Location = New System.Drawing.Point(248, 72)
        Me.LEDF0F68_12.Name = "LEDF0F68_12"
        Me.LEDF0F68_12.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_12.Round = False
        Me.LEDF0F68_12.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_13
        '
        Me.LEDF0F68_13.Location = New System.Drawing.Point(184, 88)
        Me.LEDF0F68_13.Name = "LEDF0F68_13"
        Me.LEDF0F68_13.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_13.Round = False
        Me.LEDF0F68_13.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_14
        '
        Me.LEDF0F68_14.Location = New System.Drawing.Point(192, 88)
        Me.LEDF0F68_14.Name = "LEDF0F68_14"
        Me.LEDF0F68_14.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_14.Round = False
        Me.LEDF0F68_14.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_15
        '
        Me.LEDF0F68_15.Location = New System.Drawing.Point(200, 88)
        Me.LEDF0F68_15.Name = "LEDF0F68_15"
        Me.LEDF0F68_15.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_15.Round = False
        Me.LEDF0F68_15.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_16
        '
        Me.LEDF0F68_16.Location = New System.Drawing.Point(208, 88)
        Me.LEDF0F68_16.Name = "LEDF0F68_16"
        Me.LEDF0F68_16.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_16.Round = False
        Me.LEDF0F68_16.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_17
        '
        Me.LEDF0F68_17.Location = New System.Drawing.Point(224, 88)
        Me.LEDF0F68_17.Name = "LEDF0F68_17"
        Me.LEDF0F68_17.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_17.Round = False
        Me.LEDF0F68_17.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_18
        '
        Me.LEDF0F68_18.Location = New System.Drawing.Point(232, 88)
        Me.LEDF0F68_18.Name = "LEDF0F68_18"
        Me.LEDF0F68_18.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_18.Round = False
        Me.LEDF0F68_18.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_19
        '
        Me.LEDF0F68_19.Location = New System.Drawing.Point(240, 88)
        Me.LEDF0F68_19.Name = "LEDF0F68_19"
        Me.LEDF0F68_19.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_19.Round = False
        Me.LEDF0F68_19.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_20
        '
        Me.LEDF0F68_20.Location = New System.Drawing.Point(248, 88)
        Me.LEDF0F68_20.Name = "LEDF0F68_20"
        Me.LEDF0F68_20.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_20.Round = False
        Me.LEDF0F68_20.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_21
        '
        Me.LEDF0F68_21.Location = New System.Drawing.Point(184, 104)
        Me.LEDF0F68_21.Name = "LEDF0F68_21"
        Me.LEDF0F68_21.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_21.Round = False
        Me.LEDF0F68_21.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_22
        '
        Me.LEDF0F68_22.Location = New System.Drawing.Point(192, 104)
        Me.LEDF0F68_22.Name = "LEDF0F68_22"
        Me.LEDF0F68_22.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_22.Round = False
        Me.LEDF0F68_22.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_23
        '
        Me.LEDF0F68_23.Location = New System.Drawing.Point(200, 104)
        Me.LEDF0F68_23.Name = "LEDF0F68_23"
        Me.LEDF0F68_23.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_23.Round = False
        Me.LEDF0F68_23.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_24
        '
        Me.LEDF0F68_24.Location = New System.Drawing.Point(208, 104)
        Me.LEDF0F68_24.Name = "LEDF0F68_24"
        Me.LEDF0F68_24.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_24.Round = False
        Me.LEDF0F68_24.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_25
        '
        Me.LEDF0F68_25.Location = New System.Drawing.Point(224, 104)
        Me.LEDF0F68_25.Name = "LEDF0F68_25"
        Me.LEDF0F68_25.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_25.Round = False
        Me.LEDF0F68_25.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_26
        '
        Me.LEDF0F68_26.Location = New System.Drawing.Point(232, 104)
        Me.LEDF0F68_26.Name = "LEDF0F68_26"
        Me.LEDF0F68_26.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_26.Round = False
        Me.LEDF0F68_26.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_27
        '
        Me.LEDF0F68_27.Location = New System.Drawing.Point(240, 104)
        Me.LEDF0F68_27.Name = "LEDF0F68_27"
        Me.LEDF0F68_27.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_27.Round = False
        Me.LEDF0F68_27.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_28
        '
        Me.LEDF0F68_28.Location = New System.Drawing.Point(248, 104)
        Me.LEDF0F68_28.Name = "LEDF0F68_28"
        Me.LEDF0F68_28.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_28.Round = False
        Me.LEDF0F68_28.Size = New System.Drawing.Size(9, 9)
        '
        'LabelFunctions_1
        '
        Me.LabelFunctions_1.Location = New System.Drawing.Point(137, 72)
        Me.LabelFunctions_1.Name = "LabelFunctions_1"
        Me.LabelFunctions_1.AutoSize = False
        Me.LabelFunctions_1.Size = New System.Drawing.Size(42, 15)
        Me.LabelFunctions_1.TabIndex = 60
        Me.LabelFunctions_1.Text = "F5-F12"
        Me.LabelFunctions_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_2
        '
        Me.LabelFunctions_2.Location = New System.Drawing.Point(129, 88)
        Me.LabelFunctions_2.Name = "LabelFunctions_2"
        Me.LabelFunctions_2.AutoSize = False
        Me.LabelFunctions_2.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_2.TabIndex = 59
        Me.LabelFunctions_2.Text = "F13-F20"
        Me.LabelFunctions_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_3
        '
        Me.LabelFunctions_3.Location = New System.Drawing.Point(127, 104)
        Me.LabelFunctions_3.Name = "LabelFunctions_3"
        Me.LabelFunctions_3.AutoSize = False
        Me.LabelFunctions_3.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_3.TabIndex = 58
        Me.LabelFunctions_3.Text = "F21-F28"
        Me.LabelFunctions_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'FrameLoks
        '
        Me.FrameLoks.Controls.Add(Me.CmdAlleLoksAnhalten)
        Me.FrameLoks.Controls.Add(Me.CmdEineLokAnhalten)
        Me.FrameLoks.Controls.Add(Me.txtLokAddress)
        Me.FrameLoks.Controls.Add(Me.TxtLokSpeed)
        Me.FrameLoks.Controls.Add(Me.ComboLokSteps)
        Me.FrameLoks.Controls.Add(Me.CmdLokReverse)
        Me.FrameLoks.Controls.Add(Me.CmdLokForward)
        Me.FrameLoks.Controls.Add(Me.TxtLokFNumber)
        Me.FrameLoks.Controls.Add(Me.CmdFunctionOn)
        Me.FrameLoks.Controls.Add(Me.CmdFunctionOff)
        Me.FrameLoks.Controls.Add(Me.CmdGetLokInfo)
        Me.FrameLoks.Controls.Add(Me.CmdLokEnter)
        Me.FrameLoks.Controls.Add(Me.CmdLokAusStack)
        Me.FrameLoks.Controls.Add(Me.LabelLokAddress)
        Me.FrameLoks.Controls.Add(Me.LabelLokSpeed)
        Me.FrameLoks.Controls.Add(Me.LabelLokSteps)
        Me.FrameLoks.Controls.Add(Me.LabelFunktion)
        Me.FrameLoks.Controls.Add(Me.LabelFunctions_0)
        Me.FrameLoks.Controls.Add(Me.LabelFunctions_1)
        Me.FrameLoks.Controls.Add(Me.LabelFunctions_2)
        Me.FrameLoks.Controls.Add(Me.LabelFunctions_3)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_0)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_1)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_2)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_3)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_4)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_5)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_6)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_7)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_8)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_9)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_10)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_11)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_12)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_13)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_14)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_15)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_16)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_17)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_18)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_19)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_20)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_21)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_22)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_23)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_24)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_25)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_26)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_27)
        Me.FrameLoks.Controls.Add(Me.LEDF0F68_28)
        Me.FrameLoks.Location = New System.Drawing.Point(296, 272)
        Me.FrameLoks.Name = "FrameLoks"
        Me.FrameLoks.Size = New System.Drawing.Size(312, 257)
        Me.FrameLoks.TabIndex = 44
        Me.FrameLoks.Text = "Loks"
        '
        'labeFirmwareVersion
        '
        Me.labeFirmwareVersion.Location = New System.Drawing.Point(-8, 32)
        Me.labeFirmwareVersion.Name = "labeFirmwareVersion"
        Me.labeFirmwareVersion.AutoSize = False
        Me.labeFirmwareVersion.Size = New System.Drawing.Size(118, 15)
        Me.labeFirmwareVersion.TabIndex = 43
        Me.labeFirmwareVersion.Text = "Z21 Firmware Version:"
        Me.labeFirmwareVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtXZ21Firmware
        '
        Me.txtXZ21Firmware.Location = New System.Drawing.Point(113, 32)
        Me.txtXZ21Firmware.Name = "txtXZ21Firmware"
        Me.txtXZ21Firmware.Size = New System.Drawing.Size(57, 17)
        Me.txtXZ21Firmware.TabIndex = 42
        Me.txtXZ21Firmware.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelMasterVersion
        '
        Me.LabelMasterVersion.Location = New System.Drawing.Point(29, 16)
        Me.LabelMasterVersion.Name = "LabelMasterVersion"
        Me.LabelMasterVersion.AutoSize = False
        Me.LabelMasterVersion.Size = New System.Drawing.Size(84, 15)
        Me.LabelMasterVersion.TabIndex = 41
        Me.LabelMasterVersion.Text = "Master Version:"
        Me.LabelMasterVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtMasterVersion
        '
        Me.txtMasterVersion.Location = New System.Drawing.Point(113, 16)
        Me.txtMasterVersion.Name = "txtMasterVersion"
        Me.txtMasterVersion.Size = New System.Drawing.Size(57, 17)
        Me.txtMasterVersion.TabIndex = 40
        Me.txtMasterVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'FrameVersion
        '
        Me.FrameVersion.Controls.Add(Me.labeFirmwareVersion)
        Me.FrameVersion.Controls.Add(Me.txtXZ21Firmware)
        Me.FrameVersion.Controls.Add(Me.LabelMasterVersion)
        Me.FrameVersion.Controls.Add(Me.txtMasterVersion)
        Me.FrameVersion.Location = New System.Drawing.Point(528, 8)
        Me.FrameVersion.Name = "FrameVersion"
        Me.FrameVersion.Size = New System.Drawing.Size(185, 57)
        Me.FrameVersion.TabIndex = 39
        '
        'CmdAllesAn
        '
        Me.CmdAllesAn.Location = New System.Drawing.Point(16, 24)
        Me.CmdAllesAn.Name = "CmdAllesAn"
        Me.CmdAllesAn.Size = New System.Drawing.Size(107, 33)
        Me.CmdAllesAn.TabIndex = 38
        Me.CmdAllesAn.Text = "Alles An"
        Me.CmdAllesAn.UseVisualStyleBackColor = True
        '
        'CmdAllesAus
        '
        Me.CmdAllesAus.Location = New System.Drawing.Point(16, 64)
        Me.CmdAllesAus.Name = "CmdAllesAus"
        Me.CmdAllesAus.Size = New System.Drawing.Size(107, 33)
        Me.CmdAllesAus.TabIndex = 37
        Me.CmdAllesAus.Text = "Alles Aus"
        Me.CmdAllesAus.UseVisualStyleBackColor = True
        '
        'FrameAnAus
        '
        Me.FrameAnAus.Controls.Add(Me.CmdAllesAn)
        Me.FrameAnAus.Controls.Add(Me.CmdAllesAus)
        Me.FrameAnAus.Location = New System.Drawing.Point(376, 8)
        Me.FrameAnAus.Name = "FrameAnAus"
        Me.FrameAnAus.Size = New System.Drawing.Size(145, 121)
        Me.FrameAnAus.TabIndex = 36
        Me.FrameAnAus.Text = "An / Aus"
        '
        'TimerMessageCounter
        '
        Me.TimerMessageCounter.Interval = 1000
        Me.TimerMessageCounter.Enabled = True
        '
        'LabelZ21tMessageCounter
        '
        Me.LabelZ21tMessageCounter.Location = New System.Drawing.Point(-4, 16)
        Me.LabelZ21tMessageCounter.Name = "LabelZ21tMessageCounter"
        Me.LabelZ21tMessageCounter.AutoSize = False
        Me.LabelZ21tMessageCounter.Size = New System.Drawing.Size(134, 15)
        Me.LabelZ21tMessageCounter.TabIndex = 34
        Me.LabelZ21tMessageCounter.Text = "Number of Z21Messages:"
        Me.LabelZ21tMessageCounter.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtNmbPckts
        '
        Me.txtNmbPckts.Location = New System.Drawing.Point(136, 16)
        Me.txtNmbPckts.Name = "txtNmbPckts"
        Me.txtNmbPckts.Size = New System.Drawing.Size(41, 17)
        Me.txtNmbPckts.TabIndex = 33
        Me.txtNmbPckts.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'FrameNumberOfmessages
        '
        Me.FrameNumberOfmessages.Controls.Add(Me.LabelZ21tMessageCounter)
        Me.FrameNumberOfmessages.Controls.Add(Me.txtNmbPckts)
        Me.FrameNumberOfmessages.Location = New System.Drawing.Point(528, 72)
        Me.FrameNumberOfmessages.Name = "FrameNumberOfmessages"
        Me.FrameNumberOfmessages.Size = New System.Drawing.Size(185, 41)
        Me.FrameNumberOfmessages.TabIndex = 32
        '
        'TxtExtAccSwitch
        '
        Me.TxtExtAccSwitch.Location = New System.Drawing.Point(72, 24)
        Me.TxtExtAccSwitch.Name = "TxtExtAccSwitch"
        Me.TxtExtAccSwitch.Size = New System.Drawing.Size(57, 25)
        Me.TxtExtAccSwitch.TabIndex = 29
        Me.TxtExtAccSwitch.Text = ""
        '
        'TxtExtAccValue
        '
        Me.TxtExtAccValue.Location = New System.Drawing.Point(72, 56)
        Me.TxtExtAccValue.Name = "TxtExtAccValue"
        Me.TxtExtAccValue.Size = New System.Drawing.Size(57, 25)
        Me.TxtExtAccValue.TabIndex = 28
        Me.TxtExtAccValue.Text = ""
        '
        'CmdExtAccSet
        '
        Me.CmdExtAccSet.Location = New System.Drawing.Point(72, 88)
        Me.CmdExtAccSet.Name = "CmdExtAccSet"
        Me.CmdExtAccSet.Size = New System.Drawing.Size(57, 25)
        Me.CmdExtAccSet.TabIndex = 27
        Me.CmdExtAccSet.Text = "Set"
        Me.CmdExtAccSet.UseVisualStyleBackColor = True
        '
        'LabelExtSwitch
        '
        Me.LabelExtSwitch.Location = New System.Drawing.Point(24, 24)
        Me.LabelExtSwitch.Name = "LabelExtSwitch"
        Me.LabelExtSwitch.AutoSize = True
        Me.LabelExtSwitch.Size = New System.Drawing.Size(51, 16)
        Me.LabelExtSwitch.TabIndex = 31
        Me.LabelExtSwitch.Text = "Address"
        Me.LabelExtSwitch.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelExtValue
        '
        Me.LabelExtValue.Location = New System.Drawing.Point(27, 56)
        Me.LabelExtValue.Name = "LabelExtValue"
        Me.LabelExtValue.AutoSize = True
        Me.LabelExtValue.Size = New System.Drawing.Size(35, 16)
        Me.LabelExtValue.TabIndex = 30
        Me.LabelExtValue.Text = "Value"
        Me.LabelExtValue.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'FrameExtAcc
        '
        Me.FrameExtAcc.Controls.Add(Me.TxtExtAccSwitch)
        Me.FrameExtAcc.Controls.Add(Me.TxtExtAccValue)
        Me.FrameExtAcc.Controls.Add(Me.CmdExtAccSet)
        Me.FrameExtAcc.Controls.Add(Me.LabelExtSwitch)
        Me.FrameExtAcc.Controls.Add(Me.LabelExtValue)
        Me.FrameExtAcc.Location = New System.Drawing.Point(296, 136)
        Me.FrameExtAcc.Name = "FrameExtAcc"
        Me.FrameExtAcc.Size = New System.Drawing.Size(145, 129)
        Me.FrameExtAcc.TabIndex = 26
        Me.FrameExtAcc.Text = "Extended Accessory"
        '
        'cmdRecht
        '
        Me.cmdRecht.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRecht.Location = New System.Drawing.Point(216, 24)
        Me.cmdRecht.Name = "cmdRecht"
        Me.cmdRecht.Size = New System.Drawing.Size(49, 25)
        Me.cmdRecht.TabIndex = 19
        Me.cmdRecht.Text = "=>"
        Me.cmdRecht.UseVisualStyleBackColor = True
        '
        'cmdAf
        '
        Me.cmdAf.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAf.Location = New System.Drawing.Point(160, 24)
        Me.cmdAf.Name = "cmdAf"
        Me.cmdAf.Size = New System.Drawing.Size(49, 25)
        Me.cmdAf.TabIndex = 18
        Me.cmdAf.Text = "<="
        Me.cmdAf.UseVisualStyleBackColor = True
        '
        'txtWissel
        '
        Me.txtWissel.Location = New System.Drawing.Point(64, 24)
        Me.txtWissel.Name = "txtWissel"
        Me.txtWissel.Size = New System.Drawing.Size(81, 25)
        Me.txtWissel.TabIndex = 17
        Me.txtWissel.Text = ""
        '
        'cmdSwitchFeedback
        '
        Me.cmdSwitchFeedback.Location = New System.Drawing.Point(64, 64)
        Me.cmdSwitchFeedback.Name = "cmdSwitchFeedback"
        Me.cmdSwitchFeedback.Size = New System.Drawing.Size(81, 33)
        Me.cmdSwitchFeedback.TabIndex = 16
        Me.cmdSwitchFeedback.Text = "Feedback"
        Me.cmdSwitchFeedback.UseVisualStyleBackColor = True
        '
        'TxtBoxSwitchAddress
        '
        Me.TxtBoxSwitchAddress.Location = New System.Drawing.Point(24, 144)
        Me.TxtBoxSwitchAddress.Name = "TxtBoxSwitchAddress"
        Me.TxtBoxSwitchAddress.Size = New System.Drawing.Size(97, 265)
        Me.TxtBoxSwitchAddress.TabIndex = 15
        Me.TxtBoxSwitchAddress.Text = ""
        Me.TxtBoxSwitchAddress.Multiline = True
        '
        'TxtBoxSwitchData
        '
        Me.TxtBoxSwitchData.Location = New System.Drawing.Point(120, 144)
        Me.TxtBoxSwitchData.Name = "TxtBoxSwitchData"
        Me.TxtBoxSwitchData.Size = New System.Drawing.Size(73, 265)
        Me.TxtBoxSwitchData.TabIndex = 14
        Me.TxtBoxSwitchData.Text = ""
        Me.TxtBoxSwitchData.Multiline = True
        '
        'TxtBoxSwitchTtBits
        '
        Me.TxtBoxSwitchTtBits.Location = New System.Drawing.Point(192, 144)
        Me.TxtBoxSwitchTtBits.Name = "TxtBoxSwitchTtBits"
        Me.TxtBoxSwitchTtBits.Size = New System.Drawing.Size(41, 265)
        Me.TxtBoxSwitchTtBits.TabIndex = 13
        Me.TxtBoxSwitchTtBits.Text = ""
        Me.TxtBoxSwitchTtBits.Multiline = True
        '
        'TxtBoxSwitchIBit
        '
        Me.TxtBoxSwitchIBit.Location = New System.Drawing.Point(232, 144)
        Me.TxtBoxSwitchIBit.Name = "TxtBoxSwitchIBit"
        Me.TxtBoxSwitchIBit.Size = New System.Drawing.Size(25, 265)
        Me.TxtBoxSwitchIBit.TabIndex = 12
        Me.TxtBoxSwitchIBit.Text = ""
        Me.TxtBoxSwitchIBit.Multiline = True
        '
        'CheckActivate
        '
        Me.CheckActivate.Location = New System.Drawing.Point(150, 56)
        Me.CheckActivate.Name = "CheckActivate"
        Me.CheckActivate.Size = New System.Drawing.Size(126, 25)
        Me.CheckActivate.TabIndex = 11
        Me.CheckActivate.Text = "activate"
        Me.CheckActivate.Checked = True
        Me.CheckActivate.UseVisualStyleBackColor = True
        '
        'CheckAutoDeactivate
        '
        Me.CheckAutoDeactivate.Location = New System.Drawing.Point(150, 80)
        Me.CheckAutoDeactivate.Name = "CheckAutoDeactivate"
        Me.CheckAutoDeactivate.Size = New System.Drawing.Size(126, 25)
        Me.CheckAutoDeactivate.TabIndex = 10
        Me.CheckAutoDeactivate.Text = "auto deactivate"
        Me.CheckAutoDeactivate.Checked = True
        Me.CheckAutoDeactivate.UseVisualStyleBackColor = True
        '
        'LabeLabelInfo1
        '
        Me.LabeLabelInfo1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeLabelInfo1.Location = New System.Drawing.Point(72, 408)
        Me.LabeLabelInfo1.Name = "LabeLabelInfo1"
        Me.LabeLabelInfo1.AutoSize = True
        Me.LabeLabelInfo1.Size = New System.Drawing.Size(135, 16)
        Me.LabeLabelInfo1.TabIndex = 25
        Me.LabeLabelInfo1.Text = "latest message at top"
        Me.LabeLabelInfo1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelSwitch
        '
        Me.LabelSwitch.Location = New System.Drawing.Point(16, 24)
        Me.LabelSwitch.Name = "LabelSwitch"
        Me.LabelSwitch.AutoSize = True
        Me.LabelSwitch.Size = New System.Drawing.Size(41, 16)
        Me.LabelSwitch.TabIndex = 24
        Me.LabelSwitch.Text = "Switch:"
        Me.LabelSwitch.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelSwitchAddress
        '
        Me.LabelSwitchAddress.Location = New System.Drawing.Point(24, 120)
        Me.LabelSwitchAddress.Name = "LabelSwitchAddress"
        Me.LabelSwitchAddress.AutoSize = False
        Me.LabelSwitchAddress.Size = New System.Drawing.Size(97, 16)
        Me.LabelSwitchAddress.TabIndex = 23
        Me.LabelSwitchAddress.Text = "Address"
        Me.LabelSwitchAddress.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchdata
        '
        Me.LabelSwitchdata.Location = New System.Drawing.Point(120, 120)
        Me.LabelSwitchdata.Name = "LabelSwitchdata"
        Me.LabelSwitchdata.AutoSize = False
        Me.LabelSwitchdata.Size = New System.Drawing.Size(73, 16)
        Me.LabelSwitchdata.TabIndex = 22
        Me.LabelSwitchdata.Text = "Data"
        Me.LabelSwitchdata.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchTtBits
        '
        Me.LabelSwitchTtBits.Location = New System.Drawing.Point(192, 120)
        Me.LabelSwitchTtBits.Name = "LabelSwitchTtBits"
        Me.LabelSwitchTtBits.AutoSize = False
        Me.LabelSwitchTtBits.Size = New System.Drawing.Size(41, 16)
        Me.LabelSwitchTtBits.TabIndex = 21
        Me.LabelSwitchTtBits.Text = "TT"
        Me.LabelSwitchTtBits.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchIBit
        '
        Me.LabelSwitchIBit.Location = New System.Drawing.Point(232, 120)
        Me.LabelSwitchIBit.Name = "LabelSwitchIBit"
        Me.LabelSwitchIBit.AutoSize = False
        Me.LabelSwitchIBit.Size = New System.Drawing.Size(25, 16)
        Me.LabelSwitchIBit.TabIndex = 20
        Me.LabelSwitchIBit.Text = "I"
        Me.LabelSwitchIBit.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'FrameAcc
        '
        Me.FrameAcc.Controls.Add(Me.cmdRecht)
        Me.FrameAcc.Controls.Add(Me.cmdAf)
        Me.FrameAcc.Controls.Add(Me.txtWissel)
        Me.FrameAcc.Controls.Add(Me.cmdSwitchFeedback)
        Me.FrameAcc.Controls.Add(Me.TxtBoxSwitchAddress)
        Me.FrameAcc.Controls.Add(Me.TxtBoxSwitchData)
        Me.FrameAcc.Controls.Add(Me.TxtBoxSwitchTtBits)
        Me.FrameAcc.Controls.Add(Me.TxtBoxSwitchIBit)
        Me.FrameAcc.Controls.Add(Me.CheckActivate)
        Me.FrameAcc.Controls.Add(Me.CheckAutoDeactivate)
        Me.FrameAcc.Controls.Add(Me.LabeLabelInfo1)
        Me.FrameAcc.Controls.Add(Me.LabelSwitch)
        Me.FrameAcc.Controls.Add(Me.LabelSwitchAddress)
        Me.FrameAcc.Controls.Add(Me.LabelSwitchdata)
        Me.FrameAcc.Controls.Add(Me.LabelSwitchTtBits)
        Me.FrameAcc.Controls.Add(Me.LabelSwitchIBit)
        Me.FrameAcc.Location = New System.Drawing.Point(8, 136)
        Me.FrameAcc.Name = "FrameAcc"
        Me.FrameAcc.Size = New System.Drawing.Size(281, 433)
        Me.FrameAcc.TabIndex = 9
        Me.FrameAcc.Text = "Accessory"
        '
        'ledNotHalt
        '
        Me.ledNotHalt.Location = New System.Drawing.Point(8, 56)
        Me.ledNotHalt.Name = "ledNotHalt"
        Me.ledNotHalt.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledNotHalt.Round = True
        Me.ledNotHalt.Size = New System.Drawing.Size(17, 17)
        '
        'ledNotAus
        '
        Me.ledNotAus.Location = New System.Drawing.Point(8, 24)
        Me.ledNotAus.Name = "ledNotAus"
        Me.ledNotAus.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledNotAus.Round = True
        Me.ledNotAus.Size = New System.Drawing.Size(17, 17)
        '
        'ledProgrammiermode
        '
        Me.ledProgrammiermode.Location = New System.Drawing.Point(8, 88)
        Me.ledProgrammiermode.Name = "ledProgrammiermode"
        Me.ledProgrammiermode.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledProgrammiermode.Round = True
        Me.ledProgrammiermode.Size = New System.Drawing.Size(17, 17)
        '
        'LabelNotHalt
        '
        Me.LabelNotHalt.Location = New System.Drawing.Point(33, 58)
        Me.LabelNotHalt.Name = "LabelNotHalt"
        Me.LabelNotHalt.AutoSize = True
        Me.LabelNotHalt.Size = New System.Drawing.Size(45, 16)
        Me.LabelNotHalt.TabIndex = 8
        Me.LabelNotHalt.Text = "NotHalt"
        Me.LabelNotHalt.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelNotAus
        '
        Me.LabelNotAus.Location = New System.Drawing.Point(33, 26)
        Me.LabelNotAus.Name = "LabelNotAus"
        Me.LabelNotAus.AutoSize = True
        Me.LabelNotAus.Size = New System.Drawing.Size(44, 16)
        Me.LabelNotAus.TabIndex = 7
        Me.LabelNotAus.Text = "NotAus"
        Me.LabelNotAus.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelProgrammiermode
        '
        Me.LabelProgrammiermode.Location = New System.Drawing.Point(32, 88)
        Me.LabelProgrammiermode.Name = "LabelProgrammiermode"
        Me.LabelProgrammiermode.AutoSize = True
        Me.LabelProgrammiermode.Size = New System.Drawing.Size(92, 16)
        Me.LabelProgrammiermode.TabIndex = 6
        Me.LabelProgrammiermode.Text = "Programmier Mode"
        Me.LabelProgrammiermode.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'FrameStatus
        '
        Me.FrameStatus.Controls.Add(Me.LabelNotHalt)
        Me.FrameStatus.Controls.Add(Me.LabelNotAus)
        Me.FrameStatus.Controls.Add(Me.LabelProgrammiermode)
        Me.FrameStatus.Controls.Add(Me.ledNotHalt)
        Me.FrameStatus.Controls.Add(Me.ledNotAus)
        Me.FrameStatus.Controls.Add(Me.ledProgrammiermode)
        Me.FrameStatus.Location = New System.Drawing.Point(216, 8)
        Me.FrameStatus.Name = "FrameStatus"
        Me.FrameStatus.Size = New System.Drawing.Size(145, 121)
        Me.FrameStatus.TabIndex = 5
        Me.FrameStatus.Text = "Status"
        '
        'cmdUdpSet
        '
        Me.cmdUdpSet.Location = New System.Drawing.Point(120, 88)
        Me.cmdUdpSet.Name = "cmdUdpSet"
        Me.cmdUdpSet.Size = New System.Drawing.Size(57, 25)
        Me.cmdUdpSet.TabIndex = 35
        Me.cmdUdpSet.Text = "Set"
        Me.cmdUdpSet.UseVisualStyleBackColor = True
        '
        'txtIP
        '
        Me.txtIP.Location = New System.Drawing.Point(85, 24)
        Me.txtIP.Name = "txtIP"
        Me.txtIP.Size = New System.Drawing.Size(97, 25)
        Me.txtIP.TabIndex = 2
        Me.txtIP.Text = "192.168.0.111"
        '
        'txtUdpPort
        '
        Me.txtUdpPort.Location = New System.Drawing.Point(85, 56)
        Me.txtUdpPort.Name = "txtUdpPort"
        Me.txtUdpPort.Size = New System.Drawing.Size(97, 25)
        Me.txtUdpPort.TabIndex = 1
        Me.txtUdpPort.Text = "21105"
        '
        'LabelIpAddress
        '
        Me.LabelIpAddress.Location = New System.Drawing.Point(8, 24)
        Me.LabelIpAddress.Name = "LabelIpAddress"
        Me.LabelIpAddress.AutoSize = False
        Me.LabelIpAddress.Size = New System.Drawing.Size(66, 16)
        Me.LabelIpAddress.TabIndex = 4
        Me.LabelIpAddress.Text = "IP Address"
        Me.LabelIpAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelUdpPort
        '
        Me.LabelUdpPort.Location = New System.Drawing.Point(22, 56)
        Me.LabelUdpPort.Name = "LabelUdpPort"
        Me.LabelUdpPort.AutoSize = False
        Me.LabelUdpPort.Size = New System.Drawing.Size(53, 15)
        Me.LabelUdpPort.TabIndex = 3
        Me.LabelUdpPort.Text = "UDP Port"
        Me.LabelUdpPort.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'FrameUDP
        '
        Me.FrameUDP.Controls.Add(Me.cmdUdpSet)
        Me.FrameUDP.Controls.Add(Me.txtIP)
        Me.FrameUDP.Controls.Add(Me.txtUdpPort)
        Me.FrameUDP.Controls.Add(Me.LabelIpAddress)
        Me.FrameUDP.Controls.Add(Me.LabelUdpPort)
        Me.FrameUDP.Location = New System.Drawing.Point(8, 8)
        Me.FrameUDP.Name = "FrameUDP"
        Me.FrameUDP.Size = New System.Drawing.Size(193, 121)
        Me.FrameUDP.TabIndex = 0
        Me.FrameUDP.Text = "Z21 Network Settings"
        Me.ToolTip1.SetToolTip(Me.txtWissel, "1..2048")
        Me.ToolTip1.SetToolTip(Me.TxtBoxSwitchTtBits, "C=Command Station, D=Decoder")
        Me.ToolTip1.SetToolTip(Me.cmdUdpSet, "Set IP address and Port")
        Me.ToolTip1.SetToolTip(Me.txtUdpPort, "Yamorc: 21105 / Z21: 21106")
        Me.ToolTip1.SetToolTip(Me.LabelUdpPort, "Yamorc: 21105 / Z21: 21106")
        '
        'FrmZ21Test
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(726, 593)
        Me.Controls.Add(Me.FrameLoks)
        Me.Controls.Add(Me.FrameVersion)
        Me.Controls.Add(Me.FrameAnAus)
        Me.Controls.Add(Me.FrameNumberOfmessages)
        Me.Controls.Add(Me.FrameExtAcc)
        Me.Controls.Add(Me.FrameAcc)
        Me.Controls.Add(Me.FrameStatus)
        Me.Controls.Add(Me.FrameUDP)
        Me.Controls.Add(Me.LabelStatus)
        Me.Name = "FrmZ21Test"
        Me.Text = "TestZ21"
        Me.FrameLoks.ResumeLayout(False)
        Me.FrameLoks.PerformLayout()
        Me.FrameVersion.ResumeLayout(False)
        Me.FrameVersion.PerformLayout()
        Me.FrameAnAus.ResumeLayout(False)
        Me.FrameAnAus.PerformLayout()
        Me.FrameNumberOfmessages.ResumeLayout(False)
        Me.FrameNumberOfmessages.PerformLayout()
        Me.FrameExtAcc.ResumeLayout(False)
        Me.FrameExtAcc.PerformLayout()
        Me.FrameAcc.ResumeLayout(False)
        Me.FrameAcc.PerformLayout()
        Me.FrameStatus.ResumeLayout(False)
        Me.FrameStatus.PerformLayout()
        Me.FrameUDP.ResumeLayout(False)
        Me.FrameUDP.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FrameLoks As System.Windows.Forms.GroupBox
    Friend WithEvents CmdAlleLoksAnhalten As System.Windows.Forms.Button
    Friend WithEvents CmdEineLokAnhalten As System.Windows.Forms.Button
    Friend WithEvents txtLokAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtLokSpeed As System.Windows.Forms.TextBox
    Friend WithEvents ComboLokSteps As System.Windows.Forms.ComboBox
    Friend WithEvents CmdLokReverse As System.Windows.Forms.Button
    Friend WithEvents CmdLokForward As System.Windows.Forms.Button
    Friend WithEvents TxtLokFNumber As System.Windows.Forms.TextBox
    Friend WithEvents CmdFunctionOn As System.Windows.Forms.Button
    Friend WithEvents CmdFunctionOff As System.Windows.Forms.Button
    Friend WithEvents CmdGetLokInfo As System.Windows.Forms.Button
    Friend WithEvents CmdLokEnter As System.Windows.Forms.Button
    Friend WithEvents CmdLokAusStack As System.Windows.Forms.Button
    Friend WithEvents LabelLokAddress As System.Windows.Forms.Label
    Friend WithEvents LabelLokSpeed As System.Windows.Forms.Label
    Friend WithEvents LabelLokSteps As System.Windows.Forms.Label
    Friend WithEvents LabelFunktion As System.Windows.Forms.Label
    Friend WithEvents LEDF0F68_0 As Led
    Friend WithEvents LEDF0F68_1 As Led
    Friend WithEvents LEDF0F68_2 As Led
    Friend WithEvents LEDF0F68_3 As Led
    Friend WithEvents LEDF0F68_4 As Led
    Friend WithEvents LabelFunctions_0 As System.Windows.Forms.Label
    Friend WithEvents LEDF0F68_5 As Led
    Friend WithEvents LEDF0F68_6 As Led
    Friend WithEvents LEDF0F68_7 As Led
    Friend WithEvents LEDF0F68_8 As Led
    Friend WithEvents LEDF0F68_9 As Led
    Friend WithEvents LEDF0F68_10 As Led
    Friend WithEvents LEDF0F68_11 As Led
    Friend WithEvents LEDF0F68_12 As Led
    Friend WithEvents LEDF0F68_13 As Led
    Friend WithEvents LEDF0F68_14 As Led
    Friend WithEvents LEDF0F68_15 As Led
    Friend WithEvents LEDF0F68_16 As Led
    Friend WithEvents LEDF0F68_17 As Led
    Friend WithEvents LEDF0F68_18 As Led
    Friend WithEvents LEDF0F68_19 As Led
    Friend WithEvents LEDF0F68_20 As Led
    Friend WithEvents LEDF0F68_21 As Led
    Friend WithEvents LEDF0F68_22 As Led
    Friend WithEvents LEDF0F68_23 As Led
    Friend WithEvents LEDF0F68_24 As Led
    Friend WithEvents LEDF0F68_25 As Led
    Friend WithEvents LEDF0F68_26 As Led
    Friend WithEvents LEDF0F68_27 As Led
    Friend WithEvents LEDF0F68_28 As Led
    Friend WithEvents LabelFunctions_1 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_2 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_3 As System.Windows.Forms.Label
    Friend WithEvents FrameVersion As System.Windows.Forms.GroupBox
    Friend WithEvents labeFirmwareVersion As System.Windows.Forms.Label
    Friend WithEvents txtXZ21Firmware As System.Windows.Forms.Label
    Friend WithEvents LabelMasterVersion As System.Windows.Forms.Label
    Friend WithEvents txtMasterVersion As System.Windows.Forms.Label
    Friend WithEvents FrameAnAus As System.Windows.Forms.GroupBox
    Friend WithEvents CmdAllesAn As System.Windows.Forms.Button
    Friend WithEvents CmdAllesAus As System.Windows.Forms.Button
    Friend WithEvents TimerMessageCounter As System.Windows.Forms.Timer
    Friend WithEvents FrameNumberOfmessages As System.Windows.Forms.GroupBox
    Friend WithEvents LabelZ21tMessageCounter As System.Windows.Forms.Label
    Friend WithEvents txtNmbPckts As System.Windows.Forms.Label
    Friend WithEvents FrameExtAcc As System.Windows.Forms.GroupBox
    Friend WithEvents TxtExtAccSwitch As System.Windows.Forms.TextBox
    Friend WithEvents TxtExtAccValue As System.Windows.Forms.TextBox
    Friend WithEvents CmdExtAccSet As System.Windows.Forms.Button
    Friend WithEvents LabelExtSwitch As System.Windows.Forms.Label
    Friend WithEvents LabelExtValue As System.Windows.Forms.Label
    Friend WithEvents FrameAcc As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRecht As System.Windows.Forms.Button
    Friend WithEvents cmdAf As System.Windows.Forms.Button
    Friend WithEvents txtWissel As System.Windows.Forms.TextBox
    Friend WithEvents cmdSwitchFeedback As System.Windows.Forms.Button
    Friend WithEvents TxtBoxSwitchAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchData As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchTtBits As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchIBit As System.Windows.Forms.TextBox
    Friend WithEvents CheckActivate As System.Windows.Forms.CheckBox
    Friend WithEvents CheckAutoDeactivate As System.Windows.Forms.CheckBox
    Friend WithEvents LabeLabelInfo1 As System.Windows.Forms.Label
    Friend WithEvents LabelSwitch As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchAddress As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchdata As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchTtBits As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchIBit As System.Windows.Forms.Label
    Friend WithEvents FrameStatus As System.Windows.Forms.GroupBox
    Friend WithEvents ledNotHalt As Led
    Friend WithEvents ledNotAus As Led
    Friend WithEvents ledProgrammiermode As Led
    Friend WithEvents LabelNotHalt As System.Windows.Forms.Label
    Friend WithEvents LabelNotAus As System.Windows.Forms.Label
    Friend WithEvents LabelProgrammiermode As System.Windows.Forms.Label
    Friend WithEvents FrameUDP As System.Windows.Forms.GroupBox
    Friend WithEvents cmdUdpSet As System.Windows.Forms.Button
    Friend WithEvents txtIP As System.Windows.Forms.TextBox
    Friend WithEvents txtUdpPort As System.Windows.Forms.TextBox
    Friend WithEvents LabelIpAddress As System.Windows.Forms.Label
    Friend WithEvents LabelUdpPort As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents LabelStatus As System.Windows.Forms.Label
End Class
