<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmExtendedTest
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
        Me.CheckHandshake = New System.Windows.Forms.CheckBox()
        Me.ComboBaudUsb = New System.Windows.Forms.ComboBox()
        Me.CmdExtAccSet = New System.Windows.Forms.Button()
        Me.TxtExtAccValue = New System.Windows.Forms.TextBox()
        Me.TxtExtAccSwitch = New System.Windows.Forms.TextBox()
        Me.CmdLokAusStack = New System.Windows.Forms.Button()
        Me.CmdStartMan = New System.Windows.Forms.Button()
        Me.SmPage = New System.Windows.Forms.RadioButton()
        Me.SmRegister = New System.Windows.Forms.RadioButton()
        Me.SmDirect = New System.Windows.Forms.RadioButton()
        Me.CheckSM = New System.Windows.Forms.CheckBox()
        Me.CmdLokEnter = New System.Windows.Forms.Button()
        Me.CmdGetF0_F12 = New System.Windows.Forms.Button()
        Me.CmdFunctionOff = New System.Windows.Forms.Button()
        Me.CmdFunctionOn = New System.Windows.Forms.Button()
        Me.CmdGetF29_68 = New System.Windows.Forms.Button()
        Me.CmdGetF13_28 = New System.Windows.Forms.Button()
        Me.CmdPomSet = New System.Windows.Forms.Button()
        Me.CmdPomGet = New System.Windows.Forms.Button()
        Me.TxtPomValue = New System.Windows.Forms.TextBox()
        Me.TxtPomCV = New System.Windows.Forms.TextBox()
        Me.TxtPomAddress = New System.Windows.Forms.TextBox()
        Me.TxtLokFNumber = New System.Windows.Forms.TextBox()
        Me.CheckAutoDeactivate = New System.Windows.Forms.CheckBox()
        Me.CheckActivate = New System.Windows.Forms.CheckBox()
        Me.CmdLokForward = New System.Windows.Forms.Button()
        Me.CmdLokReverse = New System.Windows.Forms.Button()
        Me.ComboLokSteps = New System.Windows.Forms.ComboBox()
        Me.TxtLokSpeed = New System.Windows.Forms.TextBox()
        Me.ComboDay = New System.Windows.Forms.ComboBox()
        Me.TxtHour = New System.Windows.Forms.TextBox()
        Me.TxtMinute = New System.Windows.Forms.TextBox()
        Me.TxtFactor = New System.Windows.Forms.TextBox()
        Me.CheckRun = New System.Windows.Forms.CheckBox()
        Me.CmdModellzeitStellen = New System.Windows.Forms.Button()
        Me.CmdModellzeitAnfordern = New System.Windows.Forms.Button()
        Me.CmdModellzeitAnhalten = New System.Windows.Forms.Button()
        Me.CmdModellzeitStarten = New System.Windows.Forms.Button()
        Me.cmdStartAuto = New System.Windows.Forms.Button()
        Me.CmdErweiterteVersion = New System.Windows.Forms.Button()
        Me.CmdLanInfo = New System.Windows.Forms.Button()
        Me.TimerMessageCounter = New System.Windows.Forms.Timer(Me.components)
        Me.CmdRsFeedback = New System.Windows.Forms.Button()
        Me.TxtRsBusAddress = New System.Windows.Forms.TextBox()
        Me.TxtBoxRsNibble = New System.Windows.Forms.TextBox()
        Me.TxtBoxRsData = New System.Windows.Forms.TextBox()
        Me.TxBoxtRsAddress = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchIBit = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchTtBits = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchData = New System.Windows.Forms.TextBox()
        Me.TxtBoxSwitchAddress = New System.Windows.Forms.TextBox()
        Me.CmdSerialConnect = New System.Windows.Forms.Button()
        Me.CheckCts = New System.Windows.Forms.CheckBox()
        Me.TxtSerialPort = New System.Windows.Forms.TextBox()
        Me.ComboBaud = New System.Windows.Forms.ComboBox()
        Me.Cmdtest = New System.Windows.Forms.Button()
        Me.CmdUsbConnect = New System.Windows.Forms.Button()
        Me.TxtUsbPort = New System.Windows.Forms.TextBox()
        Me.CmdInterfaceInfo = New System.Windows.Forms.Button()
        Me.cmdPomHolen = New System.Windows.Forms.Button()
        Me.txtLokAddress = New System.Windows.Forms.TextBox()
        Me.CmdEineLokAnhalten = New System.Windows.Forms.Button()
        Me.CmdAlleLoksAnhalten = New System.Windows.Forms.Button()
        Me.CmdAllesAus = New System.Windows.Forms.Button()
        Me.CmdAllesAn = New System.Windows.Forms.Button()
        Me.cmdSwitchFeedback = New System.Windows.Forms.Button()
        Me.txtWissel = New System.Windows.Forms.TextBox()
        Me.cmdAf = New System.Windows.Forms.Button()
        Me.cmdRecht = New System.Windows.Forms.Button()
        Me.txtTcpPort = New System.Windows.Forms.TextBox()
        Me.txtIP = New System.Windows.Forms.TextBox()
        Me.cmdTcpConnect = New System.Windows.Forms.Button()
        Me.LabelBaudrateUSB_1 = New System.Windows.Forms.Label()
        Me.Line9 = New System.Windows.Forms.Label()
        Me.LabelExtValue = New System.Windows.Forms.Label()
        Me.LabelExtSwitch = New System.Windows.Forms.Label()
        Me.LabelExtAccCmd = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Line8 = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_T = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_B = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_L = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_R = New System.Windows.Forms.Label()
        Me.LEDF0F68_68 = New Led()
        Me.LEDF0F68_67 = New Led()
        Me.LEDF0F68_66 = New Led()
        Me.LEDF0F68_65 = New Led()
        Me.LEDF0F68_64 = New Led()
        Me.LEDF0F68_63 = New Led()
        Me.LEDF0F68_62 = New Led()
        Me.LEDF0F68_61 = New Led()
        Me.LabelFunctions_8 = New System.Windows.Forms.Label()
        Me.LabelFunctions_7 = New System.Windows.Forms.Label()
        Me.LabelFunctions_6 = New System.Windows.Forms.Label()
        Me.LabelFunctions_5 = New System.Windows.Forms.Label()
        Me.LabelFunctions_4 = New System.Windows.Forms.Label()
        Me.LabelFunctions_3 = New System.Windows.Forms.Label()
        Me.LabelFunctions_2 = New System.Windows.Forms.Label()
        Me.LabelFunctions_1 = New System.Windows.Forms.Label()
        Me.LEDF0F68_60 = New Led()
        Me.LEDF0F68_59 = New Led()
        Me.LEDF0F68_58 = New Led()
        Me.LEDF0F68_57 = New Led()
        Me.LEDF0F68_56 = New Led()
        Me.LEDF0F68_55 = New Led()
        Me.LEDF0F68_54 = New Led()
        Me.LEDF0F68_53 = New Led()
        Me.LEDF0F68_52 = New Led()
        Me.LEDF0F68_51 = New Led()
        Me.LEDF0F68_50 = New Led()
        Me.LEDF0F68_49 = New Led()
        Me.LEDF0F68_48 = New Led()
        Me.LEDF0F68_47 = New Led()
        Me.LEDF0F68_46 = New Led()
        Me.LEDF0F68_45 = New Led()
        Me.LEDF0F68_44 = New Led()
        Me.LEDF0F68_43 = New Led()
        Me.LEDF0F68_42 = New Led()
        Me.LEDF0F68_41 = New Led()
        Me.LEDF0F68_40 = New Led()
        Me.LEDF0F68_39 = New Led()
        Me.LEDF0F68_38 = New Led()
        Me.LEDF0F68_37 = New Led()
        Me.LEDF0F68_36 = New Led()
        Me.LEDF0F68_35 = New Led()
        Me.LEDF0F68_34 = New Led()
        Me.LEDF0F68_33 = New Led()
        Me.LEDF0F68_32 = New Led()
        Me.LEDF0F68_31 = New Led()
        Me.LEDF0F68_30 = New Led()
        Me.LEDF0F68_29 = New Led()
        Me.LEDF0F68_28 = New Led()
        Me.LEDF0F68_27 = New Led()
        Me.LEDF0F68_26 = New Led()
        Me.LEDF0F68_25 = New Led()
        Me.LEDF0F68_24 = New Led()
        Me.LEDF0F68_23 = New Led()
        Me.LEDF0F68_22 = New Led()
        Me.LEDF0F68_21 = New Led()
        Me.LEDF0F68_20 = New Led()
        Me.LEDF0F68_19 = New Led()
        Me.LEDF0F68_18 = New Led()
        Me.LEDF0F68_17 = New Led()
        Me.LEDF0F68_16 = New Led()
        Me.LEDF0F68_15 = New Led()
        Me.LEDF0F68_14 = New Led()
        Me.LEDF0F68_13 = New Led()
        Me.LEDF0F68_12 = New Led()
        Me.LEDF0F68_11 = New Led()
        Me.LEDF0F68_10 = New Led()
        Me.LEDF0F68_9 = New Led()
        Me.LEDF0F68_8 = New Led()
        Me.LEDF0F68_7 = New Led()
        Me.LEDF0F68_6 = New Led()
        Me.LEDF0F68_5 = New Led()
        Me.LabelFunctions_0 = New System.Windows.Forms.Label()
        Me.LEDF0F68_4 = New Led()
        Me.LEDF0F68_3 = New Led()
        Me.LEDF0F68_2 = New Led()
        Me.LEDF0F68_1 = New Led()
        Me.LEDF0F68_0 = New Led()
        Me.Line_Shape4_T = New System.Windows.Forms.Label()
        Me.Line_Shape4_B = New System.Windows.Forms.Label()
        Me.Line_Shape4_L = New System.Windows.Forms.Label()
        Me.Line_Shape4_R = New System.Windows.Forms.Label()
        Me.LabelPom = New System.Windows.Forms.Label()
        Me.LabelPomValue = New System.Windows.Forms.Label()
        Me.LabelPomCV = New System.Windows.Forms.Label()
        Me.LabelPomAddress = New System.Windows.Forms.Label()
        Me.LabelFunktion = New System.Windows.Forms.Label()
        Me.LabelLoks = New System.Windows.Forms.Label()
        Me.LabelLokSteps = New System.Windows.Forms.Label()
        Me.LabelLokSpeed = New System.Windows.Forms.Label()
        Me.LabelLokAddress = New System.Windows.Forms.Label()
        Me.LabelModellzeit = New System.Windows.Forms.Label()
        Me.Line_Shape3_T = New System.Windows.Forms.Label()
        Me.Line_Shape3_B = New System.Windows.Forms.Label()
        Me.Line_Shape3_L = New System.Windows.Forms.Label()
        Me.Line_Shape3_R = New System.Windows.Forms.Label()
        Me.Line7 = New System.Windows.Forms.Label()
        Me.Line_Shape2_T = New System.Windows.Forms.Label()
        Me.Line_Shape2_B = New System.Windows.Forms.Label()
        Me.Line_Shape2_L = New System.Windows.Forms.Label()
        Me.Line_Shape2_R = New System.Windows.Forms.Label()
        Me.LabelDay = New System.Windows.Forms.Label()
        Me.LabelHour = New System.Windows.Forms.Label()
        Me.LabelMinute = New System.Windows.Forms.Label()
        Me.LabelFacor = New System.Windows.Forms.Label()
        Me.TxtMinutes = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TxtHours = New System.Windows.Forms.Label()
        Me.txtVersionBootloader = New System.Windows.Forms.Label()
        Me.txtBuildRuckmelder = New System.Windows.Forms.Label()
        Me.TxtVersionRuckmelder = New System.Windows.Forms.Label()
        Me.txtBuildNr = New System.Windows.Forms.Label()
        Me.LabelVersionBootloader = New System.Windows.Forms.Label()
        Me.LabelBuildruckmeldedecoder = New System.Windows.Forms.Label()
        Me.LabelVersionRuckmeldedecoder = New System.Windows.Forms.Label()
        Me.LabelBuildNr = New System.Windows.Forms.Label()
        Me.Line6 = New System.Windows.Forms.Label()
        Me.Line_Shape1_T = New System.Windows.Forms.Label()
        Me.Line_Shape1_B = New System.Windows.Forms.Label()
        Me.Line_Shape1_L = New System.Windows.Forms.Label()
        Me.Line_Shape1_R = New System.Windows.Forms.Label()
        Me.LabelInfo2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LabelRsBusAddress = New System.Windows.Forms.Label()
        Me.Line_ShapeFeedback_T = New System.Windows.Forms.Label()
        Me.Line_ShapeFeedback_B = New System.Windows.Forms.Label()
        Me.Line_ShapeFeedback_L = New System.Windows.Forms.Label()
        Me.Line_ShapeFeedback_R = New System.Windows.Forms.Label()
        Me.LabelSwitchIBit = New System.Windows.Forms.Label()
        Me.LabelSwitchTtBits = New System.Windows.Forms.Label()
        Me.LabelSwitchdata = New System.Windows.Forms.Label()
        Me.LabelSwitchAddress = New System.Windows.Forms.Label()
        Me.Line_ShapeSwitch_T = New System.Windows.Forms.Label()
        Me.Line_ShapeSwitch_B = New System.Windows.Forms.Label()
        Me.Line_ShapeSwitch_L = New System.Windows.Forms.Label()
        Me.Line_ShapeSwitch_R = New System.Windows.Forms.Label()
        Me.Line_ShapeInfo_T = New System.Windows.Forms.Label()
        Me.Line_ShapeInfo_B = New System.Windows.Forms.Label()
        Me.Line_ShapeInfo_L = New System.Windows.Forms.Label()
        Me.Line_ShapeInfo_R = New System.Windows.Forms.Label()
        Me.LabelConnected = New System.Windows.Forms.Label()
        Me.Line_ShapeSerial_T = New System.Windows.Forms.Label()
        Me.Line_ShapeSerial_B = New System.Windows.Forms.Label()
        Me.Line_ShapeSerial_L = New System.Windows.Forms.Label()
        Me.Line_ShapeSerial_R = New System.Windows.Forms.Label()
        Me.Line_ShapeUsb_T = New System.Windows.Forms.Label()
        Me.Line_ShapeUsb_B = New System.Windows.Forms.Label()
        Me.Line_ShapeUsb_L = New System.Windows.Forms.Label()
        Me.Line_ShapeUsb_R = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_T = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_B = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_L = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_R = New System.Windows.Forms.Label()
        Me.LabelBaudrate_0 = New System.Windows.Forms.Label()
        Me.LabelSerialPort = New System.Windows.Forms.Label()
        Me.LabelUSBPort = New System.Windows.Forms.Label()
        Me.LabelProgrammiermode = New System.Windows.Forms.Label()
        Me.LabelNotAus = New System.Windows.Forms.Label()
        Me.LabelNotHalt = New System.Windows.Forms.Label()
        Me.ledProgrammiermode = New Led()
        Me.ledNotAus = New Led()
        Me.ledNotHalt = New Led()
        Me.txtMasterVersion = New System.Windows.Forms.Label()
        Me.LabelMasterVersion = New System.Windows.Forms.Label()
        Me.Line5 = New System.Windows.Forms.Label()
        Me.Line4 = New System.Windows.Forms.Label()
        Me.txtNmbPckts = New System.Windows.Forms.Label()
        Me.txtTcpFree = New System.Windows.Forms.Label()
        Me.txtXVersion = New System.Windows.Forms.Label()
        Me.txtXAddress = New System.Windows.Forms.Label()
        Me.txtInterfaceCode = New System.Windows.Forms.Label()
        Me.txtInterfaceVersion = New System.Windows.Forms.Label()
        Me.Line3 = New System.Windows.Forms.Label()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.labelTcpFree = New System.Windows.Forms.Label()
        Me.labelXVersion = New System.Windows.Forms.Label()
        Me.labelXAddress = New System.Windows.Forms.Label()
        Me.labelCode = New System.Windows.Forms.Label()
        Me.labelInterfaceVersion = New System.Windows.Forms.Label()
        Me.LabelXpressNetMessageCounter = New System.Windows.Forms.Label()
        Me.LabelSwitch = New System.Windows.Forms.Label()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.LabelStatus = New System.Windows.Forms.Label()
        Me.ledInterfaceStatus = New Led()
        Me.LabelTcpPort = New System.Windows.Forms.Label()
        Me.LabelIpAddress = New System.Windows.Forms.Label()
        Me.LabeLabelInfo1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'CheckHandshake
        '
        Me.CheckHandshake.Location = New System.Drawing.Point(230, 100)
        Me.CheckHandshake.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckHandshake.Name = "CheckHandshake"
        Me.CheckHandshake.Size = New System.Drawing.Size(86, 20)
        Me.CheckHandshake.TabIndex = 144
        Me.CheckHandshake.Text = "Handshake"
        Me.CheckHandshake.UseVisualStyleBackColor = True
        '
        'ComboBaudUsb
        '
        Me.ComboBaudUsb.Location = New System.Drawing.Point(296, 64)
        Me.ComboBaudUsb.Name = "ComboBaudUsb"
        Me.ComboBaudUsb.Size = New System.Drawing.Size(136, 24)
        Me.ComboBaudUsb.TabIndex = 142
        Me.ComboBaudUsb.Text = "57600"
        '
        'CmdExtAccSet
        '
        Me.CmdExtAccSet.Location = New System.Drawing.Point(916, 480)
        Me.CmdExtAccSet.Name = "CmdExtAccSet"
        Me.CmdExtAccSet.Size = New System.Drawing.Size(64, 25)
        Me.CmdExtAccSet.TabIndex = 141
        Me.CmdExtAccSet.Text = "Set"
        Me.CmdExtAccSet.UseVisualStyleBackColor = True
        '
        'TxtExtAccValue
        '
        Me.TxtExtAccValue.Location = New System.Drawing.Point(916, 448)
        Me.TxtExtAccValue.Name = "TxtExtAccValue"
        Me.TxtExtAccValue.Size = New System.Drawing.Size(64, 22)
        Me.TxtExtAccValue.TabIndex = 140
        Me.TxtExtAccValue.Text = ""
        '
        'TxtExtAccSwitch
        '
        Me.TxtExtAccSwitch.Location = New System.Drawing.Point(916, 400)
        Me.TxtExtAccSwitch.Name = "TxtExtAccSwitch"
        Me.TxtExtAccSwitch.Size = New System.Drawing.Size(64, 22)
        Me.TxtExtAccSwitch.TabIndex = 138
        Me.TxtExtAccSwitch.Text = ""
        '
        'CmdLokAusStack
        '
        Me.CmdLokAusStack.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdLokAusStack.Location = New System.Drawing.Point(694, 384)
        Me.CmdLokAusStack.Name = "CmdLokAusStack"
        Me.CmdLokAusStack.Size = New System.Drawing.Size(20, 22)
        Me.CmdLokAusStack.TabIndex = 135
        Me.CmdLokAusStack.Text = "-"
        Me.CmdLokAusStack.UseVisualStyleBackColor = True
        '
        'CmdStartMan
        '
        Me.CmdStartMan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdStartMan.Location = New System.Drawing.Point(808, 200)
        Me.CmdStartMan.Name = "CmdStartMan"
        Me.CmdStartMan.Size = New System.Drawing.Size(46, 25)
        Me.CmdStartMan.TabIndex = 134
        Me.CmdStartMan.Text = "Man."
        Me.CmdStartMan.UseVisualStyleBackColor = True
        '
        'SmPage
        '
        Me.SmPage.Location = New System.Drawing.Point(1074, 430)
        Me.SmPage.Name = "SmPage"
        Me.SmPage.Size = New System.Drawing.Size(92, 22)
        Me.SmPage.TabIndex = 132
        Me.SmPage.Text = "Page"
        Me.SmPage.UseVisualStyleBackColor = True
        '
        'SmRegister
        '
        Me.SmRegister.Location = New System.Drawing.Point(1074, 406)
        Me.SmRegister.Name = "SmRegister"
        Me.SmRegister.Size = New System.Drawing.Size(92, 22)
        Me.SmRegister.TabIndex = 131
        Me.SmRegister.Text = "Register"
        Me.SmRegister.UseVisualStyleBackColor = True
        '
        'SmDirect
        '
        Me.SmDirect.Location = New System.Drawing.Point(1074, 382)
        Me.SmDirect.Name = "SmDirect"
        Me.SmDirect.Size = New System.Drawing.Size(92, 22)
        Me.SmDirect.TabIndex = 130
        Me.SmDirect.Text = "Direct"
        Me.SmDirect.Checked = True
        Me.SmDirect.UseVisualStyleBackColor = True
        '
        'CheckSM
        '
        Me.CheckSM.Location = New System.Drawing.Point(1014, 400)
        Me.CheckSM.Name = "CheckSM"
        Me.CheckSM.Size = New System.Drawing.Size(50, 25)
        Me.CheckSM.TabIndex = 129
        Me.CheckSM.Text = "SM"
        Me.CheckSM.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckSM.UseVisualStyleBackColor = True
        '
        'CmdLokEnter
        '
        Me.CmdLokEnter.Location = New System.Drawing.Point(640, 416)
        Me.CmdLokEnter.Name = "CmdLokEnter"
        Me.CmdLokEnter.Size = New System.Drawing.Size(49, 25)
        Me.CmdLokEnter.TabIndex = 128
        Me.CmdLokEnter.Text = "Enter"
        Me.CmdLokEnter.UseVisualStyleBackColor = True
        '
        'CmdGetF0_F12
        '
        Me.CmdGetF0_F12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdGetF0_F12.Location = New System.Drawing.Point(846, 408)
        Me.CmdGetF0_F12.Name = "CmdGetF0_F12"
        Me.CmdGetF0_F12.Size = New System.Drawing.Size(38, 25)
        Me.CmdGetF0_F12.TabIndex = 127
        Me.CmdGetF0_F12.Text = "Get"
        Me.CmdGetF0_F12.UseVisualStyleBackColor = True
        '
        'CmdFunctionOff
        '
        Me.CmdFunctionOff.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdFunctionOff.Location = New System.Drawing.Point(816, 384)
        Me.CmdFunctionOff.Name = "CmdFunctionOff"
        Me.CmdFunctionOff.Size = New System.Drawing.Size(25, 25)
        Me.CmdFunctionOff.TabIndex = 126
        Me.CmdFunctionOff.Text = "-"
        Me.CmdFunctionOff.UseVisualStyleBackColor = True
        '
        'CmdFunctionOn
        '
        Me.CmdFunctionOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdFunctionOn.Location = New System.Drawing.Point(776, 384)
        Me.CmdFunctionOn.Name = "CmdFunctionOn"
        Me.CmdFunctionOn.Size = New System.Drawing.Size(25, 25)
        Me.CmdFunctionOn.TabIndex = 125
        Me.CmdFunctionOn.Text = "+"
        Me.CmdFunctionOn.UseVisualStyleBackColor = True
        '
        'CmdGetF29_68
        '
        Me.CmdGetF29_68.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdGetF29_68.Location = New System.Drawing.Point(846, 504)
        Me.CmdGetF29_68.Name = "CmdGetF29_68"
        Me.CmdGetF29_68.Size = New System.Drawing.Size(38, 25)
        Me.CmdGetF29_68.TabIndex = 124
        Me.CmdGetF29_68.Text = "Get"
        Me.CmdGetF29_68.UseVisualStyleBackColor = True
        '
        'CmdGetF13_28
        '
        Me.CmdGetF13_28.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdGetF13_28.Location = New System.Drawing.Point(846, 448)
        Me.CmdGetF13_28.Name = "CmdGetF13_28"
        Me.CmdGetF13_28.Size = New System.Drawing.Size(38, 25)
        Me.CmdGetF13_28.TabIndex = 123
        Me.CmdGetF13_28.Text = "Get"
        Me.CmdGetF13_28.UseVisualStyleBackColor = True
        '
        'CmdPomSet
        '
        Me.CmdPomSet.Location = New System.Drawing.Point(1096, 560)
        Me.CmdPomSet.Name = "CmdPomSet"
        Me.CmdPomSet.Size = New System.Drawing.Size(64, 25)
        Me.CmdPomSet.TabIndex = 113
        Me.CmdPomSet.Text = "Set"
        Me.CmdPomSet.UseVisualStyleBackColor = True
        '
        'CmdPomGet
        '
        Me.CmdPomGet.Location = New System.Drawing.Point(1020, 560)
        Me.CmdPomGet.Name = "CmdPomGet"
        Me.CmdPomGet.Size = New System.Drawing.Size(64, 25)
        Me.CmdPomGet.TabIndex = 111
        Me.CmdPomGet.Text = "Get"
        Me.CmdPomGet.UseVisualStyleBackColor = True
        '
        'TxtPomValue
        '
        Me.TxtPomValue.Location = New System.Drawing.Point(1084, 528)
        Me.TxtPomValue.Name = "TxtPomValue"
        Me.TxtPomValue.Size = New System.Drawing.Size(80, 22)
        Me.TxtPomValue.TabIndex = 107
        Me.TxtPomValue.Text = "0"
        '
        'TxtPomCV
        '
        Me.TxtPomCV.Location = New System.Drawing.Point(1084, 496)
        Me.TxtPomCV.Name = "TxtPomCV"
        Me.TxtPomCV.Size = New System.Drawing.Size(80, 22)
        Me.TxtPomCV.TabIndex = 106
        Me.TxtPomCV.Text = "4"
        '
        'TxtPomAddress
        '
        Me.TxtPomAddress.Location = New System.Drawing.Point(1084, 464)
        Me.TxtPomAddress.Name = "TxtPomAddress"
        Me.TxtPomAddress.Size = New System.Drawing.Size(80, 22)
        Me.TxtPomAddress.TabIndex = 105
        Me.TxtPomAddress.Text = "3"
        '
        'TxtLokFNumber
        '
        Me.TxtLokFNumber.Location = New System.Drawing.Point(744, 384)
        Me.TxtLokFNumber.Name = "TxtLokFNumber"
        Me.TxtLokFNumber.Size = New System.Drawing.Size(25, 22)
        Me.TxtLokFNumber.TabIndex = 104
        Me.TxtLokFNumber.Text = "1"
        '
        'CheckAutoDeactivate
        '
        Me.CheckAutoDeactivate.Location = New System.Drawing.Point(150, 264)
        Me.CheckAutoDeactivate.Name = "CheckAutoDeactivate"
        Me.CheckAutoDeactivate.Size = New System.Drawing.Size(128, 25)
        Me.CheckAutoDeactivate.TabIndex = 102
        Me.CheckAutoDeactivate.Text = "auto deactivate"
        Me.CheckAutoDeactivate.Checked = True
        Me.CheckAutoDeactivate.UseVisualStyleBackColor = True
        '
        'CheckActivate
        '
        Me.CheckActivate.Location = New System.Drawing.Point(150, 240)
        Me.CheckActivate.Name = "CheckActivate"
        Me.CheckActivate.Size = New System.Drawing.Size(128, 25)
        Me.CheckActivate.TabIndex = 101
        Me.CheckActivate.Text = "activate"
        Me.CheckActivate.Checked = True
        Me.CheckActivate.UseVisualStyleBackColor = True
        '
        'CmdLokForward
        '
        Me.CmdLokForward.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdLokForward.Location = New System.Drawing.Point(643, 496)
        Me.CmdLokForward.Name = "CmdLokForward"
        Me.CmdLokForward.Size = New System.Drawing.Size(46, 25)
        Me.CmdLokForward.TabIndex = 100
        Me.CmdLokForward.Text = "=>"
        Me.CmdLokForward.UseVisualStyleBackColor = True
        '
        'CmdLokReverse
        '
        Me.CmdLokReverse.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdLokReverse.Location = New System.Drawing.Point(588, 496)
        Me.CmdLokReverse.Name = "CmdLokReverse"
        Me.CmdLokReverse.Size = New System.Drawing.Size(46, 25)
        Me.CmdLokReverse.TabIndex = 99
        Me.CmdLokReverse.Text = "<="
        Me.CmdLokReverse.UseVisualStyleBackColor = True
        '
        'ComboLokSteps
        '
        Me.ComboLokSteps.Location = New System.Drawing.Point(640, 568)
        Me.ComboLokSteps.Name = "ComboLokSteps"
        Me.ComboLokSteps.Size = New System.Drawing.Size(49, 24)
        Me.ComboLokSteps.TabIndex = 98
        Me.ComboLokSteps.Text = "28"
        '
        'TxtLokSpeed
        '
        Me.TxtLokSpeed.Location = New System.Drawing.Point(640, 464)
        Me.TxtLokSpeed.Name = "TxtLokSpeed"
        Me.TxtLokSpeed.Size = New System.Drawing.Size(49, 22)
        Me.TxtLokSpeed.TabIndex = 97
        Me.TxtLokSpeed.Text = "0"
        '
        'ComboDay
        '
        Me.ComboDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboDay.Location = New System.Drawing.Point(618, 216)
        Me.ComboDay.Name = "ComboDay"
        Me.ComboDay.Size = New System.Drawing.Size(88, 21)
        Me.ComboDay.TabIndex = 87
        Me.ComboDay.Text = "Monday"
        '
        'TxtHour
        '
        Me.TxtHour.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtHour.Location = New System.Drawing.Point(618, 240)
        Me.TxtHour.Name = "TxtHour"
        Me.TxtHour.Size = New System.Drawing.Size(25, 20)
        Me.TxtHour.TabIndex = 86
        Me.TxtHour.Text = "0"
        '
        'TxtMinute
        '
        Me.TxtMinute.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtMinute.Location = New System.Drawing.Point(618, 272)
        Me.TxtMinute.Name = "TxtMinute"
        Me.TxtMinute.Size = New System.Drawing.Size(25, 20)
        Me.TxtMinute.TabIndex = 85
        Me.TxtMinute.Text = "0"
        '
        'TxtFactor
        '
        Me.TxtFactor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtFactor.Location = New System.Drawing.Point(618, 304)
        Me.TxtFactor.Name = "TxtFactor"
        Me.TxtFactor.Size = New System.Drawing.Size(25, 20)
        Me.TxtFactor.TabIndex = 84
        Me.TxtFactor.Text = "0"
        '
        'CheckRun
        '
        Me.CheckRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckRun.Location = New System.Drawing.Point(650, 272)
        Me.CheckRun.Name = "CheckRun"
        Me.CheckRun.Size = New System.Drawing.Size(54, 25)
        Me.CheckRun.TabIndex = 83
        Me.CheckRun.Text = "Run"
        Me.CheckRun.Checked = True
        Me.CheckRun.UseVisualStyleBackColor = True
        '
        'CmdModellzeitStellen
        '
        Me.CmdModellzeitStellen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdModellzeitStellen.Location = New System.Drawing.Point(648, 304)
        Me.CmdModellzeitStellen.Name = "CmdModellzeitStellen"
        Me.CmdModellzeitStellen.Size = New System.Drawing.Size(58, 25)
        Me.CmdModellzeitStellen.TabIndex = 82
        Me.CmdModellzeitStellen.Text = "Stellen"
        Me.CmdModellzeitStellen.UseVisualStyleBackColor = True
        '
        'CmdModellzeitAnfordern
        '
        Me.CmdModellzeitAnfordern.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdModellzeitAnfordern.Location = New System.Drawing.Point(720, 304)
        Me.CmdModellzeitAnfordern.Name = "CmdModellzeitAnfordern"
        Me.CmdModellzeitAnfordern.Size = New System.Drawing.Size(72, 25)
        Me.CmdModellzeitAnfordern.TabIndex = 81
        Me.CmdModellzeitAnfordern.Text = "Anfordern"
        Me.CmdModellzeitAnfordern.UseVisualStyleBackColor = True
        '
        'CmdModellzeitAnhalten
        '
        Me.CmdModellzeitAnhalten.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdModellzeitAnhalten.Location = New System.Drawing.Point(720, 264)
        Me.CmdModellzeitAnhalten.Name = "CmdModellzeitAnhalten"
        Me.CmdModellzeitAnhalten.Size = New System.Drawing.Size(72, 25)
        Me.CmdModellzeitAnhalten.TabIndex = 80
        Me.CmdModellzeitAnhalten.Text = "Anhalten"
        Me.CmdModellzeitAnhalten.UseVisualStyleBackColor = True
        '
        'CmdModellzeitStarten
        '
        Me.CmdModellzeitStarten.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdModellzeitStarten.Location = New System.Drawing.Point(720, 224)
        Me.CmdModellzeitStarten.Name = "CmdModellzeitStarten"
        Me.CmdModellzeitStarten.Size = New System.Drawing.Size(72, 25)
        Me.CmdModellzeitStarten.TabIndex = 79
        Me.CmdModellzeitStarten.Text = "Starten"
        Me.CmdModellzeitStarten.UseVisualStyleBackColor = True
        '
        'cmdStartAuto
        '
        Me.cmdStartAuto.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStartAuto.Location = New System.Drawing.Point(856, 200)
        Me.cmdStartAuto.Name = "cmdStartAuto"
        Me.cmdStartAuto.Size = New System.Drawing.Size(46, 25)
        Me.cmdStartAuto.TabIndex = 78
        Me.cmdStartAuto.Text = "Auto"
        Me.cmdStartAuto.UseVisualStyleBackColor = True
        '
        'CmdErweiterteVersion
        '
        Me.CmdErweiterteVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdErweiterteVersion.Location = New System.Drawing.Point(1093, 202)
        Me.CmdErweiterteVersion.Name = "CmdErweiterteVersion"
        Me.CmdErweiterteVersion.Size = New System.Drawing.Size(67, 40)
        Me.CmdErweiterteVersion.TabIndex = 74
        Me.CmdErweiterteVersion.Text = "Erweiterte Info"
        Me.CmdErweiterteVersion.UseVisualStyleBackColor = True
        '
        'CmdLanInfo
        '
        Me.CmdLanInfo.Location = New System.Drawing.Point(1080, 114)
        Me.CmdLanInfo.Name = "CmdLanInfo"
        Me.CmdLanInfo.Size = New System.Drawing.Size(81, 40)
        Me.CmdLanInfo.TabIndex = 65
        Me.CmdLanInfo.Text = "LAN Info"
        Me.CmdLanInfo.UseVisualStyleBackColor = True
        '
        'TimerMessageCounter
        '
        Me.TimerMessageCounter.Interval = 1000
        Me.TimerMessageCounter.Enabled = True
        '
        'CmdRsFeedback
        '
        Me.CmdRsFeedback.Location = New System.Drawing.Point(424, 248)
        Me.CmdRsFeedback.Name = "CmdRsFeedback"
        Me.CmdRsFeedback.Size = New System.Drawing.Size(89, 33)
        Me.CmdRsFeedback.TabIndex = 63
        Me.CmdRsFeedback.Text = "Feedback"
        Me.CmdRsFeedback.UseVisualStyleBackColor = True
        '
        'TxtRsBusAddress
        '
        Me.TxtRsBusAddress.Location = New System.Drawing.Point(424, 208)
        Me.TxtRsBusAddress.Name = "TxtRsBusAddress"
        Me.TxtRsBusAddress.Size = New System.Drawing.Size(89, 22)
        Me.TxtRsBusAddress.TabIndex = 58
        Me.TxtRsBusAddress.Text = ""
        '
        'TxtBoxRsNibble
        '
        Me.TxtBoxRsNibble.Location = New System.Drawing.Point(480, 328)
        Me.TxtBoxRsNibble.Name = "TxtBoxRsNibble"
        Me.TxtBoxRsNibble.Size = New System.Drawing.Size(41, 265)
        Me.TxtBoxRsNibble.TabIndex = 57
        Me.TxtBoxRsNibble.Text = ""
        Me.TxtBoxRsNibble.Multiline = True
        '
        'TxtBoxRsData
        '
        Me.TxtBoxRsData.Location = New System.Drawing.Point(408, 328)
        Me.TxtBoxRsData.Name = "TxtBoxRsData"
        Me.TxtBoxRsData.Size = New System.Drawing.Size(73, 265)
        Me.TxtBoxRsData.TabIndex = 56
        Me.TxtBoxRsData.Text = ""
        Me.TxtBoxRsData.Multiline = True
        '
        'TxBoxtRsAddress
        '
        Me.TxBoxtRsAddress.Location = New System.Drawing.Point(312, 328)
        Me.TxBoxtRsAddress.Name = "TxBoxtRsAddress"
        Me.TxBoxtRsAddress.Size = New System.Drawing.Size(97, 265)
        Me.TxBoxtRsAddress.TabIndex = 55
        Me.TxBoxtRsAddress.Text = ""
        Me.TxBoxtRsAddress.Multiline = True
        '
        'TxtBoxSwitchIBit
        '
        Me.TxtBoxSwitchIBit.Location = New System.Drawing.Point(232, 328)
        Me.TxtBoxSwitchIBit.Name = "TxtBoxSwitchIBit"
        Me.TxtBoxSwitchIBit.Size = New System.Drawing.Size(25, 265)
        Me.TxtBoxSwitchIBit.TabIndex = 50
        Me.TxtBoxSwitchIBit.Text = ""
        Me.TxtBoxSwitchIBit.Multiline = True
        '
        'TxtBoxSwitchTtBits
        '
        Me.TxtBoxSwitchTtBits.Location = New System.Drawing.Point(192, 328)
        Me.TxtBoxSwitchTtBits.Name = "TxtBoxSwitchTtBits"
        Me.TxtBoxSwitchTtBits.Size = New System.Drawing.Size(41, 265)
        Me.TxtBoxSwitchTtBits.TabIndex = 49
        Me.TxtBoxSwitchTtBits.Text = ""
        Me.TxtBoxSwitchTtBits.Multiline = True
        '
        'TxtBoxSwitchData
        '
        Me.TxtBoxSwitchData.Location = New System.Drawing.Point(120, 328)
        Me.TxtBoxSwitchData.Name = "TxtBoxSwitchData"
        Me.TxtBoxSwitchData.Size = New System.Drawing.Size(73, 265)
        Me.TxtBoxSwitchData.TabIndex = 48
        Me.TxtBoxSwitchData.Text = ""
        Me.TxtBoxSwitchData.Multiline = True
        '
        'TxtBoxSwitchAddress
        '
        Me.TxtBoxSwitchAddress.Location = New System.Drawing.Point(24, 328)
        Me.TxtBoxSwitchAddress.Name = "TxtBoxSwitchAddress"
        Me.TxtBoxSwitchAddress.Size = New System.Drawing.Size(97, 265)
        Me.TxtBoxSwitchAddress.TabIndex = 47
        Me.TxtBoxSwitchAddress.Text = ""
        Me.TxtBoxSwitchAddress.Multiline = True
        '
        'CmdSerialConnect
        '
        Me.CmdSerialConnect.Location = New System.Drawing.Point(512, 96)
        Me.CmdSerialConnect.Name = "CmdSerialConnect"
        Me.CmdSerialConnect.Size = New System.Drawing.Size(124, 25)
        Me.CmdSerialConnect.TabIndex = 45
        Me.CmdSerialConnect.Text = "Connect Serial"
        Me.CmdSerialConnect.UseVisualStyleBackColor = True
        '
        'CheckCts
        '
        Me.CheckCts.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckCts.Location = New System.Drawing.Point(636, 56)
        Me.CheckCts.Name = "CheckCts"
        Me.CheckCts.Size = New System.Drawing.Size(97, 33)
        Me.CheckCts.TabIndex = 44
        Me.CheckCts.Text = "CTS flowcontrol"
        Me.CheckCts.Checked = True
        Me.CheckCts.UseVisualStyleBackColor = True
        '
        'TxtSerialPort
        '
        Me.TxtSerialPort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtSerialPort.Location = New System.Drawing.Point(532, 32)
        Me.TxtSerialPort.Name = "TxtSerialPort"
        Me.TxtSerialPort.Size = New System.Drawing.Size(81, 20)
        Me.TxtSerialPort.TabIndex = 41
        Me.TxtSerialPort.Text = "1"
        '
        'ComboBaud
        '
        Me.ComboBaud.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBaud.Location = New System.Drawing.Point(532, 64)
        Me.ComboBaud.Name = "ComboBaud"
        Me.ComboBaud.Size = New System.Drawing.Size(81, 21)
        Me.ComboBaud.TabIndex = 40
        Me.ComboBaud.Text = "19200"
        '
        'Cmdtest
        '
        Me.Cmdtest.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmdtest.Location = New System.Drawing.Point(814, 296)
        Me.Cmdtest.Name = "Cmdtest"
        Me.Cmdtest.Size = New System.Drawing.Size(80, 33)
        Me.Cmdtest.TabIndex = 39
        Me.Cmdtest.Text = "Send Test"
        Me.Cmdtest.UseVisualStyleBackColor = True
        '
        'CmdUsbConnect
        '
        Me.CmdUsbConnect.Location = New System.Drawing.Point(320, 96)
        Me.CmdUsbConnect.Name = "CmdUsbConnect"
        Me.CmdUsbConnect.Size = New System.Drawing.Size(112, 25)
        Me.CmdUsbConnect.TabIndex = 38
        Me.CmdUsbConnect.Text = "Connect USB"
        Me.CmdUsbConnect.UseVisualStyleBackColor = True
        '
        'TxtUsbPort
        '
        Me.TxtUsbPort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtUsbPort.Location = New System.Drawing.Point(296, 32)
        Me.TxtUsbPort.Name = "TxtUsbPort"
        Me.TxtUsbPort.Size = New System.Drawing.Size(136, 20)
        Me.TxtUsbPort.TabIndex = 37
        Me.TxtUsbPort.Text = "3"
        '
        'CmdInterfaceInfo
        '
        Me.CmdInterfaceInfo.Location = New System.Drawing.Point(1079, 47)
        Me.CmdInterfaceInfo.Name = "CmdInterfaceInfo"
        Me.CmdInterfaceInfo.Size = New System.Drawing.Size(81, 38)
        Me.CmdInterfaceInfo.TabIndex = 30
        Me.CmdInterfaceInfo.Text = "Interface Info"
        Me.CmdInterfaceInfo.UseVisualStyleBackColor = True
        '
        'cmdPomHolen
        '
        Me.cmdPomHolen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPomHolen.Location = New System.Drawing.Point(1020, 592)
        Me.cmdPomHolen.Name = "cmdPomHolen"
        Me.cmdPomHolen.Size = New System.Drawing.Size(140, 25)
        Me.cmdPomHolen.TabIndex = 29
        Me.cmdPomHolen.Text = "Holen"
        Me.cmdPomHolen.UseVisualStyleBackColor = True
        '
        'txtLokAddress
        '
        Me.txtLokAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLokAddress.Location = New System.Drawing.Point(640, 384)
        Me.txtLokAddress.Name = "txtLokAddress"
        Me.txtLokAddress.Size = New System.Drawing.Size(49, 20)
        Me.txtLokAddress.TabIndex = 28
        Me.txtLokAddress.Text = "3"
        '
        'CmdEineLokAnhalten
        '
        Me.CmdEineLokAnhalten.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdEineLokAnhalten.Location = New System.Drawing.Point(588, 528)
        Me.CmdEineLokAnhalten.Name = "CmdEineLokAnhalten"
        Me.CmdEineLokAnhalten.Size = New System.Drawing.Size(101, 25)
        Me.CmdEineLokAnhalten.TabIndex = 27
        Me.CmdEineLokAnhalten.Text = "Anhalten"
        Me.CmdEineLokAnhalten.UseVisualStyleBackColor = True
        '
        'CmdAlleLoksAnhalten
        '
        Me.CmdAlleLoksAnhalten.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdAlleLoksAnhalten.Location = New System.Drawing.Point(728, 568)
        Me.CmdAlleLoksAnhalten.Name = "CmdAlleLoksAnhalten"
        Me.CmdAlleLoksAnhalten.Size = New System.Drawing.Size(107, 33)
        Me.CmdAlleLoksAnhalten.TabIndex = 26
        Me.CmdAlleLoksAnhalten.Text = "Alle Loks anhalten"
        Me.CmdAlleLoksAnhalten.UseVisualStyleBackColor = True
        '
        'CmdAllesAus
        '
        Me.CmdAllesAus.Location = New System.Drawing.Point(1046, 288)
        Me.CmdAllesAus.Name = "CmdAllesAus"
        Me.CmdAllesAus.Size = New System.Drawing.Size(110, 33)
        Me.CmdAllesAus.TabIndex = 25
        Me.CmdAllesAus.Text = "Alles Aus"
        Me.CmdAllesAus.UseVisualStyleBackColor = True
        '
        'CmdAllesAn
        '
        Me.CmdAllesAn.Location = New System.Drawing.Point(924, 288)
        Me.CmdAllesAn.Name = "CmdAllesAn"
        Me.CmdAllesAn.Size = New System.Drawing.Size(110, 33)
        Me.CmdAllesAn.TabIndex = 24
        Me.CmdAllesAn.Text = "Alles An"
        Me.CmdAllesAn.UseVisualStyleBackColor = True
        '
        'cmdSwitchFeedback
        '
        Me.cmdSwitchFeedback.Location = New System.Drawing.Point(64, 248)
        Me.cmdSwitchFeedback.Name = "cmdSwitchFeedback"
        Me.cmdSwitchFeedback.Size = New System.Drawing.Size(81, 33)
        Me.cmdSwitchFeedback.TabIndex = 17
        Me.cmdSwitchFeedback.Text = "Feedback"
        Me.cmdSwitchFeedback.UseVisualStyleBackColor = True
        '
        'txtWissel
        '
        Me.txtWissel.Location = New System.Drawing.Point(64, 208)
        Me.txtWissel.Name = "txtWissel"
        Me.txtWissel.Size = New System.Drawing.Size(81, 22)
        Me.txtWissel.TabIndex = 9
        Me.txtWissel.Text = ""
        '
        'cmdAf
        '
        Me.cmdAf.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAf.Location = New System.Drawing.Point(160, 208)
        Me.cmdAf.Name = "cmdAf"
        Me.cmdAf.Size = New System.Drawing.Size(49, 25)
        Me.cmdAf.TabIndex = 8
        Me.cmdAf.Text = "<="
        Me.cmdAf.UseVisualStyleBackColor = True
        '
        'cmdRecht
        '
        Me.cmdRecht.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRecht.Location = New System.Drawing.Point(216, 208)
        Me.cmdRecht.Name = "cmdRecht"
        Me.cmdRecht.Size = New System.Drawing.Size(49, 25)
        Me.cmdRecht.TabIndex = 7
        Me.cmdRecht.Text = "=>"
        Me.cmdRecht.UseVisualStyleBackColor = True
        '
        'txtTcpPort
        '
        Me.txtTcpPort.Location = New System.Drawing.Point(96, 64)
        Me.txtTcpPort.Name = "txtTcpPort"
        Me.txtTcpPort.Size = New System.Drawing.Size(97, 22)
        Me.txtTcpPort.TabIndex = 3
        Me.txtTcpPort.Text = "5550"
        '
        'txtIP
        '
        Me.txtIP.Location = New System.Drawing.Point(96, 32)
        Me.txtIP.Name = "txtIP"
        Me.txtIP.Size = New System.Drawing.Size(97, 22)
        Me.txtIP.TabIndex = 2
        Me.txtIP.Text = "192.168.0.200"
        '
        'cmdTcpConnect
        '
        Me.cmdTcpConnect.Location = New System.Drawing.Point(64, 96)
        Me.cmdTcpConnect.Name = "cmdTcpConnect"
        Me.cmdTcpConnect.Size = New System.Drawing.Size(136, 33)
        Me.cmdTcpConnect.TabIndex = 0
        Me.cmdTcpConnect.Text = "Connect TCP/IP"
        Me.cmdTcpConnect.UseVisualStyleBackColor = True
        '
        'LabelBaudrateUSB_1
        '
        Me.LabelBaudrateUSB_1.Location = New System.Drawing.Point(232, 64)
        Me.LabelBaudrateUSB_1.Name = "LabelBaudrateUSB_1"
        Me.LabelBaudrateUSB_1.AutoSize = True
        Me.LabelBaudrateUSB_1.Size = New System.Drawing.Size(63, 16)
        Me.LabelBaudrateUSB_1.TabIndex = 143
        Me.LabelBaudrateUSB_1.Text = "Baudrate"
        Me.LabelBaudrateUSB_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Line9
        '
        Me.Line9.BackColor = System.Drawing.Color.Black
        Me.Line9.Location = New System.Drawing.Point(904, 520)
        Me.Line9.Name = "Line9"
        Me.Line9.Size = New System.Drawing.Size(88, 1)
        '
        'LabelExtValue
        '
        Me.LabelExtValue.Location = New System.Drawing.Point(900, 432)
        Me.LabelExtValue.Name = "LabelExtValue"
        Me.LabelExtValue.AutoSize = False
        Me.LabelExtValue.Size = New System.Drawing.Size(96, 16)
        Me.LabelExtValue.TabIndex = 139
        Me.LabelExtValue.Text = "Value"
        Me.LabelExtValue.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelExtSwitch
        '
        Me.LabelExtSwitch.Location = New System.Drawing.Point(900, 384)
        Me.LabelExtSwitch.Name = "LabelExtSwitch"
        Me.LabelExtSwitch.AutoSize = False
        Me.LabelExtSwitch.Size = New System.Drawing.Size(96, 16)
        Me.LabelExtSwitch.TabIndex = 137
        Me.LabelExtSwitch.Text = "Address"
        Me.LabelExtSwitch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelExtAccCmd
        '
        Me.LabelExtAccCmd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelExtAccCmd.Location = New System.Drawing.Point(900, 360)
        Me.LabelExtAccCmd.Name = "LabelExtAccCmd"
        Me.LabelExtAccCmd.AutoSize = False
        Me.LabelExtAccCmd.Size = New System.Drawing.Size(96, 17)
        Me.LabelExtAccCmd.TabIndex = 136
        Me.LabelExtAccCmd.Text = "ExtAccCmd"
        Me.LabelExtAccCmd.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(808, 184)
        Me.Label4.Name = "Label4"
        Me.Label4.AutoSize = False
        Me.Label4.Size = New System.Drawing.Size(94, 15)
        Me.Label4.TabIndex = 133
        Me.Label4.Text = "Start modus"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Line8
        '
        Me.Line8.BackColor = System.Drawing.Color.Black
        Me.Line8.Location = New System.Drawing.Point(1068, 384)
        Me.Line8.Name = "Line8"
        Me.Line8.Size = New System.Drawing.Size(1, 64)
        '
        'Line_ShapeAnAus_T
        '
        Me.Line_ShapeAnAus_T.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeAnAus_T.Location = New System.Drawing.Point(568, 352)
        Me.Line_ShapeAnAus_T.Name = "Line_ShapeAnAus_T"
        Me.Line_ShapeAnAus_T.Size = New System.Drawing.Size(321, 1)
        '
        'Line_ShapeAnAus_B
        '
        Me.Line_ShapeAnAus_B.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeAnAus_B.Location = New System.Drawing.Point(568, 632)
        Me.Line_ShapeAnAus_B.Name = "Line_ShapeAnAus_B"
        Me.Line_ShapeAnAus_B.Size = New System.Drawing.Size(321, 1)
        '
        'Line_ShapeAnAus_L
        '
        Me.Line_ShapeAnAus_L.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeAnAus_L.Location = New System.Drawing.Point(568, 352)
        Me.Line_ShapeAnAus_L.Name = "Line_ShapeAnAus_L"
        Me.Line_ShapeAnAus_L.Size = New System.Drawing.Size(1, 281)
        '
        'Line_ShapeAnAus_R
        '
        Me.Line_ShapeAnAus_R.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeAnAus_R.Location = New System.Drawing.Point(888, 352)
        Me.Line_ShapeAnAus_R.Name = "Line_ShapeAnAus_R"
        Me.Line_ShapeAnAus_R.Size = New System.Drawing.Size(1, 281)
        '
        'LEDF0F68_68
        '
        Me.LEDF0F68_68.Location = New System.Drawing.Point(832, 544)
        Me.LEDF0F68_68.Name = "LEDF0F68_68"
        Me.LEDF0F68_68.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_68.Round = False
        Me.LEDF0F68_68.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_67
        '
        Me.LEDF0F68_67.Location = New System.Drawing.Point(824, 544)
        Me.LEDF0F68_67.Name = "LEDF0F68_67"
        Me.LEDF0F68_67.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_67.Round = False
        Me.LEDF0F68_67.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_66
        '
        Me.LEDF0F68_66.Location = New System.Drawing.Point(816, 544)
        Me.LEDF0F68_66.Name = "LEDF0F68_66"
        Me.LEDF0F68_66.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_66.Round = False
        Me.LEDF0F68_66.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_65
        '
        Me.LEDF0F68_65.Location = New System.Drawing.Point(808, 544)
        Me.LEDF0F68_65.Name = "LEDF0F68_65"
        Me.LEDF0F68_65.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_65.Round = False
        Me.LEDF0F68_65.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_64
        '
        Me.LEDF0F68_64.Location = New System.Drawing.Point(792, 544)
        Me.LEDF0F68_64.Name = "LEDF0F68_64"
        Me.LEDF0F68_64.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_64.Round = False
        Me.LEDF0F68_64.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_63
        '
        Me.LEDF0F68_63.Location = New System.Drawing.Point(784, 544)
        Me.LEDF0F68_63.Name = "LEDF0F68_63"
        Me.LEDF0F68_63.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_63.Round = False
        Me.LEDF0F68_63.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_62
        '
        Me.LEDF0F68_62.Location = New System.Drawing.Point(776, 544)
        Me.LEDF0F68_62.Name = "LEDF0F68_62"
        Me.LEDF0F68_62.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_62.Round = False
        Me.LEDF0F68_62.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_61
        '
        Me.LEDF0F68_61.Location = New System.Drawing.Point(768, 544)
        Me.LEDF0F68_61.Name = "LEDF0F68_61"
        Me.LEDF0F68_61.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_61.Round = False
        Me.LEDF0F68_61.Size = New System.Drawing.Size(9, 9)
        '
        'LabelFunctions_8
        '
        Me.LabelFunctions_8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_8.Location = New System.Drawing.Point(711, 544)
        Me.LabelFunctions_8.Name = "LabelFunctions_8"
        Me.LabelFunctions_8.AutoSize = False
        Me.LabelFunctions_8.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_8.TabIndex = 122
        Me.LabelFunctions_8.Text = "F61-F68"
        Me.LabelFunctions_8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_7
        '
        Me.LabelFunctions_7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_7.Location = New System.Drawing.Point(711, 528)
        Me.LabelFunctions_7.Name = "LabelFunctions_7"
        Me.LabelFunctions_7.AutoSize = False
        Me.LabelFunctions_7.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_7.TabIndex = 121
        Me.LabelFunctions_7.Text = "F53-F60"
        Me.LabelFunctions_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_6
        '
        Me.LabelFunctions_6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_6.Location = New System.Drawing.Point(711, 512)
        Me.LabelFunctions_6.Name = "LabelFunctions_6"
        Me.LabelFunctions_6.AutoSize = False
        Me.LabelFunctions_6.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_6.TabIndex = 120
        Me.LabelFunctions_6.Text = "F45-F52"
        Me.LabelFunctions_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_5
        '
        Me.LabelFunctions_5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_5.Location = New System.Drawing.Point(711, 496)
        Me.LabelFunctions_5.Name = "LabelFunctions_5"
        Me.LabelFunctions_5.AutoSize = False
        Me.LabelFunctions_5.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_5.TabIndex = 119
        Me.LabelFunctions_5.Text = "F37-F44"
        Me.LabelFunctions_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_4
        '
        Me.LabelFunctions_4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_4.Location = New System.Drawing.Point(711, 480)
        Me.LabelFunctions_4.Name = "LabelFunctions_4"
        Me.LabelFunctions_4.AutoSize = False
        Me.LabelFunctions_4.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_4.TabIndex = 118
        Me.LabelFunctions_4.Text = "F29-F36"
        Me.LabelFunctions_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_3
        '
        Me.LabelFunctions_3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_3.Location = New System.Drawing.Point(711, 464)
        Me.LabelFunctions_3.Name = "LabelFunctions_3"
        Me.LabelFunctions_3.AutoSize = False
        Me.LabelFunctions_3.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_3.TabIndex = 117
        Me.LabelFunctions_3.Text = "F21-F28"
        Me.LabelFunctions_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_2
        '
        Me.LabelFunctions_2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_2.Location = New System.Drawing.Point(713, 448)
        Me.LabelFunctions_2.Name = "LabelFunctions_2"
        Me.LabelFunctions_2.AutoSize = False
        Me.LabelFunctions_2.Size = New System.Drawing.Size(48, 15)
        Me.LabelFunctions_2.TabIndex = 116
        Me.LabelFunctions_2.Text = "F13-F20"
        Me.LabelFunctions_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunctions_1
        '
        Me.LabelFunctions_1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_1.Location = New System.Drawing.Point(721, 432)
        Me.LabelFunctions_1.Name = "LabelFunctions_1"
        Me.LabelFunctions_1.AutoSize = False
        Me.LabelFunctions_1.Size = New System.Drawing.Size(42, 15)
        Me.LabelFunctions_1.TabIndex = 115
        Me.LabelFunctions_1.Text = "F5-F12"
        Me.LabelFunctions_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LEDF0F68_60
        '
        Me.LEDF0F68_60.Location = New System.Drawing.Point(832, 528)
        Me.LEDF0F68_60.Name = "LEDF0F68_60"
        Me.LEDF0F68_60.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_60.Round = False
        Me.LEDF0F68_60.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_59
        '
        Me.LEDF0F68_59.Location = New System.Drawing.Point(824, 528)
        Me.LEDF0F68_59.Name = "LEDF0F68_59"
        Me.LEDF0F68_59.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_59.Round = False
        Me.LEDF0F68_59.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_58
        '
        Me.LEDF0F68_58.Location = New System.Drawing.Point(816, 528)
        Me.LEDF0F68_58.Name = "LEDF0F68_58"
        Me.LEDF0F68_58.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_58.Round = False
        Me.LEDF0F68_58.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_57
        '
        Me.LEDF0F68_57.Location = New System.Drawing.Point(808, 528)
        Me.LEDF0F68_57.Name = "LEDF0F68_57"
        Me.LEDF0F68_57.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_57.Round = False
        Me.LEDF0F68_57.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_56
        '
        Me.LEDF0F68_56.Location = New System.Drawing.Point(792, 528)
        Me.LEDF0F68_56.Name = "LEDF0F68_56"
        Me.LEDF0F68_56.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_56.Round = False
        Me.LEDF0F68_56.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_55
        '
        Me.LEDF0F68_55.Location = New System.Drawing.Point(784, 528)
        Me.LEDF0F68_55.Name = "LEDF0F68_55"
        Me.LEDF0F68_55.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_55.Round = False
        Me.LEDF0F68_55.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_54
        '
        Me.LEDF0F68_54.Location = New System.Drawing.Point(776, 528)
        Me.LEDF0F68_54.Name = "LEDF0F68_54"
        Me.LEDF0F68_54.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_54.Round = False
        Me.LEDF0F68_54.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_53
        '
        Me.LEDF0F68_53.Location = New System.Drawing.Point(768, 528)
        Me.LEDF0F68_53.Name = "LEDF0F68_53"
        Me.LEDF0F68_53.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_53.Round = False
        Me.LEDF0F68_53.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_52
        '
        Me.LEDF0F68_52.Location = New System.Drawing.Point(832, 512)
        Me.LEDF0F68_52.Name = "LEDF0F68_52"
        Me.LEDF0F68_52.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_52.Round = False
        Me.LEDF0F68_52.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_51
        '
        Me.LEDF0F68_51.Location = New System.Drawing.Point(824, 512)
        Me.LEDF0F68_51.Name = "LEDF0F68_51"
        Me.LEDF0F68_51.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_51.Round = False
        Me.LEDF0F68_51.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_50
        '
        Me.LEDF0F68_50.Location = New System.Drawing.Point(816, 512)
        Me.LEDF0F68_50.Name = "LEDF0F68_50"
        Me.LEDF0F68_50.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_50.Round = False
        Me.LEDF0F68_50.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_49
        '
        Me.LEDF0F68_49.Location = New System.Drawing.Point(808, 512)
        Me.LEDF0F68_49.Name = "LEDF0F68_49"
        Me.LEDF0F68_49.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_49.Round = False
        Me.LEDF0F68_49.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_48
        '
        Me.LEDF0F68_48.Location = New System.Drawing.Point(792, 512)
        Me.LEDF0F68_48.Name = "LEDF0F68_48"
        Me.LEDF0F68_48.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_48.Round = False
        Me.LEDF0F68_48.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_47
        '
        Me.LEDF0F68_47.Location = New System.Drawing.Point(784, 512)
        Me.LEDF0F68_47.Name = "LEDF0F68_47"
        Me.LEDF0F68_47.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_47.Round = False
        Me.LEDF0F68_47.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_46
        '
        Me.LEDF0F68_46.Location = New System.Drawing.Point(776, 512)
        Me.LEDF0F68_46.Name = "LEDF0F68_46"
        Me.LEDF0F68_46.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_46.Round = False
        Me.LEDF0F68_46.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_45
        '
        Me.LEDF0F68_45.Location = New System.Drawing.Point(768, 512)
        Me.LEDF0F68_45.Name = "LEDF0F68_45"
        Me.LEDF0F68_45.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_45.Round = False
        Me.LEDF0F68_45.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_44
        '
        Me.LEDF0F68_44.Location = New System.Drawing.Point(832, 496)
        Me.LEDF0F68_44.Name = "LEDF0F68_44"
        Me.LEDF0F68_44.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_44.Round = False
        Me.LEDF0F68_44.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_43
        '
        Me.LEDF0F68_43.Location = New System.Drawing.Point(824, 496)
        Me.LEDF0F68_43.Name = "LEDF0F68_43"
        Me.LEDF0F68_43.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_43.Round = False
        Me.LEDF0F68_43.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_42
        '
        Me.LEDF0F68_42.Location = New System.Drawing.Point(816, 496)
        Me.LEDF0F68_42.Name = "LEDF0F68_42"
        Me.LEDF0F68_42.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_42.Round = False
        Me.LEDF0F68_42.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_41
        '
        Me.LEDF0F68_41.Location = New System.Drawing.Point(808, 496)
        Me.LEDF0F68_41.Name = "LEDF0F68_41"
        Me.LEDF0F68_41.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_41.Round = False
        Me.LEDF0F68_41.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_40
        '
        Me.LEDF0F68_40.Location = New System.Drawing.Point(792, 496)
        Me.LEDF0F68_40.Name = "LEDF0F68_40"
        Me.LEDF0F68_40.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_40.Round = False
        Me.LEDF0F68_40.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_39
        '
        Me.LEDF0F68_39.Location = New System.Drawing.Point(784, 496)
        Me.LEDF0F68_39.Name = "LEDF0F68_39"
        Me.LEDF0F68_39.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_39.Round = False
        Me.LEDF0F68_39.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_38
        '
        Me.LEDF0F68_38.Location = New System.Drawing.Point(776, 496)
        Me.LEDF0F68_38.Name = "LEDF0F68_38"
        Me.LEDF0F68_38.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_38.Round = False
        Me.LEDF0F68_38.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_37
        '
        Me.LEDF0F68_37.Location = New System.Drawing.Point(768, 496)
        Me.LEDF0F68_37.Name = "LEDF0F68_37"
        Me.LEDF0F68_37.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_37.Round = False
        Me.LEDF0F68_37.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_36
        '
        Me.LEDF0F68_36.Location = New System.Drawing.Point(832, 480)
        Me.LEDF0F68_36.Name = "LEDF0F68_36"
        Me.LEDF0F68_36.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_36.Round = False
        Me.LEDF0F68_36.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_35
        '
        Me.LEDF0F68_35.Location = New System.Drawing.Point(824, 480)
        Me.LEDF0F68_35.Name = "LEDF0F68_35"
        Me.LEDF0F68_35.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_35.Round = False
        Me.LEDF0F68_35.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_34
        '
        Me.LEDF0F68_34.Location = New System.Drawing.Point(816, 480)
        Me.LEDF0F68_34.Name = "LEDF0F68_34"
        Me.LEDF0F68_34.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_34.Round = False
        Me.LEDF0F68_34.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_33
        '
        Me.LEDF0F68_33.Location = New System.Drawing.Point(808, 480)
        Me.LEDF0F68_33.Name = "LEDF0F68_33"
        Me.LEDF0F68_33.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_33.Round = False
        Me.LEDF0F68_33.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_32
        '
        Me.LEDF0F68_32.Location = New System.Drawing.Point(792, 480)
        Me.LEDF0F68_32.Name = "LEDF0F68_32"
        Me.LEDF0F68_32.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_32.Round = False
        Me.LEDF0F68_32.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_31
        '
        Me.LEDF0F68_31.Location = New System.Drawing.Point(784, 480)
        Me.LEDF0F68_31.Name = "LEDF0F68_31"
        Me.LEDF0F68_31.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_31.Round = False
        Me.LEDF0F68_31.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_30
        '
        Me.LEDF0F68_30.Location = New System.Drawing.Point(776, 480)
        Me.LEDF0F68_30.Name = "LEDF0F68_30"
        Me.LEDF0F68_30.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_30.Round = False
        Me.LEDF0F68_30.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_29
        '
        Me.LEDF0F68_29.Location = New System.Drawing.Point(768, 480)
        Me.LEDF0F68_29.Name = "LEDF0F68_29"
        Me.LEDF0F68_29.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_29.Round = False
        Me.LEDF0F68_29.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_28
        '
        Me.LEDF0F68_28.Location = New System.Drawing.Point(832, 464)
        Me.LEDF0F68_28.Name = "LEDF0F68_28"
        Me.LEDF0F68_28.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_28.Round = False
        Me.LEDF0F68_28.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_27
        '
        Me.LEDF0F68_27.Location = New System.Drawing.Point(824, 464)
        Me.LEDF0F68_27.Name = "LEDF0F68_27"
        Me.LEDF0F68_27.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_27.Round = False
        Me.LEDF0F68_27.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_26
        '
        Me.LEDF0F68_26.Location = New System.Drawing.Point(816, 464)
        Me.LEDF0F68_26.Name = "LEDF0F68_26"
        Me.LEDF0F68_26.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_26.Round = False
        Me.LEDF0F68_26.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_25
        '
        Me.LEDF0F68_25.Location = New System.Drawing.Point(808, 464)
        Me.LEDF0F68_25.Name = "LEDF0F68_25"
        Me.LEDF0F68_25.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_25.Round = False
        Me.LEDF0F68_25.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_24
        '
        Me.LEDF0F68_24.Location = New System.Drawing.Point(792, 464)
        Me.LEDF0F68_24.Name = "LEDF0F68_24"
        Me.LEDF0F68_24.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_24.Round = False
        Me.LEDF0F68_24.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_23
        '
        Me.LEDF0F68_23.Location = New System.Drawing.Point(784, 464)
        Me.LEDF0F68_23.Name = "LEDF0F68_23"
        Me.LEDF0F68_23.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_23.Round = False
        Me.LEDF0F68_23.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_22
        '
        Me.LEDF0F68_22.Location = New System.Drawing.Point(776, 464)
        Me.LEDF0F68_22.Name = "LEDF0F68_22"
        Me.LEDF0F68_22.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_22.Round = False
        Me.LEDF0F68_22.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_21
        '
        Me.LEDF0F68_21.Location = New System.Drawing.Point(768, 464)
        Me.LEDF0F68_21.Name = "LEDF0F68_21"
        Me.LEDF0F68_21.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_21.Round = False
        Me.LEDF0F68_21.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_20
        '
        Me.LEDF0F68_20.Location = New System.Drawing.Point(832, 448)
        Me.LEDF0F68_20.Name = "LEDF0F68_20"
        Me.LEDF0F68_20.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_20.Round = False
        Me.LEDF0F68_20.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_19
        '
        Me.LEDF0F68_19.Location = New System.Drawing.Point(824, 448)
        Me.LEDF0F68_19.Name = "LEDF0F68_19"
        Me.LEDF0F68_19.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_19.Round = False
        Me.LEDF0F68_19.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_18
        '
        Me.LEDF0F68_18.Location = New System.Drawing.Point(816, 448)
        Me.LEDF0F68_18.Name = "LEDF0F68_18"
        Me.LEDF0F68_18.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_18.Round = False
        Me.LEDF0F68_18.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_17
        '
        Me.LEDF0F68_17.Location = New System.Drawing.Point(808, 448)
        Me.LEDF0F68_17.Name = "LEDF0F68_17"
        Me.LEDF0F68_17.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_17.Round = False
        Me.LEDF0F68_17.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_16
        '
        Me.LEDF0F68_16.Location = New System.Drawing.Point(792, 448)
        Me.LEDF0F68_16.Name = "LEDF0F68_16"
        Me.LEDF0F68_16.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_16.Round = False
        Me.LEDF0F68_16.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_15
        '
        Me.LEDF0F68_15.Location = New System.Drawing.Point(784, 448)
        Me.LEDF0F68_15.Name = "LEDF0F68_15"
        Me.LEDF0F68_15.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_15.Round = False
        Me.LEDF0F68_15.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_14
        '
        Me.LEDF0F68_14.Location = New System.Drawing.Point(776, 448)
        Me.LEDF0F68_14.Name = "LEDF0F68_14"
        Me.LEDF0F68_14.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_14.Round = False
        Me.LEDF0F68_14.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_13
        '
        Me.LEDF0F68_13.Location = New System.Drawing.Point(768, 448)
        Me.LEDF0F68_13.Name = "LEDF0F68_13"
        Me.LEDF0F68_13.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_13.Round = False
        Me.LEDF0F68_13.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_12
        '
        Me.LEDF0F68_12.Location = New System.Drawing.Point(832, 432)
        Me.LEDF0F68_12.Name = "LEDF0F68_12"
        Me.LEDF0F68_12.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_12.Round = False
        Me.LEDF0F68_12.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_11
        '
        Me.LEDF0F68_11.Location = New System.Drawing.Point(824, 432)
        Me.LEDF0F68_11.Name = "LEDF0F68_11"
        Me.LEDF0F68_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_11.Round = False
        Me.LEDF0F68_11.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_10
        '
        Me.LEDF0F68_10.Location = New System.Drawing.Point(816, 432)
        Me.LEDF0F68_10.Name = "LEDF0F68_10"
        Me.LEDF0F68_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_10.Round = False
        Me.LEDF0F68_10.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_9
        '
        Me.LEDF0F68_9.Location = New System.Drawing.Point(808, 432)
        Me.LEDF0F68_9.Name = "LEDF0F68_9"
        Me.LEDF0F68_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_9.Round = False
        Me.LEDF0F68_9.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_8
        '
        Me.LEDF0F68_8.Location = New System.Drawing.Point(792, 432)
        Me.LEDF0F68_8.Name = "LEDF0F68_8"
        Me.LEDF0F68_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_8.Round = False
        Me.LEDF0F68_8.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_7
        '
        Me.LEDF0F68_7.Location = New System.Drawing.Point(784, 432)
        Me.LEDF0F68_7.Name = "LEDF0F68_7"
        Me.LEDF0F68_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_7.Round = False
        Me.LEDF0F68_7.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_6
        '
        Me.LEDF0F68_6.Location = New System.Drawing.Point(776, 432)
        Me.LEDF0F68_6.Name = "LEDF0F68_6"
        Me.LEDF0F68_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_6.Round = False
        Me.LEDF0F68_6.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_5
        '
        Me.LEDF0F68_5.Location = New System.Drawing.Point(768, 432)
        Me.LEDF0F68_5.Name = "LEDF0F68_5"
        Me.LEDF0F68_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_5.Round = False
        Me.LEDF0F68_5.Size = New System.Drawing.Size(9, 9)
        '
        'LabelFunctions_0
        '
        Me.LabelFunctions_0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFunctions_0.Location = New System.Drawing.Point(727, 415)
        Me.LabelFunctions_0.Name = "LabelFunctions_0"
        Me.LabelFunctions_0.AutoSize = False
        Me.LabelFunctions_0.Size = New System.Drawing.Size(36, 15)
        Me.LabelFunctions_0.TabIndex = 114
        Me.LabelFunctions_0.Text = "F0-F4"
        Me.LabelFunctions_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LEDF0F68_4
        '
        Me.LEDF0F68_4.Location = New System.Drawing.Point(832, 416)
        Me.LEDF0F68_4.Name = "LEDF0F68_4"
        Me.LEDF0F68_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_4.Round = False
        Me.LEDF0F68_4.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_3
        '
        Me.LEDF0F68_3.Location = New System.Drawing.Point(824, 416)
        Me.LEDF0F68_3.Name = "LEDF0F68_3"
        Me.LEDF0F68_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_3.Round = False
        Me.LEDF0F68_3.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_2
        '
        Me.LEDF0F68_2.Location = New System.Drawing.Point(816, 416)
        Me.LEDF0F68_2.Name = "LEDF0F68_2"
        Me.LEDF0F68_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_2.Round = False
        Me.LEDF0F68_2.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_1
        '
        Me.LEDF0F68_1.Location = New System.Drawing.Point(808, 416)
        Me.LEDF0F68_1.Name = "LEDF0F68_1"
        Me.LEDF0F68_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_1.Round = False
        Me.LEDF0F68_1.Size = New System.Drawing.Size(9, 9)
        '
        'LEDF0F68_0
        '
        Me.LEDF0F68_0.Location = New System.Drawing.Point(792, 416)
        Me.LEDF0F68_0.Name = "LEDF0F68_0"
        Me.LEDF0F68_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.LEDF0F68_0.Round = False
        Me.LEDF0F68_0.Size = New System.Drawing.Size(9, 9)
        '
        'Line_Shape4_T
        '
        Me.Line_Shape4_T.BackColor = System.Drawing.Color.Black
        Me.Line_Shape4_T.Location = New System.Drawing.Point(1008, 352)
        Me.Line_Shape4_T.Name = "Line_Shape4_T"
        Me.Line_Shape4_T.Size = New System.Drawing.Size(165, 1)
        '
        'Line_Shape4_B
        '
        Me.Line_Shape4_B.BackColor = System.Drawing.Color.Black
        Me.Line_Shape4_B.Location = New System.Drawing.Point(1008, 632)
        Me.Line_Shape4_B.Name = "Line_Shape4_B"
        Me.Line_Shape4_B.Size = New System.Drawing.Size(165, 1)
        '
        'Line_Shape4_L
        '
        Me.Line_Shape4_L.BackColor = System.Drawing.Color.Black
        Me.Line_Shape4_L.Location = New System.Drawing.Point(1008, 352)
        Me.Line_Shape4_L.Name = "Line_Shape4_L"
        Me.Line_Shape4_L.Size = New System.Drawing.Size(1, 281)
        '
        'Line_Shape4_R
        '
        Me.Line_Shape4_R.BackColor = System.Drawing.Color.Black
        Me.Line_Shape4_R.Location = New System.Drawing.Point(1172, 352)
        Me.Line_Shape4_R.Name = "Line_Shape4_R"
        Me.Line_Shape4_R.Size = New System.Drawing.Size(1, 281)
        '
        'LabelPom
        '
        Me.LabelPom.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPom.Location = New System.Drawing.Point(1020, 358)
        Me.LabelPom.Name = "LabelPom"
        Me.LabelPom.AutoSize = False
        Me.LabelPom.Size = New System.Drawing.Size(140, 22)
        Me.LabelPom.TabIndex = 112
        Me.LabelPom.Text = "Programmieren"
        Me.LabelPom.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelPomValue
        '
        Me.LabelPomValue.Location = New System.Drawing.Point(1014, 528)
        Me.LabelPomValue.Name = "LabelPomValue"
        Me.LabelPomValue.AutoSize = False
        Me.LabelPomValue.Size = New System.Drawing.Size(66, 17)
        Me.LabelPomValue.TabIndex = 110
        Me.LabelPomValue.Text = "Value"
        Me.LabelPomValue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPomCV
        '
        Me.LabelPomCV.Location = New System.Drawing.Point(1014, 496)
        Me.LabelPomCV.Name = "LabelPomCV"
        Me.LabelPomCV.AutoSize = False
        Me.LabelPomCV.Size = New System.Drawing.Size(66, 17)
        Me.LabelPomCV.TabIndex = 109
        Me.LabelPomCV.Text = "CV / Reg"
        Me.LabelPomCV.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelPomAddress
        '
        Me.LabelPomAddress.Location = New System.Drawing.Point(1014, 464)
        Me.LabelPomAddress.Name = "LabelPomAddress"
        Me.LabelPomAddress.AutoSize = False
        Me.LabelPomAddress.Size = New System.Drawing.Size(66, 17)
        Me.LabelPomAddress.TabIndex = 108
        Me.LabelPomAddress.Text = "Address"
        Me.LabelPomAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFunktion
        '
        Me.LabelFunktion.Location = New System.Drawing.Point(725, 387)
        Me.LabelFunktion.Name = "LabelFunktion"
        Me.LabelFunktion.AutoSize = False
        Me.LabelFunktion.Size = New System.Drawing.Size(19, 17)
        Me.LabelFunktion.TabIndex = 103
        Me.LabelFunktion.Text = "F:"
        Me.LabelFunktion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelLoks
        '
        Me.LabelLoks.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelLoks.Location = New System.Drawing.Point(613, 355)
        Me.LabelLoks.Name = "LabelLoks"
        Me.LabelLoks.AutoSize = False
        Me.LabelLoks.Size = New System.Drawing.Size(60, 24)
        Me.LabelLoks.TabIndex = 96
        Me.LabelLoks.Text = "Loks"
        Me.LabelLoks.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelLokSteps
        '
        Me.LabelLokSteps.Location = New System.Drawing.Point(572, 568)
        Me.LabelLokSteps.Name = "LabelLokSteps"
        Me.LabelLokSteps.Size = New System.Drawing.Size(62, 17)
        Me.LabelLokSteps.TabIndex = 95
        Me.LabelLokSteps.Text = "Steps"
        Me.LabelLokSteps.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelLokSpeed
        '
        Me.LabelLokSpeed.Location = New System.Drawing.Point(572, 464)
        Me.LabelLokSpeed.Name = "LabelLokSpeed"
        Me.LabelLokSpeed.Size = New System.Drawing.Size(62, 17)
        Me.LabelLokSpeed.TabIndex = 94
        Me.LabelLokSpeed.Text = "Speed"
        Me.LabelLokSpeed.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelLokAddress
        '
        Me.LabelLokAddress.Location = New System.Drawing.Point(572, 384)
        Me.LabelLokAddress.Name = "LabelLokAddress"
        Me.LabelLokAddress.Size = New System.Drawing.Size(62, 17)
        Me.LabelLokAddress.TabIndex = 93
        Me.LabelLokAddress.Text = "Address"
        Me.LabelLokAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelModellzeit
        '
        Me.LabelModellzeit.Location = New System.Drawing.Point(716, 192)
        Me.LabelModellzeit.Name = "LabelModellzeit"
        Me.LabelModellzeit.Size = New System.Drawing.Size(80, 17)
        Me.LabelModellzeit.TabIndex = 92
        Me.LabelModellzeit.Text = "Modellzeit"
        Me.LabelModellzeit.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Line_Shape3_T
        '
        Me.Line_Shape3_T.BackColor = System.Drawing.Color.Black
        Me.Line_Shape3_T.Location = New System.Drawing.Point(568, 184)
        Me.Line_Shape3_T.Name = "Line_Shape3_T"
        Me.Line_Shape3_T.Size = New System.Drawing.Size(233, 1)
        '
        'Line_Shape3_B
        '
        Me.Line_Shape3_B.BackColor = System.Drawing.Color.Black
        Me.Line_Shape3_B.Location = New System.Drawing.Point(568, 336)
        Me.Line_Shape3_B.Name = "Line_Shape3_B"
        Me.Line_Shape3_B.Size = New System.Drawing.Size(233, 1)
        '
        'Line_Shape3_L
        '
        Me.Line_Shape3_L.BackColor = System.Drawing.Color.Black
        Me.Line_Shape3_L.Location = New System.Drawing.Point(568, 184)
        Me.Line_Shape3_L.Name = "Line_Shape3_L"
        Me.Line_Shape3_L.Size = New System.Drawing.Size(1, 153)
        '
        'Line_Shape3_R
        '
        Me.Line_Shape3_R.BackColor = System.Drawing.Color.Black
        Me.Line_Shape3_R.Location = New System.Drawing.Point(800, 184)
        Me.Line_Shape3_R.Name = "Line_Shape3_R"
        Me.Line_Shape3_R.Size = New System.Drawing.Size(1, 153)
        '
        'Line7
        '
        Me.Line7.BackColor = System.Drawing.Color.Black
        Me.Line7.Location = New System.Drawing.Point(712, 216)
        Me.Line7.Name = "Line7"
        Me.Line7.Size = New System.Drawing.Size(1, 112)
        '
        'Line_Shape2_T
        '
        Me.Line_Shape2_T.BackColor = System.Drawing.Color.Black
        Me.Line_Shape2_T.Location = New System.Drawing.Point(908, 184)
        Me.Line_Shape2_T.Name = "Line_Shape2_T"
        Me.Line_Shape2_T.Size = New System.Drawing.Size(265, 1)
        '
        'Line_Shape2_B
        '
        Me.Line_Shape2_B.BackColor = System.Drawing.Color.Black
        Me.Line_Shape2_B.Location = New System.Drawing.Point(908, 264)
        Me.Line_Shape2_B.Name = "Line_Shape2_B"
        Me.Line_Shape2_B.Size = New System.Drawing.Size(265, 1)
        '
        'Line_Shape2_L
        '
        Me.Line_Shape2_L.BackColor = System.Drawing.Color.Black
        Me.Line_Shape2_L.Location = New System.Drawing.Point(908, 184)
        Me.Line_Shape2_L.Name = "Line_Shape2_L"
        Me.Line_Shape2_L.Size = New System.Drawing.Size(1, 81)
        '
        'Line_Shape2_R
        '
        Me.Line_Shape2_R.BackColor = System.Drawing.Color.Black
        Me.Line_Shape2_R.Location = New System.Drawing.Point(1172, 184)
        Me.Line_Shape2_R.Name = "Line_Shape2_R"
        Me.Line_Shape2_R.Size = New System.Drawing.Size(1, 81)
        '
        'LabelDay
        '
        Me.LabelDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDay.Location = New System.Drawing.Point(572, 216)
        Me.LabelDay.Name = "LabelDay"
        Me.LabelDay.Size = New System.Drawing.Size(42, 25)
        Me.LabelDay.TabIndex = 91
        Me.LabelDay.Text = "Day"
        Me.LabelDay.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelHour
        '
        Me.LabelHour.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelHour.Location = New System.Drawing.Point(572, 240)
        Me.LabelHour.Name = "LabelHour"
        Me.LabelHour.Size = New System.Drawing.Size(42, 25)
        Me.LabelHour.TabIndex = 90
        Me.LabelHour.Text = "Hour"
        Me.LabelHour.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelMinute
        '
        Me.LabelMinute.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMinute.Location = New System.Drawing.Point(572, 272)
        Me.LabelMinute.Name = "LabelMinute"
        Me.LabelMinute.Size = New System.Drawing.Size(42, 25)
        Me.LabelMinute.TabIndex = 89
        Me.LabelMinute.Text = "Minute"
        Me.LabelMinute.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelFacor
        '
        Me.LabelFacor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFacor.Location = New System.Drawing.Point(572, 304)
        Me.LabelFacor.Name = "LabelFacor"
        Me.LabelFacor.Size = New System.Drawing.Size(42, 25)
        Me.LabelFacor.TabIndex = 88
        Me.LabelFacor.Text = "Factor"
        Me.LabelFacor.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TxtMinutes
        '
        Me.TxtMinutes.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtMinutes.Location = New System.Drawing.Point(646, 188)
        Me.TxtMinutes.Name = "TxtMinutes"
        Me.TxtMinutes.Size = New System.Drawing.Size(30, 24)
        Me.TxtMinutes.TabIndex = 77
        Me.TxtMinutes.Text = "00"
        Me.TxtMinutes.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(634, 188)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(12, 24)
        Me.Label5.TabIndex = 76
        Me.Label5.Text = ":"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TxtHours
        '
        Me.TxtHours.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtHours.Location = New System.Drawing.Point(604, 188)
        Me.TxtHours.Name = "TxtHours"
        Me.TxtHours.Size = New System.Drawing.Size(30, 24)
        Me.TxtHours.TabIndex = 75
        Me.TxtHours.Text = "00"
        Me.TxtHours.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtVersionBootloader
        '
        Me.txtVersionBootloader.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVersionBootloader.Location = New System.Drawing.Point(1056, 240)
        Me.txtVersionBootloader.Name = "txtVersionBootloader"
        Me.txtVersionBootloader.Size = New System.Drawing.Size(38, 15)
        Me.txtVersionBootloader.TabIndex = 73
        Me.txtVersionBootloader.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtBuildRuckmelder
        '
        Me.txtBuildRuckmelder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBuildRuckmelder.Location = New System.Drawing.Point(1056, 224)
        Me.txtBuildRuckmelder.Name = "txtBuildRuckmelder"
        Me.txtBuildRuckmelder.Size = New System.Drawing.Size(38, 15)
        Me.txtBuildRuckmelder.TabIndex = 72
        Me.txtBuildRuckmelder.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'TxtVersionRuckmelder
        '
        Me.TxtVersionRuckmelder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtVersionRuckmelder.Location = New System.Drawing.Point(1056, 208)
        Me.TxtVersionRuckmelder.Name = "TxtVersionRuckmelder"
        Me.TxtVersionRuckmelder.Size = New System.Drawing.Size(38, 15)
        Me.TxtVersionRuckmelder.TabIndex = 71
        Me.TxtVersionRuckmelder.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtBuildNr
        '
        Me.txtBuildNr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBuildNr.Location = New System.Drawing.Point(1056, 192)
        Me.txtBuildNr.Name = "txtBuildNr"
        Me.txtBuildNr.Size = New System.Drawing.Size(38, 15)
        Me.txtBuildNr.TabIndex = 70
        Me.txtBuildNr.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelVersionBootloader
        '
        Me.LabelVersionBootloader.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVersionBootloader.Location = New System.Drawing.Point(912, 240)
        Me.LabelVersionBootloader.Name = "LabelVersionBootloader"
        Me.LabelVersionBootloader.AutoSize = False
        Me.LabelVersionBootloader.Size = New System.Drawing.Size(140, 15)
        Me.LabelVersionBootloader.TabIndex = 69
        Me.LabelVersionBootloader.Text = "Version Bootloader"
        Me.LabelVersionBootloader.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelBuildruckmeldedecoder
        '
        Me.LabelBuildruckmeldedecoder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelBuildruckmeldedecoder.Location = New System.Drawing.Point(912, 224)
        Me.LabelBuildruckmeldedecoder.Name = "LabelBuildruckmeldedecoder"
        Me.LabelBuildruckmeldedecoder.AutoSize = False
        Me.LabelBuildruckmeldedecoder.Size = New System.Drawing.Size(140, 15)
        Me.LabelBuildruckmeldedecoder.TabIndex = 68
        Me.LabelBuildruckmeldedecoder.Text = "Build Ruckmeldedecoder"
        Me.LabelBuildruckmeldedecoder.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelVersionRuckmeldedecoder
        '
        Me.LabelVersionRuckmeldedecoder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVersionRuckmeldedecoder.Location = New System.Drawing.Point(912, 208)
        Me.LabelVersionRuckmeldedecoder.Name = "LabelVersionRuckmeldedecoder"
        Me.LabelVersionRuckmeldedecoder.AutoSize = False
        Me.LabelVersionRuckmeldedecoder.Size = New System.Drawing.Size(140, 15)
        Me.LabelVersionRuckmeldedecoder.TabIndex = 67
        Me.LabelVersionRuckmeldedecoder.Text = "Version Ruckmeldedecoder"
        Me.LabelVersionRuckmeldedecoder.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelBuildNr
        '
        Me.LabelBuildNr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelBuildNr.Location = New System.Drawing.Point(912, 192)
        Me.LabelBuildNr.Name = "LabelBuildNr"
        Me.LabelBuildNr.AutoSize = False
        Me.LabelBuildNr.Size = New System.Drawing.Size(140, 15)
        Me.LabelBuildNr.TabIndex = 66
        Me.LabelBuildNr.Text = "Build-Nr."
        Me.LabelBuildNr.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Line6
        '
        Me.Line6.BackColor = System.Drawing.Color.Black
        Me.Line6.Location = New System.Drawing.Point(916, 40)
        Me.Line6.Name = "Line6"
        Me.Line6.Size = New System.Drawing.Size(248, 1)
        '
        'Line_Shape1_T
        '
        Me.Line_Shape1_T.BackColor = System.Drawing.Color.Black
        Me.Line_Shape1_T.Location = New System.Drawing.Point(908, 272)
        Me.Line_Shape1_T.Name = "Line_Shape1_T"
        Me.Line_Shape1_T.Size = New System.Drawing.Size(265, 1)
        '
        'Line_Shape1_B
        '
        Me.Line_Shape1_B.BackColor = System.Drawing.Color.Black
        Me.Line_Shape1_B.Location = New System.Drawing.Point(908, 336)
        Me.Line_Shape1_B.Name = "Line_Shape1_B"
        Me.Line_Shape1_B.Size = New System.Drawing.Size(265, 1)
        '
        'Line_Shape1_L
        '
        Me.Line_Shape1_L.BackColor = System.Drawing.Color.Black
        Me.Line_Shape1_L.Location = New System.Drawing.Point(908, 272)
        Me.Line_Shape1_L.Name = "Line_Shape1_L"
        Me.Line_Shape1_L.Size = New System.Drawing.Size(1, 65)
        '
        'Line_Shape1_R
        '
        Me.Line_Shape1_R.BackColor = System.Drawing.Color.Black
        Me.Line_Shape1_R.Location = New System.Drawing.Point(1172, 272)
        Me.Line_Shape1_R.Name = "Line_Shape1_R"
        Me.Line_Shape1_R.Size = New System.Drawing.Size(1, 65)
        '
        'LabelInfo2
        '
        Me.LabelInfo2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelInfo2.Location = New System.Drawing.Point(344, 592)
        Me.LabelInfo2.Name = "LabelInfo2"
        Me.LabelInfo2.AutoSize = True
        Me.LabelInfo2.Size = New System.Drawing.Size(136, 16)
        Me.LabelInfo2.TabIndex = 64
        Me.LabelInfo2.Text = "latest message at top"
        Me.LabelInfo2.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(480, 304)
        Me.Label3.Name = "Label3"
        Me.Label3.AutoSize = False
        Me.Label3.Size = New System.Drawing.Size(41, 15)
        Me.Label3.TabIndex = 62
        Me.Label3.Text = "Nibble"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(408, 304)
        Me.Label2.Name = "Label2"
        Me.Label2.AutoSize = False
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 61
        Me.Label2.Text = "Data"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(312, 304)
        Me.Label1.Name = "Label1"
        Me.Label1.AutoSize = False
        Me.Label1.Size = New System.Drawing.Size(97, 17)
        Me.Label1.TabIndex = 60
        Me.Label1.Text = "Address"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelRsBusAddress
        '
        Me.LabelRsBusAddress.Location = New System.Drawing.Point(308, 208)
        Me.LabelRsBusAddress.Name = "LabelRsBusAddress"
        Me.LabelRsBusAddress.AutoSize = False
        Me.LabelRsBusAddress.Size = New System.Drawing.Size(108, 17)
        Me.LabelRsBusAddress.TabIndex = 59
        Me.LabelRsBusAddress.Text = "RSBus Address"
        Me.LabelRsBusAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Line_ShapeFeedback_T
        '
        Me.Line_ShapeFeedback_T.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeFeedback_T.Location = New System.Drawing.Point(296, 184)
        Me.Line_ShapeFeedback_T.Name = "Line_ShapeFeedback_T"
        Me.Line_ShapeFeedback_T.Size = New System.Drawing.Size(241, 2)
        '
        'Line_ShapeFeedback_B
        '
        Me.Line_ShapeFeedback_B.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeFeedback_B.Location = New System.Drawing.Point(296, 631)
        Me.Line_ShapeFeedback_B.Name = "Line_ShapeFeedback_B"
        Me.Line_ShapeFeedback_B.Size = New System.Drawing.Size(241, 2)
        '
        'Line_ShapeFeedback_L
        '
        Me.Line_ShapeFeedback_L.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeFeedback_L.Location = New System.Drawing.Point(296, 184)
        Me.Line_ShapeFeedback_L.Name = "Line_ShapeFeedback_L"
        Me.Line_ShapeFeedback_L.Size = New System.Drawing.Size(2, 449)
        '
        'Line_ShapeFeedback_R
        '
        Me.Line_ShapeFeedback_R.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeFeedback_R.Location = New System.Drawing.Point(535, 184)
        Me.Line_ShapeFeedback_R.Name = "Line_ShapeFeedback_R"
        Me.Line_ShapeFeedback_R.Size = New System.Drawing.Size(2, 449)
        '
        'LabelSwitchIBit
        '
        Me.LabelSwitchIBit.Location = New System.Drawing.Point(232, 304)
        Me.LabelSwitchIBit.Name = "LabelSwitchIBit"
        Me.LabelSwitchIBit.AutoSize = False
        Me.LabelSwitchIBit.Size = New System.Drawing.Size(25, 17)
        Me.LabelSwitchIBit.TabIndex = 54
        Me.LabelSwitchIBit.Text = "I"
        Me.LabelSwitchIBit.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchTtBits
        '
        Me.LabelSwitchTtBits.Location = New System.Drawing.Point(192, 304)
        Me.LabelSwitchTtBits.Name = "LabelSwitchTtBits"
        Me.LabelSwitchTtBits.AutoSize = False
        Me.LabelSwitchTtBits.Size = New System.Drawing.Size(41, 17)
        Me.LabelSwitchTtBits.TabIndex = 53
        Me.LabelSwitchTtBits.Text = "TT"
        Me.LabelSwitchTtBits.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchdata
        '
        Me.LabelSwitchdata.Location = New System.Drawing.Point(120, 304)
        Me.LabelSwitchdata.Name = "LabelSwitchdata"
        Me.LabelSwitchdata.AutoSize = False
        Me.LabelSwitchdata.Size = New System.Drawing.Size(73, 17)
        Me.LabelSwitchdata.TabIndex = 52
        Me.LabelSwitchdata.Text = "Data"
        Me.LabelSwitchdata.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchAddress
        '
        Me.LabelSwitchAddress.Location = New System.Drawing.Point(24, 304)
        Me.LabelSwitchAddress.Name = "LabelSwitchAddress"
        Me.LabelSwitchAddress.AutoSize = False
        Me.LabelSwitchAddress.Size = New System.Drawing.Size(97, 17)
        Me.LabelSwitchAddress.TabIndex = 51
        Me.LabelSwitchAddress.Text = "Address"
        Me.LabelSwitchAddress.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Line_ShapeSwitch_T
        '
        Me.Line_ShapeSwitch_T.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSwitch_T.Location = New System.Drawing.Point(8, 184)
        Me.Line_ShapeSwitch_T.Name = "Line_ShapeSwitch_T"
        Me.Line_ShapeSwitch_T.Size = New System.Drawing.Size(273, 2)
        '
        'Line_ShapeSwitch_B
        '
        Me.Line_ShapeSwitch_B.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSwitch_B.Location = New System.Drawing.Point(8, 631)
        Me.Line_ShapeSwitch_B.Name = "Line_ShapeSwitch_B"
        Me.Line_ShapeSwitch_B.Size = New System.Drawing.Size(273, 2)
        '
        'Line_ShapeSwitch_L
        '
        Me.Line_ShapeSwitch_L.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSwitch_L.Location = New System.Drawing.Point(8, 184)
        Me.Line_ShapeSwitch_L.Name = "Line_ShapeSwitch_L"
        Me.Line_ShapeSwitch_L.Size = New System.Drawing.Size(2, 449)
        '
        'Line_ShapeSwitch_R
        '
        Me.Line_ShapeSwitch_R.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSwitch_R.Location = New System.Drawing.Point(279, 184)
        Me.Line_ShapeSwitch_R.Name = "Line_ShapeSwitch_R"
        Me.Line_ShapeSwitch_R.Size = New System.Drawing.Size(2, 449)
        '
        'Line_ShapeInfo_T
        '
        Me.Line_ShapeInfo_T.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeInfo_T.Location = New System.Drawing.Point(908, 8)
        Me.Line_ShapeInfo_T.Name = "Line_ShapeInfo_T"
        Me.Line_ShapeInfo_T.Size = New System.Drawing.Size(265, 1)
        '
        'Line_ShapeInfo_B
        '
        Me.Line_ShapeInfo_B.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeInfo_B.Location = New System.Drawing.Point(908, 168)
        Me.Line_ShapeInfo_B.Name = "Line_ShapeInfo_B"
        Me.Line_ShapeInfo_B.Size = New System.Drawing.Size(265, 1)
        '
        'Line_ShapeInfo_L
        '
        Me.Line_ShapeInfo_L.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeInfo_L.Location = New System.Drawing.Point(908, 8)
        Me.Line_ShapeInfo_L.Name = "Line_ShapeInfo_L"
        Me.Line_ShapeInfo_L.Size = New System.Drawing.Size(1, 161)
        '
        'Line_ShapeInfo_R
        '
        Me.Line_ShapeInfo_R.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeInfo_R.Location = New System.Drawing.Point(1172, 8)
        Me.Line_ShapeInfo_R.Name = "Line_ShapeInfo_R"
        Me.Line_ShapeInfo_R.Size = New System.Drawing.Size(1, 161)
        '
        'LabelConnected
        '
        Me.LabelConnected.Location = New System.Drawing.Point(777, 18)
        Me.LabelConnected.Name = "LabelConnected"
        Me.LabelConnected.AutoSize = True
        Me.LabelConnected.Size = New System.Drawing.Size(73, 16)
        Me.LabelConnected.TabIndex = 46
        Me.LabelConnected.Text = "Connected"
        Me.LabelConnected.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Line_ShapeSerial_T
        '
        Me.Line_ShapeSerial_T.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSerial_T.Location = New System.Drawing.Point(452, 16)
        Me.Line_ShapeSerial_T.Name = "Line_ShapeSerial_T"
        Me.Line_ShapeSerial_T.Size = New System.Drawing.Size(289, 1)
        '
        'Line_ShapeSerial_B
        '
        Me.Line_ShapeSerial_B.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSerial_B.Location = New System.Drawing.Point(452, 136)
        Me.Line_ShapeSerial_B.Name = "Line_ShapeSerial_B"
        Me.Line_ShapeSerial_B.Size = New System.Drawing.Size(289, 1)
        '
        'Line_ShapeSerial_L
        '
        Me.Line_ShapeSerial_L.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSerial_L.Location = New System.Drawing.Point(452, 16)
        Me.Line_ShapeSerial_L.Name = "Line_ShapeSerial_L"
        Me.Line_ShapeSerial_L.Size = New System.Drawing.Size(1, 121)
        '
        'Line_ShapeSerial_R
        '
        Me.Line_ShapeSerial_R.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeSerial_R.Location = New System.Drawing.Point(740, 16)
        Me.Line_ShapeSerial_R.Name = "Line_ShapeSerial_R"
        Me.Line_ShapeSerial_R.Size = New System.Drawing.Size(1, 121)
        '
        'Line_ShapeUsb_T
        '
        Me.Line_ShapeUsb_T.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeUsb_T.Location = New System.Drawing.Point(224, 16)
        Me.Line_ShapeUsb_T.Name = "Line_ShapeUsb_T"
        Me.Line_ShapeUsb_T.Size = New System.Drawing.Size(213, 1)
        '
        'Line_ShapeUsb_B
        '
        Me.Line_ShapeUsb_B.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeUsb_B.Location = New System.Drawing.Point(224, 136)
        Me.Line_ShapeUsb_B.Name = "Line_ShapeUsb_B"
        Me.Line_ShapeUsb_B.Size = New System.Drawing.Size(213, 1)
        '
        'Line_ShapeUsb_L
        '
        Me.Line_ShapeUsb_L.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeUsb_L.Location = New System.Drawing.Point(224, 16)
        Me.Line_ShapeUsb_L.Name = "Line_ShapeUsb_L"
        Me.Line_ShapeUsb_L.Size = New System.Drawing.Size(1, 121)
        '
        'Line_ShapeUsb_R
        '
        Me.Line_ShapeUsb_R.BackColor = System.Drawing.Color.Black
        Me.Line_ShapeUsb_R.Location = New System.Drawing.Point(436, 16)
        Me.Line_ShapeUsb_R.Name = "Line_ShapeUsb_R"
        Me.Line_ShapeUsb_R.Size = New System.Drawing.Size(1, 121)
        '
        'Line_ShapeTcp_T
        '
        Me.Line_ShapeTcp_T.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_T.Location = New System.Drawing.Point(8, 16)
        Me.Line_ShapeTcp_T.Name = "Line_ShapeTcp_T"
        Me.Line_ShapeTcp_T.Size = New System.Drawing.Size(201, 1)
        '
        'Line_ShapeTcp_B
        '
        Me.Line_ShapeTcp_B.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_B.Location = New System.Drawing.Point(8, 136)
        Me.Line_ShapeTcp_B.Name = "Line_ShapeTcp_B"
        Me.Line_ShapeTcp_B.Size = New System.Drawing.Size(201, 1)
        '
        'Line_ShapeTcp_L
        '
        Me.Line_ShapeTcp_L.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_L.Location = New System.Drawing.Point(8, 16)
        Me.Line_ShapeTcp_L.Name = "Line_ShapeTcp_L"
        Me.Line_ShapeTcp_L.Size = New System.Drawing.Size(1, 121)
        '
        'Line_ShapeTcp_R
        '
        Me.Line_ShapeTcp_R.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_R.Location = New System.Drawing.Point(208, 16)
        Me.Line_ShapeTcp_R.Name = "Line_ShapeTcp_R"
        Me.Line_ShapeTcp_R.Size = New System.Drawing.Size(1, 121)
        '
        'LabelBaudrate_0
        '
        Me.LabelBaudrate_0.Location = New System.Drawing.Point(462, 64)
        Me.LabelBaudrate_0.Name = "LabelBaudrate_0"
        Me.LabelBaudrate_0.AutoSize = False
        Me.LabelBaudrate_0.Size = New System.Drawing.Size(65, 17)
        Me.LabelBaudrate_0.TabIndex = 43
        Me.LabelBaudrate_0.Text = "Baudrate"
        Me.LabelBaudrate_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelSerialPort
        '
        Me.LabelSerialPort.Location = New System.Drawing.Point(456, 32)
        Me.LabelSerialPort.Name = "LabelSerialPort"
        Me.LabelSerialPort.AutoSize = False
        Me.LabelSerialPort.Size = New System.Drawing.Size(68, 17)
        Me.LabelSerialPort.TabIndex = 42
        Me.LabelSerialPort.Text = "Serial Port"
        Me.LabelSerialPort.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelUSBPort
        '
        Me.LabelUSBPort.Location = New System.Drawing.Point(227, 32)
        Me.LabelUSBPort.Name = "LabelUSBPort"
        Me.LabelUSBPort.AutoSize = False
        Me.LabelUSBPort.Size = New System.Drawing.Size(66, 17)
        Me.LabelUSBPort.TabIndex = 36
        Me.LabelUSBPort.Text = "USB Port"
        Me.LabelUSBPort.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelProgrammiermode
        '
        Me.LabelProgrammiermode.Location = New System.Drawing.Point(777, 114)
        Me.LabelProgrammiermode.Name = "LabelProgrammiermode"
        Me.LabelProgrammiermode.AutoSize = True
        Me.LabelProgrammiermode.Size = New System.Drawing.Size(124, 16)
        Me.LabelProgrammiermode.TabIndex = 35
        Me.LabelProgrammiermode.Text = "Programmier Mode"
        Me.LabelProgrammiermode.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelNotAus
        '
        Me.LabelNotAus.Location = New System.Drawing.Point(777, 50)
        Me.LabelNotAus.Name = "LabelNotAus"
        Me.LabelNotAus.AutoSize = True
        Me.LabelNotAus.Size = New System.Drawing.Size(52, 16)
        Me.LabelNotAus.TabIndex = 34
        Me.LabelNotAus.Text = "NotAus"
        Me.LabelNotAus.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelNotHalt
        '
        Me.LabelNotHalt.Location = New System.Drawing.Point(777, 82)
        Me.LabelNotHalt.Name = "LabelNotHalt"
        Me.LabelNotHalt.AutoSize = True
        Me.LabelNotHalt.Size = New System.Drawing.Size(53, 16)
        Me.LabelNotHalt.TabIndex = 33
        Me.LabelNotHalt.Text = "NotHalt"
        Me.LabelNotHalt.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'ledProgrammiermode
        '
        Me.ledProgrammiermode.Location = New System.Drawing.Point(755, 112)
        Me.ledProgrammiermode.Name = "ledProgrammiermode"
        Me.ledProgrammiermode.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledProgrammiermode.Round = True
        Me.ledProgrammiermode.Size = New System.Drawing.Size(17, 17)
        '
        'ledNotAus
        '
        Me.ledNotAus.Location = New System.Drawing.Point(755, 48)
        Me.ledNotAus.Name = "ledNotAus"
        Me.ledNotAus.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledNotAus.Round = True
        Me.ledNotAus.Size = New System.Drawing.Size(17, 17)
        '
        'ledNotHalt
        '
        Me.ledNotHalt.Location = New System.Drawing.Point(755, 80)
        Me.ledNotHalt.Name = "ledNotHalt"
        Me.ledNotHalt.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledNotHalt.Round = True
        Me.ledNotHalt.Size = New System.Drawing.Size(17, 17)
        '
        'txtMasterVersion
        '
        Me.txtMasterVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMasterVersion.Location = New System.Drawing.Point(1017, 16)
        Me.txtMasterVersion.Name = "txtMasterVersion"
        Me.txtMasterVersion.Size = New System.Drawing.Size(57, 17)
        Me.txtMasterVersion.TabIndex = 32
        Me.txtMasterVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelMasterVersion
        '
        Me.LabelMasterVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMasterVersion.Location = New System.Drawing.Point(929, 16)
        Me.LabelMasterVersion.Name = "LabelMasterVersion"
        Me.LabelMasterVersion.AutoSize = False
        Me.LabelMasterVersion.Size = New System.Drawing.Size(84, 15)
        Me.LabelMasterVersion.TabIndex = 31
        Me.LabelMasterVersion.Text = "Master Version:"
        Me.LabelMasterVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Line5
        '
        Me.Line5.BackColor = System.Drawing.Color.Black
        Me.Line5.Location = New System.Drawing.Point(560, 344)
        Me.Line5.Name = "Line5"
        Me.Line5.Size = New System.Drawing.Size(612, 2)
        '
        'Line4
        '
        Me.Line4.BackColor = System.Drawing.Color.Black
        Me.Line4.Location = New System.Drawing.Point(552, 184)
        Me.Line4.Name = "Line4"
        Me.Line4.Size = New System.Drawing.Size(2, 448)
        '
        'txtNmbPckts
        '
        Me.txtNmbPckts.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNmbPckts.Location = New System.Drawing.Point(812, 147)
        Me.txtNmbPckts.Name = "txtNmbPckts"
        Me.txtNmbPckts.Size = New System.Drawing.Size(57, 17)
        Me.txtNmbPckts.TabIndex = 23
        Me.txtNmbPckts.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtTcpFree
        '
        Me.txtTcpFree.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTcpFree.Location = New System.Drawing.Point(1017, 152)
        Me.txtTcpFree.Name = "txtTcpFree"
        Me.txtTcpFree.Size = New System.Drawing.Size(57, 15)
        Me.txtTcpFree.TabIndex = 22
        Me.txtTcpFree.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtXVersion
        '
        Me.txtXVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtXVersion.Location = New System.Drawing.Point(1017, 128)
        Me.txtXVersion.Name = "txtXVersion"
        Me.txtXVersion.Size = New System.Drawing.Size(57, 17)
        Me.txtXVersion.TabIndex = 21
        Me.txtXVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtXAddress
        '
        Me.txtXAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtXAddress.Location = New System.Drawing.Point(1017, 104)
        Me.txtXAddress.Name = "txtXAddress"
        Me.txtXAddress.Size = New System.Drawing.Size(57, 17)
        Me.txtXAddress.TabIndex = 20
        Me.txtXAddress.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtInterfaceCode
        '
        Me.txtInterfaceCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInterfaceCode.Location = New System.Drawing.Point(1017, 72)
        Me.txtInterfaceCode.Name = "txtInterfaceCode"
        Me.txtInterfaceCode.Size = New System.Drawing.Size(57, 17)
        Me.txtInterfaceCode.TabIndex = 19
        Me.txtInterfaceCode.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'txtInterfaceVersion
        '
        Me.txtInterfaceVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInterfaceVersion.Location = New System.Drawing.Point(1017, 48)
        Me.txtInterfaceVersion.Name = "txtInterfaceVersion"
        Me.txtInterfaceVersion.Size = New System.Drawing.Size(57, 17)
        Me.txtInterfaceVersion.TabIndex = 18
        Me.txtInterfaceVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Line3
        '
        Me.Line3.BackColor = System.Drawing.Color.Black
        Me.Line3.Location = New System.Drawing.Point(916, 96)
        Me.Line3.Name = "Line3"
        Me.Line3.Size = New System.Drawing.Size(248, 1)
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.Color.Black
        Me.Line1.Location = New System.Drawing.Point(8, 176)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(1164, 2)
        '
        'labelTcpFree
        '
        Me.labelTcpFree.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelTcpFree.Location = New System.Drawing.Point(914, 152)
        Me.labelTcpFree.Name = "labelTcpFree"
        Me.labelTcpFree.AutoSize = False
        Me.labelTcpFree.Size = New System.Drawing.Size(98, 15)
        Me.labelTcpFree.TabIndex = 16
        Me.labelTcpFree.Text = "Free TCP sockets:"
        Me.labelTcpFree.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'labelXVersion
        '
        Me.labelXVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelXVersion.Location = New System.Drawing.Point(913, 128)
        Me.labelXVersion.Name = "labelXVersion"
        Me.labelXVersion.AutoSize = False
        Me.labelXVersion.Size = New System.Drawing.Size(101, 15)
        Me.labelXVersion.TabIndex = 15
        Me.labelXVersion.Text = "XpressNet Version:"
        Me.labelXVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'labelXAddress
        '
        Me.labelXAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelXAddress.Location = New System.Drawing.Point(913, 104)
        Me.labelXAddress.Name = "labelXAddress"
        Me.labelXAddress.AutoSize = False
        Me.labelXAddress.Size = New System.Drawing.Size(104, 15)
        Me.labelXAddress.TabIndex = 14
        Me.labelXAddress.Text = "XpressNet Address:"
        Me.labelXAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'labelCode
        '
        Me.labelCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelCode.Location = New System.Drawing.Point(931, 72)
        Me.labelCode.Name = "labelCode"
        Me.labelCode.AutoSize = False
        Me.labelCode.Size = New System.Drawing.Size(82, 15)
        Me.labelCode.TabIndex = 13
        Me.labelCode.Text = "Interface Code:"
        Me.labelCode.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'labelInterfaceVersion
        '
        Me.labelInterfaceVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelInterfaceVersion.Location = New System.Drawing.Point(922, 48)
        Me.labelInterfaceVersion.Name = "labelInterfaceVersion"
        Me.labelInterfaceVersion.AutoSize = False
        Me.labelInterfaceVersion.Size = New System.Drawing.Size(93, 15)
        Me.labelInterfaceVersion.TabIndex = 12
        Me.labelInterfaceVersion.Text = "Interface Version:"
        Me.labelInterfaceVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelXpressNetMessageCounter
        '
        Me.LabelXpressNetMessageCounter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelXpressNetMessageCounter.Location = New System.Drawing.Point(640, 147)
        Me.LabelXpressNetMessageCounter.Name = "LabelXpressNetMessageCounter"
        Me.LabelXpressNetMessageCounter.AutoSize = False
        Me.LabelXpressNetMessageCounter.Size = New System.Drawing.Size(170, 15)
        Me.LabelXpressNetMessageCounter.TabIndex = 11
        Me.LabelXpressNetMessageCounter.Text = "Number of XpressNet Messages:"
        Me.LabelXpressNetMessageCounter.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelSwitch
        '
        Me.LabelSwitch.Location = New System.Drawing.Point(16, 208)
        Me.LabelSwitch.Name = "LabelSwitch"
        Me.LabelSwitch.AutoSize = True
        Me.LabelSwitch.Size = New System.Drawing.Size(49, 16)
        Me.LabelSwitch.TabIndex = 10
        Me.LabelSwitch.Text = "Switch:"
        Me.LabelSwitch.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.Color.Black
        Me.Line2.Location = New System.Drawing.Point(392, 176)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(432, 2)
        '
        'LabelStatus
        '
        Me.LabelStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStatus.Location = New System.Drawing.Point(8, 147)
        Me.LabelStatus.Name = "LabelStatus"
        Me.LabelStatus.Size = New System.Drawing.Size(545, 21)
        Me.LabelStatus.TabIndex = 6
        Me.LabelStatus.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'ledInterfaceStatus
        '
        Me.ledInterfaceStatus.Location = New System.Drawing.Point(755, 16)
        Me.ledInterfaceStatus.Name = "ledInterfaceStatus"
        Me.ledInterfaceStatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledInterfaceStatus.Round = True
        Me.ledInterfaceStatus.Size = New System.Drawing.Size(17, 17)
        '
        'LabelTcpPort
        '
        Me.LabelTcpPort.Location = New System.Drawing.Point(20, 64)
        Me.LabelTcpPort.Name = "LabelTcpPort"
        Me.LabelTcpPort.AutoSize = False
        Me.LabelTcpPort.Size = New System.Drawing.Size(66, 17)
        Me.LabelTcpPort.TabIndex = 5
        Me.LabelTcpPort.Text = "TCP Port"
        Me.LabelTcpPort.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelIpAddress
        '
        Me.LabelIpAddress.Location = New System.Drawing.Point(9, 32)
        Me.LabelIpAddress.Name = "LabelIpAddress"
        Me.LabelIpAddress.AutoSize = False
        Me.LabelIpAddress.Size = New System.Drawing.Size(76, 17)
        Me.LabelIpAddress.TabIndex = 4
        Me.LabelIpAddress.Text = "IP Address"
        Me.LabelIpAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeLabelInfo1
        '
        Me.LabeLabelInfo1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeLabelInfo1.Location = New System.Drawing.Point(72, 592)
        Me.LabeLabelInfo1.Name = "LabeLabelInfo1"
        Me.LabeLabelInfo1.AutoSize = True
        Me.LabeLabelInfo1.Size = New System.Drawing.Size(136, 16)
        Me.LabeLabelInfo1.TabIndex = 1
        Me.LabeLabelInfo1.Text = "latest message at top"
        Me.LabeLabelInfo1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ToolTip1.SetToolTip(Me.ComboBaudUsb, "Lenz: 57600 / Yamorc: 115200")
        Me.ToolTip1.SetToolTip(Me.TxtRsBusAddress, "1..128")
        Me.ToolTip1.SetToolTip(Me.TxtBoxSwitchTtBits, "C=Command Station, D=Decoder")
        Me.ToolTip1.SetToolTip(Me.txtWissel, "1..2048")
        Me.ToolTip1.SetToolTip(Me.cmdTcpConnect, "Schakel wissel 1 om")
        Me.ToolTip1.SetToolTip(Me.LabelBaudrateUSB_1, "Lenz: 57600 / Yamorc: 115200")
        '
        'FrmExtendedTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1190, 643)
        Me.Controls.Add(Me.CheckHandshake)
        Me.Controls.Add(Me.ComboBaudUsb)
        Me.Controls.Add(Me.CmdExtAccSet)
        Me.Controls.Add(Me.TxtExtAccValue)
        Me.Controls.Add(Me.TxtExtAccSwitch)
        Me.Controls.Add(Me.CmdLokAusStack)
        Me.Controls.Add(Me.CmdStartMan)
        Me.Controls.Add(Me.SmPage)
        Me.Controls.Add(Me.SmRegister)
        Me.Controls.Add(Me.SmDirect)
        Me.Controls.Add(Me.CheckSM)
        Me.Controls.Add(Me.CmdLokEnter)
        Me.Controls.Add(Me.CmdGetF0_F12)
        Me.Controls.Add(Me.CmdFunctionOff)
        Me.Controls.Add(Me.CmdFunctionOn)
        Me.Controls.Add(Me.CmdGetF29_68)
        Me.Controls.Add(Me.CmdGetF13_28)
        Me.Controls.Add(Me.CmdPomSet)
        Me.Controls.Add(Me.CmdPomGet)
        Me.Controls.Add(Me.TxtPomValue)
        Me.Controls.Add(Me.TxtPomCV)
        Me.Controls.Add(Me.TxtPomAddress)
        Me.Controls.Add(Me.TxtLokFNumber)
        Me.Controls.Add(Me.CheckAutoDeactivate)
        Me.Controls.Add(Me.CheckActivate)
        Me.Controls.Add(Me.CmdLokForward)
        Me.Controls.Add(Me.CmdLokReverse)
        Me.Controls.Add(Me.ComboLokSteps)
        Me.Controls.Add(Me.TxtLokSpeed)
        Me.Controls.Add(Me.ComboDay)
        Me.Controls.Add(Me.TxtHour)
        Me.Controls.Add(Me.TxtMinute)
        Me.Controls.Add(Me.TxtFactor)
        Me.Controls.Add(Me.CheckRun)
        Me.Controls.Add(Me.CmdModellzeitStellen)
        Me.Controls.Add(Me.CmdModellzeitAnfordern)
        Me.Controls.Add(Me.CmdModellzeitAnhalten)
        Me.Controls.Add(Me.CmdModellzeitStarten)
        Me.Controls.Add(Me.cmdStartAuto)
        Me.Controls.Add(Me.CmdErweiterteVersion)
        Me.Controls.Add(Me.CmdLanInfo)
        Me.Controls.Add(Me.CmdRsFeedback)
        Me.Controls.Add(Me.TxtRsBusAddress)
        Me.Controls.Add(Me.TxtBoxRsNibble)
        Me.Controls.Add(Me.TxtBoxRsData)
        Me.Controls.Add(Me.TxBoxtRsAddress)
        Me.Controls.Add(Me.TxtBoxSwitchIBit)
        Me.Controls.Add(Me.TxtBoxSwitchTtBits)
        Me.Controls.Add(Me.TxtBoxSwitchData)
        Me.Controls.Add(Me.TxtBoxSwitchAddress)
        Me.Controls.Add(Me.CmdSerialConnect)
        Me.Controls.Add(Me.CheckCts)
        Me.Controls.Add(Me.TxtSerialPort)
        Me.Controls.Add(Me.ComboBaud)
        Me.Controls.Add(Me.Cmdtest)
        Me.Controls.Add(Me.CmdUsbConnect)
        Me.Controls.Add(Me.TxtUsbPort)
        Me.Controls.Add(Me.CmdInterfaceInfo)
        Me.Controls.Add(Me.cmdPomHolen)
        Me.Controls.Add(Me.txtLokAddress)
        Me.Controls.Add(Me.CmdEineLokAnhalten)
        Me.Controls.Add(Me.CmdAlleLoksAnhalten)
        Me.Controls.Add(Me.CmdAllesAus)
        Me.Controls.Add(Me.CmdAllesAn)
        Me.Controls.Add(Me.cmdSwitchFeedback)
        Me.Controls.Add(Me.txtWissel)
        Me.Controls.Add(Me.cmdAf)
        Me.Controls.Add(Me.cmdRecht)
        Me.Controls.Add(Me.txtTcpPort)
        Me.Controls.Add(Me.txtIP)
        Me.Controls.Add(Me.cmdTcpConnect)
        Me.Controls.Add(Me.Line9)
        Me.Controls.Add(Me.Line8)
        Me.Controls.Add(Me.Line_ShapeAnAus_T)
        Me.Controls.Add(Me.Line_ShapeAnAus_B)
        Me.Controls.Add(Me.Line_ShapeAnAus_L)
        Me.Controls.Add(Me.Line_ShapeAnAus_R)
        Me.Controls.Add(Me.Line_Shape4_T)
        Me.Controls.Add(Me.Line_Shape4_B)
        Me.Controls.Add(Me.Line_Shape4_L)
        Me.Controls.Add(Me.Line_Shape4_R)
        Me.Controls.Add(Me.Line_Shape3_T)
        Me.Controls.Add(Me.Line_Shape3_B)
        Me.Controls.Add(Me.Line_Shape3_L)
        Me.Controls.Add(Me.Line_Shape3_R)
        Me.Controls.Add(Me.Line7)
        Me.Controls.Add(Me.Line_Shape2_T)
        Me.Controls.Add(Me.Line_Shape2_B)
        Me.Controls.Add(Me.Line_Shape2_L)
        Me.Controls.Add(Me.Line_Shape2_R)
        Me.Controls.Add(Me.Line6)
        Me.Controls.Add(Me.Line_Shape1_T)
        Me.Controls.Add(Me.Line_Shape1_B)
        Me.Controls.Add(Me.Line_Shape1_L)
        Me.Controls.Add(Me.Line_Shape1_R)
        Me.Controls.Add(Me.Line_ShapeFeedback_T)
        Me.Controls.Add(Me.Line_ShapeFeedback_B)
        Me.Controls.Add(Me.Line_ShapeFeedback_L)
        Me.Controls.Add(Me.Line_ShapeFeedback_R)
        Me.Controls.Add(Me.Line_ShapeSwitch_T)
        Me.Controls.Add(Me.Line_ShapeSwitch_B)
        Me.Controls.Add(Me.Line_ShapeSwitch_L)
        Me.Controls.Add(Me.Line_ShapeSwitch_R)
        Me.Controls.Add(Me.Line_ShapeInfo_T)
        Me.Controls.Add(Me.Line_ShapeInfo_B)
        Me.Controls.Add(Me.Line_ShapeInfo_L)
        Me.Controls.Add(Me.Line_ShapeInfo_R)
        Me.Controls.Add(Me.Line_ShapeSerial_T)
        Me.Controls.Add(Me.Line_ShapeSerial_B)
        Me.Controls.Add(Me.Line_ShapeSerial_L)
        Me.Controls.Add(Me.Line_ShapeSerial_R)
        Me.Controls.Add(Me.Line_ShapeUsb_T)
        Me.Controls.Add(Me.Line_ShapeUsb_B)
        Me.Controls.Add(Me.Line_ShapeUsb_L)
        Me.Controls.Add(Me.Line_ShapeUsb_R)
        Me.Controls.Add(Me.Line_ShapeTcp_T)
        Me.Controls.Add(Me.Line_ShapeTcp_B)
        Me.Controls.Add(Me.Line_ShapeTcp_L)
        Me.Controls.Add(Me.Line_ShapeTcp_R)
        Me.Controls.Add(Me.Line5)
        Me.Controls.Add(Me.Line4)
        Me.Controls.Add(Me.Line3)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.LabelBaudrateUSB_1)
        Me.Controls.Add(Me.LabelExtValue)
        Me.Controls.Add(Me.LabelExtSwitch)
        Me.Controls.Add(Me.LabelExtAccCmd)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LabelFunctions_8)
        Me.Controls.Add(Me.LabelFunctions_7)
        Me.Controls.Add(Me.LabelFunctions_6)
        Me.Controls.Add(Me.LabelFunctions_5)
        Me.Controls.Add(Me.LabelFunctions_4)
        Me.Controls.Add(Me.LabelFunctions_3)
        Me.Controls.Add(Me.LabelFunctions_2)
        Me.Controls.Add(Me.LabelFunctions_1)
        Me.Controls.Add(Me.LabelFunctions_0)
        Me.Controls.Add(Me.LabelPom)
        Me.Controls.Add(Me.LabelPomValue)
        Me.Controls.Add(Me.LabelPomCV)
        Me.Controls.Add(Me.LabelPomAddress)
        Me.Controls.Add(Me.LabelFunktion)
        Me.Controls.Add(Me.LabelLoks)
        Me.Controls.Add(Me.LabelLokSteps)
        Me.Controls.Add(Me.LabelLokSpeed)
        Me.Controls.Add(Me.LabelLokAddress)
        Me.Controls.Add(Me.LabelModellzeit)
        Me.Controls.Add(Me.LabelDay)
        Me.Controls.Add(Me.LabelHour)
        Me.Controls.Add(Me.LabelMinute)
        Me.Controls.Add(Me.LabelFacor)
        Me.Controls.Add(Me.TxtMinutes)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtHours)
        Me.Controls.Add(Me.txtVersionBootloader)
        Me.Controls.Add(Me.txtBuildRuckmelder)
        Me.Controls.Add(Me.TxtVersionRuckmelder)
        Me.Controls.Add(Me.txtBuildNr)
        Me.Controls.Add(Me.LabelVersionBootloader)
        Me.Controls.Add(Me.LabelBuildruckmeldedecoder)
        Me.Controls.Add(Me.LabelVersionRuckmeldedecoder)
        Me.Controls.Add(Me.LabelBuildNr)
        Me.Controls.Add(Me.LabelInfo2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LabelRsBusAddress)
        Me.Controls.Add(Me.LabelSwitchIBit)
        Me.Controls.Add(Me.LabelSwitchTtBits)
        Me.Controls.Add(Me.LabelSwitchdata)
        Me.Controls.Add(Me.LabelSwitchAddress)
        Me.Controls.Add(Me.LabelConnected)
        Me.Controls.Add(Me.LabelBaudrate_0)
        Me.Controls.Add(Me.LabelSerialPort)
        Me.Controls.Add(Me.LabelUSBPort)
        Me.Controls.Add(Me.LabelProgrammiermode)
        Me.Controls.Add(Me.LabelNotAus)
        Me.Controls.Add(Me.LabelNotHalt)
        Me.Controls.Add(Me.txtMasterVersion)
        Me.Controls.Add(Me.LabelMasterVersion)
        Me.Controls.Add(Me.txtNmbPckts)
        Me.Controls.Add(Me.txtTcpFree)
        Me.Controls.Add(Me.txtXVersion)
        Me.Controls.Add(Me.txtXAddress)
        Me.Controls.Add(Me.txtInterfaceCode)
        Me.Controls.Add(Me.txtInterfaceVersion)
        Me.Controls.Add(Me.labelTcpFree)
        Me.Controls.Add(Me.labelXVersion)
        Me.Controls.Add(Me.labelXAddress)
        Me.Controls.Add(Me.labelCode)
        Me.Controls.Add(Me.labelInterfaceVersion)
        Me.Controls.Add(Me.LabelXpressNetMessageCounter)
        Me.Controls.Add(Me.LabelSwitch)
        Me.Controls.Add(Me.LabelStatus)
        Me.Controls.Add(Me.LabelTcpPort)
        Me.Controls.Add(Me.LabelIpAddress)
        Me.Controls.Add(Me.LabeLabelInfo1)
        Me.Controls.Add(Me.LEDF0F68_68)
        Me.Controls.Add(Me.LEDF0F68_67)
        Me.Controls.Add(Me.LEDF0F68_66)
        Me.Controls.Add(Me.LEDF0F68_65)
        Me.Controls.Add(Me.LEDF0F68_64)
        Me.Controls.Add(Me.LEDF0F68_63)
        Me.Controls.Add(Me.LEDF0F68_62)
        Me.Controls.Add(Me.LEDF0F68_61)
        Me.Controls.Add(Me.LEDF0F68_60)
        Me.Controls.Add(Me.LEDF0F68_59)
        Me.Controls.Add(Me.LEDF0F68_58)
        Me.Controls.Add(Me.LEDF0F68_57)
        Me.Controls.Add(Me.LEDF0F68_56)
        Me.Controls.Add(Me.LEDF0F68_55)
        Me.Controls.Add(Me.LEDF0F68_54)
        Me.Controls.Add(Me.LEDF0F68_53)
        Me.Controls.Add(Me.LEDF0F68_52)
        Me.Controls.Add(Me.LEDF0F68_51)
        Me.Controls.Add(Me.LEDF0F68_50)
        Me.Controls.Add(Me.LEDF0F68_49)
        Me.Controls.Add(Me.LEDF0F68_48)
        Me.Controls.Add(Me.LEDF0F68_47)
        Me.Controls.Add(Me.LEDF0F68_46)
        Me.Controls.Add(Me.LEDF0F68_45)
        Me.Controls.Add(Me.LEDF0F68_44)
        Me.Controls.Add(Me.LEDF0F68_43)
        Me.Controls.Add(Me.LEDF0F68_42)
        Me.Controls.Add(Me.LEDF0F68_41)
        Me.Controls.Add(Me.LEDF0F68_40)
        Me.Controls.Add(Me.LEDF0F68_39)
        Me.Controls.Add(Me.LEDF0F68_38)
        Me.Controls.Add(Me.LEDF0F68_37)
        Me.Controls.Add(Me.LEDF0F68_36)
        Me.Controls.Add(Me.LEDF0F68_35)
        Me.Controls.Add(Me.LEDF0F68_34)
        Me.Controls.Add(Me.LEDF0F68_33)
        Me.Controls.Add(Me.LEDF0F68_32)
        Me.Controls.Add(Me.LEDF0F68_31)
        Me.Controls.Add(Me.LEDF0F68_30)
        Me.Controls.Add(Me.LEDF0F68_29)
        Me.Controls.Add(Me.LEDF0F68_28)
        Me.Controls.Add(Me.LEDF0F68_27)
        Me.Controls.Add(Me.LEDF0F68_26)
        Me.Controls.Add(Me.LEDF0F68_25)
        Me.Controls.Add(Me.LEDF0F68_24)
        Me.Controls.Add(Me.LEDF0F68_23)
        Me.Controls.Add(Me.LEDF0F68_22)
        Me.Controls.Add(Me.LEDF0F68_21)
        Me.Controls.Add(Me.LEDF0F68_20)
        Me.Controls.Add(Me.LEDF0F68_19)
        Me.Controls.Add(Me.LEDF0F68_18)
        Me.Controls.Add(Me.LEDF0F68_17)
        Me.Controls.Add(Me.LEDF0F68_16)
        Me.Controls.Add(Me.LEDF0F68_15)
        Me.Controls.Add(Me.LEDF0F68_14)
        Me.Controls.Add(Me.LEDF0F68_13)
        Me.Controls.Add(Me.LEDF0F68_12)
        Me.Controls.Add(Me.LEDF0F68_11)
        Me.Controls.Add(Me.LEDF0F68_10)
        Me.Controls.Add(Me.LEDF0F68_9)
        Me.Controls.Add(Me.LEDF0F68_8)
        Me.Controls.Add(Me.LEDF0F68_7)
        Me.Controls.Add(Me.LEDF0F68_6)
        Me.Controls.Add(Me.LEDF0F68_5)
        Me.Controls.Add(Me.LEDF0F68_4)
        Me.Controls.Add(Me.LEDF0F68_3)
        Me.Controls.Add(Me.LEDF0F68_2)
        Me.Controls.Add(Me.LEDF0F68_1)
        Me.Controls.Add(Me.LEDF0F68_0)
        Me.Controls.Add(Me.ledProgrammiermode)
        Me.Controls.Add(Me.ledNotAus)
        Me.Controls.Add(Me.ledNotHalt)
        Me.Controls.Add(Me.ledInterfaceStatus)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "FrmExtendedTest"
        Me.Text = "Test XPressNet"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBaudUsb As System.Windows.Forms.ComboBox
    Friend WithEvents CmdExtAccSet As System.Windows.Forms.Button
    Friend WithEvents TxtExtAccValue As System.Windows.Forms.TextBox
    Friend WithEvents TxtExtAccSwitch As System.Windows.Forms.TextBox
    Friend WithEvents CmdLokAusStack As System.Windows.Forms.Button
    Friend WithEvents CmdStartMan As System.Windows.Forms.Button
    Friend WithEvents SmPage As System.Windows.Forms.RadioButton
    Friend WithEvents SmRegister As System.Windows.Forms.RadioButton
    Friend WithEvents SmDirect As System.Windows.Forms.RadioButton
    Friend WithEvents CheckSM As System.Windows.Forms.CheckBox
    Friend WithEvents CmdLokEnter As System.Windows.Forms.Button
    Friend WithEvents CmdGetF0_F12 As System.Windows.Forms.Button
    Friend WithEvents CmdFunctionOff As System.Windows.Forms.Button
    Friend WithEvents CmdFunctionOn As System.Windows.Forms.Button
    Friend WithEvents CmdGetF29_68 As System.Windows.Forms.Button
    Friend WithEvents CmdGetF13_28 As System.Windows.Forms.Button
    Friend WithEvents CmdPomSet As System.Windows.Forms.Button
    Friend WithEvents CmdPomGet As System.Windows.Forms.Button
    Friend WithEvents TxtPomValue As System.Windows.Forms.TextBox
    Friend WithEvents TxtPomCV As System.Windows.Forms.TextBox
    Friend WithEvents TxtPomAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtLokFNumber As System.Windows.Forms.TextBox
    Friend WithEvents CheckAutoDeactivate As System.Windows.Forms.CheckBox
    Friend WithEvents CheckActivate As System.Windows.Forms.CheckBox
    Friend WithEvents CmdLokForward As System.Windows.Forms.Button
    Friend WithEvents CmdLokReverse As System.Windows.Forms.Button
    Friend WithEvents ComboLokSteps As System.Windows.Forms.ComboBox
    Friend WithEvents TxtLokSpeed As System.Windows.Forms.TextBox
    Friend WithEvents ComboDay As System.Windows.Forms.ComboBox
    Friend WithEvents TxtHour As System.Windows.Forms.TextBox
    Friend WithEvents TxtMinute As System.Windows.Forms.TextBox
    Friend WithEvents TxtFactor As System.Windows.Forms.TextBox
    Friend WithEvents CheckRun As System.Windows.Forms.CheckBox
    Friend WithEvents CmdModellzeitStellen As System.Windows.Forms.Button
    Friend WithEvents CmdModellzeitAnfordern As System.Windows.Forms.Button
    Friend WithEvents CmdModellzeitAnhalten As System.Windows.Forms.Button
    Friend WithEvents CmdModellzeitStarten As System.Windows.Forms.Button
    Friend WithEvents cmdStartAuto As System.Windows.Forms.Button
    Friend WithEvents CmdErweiterteVersion As System.Windows.Forms.Button
    Friend WithEvents CmdLanInfo As System.Windows.Forms.Button
    Friend WithEvents TimerMessageCounter As System.Windows.Forms.Timer
    Friend WithEvents CmdRsFeedback As System.Windows.Forms.Button
    Friend WithEvents TxtRsBusAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxRsNibble As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxRsData As System.Windows.Forms.TextBox
    Friend WithEvents TxBoxtRsAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchIBit As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchTtBits As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchData As System.Windows.Forms.TextBox
    Friend WithEvents TxtBoxSwitchAddress As System.Windows.Forms.TextBox
    Friend WithEvents CmdSerialConnect As System.Windows.Forms.Button
    Friend WithEvents CheckCts As System.Windows.Forms.CheckBox
    Friend WithEvents TxtSerialPort As System.Windows.Forms.TextBox
    Friend WithEvents ComboBaud As System.Windows.Forms.ComboBox
    Friend WithEvents Cmdtest As System.Windows.Forms.Button
    Friend WithEvents CmdUsbConnect As System.Windows.Forms.Button
    Friend WithEvents TxtUsbPort As System.Windows.Forms.TextBox
    Friend WithEvents CmdInterfaceInfo As System.Windows.Forms.Button
    Friend WithEvents cmdPomHolen As System.Windows.Forms.Button
    Friend WithEvents txtLokAddress As System.Windows.Forms.TextBox
    Friend WithEvents CmdEineLokAnhalten As System.Windows.Forms.Button
    Friend WithEvents CmdAlleLoksAnhalten As System.Windows.Forms.Button
    Friend WithEvents CmdAllesAus As System.Windows.Forms.Button
    Friend WithEvents CmdAllesAn As System.Windows.Forms.Button
    Friend WithEvents cmdSwitchFeedback As System.Windows.Forms.Button
    Friend WithEvents txtWissel As System.Windows.Forms.TextBox
    Friend WithEvents cmdAf As System.Windows.Forms.Button
    Friend WithEvents cmdRecht As System.Windows.Forms.Button
    Friend WithEvents txtTcpPort As System.Windows.Forms.TextBox
    Friend WithEvents txtIP As System.Windows.Forms.TextBox
    Friend WithEvents cmdTcpConnect As System.Windows.Forms.Button
    Friend WithEvents LabelBaudrateUSB_1 As System.Windows.Forms.Label
    Friend WithEvents Line9 As System.Windows.Forms.Label
    Friend WithEvents LabelExtValue As System.Windows.Forms.Label
    Friend WithEvents LabelExtSwitch As System.Windows.Forms.Label
    Friend WithEvents LabelExtAccCmd As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Line8 As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_R As System.Windows.Forms.Label
    Friend WithEvents LEDF0F68_68 As Led
    Friend WithEvents LEDF0F68_67 As Led
    Friend WithEvents LEDF0F68_66 As Led
    Friend WithEvents LEDF0F68_65 As Led
    Friend WithEvents LEDF0F68_64 As Led
    Friend WithEvents LEDF0F68_63 As Led
    Friend WithEvents LEDF0F68_62 As Led
    Friend WithEvents LEDF0F68_61 As Led
    Friend WithEvents LabelFunctions_8 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_7 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_6 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_5 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_4 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_3 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_2 As System.Windows.Forms.Label
    Friend WithEvents LabelFunctions_1 As System.Windows.Forms.Label
    Friend WithEvents LEDF0F68_60 As Led
    Friend WithEvents LEDF0F68_59 As Led
    Friend WithEvents LEDF0F68_58 As Led
    Friend WithEvents LEDF0F68_57 As Led
    Friend WithEvents LEDF0F68_56 As Led
    Friend WithEvents LEDF0F68_55 As Led
    Friend WithEvents LEDF0F68_54 As Led
    Friend WithEvents LEDF0F68_53 As Led
    Friend WithEvents LEDF0F68_52 As Led
    Friend WithEvents LEDF0F68_51 As Led
    Friend WithEvents LEDF0F68_50 As Led
    Friend WithEvents LEDF0F68_49 As Led
    Friend WithEvents LEDF0F68_48 As Led
    Friend WithEvents LEDF0F68_47 As Led
    Friend WithEvents LEDF0F68_46 As Led
    Friend WithEvents LEDF0F68_45 As Led
    Friend WithEvents LEDF0F68_44 As Led
    Friend WithEvents LEDF0F68_43 As Led
    Friend WithEvents LEDF0F68_42 As Led
    Friend WithEvents LEDF0F68_41 As Led
    Friend WithEvents LEDF0F68_40 As Led
    Friend WithEvents LEDF0F68_39 As Led
    Friend WithEvents LEDF0F68_38 As Led
    Friend WithEvents LEDF0F68_37 As Led
    Friend WithEvents LEDF0F68_36 As Led
    Friend WithEvents LEDF0F68_35 As Led
    Friend WithEvents LEDF0F68_34 As Led
    Friend WithEvents LEDF0F68_33 As Led
    Friend WithEvents LEDF0F68_32 As Led
    Friend WithEvents LEDF0F68_31 As Led
    Friend WithEvents LEDF0F68_30 As Led
    Friend WithEvents LEDF0F68_29 As Led
    Friend WithEvents LEDF0F68_28 As Led
    Friend WithEvents LEDF0F68_27 As Led
    Friend WithEvents LEDF0F68_26 As Led
    Friend WithEvents LEDF0F68_25 As Led
    Friend WithEvents LEDF0F68_24 As Led
    Friend WithEvents LEDF0F68_23 As Led
    Friend WithEvents LEDF0F68_22 As Led
    Friend WithEvents LEDF0F68_21 As Led
    Friend WithEvents LEDF0F68_20 As Led
    Friend WithEvents LEDF0F68_19 As Led
    Friend WithEvents LEDF0F68_18 As Led
    Friend WithEvents LEDF0F68_17 As Led
    Friend WithEvents LEDF0F68_16 As Led
    Friend WithEvents LEDF0F68_15 As Led
    Friend WithEvents LEDF0F68_14 As Led
    Friend WithEvents LEDF0F68_13 As Led
    Friend WithEvents LEDF0F68_12 As Led
    Friend WithEvents LEDF0F68_11 As Led
    Friend WithEvents LEDF0F68_10 As Led
    Friend WithEvents LEDF0F68_9 As Led
    Friend WithEvents LEDF0F68_8 As Led
    Friend WithEvents LEDF0F68_7 As Led
    Friend WithEvents LEDF0F68_6 As Led
    Friend WithEvents LEDF0F68_5 As Led
    Friend WithEvents LabelFunctions_0 As System.Windows.Forms.Label
    Friend WithEvents LEDF0F68_4 As Led
    Friend WithEvents LEDF0F68_3 As Led
    Friend WithEvents LEDF0F68_2 As Led
    Friend WithEvents LEDF0F68_1 As Led
    Friend WithEvents LEDF0F68_0 As Led
    Friend WithEvents Line_Shape4_T As System.Windows.Forms.Label
    Friend WithEvents Line_Shape4_B As System.Windows.Forms.Label
    Friend WithEvents Line_Shape4_L As System.Windows.Forms.Label
    Friend WithEvents Line_Shape4_R As System.Windows.Forms.Label
    Friend WithEvents LabelPom As System.Windows.Forms.Label
    Friend WithEvents LabelPomValue As System.Windows.Forms.Label
    Friend WithEvents LabelPomCV As System.Windows.Forms.Label
    Friend WithEvents LabelPomAddress As System.Windows.Forms.Label
    Friend WithEvents LabelFunktion As System.Windows.Forms.Label
    Friend WithEvents LabelLoks As System.Windows.Forms.Label
    Friend WithEvents LabelLokSteps As System.Windows.Forms.Label
    Friend WithEvents LabelLokSpeed As System.Windows.Forms.Label
    Friend WithEvents LabelLokAddress As System.Windows.Forms.Label
    Friend WithEvents LabelModellzeit As System.Windows.Forms.Label
    Friend WithEvents Line_Shape3_T As System.Windows.Forms.Label
    Friend WithEvents Line_Shape3_B As System.Windows.Forms.Label
    Friend WithEvents Line_Shape3_L As System.Windows.Forms.Label
    Friend WithEvents Line_Shape3_R As System.Windows.Forms.Label
    Friend WithEvents Line7 As System.Windows.Forms.Label
    Friend WithEvents Line_Shape2_T As System.Windows.Forms.Label
    Friend WithEvents Line_Shape2_B As System.Windows.Forms.Label
    Friend WithEvents Line_Shape2_L As System.Windows.Forms.Label
    Friend WithEvents Line_Shape2_R As System.Windows.Forms.Label
    Friend WithEvents LabelDay As System.Windows.Forms.Label
    Friend WithEvents LabelHour As System.Windows.Forms.Label
    Friend WithEvents LabelMinute As System.Windows.Forms.Label
    Friend WithEvents LabelFacor As System.Windows.Forms.Label
    Friend WithEvents TxtMinutes As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtHours As System.Windows.Forms.Label
    Friend WithEvents txtVersionBootloader As System.Windows.Forms.Label
    Friend WithEvents txtBuildRuckmelder As System.Windows.Forms.Label
    Friend WithEvents TxtVersionRuckmelder As System.Windows.Forms.Label
    Friend WithEvents txtBuildNr As System.Windows.Forms.Label
    Friend WithEvents LabelVersionBootloader As System.Windows.Forms.Label
    Friend WithEvents LabelBuildruckmeldedecoder As System.Windows.Forms.Label
    Friend WithEvents LabelVersionRuckmeldedecoder As System.Windows.Forms.Label
    Friend WithEvents LabelBuildNr As System.Windows.Forms.Label
    Friend WithEvents Line6 As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_T As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_B As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_L As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_R As System.Windows.Forms.Label
    Friend WithEvents LabelInfo2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LabelRsBusAddress As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeFeedback_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeFeedback_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeFeedback_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeFeedback_R As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchIBit As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchTtBits As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchdata As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchAddress As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSwitch_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSwitch_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSwitch_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSwitch_R As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeInfo_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeInfo_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeInfo_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeInfo_R As System.Windows.Forms.Label
    Friend WithEvents LabelConnected As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSerial_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSerial_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSerial_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeSerial_R As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeUsb_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeUsb_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeUsb_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeUsb_R As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_R As System.Windows.Forms.Label
    Friend WithEvents LabelBaudrate_0 As System.Windows.Forms.Label
    Friend WithEvents LabelSerialPort As System.Windows.Forms.Label
    Friend WithEvents LabelUSBPort As System.Windows.Forms.Label
    Friend WithEvents LabelProgrammiermode As System.Windows.Forms.Label
    Friend WithEvents LabelNotAus As System.Windows.Forms.Label
    Friend WithEvents LabelNotHalt As System.Windows.Forms.Label
    Friend WithEvents ledProgrammiermode As Led
    Friend WithEvents ledNotAus As Led
    Friend WithEvents ledNotHalt As Led
    Friend WithEvents txtMasterVersion As System.Windows.Forms.Label
    Friend WithEvents LabelMasterVersion As System.Windows.Forms.Label
    Friend WithEvents Line5 As System.Windows.Forms.Label
    Friend WithEvents Line4 As System.Windows.Forms.Label
    Friend WithEvents txtNmbPckts As System.Windows.Forms.Label
    Friend WithEvents txtTcpFree As System.Windows.Forms.Label
    Friend WithEvents txtXVersion As System.Windows.Forms.Label
    Friend WithEvents txtXAddress As System.Windows.Forms.Label
    Friend WithEvents txtInterfaceCode As System.Windows.Forms.Label
    Friend WithEvents txtInterfaceVersion As System.Windows.Forms.Label
    Friend WithEvents Line3 As System.Windows.Forms.Label
    Friend WithEvents Line1 As System.Windows.Forms.Label
    Friend WithEvents labelTcpFree As System.Windows.Forms.Label
    Friend WithEvents labelXVersion As System.Windows.Forms.Label
    Friend WithEvents labelXAddress As System.Windows.Forms.Label
    Friend WithEvents labelCode As System.Windows.Forms.Label
    Friend WithEvents labelInterfaceVersion As System.Windows.Forms.Label
    Friend WithEvents LabelXpressNetMessageCounter As System.Windows.Forms.Label
    Friend WithEvents LabelSwitch As System.Windows.Forms.Label
    Friend WithEvents Line2 As System.Windows.Forms.Label
    Friend WithEvents LabelStatus As System.Windows.Forms.Label
    Friend WithEvents ledInterfaceStatus As Led
    Friend WithEvents LabelTcpPort As System.Windows.Forms.Label
    Friend WithEvents LabelIpAddress As System.Windows.Forms.Label
    Friend WithEvents LabeLabelInfo1 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CheckHandshake As System.Windows.Forms.CheckBox
End Class
