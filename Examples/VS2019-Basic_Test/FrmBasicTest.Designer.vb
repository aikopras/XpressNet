<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmBasicTest
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
        Me.GbTcpConnect = New System.Windows.Forms.GroupBox()
        Me.BtnConnectTcpIp = New System.Windows.Forms.Button()
        Me.TxtTcpPort = New System.Windows.Forms.TextBox()
        Me.TxtIpAddress = New System.Windows.Forms.TextBox()
        Me.LbTcpPort = New System.Windows.Forms.Label()
        Me.LbIpAddress = New System.Windows.Forms.Label()
        Me.GbConnectUsb = New System.Windows.Forms.GroupBox()
        Me.CbBaudrate = New System.Windows.Forms.ComboBox()
        Me.BtnConnectUsb = New System.Windows.Forms.Button()
        Me.TxtUsbPort = New System.Windows.Forms.TextBox()
        Me.LbBaudrate = New System.Windows.Forms.Label()
        Me.LbUsbPort = New System.Windows.Forms.Label()
        Me.GbConnectInfo = New System.Windows.Forms.GroupBox()
        Me.LbNotHalt = New System.Windows.Forms.Label()
        Me.LbNotAus = New System.Windows.Forms.Label()
        Me.LbConnected = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.ShNotHalt = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        Me.ShNotAus = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        Me.ShConnected = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        Me.GbXpressNetVersion = New System.Windows.Forms.GroupBox()
        Me.BtnLanInfo = New System.Windows.Forms.Button()
        Me.TxtFreeTcpSockets = New System.Windows.Forms.TextBox()
        Me.TxtXpressNetVersion = New System.Windows.Forms.TextBox()
        Me.TxtXpressNetAddress = New System.Windows.Forms.TextBox()
        Me.TxtInterfaceCode = New System.Windows.Forms.TextBox()
        Me.LbFreeTcpSockets = New System.Windows.Forms.Label()
        Me.LbXpressNetVersion = New System.Windows.Forms.Label()
        Me.LbXpressNetAddress = New System.Windows.Forms.Label()
        Me.LbInterfaceCode = New System.Windows.Forms.Label()
        Me.BtnInterfaceInfo = New System.Windows.Forms.Button()
        Me.TxtInterfaceVersion = New System.Windows.Forms.TextBox()
        Me.TxtMasterVersion = New System.Windows.Forms.TextBox()
        Me.LbInterfaceVersion = New System.Windows.Forms.Label()
        Me.LbMasterVersion = New System.Windows.Forms.Label()
        Me.ShapeContainer2 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape2 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.GbAllesReset = New System.Windows.Forms.GroupBox()
        Me.BtnAllesUit = New System.Windows.Forms.Button()
        Me.BtnAllesAan = New System.Windows.Forms.Button()
        Me.GbSwitch = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LbSwitchI = New System.Windows.Forms.Label()
        Me.LbSwitchTT = New System.Windows.Forms.Label()
        Me.LbSwitchData = New System.Windows.Forms.Label()
        Me.LbSwitchAdrress = New System.Windows.Forms.Label()
        Me.TxtSwitchI = New System.Windows.Forms.TextBox()
        Me.TxtSwitchTT = New System.Windows.Forms.TextBox()
        Me.TxtSwitchData = New System.Windows.Forms.TextBox()
        Me.TxtSwitchAddress = New System.Windows.Forms.TextBox()
        Me.BtnRecht = New System.Windows.Forms.Button()
        Me.BtnAf = New System.Windows.Forms.Button()
        Me.TxtSwitch = New System.Windows.Forms.TextBox()
        Me.LbSwitch = New System.Windows.Forms.Label()
        Me.GbRSBus = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LbRSBusNibble = New System.Windows.Forms.Label()
        Me.LbRSBusData = New System.Windows.Forms.Label()
        Me.LbAddresBusAddress = New System.Windows.Forms.Label()
        Me.TxtRSBusNibble = New System.Windows.Forms.TextBox()
        Me.TxtRSBusData = New System.Windows.Forms.TextBox()
        Me.TxtRSBusAddress = New System.Windows.Forms.TextBox()
        Me.TxtRSBusAddressCall = New System.Windows.Forms.TextBox()
        Me.LbRSBusAddress = New System.Windows.Forms.Label()
        Me.GbTcpConnect.SuspendLayout()
        Me.GbConnectUsb.SuspendLayout()
        Me.GbConnectInfo.SuspendLayout()
        Me.GbXpressNetVersion.SuspendLayout()
        Me.GbAllesReset.SuspendLayout()
        Me.GbSwitch.SuspendLayout()
        Me.GbRSBus.SuspendLayout()
        Me.SuspendLayout()
        '
        'GbTcpConnect
        '
        Me.GbTcpConnect.Controls.Add(Me.BtnConnectTcpIp)
        Me.GbTcpConnect.Controls.Add(Me.TxtTcpPort)
        Me.GbTcpConnect.Controls.Add(Me.TxtIpAddress)
        Me.GbTcpConnect.Controls.Add(Me.LbTcpPort)
        Me.GbTcpConnect.Controls.Add(Me.LbIpAddress)
        Me.GbTcpConnect.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbTcpConnect.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbTcpConnect.Location = New System.Drawing.Point(10, 10)
        Me.GbTcpConnect.Name = "GbTcpConnect"
        Me.GbTcpConnect.Size = New System.Drawing.Size(259, 169)
        Me.GbTcpConnect.TabIndex = 0
        Me.GbTcpConnect.TabStop = False
        Me.GbTcpConnect.Text = "Connect TCP/IP"
        '
        'BtnConnectTcpIp
        '
        Me.BtnConnectTcpIp.ForeColor = System.Drawing.Color.Black
        Me.BtnConnectTcpIp.Location = New System.Drawing.Point(100, 119)
        Me.BtnConnectTcpIp.Name = "BtnConnectTcpIp"
        Me.BtnConnectTcpIp.Size = New System.Drawing.Size(150, 30)
        Me.BtnConnectTcpIp.TabIndex = 4
        Me.BtnConnectTcpIp.Text = "Connect TCP/IP"
        Me.BtnConnectTcpIp.UseVisualStyleBackColor = True
        '
        'TxtTcpPort
        '
        Me.TxtTcpPort.Location = New System.Drawing.Point(100, 75)
        Me.TxtTcpPort.Name = "TxtTcpPort"
        Me.TxtTcpPort.Size = New System.Drawing.Size(150, 27)
        Me.TxtTcpPort.TabIndex = 3
        Me.TxtTcpPort.Text = "5550"
        '
        'TxtIpAddress
        '
        Me.TxtIpAddress.Location = New System.Drawing.Point(100, 28)
        Me.TxtIpAddress.Name = "TxtIpAddress"
        Me.TxtIpAddress.Size = New System.Drawing.Size(150, 27)
        Me.TxtIpAddress.TabIndex = 2
        Me.TxtIpAddress.Text = "oldenzaal"
        '
        'LbTcpPort
        '
        Me.LbTcpPort.ForeColor = System.Drawing.Color.Black
        Me.LbTcpPort.Location = New System.Drawing.Point(10, 75)
        Me.LbTcpPort.Name = "LbTcpPort"
        Me.LbTcpPort.Size = New System.Drawing.Size(90, 20)
        Me.LbTcpPort.TabIndex = 1
        Me.LbTcpPort.Text = "TCP Port:"
        Me.LbTcpPort.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbIpAddress
        '
        Me.LbIpAddress.ForeColor = System.Drawing.Color.Black
        Me.LbIpAddress.Location = New System.Drawing.Point(8, 31)
        Me.LbIpAddress.Name = "LbIpAddress"
        Me.LbIpAddress.Size = New System.Drawing.Size(90, 20)
        Me.LbIpAddress.TabIndex = 0
        Me.LbIpAddress.Text = "IP Address:"
        Me.LbIpAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GbConnectUsb
        '
        Me.GbConnectUsb.Controls.Add(Me.CbBaudrate)
        Me.GbConnectUsb.Controls.Add(Me.BtnConnectUsb)
        Me.GbConnectUsb.Controls.Add(Me.TxtUsbPort)
        Me.GbConnectUsb.Controls.Add(Me.LbBaudrate)
        Me.GbConnectUsb.Controls.Add(Me.LbUsbPort)
        Me.GbConnectUsb.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbConnectUsb.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbConnectUsb.Location = New System.Drawing.Point(280, 10)
        Me.GbConnectUsb.Name = "GbConnectUsb"
        Me.GbConnectUsb.Size = New System.Drawing.Size(231, 169)
        Me.GbConnectUsb.TabIndex = 1
        Me.GbConnectUsb.TabStop = False
        Me.GbConnectUsb.Text = "Connect USB"
        '
        'CbBaudrate
        '
        Me.CbBaudrate.FormattingEnabled = True
        Me.CbBaudrate.Location = New System.Drawing.Point(100, 76)
        Me.CbBaudrate.Name = "CbBaudrate"
        Me.CbBaudrate.Size = New System.Drawing.Size(110, 27)
        Me.CbBaudrate.TabIndex = 5
        '
        'BtnConnectUsb
        '
        Me.BtnConnectUsb.ForeColor = System.Drawing.Color.Black
        Me.BtnConnectUsb.Location = New System.Drawing.Point(100, 119)
        Me.BtnConnectUsb.Name = "BtnConnectUsb"
        Me.BtnConnectUsb.Size = New System.Drawing.Size(110, 30)
        Me.BtnConnectUsb.TabIndex = 4
        Me.BtnConnectUsb.Text = "Connect USB"
        Me.BtnConnectUsb.UseVisualStyleBackColor = True
        '
        'TxtUsbPort
        '
        Me.TxtUsbPort.Location = New System.Drawing.Point(100, 28)
        Me.TxtUsbPort.Name = "TxtUsbPort"
        Me.TxtUsbPort.Size = New System.Drawing.Size(110, 27)
        Me.TxtUsbPort.TabIndex = 2
        Me.TxtUsbPort.Text = "3"
        '
        'LbBaudrate
        '
        Me.LbBaudrate.ForeColor = System.Drawing.Color.Black
        Me.LbBaudrate.Location = New System.Drawing.Point(10, 75)
        Me.LbBaudrate.Name = "LbBaudrate"
        Me.LbBaudrate.Size = New System.Drawing.Size(90, 20)
        Me.LbBaudrate.TabIndex = 1
        Me.LbBaudrate.Text = "Baudrate:"
        Me.LbBaudrate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbUsbPort
        '
        Me.LbUsbPort.ForeColor = System.Drawing.Color.Black
        Me.LbUsbPort.Location = New System.Drawing.Point(8, 31)
        Me.LbUsbPort.Name = "LbUsbPort"
        Me.LbUsbPort.Size = New System.Drawing.Size(90, 20)
        Me.LbUsbPort.TabIndex = 0
        Me.LbUsbPort.Text = "USB Port:"
        Me.LbUsbPort.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GbConnectInfo
        '
        Me.GbConnectInfo.Controls.Add(Me.LbNotHalt)
        Me.GbConnectInfo.Controls.Add(Me.LbNotAus)
        Me.GbConnectInfo.Controls.Add(Me.LbConnected)
        Me.GbConnectInfo.Controls.Add(Me.ShapeContainer1)
        Me.GbConnectInfo.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbConnectInfo.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbConnectInfo.Location = New System.Drawing.Point(520, 10)
        Me.GbConnectInfo.Name = "GbConnectInfo"
        Me.GbConnectInfo.Size = New System.Drawing.Size(182, 169)
        Me.GbConnectInfo.TabIndex = 2
        Me.GbConnectInfo.TabStop = False
        Me.GbConnectInfo.Text = "Connect Info"
        '
        'LbNotHalt
        '
        Me.LbNotHalt.ForeColor = System.Drawing.Color.Black
        Me.LbNotHalt.Location = New System.Drawing.Point(60, 110)
        Me.LbNotHalt.Name = "LbNotHalt"
        Me.LbNotHalt.Size = New System.Drawing.Size(90, 20)
        Me.LbNotHalt.TabIndex = 2
        Me.LbNotHalt.Text = "Not Halt"
        Me.LbNotHalt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LbNotAus
        '
        Me.LbNotAus.ForeColor = System.Drawing.Color.Black
        Me.LbNotAus.Location = New System.Drawing.Point(60, 70)
        Me.LbNotAus.Name = "LbNotAus"
        Me.LbNotAus.Size = New System.Drawing.Size(90, 20)
        Me.LbNotAus.TabIndex = 1
        Me.LbNotAus.Text = "Not Aus"
        Me.LbNotAus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LbConnected
        '
        Me.LbConnected.ForeColor = System.Drawing.Color.Black
        Me.LbConnected.Location = New System.Drawing.Point(60, 30)
        Me.LbConnected.Name = "LbConnected"
        Me.LbConnected.Size = New System.Drawing.Size(90, 20)
        Me.LbConnected.TabIndex = 0
        Me.LbConnected.Text = "Connected"
        Me.LbConnected.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(3, 23)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.ShNotHalt, Me.ShNotAus, Me.ShConnected})
        Me.ShapeContainer1.Size = New System.Drawing.Size(176, 143)
        Me.ShapeContainer1.TabIndex = 3
        Me.ShapeContainer1.TabStop = False
        '
        'ShNotHalt
        '
        Me.ShNotHalt.BackColor = System.Drawing.Color.Yellow
        Me.ShNotHalt.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Opaque
        Me.ShNotHalt.Location = New System.Drawing.Point(18, 87)
        Me.ShNotHalt.Name = "ShNotHalt"
        Me.ShNotHalt.Size = New System.Drawing.Size(25, 25)
        '
        'ShNotAus
        '
        Me.ShNotAus.BackColor = System.Drawing.Color.Yellow
        Me.ShNotAus.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Opaque
        Me.ShNotAus.Location = New System.Drawing.Point(18, 47)
        Me.ShNotAus.Name = "ShNotAus"
        Me.ShNotAus.Size = New System.Drawing.Size(25, 25)
        '
        'ShConnected
        '
        Me.ShConnected.BackColor = System.Drawing.Color.Yellow
        Me.ShConnected.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Opaque
        Me.ShConnected.Location = New System.Drawing.Point(18, 7)
        Me.ShConnected.Name = "ShConnected"
        Me.ShConnected.Size = New System.Drawing.Size(25, 25)
        '
        'GbXpressNetVersion
        '
        Me.GbXpressNetVersion.Controls.Add(Me.BtnLanInfo)
        Me.GbXpressNetVersion.Controls.Add(Me.TxtFreeTcpSockets)
        Me.GbXpressNetVersion.Controls.Add(Me.TxtXpressNetVersion)
        Me.GbXpressNetVersion.Controls.Add(Me.TxtXpressNetAddress)
        Me.GbXpressNetVersion.Controls.Add(Me.TxtInterfaceCode)
        Me.GbXpressNetVersion.Controls.Add(Me.LbFreeTcpSockets)
        Me.GbXpressNetVersion.Controls.Add(Me.LbXpressNetVersion)
        Me.GbXpressNetVersion.Controls.Add(Me.LbXpressNetAddress)
        Me.GbXpressNetVersion.Controls.Add(Me.LbInterfaceCode)
        Me.GbXpressNetVersion.Controls.Add(Me.BtnInterfaceInfo)
        Me.GbXpressNetVersion.Controls.Add(Me.TxtInterfaceVersion)
        Me.GbXpressNetVersion.Controls.Add(Me.TxtMasterVersion)
        Me.GbXpressNetVersion.Controls.Add(Me.LbInterfaceVersion)
        Me.GbXpressNetVersion.Controls.Add(Me.LbMasterVersion)
        Me.GbXpressNetVersion.Controls.Add(Me.ShapeContainer2)
        Me.GbXpressNetVersion.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbXpressNetVersion.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbXpressNetVersion.Location = New System.Drawing.Point(714, 10)
        Me.GbXpressNetVersion.Name = "GbXpressNetVersion"
        Me.GbXpressNetVersion.Size = New System.Drawing.Size(368, 272)
        Me.GbXpressNetVersion.TabIndex = 3
        Me.GbXpressNetVersion.TabStop = False
        Me.GbXpressNetVersion.Text = "XpressNet Versie"
        '
        'BtnLanInfo
        '
        Me.BtnLanInfo.ForeColor = System.Drawing.Color.Black
        Me.BtnLanInfo.Location = New System.Drawing.Point(272, 175)
        Me.BtnLanInfo.Name = "BtnLanInfo"
        Me.BtnLanInfo.Size = New System.Drawing.Size(84, 59)
        Me.BtnLanInfo.TabIndex = 14
        Me.BtnLanInfo.Text = "LAN info"
        Me.BtnLanInfo.UseVisualStyleBackColor = True
        '
        'TxtFreeTcpSockets
        '
        Me.TxtFreeTcpSockets.Location = New System.Drawing.Point(150, 227)
        Me.TxtFreeTcpSockets.Name = "TxtFreeTcpSockets"
        Me.TxtFreeTcpSockets.Size = New System.Drawing.Size(100, 27)
        Me.TxtFreeTcpSockets.TabIndex = 12
        '
        'TxtXpressNetVersion
        '
        Me.TxtXpressNetVersion.Location = New System.Drawing.Point(150, 192)
        Me.TxtXpressNetVersion.Name = "TxtXpressNetVersion"
        Me.TxtXpressNetVersion.Size = New System.Drawing.Size(100, 27)
        Me.TxtXpressNetVersion.TabIndex = 11
        '
        'TxtXpressNetAddress
        '
        Me.TxtXpressNetAddress.Location = New System.Drawing.Point(150, 157)
        Me.TxtXpressNetAddress.Name = "TxtXpressNetAddress"
        Me.TxtXpressNetAddress.Size = New System.Drawing.Size(100, 27)
        Me.TxtXpressNetAddress.TabIndex = 10
        '
        'TxtInterfaceCode
        '
        Me.TxtInterfaceCode.Location = New System.Drawing.Point(150, 110)
        Me.TxtInterfaceCode.Name = "TxtInterfaceCode"
        Me.TxtInterfaceCode.Size = New System.Drawing.Size(100, 27)
        Me.TxtInterfaceCode.TabIndex = 9
        '
        'LbFreeTcpSockets
        '
        Me.LbFreeTcpSockets.ForeColor = System.Drawing.Color.Black
        Me.LbFreeTcpSockets.Location = New System.Drawing.Point(10, 230)
        Me.LbFreeTcpSockets.Name = "LbFreeTcpSockets"
        Me.LbFreeTcpSockets.Size = New System.Drawing.Size(140, 20)
        Me.LbFreeTcpSockets.TabIndex = 8
        Me.LbFreeTcpSockets.Text = "Free TCP sockets:"
        Me.LbFreeTcpSockets.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbXpressNetVersion
        '
        Me.LbXpressNetVersion.ForeColor = System.Drawing.Color.Black
        Me.LbXpressNetVersion.Location = New System.Drawing.Point(10, 195)
        Me.LbXpressNetVersion.Name = "LbXpressNetVersion"
        Me.LbXpressNetVersion.Size = New System.Drawing.Size(140, 20)
        Me.LbXpressNetVersion.TabIndex = 7
        Me.LbXpressNetVersion.Text = "XpressNet Version:"
        Me.LbXpressNetVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbXpressNetAddress
        '
        Me.LbXpressNetAddress.ForeColor = System.Drawing.Color.Black
        Me.LbXpressNetAddress.Location = New System.Drawing.Point(10, 160)
        Me.LbXpressNetAddress.Name = "LbXpressNetAddress"
        Me.LbXpressNetAddress.Size = New System.Drawing.Size(140, 20)
        Me.LbXpressNetAddress.TabIndex = 6
        Me.LbXpressNetAddress.Text = "XpressNet Address:"
        Me.LbXpressNetAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbInterfaceCode
        '
        Me.LbInterfaceCode.ForeColor = System.Drawing.Color.Black
        Me.LbInterfaceCode.Location = New System.Drawing.Point(10, 110)
        Me.LbInterfaceCode.Name = "LbInterfaceCode"
        Me.LbInterfaceCode.Size = New System.Drawing.Size(140, 20)
        Me.LbInterfaceCode.TabIndex = 5
        Me.LbInterfaceCode.Text = "Interface Code:"
        Me.LbInterfaceCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BtnInterfaceInfo
        '
        Me.BtnInterfaceInfo.ForeColor = System.Drawing.Color.Black
        Me.BtnInterfaceInfo.Location = New System.Drawing.Point(272, 75)
        Me.BtnInterfaceInfo.Name = "BtnInterfaceInfo"
        Me.BtnInterfaceInfo.Size = New System.Drawing.Size(84, 62)
        Me.BtnInterfaceInfo.TabIndex = 4
        Me.BtnInterfaceInfo.Text = "Interface Info"
        Me.BtnInterfaceInfo.UseVisualStyleBackColor = True
        '
        'TxtInterfaceVersion
        '
        Me.TxtInterfaceVersion.Location = New System.Drawing.Point(150, 75)
        Me.TxtInterfaceVersion.Name = "TxtInterfaceVersion"
        Me.TxtInterfaceVersion.Size = New System.Drawing.Size(100, 27)
        Me.TxtInterfaceVersion.TabIndex = 3
        '
        'TxtMasterVersion
        '
        Me.TxtMasterVersion.Location = New System.Drawing.Point(150, 28)
        Me.TxtMasterVersion.Name = "TxtMasterVersion"
        Me.TxtMasterVersion.Size = New System.Drawing.Size(100, 27)
        Me.TxtMasterVersion.TabIndex = 2
        '
        'LbInterfaceVersion
        '
        Me.LbInterfaceVersion.ForeColor = System.Drawing.Color.Black
        Me.LbInterfaceVersion.Location = New System.Drawing.Point(10, 75)
        Me.LbInterfaceVersion.Name = "LbInterfaceVersion"
        Me.LbInterfaceVersion.Size = New System.Drawing.Size(140, 20)
        Me.LbInterfaceVersion.TabIndex = 1
        Me.LbInterfaceVersion.Text = "Interface Version:"
        Me.LbInterfaceVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbMasterVersion
        '
        Me.LbMasterVersion.ForeColor = System.Drawing.Color.Black
        Me.LbMasterVersion.Location = New System.Drawing.Point(10, 31)
        Me.LbMasterVersion.Name = "LbMasterVersion"
        Me.LbMasterVersion.Size = New System.Drawing.Size(140, 20)
        Me.LbMasterVersion.TabIndex = 0
        Me.LbMasterVersion.Text = "Master Version:"
        Me.LbMasterVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ShapeContainer2
        '
        Me.ShapeContainer2.Location = New System.Drawing.Point(3, 23)
        Me.ShapeContainer2.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer2.Name = "ShapeContainer2"
        Me.ShapeContainer2.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape2, Me.LineShape1})
        Me.ShapeContainer2.Size = New System.Drawing.Size(362, 246)
        Me.ShapeContainer2.TabIndex = 13
        Me.ShapeContainer2.TabStop = False
        '
        'LineShape2
        '
        Me.LineShape2.Name = "LineShape2"
        Me.LineShape2.X1 = 10
        Me.LineShape2.X2 = 350
        Me.LineShape2.Y1 = 125
        Me.LineShape2.Y2 = 125
        '
        'LineShape1
        '
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 10
        Me.LineShape1.X2 = 350
        Me.LineShape1.Y1 = 40
        Me.LineShape1.Y2 = 40
        '
        'GbAllesReset
        '
        Me.GbAllesReset.Controls.Add(Me.BtnAllesUit)
        Me.GbAllesReset.Controls.Add(Me.BtnAllesAan)
        Me.GbAllesReset.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GbAllesReset.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbAllesReset.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbAllesReset.Location = New System.Drawing.Point(714, 297)
        Me.GbAllesReset.Name = "GbAllesReset"
        Me.GbAllesReset.Size = New System.Drawing.Size(368, 120)
        Me.GbAllesReset.TabIndex = 4
        Me.GbAllesReset.TabStop = False
        Me.GbAllesReset.Text = "Alles Resetten"
        '
        'BtnAllesUit
        '
        Me.BtnAllesUit.ForeColor = System.Drawing.Color.Black
        Me.BtnAllesUit.Location = New System.Drawing.Point(216, 40)
        Me.BtnAllesUit.Name = "BtnAllesUit"
        Me.BtnAllesUit.Size = New System.Drawing.Size(110, 50)
        Me.BtnAllesUit.TabIndex = 5
        Me.BtnAllesUit.Text = "Alles Uit"
        Me.BtnAllesUit.UseVisualStyleBackColor = True
        '
        'BtnAllesAan
        '
        Me.BtnAllesAan.ForeColor = System.Drawing.Color.Black
        Me.BtnAllesAan.Location = New System.Drawing.Point(37, 40)
        Me.BtnAllesAan.Name = "BtnAllesAan"
        Me.BtnAllesAan.Size = New System.Drawing.Size(110, 50)
        Me.BtnAllesAan.TabIndex = 4
        Me.BtnAllesAan.Text = "Alles Aan"
        Me.BtnAllesAan.UseVisualStyleBackColor = True
        '
        'GbSwitch
        '
        Me.GbSwitch.Controls.Add(Me.Label1)
        Me.GbSwitch.Controls.Add(Me.LbSwitchI)
        Me.GbSwitch.Controls.Add(Me.LbSwitchTT)
        Me.GbSwitch.Controls.Add(Me.LbSwitchData)
        Me.GbSwitch.Controls.Add(Me.LbSwitchAdrress)
        Me.GbSwitch.Controls.Add(Me.TxtSwitchI)
        Me.GbSwitch.Controls.Add(Me.TxtSwitchTT)
        Me.GbSwitch.Controls.Add(Me.TxtSwitchData)
        Me.GbSwitch.Controls.Add(Me.TxtSwitchAddress)
        Me.GbSwitch.Controls.Add(Me.BtnRecht)
        Me.GbSwitch.Controls.Add(Me.BtnAf)
        Me.GbSwitch.Controls.Add(Me.TxtSwitch)
        Me.GbSwitch.Controls.Add(Me.LbSwitch)
        Me.GbSwitch.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbSwitch.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbSwitch.Location = New System.Drawing.Point(10, 180)
        Me.GbSwitch.Name = "GbSwitch"
        Me.GbSwitch.Size = New System.Drawing.Size(317, 493)
        Me.GbSwitch.TabIndex = 5
        Me.GbSwitch.TabStop = False
        Me.GbSwitch.Text = "Switches"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(66, 463)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(138, 19)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Last message at top"
        '
        'LbSwitchI
        '
        Me.LbSwitchI.AutoSize = True
        Me.LbSwitchI.ForeColor = System.Drawing.Color.Black
        Me.LbSwitchI.Location = New System.Drawing.Point(268, 88)
        Me.LbSwitchI.Name = "LbSwitchI"
        Me.LbSwitchI.Size = New System.Drawing.Size(13, 19)
        Me.LbSwitchI.TabIndex = 11
        Me.LbSwitchI.Text = "I"
        '
        'LbSwitchTT
        '
        Me.LbSwitchTT.AutoSize = True
        Me.LbSwitchTT.ForeColor = System.Drawing.Color.Black
        Me.LbSwitchTT.Location = New System.Drawing.Point(210, 88)
        Me.LbSwitchTT.Name = "LbSwitchTT"
        Me.LbSwitchTT.Size = New System.Drawing.Size(25, 19)
        Me.LbSwitchTT.TabIndex = 10
        Me.LbSwitchTT.Text = "TT"
        '
        'LbSwitchData
        '
        Me.LbSwitchData.AutoSize = True
        Me.LbSwitchData.ForeColor = System.Drawing.Color.Black
        Me.LbSwitchData.Location = New System.Drawing.Point(130, 88)
        Me.LbSwitchData.Name = "LbSwitchData"
        Me.LbSwitchData.Size = New System.Drawing.Size(41, 19)
        Me.LbSwitchData.TabIndex = 9
        Me.LbSwitchData.Text = "Data"
        '
        'LbSwitchAdrress
        '
        Me.LbSwitchAdrress.AutoSize = True
        Me.LbSwitchAdrress.ForeColor = System.Drawing.Color.Black
        Me.LbSwitchAdrress.Location = New System.Drawing.Point(25, 88)
        Me.LbSwitchAdrress.Name = "LbSwitchAdrress"
        Me.LbSwitchAdrress.Size = New System.Drawing.Size(63, 19)
        Me.LbSwitchAdrress.TabIndex = 8
        Me.LbSwitchAdrress.Text = "Address"
        '
        'TxtSwitchI
        '
        Me.TxtSwitchI.Location = New System.Drawing.Point(248, 110)
        Me.TxtSwitchI.Multiline = True
        Me.TxtSwitchI.Name = "TxtSwitchI"
        Me.TxtSwitchI.Size = New System.Drawing.Size(52, 350)
        Me.TxtSwitchI.TabIndex = 7
        '
        'TxtSwitchTT
        '
        Me.TxtSwitchTT.Location = New System.Drawing.Point(198, 110)
        Me.TxtSwitchTT.Multiline = True
        Me.TxtSwitchTT.Name = "TxtSwitchTT"
        Me.TxtSwitchTT.Size = New System.Drawing.Size(52, 350)
        Me.TxtSwitchTT.TabIndex = 6
        '
        'TxtSwitchData
        '
        Me.TxtSwitchData.Location = New System.Drawing.Point(104, 110)
        Me.TxtSwitchData.Multiline = True
        Me.TxtSwitchData.Name = "TxtSwitchData"
        Me.TxtSwitchData.Size = New System.Drawing.Size(95, 350)
        Me.TxtSwitchData.TabIndex = 5
        '
        'TxtSwitchAddress
        '
        Me.TxtSwitchAddress.Location = New System.Drawing.Point(10, 110)
        Me.TxtSwitchAddress.Multiline = True
        Me.TxtSwitchAddress.Name = "TxtSwitchAddress"
        Me.TxtSwitchAddress.Size = New System.Drawing.Size(95, 350)
        Me.TxtSwitchAddress.TabIndex = 4
        '
        'BtnRecht
        '
        Me.BtnRecht.Font = New System.Drawing.Font("Calibri", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnRecht.ForeColor = System.Drawing.Color.Black
        Me.BtnRecht.Location = New System.Drawing.Point(230, 27)
        Me.BtnRecht.Name = "BtnRecht"
        Me.BtnRecht.Size = New System.Drawing.Size(70, 41)
        Me.BtnRecht.TabIndex = 3
        Me.BtnRecht.Text = "=>"
        Me.BtnRecht.UseVisualStyleBackColor = True
        '
        'BtnAf
        '
        Me.BtnAf.Font = New System.Drawing.Font("Calibri", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAf.ForeColor = System.Drawing.Color.Black
        Me.BtnAf.Location = New System.Drawing.Point(144, 28)
        Me.BtnAf.Margin = New System.Windows.Forms.Padding(1)
        Me.BtnAf.Name = "BtnAf"
        Me.BtnAf.Size = New System.Drawing.Size(70, 41)
        Me.BtnAf.TabIndex = 2
        Me.BtnAf.Text = "<="
        Me.BtnAf.UseVisualStyleBackColor = True
        '
        'TxtSwitch
        '
        Me.TxtSwitch.Location = New System.Drawing.Point(70, 35)
        Me.TxtSwitch.Name = "TxtSwitch"
        Me.TxtSwitch.Size = New System.Drawing.Size(67, 27)
        Me.TxtSwitch.TabIndex = 1
        '
        'LbSwitch
        '
        Me.LbSwitch.AutoSize = True
        Me.LbSwitch.ForeColor = System.Drawing.Color.Black
        Me.LbSwitch.Location = New System.Drawing.Point(10, 35)
        Me.LbSwitch.Name = "LbSwitch"
        Me.LbSwitch.Size = New System.Drawing.Size(63, 19)
        Me.LbSwitch.TabIndex = 0
        Me.LbSwitch.Text = "Switch: "
        '
        'GbRSBus
        '
        Me.GbRSBus.Controls.Add(Me.Label2)
        Me.GbRSBus.Controls.Add(Me.LbRSBusNibble)
        Me.GbRSBus.Controls.Add(Me.LbRSBusData)
        Me.GbRSBus.Controls.Add(Me.LbAddresBusAddress)
        Me.GbRSBus.Controls.Add(Me.TxtRSBusNibble)
        Me.GbRSBus.Controls.Add(Me.TxtRSBusData)
        Me.GbRSBus.Controls.Add(Me.TxtRSBusAddress)
        Me.GbRSBus.Controls.Add(Me.TxtRSBusAddressCall)
        Me.GbRSBus.Controls.Add(Me.LbRSBusAddress)
        Me.GbRSBus.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GbRSBus.ForeColor = System.Drawing.Color.DodgerBlue
        Me.GbRSBus.Location = New System.Drawing.Point(340, 180)
        Me.GbRSBus.Name = "GbRSBus"
        Me.GbRSBus.Size = New System.Drawing.Size(279, 493)
        Me.GbRSBus.TabIndex = 6
        Me.GbRSBus.TabStop = False
        Me.GbRSBus.Text = "RSBus"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(61, 463)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(138, 19)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Last message at top"
        '
        'LbRSBusNibble
        '
        Me.LbRSBusNibble.AutoSize = True
        Me.LbRSBusNibble.ForeColor = System.Drawing.Color.Black
        Me.LbRSBusNibble.Location = New System.Drawing.Point(200, 88)
        Me.LbRSBusNibble.Name = "LbRSBusNibble"
        Me.LbRSBusNibble.Size = New System.Drawing.Size(54, 19)
        Me.LbRSBusNibble.TabIndex = 10
        Me.LbRSBusNibble.Text = "Nibble"
        '
        'LbRSBusData
        '
        Me.LbRSBusData.AutoSize = True
        Me.LbRSBusData.ForeColor = System.Drawing.Color.Black
        Me.LbRSBusData.Location = New System.Drawing.Point(130, 88)
        Me.LbRSBusData.Name = "LbRSBusData"
        Me.LbRSBusData.Size = New System.Drawing.Size(41, 19)
        Me.LbRSBusData.TabIndex = 9
        Me.LbRSBusData.Text = "Data"
        '
        'LbAddresBusAddress
        '
        Me.LbAddresBusAddress.AutoSize = True
        Me.LbAddresBusAddress.ForeColor = System.Drawing.Color.Black
        Me.LbAddresBusAddress.Location = New System.Drawing.Point(25, 88)
        Me.LbAddresBusAddress.Name = "LbAddresBusAddress"
        Me.LbAddresBusAddress.Size = New System.Drawing.Size(63, 19)
        Me.LbAddresBusAddress.TabIndex = 8
        Me.LbAddresBusAddress.Text = "Address"
        '
        'TxtRSBusNibble
        '
        Me.TxtRSBusNibble.Location = New System.Drawing.Point(198, 110)
        Me.TxtRSBusNibble.Multiline = True
        Me.TxtRSBusNibble.Name = "TxtRSBusNibble"
        Me.TxtRSBusNibble.Size = New System.Drawing.Size(60, 350)
        Me.TxtRSBusNibble.TabIndex = 6
        '
        'TxtRSBusData
        '
        Me.TxtRSBusData.Location = New System.Drawing.Point(104, 110)
        Me.TxtRSBusData.Multiline = True
        Me.TxtRSBusData.Name = "TxtRSBusData"
        Me.TxtRSBusData.Size = New System.Drawing.Size(95, 350)
        Me.TxtRSBusData.TabIndex = 5
        '
        'TxtRSBusAddress
        '
        Me.TxtRSBusAddress.Location = New System.Drawing.Point(10, 110)
        Me.TxtRSBusAddress.Multiline = True
        Me.TxtRSBusAddress.Name = "TxtRSBusAddress"
        Me.TxtRSBusAddress.Size = New System.Drawing.Size(95, 350)
        Me.TxtRSBusAddress.TabIndex = 4
        '
        'TxtRSBusAddressCall
        '
        Me.TxtRSBusAddressCall.Location = New System.Drawing.Point(124, 32)
        Me.TxtRSBusAddressCall.Name = "TxtRSBusAddressCall"
        Me.TxtRSBusAddressCall.Size = New System.Drawing.Size(103, 27)
        Me.TxtRSBusAddressCall.TabIndex = 1
        '
        'LbRSBusAddress
        '
        Me.LbRSBusAddress.AutoSize = True
        Me.LbRSBusAddress.ForeColor = System.Drawing.Color.Black
        Me.LbRSBusAddress.Location = New System.Drawing.Point(10, 35)
        Me.LbRSBusAddress.Name = "LbRSBusAddress"
        Me.LbRSBusAddress.Size = New System.Drawing.Size(116, 19)
        Me.LbRSBusAddress.TabIndex = 0
        Me.LbRSBusAddress.Text = "RSBus Address: "
        '
        'FrmBasicTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1094, 678)
        Me.Controls.Add(Me.GbRSBus)
        Me.Controls.Add(Me.GbSwitch)
        Me.Controls.Add(Me.GbAllesReset)
        Me.Controls.Add(Me.GbXpressNetVersion)
        Me.Controls.Add(Me.GbConnectInfo)
        Me.Controls.Add(Me.GbConnectUsb)
        Me.Controls.Add(Me.GbTcpConnect)
        Me.Location = New System.Drawing.Point(18, 0)
        Me.Name = "FrmBasicTest"
        Me.Padding = New System.Windows.Forms.Padding(0, 0, 0, 3)
        Me.Text = "Basic Test"
        Me.GbTcpConnect.ResumeLayout(False)
        Me.GbTcpConnect.PerformLayout()
        Me.GbConnectUsb.ResumeLayout(False)
        Me.GbConnectUsb.PerformLayout()
        Me.GbConnectInfo.ResumeLayout(False)
        Me.GbXpressNetVersion.ResumeLayout(False)
        Me.GbXpressNetVersion.PerformLayout()
        Me.GbAllesReset.ResumeLayout(False)
        Me.GbSwitch.ResumeLayout(False)
        Me.GbSwitch.PerformLayout()
        Me.GbRSBus.ResumeLayout(False)
        Me.GbRSBus.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GbTcpConnect As GroupBox
    Friend WithEvents BtnConnectTcpIp As Button
    Friend WithEvents TxtTcpPort As TextBox
    Friend WithEvents TxtIpAddress As TextBox
    Friend WithEvents LbTcpPort As Label
    Friend WithEvents LbIpAddress As Label
    Friend WithEvents GbConnectUsb As GroupBox
    Friend WithEvents CbBaudrate As ComboBox
    Friend WithEvents BtnConnectUsb As Button
    Friend WithEvents TxtUsbPort As TextBox
    Friend WithEvents LbBaudrate As Label
    Friend WithEvents LbUsbPort As Label
    Friend WithEvents GbConnectInfo As GroupBox
    Friend WithEvents LbNotHalt As Label
    Friend WithEvents LbNotAus As Label
    Friend WithEvents LbConnected As Label
    Friend WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Friend WithEvents ShNotAus As PowerPacks.OvalShape
    Friend WithEvents ShConnected As PowerPacks.OvalShape
    Friend WithEvents ShNotHalt As PowerPacks.OvalShape
    Friend WithEvents GbXpressNetVersion As GroupBox
    Friend WithEvents BtnLanInfo As Button
    Friend WithEvents TxtFreeTcpSockets As TextBox
    Friend WithEvents TxtXpressNetVersion As TextBox
    Friend WithEvents TxtXpressNetAddress As TextBox
    Friend WithEvents TxtInterfaceCode As TextBox
    Friend WithEvents LbFreeTcpSockets As Label
    Friend WithEvents LbXpressNetVersion As Label
    Friend WithEvents LbXpressNetAddress As Label
    Friend WithEvents LbInterfaceCode As Label
    Friend WithEvents BtnInterfaceInfo As Button
    Friend WithEvents TxtInterfaceVersion As TextBox
    Friend WithEvents TxtMasterVersion As TextBox
    Friend WithEvents LbInterfaceVersion As Label
    Friend WithEvents LbMasterVersion As Label
    Friend WithEvents ShapeContainer2 As PowerPacks.ShapeContainer
    Friend WithEvents LineShape2 As PowerPacks.LineShape
    Friend WithEvents LineShape1 As PowerPacks.LineShape
    Friend WithEvents GbAllesReset As GroupBox
    Friend WithEvents BtnAllesUit As Button
    Friend WithEvents BtnAllesAan As Button
    Friend WithEvents GbSwitch As GroupBox
    Friend WithEvents BtnRecht As Button
    Friend WithEvents BtnAf As Button
    Friend WithEvents TxtSwitch As TextBox
    Friend WithEvents LbSwitch As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents LbSwitchI As Label
    Friend WithEvents LbSwitchTT As Label
    Friend WithEvents LbSwitchData As Label
    Friend WithEvents LbSwitchAdrress As Label
    Friend WithEvents TxtSwitchI As TextBox
    Friend WithEvents TxtSwitchTT As TextBox
    Friend WithEvents TxtSwitchData As TextBox
    Friend WithEvents TxtSwitchAddress As TextBox
    Friend WithEvents GbRSBus As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents LbRSBusNibble As Label
    Friend WithEvents LbRSBusData As Label
    Friend WithEvents LbAddresBusAddress As Label
    Friend WithEvents TxtRSBusNibble As TextBox
    Friend WithEvents TxtRSBusData As TextBox
    Friend WithEvents TxtRSBusAddress As TextBox
    Friend WithEvents TxtRSBusAddressCall As TextBox
    Friend WithEvents LbRSBusAddress As Label
End Class
