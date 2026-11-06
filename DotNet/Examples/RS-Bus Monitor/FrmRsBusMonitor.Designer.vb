<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmRsBusMonitor
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
        Me.Switch2RS = New System.Windows.Forms.Button()
        Me.RS2Switch = New System.Windows.Forms.Button()
        Me.TextSwitch2RS = New System.Windows.Forms.TextBox()
        Me.TextRS2Switch = New System.Windows.Forms.TextBox()
        Me.SchaltInfo = New System.Windows.Forms.Button()
        Me.CmdAllesAus = New System.Windows.Forms.Button()
        Me.CmdAllesAn = New System.Windows.Forms.Button()
        Me.cmdTcpConnect = New System.Windows.Forms.Button()
        Me.txtIP = New System.Windows.Forms.TextBox()
        Me.picPanel = New System.Windows.Forms.PictureBox()
        Me.Line_Shape1_T = New System.Windows.Forms.Label()
        Me.Line_Shape1_B = New System.Windows.Forms.Label()
        Me.Line_Shape1_L = New System.Windows.Forms.Label()
        Me.Line_Shape1_R = New System.Windows.Forms.Label()
        Me.LabelAddressConversion = New System.Windows.Forms.Label()
        Me.LabelSwitchAddress = New System.Windows.Forms.Label()
        Me.LabelRSAddress = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_T = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_B = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_L = New System.Windows.Forms.Label()
        Me.Line_ShapeAnAus_R = New System.Windows.Forms.Label()
        Me.ledInterfaceStatus = New Led()
        Me.ledNotHalt = New Led()
        Me.ledNotAus = New Led()
        Me.LabelNotHalt = New System.Windows.Forms.Label()
        Me.LabelNotAus = New System.Windows.Forms.Label()
        Me.LabelConnected = New System.Windows.Forms.Label()
        Me.LabelStatus = New System.Windows.Forms.Label()
        Me.LabelIpAddress = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_T = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_B = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_L = New System.Windows.Forms.Label()
        Me.Line_ShapeTcp_R = New System.Windows.Forms.Label()
        Me.Label_Header = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Switch2RS
        '
        Me.Switch2RS.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Switch2RS.Location = New System.Drawing.Point(112, 416)
        Me.Switch2RS.Name = "Switch2RS"
        Me.Switch2RS.Size = New System.Drawing.Size(25, 25)
        Me.Switch2RS.TabIndex = 16
        Me.Switch2RS.Text = "<"
        Me.Switch2RS.UseVisualStyleBackColor = True
        '
        'RS2Switch
        '
        Me.RS2Switch.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RS2Switch.Location = New System.Drawing.Point(112, 387)
        Me.RS2Switch.Name = "RS2Switch"
        Me.RS2Switch.Size = New System.Drawing.Size(25, 27)
        Me.RS2Switch.TabIndex = 15
        Me.RS2Switch.Text = ">"
        Me.RS2Switch.UseVisualStyleBackColor = True
        '
        'TextSwitch2RS
        '
        Me.TextSwitch2RS.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextSwitch2RS.Location = New System.Drawing.Point(144, 408)
        Me.TextSwitch2RS.Name = "TextSwitch2RS"
        Me.TextSwitch2RS.Size = New System.Drawing.Size(73, 22)
        Me.TextSwitch2RS.TabIndex = 14
        Me.TextSwitch2RS.Text = ""
        Me.TextSwitch2RS.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextRS2Switch
        '
        Me.TextRS2Switch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextRS2Switch.Location = New System.Drawing.Point(32, 408)
        Me.TextRS2Switch.Name = "TextRS2Switch"
        Me.TextRS2Switch.Size = New System.Drawing.Size(73, 22)
        Me.TextRS2Switch.TabIndex = 13
        Me.TextRS2Switch.Text = ""
        Me.TextRS2Switch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'SchaltInfo
        '
        Me.SchaltInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SchaltInfo.Location = New System.Drawing.Point(40, 480)
        Me.SchaltInfo.Name = "SchaltInfo"
        Me.SchaltInfo.Size = New System.Drawing.Size(177, 33)
        Me.SchaltInfo.TabIndex = 12
        Me.SchaltInfo.Text = "Get RS-Bus data"
        Me.SchaltInfo.UseVisualStyleBackColor = True
        '
        'CmdAllesAus
        '
        Me.CmdAllesAus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdAllesAus.Location = New System.Drawing.Point(136, 208)
        Me.CmdAllesAus.Name = "CmdAllesAus"
        Me.CmdAllesAus.Size = New System.Drawing.Size(75, 25)
        Me.CmdAllesAus.TabIndex = 11
        Me.CmdAllesAus.Text = "Alles Aus"
        Me.CmdAllesAus.UseVisualStyleBackColor = True
        '
        'CmdAllesAn
        '
        Me.CmdAllesAn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdAllesAn.Location = New System.Drawing.Point(32, 208)
        Me.CmdAllesAn.Name = "CmdAllesAn"
        Me.CmdAllesAn.Size = New System.Drawing.Size(75, 25)
        Me.CmdAllesAn.TabIndex = 10
        Me.CmdAllesAn.Text = "Alles An"
        Me.CmdAllesAn.UseVisualStyleBackColor = True
        '
        'cmdTcpConnect
        '
        Me.cmdTcpConnect.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTcpConnect.Location = New System.Drawing.Point(40, 120)
        Me.cmdTcpConnect.Name = "cmdTcpConnect"
        Me.cmdTcpConnect.Size = New System.Drawing.Size(169, 33)
        Me.cmdTcpConnect.TabIndex = 4
        Me.cmdTcpConnect.Text = "Connect TCP/IP"
        Me.cmdTcpConnect.UseVisualStyleBackColor = True
        '
        'txtIP
        '
        Me.txtIP.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIP.Location = New System.Drawing.Point(112, 88)
        Me.txtIP.Name = "txtIP"
        Me.txtIP.Size = New System.Drawing.Size(97, 22)
        Me.txtIP.TabIndex = 3
        Me.txtIP.Text = "192.168.0.200"
        '
        'picPanel
        '
        Me.picPanel.Location = New System.Drawing.Point(248, 16)
        Me.picPanel.Name = "picPanel"
        Me.picPanel.Size = New System.Drawing.Size(633, 417)
        Me.picPanel.TabIndex = 0
        '
        'Line_Shape1_T
        '
        Me.Line_Shape1_T.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_Shape1_T.Location = New System.Drawing.Point(24, 360)
        Me.Line_Shape1_T.Name = "Line_Shape1_T"
        Me.Line_Shape1_T.Size = New System.Drawing.Size(201, 1)
        '
        'Line_Shape1_B
        '
        Me.Line_Shape1_B.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_Shape1_B.Location = New System.Drawing.Point(24, 448)
        Me.Line_Shape1_B.Name = "Line_Shape1_B"
        Me.Line_Shape1_B.Size = New System.Drawing.Size(201, 1)
        '
        'Line_Shape1_L
        '
        Me.Line_Shape1_L.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_Shape1_L.Location = New System.Drawing.Point(24, 360)
        Me.Line_Shape1_L.Name = "Line_Shape1_L"
        Me.Line_Shape1_L.Size = New System.Drawing.Size(1, 89)
        '
        'Line_Shape1_R
        '
        Me.Line_Shape1_R.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_Shape1_R.Location = New System.Drawing.Point(224, 360)
        Me.Line_Shape1_R.Name = "Line_Shape1_R"
        Me.Line_Shape1_R.Size = New System.Drawing.Size(1, 89)
        '
        'LabelAddressConversion
        '
        Me.LabelAddressConversion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelAddressConversion.Location = New System.Drawing.Point(32, 368)
        Me.LabelAddressConversion.Name = "LabelAddressConversion"
        Me.LabelAddressConversion.Size = New System.Drawing.Size(185, 17)
        Me.LabelAddressConversion.TabIndex = 19
        Me.LabelAddressConversion.Text = "Address Conversions"
        Me.LabelAddressConversion.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelSwitchAddress
        '
        Me.LabelSwitchAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSwitchAddress.Location = New System.Drawing.Point(144, 384)
        Me.LabelSwitchAddress.Name = "LabelSwitchAddress"
        Me.LabelSwitchAddress.Size = New System.Drawing.Size(73, 17)
        Me.LabelSwitchAddress.TabIndex = 18
        Me.LabelSwitchAddress.Text = "Switch"
        Me.LabelSwitchAddress.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelRSAddress
        '
        Me.LabelRSAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelRSAddress.Location = New System.Drawing.Point(32, 384)
        Me.LabelRSAddress.Name = "LabelRSAddress"
        Me.LabelRSAddress.Size = New System.Drawing.Size(73, 17)
        Me.LabelRSAddress.TabIndex = 17
        Me.LabelRSAddress.Text = "RS"
        Me.LabelRSAddress.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Line_ShapeAnAus_T
        '
        Me.Line_ShapeAnAus_T.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeAnAus_T.Location = New System.Drawing.Point(24, 200)
        Me.Line_ShapeAnAus_T.Name = "Line_ShapeAnAus_T"
        Me.Line_ShapeAnAus_T.Size = New System.Drawing.Size(201, 1)
        '
        'Line_ShapeAnAus_B
        '
        Me.Line_ShapeAnAus_B.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeAnAus_B.Location = New System.Drawing.Point(24, 336)
        Me.Line_ShapeAnAus_B.Name = "Line_ShapeAnAus_B"
        Me.Line_ShapeAnAus_B.Size = New System.Drawing.Size(201, 1)
        '
        'Line_ShapeAnAus_L
        '
        Me.Line_ShapeAnAus_L.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeAnAus_L.Location = New System.Drawing.Point(24, 200)
        Me.Line_ShapeAnAus_L.Name = "Line_ShapeAnAus_L"
        Me.Line_ShapeAnAus_L.Size = New System.Drawing.Size(1, 137)
        '
        'Line_ShapeAnAus_R
        '
        Me.Line_ShapeAnAus_R.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeAnAus_R.Location = New System.Drawing.Point(224, 200)
        Me.Line_ShapeAnAus_R.Name = "Line_ShapeAnAus_R"
        Me.Line_ShapeAnAus_R.Size = New System.Drawing.Size(1, 137)
        '
        'ledInterfaceStatus
        '
        Me.ledInterfaceStatus.Location = New System.Drawing.Point(32, 248)
        Me.ledInterfaceStatus.Name = "ledInterfaceStatus"
        Me.ledInterfaceStatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledInterfaceStatus.Round = True
        Me.ledInterfaceStatus.Size = New System.Drawing.Size(17, 17)
        '
        'ledNotHalt
        '
        Me.ledNotHalt.Location = New System.Drawing.Point(32, 312)
        Me.ledNotHalt.Name = "ledNotHalt"
        Me.ledNotHalt.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledNotHalt.Round = True
        Me.ledNotHalt.Size = New System.Drawing.Size(17, 17)
        '
        'ledNotAus
        '
        Me.ledNotAus.Location = New System.Drawing.Point(32, 280)
        Me.ledNotAus.Name = "ledNotAus"
        Me.ledNotAus.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ledNotAus.Round = True
        Me.ledNotAus.Size = New System.Drawing.Size(17, 17)
        '
        'LabelNotHalt
        '
        Me.LabelNotHalt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNotHalt.Location = New System.Drawing.Point(57, 314)
        Me.LabelNotHalt.Name = "LabelNotHalt"
        Me.LabelNotHalt.AutoSize = True
        Me.LabelNotHalt.Size = New System.Drawing.Size(53, 16)
        Me.LabelNotHalt.TabIndex = 9
        Me.LabelNotHalt.Text = "NotHalt"
        Me.LabelNotHalt.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelNotAus
        '
        Me.LabelNotAus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNotAus.Location = New System.Drawing.Point(57, 282)
        Me.LabelNotAus.Name = "LabelNotAus"
        Me.LabelNotAus.AutoSize = True
        Me.LabelNotAus.Size = New System.Drawing.Size(52, 16)
        Me.LabelNotAus.TabIndex = 8
        Me.LabelNotAus.Text = "NotAus"
        Me.LabelNotAus.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelConnected
        '
        Me.LabelConnected.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelConnected.Location = New System.Drawing.Point(57, 250)
        Me.LabelConnected.Name = "LabelConnected"
        Me.LabelConnected.AutoSize = True
        Me.LabelConnected.Size = New System.Drawing.Size(73, 16)
        Me.LabelConnected.TabIndex = 7
        Me.LabelConnected.Text = "Connected"
        Me.LabelConnected.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelStatus
        '
        Me.LabelStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelStatus.Location = New System.Drawing.Point(24, 168)
        Me.LabelStatus.Name = "LabelStatus"
        Me.LabelStatus.Size = New System.Drawing.Size(201, 21)
        Me.LabelStatus.TabIndex = 6
        Me.LabelStatus.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'LabelIpAddress
        '
        Me.LabelIpAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelIpAddress.Location = New System.Drawing.Point(35, 88)
        Me.LabelIpAddress.Name = "LabelIpAddress"
        Me.LabelIpAddress.AutoSize = True
        Me.LabelIpAddress.Size = New System.Drawing.Size(74, 16)
        Me.LabelIpAddress.TabIndex = 5
        Me.LabelIpAddress.Text = "IP Address"
        Me.LabelIpAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Line_ShapeTcp_T
        '
        Me.Line_ShapeTcp_T.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_T.Location = New System.Drawing.Point(24, 72)
        Me.Line_ShapeTcp_T.Name = "Line_ShapeTcp_T"
        Me.Line_ShapeTcp_T.Size = New System.Drawing.Size(201, 1)
        '
        'Line_ShapeTcp_B
        '
        Me.Line_ShapeTcp_B.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_B.Location = New System.Drawing.Point(24, 160)
        Me.Line_ShapeTcp_B.Name = "Line_ShapeTcp_B"
        Me.Line_ShapeTcp_B.Size = New System.Drawing.Size(201, 1)
        '
        'Line_ShapeTcp_L
        '
        Me.Line_ShapeTcp_L.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_L.Location = New System.Drawing.Point(24, 72)
        Me.Line_ShapeTcp_L.Name = "Line_ShapeTcp_L"
        Me.Line_ShapeTcp_L.Size = New System.Drawing.Size(1, 89)
        '
        'Line_ShapeTcp_R
        '
        Me.Line_ShapeTcp_R.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line_ShapeTcp_R.Location = New System.Drawing.Point(224, 72)
        Me.Line_ShapeTcp_R.Name = "Line_ShapeTcp_R"
        Me.Line_ShapeTcp_R.Size = New System.Drawing.Size(1, 89)
        '
        'Label_Header
        '
        Me.Label_Header.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Header.Location = New System.Drawing.Point(24, 24)
        Me.Label_Header.Name = "Label_Header"
        Me.Label_Header.Size = New System.Drawing.Size(193, 25)
        Me.Label_Header.TabIndex = 2
        Me.Label_Header.Text = "RS-Bus Matrix display"
        Me.Label_Header.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(512, 312)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'FrmRsBusMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 653)
        Me.Controls.Add(Me.Switch2RS)
        Me.Controls.Add(Me.RS2Switch)
        Me.Controls.Add(Me.TextSwitch2RS)
        Me.Controls.Add(Me.TextRS2Switch)
        Me.Controls.Add(Me.SchaltInfo)
        Me.Controls.Add(Me.CmdAllesAus)
        Me.Controls.Add(Me.CmdAllesAn)
        Me.Controls.Add(Me.cmdTcpConnect)
        Me.Controls.Add(Me.txtIP)
        Me.Controls.Add(Me.picPanel)
        Me.Controls.Add(Me.LabelAddressConversion)
        Me.Controls.Add(Me.LabelSwitchAddress)
        Me.Controls.Add(Me.LabelRSAddress)
        Me.Controls.Add(Me.LabelNotHalt)
        Me.Controls.Add(Me.LabelNotAus)
        Me.Controls.Add(Me.LabelConnected)
        Me.Controls.Add(Me.LabelStatus)
        Me.Controls.Add(Me.LabelIpAddress)
        Me.Controls.Add(Me.Label_Header)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ledInterfaceStatus)
        Me.Controls.Add(Me.ledNotHalt)
        Me.Controls.Add(Me.ledNotAus)
        Me.Controls.Add(Me.Line_Shape1_T)
        Me.Controls.Add(Me.Line_Shape1_B)
        Me.Controls.Add(Me.Line_Shape1_L)
        Me.Controls.Add(Me.Line_Shape1_R)
        Me.Controls.Add(Me.Line_ShapeAnAus_T)
        Me.Controls.Add(Me.Line_ShapeAnAus_B)
        Me.Controls.Add(Me.Line_ShapeAnAus_L)
        Me.Controls.Add(Me.Line_ShapeAnAus_R)
        Me.Controls.Add(Me.Line_ShapeTcp_T)
        Me.Controls.Add(Me.Line_ShapeTcp_B)
        Me.Controls.Add(Me.Line_ShapeTcp_L)
        Me.Controls.Add(Me.Line_ShapeTcp_R)
        Me.Name = "FrmRsBusMonitor"
        Me.Text = "RsBusMonitor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Switch2RS As System.Windows.Forms.Button
    Friend WithEvents RS2Switch As System.Windows.Forms.Button
    Friend WithEvents TextSwitch2RS As System.Windows.Forms.TextBox
    Friend WithEvents TextRS2Switch As System.Windows.Forms.TextBox
    Friend WithEvents SchaltInfo As System.Windows.Forms.Button
    Friend WithEvents CmdAllesAus As System.Windows.Forms.Button
    Friend WithEvents CmdAllesAn As System.Windows.Forms.Button
    Friend WithEvents cmdTcpConnect As System.Windows.Forms.Button
    Friend WithEvents txtIP As System.Windows.Forms.TextBox
    Friend WithEvents picPanel As System.Windows.Forms.PictureBox
    Friend WithEvents Line_Shape1_T As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_B As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_L As System.Windows.Forms.Label
    Friend WithEvents Line_Shape1_R As System.Windows.Forms.Label
    Friend WithEvents LabelAddressConversion As System.Windows.Forms.Label
    Friend WithEvents LabelSwitchAddress As System.Windows.Forms.Label
    Friend WithEvents LabelRSAddress As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeAnAus_R As System.Windows.Forms.Label
    Friend WithEvents ledInterfaceStatus As Led
    Friend WithEvents ledNotHalt As Led
    Friend WithEvents ledNotAus As Led
    Friend WithEvents LabelNotHalt As System.Windows.Forms.Label
    Friend WithEvents LabelNotAus As System.Windows.Forms.Label
    Friend WithEvents LabelConnected As System.Windows.Forms.Label
    Friend WithEvents LabelStatus As System.Windows.Forms.Label
    Friend WithEvents LabelIpAddress As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_T As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_B As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_L As System.Windows.Forms.Label
    Friend WithEvents Line_ShapeTcp_R As System.Windows.Forms.Label
    Friend WithEvents Label_Header As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
