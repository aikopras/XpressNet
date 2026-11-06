<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmStart
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
        Me.cmdNewZ21 = New System.Windows.Forms.Button()
        Me.cmdNewXpressNet = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'cmdNewZ21
        '
        Me.cmdNewZ21.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNewZ21.Location = New System.Drawing.Point(200, 16)
        Me.cmdNewZ21.Name = "cmdNewZ21"
        Me.cmdNewZ21.Size = New System.Drawing.Size(145, 73)
        Me.cmdNewZ21.TabIndex = 1
        Me.cmdNewZ21.Text = "Add Z21"
        Me.cmdNewZ21.UseVisualStyleBackColor = True
        '
        'cmdNewXpressNet
        '
        Me.cmdNewXpressNet.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdNewXpressNet.Location = New System.Drawing.Point(24, 16)
        Me.cmdNewXpressNet.Name = "cmdNewXpressNet"
        Me.cmdNewXpressNet.Size = New System.Drawing.Size(145, 73)
        Me.cmdNewXpressNet.TabIndex = 0
        Me.cmdNewXpressNet.Text = "Add XpressNet"
        Me.cmdNewXpressNet.UseVisualStyleBackColor = True
        Me.ToolTip1.SetToolTip(Me.cmdNewXpressNet, "Schakel wissel 1 om")
        '
        'FrmStart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 116)
        Me.Controls.Add(Me.cmdNewZ21)
        Me.Controls.Add(Me.cmdNewXpressNet)
        Me.Name = "FrmStart"
        Me.Text = "Start"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdNewZ21 As System.Windows.Forms.Button
    Friend WithEvents cmdNewXpressNet As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
