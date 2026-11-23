Option Explicit On
' 2026-11: new. A LED, as the Shape control in the VB6 examples (and the PowerPacks OvalShape before).
' The colour of the LED is the BackColor; FillColor is the same property, with the name of the VB6 Shape control.
' Round = True: a circle (VB6: Shape = 3), Round = False: a rectangle (VB6: Shape = 0).
' After a build, the Led control is available in the Toolbox.

Public Class Led
    Inherits Control

    Private _round As Boolean = True

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Size = New Size(25, 25)
        BackColor = Color.Yellow
    End Sub

    ' Circle or rectangle
    Public Property Round As Boolean
        Get
            Return _round
        End Get
        Set(value As Boolean)
            _round = value
            Invalidate()
        End Set
    End Property

    ' The colour of the LED (VB6: FillColor of the Shape control)
    <System.ComponentModel.Browsable(False)>
    <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
    Public Property FillColor As Color
        Get
            Return BackColor
        End Get
        Set(value As Color)
            BackColor = value
        End Set
    End Property

    ' Around a round LED we show the background of the parent (for example a GroupBox)
    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        If Parent IsNot Nothing Then
            e.Graphics.Clear(Parent.BackColor)
        Else
            MyBase.OnPaintBackground(e)
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Using brush As New SolidBrush(BackColor)
            If _round Then
                e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                e.Graphics.FillEllipse(brush, 0, 0, Width - 1, Height - 1)
                e.Graphics.DrawEllipse(Pens.Black, 0, 0, Width - 1, Height - 1)
            Else
                e.Graphics.FillRectangle(brush, 0, 0, Width - 1, Height - 1)
                e.Graphics.DrawRectangle(Pens.Black, 0, 0, Width - 1, Height - 1)
            End If
        End Using
    End Sub

    ' A new BackColor (colour of the LED) is shown immediately
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        Invalidate()
    End Sub

End Class
