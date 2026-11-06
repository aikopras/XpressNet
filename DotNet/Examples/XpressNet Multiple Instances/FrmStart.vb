Option Explicit On
' 2026-11: VB.NET version of the VB6 example XpressNet Multiple Instances (B_UserInterfaceFormStart.frm), for XpressNet.NET 2.0.
' Every click opens a new form with its own XpressNetClass object: the Extended test (FrmExtendedTest) or the
' Z21 Basic test (FrmZ21Test). These forms are the same as in the examples XpressNet Test and Z21 Basic Test.
' Differences with VB6: a List instead of an array with ReDim Preserve; Show instead of Set ... = New and .Show.

Public Class FrmStart
    Private XpnConnections As New List(Of FrmExtendedTest)
    Private Z21Connections As New List(Of FrmZ21Test)


    Private Sub cmdNewXpressNet_Click(sender As Object, e As EventArgs) Handles cmdNewXpressNet.Click
        Dim XpnConnection As New FrmExtendedTest
        XpnConnections.Add(XpnConnection)
        XpnConnection.Show()
    End Sub

    Private Sub cmdNewZ21_Click(sender As Object, e As EventArgs) Handles cmdNewZ21.Click
        Dim Z21Connection As New FrmZ21Test
        Z21Connections.Add(Z21Connection)
        Z21Connection.Show()
    End Sub

End Class
