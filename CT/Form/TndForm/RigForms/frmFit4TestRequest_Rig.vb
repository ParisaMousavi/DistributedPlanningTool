Public Class frmFit4TestRequest_Rig
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmFit4TestRequest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wBrowserFit4Test.Navigate(CT.My.Resources.Fit4TestURL.ToString + "?hcid=" + CT.Form.DataCenter.ProgramConfig.HCID.ToString)
    End Sub
End Class