Public Class frmProgressbar
    Private Sub frmProgressbar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SmoothProgressBar1.Value = 100
        Me.SmoothProgressBar2.Value = 0
        Me.Text = "loading..."
    End Sub

    'ProgressBar update
    Public Sub UpdateProgressBar(intprogressvalue As Double)
        If (Me.SmoothProgressBar1.Value > 0) Then
            Me.SmoothProgressBar1.Value -= intprogressvalue
            Me.SmoothProgressBar2.Value += intprogressvalue
            Me.Refresh()
        End If
    End Sub

    Private Sub SmoothProgressBar2_Load(sender As Object, e As EventArgs) Handles SmoothProgressBar2.Load

    End Sub
End Class