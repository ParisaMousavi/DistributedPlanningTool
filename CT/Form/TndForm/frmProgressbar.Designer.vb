<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProgressbar
    Inherits frmBase  'System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SmoothProgressBar1 = New CT.SmoothProgressBar.SmoothProgressBar()
        Me.SmoothProgressBar2 = New CT.SmoothProgressBar.SmoothProgressBar()
        Me.SuspendLayout()
        '
        'SmoothProgressBar1
        '
        Me.SmoothProgressBar1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SmoothProgressBar1.Location = New System.Drawing.Point(11, 11)
        Me.SmoothProgressBar1.Margin = New System.Windows.Forms.Padding(2)
        Me.SmoothProgressBar1.Maximum = 100
        Me.SmoothProgressBar1.Minimum = 0
        Me.SmoothProgressBar1.Name = "SmoothProgressBar1"
        Me.SmoothProgressBar1.ProgressBarColor = System.Drawing.Color.Blue
        Me.SmoothProgressBar1.Size = New System.Drawing.Size(456, 33)
        Me.SmoothProgressBar1.TabIndex = 0
        Me.SmoothProgressBar1.Value = 0R
        '
        'SmoothProgressBar2
        '
        Me.SmoothProgressBar2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SmoothProgressBar2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.SmoothProgressBar2.Location = New System.Drawing.Point(11, 11)
        Me.SmoothProgressBar2.Margin = New System.Windows.Forms.Padding(2)
        Me.SmoothProgressBar2.Maximum = 100
        Me.SmoothProgressBar2.Minimum = 0
        Me.SmoothProgressBar2.Name = "SmoothProgressBar2"
        Me.SmoothProgressBar2.ProgressBarColor = System.Drawing.Color.Blue
        Me.SmoothProgressBar2.Size = New System.Drawing.Size(456, 33)
        Me.SmoothProgressBar2.TabIndex = 1
        Me.SmoothProgressBar2.Value = 0R
        '
        'frmProgressbar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(478, 55)
        Me.ControlBox = False
        Me.Controls.Add(Me.SmoothProgressBar2)
        Me.Controls.Add(Me.SmoothProgressBar1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProgressbar"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SmoothProgressBar1 As SmoothProgressBar.SmoothProgressBar
    Friend WithEvents SmoothProgressBar2 As SmoothProgressBar.SmoothProgressBar
End Class
