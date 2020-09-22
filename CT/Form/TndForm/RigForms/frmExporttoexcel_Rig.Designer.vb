<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmExporttoexcel_Rig
    Inherits frmBase  'System.Windows.Forms.Form

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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.chkDvpteam = New System.Windows.Forms.CheckBox()
        Me.chkChangelogs = New System.Windows.Forms.CheckBox()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.chkTndplan = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkDvpteam)
        Me.Panel1.Controls.Add(Me.chkChangelogs)
        Me.Panel1.Controls.Add(Me.btnExport)
        Me.Panel1.Controls.Add(Me.chkTndplan)
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Location = New System.Drawing.Point(8, 10)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(511, 82)
        Me.Panel1.TabIndex = 1
        '
        'chkDvpteam
        '
        Me.chkDvpteam.AutoSize = True
        Me.chkDvpteam.Checked = True
        Me.chkDvpteam.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDvpteam.Location = New System.Drawing.Point(227, 12)
        Me.chkDvpteam.Margin = New System.Windows.Forms.Padding(4)
        Me.chkDvpteam.Name = "chkDvpteam"
        Me.chkDvpteam.Size = New System.Drawing.Size(163, 21)
        Me.chkDvpteam.TabIndex = 8
        Me.chkDvpteam.Text = "DVP Team && CDSIDs"
        Me.chkDvpteam.UseVisualStyleBackColor = True
        '
        'chkChangelogs
        '
        Me.chkChangelogs.AutoSize = True
        Me.chkChangelogs.Checked = True
        Me.chkChangelogs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkChangelogs.Location = New System.Drawing.Point(105, 12)
        Me.chkChangelogs.Margin = New System.Windows.Forms.Padding(4)
        Me.chkChangelogs.Name = "chkChangelogs"
        Me.chkChangelogs.Size = New System.Drawing.Size(114, 21)
        Me.chkChangelogs.TabIndex = 7
        Me.chkChangelogs.Text = "Change Logs"
        Me.chkChangelogs.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExport.Location = New System.Drawing.Point(298, 46)
        Me.btnExport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(100, 28)
        Me.btnExport.TabIndex = 1
        Me.btnExport.Text = "&Export"
        Me.ToolTip1.SetToolTip(Me.btnExport, "Export [F7]")
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'chkTndplan
        '
        Me.chkTndplan.AutoSize = True
        Me.chkTndplan.Checked = True
        Me.chkTndplan.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTndplan.Location = New System.Drawing.Point(10, 12)
        Me.chkTndplan.Margin = New System.Windows.Forms.Padding(4)
        Me.chkTndplan.Name = "chkTndplan"
        Me.chkTndplan.Size = New System.Drawing.Size(87, 21)
        Me.chkTndplan.TabIndex = 4
        Me.chkTndplan.Text = "Tnd Plan"
        Me.chkTndplan.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(401, 46)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Close from [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmExporttoexcel_Rig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(527, 102)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmExporttoexcel_Rig"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Export to excel [Choose your report(s)]"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents chkDvpteam As System.Windows.Forms.CheckBox
    Friend WithEvents chkChangelogs As System.Windows.Forms.CheckBox
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents chkTndplan As System.Windows.Forms.CheckBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
