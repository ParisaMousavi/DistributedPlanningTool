<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmNewVehicle_Rig
    'Inherits System.Windows.Forms.Form
    Inherits frmBase

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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnAddUnit = New System.Windows.Forms.Button()
        Me.txtCounts = New System.Windows.Forms.TextBox()
        Me.lblCountsTitle = New System.Windows.Forms.Label()
        Me.lblBuildTypeTitle = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblBPTitle = New System.Windows.Forms.Label()
        Me.lblBuildPhase = New System.Windows.Forms.Label()
        Me.lblHCNTitle = New System.Windows.Forms.Label()
        Me.lblHCName = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblHCIDTitle = New System.Windows.Forms.Label()
        Me.lblHCID = New System.Windows.Forms.Label()
        Me.pnHealthChartName = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.grpNewVeh = New System.Windows.Forms.GroupBox()
        Me.lblBuildType = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pnHealthChartName.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.grpNewVeh.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(425, 14)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Close the form [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnAddUnit
        '
        Me.btnAddUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddUnit.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnAddUnit.Location = New System.Drawing.Point(317, 14)
        Me.btnAddUnit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAddUnit.Name = "btnAddUnit"
        Me.btnAddUnit.Size = New System.Drawing.Size(100, 28)
        Me.btnAddUnit.TabIndex = 0
        Me.btnAddUnit.Text = "&Add Unit"
        Me.ToolTip1.SetToolTip(Me.btnAddUnit, "Add Unit [F7]")
        Me.btnAddUnit.UseVisualStyleBackColor = True
        '
        'txtCounts
        '
        Me.txtCounts.Location = New System.Drawing.Point(431, 100)
        Me.txtCounts.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCounts.Name = "txtCounts"
        Me.txtCounts.Size = New System.Drawing.Size(91, 22)
        Me.txtCounts.TabIndex = 1
        Me.txtCounts.Text = "1"
        '
        'lblCountsTitle
        '
        Me.lblCountsTitle.AutoSize = True
        Me.lblCountsTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCountsTitle.Location = New System.Drawing.Point(324, 100)
        Me.lblCountsTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCountsTitle.Name = "lblCountsTitle"
        Me.lblCountsTitle.Size = New System.Drawing.Size(56, 18)
        Me.lblCountsTitle.TabIndex = 7
        Me.lblCountsTitle.Text = "Count&s"
        '
        'lblBuildTypeTitle
        '
        Me.lblBuildTypeTitle.AutoSize = True
        Me.lblBuildTypeTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBuildTypeTitle.Location = New System.Drawing.Point(19, 100)
        Me.lblBuildTypeTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBuildTypeTitle.Name = "lblBuildTypeTitle"
        Me.lblBuildTypeTitle.Size = New System.Drawing.Size(71, 18)
        Me.lblBuildTypeTitle.TabIndex = 5
        Me.lblBuildTypeTitle.Text = "&Build type"
        '
        'Panel2
        '
        Me.Panel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel2.Controls.Add(Me.lblBPTitle)
        Me.Panel2.Controls.Add(Me.lblBuildPhase)
        Me.Panel2.Location = New System.Drawing.Point(317, 59)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(211, 30)
        Me.Panel2.TabIndex = 4
        '
        'lblBPTitle
        '
        Me.lblBPTitle.AutoSize = True
        Me.lblBPTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBPTitle.Location = New System.Drawing.Point(7, 6)
        Me.lblBPTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBPTitle.Name = "lblBPTitle"
        Me.lblBPTitle.Size = New System.Drawing.Size(84, 18)
        Me.lblBPTitle.TabIndex = 0
        Me.lblBPTitle.Text = "Build phase"
        '
        'lblBuildPhase
        '
        Me.lblBuildPhase.AutoSize = True
        Me.lblBuildPhase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBuildPhase.ForeColor = System.Drawing.Color.Blue
        Me.lblBuildPhase.Location = New System.Drawing.Point(111, 6)
        Me.lblBuildPhase.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBuildPhase.Name = "lblBuildPhase"
        Me.lblBuildPhase.Size = New System.Drawing.Size(84, 18)
        Me.lblBuildPhase.TabIndex = 1
        Me.lblBuildPhase.Text = "Build phase"
        '
        'lblHCNTitle
        '
        Me.lblHCNTitle.AutoSize = True
        Me.lblHCNTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCNTitle.Location = New System.Drawing.Point(8, 6)
        Me.lblHCNTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblHCNTitle.Name = "lblHCNTitle"
        Me.lblHCNTitle.Size = New System.Drawing.Size(128, 18)
        Me.lblHCNTitle.TabIndex = 0
        Me.lblHCNTitle.Text = "Health chart name"
        '
        'lblHCName
        '
        Me.lblHCName.AutoSize = True
        Me.lblHCName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCName.ForeColor = System.Drawing.Color.Blue
        Me.lblHCName.Location = New System.Drawing.Point(160, 6)
        Me.lblHCName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblHCName.Name = "lblHCName"
        Me.lblHCName.Size = New System.Drawing.Size(128, 18)
        Me.lblHCName.TabIndex = 1
        Me.lblHCName.Text = "Health chart name"
        '
        'Panel1
        '
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Controls.Add(Me.lblHCIDTitle)
        Me.Panel1.Controls.Add(Me.lblHCID)
        Me.Panel1.Location = New System.Drawing.Point(11, 59)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(303, 30)
        Me.Panel1.TabIndex = 3
        '
        'lblHCIDTitle
        '
        Me.lblHCIDTitle.AutoSize = True
        Me.lblHCIDTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCIDTitle.Location = New System.Drawing.Point(8, 6)
        Me.lblHCIDTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblHCIDTitle.Name = "lblHCIDTitle"
        Me.lblHCIDTitle.Size = New System.Drawing.Size(105, 18)
        Me.lblHCIDTitle.TabIndex = 0
        Me.lblHCIDTitle.Text = "Health chart ID"
        '
        'lblHCID
        '
        Me.lblHCID.AutoSize = True
        Me.lblHCID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCID.ForeColor = System.Drawing.Color.Blue
        Me.lblHCID.Location = New System.Drawing.Point(160, 6)
        Me.lblHCID.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblHCID.Name = "lblHCID"
        Me.lblHCID.Size = New System.Drawing.Size(105, 18)
        Me.lblHCID.TabIndex = 1
        Me.lblHCID.Text = "Health chart ID"
        '
        'pnHealthChartName
        '
        Me.pnHealthChartName.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnHealthChartName.Controls.Add(Me.lblHCNTitle)
        Me.pnHealthChartName.Controls.Add(Me.lblHCName)
        Me.pnHealthChartName.Location = New System.Drawing.Point(11, 21)
        Me.pnHealthChartName.Margin = New System.Windows.Forms.Padding(4)
        Me.pnHealthChartName.Name = "pnHealthChartName"
        Me.pnHealthChartName.Size = New System.Drawing.Size(303, 30)
        Me.pnHealthChartName.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnAddUnit)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 155)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(539, 49)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'grpNewVeh
        '
        Me.grpNewVeh.Controls.Add(Me.lblBuildType)
        Me.grpNewVeh.Controls.Add(Me.txtCounts)
        Me.grpNewVeh.Controls.Add(Me.lblCountsTitle)
        Me.grpNewVeh.Controls.Add(Me.lblBuildTypeTitle)
        Me.grpNewVeh.Controls.Add(Me.Panel2)
        Me.grpNewVeh.Controls.Add(Me.Panel1)
        Me.grpNewVeh.Controls.Add(Me.pnHealthChartName)
        Me.grpNewVeh.Location = New System.Drawing.Point(10, 9)
        Me.grpNewVeh.Margin = New System.Windows.Forms.Padding(4)
        Me.grpNewVeh.Name = "grpNewVeh"
        Me.grpNewVeh.Padding = New System.Windows.Forms.Padding(4)
        Me.grpNewVeh.Size = New System.Drawing.Size(539, 138)
        Me.grpNewVeh.TabIndex = 5
        Me.grpNewVeh.TabStop = False
        Me.grpNewVeh.Text = "Units Count and Type"
        '
        'lblBuildType
        '
        Me.lblBuildType.AutoSize = True
        Me.lblBuildType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblBuildType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBuildType.Location = New System.Drawing.Point(171, 104)
        Me.lblBuildType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBuildType.Name = "lblBuildType"
        Me.lblBuildType.Size = New System.Drawing.Size(32, 20)
        Me.lblBuildType.TabIndex = 8
        Me.lblBuildType.Text = "Rig"
        '
        'frmNewRig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(559, 212)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpNewVeh)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNewRig"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add new unit to program"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pnHealthChartName.ResumeLayout(False)
        Me.pnHealthChartName.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.grpNewVeh.ResumeLayout(False)
        Me.grpNewVeh.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnAddUnit As System.Windows.Forms.Button
    Friend WithEvents txtCounts As System.Windows.Forms.TextBox
    Friend WithEvents lblCountsTitle As System.Windows.Forms.Label
    Friend WithEvents lblBuildTypeTitle As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblBPTitle As System.Windows.Forms.Label
    Friend WithEvents lblBuildPhase As System.Windows.Forms.Label
    Friend WithEvents lblHCNTitle As System.Windows.Forms.Label
    Friend WithEvents lblHCName As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblHCIDTitle As System.Windows.Forms.Label
    Friend WithEvents lblHCID As System.Windows.Forms.Label
    Friend WithEvents pnHealthChartName As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents grpNewVeh As System.Windows.Forms.GroupBox
    Friend WithEvents lblBuildType As System.Windows.Forms.Label
End Class
