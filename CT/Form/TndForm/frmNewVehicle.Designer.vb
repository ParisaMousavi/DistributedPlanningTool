<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmNewVehicle
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
        Me.grpNewVeh = New System.Windows.Forms.GroupBox()
        Me.txtCounts = New System.Windows.Forms.TextBox()
        Me.lblCountsTitle = New System.Windows.Forms.Label()
        Me.cboBuildType = New System.Windows.Forms.ComboBox()
        Me.lblBuildTypeTitle = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblBPTitle = New System.Windows.Forms.Label()
        Me.lblBuildPhase = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblHCIDTitle = New System.Windows.Forms.Label()
        Me.lblHCID = New System.Windows.Forms.Label()
        Me.pnHealthChartName = New System.Windows.Forms.Panel()
        Me.lblHCNTitle = New System.Windows.Forms.Label()
        Me.lblHCName = New System.Windows.Forms.Label()
        Me.btnAddUnit = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpNewVeh.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pnHealthChartName.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpNewVeh
        '
        Me.grpNewVeh.Controls.Add(Me.txtCounts)
        Me.grpNewVeh.Controls.Add(Me.lblCountsTitle)
        Me.grpNewVeh.Controls.Add(Me.cboBuildType)
        Me.grpNewVeh.Controls.Add(Me.lblBuildTypeTitle)
        Me.grpNewVeh.Controls.Add(Me.Panel2)
        Me.grpNewVeh.Controls.Add(Me.Panel1)
        Me.grpNewVeh.Controls.Add(Me.pnHealthChartName)
        Me.grpNewVeh.Location = New System.Drawing.Point(7, 6)
        Me.grpNewVeh.Name = "grpNewVeh"
        Me.grpNewVeh.Size = New System.Drawing.Size(404, 112)
        Me.grpNewVeh.TabIndex = 0
        Me.grpNewVeh.TabStop = False
        Me.grpNewVeh.Text = "Units Count and Type"
        '
        'txtCounts
        '
        Me.txtCounts.Location = New System.Drawing.Point(323, 81)
        Me.txtCounts.Name = "txtCounts"
        Me.txtCounts.Size = New System.Drawing.Size(69, 20)
        Me.txtCounts.TabIndex = 1
        Me.txtCounts.Text = "1"
        '
        'lblCountsTitle
        '
        Me.lblCountsTitle.AutoSize = True
        Me.lblCountsTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCountsTitle.Location = New System.Drawing.Point(243, 81)
        Me.lblCountsTitle.Name = "lblCountsTitle"
        Me.lblCountsTitle.Size = New System.Drawing.Size(45, 15)
        Me.lblCountsTitle.TabIndex = 7
        Me.lblCountsTitle.Text = "Count&s"
        '
        'cboBuildType
        '
        Me.cboBuildType.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cboBuildType.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboBuildType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBuildType.DropDownWidth = 107
        Me.cboBuildType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboBuildType.FormattingEnabled = True
        Me.cboBuildType.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cboBuildType.Location = New System.Drawing.Point(130, 81)
        Me.cboBuildType.Name = "cboBuildType"
        Me.cboBuildType.Size = New System.Drawing.Size(107, 23)
        Me.cboBuildType.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cboBuildType, "Build type [F4]")
        '
        'lblBuildTypeTitle
        '
        Me.lblBuildTypeTitle.AutoSize = True
        Me.lblBuildTypeTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBuildTypeTitle.Location = New System.Drawing.Point(14, 81)
        Me.lblBuildTypeTitle.Name = "lblBuildTypeTitle"
        Me.lblBuildTypeTitle.Size = New System.Drawing.Size(60, 15)
        Me.lblBuildTypeTitle.TabIndex = 5
        Me.lblBuildTypeTitle.Text = "&Build type"
        '
        'Panel2
        '
        Me.Panel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel2.Controls.Add(Me.lblBPTitle)
        Me.Panel2.Controls.Add(Me.lblBuildPhase)
        Me.Panel2.Location = New System.Drawing.Point(238, 48)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(158, 24)
        Me.Panel2.TabIndex = 4
        '
        'lblBPTitle
        '
        Me.lblBPTitle.AutoSize = True
        Me.lblBPTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBPTitle.Location = New System.Drawing.Point(5, 5)
        Me.lblBPTitle.Name = "lblBPTitle"
        Me.lblBPTitle.Size = New System.Drawing.Size(72, 15)
        Me.lblBPTitle.TabIndex = 0
        Me.lblBPTitle.Text = "Build phase"
        '
        'lblBuildPhase
        '
        Me.lblBuildPhase.AutoSize = True
        Me.lblBuildPhase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBuildPhase.ForeColor = System.Drawing.Color.Blue
        Me.lblBuildPhase.Location = New System.Drawing.Point(83, 5)
        Me.lblBuildPhase.Name = "lblBuildPhase"
        Me.lblBuildPhase.Size = New System.Drawing.Size(72, 15)
        Me.lblBuildPhase.TabIndex = 1
        Me.lblBuildPhase.Text = "Build phase"
        '
        'Panel1
        '
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Controls.Add(Me.lblHCIDTitle)
        Me.Panel1.Controls.Add(Me.lblHCID)
        Me.Panel1.Location = New System.Drawing.Point(8, 48)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(227, 24)
        Me.Panel1.TabIndex = 3
        '
        'lblHCIDTitle
        '
        Me.lblHCIDTitle.AutoSize = True
        Me.lblHCIDTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCIDTitle.Location = New System.Drawing.Point(6, 5)
        Me.lblHCIDTitle.Name = "lblHCIDTitle"
        Me.lblHCIDTitle.Size = New System.Drawing.Size(88, 15)
        Me.lblHCIDTitle.TabIndex = 0
        Me.lblHCIDTitle.Text = "Health chart ID"
        '
        'lblHCID
        '
        Me.lblHCID.AutoSize = True
        Me.lblHCID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCID.ForeColor = System.Drawing.Color.Blue
        Me.lblHCID.Location = New System.Drawing.Point(120, 5)
        Me.lblHCID.Name = "lblHCID"
        Me.lblHCID.Size = New System.Drawing.Size(88, 15)
        Me.lblHCID.TabIndex = 1
        Me.lblHCID.Text = "Health chart ID"
        '
        'pnHealthChartName
        '
        Me.pnHealthChartName.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnHealthChartName.Controls.Add(Me.lblHCNTitle)
        Me.pnHealthChartName.Controls.Add(Me.lblHCName)
        Me.pnHealthChartName.Location = New System.Drawing.Point(8, 17)
        Me.pnHealthChartName.Name = "pnHealthChartName"
        Me.pnHealthChartName.Size = New System.Drawing.Size(227, 24)
        Me.pnHealthChartName.TabIndex = 2
        '
        'lblHCNTitle
        '
        Me.lblHCNTitle.AutoSize = True
        Me.lblHCNTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCNTitle.Location = New System.Drawing.Point(6, 5)
        Me.lblHCNTitle.Name = "lblHCNTitle"
        Me.lblHCNTitle.Size = New System.Drawing.Size(108, 15)
        Me.lblHCNTitle.TabIndex = 0
        Me.lblHCNTitle.Text = "Health chart name"
        '
        'lblHCName
        '
        Me.lblHCName.AutoSize = True
        Me.lblHCName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHCName.ForeColor = System.Drawing.Color.Blue
        Me.lblHCName.Location = New System.Drawing.Point(120, 5)
        Me.lblHCName.Name = "lblHCName"
        Me.lblHCName.Size = New System.Drawing.Size(108, 15)
        Me.lblHCName.TabIndex = 1
        Me.lblHCName.Text = "Health chart name"
        '
        'btnAddUnit
        '
        Me.btnAddUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddUnit.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnAddUnit.Location = New System.Drawing.Point(238, 11)
        Me.btnAddUnit.Name = "btnAddUnit"
        Me.btnAddUnit.Size = New System.Drawing.Size(75, 23)
        Me.btnAddUnit.TabIndex = 0
        Me.btnAddUnit.Text = "&Add Unit"
        Me.ToolTip1.SetToolTip(Me.btnAddUnit, "Add Unit [F7]")
        Me.btnAddUnit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(319, 11)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Close the form [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnAddUnit)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 124)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(404, 40)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'frmNewVehicle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(416, 169)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpNewVeh)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNewVehicle"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add new unit to program"
        Me.grpNewVeh.ResumeLayout(False)
        Me.grpNewVeh.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pnHealthChartName.ResumeLayout(False)
        Me.pnHealthChartName.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grpNewVeh As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblHCIDTitle As System.Windows.Forms.Label
    Friend WithEvents lblHCID As System.Windows.Forms.Label
    Friend WithEvents pnHealthChartName As System.Windows.Forms.Panel
    Friend WithEvents lblHCNTitle As System.Windows.Forms.Label
    Friend WithEvents lblHCName As System.Windows.Forms.Label
    Friend WithEvents lblCountsTitle As System.Windows.Forms.Label
    Friend WithEvents cboBuildType As System.Windows.Forms.ComboBox
    Friend WithEvents lblBuildTypeTitle As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblBPTitle As System.Windows.Forms.Label
    Friend WithEvents lblBuildPhase As System.Windows.Forms.Label
    Friend WithEvents txtCounts As System.Windows.Forms.TextBox
    Friend WithEvents btnAddUnit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
