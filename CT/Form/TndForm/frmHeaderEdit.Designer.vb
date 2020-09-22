<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmHeaderEdit
    Inherits frmBase 'System.Windows.Forms.Form

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
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnCancel2 = New System.Windows.Forms.Button()
        Me.btnReset2 = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lblHCID = New System.Windows.Forms.Label()
        Me.nudIssueMin = New System.Windows.Forms.NumericUpDown()
        Me.nudIssueM = New System.Windows.Forms.NumericUpDown()
        Me.txtProgDesc = New System.Windows.Forms.TextBox()
        Me.txtBuildPhase = New System.Windows.Forms.TextBox()
        Me.txtHardwareType = New System.Windows.Forms.TextBox()
        Me.lblHCIDTitle = New System.Windows.Forms.Label()
        Me.lblDecimal = New System.Windows.Forms.Label()
        Me.lblDescriptionTitle = New System.Windows.Forms.Label()
        Me.lblPlanIssueTitle = New System.Windows.Forms.Label()
        Me.lblBuildPhaseTitle = New System.Windows.Forms.Label()
        Me.lblHardwareTitle = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grdPermissions = New System.Windows.Forms.DataGridView()
        Me.pe04_TnDProgramAuthorization_PK = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pe27_Regions_FK = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pe10_SecurityLevel_FK = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CDSID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Regions = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.cmbSecurityLevel = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ProgramFunction = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnDelete = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.txtTndPlanner = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.nudIssueMin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudIssueM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.grdPermissions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOk.Location = New System.Drawing.Point(363, 252)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(84, 23)
        Me.btnOk.TabIndex = 11
        Me.btnOk.Text = "&Update"
        Me.ToolTip1.SetToolTip(Me.btnOk, "Add / Update [F7]")
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(453, 252)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Close the form [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(377, 251)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(85, 23)
        Me.btnSave.TabIndex = 23
        Me.btnSave.Text = "&Save changes"
        Me.ToolTip1.SetToolTip(Me.btnSave, "Add / Update [F7]")
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnCancel2
        '
        Me.btnCancel2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel2.Location = New System.Drawing.Point(468, 251)
        Me.btnCancel2.Name = "btnCancel2"
        Me.btnCancel2.Size = New System.Drawing.Size(60, 23)
        Me.btnCancel2.TabIndex = 24
        Me.btnCancel2.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel2, "Close the form [Esc]")
        Me.btnCancel2.UseVisualStyleBackColor = True
        '
        'btnReset2
        '
        Me.btnReset2.Location = New System.Drawing.Point(317, 251)
        Me.btnReset2.Name = "btnReset2"
        Me.btnReset2.Size = New System.Drawing.Size(54, 23)
        Me.btnReset2.TabIndex = 25
        Me.btnReset2.Text = "&Reset"
        Me.ToolTip1.SetToolTip(Me.btnReset2, "Add / Update [F7]")
        Me.btnReset2.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(542, 307)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TabPage1.Controls.Add(Me.lblHCID)
        Me.TabPage1.Controls.Add(Me.nudIssueMin)
        Me.TabPage1.Controls.Add(Me.btnOk)
        Me.TabPage1.Controls.Add(Me.nudIssueM)
        Me.TabPage1.Controls.Add(Me.btnCancel)
        Me.TabPage1.Controls.Add(Me.txtProgDesc)
        Me.TabPage1.Controls.Add(Me.txtBuildPhase)
        Me.TabPage1.Controls.Add(Me.txtHardwareType)
        Me.TabPage1.Controls.Add(Me.lblHCIDTitle)
        Me.TabPage1.Controls.Add(Me.lblDecimal)
        Me.TabPage1.Controls.Add(Me.lblDescriptionTitle)
        Me.TabPage1.Controls.Add(Me.lblPlanIssueTitle)
        Me.TabPage1.Controls.Add(Me.lblBuildPhaseTitle)
        Me.TabPage1.Controls.Add(Me.lblHardwareTitle)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(534, 281)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Header info"
        '
        'lblHCID
        '
        Me.lblHCID.BackColor = System.Drawing.SystemColors.Control
        Me.lblHCID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblHCID.Location = New System.Drawing.Point(118, 22)
        Me.lblHCID.Name = "lblHCID"
        Me.lblHCID.Size = New System.Drawing.Size(284, 20)
        Me.lblHCID.TabIndex = 1
        '
        'nudIssueMin
        '
        Me.nudIssueMin.Location = New System.Drawing.Point(168, 105)
        Me.nudIssueMin.Name = "nudIssueMin"
        Me.nudIssueMin.Size = New System.Drawing.Size(38, 20)
        Me.nudIssueMin.TabIndex = 5
        '
        'nudIssueM
        '
        Me.nudIssueM.Location = New System.Drawing.Point(118, 105)
        Me.nudIssueM.Name = "nudIssueM"
        Me.nudIssueM.Size = New System.Drawing.Size(40, 20)
        Me.nudIssueM.TabIndex = 4
        '
        'txtProgDesc
        '
        Me.txtProgDesc.Location = New System.Drawing.Point(118, 48)
        Me.txtProgDesc.MaxLength = 70
        Me.txtProgDesc.Name = "txtProgDesc"
        Me.txtProgDesc.Size = New System.Drawing.Size(284, 20)
        Me.txtProgDesc.TabIndex = 2
        '
        'txtBuildPhase
        '
        Me.txtBuildPhase.Location = New System.Drawing.Point(118, 76)
        Me.txtBuildPhase.MaxLength = 25
        Me.txtBuildPhase.Name = "txtBuildPhase"
        Me.txtBuildPhase.Size = New System.Drawing.Size(148, 20)
        Me.txtBuildPhase.TabIndex = 3
        '
        'txtHardwareType
        '
        Me.txtHardwareType.Location = New System.Drawing.Point(118, 132)
        Me.txtHardwareType.MaxLength = 50
        Me.txtHardwareType.Name = "txtHardwareType"
        Me.txtHardwareType.Size = New System.Drawing.Size(148, 20)
        Me.txtHardwareType.TabIndex = 6
        '
        'lblHCIDTitle
        '
        Me.lblHCIDTitle.AutoSize = True
        Me.lblHCIDTitle.Location = New System.Drawing.Point(12, 23)
        Me.lblHCIDTitle.Name = "lblHCIDTitle"
        Me.lblHCIDTitle.Size = New System.Drawing.Size(33, 13)
        Me.lblHCIDTitle.TabIndex = 12
        Me.lblHCIDTitle.Text = "HCID"
        '
        'lblDecimal
        '
        Me.lblDecimal.AutoSize = True
        Me.lblDecimal.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDecimal.Location = New System.Drawing.Point(155, 79)
        Me.lblDecimal.Name = "lblDecimal"
        Me.lblDecimal.Size = New System.Drawing.Size(16, 24)
        Me.lblDecimal.TabIndex = 17
        Me.lblDecimal.Text = "."
        '
        'lblDescriptionTitle
        '
        Me.lblDescriptionTitle.AutoSize = True
        Me.lblDescriptionTitle.Location = New System.Drawing.Point(12, 51)
        Me.lblDescriptionTitle.Name = "lblDescriptionTitle"
        Me.lblDescriptionTitle.Size = New System.Drawing.Size(100, 13)
        Me.lblDescriptionTitle.TabIndex = 13
        Me.lblDescriptionTitle.Text = "Program description"
        '
        'lblPlanIssueTitle
        '
        Me.lblPlanIssueTitle.AutoSize = True
        Me.lblPlanIssueTitle.Location = New System.Drawing.Point(12, 107)
        Me.lblPlanIssueTitle.Name = "lblPlanIssueTitle"
        Me.lblPlanIssueTitle.Size = New System.Drawing.Size(80, 13)
        Me.lblPlanIssueTitle.TabIndex = 16
        Me.lblPlanIssueTitle.Text = "T&&D Plan Issue"
        '
        'lblBuildPhaseTitle
        '
        Me.lblBuildPhaseTitle.AutoSize = True
        Me.lblBuildPhaseTitle.Location = New System.Drawing.Point(12, 79)
        Me.lblBuildPhaseTitle.Name = "lblBuildPhaseTitle"
        Me.lblBuildPhaseTitle.Size = New System.Drawing.Size(62, 13)
        Me.lblBuildPhaseTitle.TabIndex = 14
        Me.lblBuildPhaseTitle.Text = "Build phase"
        '
        'lblHardwareTitle
        '
        Me.lblHardwareTitle.AutoSize = True
        Me.lblHardwareTitle.Location = New System.Drawing.Point(12, 135)
        Me.lblHardwareTitle.Name = "lblHardwareTitle"
        Me.lblHardwareTitle.Size = New System.Drawing.Size(80, 13)
        Me.lblHardwareTitle.TabIndex = 15
        Me.lblHardwareTitle.Text = "Hardware Type"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.btnReset2)
        Me.TabPage2.Controls.Add(Me.btnSave)
        Me.TabPage2.Controls.Add(Me.btnCancel2)
        Me.TabPage2.Controls.Add(Me.grdPermissions)
        Me.TabPage2.Controls.Add(Me.txtTndPlanner)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(534, 281)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "User access"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(229, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(289, 26)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "* To delete a user, please select the row and press ""Delete"" button"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grdPermissions
        '
        Me.grdPermissions.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.grdPermissions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdPermissions.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.pe04_TnDProgramAuthorization_PK, Me.pe27_Regions_FK, Me.pe10_SecurityLevel_FK, Me.CDSID, Me.Regions, Me.cmbSecurityLevel, Me.ProgramFunction, Me.btnDelete})
        Me.grdPermissions.Location = New System.Drawing.Point(6, 41)
        Me.grdPermissions.Name = "grdPermissions"
        Me.grdPermissions.Size = New System.Drawing.Size(522, 204)
        Me.grdPermissions.TabIndex = 22
        '
        'pe04_TnDProgramAuthorization_PK
        '
        Me.pe04_TnDProgramAuthorization_PK.DataPropertyName = "pe04_TnDProgramAuthorization_PK"
        Me.pe04_TnDProgramAuthorization_PK.HeaderText = "pe04_TnDProgramAuthorization_PK"
        Me.pe04_TnDProgramAuthorization_PK.Name = "pe04_TnDProgramAuthorization_PK"
        Me.pe04_TnDProgramAuthorization_PK.Visible = False
        '
        'pe27_Regions_FK
        '
        Me.pe27_Regions_FK.DataPropertyName = "pe27_Regions_FK"
        Me.pe27_Regions_FK.HeaderText = "pe27_Regions_FK"
        Me.pe27_Regions_FK.Name = "pe27_Regions_FK"
        Me.pe27_Regions_FK.Visible = False
        '
        'pe10_SecurityLevel_FK
        '
        Me.pe10_SecurityLevel_FK.DataPropertyName = "pe10_SecurityLevel_FK"
        Me.pe10_SecurityLevel_FK.HeaderText = "pe10_SecurityLevel_FK"
        Me.pe10_SecurityLevel_FK.Name = "pe10_SecurityLevel_FK"
        Me.pe10_SecurityLevel_FK.Visible = False
        '
        'CDSID
        '
        Me.CDSID.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.CDSID.DataPropertyName = "CDSID"
        Me.CDSID.FillWeight = 75.0!
        Me.CDSID.HeaderText = "CDSID"
        Me.CDSID.MaxInputLength = 16
        Me.CDSID.Name = "CDSID"
        '
        'Regions
        '
        Me.Regions.DataPropertyName = "pe27_Regions_FK"
        Me.Regions.FillWeight = 40.0!
        Me.Regions.HeaderText = "Regions"
        Me.Regions.Name = "Regions"
        '
        'cmbSecurityLevel
        '
        Me.cmbSecurityLevel.DataPropertyName = "pe10_SecurityLevel_FK"
        Me.cmbSecurityLevel.FillWeight = 60.0!
        Me.cmbSecurityLevel.HeaderText = "Security Level"
        Me.cmbSecurityLevel.Name = "cmbSecurityLevel"
        '
        'ProgramFunction
        '
        Me.ProgramFunction.DataPropertyName = "ProgramFunction"
        Me.ProgramFunction.FillWeight = 85.0!
        Me.ProgramFunction.HeaderText = "Program Function"
        Me.ProgramFunction.MaxInputLength = 50
        Me.ProgramFunction.Name = "ProgramFunction"
        '
        'btnDelete
        '
        Me.btnDelete.FillWeight = 25.0!
        Me.btnDelete.HeaderText = "Delete"
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Text = "X"
        Me.btnDelete.UseColumnTextForButtonValue = True
        '
        'txtTndPlanner
        '
        Me.txtTndPlanner.Location = New System.Drawing.Point(79, 14)
        Me.txtTndPlanner.MaxLength = 16
        Me.txtTndPlanner.Name = "txtTndPlanner"
        Me.txtTndPlanner.Size = New System.Drawing.Size(148, 20)
        Me.txtTndPlanner.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "TnD Planner"
        '
        'frmHeaderEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(563, 320)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHeaderEdit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Edit header information"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.nudIssueMin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudIssueM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.grdPermissions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblHCID As System.Windows.Forms.Label
    Friend WithEvents nudIssueMin As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudIssueM As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtProgDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtBuildPhase As System.Windows.Forms.TextBox
    Friend WithEvents txtHardwareType As System.Windows.Forms.TextBox
    Friend WithEvents lblDecimal As System.Windows.Forms.Label
    Friend WithEvents lblPlanIssueTitle As System.Windows.Forms.Label
    Friend WithEvents lblHardwareTitle As System.Windows.Forms.Label
    Friend WithEvents lblBuildPhaseTitle As System.Windows.Forms.Label
    Friend WithEvents lblDescriptionTitle As System.Windows.Forms.Label
    Friend WithEvents lblHCIDTitle As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents grdPermissions As System.Windows.Forms.DataGridView
    Friend WithEvents txtTndPlanner As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel2 As System.Windows.Forms.Button
    Friend WithEvents btnReset2 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pe04_TnDProgramAuthorization_PK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pe27_Regions_FK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pe10_SecurityLevel_FK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CDSID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Regions As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents cmbSecurityLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ProgramFunction As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnDelete As System.Windows.Forms.DataGridViewButtonColumn
End Class
