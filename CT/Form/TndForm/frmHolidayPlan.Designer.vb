<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmHolidayPlan
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHolidayPlan))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Lbl_Specific_Total = New System.Windows.Forms.Label()
        Me.Lbl_Generic_Total = New System.Windows.Forms.Label()
        Me.btnAddRow = New System.Windows.Forms.Button()
        Me.dgvSpecific = New System.Windows.Forms.DataGridView()
        Me.HolidayName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pe83 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HolidayTypeKey = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgvDefault = New System.Windows.Forms.DataGridView()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblBuildPhase = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblHCID = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblHCName = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvSpecific, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDefault, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Lbl_Specific_Total)
        Me.GroupBox1.Controls.Add(Me.Lbl_Generic_Total)
        Me.GroupBox1.Controls.Add(Me.btnAddRow)
        Me.GroupBox1.Controls.Add(Me.dgvSpecific)
        Me.GroupBox1.Controls.Add(Me.dgvDefault)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Controls.Add(Me.btnRemove)
        Me.GroupBox1.Controls.Add(Me.btnAdd)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.lblBuildPhase)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.lblHCID)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.lblHCName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 5)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(886, 604)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Custom Plan Holidays"
        '
        'Lbl_Specific_Total
        '
        Me.Lbl_Specific_Total.AutoSize = True
        Me.Lbl_Specific_Total.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Lbl_Specific_Total.Location = New System.Drawing.Point(829, 566)
        Me.Lbl_Specific_Total.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lbl_Specific_Total.Name = "Lbl_Specific_Total"
        Me.Lbl_Specific_Total.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Lbl_Specific_Total.Size = New System.Drawing.Size(40, 13)
        Me.Lbl_Specific_Total.TabIndex = 20
        Me.Lbl_Specific_Total.Text = "Total : "
        Me.Lbl_Specific_Total.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Lbl_Generic_Total
        '
        Me.Lbl_Generic_Total.AutoSize = True
        Me.Lbl_Generic_Total.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Lbl_Generic_Total.Location = New System.Drawing.Point(829, 302)
        Me.Lbl_Generic_Total.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lbl_Generic_Total.Name = "Lbl_Generic_Total"
        Me.Lbl_Generic_Total.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Lbl_Generic_Total.Size = New System.Drawing.Size(40, 13)
        Me.Lbl_Generic_Total.TabIndex = 19
        Me.Lbl_Generic_Total.Text = "Total : "
        Me.Lbl_Generic_Total.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnAddRow
        '
        Me.btnAddRow.Location = New System.Drawing.Point(538, 309)
        Me.btnAddRow.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnAddRow.Name = "btnAddRow"
        Me.btnAddRow.Size = New System.Drawing.Size(152, 23)
        Me.btnAddRow.TabIndex = 3
        Me.btnAddRow.Text = "Add Holiday for &Manual Entry"
        Me.ToolTip1.SetToolTip(Me.btnAddRow, "[F6]")
        Me.btnAddRow.UseVisualStyleBackColor = True
        '
        'dgvSpecific
        '
        Me.dgvSpecific.AllowUserToAddRows = False
        Me.dgvSpecific.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvSpecific.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSpecific.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.HolidayName, Me.pe83, Me.HolidayTypeKey})
        Me.dgvSpecific.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvSpecific.Location = New System.Drawing.Point(14, 351)
        Me.dgvSpecific.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvSpecific.Name = "dgvSpecific"
        Me.dgvSpecific.RowHeadersVisible = False
        Me.dgvSpecific.RowTemplate.Height = 24
        Me.dgvSpecific.Size = New System.Drawing.Size(857, 212)
        Me.dgvSpecific.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.dgvSpecific, resources.GetString("dgvSpecific.ToolTip"))
        '
        'HolidayName
        '
        Me.HolidayName.HeaderText = "Holiday Name"
        Me.HolidayName.Name = "HolidayName"
        '
        'pe83
        '
        Me.pe83.HeaderText = "pe83"
        Me.pe83.Name = "pe83"
        Me.pe83.Visible = False
        '
        'HolidayTypeKey
        '
        Me.HolidayTypeKey.HeaderText = "HolidayTypeKey"
        Me.HolidayTypeKey.Name = "HolidayTypeKey"
        Me.HolidayTypeKey.Visible = False
        '
        'dgvDefault
        '
        Me.dgvDefault.AllowUserToAddRows = False
        Me.dgvDefault.AllowUserToDeleteRows = False
        Me.dgvDefault.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvDefault.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDefault.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvDefault.Location = New System.Drawing.Point(14, 88)
        Me.dgvDefault.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgvDefault.Name = "dgvDefault"
        Me.dgvDefault.RowHeadersVisible = False
        Me.dgvDefault.RowTemplate.Height = 24
        Me.dgvDefault.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDefault.Size = New System.Drawing.Size(857, 212)
        Me.dgvDefault.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.dgvDefault, "* Select one or more holiday(s) and click 'Add Generic to Specific'" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "* F2 to focu" &
        "s the grid")
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(718, 570)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Close"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Cancel [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(632, 570)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(72, 23)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "&Populate"
        Me.ToolTip1.SetToolTip(Me.btnSave, "Save [F7]")
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(367, 309)
        Me.btnRemove.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(152, 23)
        Me.btnRemove.TabIndex = 2
        Me.btnRemove.Text = "&Remove from Specific"
        Me.ToolTip1.SetToolTip(Me.btnRemove, "[F5]")
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(195, 309)
        Me.btnAdd.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(152, 23)
        Me.btnAdd.TabIndex = 1
        Me.btnAdd.Text = "&Add Generic to Specific"
        Me.ToolTip1.SetToolTip(Me.btnAdd, "[F4]")
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 331)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Specific TnD Plan Holidays"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 67)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(87, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Generic Holidays"
        '
        'lblBuildPhase
        '
        Me.lblBuildPhase.AutoSize = True
        Me.lblBuildPhase.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblBuildPhase.Location = New System.Drawing.Point(529, 24)
        Me.lblBuildPhase.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblBuildPhase.Name = "lblBuildPhase"
        Me.lblBuildPhase.Size = New System.Drawing.Size(70, 13)
        Me.lblBuildPhase.TabIndex = 7
        Me.lblBuildPhase.Text = "lblBuildPhase"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(455, 24)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Build Phase : "
        '
        'lblHCID
        '
        Me.lblHCID.AutoSize = True
        Me.lblHCID.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblHCID.Location = New System.Drawing.Point(114, 46)
        Me.lblHCID.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblHCID.Name = "lblHCID"
        Me.lblHCID.Size = New System.Drawing.Size(43, 13)
        Me.lblHCID.TabIndex = 5
        Me.lblHCID.Text = "lblHCID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 46)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(89, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Health Chart ID : "
        '
        'lblHCName
        '
        Me.lblHCName.AutoSize = True
        Me.lblHCName.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblHCName.Location = New System.Drawing.Point(114, 24)
        Me.lblHCName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblHCName.Name = "lblHCName"
        Me.lblHCName.Size = New System.Drawing.Size(60, 13)
        Me.lblHCName.TabIndex = 3
        Me.lblHCName.Text = "lblHCName"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 24)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Health Chart Name : "
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Holiday Name"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 1140
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "pe83"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Visible = False
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "HolidayTypeKey"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Visible = False
        '
        'frmHolidayPlan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(892, 613)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHolidayPlan"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Custom Plan Holidays"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvSpecific, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDefault, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblBuildPhase As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblHCID As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblHCName As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dgvDefault As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSpecific As System.Windows.Forms.DataGridView
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnAddRow As System.Windows.Forms.Button
    Friend WithEvents HolidayName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pe83 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HolidayTypeKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Lbl_Specific_Total As System.Windows.Forms.Label
    Friend WithEvents Lbl_Generic_Total As System.Windows.Forms.Label
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
