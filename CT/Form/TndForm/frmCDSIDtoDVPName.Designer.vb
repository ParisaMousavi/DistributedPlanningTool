<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCDSIDtoDVPName
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
        Me.components = New System.ComponentModel.Container()
        Me.lblHCID = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dgvAssignCDS = New System.Windows.Forms.DataGridView()
        Me.PMTGroup = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DVPTeamname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PMTLevel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DNRLevel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Edited = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnPhonebook = New System.Windows.Forms.Button()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.GrbMain = New System.Windows.Forms.GroupBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GrbPhonebook = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ChkListPhonebook = New System.Windows.Forms.CheckedListBox()
        Me.cbPMTLevel = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.dgvAssignCDS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GrbMain.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GrbPhonebook.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblHCID
        '
        Me.lblHCID.AutoSize = True
        Me.lblHCID.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblHCID.Location = New System.Drawing.Point(150, 22)
        Me.lblHCID.Name = "lblHCID"
        Me.lblHCID.Size = New System.Drawing.Size(54, 17)
        Me.lblHCID.TabIndex = 7
        Me.lblHCID.Text = "lblHCID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(14, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(116, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Health Chart ID : "
        '
        'dgvAssignCDS
        '
        Me.dgvAssignCDS.AllowUserToAddRows = False
        Me.dgvAssignCDS.AllowUserToDeleteRows = False
        Me.dgvAssignCDS.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvAssignCDS.ColumnHeadersHeight = 30
        Me.dgvAssignCDS.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PMTGroup, Me.DVPTeamname, Me.PMTLevel, Me.DNRLevel, Me.Edited})
        Me.dgvAssignCDS.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvAssignCDS.Location = New System.Drawing.Point(17, 56)
        Me.dgvAssignCDS.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgvAssignCDS.Name = "dgvAssignCDS"
        Me.dgvAssignCDS.RowTemplate.Height = 24
        Me.dgvAssignCDS.Size = New System.Drawing.Size(1301, 476)
        Me.dgvAssignCDS.TabIndex = 10
        '
        'PMTGroup
        '
        Me.PMTGroup.FillWeight = 35.0!
        Me.PMTGroup.HeaderText = "PMT Group"
        Me.PMTGroup.Name = "PMTGroup"
        Me.PMTGroup.ReadOnly = True
        '
        'DVPTeamname
        '
        Me.DVPTeamname.FillWeight = 50.0!
        Me.DVPTeamname.HeaderText = "Global DVP Team name"
        Me.DVPTeamname.Name = "DVPTeamname"
        Me.DVPTeamname.ReadOnly = True
        '
        'PMTLevel
        '
        Me.PMTLevel.FillWeight = 55.0!
        Me.PMTLevel.HeaderText = "PMT Level"
        Me.PMTLevel.Name = "PMTLevel"
        Me.PMTLevel.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'DNRLevel
        '
        Me.DNRLevel.FillWeight = 55.0!
        Me.DNRLevel.HeaderText = "DNRLevel"
        Me.DNRLevel.Name = "DNRLevel"
        '
        'Edited
        '
        Me.Edited.FillWeight = 10.0!
        Me.Edited.HeaderText = "Edited"
        Me.Edited.Name = "Edited"
        Me.Edited.ReadOnly = True
        Me.Edited.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(1158, 15)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(150, 28)
        Me.btnCancel.TabIndex = 15
        Me.btnCancel.Text = "&Close"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Close [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(1000, 15)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(150, 28)
        Me.btnSave.TabIndex = 14
        Me.btnSave.Text = "&Save"
        Me.ToolTip1.SetToolTip(Me.btnSave, "Save [F7]")
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnPhonebook
        '
        Me.btnPhonebook.Location = New System.Drawing.Point(7, 15)
        Me.btnPhonebook.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPhonebook.Name = "btnPhonebook"
        Me.btnPhonebook.Size = New System.Drawing.Size(150, 28)
        Me.btnPhonebook.TabIndex = 16
        Me.btnPhonebook.Text = "&Phone book"
        Me.ToolTip1.SetToolTip(Me.btnPhonebook, "Save [F7]")
        Me.btnPhonebook.UseVisualStyleBackColor = True
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(165, 387)
        Me.btnInsert.Margin = New System.Windows.Forms.Padding(4)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(124, 36)
        Me.btnInsert.TabIndex = 15
        Me.btnInsert.Text = "Se&lect"
        Me.ToolTip1.SetToolTip(Me.btnInsert, "Save [F7]")
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'GrbMain
        '
        Me.GrbMain.Controls.Add(Me.Panel1)
        Me.GrbMain.Controls.Add(Me.cbPMTLevel)
        Me.GrbMain.Controls.Add(Me.Label6)
        Me.GrbMain.Controls.Add(Me.lblHCID)
        Me.GrbMain.Controls.Add(Me.dgvAssignCDS)
        Me.GrbMain.Location = New System.Drawing.Point(7, 7)
        Me.GrbMain.Name = "GrbMain"
        Me.GrbMain.Size = New System.Drawing.Size(1335, 549)
        Me.GrbMain.TabIndex = 16
        Me.GrbMain.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GrbPhonebook)
        Me.Panel1.Location = New System.Drawing.Point(512, 89)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(310, 440)
        Me.Panel1.TabIndex = 18
        Me.Panel1.Visible = False
        '
        'GrbPhonebook
        '
        Me.GrbPhonebook.Controls.Add(Me.Label1)
        Me.GrbPhonebook.Controls.Add(Me.btnInsert)
        Me.GrbPhonebook.Controls.Add(Me.ChkListPhonebook)
        Me.GrbPhonebook.Location = New System.Drawing.Point(3, 3)
        Me.GrbPhonebook.Name = "GrbPhonebook"
        Me.GrbPhonebook.Size = New System.Drawing.Size(300, 430)
        Me.GrbPhonebook.TabIndex = 17
        Me.GrbPhonebook.TabStop = False
        Me.GrbPhonebook.Text = "DNRLevel CDSID selection"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 406)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(132, 17)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Press 'Esc' to Close"
        '
        'ChkListPhonebook
        '
        Me.ChkListPhonebook.CheckOnClick = True
        Me.ChkListPhonebook.FormattingEnabled = True
        Me.ChkListPhonebook.Location = New System.Drawing.Point(11, 23)
        Me.ChkListPhonebook.Name = "ChkListPhonebook"
        Me.ChkListPhonebook.Size = New System.Drawing.Size(278, 361)
        Me.ChkListPhonebook.TabIndex = 0
        '
        'cbPMTLevel
        '
        Me.cbPMTLevel.FormattingEnabled = True
        Me.cbPMTLevel.Location = New System.Drawing.Point(702, 22)
        Me.cbPMTLevel.Name = "cbPMTLevel"
        Me.cbPMTLevel.Size = New System.Drawing.Size(121, 24)
        Me.cbPMTLevel.TabIndex = 17
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(161, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(469, 17)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "New CDSID should be inserted first in Phone book via Phone book button"
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.FillWeight = 35.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "PMT Group"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 123
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn2.HeaderText = "Global DVP Team name"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 351
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.FillWeight = 35.0!
        Me.DataGridViewTextBoxColumn3.HeaderText = "PMT Level"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewTextBoxColumn3.Width = 123
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.FillWeight = 35.0!
        Me.DataGridViewTextBoxColumn4.HeaderText = "DNRLevel"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        Me.DataGridViewTextBoxColumn4.Visible = False
        Me.DataGridViewTextBoxColumn4.Width = 122
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.FillWeight = 10.0!
        Me.DataGridViewTextBoxColumn5.HeaderText = "Edited"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        Me.DataGridViewTextBoxColumn5.Visible = False
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Text = "NotifyIcon1"
        Me.NotifyIcon1.Visible = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Controls.Add(Me.btnPhonebook)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 562)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1335, 53)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        '
        'frmCDSIDtoDVPName
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1348, 626)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GrbMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCDSIDtoDVPName"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Assign CDSID to DvpTeam"
        CType(Me.dgvAssignCDS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GrbMain.ResumeLayout(False)
        Me.GrbMain.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.GrbPhonebook.ResumeLayout(False)
        Me.GrbPhonebook.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblHCID As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dgvAssignCDS As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GrbMain As System.Windows.Forms.GroupBox
    Friend WithEvents PMTGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DVPTeamname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PMTLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DNRLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Edited As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnPhonebook As System.Windows.Forms.Button
    Friend WithEvents cbPMTLevel As System.Windows.Forms.ComboBox
    Friend WithEvents GrbPhonebook As System.Windows.Forms.GroupBox
    Friend WithEvents ChkListPhonebook As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnInsert As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
