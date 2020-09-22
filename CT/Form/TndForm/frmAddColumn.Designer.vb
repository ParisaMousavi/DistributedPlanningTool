<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAddColumn
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.grbHeaderDescription = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtHeader = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.lstColumns = New System.Windows.Forms.ListBox()
        Me.cboSections = New System.Windows.Forms.ComboBox()
        Me.cboFeatureGroup = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.grbHeaderDescription.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.grbHeaderDescription)
        Me.GroupBox1.Controls.Add(Me.lstColumns)
        Me.GroupBox1.Controls.Add(Me.cboSections)
        Me.GroupBox1.Controls.Add(Me.cboFeatureGroup)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(901, 521)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 496)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(297, 17)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "* Use 'Ctrl' key to select one or more columns."
        '
        'grbHeaderDescription
        '
        Me.grbHeaderDescription.Controls.Add(Me.cmdCancel)
        Me.grbHeaderDescription.Controls.Add(Me.cmdOk)
        Me.grbHeaderDescription.Controls.Add(Me.txtDescription)
        Me.grbHeaderDescription.Controls.Add(Me.txtHeader)
        Me.grbHeaderDescription.Controls.Add(Me.Label5)
        Me.grbHeaderDescription.Controls.Add(Me.Label4)
        Me.grbHeaderDescription.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grbHeaderDescription.Location = New System.Drawing.Point(261, 207)
        Me.grbHeaderDescription.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grbHeaderDescription.Name = "grbHeaderDescription"
        Me.grbHeaderDescription.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grbHeaderDescription.Size = New System.Drawing.Size(408, 177)
        Me.grbHeaderDescription.TabIndex = 10
        Me.grbHeaderDescription.TabStop = False
        Me.grbHeaderDescription.Text = "Header && Description"
        Me.grbHeaderDescription.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(283, 134)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(107, 28)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOk
        '
        Me.cmdOk.Location = New System.Drawing.Point(159, 134)
        Me.cmdOk.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(107, 28)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(109, 87)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(280, 22)
        Me.txtDescription.TabIndex = 1
        '
        'txtHeader
        '
        Me.txtHeader.Location = New System.Drawing.Point(109, 36)
        Me.txtHeader.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtHeader.Name = "txtHeader"
        Me.txtHeader.Size = New System.Drawing.Size(280, 22)
        Me.txtHeader.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 87)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 17)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Description : "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 36)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 17)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Header : "
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(768, 16)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(107, 28)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.Text = "&Close"
        Me.ToolTip1.SetToolTip(Me.cmdClose, "Close the form [Esc]")
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Location = New System.Drawing.Point(651, 16)
        Me.cmdDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(107, 28)
        Me.cmdDelete.TabIndex = 5
        Me.cmdDelete.Text = "&Delete"
        Me.ToolTip1.SetToolTip(Me.cmdDelete, "* Delete column [F9]." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "* Use 'Ctrl' key to select one or more columns.")
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Location = New System.Drawing.Point(534, 16)
        Me.cmdUpdate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(107, 28)
        Me.cmdUpdate.TabIndex = 4
        Me.cmdUpdate.Text = "&Rename"
        Me.ToolTip1.SetToolTip(Me.cmdUpdate, "Rename column [F8]")
        Me.cmdUpdate.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(416, 16)
        Me.cmdAdd.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(107, 28)
        Me.cmdAdd.TabIndex = 3
        Me.cmdAdd.Text = "&Add a feature"
        Me.ToolTip1.SetToolTip(Me.cmdAdd, "Add column [F7]")
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'lstColumns
        '
        Me.lstColumns.FormattingEnabled = True
        Me.lstColumns.ItemHeight = 16
        Me.lstColumns.Location = New System.Drawing.Point(16, 103)
        Me.lstColumns.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.lstColumns.Name = "lstColumns"
        Me.lstColumns.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstColumns.Size = New System.Drawing.Size(871, 388)
        Me.lstColumns.TabIndex = 2
        '
        'cboSections
        '
        Me.cboSections.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSections.FormattingEnabled = True
        Me.cboSections.Location = New System.Drawing.Point(476, 46)
        Me.cboSections.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboSections.Name = "cboSections"
        Me.cboSections.Size = New System.Drawing.Size(411, 24)
        Me.cboSections.TabIndex = 1
        '
        'cboFeatureGroup
        '
        Me.cboFeatureGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFeatureGroup.FormattingEnabled = True
        Me.cboFeatureGroup.Items.AddRange(New Object() {"Instrumentation", "Non MFC specification", "MFC specification", "Program Information", "Further Basic Specification", "User & shipping details", "Update pack"})
        Me.cboFeatureGroup.Location = New System.Drawing.Point(16, 46)
        Me.cboFeatureGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboFeatureGroup.Name = "cboFeatureGroup"
        Me.cboFeatureGroup.Size = New System.Drawing.Size(411, 24)
        Me.cboFeatureGroup.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cboFeatureGroup, "Features [F4]")
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 79)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Column Names"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(472, 18)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Section"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 17)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Features Group"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdClose)
        Me.GroupBox2.Controls.Add(Me.cmdAdd)
        Me.GroupBox2.Controls.Add(Me.cmdDelete)
        Me.GroupBox2.Controls.Add(Me.cmdUpdate)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 523)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(901, 58)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        '
        'frmAddColumn
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(914, 584)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddColumn"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Columns Add/Update/Delete"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grbHeaderDescription.ResumeLayout(False)
        Me.grbHeaderDescription.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboSections As System.Windows.Forms.ComboBox
    Friend WithEvents cboFeatureGroup As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents grbHeaderDescription As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtHeader As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lstColumns As System.Windows.Forms.ListBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
End Class
