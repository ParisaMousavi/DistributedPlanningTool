<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeleteVehicle
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblVehicleID = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblXCCTEamName = New System.Windows.Forms.Label()
        Me.lblPhase = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblTemaName = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblTransmissionType = New System.Windows.Forms.Label()
        Me.lblEngine = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblTransmission = New System.Windows.Forms.Label()
        Me.lblEngineType = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(367, 17)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(100, 28)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, "Close form [Esc]")
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdDelete.Location = New System.Drawing.Point(253, 17)
        Me.cmdDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(100, 28)
        Me.cmdDelete.TabIndex = 0
        Me.cmdDelete.Text = "&Delete unit"
        Me.ToolTip1.SetToolTip(Me.cmdDelete, "Delete unit [F7]")
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdCancel)
        Me.GroupBox1.Controls.Add(Me.cmdDelete)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 334)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(477, 57)
        Me.GroupBox1.TabIndex = 21
        Me.GroupBox1.TabStop = False
        '
        'lblVehicleID
        '
        Me.lblVehicleID.AutoSize = True
        Me.lblVehicleID.ForeColor = System.Drawing.Color.Blue
        Me.lblVehicleID.Location = New System.Drawing.Point(176, 31)
        Me.lblVehicleID.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblVehicleID.Name = "lblVehicleID"
        Me.lblVehicleID.Size = New System.Drawing.Size(81, 17)
        Me.lblVehicleID.TabIndex = 11
        Me.lblVehicleID.Text = "lblVehicleID"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(27, 282)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(129, 17)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "XCC Team Name"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(61, 250)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 17)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Team Name"
        '
        'lblXCCTEamName
        '
        Me.lblXCCTEamName.AutoSize = True
        Me.lblXCCTEamName.ForeColor = System.Drawing.Color.Blue
        Me.lblXCCTEamName.Location = New System.Drawing.Point(176, 282)
        Me.lblXCCTEamName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblXCCTEamName.Name = "lblXCCTEamName"
        Me.lblXCCTEamName.Size = New System.Drawing.Size(123, 17)
        Me.lblXCCTEamName.TabIndex = 20
        Me.lblXCCTEamName.Text = "lblXCCTEamName"
        '
        'lblPhase
        '
        Me.lblPhase.AutoSize = True
        Me.lblPhase.ForeColor = System.Drawing.Color.Blue
        Me.lblPhase.Location = New System.Drawing.Point(176, 63)
        Me.lblPhase.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPhase.Name = "lblPhase"
        Me.lblPhase.Size = New System.Drawing.Size(62, 17)
        Me.lblPhase.TabIndex = 12
        Me.lblPhase.Text = "lblPhase"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(99, 31)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Unit ID"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(11, 218)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(145, 17)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Transmission Type"
        '
        'lblTemaName
        '
        Me.lblTemaName.AutoSize = True
        Me.lblTemaName.ForeColor = System.Drawing.Color.Blue
        Me.lblTemaName.Location = New System.Drawing.Point(176, 250)
        Me.lblTemaName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTemaName.Name = "lblTemaName"
        Me.lblTemaName.Size = New System.Drawing.Size(95, 17)
        Me.lblTemaName.TabIndex = 19
        Me.lblTemaName.Text = "lblTemaName"
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.ForeColor = System.Drawing.Color.Blue
        Me.lblType.Location = New System.Drawing.Point(176, 95)
        Me.lblType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(54, 17)
        Me.lblType.TabIndex = 13
        Me.lblType.Text = "lblType"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(103, 63)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Phase"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(52, 186)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 17)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Transmission"
        '
        'lblTransmissionType
        '
        Me.lblTransmissionType.AutoSize = True
        Me.lblTransmissionType.ForeColor = System.Drawing.Color.Blue
        Me.lblTransmissionType.Location = New System.Drawing.Point(176, 218)
        Me.lblTransmissionType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTransmissionType.Name = "lblTransmissionType"
        Me.lblTransmissionType.Size = New System.Drawing.Size(138, 17)
        Me.lblTransmissionType.TabIndex = 18
        Me.lblTransmissionType.Text = "lblTransmissionType"
        '
        'lblEngine
        '
        Me.lblEngine.AutoSize = True
        Me.lblEngine.ForeColor = System.Drawing.Color.Blue
        Me.lblEngine.Location = New System.Drawing.Point(176, 122)
        Me.lblEngine.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEngine.Name = "lblEngine"
        Me.lblEngine.Size = New System.Drawing.Size(66, 17)
        Me.lblEngine.TabIndex = 15
        Me.lblEngine.Text = "lblEngine"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(112, 95)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Type"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(57, 154)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(99, 17)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Engine Type"
        '
        'lblTransmission
        '
        Me.lblTransmission.AutoSize = True
        Me.lblTransmission.ForeColor = System.Drawing.Color.Blue
        Me.lblTransmission.Location = New System.Drawing.Point(176, 186)
        Me.lblTransmission.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTransmission.Name = "lblTransmission"
        Me.lblTransmission.Size = New System.Drawing.Size(106, 17)
        Me.lblTransmission.TabIndex = 17
        Me.lblTransmission.Text = "lblTransmission"
        '
        'lblEngineType
        '
        Me.lblEngineType.AutoSize = True
        Me.lblEngineType.ForeColor = System.Drawing.Color.Blue
        Me.lblEngineType.Location = New System.Drawing.Point(176, 154)
        Me.lblEngineType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEngineType.Name = "lblEngineType"
        Me.lblEngineType.Size = New System.Drawing.Size(98, 17)
        Me.lblEngineType.TabIndex = 16
        Me.lblEngineType.Text = "lblEngineType"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(97, 122)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Engine"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblXCCTEamName)
        Me.GroupBox2.Controls.Add(Me.lblVehicleID)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.lblType)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.lblTemaName)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.lblTransmissionType)
        Me.GroupBox2.Controls.Add(Me.lblEngineType)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.lblTransmission)
        Me.GroupBox2.Controls.Add(Me.lblEngine)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.lblPhase)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 15)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(477, 319)
        Me.GroupBox2.TabIndex = 22
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Unit infomation"
        '
        'frmDeleteVehicle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(495, 400)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDeleteVehicle"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Delete unit"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblXCCTEamName As System.Windows.Forms.Label
    Friend WithEvents lblTemaName As System.Windows.Forms.Label
    Friend WithEvents lblTransmissionType As System.Windows.Forms.Label
    Friend WithEvents lblTransmission As System.Windows.Forms.Label
    Friend WithEvents lblEngineType As System.Windows.Forms.Label
    Friend WithEvents lblEngine As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblPhase As System.Windows.Forms.Label
    Friend WithEvents lblVehicleID As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
End Class
