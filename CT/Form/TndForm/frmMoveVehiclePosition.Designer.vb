<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMoveVehiclePosition
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
        Me.grpMenu = New System.Windows.Forms.GroupBox()
        Me.txtMovePositionRank = New System.Windows.Forms.TextBox()
        Me.lblXCCTeamName = New System.Windows.Forms.Label()
        Me.lblEngineType = New System.Windows.Forms.Label()
        Me.lblMovePositionRankTitle = New System.Windows.Forms.Label()
        Me.lblPhase = New System.Windows.Forms.Label()
        Me.lblEngineTypeTitle = New System.Windows.Forms.Label()
        Me.lblTeamName = New System.Windows.Forms.Label()
        Me.lblEngine = New System.Windows.Forms.Label()
        Me.lblTeamNameTitle = New System.Windows.Forms.Label()
        Me.lblXCCTeamNameTitle = New System.Windows.Forms.Label()
        Me.lblTransmissionTitle = New System.Windows.Forms.Label()
        Me.lblTypeTitle = New System.Windows.Forms.Label()
        Me.lblTransmission = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblVehicleID = New System.Windows.Forms.Label()
        Me.lblTransmissionType = New System.Windows.Forms.Label()
        Me.lblTransmissionTypeTitle = New System.Windows.Forms.Label()
        Me.lblPhaseTitle = New System.Windows.Forms.Label()
        Me.lblEngineTitle = New System.Windows.Forms.Label()
        Me.lblVehicleIDTitle = New System.Windows.Forms.Label()
        Me.btnMove = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.grpMenu.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMenu
        '
        Me.grpMenu.Controls.Add(Me.txtMovePositionRank)
        Me.grpMenu.Controls.Add(Me.lblXCCTeamName)
        Me.grpMenu.Controls.Add(Me.lblEngineType)
        Me.grpMenu.Controls.Add(Me.lblMovePositionRankTitle)
        Me.grpMenu.Controls.Add(Me.lblPhase)
        Me.grpMenu.Controls.Add(Me.lblEngineTypeTitle)
        Me.grpMenu.Controls.Add(Me.lblTeamName)
        Me.grpMenu.Controls.Add(Me.lblEngine)
        Me.grpMenu.Controls.Add(Me.lblTeamNameTitle)
        Me.grpMenu.Controls.Add(Me.lblXCCTeamNameTitle)
        Me.grpMenu.Controls.Add(Me.lblTransmissionTitle)
        Me.grpMenu.Controls.Add(Me.lblTypeTitle)
        Me.grpMenu.Controls.Add(Me.lblTransmission)
        Me.grpMenu.Controls.Add(Me.lblType)
        Me.grpMenu.Controls.Add(Me.lblVehicleID)
        Me.grpMenu.Controls.Add(Me.lblTransmissionType)
        Me.grpMenu.Controls.Add(Me.lblTransmissionTypeTitle)
        Me.grpMenu.Controls.Add(Me.lblPhaseTitle)
        Me.grpMenu.Controls.Add(Me.lblEngineTitle)
        Me.grpMenu.Controls.Add(Me.lblVehicleIDTitle)
        Me.grpMenu.Location = New System.Drawing.Point(9, 12)
        Me.grpMenu.Name = "grpMenu"
        Me.grpMenu.Size = New System.Drawing.Size(358, 264)
        Me.grpMenu.TabIndex = 0
        Me.grpMenu.TabStop = False
        Me.grpMenu.Text = "Unit Information"
        '
        'txtMovePositionRank
        '
        Me.txtMovePositionRank.Location = New System.Drawing.Point(152, 233)
        Me.txtMovePositionRank.Name = "txtMovePositionRank"
        Me.txtMovePositionRank.Size = New System.Drawing.Size(53, 20)
        Me.txtMovePositionRank.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.txtMovePositionRank, "Rank [F4]")
        '
        'lblXCCTeamName
        '
        Me.lblXCCTeamName.AutoSize = True
        Me.lblXCCTeamName.ForeColor = System.Drawing.Color.Blue
        Me.lblXCCTeamName.Location = New System.Drawing.Point(153, 212)
        Me.lblXCCTeamName.Name = "lblXCCTeamName"
        Me.lblXCCTeamName.Size = New System.Drawing.Size(10, 13)
        Me.lblXCCTeamName.TabIndex = 22
        Me.lblXCCTeamName.Text = "-"
        '
        'lblEngineType
        '
        Me.lblEngineType.AutoSize = True
        Me.lblEngineType.ForeColor = System.Drawing.Color.Blue
        Me.lblEngineType.Location = New System.Drawing.Point(153, 116)
        Me.lblEngineType.Name = "lblEngineType"
        Me.lblEngineType.Size = New System.Drawing.Size(10, 13)
        Me.lblEngineType.TabIndex = 21
        Me.lblEngineType.Text = "-"
        '
        'lblMovePositionRankTitle
        '
        Me.lblMovePositionRankTitle.AutoSize = True
        Me.lblMovePositionRankTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblMovePositionRankTitle.Location = New System.Drawing.Point(10, 236)
        Me.lblMovePositionRankTitle.Name = "lblMovePositionRankTitle"
        Me.lblMovePositionRankTitle.Size = New System.Drawing.Size(130, 13)
        Me.lblMovePositionRankTitle.TabIndex = 20
        Me.lblMovePositionRankTitle.Text = "Move to position &rank"
        '
        'lblPhase
        '
        Me.lblPhase.AutoSize = True
        Me.lblPhase.ForeColor = System.Drawing.Color.Blue
        Me.lblPhase.Location = New System.Drawing.Point(153, 42)
        Me.lblPhase.Name = "lblPhase"
        Me.lblPhase.Size = New System.Drawing.Size(10, 13)
        Me.lblPhase.TabIndex = 19
        Me.lblPhase.Text = "-"
        '
        'lblEngineTypeTitle
        '
        Me.lblEngineTypeTitle.AutoSize = True
        Me.lblEngineTypeTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblEngineTypeTitle.Location = New System.Drawing.Point(60, 116)
        Me.lblEngineTypeTitle.Name = "lblEngineTypeTitle"
        Me.lblEngineTypeTitle.Size = New System.Drawing.Size(78, 13)
        Me.lblEngineTypeTitle.TabIndex = 18
        Me.lblEngineTypeTitle.Text = "Engine Type"
        '
        'lblTeamName
        '
        Me.lblTeamName.AutoSize = True
        Me.lblTeamName.ForeColor = System.Drawing.Color.Blue
        Me.lblTeamName.Location = New System.Drawing.Point(153, 188)
        Me.lblTeamName.Name = "lblTeamName"
        Me.lblTeamName.Size = New System.Drawing.Size(10, 13)
        Me.lblTeamName.TabIndex = 17
        Me.lblTeamName.Text = "-"
        '
        'lblEngine
        '
        Me.lblEngine.AutoSize = True
        Me.lblEngine.ForeColor = System.Drawing.Color.Blue
        Me.lblEngine.Location = New System.Drawing.Point(153, 92)
        Me.lblEngine.Name = "lblEngine"
        Me.lblEngine.Size = New System.Drawing.Size(10, 13)
        Me.lblEngine.TabIndex = 16
        Me.lblEngine.Text = "-"
        '
        'lblTeamNameTitle
        '
        Me.lblTeamNameTitle.AutoSize = True
        Me.lblTeamNameTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblTeamNameTitle.Location = New System.Drawing.Point(64, 188)
        Me.lblTeamNameTitle.Name = "lblTeamNameTitle"
        Me.lblTeamNameTitle.Size = New System.Drawing.Size(74, 13)
        Me.lblTeamNameTitle.TabIndex = 15
        Me.lblTeamNameTitle.Text = "Team Name"
        '
        'lblXCCTeamNameTitle
        '
        Me.lblXCCTeamNameTitle.AutoSize = True
        Me.lblXCCTeamNameTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblXCCTeamNameTitle.Location = New System.Drawing.Point(38, 212)
        Me.lblXCCTeamNameTitle.Name = "lblXCCTeamNameTitle"
        Me.lblXCCTeamNameTitle.Size = New System.Drawing.Size(102, 13)
        Me.lblXCCTeamNameTitle.TabIndex = 14
        Me.lblXCCTeamNameTitle.Text = "XCC Team Name"
        '
        'lblTransmissionTitle
        '
        Me.lblTransmissionTitle.AutoSize = True
        Me.lblTransmissionTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblTransmissionTitle.Location = New System.Drawing.Point(56, 140)
        Me.lblTransmissionTitle.Name = "lblTransmissionTitle"
        Me.lblTransmissionTitle.Size = New System.Drawing.Size(80, 13)
        Me.lblTransmissionTitle.TabIndex = 13
        Me.lblTransmissionTitle.Text = "Transmission"
        '
        'lblTypeTitle
        '
        Me.lblTypeTitle.AutoSize = True
        Me.lblTypeTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblTypeTitle.Location = New System.Drawing.Point(101, 66)
        Me.lblTypeTitle.Name = "lblTypeTitle"
        Me.lblTypeTitle.Size = New System.Drawing.Size(35, 13)
        Me.lblTypeTitle.TabIndex = 12
        Me.lblTypeTitle.Text = "Type"
        '
        'lblTransmission
        '
        Me.lblTransmission.AutoSize = True
        Me.lblTransmission.ForeColor = System.Drawing.Color.Blue
        Me.lblTransmission.Location = New System.Drawing.Point(153, 140)
        Me.lblTransmission.Name = "lblTransmission"
        Me.lblTransmission.Size = New System.Drawing.Size(10, 13)
        Me.lblTransmission.TabIndex = 10
        Me.lblTransmission.Text = "-"
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.ForeColor = System.Drawing.Color.Blue
        Me.lblType.Location = New System.Drawing.Point(153, 66)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(10, 13)
        Me.lblType.TabIndex = 9
        Me.lblType.Text = "-"
        '
        'lblVehicleID
        '
        Me.lblVehicleID.AutoSize = True
        Me.lblVehicleID.ForeColor = System.Drawing.Color.Blue
        Me.lblVehicleID.Location = New System.Drawing.Point(153, 18)
        Me.lblVehicleID.Name = "lblVehicleID"
        Me.lblVehicleID.Size = New System.Drawing.Size(10, 13)
        Me.lblVehicleID.TabIndex = 7
        Me.lblVehicleID.Text = "-"
        '
        'lblTransmissionType
        '
        Me.lblTransmissionType.AutoSize = True
        Me.lblTransmissionType.ForeColor = System.Drawing.Color.Blue
        Me.lblTransmissionType.Location = New System.Drawing.Point(153, 164)
        Me.lblTransmissionType.Name = "lblTransmissionType"
        Me.lblTransmissionType.Size = New System.Drawing.Size(10, 13)
        Me.lblTransmissionType.TabIndex = 5
        Me.lblTransmissionType.Text = "-"
        '
        'lblTransmissionTypeTitle
        '
        Me.lblTransmissionTypeTitle.AutoSize = True
        Me.lblTransmissionTypeTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblTransmissionTypeTitle.Location = New System.Drawing.Point(26, 164)
        Me.lblTransmissionTypeTitle.Name = "lblTransmissionTypeTitle"
        Me.lblTransmissionTypeTitle.Size = New System.Drawing.Size(112, 13)
        Me.lblTransmissionTypeTitle.TabIndex = 3
        Me.lblTransmissionTypeTitle.Text = "Transmission Type"
        '
        'lblPhaseTitle
        '
        Me.lblPhaseTitle.AutoSize = True
        Me.lblPhaseTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblPhaseTitle.Location = New System.Drawing.Point(94, 42)
        Me.lblPhaseTitle.Name = "lblPhaseTitle"
        Me.lblPhaseTitle.Size = New System.Drawing.Size(42, 13)
        Me.lblPhaseTitle.TabIndex = 2
        Me.lblPhaseTitle.Text = "Phase"
        '
        'lblEngineTitle
        '
        Me.lblEngineTitle.AutoSize = True
        Me.lblEngineTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblEngineTitle.Location = New System.Drawing.Point(91, 92)
        Me.lblEngineTitle.Name = "lblEngineTitle"
        Me.lblEngineTitle.Size = New System.Drawing.Size(46, 13)
        Me.lblEngineTitle.TabIndex = 1
        Me.lblEngineTitle.Text = "Engine"
        '
        'lblVehicleIDTitle
        '
        Me.lblVehicleIDTitle.AutoSize = True
        Me.lblVehicleIDTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblVehicleIDTitle.Location = New System.Drawing.Point(46, 18)
        Me.lblVehicleIDTitle.Name = "lblVehicleIDTitle"
        Me.lblVehicleIDTitle.Size = New System.Drawing.Size(90, 13)
        Me.lblVehicleIDTitle.TabIndex = 0
        Me.lblVehicleIDTitle.Text = "Current unit ID"
        '
        'btnMove
        '
        Me.btnMove.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnMove.Location = New System.Drawing.Point(194, 14)
        Me.btnMove.Name = "btnMove"
        Me.btnMove.Size = New System.Drawing.Size(75, 23)
        Me.btnMove.TabIndex = 1
        Me.btnMove.Text = "&Move unit"
        Me.ToolTip1.SetToolTip(Me.btnMove, "Move unit [F7]")
        Me.btnMove.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(274, 14)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Close the form [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnMove)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 281)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(358, 46)
        Me.GroupBox1.TabIndex = 22
        Me.GroupBox1.TabStop = False
        '
        'frmMoveVehiclePosition
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(374, 333)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpMenu)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMoveVehiclePosition"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Move unit position"
        Me.grpMenu.ResumeLayout(False)
        Me.grpMenu.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grpMenu As System.Windows.Forms.GroupBox
    Friend WithEvents txtMovePositionRank As System.Windows.Forms.TextBox
    Friend WithEvents lblXCCTeamName As System.Windows.Forms.Label
    Friend WithEvents lblEngineType As System.Windows.Forms.Label
    Friend WithEvents lblMovePositionRankTitle As System.Windows.Forms.Label
    Friend WithEvents lblPhase As System.Windows.Forms.Label
    Friend WithEvents lblEngineTypeTitle As System.Windows.Forms.Label
    Friend WithEvents lblTeamName As System.Windows.Forms.Label
    Friend WithEvents lblEngine As System.Windows.Forms.Label
    Friend WithEvents lblTeamNameTitle As System.Windows.Forms.Label
    Friend WithEvents lblXCCTeamNameTitle As System.Windows.Forms.Label
    Friend WithEvents lblTransmissionTitle As System.Windows.Forms.Label
    Friend WithEvents lblTypeTitle As System.Windows.Forms.Label
    Friend WithEvents lblTransmission As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblVehicleID As System.Windows.Forms.Label
    Friend WithEvents lblTransmissionType As System.Windows.Forms.Label
    Friend WithEvents lblTransmissionTypeTitle As System.Windows.Forms.Label
    Friend WithEvents lblPhaseTitle As System.Windows.Forms.Label
    Friend WithEvents lblEngineTitle As System.Windows.Forms.Label
    Friend WithEvents lblVehicleIDTitle As System.Windows.Forms.Label
    Friend WithEvents btnMove As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
