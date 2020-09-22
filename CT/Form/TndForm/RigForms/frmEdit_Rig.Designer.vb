<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmEdit_Rig
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEdit))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtRemarks = New System.Windows.Forms.TextBox()
        Me.cboCDSID = New System.Windows.Forms.ComboBox()
        Me.cboSubFacility = New System.Windows.Forms.ComboBox()
        Me.cboMatchedFacility = New System.Windows.Forms.ComboBox()
        Me.cboProcessStepLocation = New System.Windows.Forms.ComboBox()
        Me.cboLocation_CBG = New System.Windows.Forms.ComboBox()
        Me.lblProcessStep = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.lblGlobal = New System.Windows.Forms.Label()
        Me.lblUserCase = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkHolidays = New System.Windows.Forms.CheckBox()
        Me.LblWorkingdays = New System.Windows.Forms.Label()
        Me.dtEnd = New System.Windows.Forms.DateTimePicker()
        Me.dtStart = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.opt7Days = New System.Windows.Forms.RadioButton()
        Me.opt6Days = New System.Windows.Forms.RadioButton()
        Me.opt5Days = New System.Windows.Forms.RadioButton()
        Me.chkStartAndEnd = New System.Windows.Forms.CheckBox()
        Me.lblKWEnd = New System.Windows.Forms.Label()
        Me.lblKWSt = New System.Windows.Forms.Label()
        Me.optWorkingDays = New System.Windows.Forms.RadioButton()
        Me.optWeeks = New System.Windows.Forms.RadioButton()
        Me.txtDuration = New System.Windows.Forms.TextBox()
        Me.chkEnd = New System.Windows.Forms.CheckBox()
        Me.chkStart = New System.Windows.Forms.CheckBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.lblDuration = New System.Windows.Forms.Label()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtRemarks)
        Me.GroupBox1.Controls.Add(Me.cboCDSID)
        Me.GroupBox1.Controls.Add(Me.cboSubFacility)
        Me.GroupBox1.Controls.Add(Me.cboMatchedFacility)
        Me.GroupBox1.Controls.Add(Me.cboProcessStepLocation)
        Me.GroupBox1.Controls.Add(Me.cboLocation_CBG)
        Me.GroupBox1.Controls.Add(Me.lblProcessStep)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'txtRemarks
        '
        resources.ApplyResources(Me.txtRemarks, "txtRemarks")
        Me.txtRemarks.Name = "txtRemarks"
        '
        'cboCDSID
        '
        Me.cboCDSID.FormattingEnabled = True
        Me.cboCDSID.Items.AddRange(New Object() {resources.GetString("cboCDSID.Items")})
        resources.ApplyResources(Me.cboCDSID, "cboCDSID")
        Me.cboCDSID.Name = "cboCDSID"
        '
        'cboSubFacility
        '
        Me.cboSubFacility.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSubFacility.FormattingEnabled = True
        resources.ApplyResources(Me.cboSubFacility, "cboSubFacility")
        Me.cboSubFacility.Name = "cboSubFacility"
        '
        'cboMatchedFacility
        '
        Me.cboMatchedFacility.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMatchedFacility.FormattingEnabled = True
        resources.ApplyResources(Me.cboMatchedFacility, "cboMatchedFacility")
        Me.cboMatchedFacility.Name = "cboMatchedFacility"
        '
        'cboProcessStepLocation
        '
        Me.cboProcessStepLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProcessStepLocation.FormattingEnabled = True
        resources.ApplyResources(Me.cboProcessStepLocation, "cboProcessStepLocation")
        Me.cboProcessStepLocation.Name = "cboProcessStepLocation"
        '
        'cboLocation_CBG
        '
        Me.cboLocation_CBG.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation_CBG.FormattingEnabled = True
        resources.ApplyResources(Me.cboLocation_CBG, "cboLocation_CBG")
        Me.cboLocation_CBG.Name = "cboLocation_CBG"
        Me.ToolTip1.SetToolTip(Me.cboLocation_CBG, resources.GetString("cboLocation_CBG.ToolTip"))
        '
        'lblProcessStep
        '
        resources.ApplyResources(Me.lblProcessStep, "lblProcessStep")
        Me.lblProcessStep.ForeColor = System.Drawing.Color.Blue
        Me.lblProcessStep.Name = "lblProcessStep"
        '
        'Label14
        '
        resources.ApplyResources(Me.Label14, "Label14")
        Me.Label14.Name = "Label14"
        '
        'Label13
        '
        resources.ApplyResources(Me.Label13, "Label13")
        Me.Label13.Name = "Label13"
        '
        'Label12
        '
        resources.ApplyResources(Me.Label12, "Label12")
        Me.Label12.Name = "Label12"
        '
        'Label11
        '
        resources.ApplyResources(Me.Label11, "Label11")
        Me.Label11.Name = "Label11"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'lblUser
        '
        resources.ApplyResources(Me.lblUser, "lblUser")
        Me.lblUser.ForeColor = System.Drawing.Color.Blue
        Me.lblUser.Name = "lblUser"
        '
        'lblGlobal
        '
        resources.ApplyResources(Me.lblGlobal, "lblGlobal")
        Me.lblGlobal.ForeColor = System.Drawing.Color.Blue
        Me.lblGlobal.Name = "lblGlobal"
        '
        'lblUserCase
        '
        resources.ApplyResources(Me.lblUserCase, "lblUserCase")
        Me.lblUserCase.ForeColor = System.Drawing.Color.Blue
        Me.lblUserCase.Name = "lblUserCase"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.Name = "Label10"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkHolidays)
        Me.GroupBox2.Controls.Add(Me.LblWorkingdays)
        Me.GroupBox2.Controls.Add(Me.dtEnd)
        Me.GroupBox2.Controls.Add(Me.dtStart)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.chkStartAndEnd)
        Me.GroupBox2.Controls.Add(Me.lblKWEnd)
        Me.GroupBox2.Controls.Add(Me.lblKWSt)
        Me.GroupBox2.Controls.Add(Me.optWorkingDays)
        Me.GroupBox2.Controls.Add(Me.optWeeks)
        Me.GroupBox2.Controls.Add(Me.txtDuration)
        Me.GroupBox2.Controls.Add(Me.chkEnd)
        Me.GroupBox2.Controls.Add(Me.chkStart)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.lblEndDate)
        Me.GroupBox2.Controls.Add(Me.lblDuration)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'chkHolidays
        '
        resources.ApplyResources(Me.chkHolidays, "chkHolidays")
        Me.chkHolidays.Checked = True
        Me.chkHolidays.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHolidays.Name = "chkHolidays"
        Me.ToolTip1.SetToolTip(Me.chkHolidays, resources.GetString("chkHolidays.ToolTip"))
        Me.chkHolidays.UseVisualStyleBackColor = True
        '
        'LblWorkingdays
        '
        resources.ApplyResources(Me.LblWorkingdays, "LblWorkingdays")
        Me.LblWorkingdays.Name = "LblWorkingdays"
        '
        'dtEnd
        '
        Me.dtEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        resources.ApplyResources(Me.dtEnd, "dtEnd")
        Me.dtEnd.Name = "dtEnd"
        '
        'dtStart
        '
        Me.dtStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        resources.ApplyResources(Me.dtStart, "dtStart")
        Me.dtStart.Name = "dtStart"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.opt7Days)
        Me.GroupBox3.Controls.Add(Me.opt6Days)
        Me.GroupBox3.Controls.Add(Me.opt5Days)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'opt7Days
        '
        resources.ApplyResources(Me.opt7Days, "opt7Days")
        Me.opt7Days.Name = "opt7Days"
        Me.opt7Days.UseVisualStyleBackColor = True
        '
        'opt6Days
        '
        resources.ApplyResources(Me.opt6Days, "opt6Days")
        Me.opt6Days.Name = "opt6Days"
        Me.opt6Days.UseVisualStyleBackColor = True
        '
        'opt5Days
        '
        resources.ApplyResources(Me.opt5Days, "opt5Days")
        Me.opt5Days.Checked = True
        Me.opt5Days.Name = "opt5Days"
        Me.opt5Days.TabStop = True
        Me.opt5Days.UseVisualStyleBackColor = True
        '
        'chkStartAndEnd
        '
        resources.ApplyResources(Me.chkStartAndEnd, "chkStartAndEnd")
        Me.chkStartAndEnd.Name = "chkStartAndEnd"
        Me.ToolTip1.SetToolTip(Me.chkStartAndEnd, resources.GetString("chkStartAndEnd.ToolTip"))
        Me.chkStartAndEnd.UseVisualStyleBackColor = True
        '
        'lblKWEnd
        '
        resources.ApplyResources(Me.lblKWEnd, "lblKWEnd")
        Me.lblKWEnd.Name = "lblKWEnd"
        '
        'lblKWSt
        '
        resources.ApplyResources(Me.lblKWSt, "lblKWSt")
        Me.lblKWSt.Name = "lblKWSt"
        '
        'optWorkingDays
        '
        resources.ApplyResources(Me.optWorkingDays, "optWorkingDays")
        Me.optWorkingDays.Checked = True
        Me.optWorkingDays.Name = "optWorkingDays"
        Me.optWorkingDays.TabStop = True
        Me.optWorkingDays.UseVisualStyleBackColor = True
        '
        'optWeeks
        '
        resources.ApplyResources(Me.optWeeks, "optWeeks")
        Me.optWeeks.Name = "optWeeks"
        Me.optWeeks.UseVisualStyleBackColor = True
        '
        'txtDuration
        '
        resources.ApplyResources(Me.txtDuration, "txtDuration")
        Me.txtDuration.Name = "txtDuration"
        '
        'chkEnd
        '
        resources.ApplyResources(Me.chkEnd, "chkEnd")
        Me.chkEnd.Name = "chkEnd"
        Me.chkEnd.UseVisualStyleBackColor = True
        '
        'chkStart
        '
        resources.ApplyResources(Me.chkStart, "chkStart")
        Me.chkStart.Name = "chkStart"
        Me.chkStart.UseVisualStyleBackColor = True
        '
        'Label15
        '
        resources.ApplyResources(Me.Label15, "Label15")
        Me.Label15.Name = "Label15"
        '
        'lblEndDate
        '
        resources.ApplyResources(Me.lblEndDate, "lblEndDate")
        Me.lblEndDate.Name = "lblEndDate"
        '
        'lblDuration
        '
        resources.ApplyResources(Me.lblDuration, "lblDuration")
        Me.lblDuration.Name = "lblDuration"
        '
        'cmdOk
        '
        resources.ApplyResources(Me.cmdOk, "cmdOk")
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Name = "cmdOk"
        Me.ToolTip1.SetToolTip(Me.cmdOk, resources.GetString("cmdOk.ToolTip"))
        Me.cmdOk.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        resources.ApplyResources(Me.cmdCancel, "cmdCancel")
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Name = "cmdCancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, resources.GetString("cmdCancel.ToolTip"))
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cmdCancel)
        Me.GroupBox4.Controls.Add(Me.cmdOk)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.lblGlobal)
        Me.GroupBox5.Controls.Add(Me.Label10)
        Me.GroupBox5.Controls.Add(Me.lblUserCase)
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Controls.Add(Me.lblUser)
        resources.ApplyResources(Me.GroupBox5, "GroupBox5")
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.TabStop = False
        '
        'frmEdit
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEdit"
        Me.ShowInTaskbar = False
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents cboCDSID As System.Windows.Forms.ComboBox
    Friend WithEvents cboSubFacility As System.Windows.Forms.ComboBox
    Friend WithEvents cboMatchedFacility As System.Windows.Forms.ComboBox
    Friend WithEvents cboProcessStepLocation As System.Windows.Forms.ComboBox
    Friend WithEvents cboLocation_CBG As System.Windows.Forms.ComboBox
    Friend WithEvents lblProcessStep As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents lblGlobal As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lblUserCase As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents optWorkingDays As System.Windows.Forms.RadioButton
    Friend WithEvents optWeeks As System.Windows.Forms.RadioButton
    Friend WithEvents txtDuration As System.Windows.Forms.TextBox
    Friend WithEvents chkEnd As System.Windows.Forms.CheckBox
    Friend WithEvents chkStart As System.Windows.Forms.CheckBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents opt7Days As System.Windows.Forms.RadioButton
    Friend WithEvents opt6Days As System.Windows.Forms.RadioButton
    Friend WithEvents opt5Days As System.Windows.Forms.RadioButton
    Friend WithEvents chkStartAndEnd As System.Windows.Forms.CheckBox
    Friend WithEvents lblKWEnd As System.Windows.Forms.Label
    Friend WithEvents lblKWSt As System.Windows.Forms.Label
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents dtStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents LblWorkingdays As System.Windows.Forms.Label
    Friend WithEvents chkHolidays As System.Windows.Forms.CheckBox
End Class
