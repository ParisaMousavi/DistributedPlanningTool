<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAddDates_Rig
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
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grpM1 = New System.Windows.Forms.GroupBox()
        Me.btnClearM1dc = New System.Windows.Forms.Button()
        Me.lblM1dc = New System.Windows.Forms.Label()
        Me.btnClearM1 = New System.Windows.Forms.Button()
        Me.dtM1dc = New System.Windows.Forms.DateTimePicker()
        Me.lblM1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dtM1 = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtHCid = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.grpVP = New System.Windows.Forms.GroupBox()
        Me.btnClearFEC = New System.Windows.Forms.Button()
        Me.btnClearPEC = New System.Windows.Forms.Button()
        Me.btnClearVP = New System.Windows.Forms.Button()
        Me.lblFec = New System.Windows.Forms.Label()
        Me.dtFec = New System.Windows.Forms.DateTimePicker()
        Me.lblPec = New System.Windows.Forms.Label()
        Me.dtPec = New System.Windows.Forms.DateTimePicker()
        Me.lblVp = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.dtVP = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.cmdBackcolor = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdFontcolor = New System.Windows.Forms.Button()
        Me.dgvDates = New System.Windows.Forms.DataGridView()
        Me.lblID = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnClearMRD = New System.Windows.Forms.Button()
        Me.lblMRD = New System.Windows.Forms.Label()
        Me.dtMRD = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdReset = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.grpM1.SuspendLayout()
        Me.grpVP.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(19, 18)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 17)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Font Color"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 20)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "HCID"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 59)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "MRD"
        '
        'grpM1
        '
        Me.grpM1.Controls.Add(Me.btnClearM1dc)
        Me.grpM1.Controls.Add(Me.lblM1dc)
        Me.grpM1.Controls.Add(Me.btnClearM1)
        Me.grpM1.Controls.Add(Me.dtM1dc)
        Me.grpM1.Controls.Add(Me.lblM1)
        Me.grpM1.Controls.Add(Me.Label5)
        Me.grpM1.Controls.Add(Me.dtM1)
        Me.grpM1.Controls.Add(Me.Label4)
        Me.grpM1.Location = New System.Drawing.Point(48, 144)
        Me.grpM1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpM1.Name = "grpM1"
        Me.grpM1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpM1.Size = New System.Drawing.Size(329, 137)
        Me.grpM1.TabIndex = 1
        Me.grpM1.TabStop = False
        Me.grpM1.Text = "M1 / X0 / XM / X1 / TPV"
        '
        'btnClearM1dc
        '
        Me.btnClearM1dc.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearM1dc.Location = New System.Drawing.Point(287, 70)
        Me.btnClearM1dc.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearM1dc.Name = "btnClearM1dc"
        Me.btnClearM1dc.Size = New System.Drawing.Size(31, 23)
        Me.btnClearM1dc.TabIndex = 32
        Me.btnClearM1dc.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearM1dc, "Clear date value")
        Me.btnClearM1dc.UseVisualStyleBackColor = True
        '
        'lblM1dc
        '
        Me.lblM1dc.AutoSize = True
        Me.lblM1dc.Location = New System.Drawing.Point(251, 74)
        Me.lblM1dc.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblM1dc.Name = "lblM1dc"
        Me.lblM1dc.Size = New System.Drawing.Size(22, 17)
        Me.lblM1dc.TabIndex = 28
        Me.lblM1dc.Text = "( )"
        '
        'btnClearM1
        '
        Me.btnClearM1.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearM1.Location = New System.Drawing.Point(287, 30)
        Me.btnClearM1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearM1.Name = "btnClearM1"
        Me.btnClearM1.Size = New System.Drawing.Size(31, 23)
        Me.btnClearM1.TabIndex = 31
        Me.btnClearM1.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearM1, "Clear date value")
        Me.btnClearM1.UseVisualStyleBackColor = True
        '
        'dtM1dc
        '
        Me.dtM1dc.CustomFormat = " "
        Me.dtM1dc.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtM1dc.Location = New System.Drawing.Point(71, 69)
        Me.dtM1dc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtM1dc.Name = "dtM1dc"
        Me.dtM1dc.Size = New System.Drawing.Size(171, 22)
        Me.dtM1dc.TabIndex = 1
        '
        'lblM1
        '
        Me.lblM1.AutoSize = True
        Me.lblM1.Location = New System.Drawing.Point(251, 33)
        Me.lblM1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblM1.Name = "lblM1"
        Me.lblM1.Size = New System.Drawing.Size(22, 17)
        Me.lblM1.TabIndex = 26
        Me.lblM1.Text = "( )"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 74)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 17)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "M1DC"
        '
        'dtM1
        '
        Me.dtM1.CustomFormat = " "
        Me.dtM1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtM1.Location = New System.Drawing.Point(71, 28)
        Me.dtM1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtM1.Name = "dtM1"
        Me.dtM1.Size = New System.Drawing.Size(171, 22)
        Me.dtM1.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 33)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 17)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "1st M1"
        '
        'txtHCid
        '
        Me.txtHCid.Location = New System.Drawing.Point(71, 20)
        Me.txtHCid.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtHCid.Name = "txtHCid"
        Me.txtHCid.Size = New System.Drawing.Size(171, 22)
        Me.txtHCid.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtHCid, "HC Id [F4]")
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(19, 50)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(76, 17)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Back Color"
        '
        'grpVP
        '
        Me.grpVP.Controls.Add(Me.btnClearFEC)
        Me.grpVP.Controls.Add(Me.btnClearPEC)
        Me.grpVP.Controls.Add(Me.btnClearVP)
        Me.grpVP.Controls.Add(Me.lblFec)
        Me.grpVP.Controls.Add(Me.dtFec)
        Me.grpVP.Controls.Add(Me.lblPec)
        Me.grpVP.Controls.Add(Me.dtPec)
        Me.grpVP.Controls.Add(Me.lblVp)
        Me.grpVP.Controls.Add(Me.Label10)
        Me.grpVP.Controls.Add(Me.dtVP)
        Me.grpVP.Controls.Add(Me.Label8)
        Me.grpVP.Controls.Add(Me.Label9)
        Me.grpVP.Location = New System.Drawing.Point(385, 144)
        Me.grpVP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpVP.Name = "grpVP"
        Me.grpVP.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpVP.Size = New System.Drawing.Size(339, 137)
        Me.grpVP.TabIndex = 3
        Me.grpVP.TabStop = False
        Me.grpVP.Text = "VP / TT / PP / DCV"
        '
        'btnClearFEC
        '
        Me.btnClearFEC.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearFEC.Location = New System.Drawing.Point(299, 107)
        Me.btnClearFEC.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearFEC.Name = "btnClearFEC"
        Me.btnClearFEC.Size = New System.Drawing.Size(31, 23)
        Me.btnClearFEC.TabIndex = 35
        Me.btnClearFEC.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearFEC, "Clear date value")
        Me.btnClearFEC.UseVisualStyleBackColor = True
        '
        'btnClearPEC
        '
        Me.btnClearPEC.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearPEC.Location = New System.Drawing.Point(299, 70)
        Me.btnClearPEC.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearPEC.Name = "btnClearPEC"
        Me.btnClearPEC.Size = New System.Drawing.Size(31, 23)
        Me.btnClearPEC.TabIndex = 32
        Me.btnClearPEC.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearPEC, "Clear date value")
        Me.btnClearPEC.UseVisualStyleBackColor = True
        '
        'btnClearVP
        '
        Me.btnClearVP.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearVP.Location = New System.Drawing.Point(299, 30)
        Me.btnClearVP.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearVP.Name = "btnClearVP"
        Me.btnClearVP.Size = New System.Drawing.Size(31, 23)
        Me.btnClearVP.TabIndex = 31
        Me.btnClearVP.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearVP, "Clear date value")
        Me.btnClearVP.UseVisualStyleBackColor = True
        '
        'lblFec
        '
        Me.lblFec.AutoSize = True
        Me.lblFec.Location = New System.Drawing.Point(269, 111)
        Me.lblFec.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFec.Name = "lblFec"
        Me.lblFec.Size = New System.Drawing.Size(22, 17)
        Me.lblFec.TabIndex = 34
        Me.lblFec.Text = "( )"
        '
        'dtFec
        '
        Me.dtFec.CustomFormat = " "
        Me.dtFec.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtFec.Location = New System.Drawing.Point(91, 106)
        Me.dtFec.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtFec.Name = "dtFec"
        Me.dtFec.Size = New System.Drawing.Size(171, 22)
        Me.dtFec.TabIndex = 2
        '
        'lblPec
        '
        Me.lblPec.AutoSize = True
        Me.lblPec.Location = New System.Drawing.Point(269, 74)
        Me.lblPec.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPec.Name = "lblPec"
        Me.lblPec.Size = New System.Drawing.Size(22, 17)
        Me.lblPec.TabIndex = 32
        Me.lblPec.Text = "( )"
        '
        'dtPec
        '
        Me.dtPec.CustomFormat = " "
        Me.dtPec.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPec.Location = New System.Drawing.Point(91, 69)
        Me.dtPec.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtPec.Name = "dtPec"
        Me.dtPec.Size = New System.Drawing.Size(171, 22)
        Me.dtPec.TabIndex = 1
        '
        'lblVp
        '
        Me.lblVp.AutoSize = True
        Me.lblVp.Location = New System.Drawing.Point(269, 33)
        Me.lblVp.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblVp.Name = "lblVp"
        Me.lblVp.Size = New System.Drawing.Size(22, 17)
        Me.lblVp.TabIndex = 30
        Me.lblVp.Text = "( )"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 111)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(34, 17)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "FEC"
        '
        'dtVP
        '
        Me.dtVP.CustomFormat = " "
        Me.dtVP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtVP.Location = New System.Drawing.Point(91, 28)
        Me.dtVP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtVP.Name = "dtVP"
        Me.dtVP.Size = New System.Drawing.Size(171, 22)
        Me.dtVP.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 74)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(35, 17)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "PEC"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 33)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(49, 17)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "1st VP"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.GroupBox5)
        Me.GroupBox3.Controls.Add(Me.dgvDates)
        Me.GroupBox3.Controls.Add(Me.lblID)
        Me.GroupBox3.Controls.Add(Me.GroupBox4)
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.cmdClose)
        Me.GroupBox3.Controls.Add(Me.cmdReset)
        Me.GroupBox3.Controls.Add(Me.cmdDelete)
        Me.GroupBox3.Controls.Add(Me.cmdUpdate)
        Me.GroupBox3.Controls.Add(Me.cmdAdd)
        Me.GroupBox3.Controls.Add(Me.grpVP)
        Me.GroupBox3.Controls.Add(Me.grpM1)
        Me.GroupBox3.Location = New System.Drawing.Point(5, -2)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Size = New System.Drawing.Size(771, 634)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.cmdBackcolor)
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Controls.Add(Me.Label12)
        Me.GroupBox5.Controls.Add(Me.cmdFontcolor)
        Me.GroupBox5.Location = New System.Drawing.Point(385, 12)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox5.Size = New System.Drawing.Size(339, 98)
        Me.GroupBox5.TabIndex = 29
        Me.GroupBox5.TabStop = False
        '
        'cmdBackcolor
        '
        Me.cmdBackcolor.Location = New System.Drawing.Point(99, 50)
        Me.cmdBackcolor.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdBackcolor.Name = "cmdBackcolor"
        Me.cmdBackcolor.Size = New System.Drawing.Size(72, 28)
        Me.cmdBackcolor.TabIndex = 2
        Me.cmdBackcolor.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(191, 18)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(91, 57)
        Me.Label12.TabIndex = 19
        Me.Label12.Text = "(Click on the boxes to select color)"
        '
        'cmdFontcolor
        '
        Me.cmdFontcolor.Location = New System.Drawing.Point(99, 18)
        Me.cmdFontcolor.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdFontcolor.Name = "cmdFontcolor"
        Me.cmdFontcolor.Size = New System.Drawing.Size(72, 28)
        Me.cmdFontcolor.TabIndex = 1
        Me.cmdFontcolor.UseVisualStyleBackColor = True
        '
        'dgvDates
        '
        Me.dgvDates.AllowUserToAddRows = False
        Me.dgvDates.AllowUserToDeleteRows = False
        Me.dgvDates.AllowUserToOrderColumns = True
        Me.dgvDates.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvDates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDates.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvDates.Location = New System.Drawing.Point(9, 352)
        Me.dgvDates.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvDates.MultiSelect = False
        Me.dgvDates.Name = "dgvDates"
        Me.dgvDates.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDates.Size = New System.Drawing.Size(752, 267)
        Me.dgvDates.TabIndex = 9
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Location = New System.Drawing.Point(635, 12)
        Me.lblID.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(0, 17)
        Me.lblID.TabIndex = 28
        Me.lblID.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnClearMRD)
        Me.GroupBox4.Controls.Add(Me.lblMRD)
        Me.GroupBox4.Controls.Add(Me.dtMRD)
        Me.GroupBox4.Controls.Add(Me.txtHCid)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Location = New System.Drawing.Point(48, 12)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(329, 98)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'btnClearMRD
        '
        Me.btnClearMRD.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearMRD.Location = New System.Drawing.Point(287, 59)
        Me.btnClearMRD.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearMRD.Name = "btnClearMRD"
        Me.btnClearMRD.Size = New System.Drawing.Size(31, 23)
        Me.btnClearMRD.TabIndex = 29
        Me.btnClearMRD.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearMRD, "Clear date value")
        Me.btnClearMRD.UseVisualStyleBackColor = True
        '
        'lblMRD
        '
        Me.lblMRD.AutoSize = True
        Me.lblMRD.Location = New System.Drawing.Point(251, 59)
        Me.lblMRD.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMRD.Name = "lblMRD"
        Me.lblMRD.Size = New System.Drawing.Size(22, 17)
        Me.lblMRD.TabIndex = 22
        Me.lblMRD.Text = "( )"
        '
        'dtMRD
        '
        Me.dtMRD.CustomFormat = " "
        Me.dtMRD.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtMRD.Location = New System.Drawing.Point(71, 59)
        Me.dtMRD.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtMRD.Name = "dtMRD"
        Me.dtMRD.Size = New System.Drawing.Size(171, 22)
        Me.dtMRD.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(8, 330)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(264, 17)
        Me.Label11.TabIndex = 18
        Me.Label11.Text = "Click row to select data for update/delete"
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(584, 292)
        Me.cmdClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(107, 28)
        Me.cmdClose.TabIndex = 8
        Me.cmdClose.Text = "&Close"
        Me.ToolTip1.SetToolTip(Me.cmdClose, "Close form [Esc]")
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdReset
        '
        Me.cmdReset.Location = New System.Drawing.Point(456, 292)
        Me.cmdReset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(107, 28)
        Me.cmdReset.TabIndex = 7
        Me.cmdReset.Text = "&Reset"
        Me.ToolTip1.SetToolTip(Me.cmdReset, "Refresh [F5]")
        Me.cmdReset.UseVisualStyleBackColor = True
        '
        'cmdDelete
        '
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Location = New System.Drawing.Point(329, 292)
        Me.cmdDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(107, 28)
        Me.cmdDelete.TabIndex = 6
        Me.cmdDelete.Text = "&Delete"
        Me.ToolTip1.SetToolTip(Me.cmdDelete, "Delete [F9]")
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Enabled = False
        Me.cmdUpdate.Location = New System.Drawing.Point(203, 292)
        Me.cmdUpdate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(107, 28)
        Me.cmdUpdate.TabIndex = 5
        Me.cmdUpdate.Text = "&Update"
        Me.ToolTip1.SetToolTip(Me.cmdUpdate, "Update [F8]")
        Me.cmdUpdate.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.Enabled = False
        Me.cmdAdd.Location = New System.Drawing.Point(75, 292)
        Me.cmdAdd.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(107, 28)
        Me.cmdAdd.TabIndex = 4
        Me.cmdAdd.Text = "&Add"
        Me.ToolTip1.SetToolTip(Me.cmdAdd, "Add [F7]")
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'frmAddDates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 639)
        Me.Controls.Add(Me.GroupBox3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddDates"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "TnD Plan dates"
        Me.grpM1.ResumeLayout(False)
        Me.grpM1.PerformLayout()
        Me.grpVP.ResumeLayout(False)
        Me.grpVP.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents grpM1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtHCid As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents grpVP As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdReset As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents dtMRD As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblMRD As System.Windows.Forms.Label
    Friend WithEvents lblM1dc As System.Windows.Forms.Label
    Friend WithEvents dtM1dc As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblM1 As System.Windows.Forms.Label
    Friend WithEvents dtM1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFec As System.Windows.Forms.Label
    Friend WithEvents dtFec As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPec As System.Windows.Forms.Label
    Friend WithEvents dtPec As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblVp As System.Windows.Forms.Label
    Friend WithEvents dtVP As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdFontcolor As System.Windows.Forms.Button
    Friend WithEvents cmdBackcolor As System.Windows.Forms.Button
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents dgvDates As System.Windows.Forms.DataGridView
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
    Friend WithEvents btnClearMRD As System.Windows.Forms.Button
    Friend WithEvents btnClearM1dc As System.Windows.Forms.Button
    Friend WithEvents btnClearM1 As System.Windows.Forms.Button
    Friend WithEvents btnClearFEC As System.Windows.Forms.Button
    Friend WithEvents btnClearPEC As System.Windows.Forms.Button
    Friend WithEvents btnClearVP As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
End Class
