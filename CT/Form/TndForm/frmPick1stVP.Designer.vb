<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPick1stVP
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
        Me.pnAll = New System.Windows.Forms.Panel()
        Me.txtTnDPlanner = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grpVP = New System.Windows.Forms.GroupBox()
        Me.pnFEC = New System.Windows.Forms.Panel()
        Me.btnClearFEC = New System.Windows.Forms.Button()
        Me.dtFEC = New System.Windows.Forms.DateTimePicker()
        Me.lblFECTitle = New System.Windows.Forms.Label()
        Me.lblFEC = New System.Windows.Forms.Label()
        Me.pnPEC = New System.Windows.Forms.Panel()
        Me.btnClearPEC = New System.Windows.Forms.Button()
        Me.dtPEC = New System.Windows.Forms.DateTimePicker()
        Me.lblPECTitle = New System.Windows.Forms.Label()
        Me.lblPEC = New System.Windows.Forms.Label()
        Me.pn1stVP = New System.Windows.Forms.Panel()
        Me.btnClearVP = New System.Windows.Forms.Button()
        Me.dtVP = New System.Windows.Forms.DateTimePicker()
        Me.lbl1stVPTitle = New System.Windows.Forms.Label()
        Me.lblVP = New System.Windows.Forms.Label()
        Me.pnMRD = New System.Windows.Forms.Panel()
        Me.btnClearMRD = New System.Windows.Forms.Button()
        Me.dtMRD = New System.Windows.Forms.DateTimePicker()
        Me.lblMRDTitle = New System.Windows.Forms.Label()
        Me.lblMRD = New System.Windows.Forms.Label()
        Me.grpM1 = New System.Windows.Forms.GroupBox()
        Me.pnM1DC = New System.Windows.Forms.Panel()
        Me.btnClearM1DC = New System.Windows.Forms.Button()
        Me.dtM1DC = New System.Windows.Forms.DateTimePicker()
        Me.lblM1DCTitle = New System.Windows.Forms.Label()
        Me.lblM1DC = New System.Windows.Forms.Label()
        Me.pn1stM1 = New System.Windows.Forms.Panel()
        Me.btnClearM1 = New System.Windows.Forms.Button()
        Me.dtM1 = New System.Windows.Forms.DateTimePicker()
        Me.lbl1stM1Title = New System.Windows.Forms.Label()
        Me.lblM1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.brnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkInduFormatting = New System.Windows.Forms.CheckBox()
        Me.pnAll.SuspendLayout()
        Me.grpVP.SuspendLayout()
        Me.pnFEC.SuspendLayout()
        Me.pnPEC.SuspendLayout()
        Me.pn1stVP.SuspendLayout()
        Me.pnMRD.SuspendLayout()
        Me.grpM1.SuspendLayout()
        Me.pnM1DC.SuspendLayout()
        Me.pn1stM1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnAll
        '
        Me.pnAll.Controls.Add(Me.txtTnDPlanner)
        Me.pnAll.Controls.Add(Me.Label1)
        Me.pnAll.Controls.Add(Me.grpVP)
        Me.pnAll.Controls.Add(Me.pnMRD)
        Me.pnAll.Controls.Add(Me.grpM1)
        Me.pnAll.Location = New System.Drawing.Point(16, 15)
        Me.pnAll.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnAll.Name = "pnAll"
        Me.pnAll.Size = New System.Drawing.Size(309, 325)
        Me.pnAll.TabIndex = 0
        '
        'txtTnDPlanner
        '
        Me.txtTnDPlanner.Location = New System.Drawing.Point(111, 290)
        Me.txtTnDPlanner.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtTnDPlanner.MaxLength = 50
        Me.txtTnDPlanner.Name = "txtTnDPlanner"
        Me.txtTnDPlanner.Size = New System.Drawing.Size(161, 22)
        Me.txtTnDPlanner.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 294)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "TnD Planner"
        '
        'grpVP
        '
        Me.grpVP.Controls.Add(Me.pnFEC)
        Me.grpVP.Controls.Add(Me.pnPEC)
        Me.grpVP.Controls.Add(Me.pn1stVP)
        Me.grpVP.Location = New System.Drawing.Point(4, 149)
        Me.grpVP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpVP.Name = "grpVP"
        Me.grpVP.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpVP.Size = New System.Drawing.Size(301, 135)
        Me.grpVP.TabIndex = 2
        Me.grpVP.TabStop = False
        Me.grpVP.Text = "VP / DCV"
        '
        'pnFEC
        '
        Me.pnFEC.Controls.Add(Me.btnClearFEC)
        Me.pnFEC.Controls.Add(Me.dtFEC)
        Me.pnFEC.Controls.Add(Me.lblFECTitle)
        Me.pnFEC.Controls.Add(Me.lblFEC)
        Me.pnFEC.Location = New System.Drawing.Point(8, 96)
        Me.pnFEC.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnFEC.Name = "pnFEC"
        Me.pnFEC.Size = New System.Drawing.Size(283, 30)
        Me.pnFEC.TabIndex = 2
        '
        'btnClearFEC
        '
        Me.btnClearFEC.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearFEC.Location = New System.Drawing.Point(229, 5)
        Me.btnClearFEC.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearFEC.Name = "btnClearFEC"
        Me.btnClearFEC.Size = New System.Drawing.Size(31, 23)
        Me.btnClearFEC.TabIndex = 1
        Me.btnClearFEC.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearFEC, "Clear date value")
        Me.btnClearFEC.UseVisualStyleBackColor = True
        '
        'dtFEC
        '
        Me.dtFEC.CustomFormat = "dd MMM yyyy"
        Me.dtFEC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtFEC.Location = New System.Drawing.Point(55, 5)
        Me.dtFEC.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtFEC.Name = "dtFEC"
        Me.dtFEC.Size = New System.Drawing.Size(132, 22)
        Me.dtFEC.TabIndex = 0
        '
        'lblFECTitle
        '
        Me.lblFECTitle.AutoSize = True
        Me.lblFECTitle.Location = New System.Drawing.Point(4, 5)
        Me.lblFECTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFECTitle.Name = "lblFECTitle"
        Me.lblFECTitle.Size = New System.Drawing.Size(34, 17)
        Me.lblFECTitle.TabIndex = 9
        Me.lblFECTitle.Text = "FEC"
        '
        'lblFEC
        '
        Me.lblFEC.AutoSize = True
        Me.lblFEC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFEC.ForeColor = System.Drawing.Color.Black
        Me.lblFEC.Location = New System.Drawing.Point(196, 5)
        Me.lblFEC.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFEC.Name = "lblFEC"
        Me.lblFEC.Size = New System.Drawing.Size(18, 18)
        Me.lblFEC.TabIndex = 7
        Me.lblFEC.Text = "()"
        '
        'pnPEC
        '
        Me.pnPEC.Controls.Add(Me.btnClearPEC)
        Me.pnPEC.Controls.Add(Me.dtPEC)
        Me.pnPEC.Controls.Add(Me.lblPECTitle)
        Me.pnPEC.Controls.Add(Me.lblPEC)
        Me.pnPEC.Location = New System.Drawing.Point(8, 59)
        Me.pnPEC.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnPEC.Name = "pnPEC"
        Me.pnPEC.Size = New System.Drawing.Size(283, 30)
        Me.pnPEC.TabIndex = 1
        '
        'btnClearPEC
        '
        Me.btnClearPEC.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearPEC.Location = New System.Drawing.Point(229, 5)
        Me.btnClearPEC.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearPEC.Name = "btnClearPEC"
        Me.btnClearPEC.Size = New System.Drawing.Size(31, 23)
        Me.btnClearPEC.TabIndex = 1
        Me.btnClearPEC.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearPEC, "Clear date value")
        Me.btnClearPEC.UseVisualStyleBackColor = True
        '
        'dtPEC
        '
        Me.dtPEC.CustomFormat = "dd MMM yyyy"
        Me.dtPEC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPEC.Location = New System.Drawing.Point(55, 5)
        Me.dtPEC.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtPEC.Name = "dtPEC"
        Me.dtPEC.Size = New System.Drawing.Size(132, 22)
        Me.dtPEC.TabIndex = 0
        '
        'lblPECTitle
        '
        Me.lblPECTitle.AutoSize = True
        Me.lblPECTitle.Location = New System.Drawing.Point(4, 5)
        Me.lblPECTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPECTitle.Name = "lblPECTitle"
        Me.lblPECTitle.Size = New System.Drawing.Size(35, 17)
        Me.lblPECTitle.TabIndex = 9
        Me.lblPECTitle.Text = "PEC"
        '
        'lblPEC
        '
        Me.lblPEC.AutoSize = True
        Me.lblPEC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPEC.ForeColor = System.Drawing.Color.Black
        Me.lblPEC.Location = New System.Drawing.Point(196, 5)
        Me.lblPEC.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPEC.Name = "lblPEC"
        Me.lblPEC.Size = New System.Drawing.Size(18, 18)
        Me.lblPEC.TabIndex = 7
        Me.lblPEC.Text = "()"
        '
        'pn1stVP
        '
        Me.pn1stVP.Controls.Add(Me.btnClearVP)
        Me.pn1stVP.Controls.Add(Me.dtVP)
        Me.pn1stVP.Controls.Add(Me.lbl1stVPTitle)
        Me.pn1stVP.Controls.Add(Me.lblVP)
        Me.pn1stVP.Location = New System.Drawing.Point(8, 23)
        Me.pn1stVP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pn1stVP.Name = "pn1stVP"
        Me.pn1stVP.Size = New System.Drawing.Size(283, 30)
        Me.pn1stVP.TabIndex = 0
        '
        'btnClearVP
        '
        Me.btnClearVP.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearVP.Location = New System.Drawing.Point(229, 5)
        Me.btnClearVP.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearVP.Name = "btnClearVP"
        Me.btnClearVP.Size = New System.Drawing.Size(31, 23)
        Me.btnClearVP.TabIndex = 2
        Me.btnClearVP.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearVP, "Clear date value")
        Me.btnClearVP.UseVisualStyleBackColor = True
        '
        'dtVP
        '
        Me.dtVP.CustomFormat = "dd MMM yyyy"
        Me.dtVP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtVP.Location = New System.Drawing.Point(55, 5)
        Me.dtVP.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtVP.Name = "dtVP"
        Me.dtVP.Size = New System.Drawing.Size(132, 22)
        Me.dtVP.TabIndex = 1
        '
        'lbl1stVPTitle
        '
        Me.lbl1stVPTitle.AutoSize = True
        Me.lbl1stVPTitle.Location = New System.Drawing.Point(4, 5)
        Me.lbl1stVPTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl1stVPTitle.Name = "lbl1stVPTitle"
        Me.lbl1stVPTitle.Size = New System.Drawing.Size(49, 17)
        Me.lbl1stVPTitle.TabIndex = 0
        Me.lbl1stVPTitle.Text = "1st VP"
        '
        'lblVP
        '
        Me.lblVP.AutoSize = True
        Me.lblVP.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVP.ForeColor = System.Drawing.Color.Black
        Me.lblVP.Location = New System.Drawing.Point(196, 5)
        Me.lblVP.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblVP.Name = "lblVP"
        Me.lblVP.Size = New System.Drawing.Size(18, 18)
        Me.lblVP.TabIndex = 7
        Me.lblVP.Text = "()"
        '
        'pnMRD
        '
        Me.pnMRD.Controls.Add(Me.btnClearMRD)
        Me.pnMRD.Controls.Add(Me.dtMRD)
        Me.pnMRD.Controls.Add(Me.lblMRDTitle)
        Me.pnMRD.Controls.Add(Me.lblMRD)
        Me.pnMRD.Location = New System.Drawing.Point(12, 4)
        Me.pnMRD.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnMRD.Name = "pnMRD"
        Me.pnMRD.Size = New System.Drawing.Size(283, 30)
        Me.pnMRD.TabIndex = 0
        '
        'btnClearMRD
        '
        Me.btnClearMRD.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearMRD.Location = New System.Drawing.Point(229, 4)
        Me.btnClearMRD.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearMRD.Name = "btnClearMRD"
        Me.btnClearMRD.Size = New System.Drawing.Size(31, 23)
        Me.btnClearMRD.TabIndex = 2
        Me.btnClearMRD.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearMRD, "Clear date value")
        Me.btnClearMRD.UseVisualStyleBackColor = True
        '
        'dtMRD
        '
        Me.dtMRD.CustomFormat = "dd MMM yyyy"
        Me.dtMRD.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtMRD.Location = New System.Drawing.Point(55, 4)
        Me.dtMRD.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtMRD.Name = "dtMRD"
        Me.dtMRD.Size = New System.Drawing.Size(132, 22)
        Me.dtMRD.TabIndex = 1
        '
        'lblMRDTitle
        '
        Me.lblMRDTitle.AutoSize = True
        Me.lblMRDTitle.Location = New System.Drawing.Point(4, 4)
        Me.lblMRDTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMRDTitle.Name = "lblMRDTitle"
        Me.lblMRDTitle.Size = New System.Drawing.Size(39, 17)
        Me.lblMRDTitle.TabIndex = 0
        Me.lblMRDTitle.Text = "MRD"
        '
        'lblMRD
        '
        Me.lblMRD.AutoSize = True
        Me.lblMRD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMRD.ForeColor = System.Drawing.Color.Black
        Me.lblMRD.Location = New System.Drawing.Point(196, 4)
        Me.lblMRD.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMRD.Name = "lblMRD"
        Me.lblMRD.Size = New System.Drawing.Size(18, 18)
        Me.lblMRD.TabIndex = 7
        Me.lblMRD.Text = "()"
        '
        'grpM1
        '
        Me.grpM1.Controls.Add(Me.pnM1DC)
        Me.grpM1.Controls.Add(Me.pn1stM1)
        Me.grpM1.Location = New System.Drawing.Point(4, 41)
        Me.grpM1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpM1.Name = "grpM1"
        Me.grpM1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpM1.Size = New System.Drawing.Size(301, 101)
        Me.grpM1.TabIndex = 1
        Me.grpM1.TabStop = False
        Me.grpM1.Text = "M1 / X0 / XM / X1 / TPV"
        '
        'pnM1DC
        '
        Me.pnM1DC.Controls.Add(Me.btnClearM1DC)
        Me.pnM1DC.Controls.Add(Me.dtM1DC)
        Me.pnM1DC.Controls.Add(Me.lblM1DCTitle)
        Me.pnM1DC.Controls.Add(Me.lblM1DC)
        Me.pnM1DC.Location = New System.Drawing.Point(8, 60)
        Me.pnM1DC.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnM1DC.Name = "pnM1DC"
        Me.pnM1DC.Size = New System.Drawing.Size(283, 30)
        Me.pnM1DC.TabIndex = 1
        '
        'btnClearM1DC
        '
        Me.btnClearM1DC.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearM1DC.Location = New System.Drawing.Point(229, 5)
        Me.btnClearM1DC.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearM1DC.Name = "btnClearM1DC"
        Me.btnClearM1DC.Size = New System.Drawing.Size(31, 23)
        Me.btnClearM1DC.TabIndex = 2
        Me.btnClearM1DC.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearM1DC, "Clear date value")
        Me.btnClearM1DC.UseVisualStyleBackColor = True
        '
        'dtM1DC
        '
        Me.dtM1DC.CustomFormat = "dd MMM yyyy"
        Me.dtM1DC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtM1DC.Location = New System.Drawing.Point(55, 5)
        Me.dtM1DC.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtM1DC.Name = "dtM1DC"
        Me.dtM1DC.Size = New System.Drawing.Size(132, 22)
        Me.dtM1DC.TabIndex = 1
        '
        'lblM1DCTitle
        '
        Me.lblM1DCTitle.AutoSize = True
        Me.lblM1DCTitle.Location = New System.Drawing.Point(4, 5)
        Me.lblM1DCTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblM1DCTitle.Name = "lblM1DCTitle"
        Me.lblM1DCTitle.Size = New System.Drawing.Size(46, 17)
        Me.lblM1DCTitle.TabIndex = 0
        Me.lblM1DCTitle.Text = "M1DC"
        '
        'lblM1DC
        '
        Me.lblM1DC.AutoSize = True
        Me.lblM1DC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblM1DC.ForeColor = System.Drawing.Color.Black
        Me.lblM1DC.Location = New System.Drawing.Point(196, 5)
        Me.lblM1DC.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblM1DC.Name = "lblM1DC"
        Me.lblM1DC.Size = New System.Drawing.Size(18, 18)
        Me.lblM1DC.TabIndex = 7
        Me.lblM1DC.Text = "()"
        '
        'pn1stM1
        '
        Me.pn1stM1.Controls.Add(Me.btnClearM1)
        Me.pn1stM1.Controls.Add(Me.dtM1)
        Me.pn1stM1.Controls.Add(Me.lbl1stM1Title)
        Me.pn1stM1.Controls.Add(Me.lblM1)
        Me.pn1stM1.Location = New System.Drawing.Point(8, 23)
        Me.pn1stM1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pn1stM1.Name = "pn1stM1"
        Me.pn1stM1.Size = New System.Drawing.Size(283, 30)
        Me.pn1stM1.TabIndex = 0
        '
        'btnClearM1
        '
        Me.btnClearM1.Font = New System.Drawing.Font("Wingdings", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClearM1.Location = New System.Drawing.Point(229, 5)
        Me.btnClearM1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClearM1.Name = "btnClearM1"
        Me.btnClearM1.Size = New System.Drawing.Size(31, 23)
        Me.btnClearM1.TabIndex = 0
        Me.btnClearM1.Text = "û"
        Me.ToolTip1.SetToolTip(Me.btnClearM1, "Clear date value")
        Me.btnClearM1.UseVisualStyleBackColor = True
        '
        'dtM1
        '
        Me.dtM1.CustomFormat = "dd MMM yyyy"
        Me.dtM1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtM1.Location = New System.Drawing.Point(55, 5)
        Me.dtM1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtM1.Name = "dtM1"
        Me.dtM1.Size = New System.Drawing.Size(132, 22)
        Me.dtM1.TabIndex = 3
        '
        'lbl1stM1Title
        '
        Me.lbl1stM1Title.AutoSize = True
        Me.lbl1stM1Title.Location = New System.Drawing.Point(4, 5)
        Me.lbl1stM1Title.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl1stM1Title.Name = "lbl1stM1Title"
        Me.lbl1stM1Title.Size = New System.Drawing.Size(50, 17)
        Me.lbl1stM1Title.TabIndex = 2
        Me.lbl1stM1Title.Text = "1st M1"
        '
        'lblM1
        '
        Me.lblM1.AutoSize = True
        Me.lblM1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblM1.ForeColor = System.Drawing.Color.Black
        Me.lblM1.Location = New System.Drawing.Point(196, 5)
        Me.lblM1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblM1.Name = "lblM1"
        Me.lblM1.Size = New System.Drawing.Size(18, 18)
        Me.lblM1.TabIndex = 7
        Me.lblM1.Text = "()"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.brnOk)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 375)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(309, 54)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'brnOk
        '
        Me.brnOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.brnOk.Location = New System.Drawing.Point(56, 17)
        Me.brnOk.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.brnOk.Name = "brnOk"
        Me.brnOk.Size = New System.Drawing.Size(100, 28)
        Me.brnOk.TabIndex = 0
        Me.brnOk.Text = "&Ok"
        Me.ToolTip1.SetToolTip(Me.brnOk, "Ok [F7]")
        Me.brnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(164, 17)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "To close form [Esc]")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'chkInduFormatting
        '
        Me.chkInduFormatting.AutoSize = True
        Me.chkInduFormatting.Location = New System.Drawing.Point(21, 348)
        Me.chkInduFormatting.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkInduFormatting.Name = "chkInduFormatting"
        Me.chkInduFormatting.Size = New System.Drawing.Size(160, 21)
        Me.chkInduFormatting.TabIndex = 1
        Me.chkInduFormatting.Text = "Custom Formatting"
        Me.chkInduFormatting.UseVisualStyleBackColor = True
        '
        'frmPick1stVP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(333, 436)
        Me.Controls.Add(Me.chkInduFormatting)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.pnAll)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPick1stVP"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Dates"
        Me.pnAll.ResumeLayout(False)
        Me.pnAll.PerformLayout()
        Me.grpVP.ResumeLayout(False)
        Me.pnFEC.ResumeLayout(False)
        Me.pnFEC.PerformLayout()
        Me.pnPEC.ResumeLayout(False)
        Me.pnPEC.PerformLayout()
        Me.pn1stVP.ResumeLayout(False)
        Me.pn1stVP.PerformLayout()
        Me.pnMRD.ResumeLayout(False)
        Me.pnMRD.PerformLayout()
        Me.grpM1.ResumeLayout(False)
        Me.pnM1DC.ResumeLayout(False)
        Me.pnM1DC.PerformLayout()
        Me.pn1stM1.ResumeLayout(False)
        Me.pn1stM1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pnAll As System.Windows.Forms.Panel
    Friend WithEvents pnMRD As System.Windows.Forms.Panel
    Friend WithEvents lblMRD As System.Windows.Forms.Label
    Friend WithEvents grpM1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMRDTitle As System.Windows.Forms.Label
    Friend WithEvents pnM1DC As System.Windows.Forms.Panel
    Friend WithEvents lblM1DCTitle As System.Windows.Forms.Label
    Friend WithEvents lblM1DC As System.Windows.Forms.Label
    Friend WithEvents pn1stM1 As System.Windows.Forms.Panel
    Friend WithEvents lbl1stM1Title As System.Windows.Forms.Label
    Friend WithEvents lblM1 As System.Windows.Forms.Label
    Friend WithEvents grpVP As System.Windows.Forms.GroupBox
    Friend WithEvents pnFEC As System.Windows.Forms.Panel
    Friend WithEvents lblFECTitle As System.Windows.Forms.Label
    Friend WithEvents lblFEC As System.Windows.Forms.Label
    Friend WithEvents pnPEC As System.Windows.Forms.Panel
    Friend WithEvents lblPECTitle As System.Windows.Forms.Label
    Friend WithEvents lblPEC As System.Windows.Forms.Label
    Friend WithEvents pn1stVP As System.Windows.Forms.Panel
    Friend WithEvents lbl1stVPTitle As System.Windows.Forms.Label
    Friend WithEvents lblVP As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents brnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents dtMRD As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPEC As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtFEC As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtVP As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtM1DC As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtM1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
    Friend WithEvents btnClearFEC As System.Windows.Forms.Button
    Friend WithEvents btnClearPEC As System.Windows.Forms.Button
    Friend WithEvents btnClearVP As System.Windows.Forms.Button
    Friend WithEvents btnClearMRD As System.Windows.Forms.Button
    Friend WithEvents btnClearM1DC As System.Windows.Forms.Button
    Friend WithEvents btnClearM1 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents chkInduFormatting As System.Windows.Forms.CheckBox
    Friend WithEvents txtTnDPlanner As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
