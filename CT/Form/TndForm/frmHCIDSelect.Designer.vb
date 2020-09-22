<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmHCIDSelect
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHCIDSelect))
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TabController1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.grdPlans = New System.Windows.Forms.DataGridView()
        Me.pe01_TnDBasicProgram_ID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Read_Only = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.GenOrSpec = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HCID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DisplayHealthChartId = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PlanVersion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FileStatus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProgramDescription = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BuildPhase = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MRDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AssyBuildScale = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.M1DC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PECDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FECDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pe02 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BuildType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.XCCpe01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.XCCpe26 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Carline = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Platform = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Region = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.grdPlansGeneric = New System.Windows.Forms.DataGridView()
        Me.Gpe01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.GIsGeneric = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GHCID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GHCIDName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GBuildPhase = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn23 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn24 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GAssyBuildScale = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn26 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn27 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn28 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Gpe02 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GBuildType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GXccPe01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GXccPe26 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GCarline = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GPlatform = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GRegion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnBuckList = New System.Windows.Forms.CheckBox()
        Me.btnRigList = New System.Windows.Forms.CheckBox()
        Me.btnVehicleList = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnPreCheck = New System.Windows.Forms.Button()
        Me.SmoothProgressBar2 = New CT.SmoothProgressBar.SmoothProgressBar()
        Me.SmoothProgressBar1 = New CT.SmoothProgressBar.SmoothProgressBar()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.btnDraft = New System.Windows.Forms.Button()
        Me.btnCheckout = New System.Windows.Forms.Button()
        Me.btnOpenLoad = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.chkLoadIndFormat = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblXCCDBStatus = New System.Windows.Forms.Label()
        Me.txtHCName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtHcid = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CheckoutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DiscardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotifyToCheckout = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStripDraft = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.GenerateDraftToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoadDraftToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn17 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn18 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn19 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn20 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn21 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn22 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn25 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn29 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn30 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn31 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn32 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn33 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn34 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn35 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn36 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn37 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn38 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn39 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ActiveusersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1.SuspendLayout()
        Me.TabController1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.grdPlans, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.grdPlansGeneric, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ContextMenuStripDraft.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TabController1)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'TabController1
        '
        Me.TabController1.Controls.Add(Me.TabPage1)
        Me.TabController1.Controls.Add(Me.TabPage2)
        resources.ApplyResources(Me.TabController1, "TabController1")
        Me.TabController1.Name = "TabController1"
        Me.TabController1.SelectedIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.grdPlans)
        resources.ApplyResources(Me.TabPage1, "TabPage1")
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'grdPlans
        '
        Me.grdPlans.AllowUserToAddRows = False
        Me.grdPlans.AllowUserToDeleteRows = False
        Me.grdPlans.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdPlans.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdPlans.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.pe01_TnDBasicProgram_ID, Me.Read_Only, Me.GenOrSpec, Me.HCID, Me.DisplayHealthChartId, Me.PlanVersion, Me.FileStatus, Me.ProgramDescription, Me.BuildPhase, Me.MRDate, Me.Qty, Me.AssyBuildScale, Me.M1DC, Me.PECDate, Me.FECDate, Me.pe02, Me.BuildType, Me.XCCpe01, Me.XCCpe26, Me.Carline, Me.Platform, Me.Region})
        resources.ApplyResources(Me.grdPlans, "grdPlans")
        Me.grdPlans.MultiSelect = False
        Me.grdPlans.Name = "grdPlans"
        Me.grdPlans.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'pe01_TnDBasicProgram_ID
        '
        Me.pe01_TnDBasicProgram_ID.DataPropertyName = "pe01_TnDBasicProgram_FK"
        resources.ApplyResources(Me.pe01_TnDBasicProgram_ID, "pe01_TnDBasicProgram_ID")
        Me.pe01_TnDBasicProgram_ID.Name = "pe01_TnDBasicProgram_ID"
        Me.pe01_TnDBasicProgram_ID.ReadOnly = True
        '
        'Read_Only
        '
        Me.Read_Only.FillWeight = 30.0!
        resources.ApplyResources(Me.Read_Only, "Read_Only")
        Me.Read_Only.Name = "Read_Only"
        '
        'GenOrSpec
        '
        Me.GenOrSpec.DataPropertyName = "GenericSpecific"
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GenOrSpec.DefaultCellStyle = DataGridViewCellStyle2
        Me.GenOrSpec.FillWeight = 13.27627!
        resources.ApplyResources(Me.GenOrSpec, "GenOrSpec")
        Me.GenOrSpec.Name = "GenOrSpec"
        Me.GenOrSpec.ReadOnly = True
        '
        'HCID
        '
        Me.HCID.DataPropertyName = "HealthChartId"
        Me.HCID.FillWeight = 20.40755!
        resources.ApplyResources(Me.HCID, "HCID")
        Me.HCID.Name = "HCID"
        Me.HCID.ReadOnly = True
        '
        'DisplayHealthChartId
        '
        Me.DisplayHealthChartId.DataPropertyName = "DisplayHealthChartId"
        Me.DisplayHealthChartId.FillWeight = 20.40755!
        resources.ApplyResources(Me.DisplayHealthChartId, "DisplayHealthChartId")
        Me.DisplayHealthChartId.Name = "DisplayHealthChartId"
        Me.DisplayHealthChartId.ReadOnly = True
        '
        'PlanVersion
        '
        Me.PlanVersion.DataPropertyName = "PlanVersion"
        Me.PlanVersion.FillWeight = 21.15431!
        resources.ApplyResources(Me.PlanVersion, "PlanVersion")
        Me.PlanVersion.Name = "PlanVersion"
        Me.PlanVersion.ReadOnly = True
        '
        'FileStatus
        '
        Me.FileStatus.DataPropertyName = "FileStatus"
        Me.FileStatus.FillWeight = 21.15431!
        resources.ApplyResources(Me.FileStatus, "FileStatus")
        Me.FileStatus.Name = "FileStatus"
        Me.FileStatus.ReadOnly = True
        '
        'ProgramDescription
        '
        Me.ProgramDescription.DataPropertyName = "HealthChartName"
        Me.ProgramDescription.FillWeight = 50.35826!
        resources.ApplyResources(Me.ProgramDescription, "ProgramDescription")
        Me.ProgramDescription.Name = "ProgramDescription"
        Me.ProgramDescription.ReadOnly = True
        '
        'BuildPhase
        '
        Me.BuildPhase.DataPropertyName = "BuildPhase"
        Me.BuildPhase.FillWeight = 14.64967!
        resources.ApplyResources(Me.BuildPhase, "BuildPhase")
        Me.BuildPhase.Name = "BuildPhase"
        Me.BuildPhase.ReadOnly = True
        '
        'MRDate
        '
        Me.MRDate.DataPropertyName = "AssyMrd"
        Me.MRDate.FillWeight = 16.02308!
        resources.ApplyResources(Me.MRDate, "MRDate")
        Me.MRDate.Name = "MRDate"
        Me.MRDate.ReadOnly = True
        '
        'Qty
        '
        Me.Qty.DataPropertyName = "Quantity"
        Me.Qty.FillWeight = 9.156051!
        resources.ApplyResources(Me.Qty, "Qty")
        Me.Qty.Name = "Qty"
        Me.Qty.ReadOnly = True
        '
        'AssyBuildScale
        '
        Me.AssyBuildScale.DataPropertyName = "AssyBuildScale"
        Me.AssyBuildScale.FillWeight = 20.60111!
        resources.ApplyResources(Me.AssyBuildScale, "AssyBuildScale")
        Me.AssyBuildScale.Name = "AssyBuildScale"
        Me.AssyBuildScale.ReadOnly = True
        '
        'M1DC
        '
        Me.M1DC.DataPropertyName = "M1DC"
        Me.M1DC.FillWeight = 13.73407!
        resources.ApplyResources(Me.M1DC, "M1DC")
        Me.M1DC.Name = "M1DC"
        Me.M1DC.ReadOnly = True
        '
        'PECDate
        '
        Me.PECDate.DataPropertyName = "PEC"
        Me.PECDate.FillWeight = 14.19188!
        resources.ApplyResources(Me.PECDate, "PECDate")
        Me.PECDate.Name = "PECDate"
        Me.PECDate.ReadOnly = True
        '
        'FECDate
        '
        Me.FECDate.DataPropertyName = "FEC"
        Me.FECDate.FillWeight = 14.19188!
        resources.ApplyResources(Me.FECDate, "FECDate")
        Me.FECDate.Name = "FECDate"
        Me.FECDate.ReadOnly = True
        '
        'pe02
        '
        Me.pe02.DataPropertyName = "pe02"
        resources.ApplyResources(Me.pe02, "pe02")
        Me.pe02.Name = "pe02"
        Me.pe02.ReadOnly = True
        '
        'BuildType
        '
        Me.BuildType.DataPropertyName = "BuildType"
        resources.ApplyResources(Me.BuildType, "BuildType")
        Me.BuildType.Name = "BuildType"
        Me.BuildType.ReadOnly = True
        '
        'XCCpe01
        '
        Me.XCCpe01.DataPropertyName = "XCCpe01"
        resources.ApplyResources(Me.XCCpe01, "XCCpe01")
        Me.XCCpe01.Name = "XCCpe01"
        Me.XCCpe01.ReadOnly = True
        '
        'XCCpe26
        '
        Me.XCCpe26.DataPropertyName = "XCCpe26"
        resources.ApplyResources(Me.XCCpe26, "XCCpe26")
        Me.XCCpe26.Name = "XCCpe26"
        Me.XCCpe26.ReadOnly = True
        '
        'Carline
        '
        Me.Carline.DataPropertyName = "Carline"
        resources.ApplyResources(Me.Carline, "Carline")
        Me.Carline.Name = "Carline"
        Me.Carline.ReadOnly = True
        '
        'Platform
        '
        Me.Platform.DataPropertyName = "Platform"
        resources.ApplyResources(Me.Platform, "Platform")
        Me.Platform.Name = "Platform"
        Me.Platform.ReadOnly = True
        '
        'Region
        '
        Me.Region.DataPropertyName = "Region"
        Me.Region.FillWeight = 11.44506!
        resources.ApplyResources(Me.Region, "Region")
        Me.Region.Name = "Region"
        Me.Region.ReadOnly = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.grdPlansGeneric)
        resources.ApplyResources(Me.TabPage2, "TabPage2")
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'grdPlansGeneric
        '
        Me.grdPlansGeneric.AllowUserToAddRows = False
        Me.grdPlansGeneric.AllowUserToDeleteRows = False
        Me.grdPlansGeneric.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdPlansGeneric.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.grdPlansGeneric.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Gpe01, Me.DataGridViewCheckBoxColumn1, Me.GIsGeneric, Me.GHCID, Me.GHCIDName, Me.GBuildPhase, Me.DataGridViewTextBoxColumn23, Me.DataGridViewTextBoxColumn24, Me.GAssyBuildScale, Me.DataGridViewTextBoxColumn26, Me.DataGridViewTextBoxColumn27, Me.DataGridViewTextBoxColumn28, Me.Gpe02, Me.GBuildType, Me.GXccPe01, Me.GXccPe26, Me.GCarline, Me.GPlatform, Me.GRegion})
        resources.ApplyResources(Me.grdPlansGeneric, "grdPlansGeneric")
        Me.grdPlansGeneric.MultiSelect = False
        Me.grdPlansGeneric.Name = "grdPlansGeneric"
        Me.grdPlansGeneric.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'Gpe01
        '
        Me.Gpe01.DataPropertyName = "pe01_TnDBasicProgram_FK"
        Me.Gpe01.FillWeight = 25.0!
        resources.ApplyResources(Me.Gpe01, "Gpe01")
        Me.Gpe01.Name = "Gpe01"
        Me.Gpe01.ReadOnly = True
        '
        'DataGridViewCheckBoxColumn1
        '
        Me.DataGridViewCheckBoxColumn1.FillWeight = 30.0!
        resources.ApplyResources(Me.DataGridViewCheckBoxColumn1, "DataGridViewCheckBoxColumn1")
        Me.DataGridViewCheckBoxColumn1.Name = "DataGridViewCheckBoxColumn1"
        '
        'GIsGeneric
        '
        Me.GIsGeneric.DataPropertyName = "GenericSpecific"
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GIsGeneric.DefaultCellStyle = DataGridViewCellStyle4
        Me.GIsGeneric.FillWeight = 29.0!
        resources.ApplyResources(Me.GIsGeneric, "GIsGeneric")
        Me.GIsGeneric.Name = "GIsGeneric"
        Me.GIsGeneric.ReadOnly = True
        '
        'GHCID
        '
        Me.GHCID.DataPropertyName = "HealthChartId"
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GHCID.DefaultCellStyle = DataGridViewCellStyle5
        Me.GHCID.FillWeight = 23.0!
        resources.ApplyResources(Me.GHCID, "GHCID")
        Me.GHCID.Name = "GHCID"
        Me.GHCID.ReadOnly = True
        '
        'GHCIDName
        '
        Me.GHCIDName.DataPropertyName = "HealthChartName"
        Me.GHCIDName.FillWeight = 110.0!
        resources.ApplyResources(Me.GHCIDName, "GHCIDName")
        Me.GHCIDName.Name = "GHCIDName"
        Me.GHCIDName.ReadOnly = True
        '
        'GBuildPhase
        '
        Me.GBuildPhase.DataPropertyName = "BuildPhase"
        Me.GBuildPhase.FillWeight = 32.0!
        resources.ApplyResources(Me.GBuildPhase, "GBuildPhase")
        Me.GBuildPhase.Name = "GBuildPhase"
        Me.GBuildPhase.ReadOnly = True
        '
        'DataGridViewTextBoxColumn23
        '
        Me.DataGridViewTextBoxColumn23.DataPropertyName = "AssyMrd"
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridViewTextBoxColumn23.DefaultCellStyle = DataGridViewCellStyle6
        Me.DataGridViewTextBoxColumn23.FillWeight = 35.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn23, "DataGridViewTextBoxColumn23")
        Me.DataGridViewTextBoxColumn23.Name = "DataGridViewTextBoxColumn23"
        Me.DataGridViewTextBoxColumn23.ReadOnly = True
        '
        'DataGridViewTextBoxColumn24
        '
        Me.DataGridViewTextBoxColumn24.DataPropertyName = "Quantity"
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridViewTextBoxColumn24.DefaultCellStyle = DataGridViewCellStyle7
        Me.DataGridViewTextBoxColumn24.FillWeight = 20.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn24, "DataGridViewTextBoxColumn24")
        Me.DataGridViewTextBoxColumn24.Name = "DataGridViewTextBoxColumn24"
        Me.DataGridViewTextBoxColumn24.ReadOnly = True
        '
        'GAssyBuildScale
        '
        Me.GAssyBuildScale.DataPropertyName = "AssyBuildScale"
        Me.GAssyBuildScale.FillWeight = 45.0!
        resources.ApplyResources(Me.GAssyBuildScale, "GAssyBuildScale")
        Me.GAssyBuildScale.Name = "GAssyBuildScale"
        Me.GAssyBuildScale.ReadOnly = True
        '
        'DataGridViewTextBoxColumn26
        '
        Me.DataGridViewTextBoxColumn26.DataPropertyName = "M1DC"
        Me.DataGridViewTextBoxColumn26.FillWeight = 30.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn26, "DataGridViewTextBoxColumn26")
        Me.DataGridViewTextBoxColumn26.Name = "DataGridViewTextBoxColumn26"
        Me.DataGridViewTextBoxColumn26.ReadOnly = True
        '
        'DataGridViewTextBoxColumn27
        '
        Me.DataGridViewTextBoxColumn27.DataPropertyName = "PEC"
        Me.DataGridViewTextBoxColumn27.FillWeight = 31.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn27, "DataGridViewTextBoxColumn27")
        Me.DataGridViewTextBoxColumn27.Name = "DataGridViewTextBoxColumn27"
        Me.DataGridViewTextBoxColumn27.ReadOnly = True
        '
        'DataGridViewTextBoxColumn28
        '
        Me.DataGridViewTextBoxColumn28.DataPropertyName = "FEC"
        Me.DataGridViewTextBoxColumn28.FillWeight = 31.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn28, "DataGridViewTextBoxColumn28")
        Me.DataGridViewTextBoxColumn28.Name = "DataGridViewTextBoxColumn28"
        Me.DataGridViewTextBoxColumn28.ReadOnly = True
        '
        'Gpe02
        '
        Me.Gpe02.DataPropertyName = "pe02"
        Me.Gpe02.FillWeight = 31.0!
        resources.ApplyResources(Me.Gpe02, "Gpe02")
        Me.Gpe02.Name = "Gpe02"
        Me.Gpe02.ReadOnly = True
        '
        'GBuildType
        '
        Me.GBuildType.DataPropertyName = "BuildType"
        resources.ApplyResources(Me.GBuildType, "GBuildType")
        Me.GBuildType.Name = "GBuildType"
        Me.GBuildType.ReadOnly = True
        '
        'GXccPe01
        '
        Me.GXccPe01.DataPropertyName = "XCCpe01"
        resources.ApplyResources(Me.GXccPe01, "GXccPe01")
        Me.GXccPe01.Name = "GXccPe01"
        Me.GXccPe01.ReadOnly = True
        '
        'GXccPe26
        '
        Me.GXccPe26.DataPropertyName = "XCCpe26"
        resources.ApplyResources(Me.GXccPe26, "GXccPe26")
        Me.GXccPe26.Name = "GXccPe26"
        Me.GXccPe26.ReadOnly = True
        '
        'GCarline
        '
        Me.GCarline.DataPropertyName = "Carline"
        resources.ApplyResources(Me.GCarline, "GCarline")
        Me.GCarline.Name = "GCarline"
        Me.GCarline.ReadOnly = True
        '
        'GPlatform
        '
        Me.GPlatform.DataPropertyName = "Platform"
        resources.ApplyResources(Me.GPlatform, "GPlatform")
        Me.GPlatform.Name = "GPlatform"
        Me.GPlatform.ReadOnly = True
        '
        'GRegion
        '
        Me.GRegion.DataPropertyName = "Region"
        Me.GRegion.FillWeight = 25.0!
        resources.ApplyResources(Me.GRegion, "GRegion")
        Me.GRegion.Name = "GRegion"
        Me.GRegion.ReadOnly = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.btnBuckList)
        Me.Panel1.Controls.Add(Me.btnRigList)
        Me.Panel1.Controls.Add(Me.btnVehicleList)
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Name = "Panel1"
        '
        'btnBuckList
        '
        resources.ApplyResources(Me.btnBuckList, "btnBuckList")
        Me.btnBuckList.Name = "btnBuckList"
        Me.btnBuckList.UseVisualStyleBackColor = True
        '
        'btnRigList
        '
        resources.ApplyResources(Me.btnRigList, "btnRigList")
        Me.btnRigList.Name = "btnRigList"
        Me.btnRigList.UseVisualStyleBackColor = True
        '
        'btnVehicleList
        '
        resources.ApplyResources(Me.btnVehicleList, "btnVehicleList")
        Me.btnVehicleList.Checked = True
        Me.btnVehicleList.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btnVehicleList.Name = "btnVehicleList"
        Me.btnVehicleList.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnPreCheck)
        Me.GroupBox2.Controls.Add(Me.SmoothProgressBar2)
        Me.GroupBox2.Controls.Add(Me.SmoothProgressBar1)
        Me.GroupBox2.Controls.Add(Me.lblProgress)
        Me.GroupBox2.Controls.Add(Me.btnDraft)
        Me.GroupBox2.Controls.Add(Me.btnCheckout)
        Me.GroupBox2.Controls.Add(Me.btnOpenLoad)
        Me.GroupBox2.Controls.Add(Me.btnCancel)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'btnPreCheck
        '
        resources.ApplyResources(Me.btnPreCheck, "btnPreCheck")
        Me.btnPreCheck.Name = "btnPreCheck"
        Me.ToolTip1.SetToolTip(Me.btnPreCheck, resources.GetString("btnPreCheck.ToolTip"))
        Me.btnPreCheck.UseVisualStyleBackColor = True
        '
        'SmoothProgressBar2
        '
        resources.ApplyResources(Me.SmoothProgressBar2, "SmoothProgressBar2")
        Me.SmoothProgressBar2.Maximum = 100
        Me.SmoothProgressBar2.Minimum = 0
        Me.SmoothProgressBar2.Name = "SmoothProgressBar2"
        Me.SmoothProgressBar2.ProgressBarColor = System.Drawing.Color.Blue
        Me.SmoothProgressBar2.Value = 0R
        '
        'SmoothProgressBar1
        '
        resources.ApplyResources(Me.SmoothProgressBar1, "SmoothProgressBar1")
        Me.SmoothProgressBar1.Maximum = 100
        Me.SmoothProgressBar1.Minimum = 0
        Me.SmoothProgressBar1.Name = "SmoothProgressBar1"
        Me.SmoothProgressBar1.ProgressBarColor = System.Drawing.Color.Blue
        Me.SmoothProgressBar1.Value = 0R
        '
        'lblProgress
        '
        resources.ApplyResources(Me.lblProgress, "lblProgress")
        Me.lblProgress.BackColor = System.Drawing.Color.Transparent
        Me.lblProgress.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblProgress.Name = "lblProgress"
        '
        'btnDraft
        '
        resources.ApplyResources(Me.btnDraft, "btnDraft")
        Me.btnDraft.Name = "btnDraft"
        Me.ToolTip1.SetToolTip(Me.btnDraft, resources.GetString("btnDraft.ToolTip"))
        Me.btnDraft.UseVisualStyleBackColor = True
        '
        'btnCheckout
        '
        resources.ApplyResources(Me.btnCheckout, "btnCheckout")
        Me.btnCheckout.Name = "btnCheckout"
        Me.ToolTip1.SetToolTip(Me.btnCheckout, resources.GetString("btnCheckout.ToolTip"))
        Me.btnCheckout.UseVisualStyleBackColor = True
        '
        'btnOpenLoad
        '
        Me.btnOpenLoad.DialogResult = System.Windows.Forms.DialogResult.OK
        resources.ApplyResources(Me.btnOpenLoad, "btnOpenLoad")
        Me.btnOpenLoad.Name = "btnOpenLoad"
        Me.ToolTip1.SetToolTip(Me.btnOpenLoad, resources.GetString("btnOpenLoad.ToolTip"))
        Me.btnOpenLoad.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        resources.ApplyResources(Me.btnCancel, "btnCancel")
        Me.btnCancel.Name = "btnCancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, resources.GetString("btnCancel.ToolTip"))
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'chkLoadIndFormat
        '
        resources.ApplyResources(Me.chkLoadIndFormat, "chkLoadIndFormat")
        Me.chkLoadIndFormat.Name = "chkLoadIndFormat"
        Me.chkLoadIndFormat.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.lblXCCDBStatus)
        Me.GroupBox3.Controls.Add(Me.txtHCName)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.txtHcid)
        Me.GroupBox3.Controls.Add(Me.Label1)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.ForeColor = System.Drawing.Color.DarkGray
        Me.Label3.Name = "Label3"
        '
        'lblXCCDBStatus
        '
        resources.ApplyResources(Me.lblXCCDBStatus, "lblXCCDBStatus")
        Me.lblXCCDBStatus.ForeColor = System.Drawing.Color.Red
        Me.lblXCCDBStatus.Name = "lblXCCDBStatus"
        '
        'txtHCName
        '
        resources.ApplyResources(Me.txtHCName, "txtHCName")
        Me.txtHCName.Name = "txtHCName"
        Me.ToolTip1.SetToolTip(Me.txtHCName, resources.GetString("txtHCName.ToolTip"))
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'txtHcid
        '
        resources.ApplyResources(Me.txtHcid, "txtHcid")
        Me.txtHcid.Name = "txtHcid"
        Me.ToolTip1.SetToolTip(Me.txtHcid, resources.GetString("txtHcid.ToolTip"))
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.White
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CheckoutToolStripMenuItem, Me.CheckinToolStripMenuItem, Me.DiscardToolStripMenuItem, Me.ActiveusersToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        resources.ApplyResources(Me.ContextMenuStrip1, "ContextMenuStrip1")
        '
        'CheckoutToolStripMenuItem
        '
        Me.CheckoutToolStripMenuItem.Name = "CheckoutToolStripMenuItem"
        resources.ApplyResources(Me.CheckoutToolStripMenuItem, "CheckoutToolStripMenuItem")
        '
        'CheckinToolStripMenuItem
        '
        Me.CheckinToolStripMenuItem.Name = "CheckinToolStripMenuItem"
        resources.ApplyResources(Me.CheckinToolStripMenuItem, "CheckinToolStripMenuItem")
        '
        'DiscardToolStripMenuItem
        '
        Me.DiscardToolStripMenuItem.Name = "DiscardToolStripMenuItem"
        resources.ApplyResources(Me.DiscardToolStripMenuItem, "DiscardToolStripMenuItem")
        '
        'NotifyToCheckout
        '
        Me.NotifyToCheckout.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        resources.ApplyResources(Me.NotifyToCheckout, "NotifyToCheckout")
        '
        'ContextMenuStripDraft
        '
        Me.ContextMenuStripDraft.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStripDraft.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GenerateDraftToolStripMenuItem, Me.LoadDraftToolStripMenuItem})
        Me.ContextMenuStripDraft.Name = "ContextMenuStripDraft"
        resources.ApplyResources(Me.ContextMenuStripDraft, "ContextMenuStripDraft")
        '
        'GenerateDraftToolStripMenuItem
        '
        Me.GenerateDraftToolStripMenuItem.Name = "GenerateDraftToolStripMenuItem"
        resources.ApplyResources(Me.GenerateDraftToolStripMenuItem, "GenerateDraftToolStripMenuItem")
        '
        'LoadDraftToolStripMenuItem
        '
        Me.LoadDraftToolStripMenuItem.Name = "LoadDraftToolStripMenuItem"
        resources.ApplyResources(Me.LoadDraftToolStripMenuItem, "LoadDraftToolStripMenuItem")
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "pe01_TnDBasicProgram_FK"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn1, "DataGridViewTextBoxColumn1")
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "GenericSpecific"
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridViewTextBoxColumn2.DefaultCellStyle = DataGridViewCellStyle8
        Me.DataGridViewTextBoxColumn2.FillWeight = 30.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn2, "DataGridViewTextBoxColumn2")
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "HealthChartId"
        Me.DataGridViewTextBoxColumn3.FillWeight = 18.21904!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn3, "DataGridViewTextBoxColumn3")
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "HealthChartName"
        Me.DataGridViewTextBoxColumn4.FillWeight = 110.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn4, "DataGridViewTextBoxColumn4")
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "BuildPhase"
        Me.DataGridViewTextBoxColumn5.FillWeight = 25.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn5, "DataGridViewTextBoxColumn5")
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.DataPropertyName = "AssyMrd"
        Me.DataGridViewTextBoxColumn6.FillWeight = 39.06458!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn6, "DataGridViewTextBoxColumn6")
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.DataPropertyName = "Quantity"
        Me.DataGridViewTextBoxColumn7.FillWeight = 20.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn7, "DataGridViewTextBoxColumn7")
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.DataPropertyName = "AssyBuildScale"
        Me.DataGridViewTextBoxColumn8.FillWeight = 45.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn8, "DataGridViewTextBoxColumn8")
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.ReadOnly = True
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.DataPropertyName = "M1DC"
        Me.DataGridViewTextBoxColumn9.FillWeight = 35.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn9, "DataGridViewTextBoxColumn9")
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.ReadOnly = True
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.DataPropertyName = "PEC"
        Me.DataGridViewTextBoxColumn10.FillWeight = 35.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn10, "DataGridViewTextBoxColumn10")
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.ReadOnly = True
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.DataPropertyName = "FEC"
        Me.DataGridViewTextBoxColumn11.FillWeight = 33.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn11, "DataGridViewTextBoxColumn11")
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        Me.DataGridViewTextBoxColumn11.ReadOnly = True
        '
        'DataGridViewTextBoxColumn12
        '
        Me.DataGridViewTextBoxColumn12.DataPropertyName = "pe02"
        Me.DataGridViewTextBoxColumn12.FillWeight = 13.73407!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn12, "DataGridViewTextBoxColumn12")
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        Me.DataGridViewTextBoxColumn12.ReadOnly = True
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.DataPropertyName = "BuildType"
        Me.DataGridViewTextBoxColumn13.FillWeight = 14.19188!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn13, "DataGridViewTextBoxColumn13")
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.ReadOnly = True
        '
        'DataGridViewTextBoxColumn14
        '
        Me.DataGridViewTextBoxColumn14.DataPropertyName = "XCCpe01"
        Me.DataGridViewTextBoxColumn14.FillWeight = 14.19188!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn14, "DataGridViewTextBoxColumn14")
        Me.DataGridViewTextBoxColumn14.Name = "DataGridViewTextBoxColumn14"
        Me.DataGridViewTextBoxColumn14.ReadOnly = True
        '
        'DataGridViewTextBoxColumn15
        '
        Me.DataGridViewTextBoxColumn15.DataPropertyName = "XCCpe26"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn15, "DataGridViewTextBoxColumn15")
        Me.DataGridViewTextBoxColumn15.Name = "DataGridViewTextBoxColumn15"
        Me.DataGridViewTextBoxColumn15.ReadOnly = True
        '
        'DataGridViewTextBoxColumn16
        '
        Me.DataGridViewTextBoxColumn16.DataPropertyName = "Carline"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn16, "DataGridViewTextBoxColumn16")
        Me.DataGridViewTextBoxColumn16.Name = "DataGridViewTextBoxColumn16"
        Me.DataGridViewTextBoxColumn16.ReadOnly = True
        '
        'DataGridViewTextBoxColumn17
        '
        Me.DataGridViewTextBoxColumn17.DataPropertyName = "Platform"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn17, "DataGridViewTextBoxColumn17")
        Me.DataGridViewTextBoxColumn17.Name = "DataGridViewTextBoxColumn17"
        Me.DataGridViewTextBoxColumn17.ReadOnly = True
        '
        'DataGridViewTextBoxColumn18
        '
        Me.DataGridViewTextBoxColumn18.DataPropertyName = "XCCpe26"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn18, "DataGridViewTextBoxColumn18")
        Me.DataGridViewTextBoxColumn18.Name = "DataGridViewTextBoxColumn18"
        Me.DataGridViewTextBoxColumn18.ReadOnly = True
        '
        'DataGridViewTextBoxColumn19
        '
        Me.DataGridViewTextBoxColumn19.DataPropertyName = "Carline"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn19, "DataGridViewTextBoxColumn19")
        Me.DataGridViewTextBoxColumn19.Name = "DataGridViewTextBoxColumn19"
        Me.DataGridViewTextBoxColumn19.ReadOnly = True
        '
        'DataGridViewTextBoxColumn20
        '
        Me.DataGridViewTextBoxColumn20.DataPropertyName = "Platform"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn20, "DataGridViewTextBoxColumn20")
        Me.DataGridViewTextBoxColumn20.Name = "DataGridViewTextBoxColumn20"
        Me.DataGridViewTextBoxColumn20.ReadOnly = True
        '
        'DataGridViewTextBoxColumn21
        '
        Me.DataGridViewTextBoxColumn21.DataPropertyName = "Region"
        Me.DataGridViewTextBoxColumn21.FillWeight = 11.44506!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn21, "DataGridViewTextBoxColumn21")
        Me.DataGridViewTextBoxColumn21.Name = "DataGridViewTextBoxColumn21"
        Me.DataGridViewTextBoxColumn21.ReadOnly = True
        '
        'DataGridViewTextBoxColumn22
        '
        Me.DataGridViewTextBoxColumn22.DataPropertyName = "pe01_TnDBasicProgram_FK"
        Me.DataGridViewTextBoxColumn22.FillWeight = 25.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn22, "DataGridViewTextBoxColumn22")
        Me.DataGridViewTextBoxColumn22.Name = "DataGridViewTextBoxColumn22"
        Me.DataGridViewTextBoxColumn22.ReadOnly = True
        '
        'DataGridViewTextBoxColumn25
        '
        Me.DataGridViewTextBoxColumn25.DataPropertyName = "HealthChartName"
        Me.DataGridViewTextBoxColumn25.FillWeight = 110.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn25, "DataGridViewTextBoxColumn25")
        Me.DataGridViewTextBoxColumn25.Name = "DataGridViewTextBoxColumn25"
        Me.DataGridViewTextBoxColumn25.ReadOnly = True
        '
        'DataGridViewTextBoxColumn29
        '
        Me.DataGridViewTextBoxColumn29.DataPropertyName = "AssyBuildScale"
        Me.DataGridViewTextBoxColumn29.FillWeight = 45.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn29, "DataGridViewTextBoxColumn29")
        Me.DataGridViewTextBoxColumn29.Name = "DataGridViewTextBoxColumn29"
        Me.DataGridViewTextBoxColumn29.ReadOnly = True
        '
        'DataGridViewTextBoxColumn30
        '
        Me.DataGridViewTextBoxColumn30.DataPropertyName = "M1DC"
        Me.DataGridViewTextBoxColumn30.FillWeight = 30.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn30, "DataGridViewTextBoxColumn30")
        Me.DataGridViewTextBoxColumn30.Name = "DataGridViewTextBoxColumn30"
        Me.DataGridViewTextBoxColumn30.ReadOnly = True
        '
        'DataGridViewTextBoxColumn31
        '
        Me.DataGridViewTextBoxColumn31.DataPropertyName = "PEC"
        Me.DataGridViewTextBoxColumn31.FillWeight = 31.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn31, "DataGridViewTextBoxColumn31")
        Me.DataGridViewTextBoxColumn31.Name = "DataGridViewTextBoxColumn31"
        Me.DataGridViewTextBoxColumn31.ReadOnly = True
        '
        'DataGridViewTextBoxColumn32
        '
        Me.DataGridViewTextBoxColumn32.DataPropertyName = "FEC"
        Me.DataGridViewTextBoxColumn32.FillWeight = 31.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn32, "DataGridViewTextBoxColumn32")
        Me.DataGridViewTextBoxColumn32.Name = "DataGridViewTextBoxColumn32"
        Me.DataGridViewTextBoxColumn32.ReadOnly = True
        '
        'DataGridViewTextBoxColumn33
        '
        Me.DataGridViewTextBoxColumn33.DataPropertyName = "pe02"
        Me.DataGridViewTextBoxColumn33.FillWeight = 31.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn33, "DataGridViewTextBoxColumn33")
        Me.DataGridViewTextBoxColumn33.Name = "DataGridViewTextBoxColumn33"
        Me.DataGridViewTextBoxColumn33.ReadOnly = True
        '
        'DataGridViewTextBoxColumn34
        '
        Me.DataGridViewTextBoxColumn34.DataPropertyName = "BuildType"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn34, "DataGridViewTextBoxColumn34")
        Me.DataGridViewTextBoxColumn34.Name = "DataGridViewTextBoxColumn34"
        Me.DataGridViewTextBoxColumn34.ReadOnly = True
        '
        'DataGridViewTextBoxColumn35
        '
        Me.DataGridViewTextBoxColumn35.DataPropertyName = "Platform"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn35, "DataGridViewTextBoxColumn35")
        Me.DataGridViewTextBoxColumn35.Name = "DataGridViewTextBoxColumn35"
        Me.DataGridViewTextBoxColumn35.ReadOnly = True
        '
        'DataGridViewTextBoxColumn36
        '
        Me.DataGridViewTextBoxColumn36.DataPropertyName = "XCCpe26"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn36, "DataGridViewTextBoxColumn36")
        Me.DataGridViewTextBoxColumn36.Name = "DataGridViewTextBoxColumn36"
        Me.DataGridViewTextBoxColumn36.ReadOnly = True
        '
        'DataGridViewTextBoxColumn37
        '
        Me.DataGridViewTextBoxColumn37.DataPropertyName = "Carline"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn37, "DataGridViewTextBoxColumn37")
        Me.DataGridViewTextBoxColumn37.Name = "DataGridViewTextBoxColumn37"
        Me.DataGridViewTextBoxColumn37.ReadOnly = True
        '
        'DataGridViewTextBoxColumn38
        '
        Me.DataGridViewTextBoxColumn38.DataPropertyName = "Platform"
        resources.ApplyResources(Me.DataGridViewTextBoxColumn38, "DataGridViewTextBoxColumn38")
        Me.DataGridViewTextBoxColumn38.Name = "DataGridViewTextBoxColumn38"
        Me.DataGridViewTextBoxColumn38.ReadOnly = True
        '
        'DataGridViewTextBoxColumn39
        '
        Me.DataGridViewTextBoxColumn39.DataPropertyName = "Region"
        Me.DataGridViewTextBoxColumn39.FillWeight = 25.0!
        resources.ApplyResources(Me.DataGridViewTextBoxColumn39, "DataGridViewTextBoxColumn39")
        Me.DataGridViewTextBoxColumn39.Name = "DataGridViewTextBoxColumn39"
        Me.DataGridViewTextBoxColumn39.ReadOnly = True
        '
        'ActiveusersToolStripMenuItem
        '
        Me.ActiveusersToolStripMenuItem.Name = "ActiveusersToolStripMenuItem"
        resources.ApplyResources(Me.ActiveusersToolStripMenuItem, "ActiveusersToolStripMenuItem")
        '
        'frmHCIDSelect
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.chkLoadIndFormat)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHCIDSelect"
        Me.ShowInTaskbar = False
        Me.GroupBox1.ResumeLayout(False)
        Me.TabController1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.grdPlans, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.grdPlansGeneric, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ContextMenuStripDraft.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnOpenLoad As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtHCName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtHcid As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents chkLoadIndFormat As System.Windows.Forms.CheckBox
    Friend WithEvents SmoothProgressBar2 As SmoothProgressBar.SmoothProgressBar
    Friend WithEvents SmoothProgressBar1 As SmoothProgressBar.SmoothProgressBar
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn16 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn17 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnPreCheck As System.Windows.Forms.Button
    Friend WithEvents TabController1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents grdPlansGeneric As System.Windows.Forms.DataGridView
    Friend WithEvents lblXCCDBStatus As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewTextBoxColumn35 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Gpe01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents GIsGeneric As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GHCID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GHCIDName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GBuildPhase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn23 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn24 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GAssyBuildScale As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn26 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn27 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn28 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Gpe02 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GBuildType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GXccPe01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GXccPe26 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GCarline As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GPlatform As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GRegion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnCheckout As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CheckoutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckinToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DiscardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyToCheckout As System.Windows.Forms.NotifyIcon
    Friend WithEvents btnDraft As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStripDraft As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents GenerateDraftToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoadDraftToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DataGridViewTextBoxColumn18 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn19 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn20 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn21 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn22 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn25 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn29 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn30 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn31 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn32 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn33 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn34 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn36 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn37 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn38 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn39 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnBuckList As System.Windows.Forms.CheckBox
    Friend WithEvents btnRigList As System.Windows.Forms.CheckBox
    Friend WithEvents btnVehicleList As System.Windows.Forms.CheckBox
    Friend WithEvents grdPlans As System.Windows.Forms.DataGridView
    Friend WithEvents pe01_TnDBasicProgram_ID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Read_Only As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents GenOrSpec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HCID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DisplayHealthChartId As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PlanVersion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FileStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProgramDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BuildPhase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MRDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AssyBuildScale As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents M1DC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PECDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FECDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pe02 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BuildType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents XCCpe01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents XCCpe26 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Carline As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Platform As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Region As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ActiveusersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
