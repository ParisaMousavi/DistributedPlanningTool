Partial Class RbnTnDControlPanel
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()
    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.tabTndPlanControlPanel = Me.Factory.CreateRibbonTab
        Me.GRPProgramTDToolbar = Me.Factory.CreateRibbonGroup
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.Separator4 = Me.Factory.CreateRibbonSeparator
        Me.GRPSearchFilterHighlight = Me.Factory.CreateRibbonGroup
        Me.GRPAddDeleteUpdateVehicle = Me.Factory.CreateRibbonGroup
        Me.GRPPlan = Me.Factory.CreateRibbonGroup
        Me.GRPShowHideSpecificationsections = Me.Factory.CreateRibbonGroup
        Me.Separator7 = Me.Factory.CreateRibbonSeparator
        Me.GRPIndicator = Me.Factory.CreateRibbonGroup
        Me.GrpMessages = Me.Factory.CreateRibbonGroup
        Me.GRPHelp = Me.Factory.CreateRibbonGroup
        Me.tabTndPlanReports = Me.Factory.CreateRibbonTab
        Me.GRPReport = Me.Factory.CreateRibbonGroup
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.GRPHolidayMaster = Me.Factory.CreateRibbonGroup
        Me.tmrDisplay = New System.Windows.Forms.Timer(Me.components)
        Me.tmrMessages = New System.Windows.Forms.Timer(Me.components)
        Me.btnLoadOpenTnDPlan = Me.Factory.CreateRibbonButton
        Me.btnConvertToSpecific = Me.Factory.CreateRibbonButton
        Me.menuDraft = Me.Factory.CreateRibbonMenu
        Me.btnGenerateDraft = Me.Factory.CreateRibbonButton
        Me.mnuGenerateDraft = Me.Factory.CreateRibbonMenu
        Me.btnDeleteDraft = Me.Factory.CreateRibbonButton
        Me.btnReplacePlanWithDraft = Me.Factory.CreateRibbonButton
        Me.menuCheckInOut = Me.Factory.CreateRibbonMenu
        Me.btnCheckOut = Me.Factory.CreateRibbonButton
        Me.btnCheckIn = Me.Factory.CreateRibbonButton
        Me.btnDiscard = Me.Factory.CreateRibbonButton
        Me.mnuActiveUsers = Me.Factory.CreateRibbonMenu
        Me.btnRefreshPlan = Me.Factory.CreateRibbonButton
        Me.btnRefreshUnit = Me.Factory.CreateRibbonButton
        Me.btnUndo = Me.Factory.CreateRibbonButton
        Me.btnRedo = Me.Factory.CreateRibbonButton
        Me.btnSearchFilter = Me.Factory.CreateRibbonButton
        Me.btnClearFilter = Me.Factory.CreateRibbonButton
        Me.btnAddUnit = Me.Factory.CreateRibbonButton
        Me.btnDeleteUnit = Me.Factory.CreateRibbonButton
        Me.btnChangeSequence = Me.Factory.CreateRibbonButton
        Me.btnUpdateMRD = Me.Factory.CreateRibbonButton
        Me.togInstrumentation = Me.Factory.CreateRibbonToggleButton
        Me.togNonMFCSpecification = Me.Factory.CreateRibbonToggleButton
        Me.togMfcSpecification = Me.Factory.CreateRibbonToggleButton
        Me.togProgramInformation = Me.Factory.CreateRibbonToggleButton
        Me.togFurtherBasicSpecification = Me.Factory.CreateRibbonToggleButton
        Me.togUserShipping = Me.Factory.CreateRibbonToggleButton
        Me.togUpdatePack = Me.Factory.CreateRibbonToggleButton
        Me.togTiming = Me.Factory.CreateRibbonToggleButton
        Me.togShowAll = Me.Factory.CreateRibbonToggleButton
        Me.btnUpdateColumns = Me.Factory.CreateRibbonButton
        Me.btnTodayIndicator = Me.Factory.CreateRibbonToggleButton
        Me.TGMessages = Me.Factory.CreateRibbonToggleButton
        Me.btnCTHelp = Me.Factory.CreateRibbonButton
        Me.btnExportToExcel = Me.Factory.CreateRibbonButton
        Me.btnUnitReport = Me.Factory.CreateRibbonButton
        Me.btnEngineTransmissionReport = Me.Factory.CreateRibbonButton
        Me.btnPrecheckF4T = Me.Factory.CreateRibbonButton
        Me.btnCountReport = Me.Factory.CreateRibbonButton
        Me.btnPustFit4Test = Me.Factory.CreateRibbonButton
        Me.tglBtnValidatePlan = Me.Factory.CreateRibbonToggleButton
        Me.btnUpdateHoliday = Me.Factory.CreateRibbonButton
        Me.btnCDSIDtoDvpTeam = Me.Factory.CreateRibbonButton
        Me.tabTndPlanControlPanel.SuspendLayout()
        Me.GRPProgramTDToolbar.SuspendLayout()
        Me.GRPSearchFilterHighlight.SuspendLayout()
        Me.GRPAddDeleteUpdateVehicle.SuspendLayout()
        Me.GRPPlan.SuspendLayout()
        Me.GRPShowHideSpecificationsections.SuspendLayout()
        Me.GRPIndicator.SuspendLayout()
        Me.GrpMessages.SuspendLayout()
        Me.GRPHelp.SuspendLayout()
        Me.tabTndPlanReports.SuspendLayout()
        Me.GRPReport.SuspendLayout()
        Me.GRPHolidayMaster.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabTndPlanControlPanel
        '
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPProgramTDToolbar)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPSearchFilterHighlight)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPAddDeleteUpdateVehicle)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPPlan)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPShowHideSpecificationsections)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPIndicator)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GrpMessages)
        Me.tabTndPlanControlPanel.Groups.Add(Me.GRPHelp)
        Me.tabTndPlanControlPanel.Label = "CT Plan Control Panel"
        Me.tabTndPlanControlPanel.Name = "tabTndPlanControlPanel"
        '
        'GRPProgramTDToolbar
        '
        Me.GRPProgramTDToolbar.Items.Add(Me.btnLoadOpenTnDPlan)
        Me.GRPProgramTDToolbar.Items.Add(Me.Separator1)
        Me.GRPProgramTDToolbar.Items.Add(Me.btnConvertToSpecific)
        Me.GRPProgramTDToolbar.Items.Add(Me.menuDraft)
        Me.GRPProgramTDToolbar.Items.Add(Me.menuCheckInOut)
        Me.GRPProgramTDToolbar.Items.Add(Me.Separator2)
        Me.GRPProgramTDToolbar.Items.Add(Me.btnRefreshPlan)
        Me.GRPProgramTDToolbar.Items.Add(Me.btnRefreshUnit)
        Me.GRPProgramTDToolbar.Items.Add(Me.Separator4)
        Me.GRPProgramTDToolbar.Items.Add(Me.btnUndo)
        Me.GRPProgramTDToolbar.Items.Add(Me.btnRedo)
        Me.GRPProgramTDToolbar.Label = "Program T&&D Toolbar"
        Me.GRPProgramTDToolbar.Name = "GRPProgramTDToolbar"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'Separator4
        '
        Me.Separator4.Name = "Separator4"
        '
        'GRPSearchFilterHighlight
        '
        Me.GRPSearchFilterHighlight.Items.Add(Me.btnSearchFilter)
        Me.GRPSearchFilterHighlight.Items.Add(Me.btnClearFilter)
        Me.GRPSearchFilterHighlight.Label = "Search, Filter && Highlight"
        Me.GRPSearchFilterHighlight.Name = "GRPSearchFilterHighlight"
        '
        'GRPAddDeleteUpdateVehicle
        '
        Me.GRPAddDeleteUpdateVehicle.Items.Add(Me.btnAddUnit)
        Me.GRPAddDeleteUpdateVehicle.Items.Add(Me.btnDeleteUnit)
        Me.GRPAddDeleteUpdateVehicle.Items.Add(Me.btnChangeSequence)
        Me.GRPAddDeleteUpdateVehicle.Label = "Units"
        Me.GRPAddDeleteUpdateVehicle.Name = "GRPAddDeleteUpdateVehicle"
        '
        'GRPPlan
        '
        Me.GRPPlan.Items.Add(Me.btnUpdateMRD)
        Me.GRPPlan.Label = "Plan Gateways"
        Me.GRPPlan.Name = "GRPPlan"
        '
        'GRPShowHideSpecificationsections
        '
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togInstrumentation)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togNonMFCSpecification)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togMfcSpecification)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togProgramInformation)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togFurtherBasicSpecification)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togUserShipping)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togUpdatePack)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togTiming)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.togShowAll)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.Separator7)
        Me.GRPShowHideSpecificationsections.Items.Add(Me.btnUpdateColumns)
        Me.GRPShowHideSpecificationsections.Label = "Show/Hide Specification sections"
        Me.GRPShowHideSpecificationsections.Name = "GRPShowHideSpecificationsections"
        '
        'Separator7
        '
        Me.Separator7.Name = "Separator7"
        '
        'GRPIndicator
        '
        Me.GRPIndicator.Items.Add(Me.btnTodayIndicator)
        Me.GRPIndicator.Label = "Indicator"
        Me.GRPIndicator.Name = "GRPIndicator"
        '
        'GrpMessages
        '
        Me.GrpMessages.Items.Add(Me.TGMessages)
        Me.GrpMessages.Label = "Messages"
        Me.GrpMessages.Name = "GrpMessages"
        '
        'GRPHelp
        '
        Me.GRPHelp.Items.Add(Me.btnCTHelp)
        Me.GRPHelp.Label = "Help"
        Me.GRPHelp.Name = "GRPHelp"
        '
        'tabTndPlanReports
        '
        Me.tabTndPlanReports.Groups.Add(Me.GRPReport)
        Me.tabTndPlanReports.Groups.Add(Me.GRPHolidayMaster)
        Me.tabTndPlanReports.Label = "CT Plan Report && Setting"
        Me.tabTndPlanReports.Name = "tabTndPlanReports"
        '
        'GRPReport
        '
        Me.GRPReport.Items.Add(Me.btnExportToExcel)
        Me.GRPReport.Items.Add(Me.btnUnitReport)
        Me.GRPReport.Items.Add(Me.btnEngineTransmissionReport)
        Me.GRPReport.Items.Add(Me.btnPrecheckF4T)
        Me.GRPReport.Items.Add(Me.btnCountReport)
        Me.GRPReport.Items.Add(Me.btnPustFit4Test)
        Me.GRPReport.Items.Add(Me.Separator3)
        Me.GRPReport.Items.Add(Me.tglBtnValidatePlan)
        Me.GRPReport.Label = "Report"
        Me.GRPReport.Name = "GRPReport"
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'GRPHolidayMaster
        '
        Me.GRPHolidayMaster.Items.Add(Me.btnUpdateHoliday)
        Me.GRPHolidayMaster.Items.Add(Me.btnCDSIDtoDvpTeam)
        Me.GRPHolidayMaster.Label = "Plan Setting"
        Me.GRPHolidayMaster.Name = "GRPHolidayMaster"
        '
        'tmrDisplay
        '
        Me.tmrDisplay.Interval = 20000
        '
        'tmrMessages
        '
        Me.tmrMessages.Enabled = True
        Me.tmrMessages.Interval = 60000
        '
        'btnLoadOpenTnDPlan
        '
        Me.btnLoadOpenTnDPlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnLoadOpenTnDPlan.Label = "Load && Open TnD Plan"
        Me.btnLoadOpenTnDPlan.Name = "btnLoadOpenTnDPlan"
        Me.btnLoadOpenTnDPlan.OfficeImageId = "ShowSchedulingPage"
        Me.btnLoadOpenTnDPlan.ShowImage = True
        '
        'btnConvertToSpecific
        '
        Me.btnConvertToSpecific.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnConvertToSpecific.Enabled = False
        Me.btnConvertToSpecific.Label = "Convert To Specific"
        Me.btnConvertToSpecific.Name = "btnConvertToSpecific"
        Me.btnConvertToSpecific.OfficeImageId = "PublishToPdfOrEdoc"
        Me.btnConvertToSpecific.ShowImage = True
        '
        'menuDraft
        '
        Me.menuDraft.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.menuDraft.Enabled = False
        Me.menuDraft.Items.Add(Me.btnGenerateDraft)
        Me.menuDraft.Items.Add(Me.mnuGenerateDraft)
        Me.menuDraft.Items.Add(Me.btnDeleteDraft)
        Me.menuDraft.Items.Add(Me.btnReplacePlanWithDraft)
        Me.menuDraft.Label = "Draft"
        Me.menuDraft.Name = "menuDraft"
        Me.menuDraft.OfficeImageId = "DefinePrintStyles"
        Me.menuDraft.ShowImage = True
        '
        'btnGenerateDraft
        '
        Me.btnGenerateDraft.Enabled = False
        Me.btnGenerateDraft.Label = "Generate Draft"
        Me.btnGenerateDraft.Name = "btnGenerateDraft"
        Me.btnGenerateDraft.OfficeImageId = "IndexInsert"
        Me.btnGenerateDraft.ShowImage = True
        '
        'mnuGenerateDraft
        '
        Me.mnuGenerateDraft.Dynamic = True
        Me.mnuGenerateDraft.Enabled = False
        Me.mnuGenerateDraft.Label = "Load Draft"
        Me.mnuGenerateDraft.Name = "mnuGenerateDraft"
        Me.mnuGenerateDraft.OfficeImageId = "GroupBlogPublish"
        Me.mnuGenerateDraft.ShowImage = True
        '
        'btnDeleteDraft
        '
        Me.btnDeleteDraft.Enabled = False
        Me.btnDeleteDraft.Label = "Delete Draft"
        Me.btnDeleteDraft.Name = "btnDeleteDraft"
        Me.btnDeleteDraft.OfficeImageId = "HeaderFooterRemoveFooterWord"
        Me.btnDeleteDraft.ShowImage = True
        '
        'btnReplacePlanWithDraft
        '
        Me.btnReplacePlanWithDraft.Enabled = False
        Me.btnReplacePlanWithDraft.Label = "Replace plan with draft"
        Me.btnReplacePlanWithDraft.Name = "btnReplacePlanWithDraft"
        Me.btnReplacePlanWithDraft.OfficeImageId = "InsertDialog"
        Me.btnReplacePlanWithDraft.ShowImage = True
        '
        'menuCheckInOut
        '
        Me.menuCheckInOut.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.menuCheckInOut.Enabled = False
        Me.menuCheckInOut.Items.Add(Me.btnCheckOut)
        Me.menuCheckInOut.Items.Add(Me.btnCheckIn)
        Me.menuCheckInOut.Items.Add(Me.btnDiscard)
        Me.menuCheckInOut.Items.Add(Me.mnuActiveUsers)
        Me.menuCheckInOut.Label = "Check-Out/-In"
        Me.menuCheckInOut.Name = "menuCheckInOut"
        Me.menuCheckInOut.OfficeImageId = "ReviewTrackChanges"
        Me.menuCheckInOut.ShowImage = True
        '
        'btnCheckOut
        '
        Me.btnCheckOut.Enabled = False
        Me.btnCheckOut.Label = "Check-Out"
        Me.btnCheckOut.Name = "btnCheckOut"
        Me.btnCheckOut.OfficeImageId = "FileCheckOut"
        Me.btnCheckOut.ShowImage = True
        Me.btnCheckOut.Visible = False
        '
        'btnCheckIn
        '
        Me.btnCheckIn.Enabled = False
        Me.btnCheckIn.Label = "Check-In"
        Me.btnCheckIn.Name = "btnCheckIn"
        Me.btnCheckIn.OfficeImageId = "FileCheckIn"
        Me.btnCheckIn.ShowImage = True
        Me.btnCheckIn.Visible = False
        '
        'btnDiscard
        '
        Me.btnDiscard.Enabled = False
        Me.btnDiscard.Label = "Discard && Close"
        Me.btnDiscard.Name = "btnDiscard"
        Me.btnDiscard.OfficeImageId = "FileCheckOutDiscard"
        Me.btnDiscard.ShowImage = True
        Me.btnDiscard.Visible = False
        '
        'mnuActiveUsers
        '
        Me.mnuActiveUsers.Dynamic = True
        Me.mnuActiveUsers.Enabled = False
        Me.mnuActiveUsers.Label = "Active users"
        Me.mnuActiveUsers.Name = "mnuActiveUsers"
        Me.mnuActiveUsers.OfficeImageId = "AccessListContacts"
        Me.mnuActiveUsers.ShowImage = True
        '
        'btnRefreshPlan
        '
        Me.btnRefreshPlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnRefreshPlan.Enabled = False
        Me.btnRefreshPlan.Label = "Refresh T&&D Plan"
        Me.btnRefreshPlan.Name = "btnRefreshPlan"
        Me.btnRefreshPlan.OfficeImageId = "DataRefreshAll"
        Me.btnRefreshPlan.ShowImage = True
        '
        'btnRefreshUnit
        '
        Me.btnRefreshUnit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnRefreshUnit.Enabled = False
        Me.btnRefreshUnit.Label = "Refresh Unit"
        Me.btnRefreshUnit.Name = "btnRefreshUnit"
        Me.btnRefreshUnit.OfficeImageId = "Refresh"
        Me.btnRefreshUnit.ShowImage = True
        '
        'btnUndo
        '
        Me.btnUndo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUndo.Enabled = False
        Me.btnUndo.Label = "Undo"
        Me.btnUndo.Name = "btnUndo"
        Me.btnUndo.OfficeImageId = "Undo"
        Me.btnUndo.ShowImage = True
        '
        'btnRedo
        '
        Me.btnRedo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnRedo.Enabled = False
        Me.btnRedo.Label = "Redo"
        Me.btnRedo.Name = "btnRedo"
        Me.btnRedo.OfficeImageId = "Redo"
        Me.btnRedo.ShowImage = True
        '
        'btnSearchFilter
        '
        Me.btnSearchFilter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSearchFilter.Enabled = False
        Me.btnSearchFilter.Label = "Search, Filter && Highlight"
        Me.btnSearchFilter.Name = "btnSearchFilter"
        Me.btnSearchFilter.OfficeImageId = "ZoomToSelection"
        Me.btnSearchFilter.ShowImage = True
        '
        'btnClearFilter
        '
        Me.btnClearFilter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnClearFilter.Enabled = False
        Me.btnClearFilter.Label = "Clear Filter"
        Me.btnClearFilter.Name = "btnClearFilter"
        Me.btnClearFilter.OfficeImageId = "TableStyleClear"
        Me.btnClearFilter.ShowImage = True
        '
        'btnAddUnit
        '
        Me.btnAddUnit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAddUnit.Enabled = False
        Me.btnAddUnit.Label = "Add Unit"
        Me.btnAddUnit.Name = "btnAddUnit"
        Me.btnAddUnit.OfficeImageId = "CellsInsertDialog"
        Me.btnAddUnit.ShowImage = True
        '
        'btnDeleteUnit
        '
        Me.btnDeleteUnit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDeleteUnit.Enabled = False
        Me.btnDeleteUnit.Label = "Delete Unit"
        Me.btnDeleteUnit.Name = "btnDeleteUnit"
        Me.btnDeleteUnit.OfficeImageId = "CellsDelete"
        Me.btnDeleteUnit.ShowImage = True
        '
        'btnChangeSequence
        '
        Me.btnChangeSequence.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnChangeSequence.Enabled = False
        Me.btnChangeSequence.Label = "Change unit Sequence"
        Me.btnChangeSequence.Name = "btnChangeSequence"
        Me.btnChangeSequence.OfficeImageId = "TableRowsInsertBelowWord"
        Me.btnChangeSequence.ShowImage = True
        '
        'btnUpdateMRD
        '
        Me.btnUpdateMRD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUpdateMRD.Enabled = False
        Me.btnUpdateMRD.Label = "Update MRD && Dates"
        Me.btnUpdateMRD.Name = "btnUpdateMRD"
        Me.btnUpdateMRD.OfficeImageId = "ControlLayoutStacked"
        Me.btnUpdateMRD.ShowImage = True
        '
        'togInstrumentation
        '
        Me.togInstrumentation.Enabled = False
        Me.togInstrumentation.Label = "Instrumentation"
        Me.togInstrumentation.Name = "togInstrumentation"
        '
        'togNonMFCSpecification
        '
        Me.togNonMFCSpecification.Enabled = False
        Me.togNonMFCSpecification.Label = "Non MFC Specification"
        Me.togNonMFCSpecification.Name = "togNonMFCSpecification"
        '
        'togMfcSpecification
        '
        Me.togMfcSpecification.Enabled = False
        Me.togMfcSpecification.Label = "MFC Specification"
        Me.togMfcSpecification.Name = "togMfcSpecification"
        '
        'togProgramInformation
        '
        Me.togProgramInformation.Enabled = False
        Me.togProgramInformation.Label = "Program Information"
        Me.togProgramInformation.Name = "togProgramInformation"
        '
        'togFurtherBasicSpecification
        '
        Me.togFurtherBasicSpecification.Enabled = False
        Me.togFurtherBasicSpecification.Label = "Further Basic Specification"
        Me.togFurtherBasicSpecification.Name = "togFurtherBasicSpecification"
        '
        'togUserShipping
        '
        Me.togUserShipping.Enabled = False
        Me.togUserShipping.Label = "User && Shipping Details"
        Me.togUserShipping.Name = "togUserShipping"
        '
        'togUpdatePack
        '
        Me.togUpdatePack.Enabled = False
        Me.togUpdatePack.Label = "Update Pack"
        Me.togUpdatePack.Name = "togUpdatePack"
        '
        'togTiming
        '
        Me.togTiming.Enabled = False
        Me.togTiming.Label = "Timing"
        Me.togTiming.Name = "togTiming"
        '
        'togShowAll
        '
        Me.togShowAll.Enabled = False
        Me.togShowAll.Label = "Show All"
        Me.togShowAll.Name = "togShowAll"
        '
        'btnUpdateColumns
        '
        Me.btnUpdateColumns.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUpdateColumns.Enabled = False
        Me.btnUpdateColumns.Label = "Columns Add/Update/Delete"
        Me.btnUpdateColumns.Name = "btnUpdateColumns"
        Me.btnUpdateColumns.OfficeImageId = "DatasheetColumnLookup"
        Me.btnUpdateColumns.ShowImage = True
        '
        'btnTodayIndicator
        '
        Me.btnTodayIndicator.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnTodayIndicator.Label = "Show/hide today indicator"
        Me.btnTodayIndicator.Name = "btnTodayIndicator"
        Me.btnTodayIndicator.OfficeImageId = "NewAppointment"
        Me.btnTodayIndicator.ShowImage = True
        '
        'TGMessages
        '
        Me.TGMessages.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TGMessages.Label = "Show/Hide messages"
        Me.TGMessages.Name = "TGMessages"
        Me.TGMessages.OfficeImageId = "NewMailMessage"
        Me.TGMessages.ShowImage = True
        '
        'btnCTHelp
        '
        Me.btnCTHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnCTHelp.Label = "CT Help"
        Me.btnCTHelp.Name = "btnCTHelp"
        Me.btnCTHelp.OfficeImageId = "Help"
        Me.btnCTHelp.ShowImage = True
        '
        'btnExportToExcel
        '
        Me.btnExportToExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnExportToExcel.Enabled = False
        Me.btnExportToExcel.Label = "Export To Excel"
        Me.btnExportToExcel.Name = "btnExportToExcel"
        Me.btnExportToExcel.OfficeImageId = "ExportExcel"
        Me.btnExportToExcel.ShowImage = True
        '
        'btnUnitReport
        '
        Me.btnUnitReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUnitReport.Enabled = False
        Me.btnUnitReport.Label = "Unit Report"
        Me.btnUnitReport.Name = "btnUnitReport"
        Me.btnUnitReport.OfficeImageId = "ShowSchedulingPage"
        Me.btnUnitReport.ShowImage = True
        '
        'btnEngineTransmissionReport
        '
        Me.btnEngineTransmissionReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnEngineTransmissionReport.Enabled = False
        Me.btnEngineTransmissionReport.Label = "Engine && Transmission Report"
        Me.btnEngineTransmissionReport.Name = "btnEngineTransmissionReport"
        Me.btnEngineTransmissionReport.OfficeImageId = "TableSharePointListsModifyColumnsAndSettings"
        Me.btnEngineTransmissionReport.ShowImage = True
        '
        'btnPrecheckF4T
        '
        Me.btnPrecheckF4T.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnPrecheckF4T.Enabled = False
        Me.btnPrecheckF4T.Label = "F4Test Precheck report"
        Me.btnPrecheckF4T.Name = "btnPrecheckF4T"
        Me.btnPrecheckF4T.OfficeImageId = "AccessTableTasks"
        Me.btnPrecheckF4T.ShowImage = True
        '
        'btnCountReport
        '
        Me.btnCountReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnCountReport.Enabled = False
        Me.btnCountReport.Label = "Total testdays report"
        Me.btnCountReport.Name = "btnCountReport"
        Me.btnCountReport.OfficeImageId = "ComAddInsDialog"
        Me.btnCountReport.ShowImage = True
        '
        'btnPustFit4Test
        '
        Me.btnPustFit4Test.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnPustFit4Test.Label = "Push to Fit4Test"
        Me.btnPustFit4Test.Name = "btnPustFit4Test"
        Me.btnPustFit4Test.OfficeImageId = "ChartResetToMatchStyle"
        Me.btnPustFit4Test.ShowImage = True
        '
        'tglBtnValidatePlan
        '
        Me.tglBtnValidatePlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tglBtnValidatePlan.Enabled = False
        Me.tglBtnValidatePlan.Label = "Validate Plan"
        Me.tglBtnValidatePlan.Name = "tglBtnValidatePlan"
        Me.tglBtnValidatePlan.OfficeImageId = "ReviewAcceptChange"
        Me.tglBtnValidatePlan.ShowImage = True
        '
        'btnUpdateHoliday
        '
        Me.btnUpdateHoliday.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUpdateHoliday.Enabled = False
        Me.btnUpdateHoliday.Label = "Holidays"
        Me.btnUpdateHoliday.Name = "btnUpdateHoliday"
        Me.btnUpdateHoliday.OfficeImageId = "CopyToPersonalCalendar"
        Me.btnUpdateHoliday.ShowImage = True
        '
        'btnCDSIDtoDvpTeam
        '
        Me.btnCDSIDtoDvpTeam.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnCDSIDtoDvpTeam.Enabled = False
        Me.btnCDSIDtoDvpTeam.Label = "Assign CDS to DVP"
        Me.btnCDSIDtoDvpTeam.Name = "btnCDSIDtoDvpTeam"
        Me.btnCDSIDtoDvpTeam.OfficeImageId = "MailMergeRecipientsEditList"
        Me.btnCDSIDtoDvpTeam.ShowImage = True
        '
        'RbnTnDControlPanel
        '
        Me.Name = "RbnTnDControlPanel"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.tabTndPlanControlPanel)
        Me.Tabs.Add(Me.tabTndPlanReports)
        Me.tabTndPlanControlPanel.ResumeLayout(False)
        Me.tabTndPlanControlPanel.PerformLayout()
        Me.GRPProgramTDToolbar.ResumeLayout(False)
        Me.GRPProgramTDToolbar.PerformLayout()
        Me.GRPSearchFilterHighlight.ResumeLayout(False)
        Me.GRPSearchFilterHighlight.PerformLayout()
        Me.GRPAddDeleteUpdateVehicle.ResumeLayout(False)
        Me.GRPAddDeleteUpdateVehicle.PerformLayout()
        Me.GRPPlan.ResumeLayout(False)
        Me.GRPPlan.PerformLayout()
        Me.GRPShowHideSpecificationsections.ResumeLayout(False)
        Me.GRPShowHideSpecificationsections.PerformLayout()
        Me.GRPIndicator.ResumeLayout(False)
        Me.GRPIndicator.PerformLayout()
        Me.GrpMessages.ResumeLayout(False)
        Me.GrpMessages.PerformLayout()
        Me.GRPHelp.ResumeLayout(False)
        Me.GRPHelp.PerformLayout()
        Me.tabTndPlanReports.ResumeLayout(False)
        Me.tabTndPlanReports.PerformLayout()
        Me.GRPReport.ResumeLayout(False)
        Me.GRPReport.PerformLayout()
        Me.GRPHolidayMaster.ResumeLayout(False)
        Me.GRPHolidayMaster.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GRPProgramTDToolbar As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnLoadOpenTnDPlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GRPShowHideSpecificationsections As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnUpdateColumns As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GRPSearchFilterHighlight As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnRefreshPlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRefreshUnit As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator4 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnUndo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRedo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GRPAddDeleteUpdateVehicle As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnAddUnit As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeleteUnit As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnChangeSequence As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GRPPlan As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents togInstrumentation As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togNonMFCSpecification As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togMfcSpecification As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togProgramInformation As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togFurtherBasicSpecification As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togUserShipping As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togUpdatePack As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togTiming As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents togShowAll As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GRPReport As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnExportToExcel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUnitReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnEngineTransmissionReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUpdateMRD As Microsoft.Office.Tools.Ribbon.RibbonButton
    Public WithEvents btnConvertToSpecific As Microsoft.Office.Tools.Ribbon.RibbonButton
    Public WithEvents tabTndPlanControlPanel As Microsoft.Office.Tools.Ribbon.RibbonTab
    Public WithEvents btnSearchFilter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Public WithEvents tabTndPlanReports As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GRPIndicator As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents mnuGenerateDraft As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Public WithEvents btnClearFilter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnPrecheckF4T As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator7 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents menuCheckInOut As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnCheckIn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDiscard As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCheckOut As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents menuDraft As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnGenerateDraft As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeleteDraft As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnReplacePlanWithDraft As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCountReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnPustFit4Test As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tmrDisplay As System.Windows.Forms.Timer
    Friend WithEvents TGMessages As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents tmrMessages As System.Windows.Forms.Timer
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnTodayIndicator As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents tglBtnValidatePlan As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GRPHolidayMaster As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnUpdateHoliday As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCDSIDtoDvpTeam As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GrpMessages As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GRPHelp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCTHelp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnuActiveUsers As Microsoft.Office.Tools.Ribbon.RibbonMenu
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property RbnTnDControlPanel() As RbnTnDControlPanel
        Get

            Return Me.GetRibbon(Of RbnTnDControlPanel)()
        End Get
    End Property
End Class
