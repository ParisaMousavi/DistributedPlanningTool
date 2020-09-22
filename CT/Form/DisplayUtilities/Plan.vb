
Imports System.Data
Imports System.Windows.Forms

Namespace Form.DisplayUtilities

    Public Class Plan

        Dim _frmProgress As frmProgressbar = Nothing
        Dim _frmHCIDSelect As frmHCIDSelect


        Public Enum LoadType
            Refreshing
            Loading
        End Enum


        Private _ErrorMessage As String
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property


        Event EventUpdateProgress(progressvalue As Double)
        'Public intPer As Double = 0
        Private WithEvents _DrawTndPlanArea As Form.DisplayUtilities.DrawTndPlanArea = New Form.DisplayUtilities.DrawTndPlanArea()

        Private WithEvents _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader()

        Private Sub UpdateProgressbar(progressvalue As Double)
            RaiseEvent EventUpdateProgress(progressvalue)
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            If Not _frmProgress Is Nothing Then
                _frmProgress.UpdateProgressBar(progressvalue)
                If _frmProgress.Tag = LoadType.Loading.ToString Then
                    _frmProgress.Text = "Loading plan : " & CInt(_frmProgress.SmoothProgressBar2.Value) & "% completed."
                Else
                    _frmProgress.Text = "Refreshing plan : " & CInt(_frmProgress.SmoothProgressBar2.Value) & "% completed."
                End If
                _frmProgress.Refresh()
            ElseIf _frmHCIDSelect IsNot Nothing Then
                _frmHCIDSelect.UpdateProgressBar(progressvalue)
            End If
        End Sub

        Public Sub UpdateProgress(intprogressvalue As Double) Handles _DrawTndPlanArea.EventUpdateProgress
            UpdateProgressbar(intprogressvalue)
        End Sub

        Public Sub UpdateProgress_Hdr(intprogressvalue As Double) Handles _DrawTndPlanHeader.EventUpdateProgress
            UpdateProgressbar(intprogressvalue)
        End Sub


        Public Sub New(Optional ByRef frmHCIDSelect As frmHCIDSelect = Nothing)
            _frmHCIDSelect = frmHCIDSelect
        End Sub

        ' Public Function LoadSpecificPlan(You parameter) As String
        Public Function LoadSpecificPlan(WithCustomFormat As Boolean) As String 'jeeva
            LoadSpecificPlan = String.Empty
            Dim _PlanActiveUsers As CT.Data.PlanActiveUsers = New Data.PlanActiveUsers
            Try
                Dim strMessage As String = String.Empty

                'Loading Plan Header data
                ' Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader()
                Dim TotalColumn As Integer = 0
                strMessage = _DrawTndPlanHeader.LoadTndPlanHeaderToWorkSheet(TotalColumn)
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                'intPer = 5 '10
                UpdateProgressbar(5)
                '
                Form.DisplayUtilities.TndSection.DetectFirstElementarySections(TotalColumn)

                'intPer = 5 '15
                UpdateProgressbar(5)
                ' 

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                'Loading Plan Information data
                Dim _DrawTndPlanInformation As Form.DisplayUtilities.DrawTndPlanInformation = New Form.DisplayUtilities.DrawTndPlanInformation()
                _DrawTndPlanInformation.LoadTndPlanInformationToWorkSheet(Nothing, Nothing)

                'intPer = 15 '25
                UpdateProgressbar(10)

                'Loading Plan Area data
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                strMessage = _DrawTndPlanArea.LoadTndPlanAreaToWorkSheet(Nothing, Nothing) '5 progress
                If strMessage <> String.Empty Then Throw New Exception(strMessage)


                'intPer = 10 '40
                'UpdateProgressbar(10)

                'Applying formatting
                ' _DrawTndPlanHeader.ApplyFormattingAfterLoading(TotalColumn, Your parameter)
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                If _DrawTndPlanHeader.ApplyFormattingAfterLoading(TotalColumn, WithCustomFormat) = False Then Throw New Exception(_DrawTndPlanHeader.ErrorMessage) 'Jeeva'15 progress
                'Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                'intPer = 5 '45
                'UpdateProgressbar(5)

                strMessage = _DrawTndPlanInformation.ApplyFormattingAfterLoading()
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                'intPer = 5 '50
                UpdateProgressbar(5)

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                strMessage = _DrawTndPlanArea.ApplyFormattingAfterLoading(Nothing, Nothing) '                 'UpdateProgressbar(25) 75
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                'intPer = 15 '90
                UpdateProgressbar(15)
                '  
                'System.Threading.Thread.Sleep(50)
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Dim _TndPlanTitle As Form.DisplayUtilities.TndPlanTitle = New Form.DisplayUtilities.TndPlanTitle
                strMessage = _TndPlanTitle.LoadAndFormatLabel()
                If strMessage <> String.Empty Then Throw New Exception(strMessage)
                strMessage = _TndPlanTitle.FillMismatchedQty()
                If strMessage <> String.Empty Then Throw New Exception(strMessage)



                '-----------------------------------------------------------------
                ' Insert the current CDSID to PlanActiveUser to control the check-in and checkout plas
                '-----------------------------------------------------------------
                If CT.Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString() Then
                    If _PlanActiveUsers.Insert(CT.Form.DataCenter.ProgramConfig.pe01, CT.Form.DataCenter.ProgramConfig.HCID, CT.Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                End If


                'intPer = 10 '100
                UpdateProgressbar(10)
                '
            Catch ex As Exception
                LoadSpecificPlan = "LoadSpecificPlan: " + ex.Message
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="XCCPe26"></param>
        ''' <param name="XCCPe01"></param>
        ''' <param name="AssyBuildScale"></param>
        ''' <param name="WithCustomFormat"> must be always false because the related table for Custom formatting doesn't existed for generic plans. After 
        ''' converting the related values are created in tables.</param>
        ''' <returns></returns>
        Public Function LoadGenericPlan(XCCPe26 As Long, XCCPe01 As Long, AssyBuildScale As Integer, WithCustomFormat As Boolean) As String
            LoadGenericPlan = String.Empty
            Try
                Dim strMessage As String = String.Empty



                '-----------------------------------------------------------------------
                'Before loading plan the generic plan must be created
                '-----------------------------------------------------------------------

                Dim _PlanInterface As Data.Interfaces.PlanInterface

                If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _PlanInterface = New Data.VehiclePlan.Plan
                ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _PlanInterface = New Data.RigPlan.Plan
                Else
                    Exit Try
                End If

                If _PlanInterface.GenerateGenericPlan(Form.DataCenter.ProgramConfig.pe01,
                                             Form.DataCenter.ProgramConfig.pe02,
                                             Form.DataCenter.ProgramConfig.HCID,
                                             XCCPe26,
                                             XCCPe01,
                                             AssyBuildScale,
                                             Form.DataCenter.ProgramConfig.BuildPhase,
                                             Form.DataCenter.ProgramConfig.BuildType,
                                             DirectCast([Enum].Parse(GetType(CT.Data.DataCenter.FileStatus), Form.DataCenter.ProgramConfig.FileStatus), CT.Data.DataCenter.FileStatus)) = False Then
                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                End If

                '-----------------------------------------------------------------------
                ' Because generic plan get user access level after generating
                '-----------------------------------------------------------------------
                Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
                Dim _strUserPermissionLevel As String = String.Empty
                Try
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel = Nothing Then
                        '--------------------------------------------------------------------------
                        ' validation for controlling the result of DAL
                        '--------------------------------------------------------------------------
                        _strUserPermissionLevel = objPer.GetPermissionLevel(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.IsGeneric)
                        If _strUserPermissionLevel Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        Form.DataCenter.GlobalValues.strUserPermissionLevel = _strUserPermissionLevel
                    End If
                Catch ex As Exception
                    Form.DataCenter.GlobalValues.strUserPermissionLevel = String.Empty
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Load generic plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Finally
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel = "" Then
                        Throw New Exception("Access denied! Please contact 'AEREN8@FORD.COM' or OMEIGEN@FORD.COM or PNEZHAD@FORD.COM OR MAGES@FORD.COM for permission!")
                    End If

                End Try


                'Loading Plan Header data
                Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader()
                Dim TotalColumn As Integer = 0
                strMessage = _DrawTndPlanHeader.LoadTndPlanHeaderToWorkSheet(TotalColumn)
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                UpdateProgressbar(5)

                Form.DisplayUtilities.TndSection.DetectFirstElementarySections(TotalColumn)

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                UpdateProgressbar(5)

                'Loading Plan Information data
                Dim _DrawTndPlanInformation As Form.DisplayUtilities.DrawTndPlanInformation = New Form.DisplayUtilities.DrawTndPlanInformation()
                strMessage = _DrawTndPlanInformation.LoadTndPlanInformationToWorkSheet(Nothing, Nothing)
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                'Globals.ThisAddIn.Application.ScreenUpdating = False
                'Globals.ThisAddIn.Application.EnableEvents = False
                'Globals.ThisAddIn.Application.DisplayAlerts = False

                UpdateProgressbar(15)

                'Loading Plan Area data

                'Dim _DrawTndPlanArea As Form.DisplayUtilities.DrawTndPlanArea = New Form.DisplayUtilities.DrawTndPlanArea()
                strMessage = _DrawTndPlanArea.LoadTndPlanAreaToWorkSheet(Nothing, Nothing)
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                UpdateProgressbar(10)

                'Globals.ThisAddIn.Application.ScreenUpdating = False
                'Globals.ThisAddIn.Application.EnableEvents = False
                'Globals.ThisAddIn.Application.DisplayAlerts = False


                'Applying formatting
                If _DrawTndPlanHeader.ApplyFormattingAfterLoading(TotalColumn, WithCustomFormat) = False Then Throw New Exception(_DrawTndPlanHeader.ErrorMessage)
                'Globals.ThisAddIn.Application.ScreenUpdating = False
                'Globals.ThisAddIn.Application.EnableEvents = False
                'Globals.ThisAddIn.Application.DisplayAlerts = False

                'intPer = 10 '50
                UpdateProgressbar(10)

                _DrawTndPlanInformation.ApplyFormattingAfterLoading()

                'intPer = 25 '75
                'UpdateProgressbar(25)

                'Globals.ThisAddIn.Application.ScreenUpdating = False
                'Globals.ThisAddIn.Application.EnableEvents = False
                'Globals.ThisAddIn.Application.DisplayAlerts = False
                _DrawTndPlanArea.ApplyFormattingAfterLoading(Nothing, Nothing)

                'intPer = 15 '90
                UpdateProgressbar(15)

                Dim _TndPlanTitle As Form.DisplayUtilities.TndPlanTitle = New Form.DisplayUtilities.TndPlanTitle
                _TndPlanTitle.LoadAndFormatLabel()
                _TndPlanTitle.FillMismatchedQty()

                'intPer = 10 '100
                UpdateProgressbar(10)

            Catch ex As Exception
                LoadGenericPlan = ex.Message
            End Try
        End Function


        Public Function LoadDraftPlan(DraftHCID As Integer, WithCustomFormat As Object, strLoadtype As LoadType, BuildType As String) As String
            Dim WB As Excel.Workbook
            Dim _obj As New Form.DataCenter.ModuleFunction

            '------------------------------------------------------------------------------------
            ' Reading the path of template from registry 192 100192 101192 102192
            '------------------------------------------------------------------------------------
            Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\CT", "Manifest", Nothing)
            readValue = readValue.ToString().Substring(0, readValue.ToString().LastIndexOf("/") + 1) + "TndTemplate/TndTemplate.xltm"
            Globals.ThisAddIn.Application.Workbooks.Add(readValue)
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False 'True

            Try

                If WithCustomFormat Is Nothing Then
                    Dim Answer As DialogResult
                    Answer = MessageBox.Show("Do you want to load Draft Plan with Custom format? If you want to can interrupt loading with Cancel button.", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If Answer = DialogResult.No Then
                        WithCustomFormat = False
                    ElseIf Answer = DialogResult.Cancel Then
                        '--------------------------------------------------------------------------------
                        ' Close currect not valid Draft
                        '--------------------------------------------------------------------------------
                        Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)
                        WB = Globals.ThisAddIn.Application.Workbooks(Globals.ThisAddIn.Application.Workbooks.Count)
                        Form.DataCenter.GlobalValues.objWBCurrent = WB
                        Throw New Exception("Loading is interrupted by user!")
                    Else
                        WithCustomFormat = True
                    End If
                End If


                WB = Globals.ThisAddIn.Application.Workbooks(Globals.ThisAddIn.Application.Workbooks.Count)
                If WB.Name Like "TndTemplate*" Then
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    WB.Application.ScreenUpdating = False
                    WB.Activate()
                    WB.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()
                    Form.DataCenter.GlobalValues.objWBCurrent = WB
                End If

                Try
                    CT.Form.DataCenter.GlobalValues.wsEve = New Form.DisplayUtilities.clsWorksheetEvents
                Catch ex As Exception

                End Try

                Form.DataCenter.GlobalValues.bolPlanIsLoading = True

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

                Form.DataCenter.GlobalValues.Clear()

                If _frmProgress Is Nothing Then
                    _frmProgress = New frmProgressbar()
                    _frmProgress.Tag = strLoadtype.ToString
                    _frmProgress.Show()
                Else
                    _frmProgress.Show()

                End If

                'Form.DisplayUtilities.TndSection.CleanRange(Form.DataCenter.GlobalValues.WS.UsedRange)

                'select a plan dedicated to get all the plans which are containede in selected plan
                'We have a main plan and each main plan contains some other plan
                'Other plans have different BuildType
                'Dim _Plan As Data.Plan = New Data.Plan

                Dim dtDraft As System.Data.DataTable

                Dim _PlanInterface As Data.Interfaces.PlanInterface

                If BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _PlanInterface = New Data.VehiclePlan.Plan
                ElseIf BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _PlanInterface = New Data.RigPlan.Plan
                Else
                    LoadDraftPlan = "BuildType cannot be blank"
                    Exit Try
                End If

                dtDraft = _PlanInterface.SelectTndDraftPlanDedicated(DraftHCID, BuildType)

                UpdateProgressbar(5)

                Form.DataCenter.ProgramConfig.pe01 = dtDraft.Rows(0).Item("Pe01")
                Form.DataCenter.ProgramConfig.HCID = dtDraft.Rows(0).Item("HealthChartID")
                Form.DataCenter.ProgramConfig.IsGeneric = False
                Form.DataCenter.ProgramConfig.pe02 = dtDraft.Rows(0).Item("pe02")
                Form.DataCenter.ProgramConfig.XccPe26 = dtDraft.Rows(0).Item("XCCpe26")
                Form.DataCenter.ProgramConfig.XccPe01 = dtDraft.Rows(0).Item("XCCpe01")
                If IsDBNull(dtDraft.Rows(0).Item("AssyBuildScale")) = False Then
                    Form.DataCenter.ProgramConfig.AssyBuildScale = dtDraft.Rows(0).Item("AssyBuildScale")
                Else
                    Form.DataCenter.ProgramConfig.AssyBuildScale = 0
                End If
                Form.DataCenter.ProgramConfig.BuildType = dtDraft.Rows(0).Item("BuildType")
                Form.DataCenter.ProgramConfig.BuildPhase = dtDraft.Rows(0).Item("BuildPhase")
                Form.DataCenter.ProgramConfig.Carline = dtDraft.Rows(0).Item("Carline")
                Form.DataCenter.ProgramConfig.Platform = dtDraft.Rows(0).Item("Platform")
                Form.DataCenter.ProgramConfig.HCIDName = dtDraft.Rows(0).Item("HealthChartName")
                Form.DataCenter.ProgramConfig.IsWithCustomFormatting = Boolean.Parse(WithCustomFormat)
                Form.DataCenter.ProgramConfig.IsMainPlan = False
                Form.DataCenter.ProgramConfig.MainPlanHCID = dtDraft.Rows(0).Item("MainPlanHCID")
                Form.DataCenter.ProgramConfig.Region = Trim(dtDraft.Rows(0).Item("Region"))
                Form.DataCenter.ProgramConfig.FileStatus = Trim(dtDraft.Rows(0).Item("FileStatus"))

                Try
                    Dim objDat As New CT.Data.MessagePassing
                    Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                    Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
                Catch ex As Exception

                End Try
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = True

                Dim strResult As String = LoadSpecificPlan(Boolean.Parse(WithCustomFormat))
                If strResult <> String.Empty Then Throw New Exception(strResult)


                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

                Try
                    With Form.DataCenter.GlobalValues.WS
                        Try
                            .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                        Catch ex As Exception
                        End Try
                        Try
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column), .Cells(Form.DataCenter.GlobalValues.TotalRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).TextToColumns()
                        Catch ex As Exception
                        End Try
                        If Form.DataCenter.ProgramConfig.IsWithCustomFormatting = False Then
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Phase
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_ID
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Specification_CBG
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_HC_ID_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_HC_ID
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_XCC_Team
                            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                                .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_dedicated_shared_deleted
                                .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Vin
                                .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Color
                                .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Bodystyle
                            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                                .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_CustomerRequiredDate
                                .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_RigCustomerPickDate
                            End If

                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Hardwaretype
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Vehicle_Number_Prefix
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Vehicle_Number
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Build_Id
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Tag_Number

                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Engine
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Transmission
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Emission_Stage
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Type_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Engine_Type
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Type_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Transmission_Type


                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Paint_Facility
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Driveside
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Team_Names
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Remarks
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Ship_to_Customer

                        End If

                        '.Range(.Cells(5, "W"), .Cells(2000, "W")).NumberFormat = "dd-MM-yyyy"
                        .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column), .Cells(Form.DataCenter.GlobalValues.TotalRow + 4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).NumberFormat = "@"

                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column)).EntireColumn.Group()
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column)).EntireColumn.Group()
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column)).EntireColumn.Group()
                        End If

                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column)).EntireColumn.Group()

                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column)).EntireColumn.Group()
                        End If

                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column)).EntireColumn.Group()
                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).EntireColumn.Group()
                        .Outline.ShowLevels(0, 1)
                        Dim lcol, fcol As Integer
                        Form.DataCenter.VehicleProgramInfoColumns.FindVehicleInfoFirstLastColumns(fcol, lcol)

                        .Cells(5, lcol - 1).Select()
                        .Application.ActiveWindow.FreezePanes = True
                        .Application.ActiveWindow.Zoom = 70

                        If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = True
                        If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = True
                        If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = True
                        If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = True
                        If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = True
                        If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = True
                        If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = True

                        'Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
                        'Dim _strUserPermissionLevel As String = String.Empty
                        'Try
                        '    If Form.DataCenter.GlobalValues.strUserPermissionLevel = Nothing Then
                        '        '--------------------------------------------------------------------------
                        '        ' validation for controlling the result of DAL
                        '        '--------------------------------------------------------------------------
                        '        _strUserPermissionLevel = objPer.GetPermissionLevel(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.IsGeneric)
                        '        If _strUserPermissionLevel Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        '        Form.DataCenter.GlobalValues.strUserPermissionLevel = _strUserPermissionLevel
                        '    End If
                        'Catch ex As Exception
                        '    Form.DataCenter.GlobalValues.strUserPermissionLevel = String.Empty
                        '    System.Windows.Forms.MessageBox.Show(ex.Message, "Load draft plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                        'End Try


                        '----------------------------------------------------------------------------------------------------
                        'The ribbon buttons are deactive at first but after loading a plan the buttons
                        'get active
                        '----------------------------------------------------------------------------------------------------
                        'Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                        '_RibbonUtilities.UpdateRibbonButtonsState()

                        If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)
                        _obj.sbProtectPlan()

                        Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")

                        LoadDraftPlan = String.Empty

                    End With
                Catch ex0 As Exception
                    LoadDraftPlan = ex0.Message

                End Try
            Catch ex1 As Exception
                _obj.sbProtectPlan()
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
                LoadDraftPlan = ex1.Message
            Finally
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
                If IsNothing(_frmProgress) = False Then _frmProgress.Close()
                Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilities.UpdateRibbonButtonsState()
                If Form.DataCenter.ProgramConfig.HCID <> 0 Then
                    Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
                    Dim _strUserPermissionLevel As String = String.Empty
                    Try
                        If Form.DataCenter.GlobalValues.strUserPermissionLevel = Nothing Then
                            '--------------------------------------------------------------------------
                            ' validation for controlling the result of DAL
                            '--------------------------------------------------------------------------
                            _strUserPermissionLevel = objPer.GetPermissionLevel(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.IsGeneric)
                            If _strUserPermissionLevel Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                            Form.DataCenter.GlobalValues.strUserPermissionLevel = _strUserPermissionLevel
                        End If
                    Catch ex As Exception
                        Form.DataCenter.GlobalValues.strUserPermissionLevel = String.Empty
                        System.Windows.Forms.MessageBox.Show(ex.Message, "Load draft plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                    Finally
                        If Not Form.DataCenter.GlobalValues.strUserPermissionLevel Is Nothing Then
                            If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                                objRestrictUser.DisableRibbonButtonsForViewer()
                            Else
                                Dim clsobj As New Form.DataCenter.ModuleFunction
                                clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                            End If
                        End If
                    End Try
                End If
                Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
                Form.DataCenter.GlobalValues.bolPlanIsLoading = False
            End Try

        End Function


        Public Function CheckOutPlan() As String


            If _frmProgress Is Nothing Then
                _frmProgress = New frmProgressbar()
                _frmProgress.Tag = LoadType.Refreshing.ToString
                _frmProgress.Show()
            Else
                _frmProgress.Show()
            End If

            Try
                CheckOutPlan = String.Empty


                Dim Answer As String = String.Empty
                Dim _HCID As Integer = 0

                Dim _PlanInterface As Data.Interfaces.PlanInterface

                If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _PlanInterface = New Data.VehiclePlan.Plan
                ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _PlanInterface = New Data.RigPlan.Plan
                Else
                    Exit Try
                End If


                Globals.ThisAddIn.Application.ScreenUpdating = False
                Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

                If Form.DataCenter.ProgramConfig.IsGeneric = True Then Throw New Exception("000:Check-out option is only for 'Master Specific' plans.")

                If Form.DataCenter.ProgramConfig.FileStatus <> CT.Data.DataCenter.FileStatus.Master.ToString Then Throw New Exception("000:Only Master plan can be checked-out.")

                _HCID = Form.DataCenter.ProgramConfig.HCID
                If _PlanInterface.GenerateDraftOrCheckout(_HCID, CT.Data.DataCenter.FileStatus.Checkedout, Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(Data.DataCenter.GlobalValues.message)

                '-----------------------------------------------------------------------------------------------
                ' load the check-in version.
                ' After calling GenerateDraftOrCheckout function the primary keys will be generated
                ' therefore the plan must be refreshed and loaded again.
                '-----------------------------------------------------------------------------------------------
                If Form.DataCenter.ProgramConfig.IsMainPlan = True And _HCID <> 0 Then
                    Answer = RefreshPlan(_HCID, Form.DataCenter.ProgramConfig.IsGeneric, Form.DataCenter.ProgramConfig.IsWithCustomFormatting, Form.DataCenter.ProgramConfig.BuildType)
                    'If Answer <> String.Empty Then Throw New Exception(Answer)
                    If Answer = False Then Throw New Exception(Data.DataCenter.GlobalValues.message)
                    '
                    CheckOutPlan = String.Empty
                Else
                    CheckOutPlan = "HealthChartId was not correct."
                End If

            Catch ex As Exception
                CheckOutPlan = ex.Message
            Finally
                If _frmProgress IsNot Nothing Then _frmProgress.Close()
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True

                Globals.ThisAddIn.Application.EnableEvents = True
                Form.DataCenter.GlobalValues.WS.Application.EnableEvents = True

                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault


            End Try
        End Function


        Public Function RefreshPlan(HCID As Integer, IsGeneric As Boolean, WithCustomFormat As Boolean, BuildType As String) As Boolean

            Dim WB As Excel.Workbook
            Dim _obj As frmHCIDSelect
            Dim objGlobal As New Form.DataCenter.ModuleFunction
            Dim strMessage As String = String.Empty
            Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities

            Try

                _RibbonUtilities.DeactiveRibbonButtonsState()

                If _frmProgress Is Nothing Then
                    _frmProgress = New frmProgressbar()
                    _frmProgress.Tag = LoadType.Refreshing.ToString
                    _frmProgress.Show()
                Else
                    _frmProgress.Show()
                End If


                WB = Form.DataCenter.GlobalValues.WS.Parent
                WB.Application.EnableEvents = False
                WB.Close(False)

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\CT", "Manifest", Nothing)
                readValue = readValue.ToString().Substring(0, readValue.ToString().LastIndexOf("/") + 1).Replace("file:///", "") + "TndTemplate/TndTemplate.xltm"
                Globals.ThisAddIn.Application.Workbooks.Add(readValue)
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = True


                WB = Globals.ThisAddIn.Application.Workbooks(Globals.ThisAddIn.Application.Workbooks.Count)
                If WB.Name Like "TndTemplate*" Then
                    WB.Activate()
                    WB.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()
                    Form.DataCenter.GlobalValues.objWBCurrent = WB
                End If

                Try
                    CT.Form.DataCenter.GlobalValues.wsEve = New Form.DisplayUtilities.clsWorksheetEvents
                Catch ex As Exception
                End Try
                Form.DataCenter.GlobalValues.bolPlanIsLoading = True


                strMessage = LoadPlan(HCID, IsGeneric, WithCustomFormat, BuildType)
                If strMessage <> String.Empty Then Throw New Exception(strMessage)

                If Form.DataCenter.GlobalValues.bolRefreshCompleted = True Then

                    Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    Globals.ThisAddIn.Application.EnableEvents = False
                    Globals.ThisAddIn.Application.DisplayAlerts = False


                    If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = True
                    If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = True
                    If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = True
                    If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = True
                    If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = True
                    If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = True
                    If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = True
                    '-------------------------------------------------------------------
                    ' This function has beed centralized to make the maintenance easier 
                    '-------------------------------------------------------------------

                    _RibbonUtilities.UpdateRibbonButtonsState()


                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)
                    objGlobal.sbProtectPlan()



                Else
                    Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)
                End If
                _ErrorMessage = String.Empty
                RefreshPlan = True

            Catch ex As Exception
                objGlobal.sbProtectPlan()
                _ErrorMessage = ex.Message
                RefreshPlan = False
            Finally
                _obj = Nothing
                objGlobal.sbProtectPlan()
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                If Form.DataCenter.ProgramConfig.HCID <> 0 Then
                    Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
                    Dim _strUserPermissionLevel As String = String.Empty
                    Try
                        If Form.DataCenter.GlobalValues.strUserPermissionLevel = Nothing Then
                            '--------------------------------------------------------------------------
                            ' validation for controlling the result of DAL
                            '--------------------------------------------------------------------------
                            _strUserPermissionLevel = objPer.GetPermissionLevel(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.IsGeneric)
                            If _strUserPermissionLevel Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                            Form.DataCenter.GlobalValues.strUserPermissionLevel = _strUserPermissionLevel
                        End If
                    Catch ex As Exception
                        Form.DataCenter.GlobalValues.strUserPermissionLevel = String.Empty
                        System.Windows.Forms.MessageBox.Show(ex.Message, "Refresh plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                    Finally
                        If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                            objRestrictUser.DisableRibbonButtonsForViewer()
                        Else
                            Dim clsobj As New Form.DataCenter.ModuleFunction
                            clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                        End If
                    End Try

                End If
                Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
                Form.DataCenter.GlobalValues.bolRefreshCompleted = True
                Form.DataCenter.GlobalValues.bolPlanIsLoading = False
                If _frmProgress IsNot Nothing Then _frmProgress.Close()
            End Try
        End Function



        ''' <summary>
        ''' This function is general loading Plan generic and specific and
        ''' here will be decided to load plan generic of specific.
        ''' </summary>
        ''' <param name="HCID"></param>
        ''' <param name="IsGeneric"></param>
        ''' <param name="WithCustomFormat"></param>
        ''' <param name="BuildType"></param>
        ''' <returns></returns>
        Public Function LoadPlan(HCID As Integer, IsGeneric As Boolean, WithCustomFormat As Boolean, BuildType As String) As String
            Dim _obj As New Form.DataCenter.ModuleFunction

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Function
            End If

            Try

                '--------------------------------------------------------------
                ' for error tracking
                '--------------------------------------------------------------
                LoadPlan = String.Empty


                'disable screen updating
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

                Form.DataCenter.GlobalValues.Clear()

                If IsGeneric = False Then
                    Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
                    Dim _strUserPermissionLevel As String = String.Empty
                    Try
                        If Form.DataCenter.GlobalValues.strUserPermissionLevel = Nothing Then
                            '--------------------------------------------------------------------------
                            ' validation for controlling the result of DAL
                            '--------------------------------------------------------------------------
                            _strUserPermissionLevel = objPer.GetPermissionLevel(BuildType, HCID, IsGeneric)
                            If _strUserPermissionLevel Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                            Form.DataCenter.GlobalValues.strUserPermissionLevel = _strUserPermissionLevel
                        End If
                    Catch ex As Exception
                        Form.DataCenter.GlobalValues.strUserPermissionLevel = String.Empty
                        System.Windows.Forms.MessageBox.Show(ex.Message, "Load plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                    Finally
                        If Form.DataCenter.GlobalValues.strUserPermissionLevel = "" Then
                            Throw New Exception("Access denied! Please contact 'AEREN8@FORD.COM' or OMEIGEN@FORD.COM or PNEZHAD@FORD.COM OR MAGES@FORD.COM for permission!")
                        End If
                    End Try

                End If
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = True


                '----------------------------------------------------------------------
                ' display progress on owner
                '----------------------------------------------------------------------
                UpdateProgressbar(3)


                If IsGeneric = False Then

                    Dim _AllPlans As System.Data.DataTable = _PlanInterface.SelectAllSpecificTndPlans
                    If _AllPlans Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    Dim _RowSelectedHCID() As DataRow = _AllPlans.Select(" [HealthChartID] = " + HCID.ToString)

                    '----------------------------------------------------
                    ' Only one plan must be found here
                    '----------------------------------------------------
                    Select Case _RowSelectedHCID.Length
                        Case 0
                            Throw New Exception("The HCID was not found.")
                        Case > 1
                            Throw New Exception("The count of plans with same HCID are more than one.")
                    End Select



                    Form.DataCenter.ProgramConfig.pe01 = Long.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.pe01_TnDBasicProgram_FK.ToString))
                    Form.DataCenter.ProgramConfig.HCID = Integer.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.HealthChartId.ToString))
                    Form.DataCenter.ProgramConfig.IsGeneric = If(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.GenericSpecific.ToString) = "Generic", True, False)
                    Form.DataCenter.ProgramConfig.pe02 = Long.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.pe02.ToString))
                    Form.DataCenter.ProgramConfig.XccPe26 = Long.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.XCCpe26.ToString))
                    Form.DataCenter.ProgramConfig.XccPe01 = Long.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.XCCpe01.ToString))
                    Form.DataCenter.ProgramConfig.AssyBuildScale = Long.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.AssyBuildScale.ToString))
                    Form.DataCenter.ProgramConfig.BuildType = _RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.BuildType.ToString).ToString
                    Form.DataCenter.ProgramConfig.BuildPhase = _RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.BuildPhase.ToString).ToString
                    Form.DataCenter.ProgramConfig.Carline = _RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.Carline.ToString).ToString
                    Form.DataCenter.ProgramConfig.Platform = _RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.Platform.ToString).ToString
                    Form.DataCenter.ProgramConfig.HCIDName = _RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.HealthChartName.ToString).ToString
                    Form.DataCenter.ProgramConfig.IsWithCustomFormatting = WithCustomFormat
                    Form.DataCenter.ProgramConfig.IsMainPlan = True
                    Form.DataCenter.ProgramConfig.Region = Trim(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.Region.ToString).ToString)
                    Form.DataCenter.ProgramConfig.FileStatus = Trim(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.FileStatus.ToString).ToString) ' for check-Out/-In logic
                    Form.DataCenter.ProgramConfig.MainPlanHCID = If(Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Checkedout.ToString, Integer.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.HealthChartId.ToString).ToString.Substring(3)), Integer.Parse(_RowSelectedHCID(0)(CT.Data.VehiclePlan.Plan.SelectAllSpecificTndPlansColumns.HealthChartId.ToString))) ' for check-Out/-In logic
                    Try
                        Dim objDat As New CT.Data.MessagePassing
                        Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                        Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
                    Catch ex As Exception

                    End Try
                ElseIf IsGeneric = True Then

                    Dim _AllPlans As System.Data.DataTable = _PlanInterface.SelectAllGenericTndPlan

                    If _AllPlans Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Dim _RowSelectedHCID() As DataRow = _AllPlans.Select(" [HealthChartID] = " + HCID.ToString)

                    '----------------------------------------------------
                    ' Only one plan must be found here
                    '----------------------------------------------------
                    Select Case _RowSelectedHCID.Length
                        Case 0
                            Throw New Exception("The HCID was not found.")
                        Case > 1
                            Throw New Exception("The count of plans with same HCID are more than one.")
                    End Select


                    Form.DataCenter.ProgramConfig.pe01 = Long.Parse(_RowSelectedHCID(0)("pe01_TnDBasicProgram_FK"))
                    Form.DataCenter.ProgramConfig.HCID = Integer.Parse(_RowSelectedHCID(0)("HealthChartID"))
                    Form.DataCenter.ProgramConfig.IsGeneric = If(_RowSelectedHCID(0)("GenericSpecific") = "Generic", True, False)
                    Form.DataCenter.ProgramConfig.pe02 = Long.Parse(_RowSelectedHCID(0)("pe02"))
                    Form.DataCenter.ProgramConfig.XccPe26 = Long.Parse(_RowSelectedHCID(0)("XCCpe26"))
                    Form.DataCenter.ProgramConfig.XccPe01 = Long.Parse(_RowSelectedHCID(0)("XCCpe01"))
                    Form.DataCenter.ProgramConfig.AssyBuildScale = Long.Parse(_RowSelectedHCID(0)("AssyBuildScale"))
                    Form.DataCenter.ProgramConfig.BuildType = _RowSelectedHCID(0)("BuildType").ToString
                    Form.DataCenter.ProgramConfig.BuildPhase = _RowSelectedHCID(0)("BuildPhase").ToString
                    Form.DataCenter.ProgramConfig.Carline = _RowSelectedHCID(0)("Carline").ToString
                    Form.DataCenter.ProgramConfig.Platform = _RowSelectedHCID(0)("Platform").ToString
                    Form.DataCenter.ProgramConfig.HCIDName = _RowSelectedHCID(0)("HealthChartName").ToString
                    Form.DataCenter.ProgramConfig.IsWithCustomFormatting = False   'In generic case always this value is False
                    Form.DataCenter.ProgramConfig.IsMainPlan = True
                    Form.DataCenter.ProgramConfig.Region = Trim(_RowSelectedHCID(0)("Region").ToString)
                    Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString
                    Form.DataCenter.ProgramConfig.MainPlanHCID = If(Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Checkedout.ToString, Integer.Parse(_RowSelectedHCID(0)("HealthChartID").ToString.Substring(3)), Integer.Parse(_RowSelectedHCID(0)("HealthChartID"))) ' for check-Out/-In logic

                End If

                UpdateProgressbar(3)

                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = True
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    Dim strResult As String = LoadSpecificPlan(WithCustomFormat)
                    If strResult <> String.Empty Then Throw New Exception(strResult)
                Else
                    '-------------------------------------------------------------------------------------------------------------
                    ' WithCustomFormat must be always false for loading generic plans.
                    ' Generic plans don't have Custom formatting.
                    ' The values for Custom formatting are not existed in related table for generic plans.
                    '-------------------------------------------------------------------------------------------------------------
                    Dim strResult As String = LoadGenericPlan(Form.DataCenter.ProgramConfig.XccPe26, Form.DataCenter.ProgramConfig.XccPe01, Form.DataCenter.ProgramConfig.AssyBuildScale, False)
                    If strResult <> String.Empty Then Throw New Exception(strResult)
                End If


                'Try
                '-----------------------------------------------------------------------------------
                ' Loading Custom formatting
                '-----------------------------------------------------------------------------------
                Form.DataCenter.GlobalValues.WS.Parent.activate()
                Form.DataCenter.GlobalValues.WS.Activate()
                With Form.DataCenter.GlobalValues.WS
                    Try
                        .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    Catch ex As Exception
                    End Try

                    If Form.DataCenter.ProgramConfig.IsWithCustomFormatting = False Then
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Phase
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_ID
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Specification_CBG
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_HC_ID_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_HC_ID
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_XCC_Team

                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_dedicated_shared_deleted
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Vin
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Bodystyle
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Color
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_CustomerRequiredDate
                            .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_RigCustomerPickDate
                        End If

                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Hardwaretype
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Vehicle_Number_Prefix
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Vehicle_Number
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Build_Id
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Tag_Number

                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Engine
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Transmission
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Emission_Stage
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Type_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Engine_Type
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Type_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Transmission_Type


                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Paint_Facility
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Driveside
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Team_Names
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Remarks
                        .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).EntireColumn.ColumnWidth = CT.Form.DataCenter.StaticColumnsWidth.VW_Ship_to_Customer
                    End If

                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column), .Cells(Form.DataCenter.GlobalValues.TotalRow + 4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).NumberFormat = "@"
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column)).EntireColumn.Group()
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column)).EntireColumn.Group()
                    ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column)).EntireColumn.Group()
                    End If

                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column)).EntireColumn.Group()
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                        .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column)).EntireColumn.Group()
                    End If
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column)).EntireColumn.Group()
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Range(.Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column), .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).EntireColumn.Group()
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Outline.ShowLevels(0, 1)
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    Dim lcol, fcol As Integer
                    Form.DataCenter.VehicleProgramInfoColumns.FindVehicleInfoFirstLastColumns(fcol, lcol)

                    .Cells(5, lcol - 1).Select()
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Application.ActiveWindow.FreezePanes = True
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()
                    .Application.ActiveWindow.Zoom = 70
                    Form.DataCenter.GlobalValues.WS.Parent.activate()
                    Form.DataCenter.GlobalValues.WS.Activate()


                    _obj.sbProtectPlan()
                End With
                'Catch ex As Exception
                'End Try
            Catch ex As Exception
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
                Form.DataCenter.GlobalValues.bolRefreshCompleted = True
                Form.DataCenter.GlobalValues.bolPlanIsLoading = False

                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.PlanClass, ex.Message), "Error in Select plan", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                LoadPlan = ex.Message
            Finally
                Form.DataCenter.GlobalValues.WS.Activate()
                Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
                Form.DataCenter.GlobalValues.bolRefreshCompleted = True
                Form.DataCenter.GlobalValues.bolPlanIsLoading = False
                _obj.sbProtectPlan()

            End Try
            'Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
            'Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Function


    End Class
End Namespace
