Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Tools.Excel
Imports System.Windows.Forms
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Collections
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Public Class RbnTnDControlPanel
    Dim menuItem(1) As String
    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
    Dim ErrorMessage As String = String.Empty
    Dim credentials As System.Net.NetworkCredential

    Private Function GetLatestVersion() As String
        Dim source
        Dim TextFilePath As String
        credentials = System.Net.CredentialCache.DefaultNetworkCredentials
        Dim fileReader As String
        Try
            '-------- Clear the return error message text ----------
            GetLatestVersion = String.Empty
            ErrorMessage = String.Empty

            '-------- Define destination path global
            Dim DestinationPath As String = String.Format("C:\Users\{0}\ct-tool", Environment.UserName)

            '---------------- Download Version file from sharepoint --------------------------
            source = New Uri("https://pd3.spt.ford.com/sites/PPEteam/SiteCollectionDocuments/PPEteam/Documents/3_PROTO_TEST_PLANNING/9_TnD_Process_Documentation/iDV-ConnectedTesting/Version.txt")
            TextFilePath = DestinationPath + "\Version.txt"
            My.Computer.Network.DownloadFile(source, TextFilePath, credentials, True, 60000I, True)

            '---------------- Read file --------------------------
            fileReader = My.Computer.FileSystem.ReadAllText(TextFilePath)

            GetLatestVersion = fileReader
        Catch ex As Exception
            GetLatestVersion = Nothing
            ErrorMessage = ex.Message
        End Try
    End Function


    Private Function StartUpdater(LatestVersion As String, UpdaterPath As String) As Boolean
        Try
            ErrorMessage = String.Empty

            MessageBox.Show(String.Format("Update package {0} is available, after pressing Ok it will be started automatic.", LatestVersion), String.Format("Update {0} to {1}", CT.My.Resources.CtVersion, LatestVersion), MessageBoxButtons.OK, MessageBoxIcon.Information)
            '------------------------------------------------------------------------------------
            ' Start updater and pass parameters to it
            '------------------------------------------------------------------------------------
            Dim startInfo As ProcessStartInfo = New ProcessStartInfo(UpdaterPath)
            startInfo.LoadUserProfile = True
            startInfo.Arguments = String.Format("CtVersion: {0},CtEnvironment:{1}", CT.My.Resources.CtVersion, CT.My.Resources.CtEnvironment)
            Dim Process As Process = Process.Start(startInfo)


            StartUpdater = True
        Catch ex As Exception
            ErrorMessage = ex.Message
            StartUpdater = False
        End Try
    End Function


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Sub btnLoadOpenTnDPlan_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLoadOpenTnDPlan.Click


        If IsEditing() = True Then Exit Sub

        Dim objGlobal As New Form.DataCenter.ModuleFunction
        Dim _frmHCIDSelect As frmHCIDSelect = Nothing
        Dim LatestVersion As String

        Try
            Dim WB As Excel.Workbook
            Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\CT", "Manifest", Nothing)

            '------------------------------------------------------------------------------------
            ' Check the current version and latest version
            '------------------------------------------------------------------------------------
            If CBool(CT.My.Resources.UsingAutomaticUpdate) = True Then
                LatestVersion = GetLatestVersion()
                If LatestVersion Is Nothing Then Throw New Exception(ErrorMessage)

                Dim LatestVersions As String() = LatestVersion.Split(".")
                Dim CurrentVersions As String() = CT.My.Resources.CtVersion.Split(".")
                Dim UpdaterPath As String = readValue.ToString().Substring(0, readValue.ToString().LastIndexOf("/") + 1) + "updater.exe"

                '--------------------------------------
                ' This logic is for stoping user to work with interface while updating DB
                '--------------------------------------
                If LatestVersions(0) = 0 Then
                    MessageBox.Show("Database is updating ...", "Update DB", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Select Case Math.Min(LatestVersions.Count, CurrentVersions.Count)
                    Case 1

                        If (CInt(LatestVersions(0)) > CInt(CurrentVersions(0))) Then ' Example latest 2 and current 1.14

                            '------------------------------------------------------------------------------------
                            ' Description
                            ' if the major latest version > major current version then don't need to check the rest and updater can get started
                            ' Start updater and pass parameters to it
                            '------------------------------------------------------------------------------------
                            If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                            Exit Sub
                        ElseIf (CInt(LatestVersions(0)) = CInt(CurrentVersions(0))) Then ' Example latest 2.1 and current 2

                            If (Math.Max(LatestVersions.Count, CurrentVersions.Count) > 1) Then
                                '------------------------------------------------------------------------------------
                                ' Description
                                ' If count of latest version is more than 1 then start updater
                                ' Start updater and pass parameters to it
                                '------------------------------------------------------------------------------------
                                If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                                Exit Sub

                            End If
                        End If
                    Case 2

                        If (CInt(LatestVersions(0)) > CInt(CurrentVersions(0))) Then ' Example latest 2.0 and current 1.15

                            '------------------------------------------------------------------------------------
                            ' Description 
                            ' if the major latest version > major current installed version then no need to check the rest and can start updater
                            ' Start updater and pass parameters to it
                            '------------------------------------------------------------------------------------
                            If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                            Exit Sub

                        ElseIf (CInt(LatestVersions(0)) = CInt(CurrentVersions(0))) Then ' Example latest 2.1 and current 2.0
                            If (CInt(LatestVersions(1)) > CInt(CurrentVersions(1))) Then ' Example latest 2.1 and current 2.0
                                '------------------------------------------------------------------------------------
                                ' Description 
                                ' if the major latest version = major current installed version then
                                '       if the minor latest version > minor current installed version then no need to check the rest and can start updater
                                ' Start updater and pass parameters to it
                                '------------------------------------------------------------------------------------
                                If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                                Exit Sub
                            ElseIf (CInt(LatestVersions(1)) = CInt(CurrentVersions(1))) Then ' Example latest 2.1.1 and current 2.1
                                If (Math.Max(LatestVersions.Count, CurrentVersions.Count) > 2) Then ' Example latest 2.1.1 and current 2.1
                                    '------------------------------------------------------------------------------------
                                    ' Description 
                                    ' if the minor latest version = minor current installed version then
                                    '       if the length of latest version is more than 2 the start updater
                                    ' Start updater and pass parameters to it
                                    '------------------------------------------------------------------------------------
                                    If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                                    Exit Sub

                                End If
                            End If

                        End If
                    Case 3
                        If (CInt(LatestVersions(0)) > CInt(CurrentVersions(0))) Then ' Example latest 3.0.0 and current 2.15.1

                            '------------------------------------------------------------------------------------
                            ' Description 
                            ' if the major latest version > major current installed version then no need to check the rest and can start updater
                            ' Start updater and pass parameters to it
                            '------------------------------------------------------------------------------------
                            If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                            Exit Sub

                        ElseIf (CInt(LatestVersions(0)) = CInt(CurrentVersions(0))) Then ' Exmple latest 3.1.0 and current 3.0.0
                            If (CInt(LatestVersions(1)) > CInt(CurrentVersions(1))) Then ' Exmple latest 3.1.0 and current 3.0.0

                                '------------------------------------------------------------------------------------
                                ' Description 
                                ' if the major latest version = major current installed version then 
                                '       if the minor latest version > the minor current version then no need to check the rest and can start the updater
                                ' Start updater and pass parameters to it
                                '------------------------------------------------------------------------------------
                                If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                                Exit Sub

                            ElseIf (CInt(LatestVersions(1)) = CInt(CurrentVersions(1))) Then ' Exmple latest 3.1.1 and current 3.1.0
                                If (CInt(LatestVersions(2)) > CInt(CurrentVersions(2))) Then

                                    '------------------------------------------------------------------------------------
                                    ' Description 
                                    ' if the minor latest version = minor current installed version then 
                                    '       if the rebuild latest version > the rebuild current version then no need to check the rest and can start the updater
                                    ' Start updater and pass parameters to it
                                    '------------------------------------------------------------------------------------
                                    If StartUpdater(LatestVersion, UpdaterPath) = False Then Throw New Exception(ErrorMessage)
                                    Exit Sub

                                ElseIf (CInt(LatestVersions(2)) = CInt(CurrentVersions(2))) Then
                                    If (Math.Max(LatestVersions.Count, CurrentVersions.Count) > 3) Then
                                        ' DO NOTHING
                                    End If

                                End If
                            End If
                        End If
                End Select
            End If

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False

            '------------------------------------------------------------------------------------
            ' Reading the path of template from registry
            '------------------------------------------------------------------------------------
            'Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\CT", "Manifest", Nothing)
            readValue = readValue.ToString().Substring(0, readValue.ToString().LastIndexOf("/") + 1) + "TndTemplate/TndTemplate.xltm"
            Globals.ThisAddIn.Application.Workbooks.Add(readValue)
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual


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

            Dim result As System.Windows.Forms.DialogResult
            _frmHCIDSelect = New frmHCIDSelect()

            '---------------------------------------------------------
            ' This part is for loading via SharePoint
            '---------------------------------------------------------
            If btnLoadOpenTnDPlan.Tag > 0 Then
                _frmHCIDSelect.txtHcid.Text = btnLoadOpenTnDPlan.Tag
                _frmHCIDSelect.Show()
                _frmHCIDSelect.Visible = True
                _frmHCIDSelect.btnOpenLoad_Click(_frmHCIDSelect.btnOpenLoad, Nothing)
            Else
                _frmHCIDSelect.Visible = False
                result = _frmHCIDSelect.ShowDialog()
            End If


            If result = System.Windows.Forms.DialogResult.OK Then


                'Hiding specification part if it has been found
                If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = True
                If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = True
                If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = True
                If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = True
                If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = True
                If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = True
                If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = True


                'The ribbon buttons are deactive at first but after loading a plan the buttons
                'get active
                ActivateRibbon()
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)


                ''-----------------------------------------
                '' Display today Marker and validate the result
                ''-----------------------------------------
                'If Form.DataCenter.GlobalSections.AddTodayMarker() = False Then Throw New Exception(Form.DataCenter.GlobalSections.ErrorMessage)

                If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)

                System.Windows.Forms.MessageBox.Show("The Plan has been loaded successfuly.", "Loading Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

                Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")

                'for listing the draft versions under load draft button.
                loadMenubutton()
                loadActiveUsers()
            ElseIf result = System.Windows.Forms.DialogResult.Cancel Then

                '--------------------------------------------------------------------------------
                ' Close currect not used template
                '--------------------------------------------------------------------------------
                Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)

            Else
                MessageBox.Show(_frmHCIDSelect.ErrorMessage, "Select and Load Plan", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)
            End If

        Catch ex As Exception

            objGlobal.sbProtectPlan()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Form.DataCenter.GlobalValues.bolPlanIsLoading = False

            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Load plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)

            Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)

        Finally
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString() Then
                Globals.ThisAddIn.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized 'Trick - to solve the ribbon frozen issue
                Globals.ThisAddIn.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized 'Trick - to solve the ribbon frozen issue
            End If


            objGlobal.sbProtectPlan()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Form.DataCenter.GlobalValues.bolPlanIsLoading = False
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()


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
                'System.Windows.Forms.MessageBox.Show(ex.Message, "CT Ribbon control", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Finally
                If IsNothing(Form.DataCenter.GlobalValues.strUserPermissionLevel) = False Then
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        objRestrictUser.DisableRibbonButtonsForViewer()
                    Else
                        Dim clsobj As New Form.DataCenter.ModuleFunction
                        clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                    End If

                End If

            End Try


        End Try
    End Sub

    Private Sub btnUpdateColumns_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdateColumns.Click

        If IsEditing() = True Then Exit Sub

        Dim _frmObject As Object
        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmAddColumn
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmAddColumn_Rig
        Else
            Exit Sub
        End If


        Try

            'disable screen updating
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            _frmObject = New frmAddColumn()
            'show the form and select a HCID
            If _frmObject.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()


            End If

        Catch ex As Exception

            'This Data.DataCenter.GlobalValues.message is only for Data Access layer if an
            'error ocurres through runnin a stored procedure, the error should be
            'writen in this variable
            'Data.DataCenter.GlobalValues.message = ex.Message

            'if an error ocurres here it must be shown direct
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Edit Column", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally


            'This part would be done either in try case or catch case
            _frmObject.Dispose()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try

    End Sub

    Private Sub btnAddUnit_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAddUnit.Click
        If IsEditing() = True Then Exit Sub
        Try
            'disable screen updating
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            Dim _frmObject As Object

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _frmObject = New frmNewVehicle
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _frmObject = New frmNewVehicle_Rig
            Else
                Exit Try
            End If

            'show the form and select a HCID
            If _frmObject.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            End If
        Catch ex As Exception
            'if an error ocurres here it must be shown direct
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Add Unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            'This part would be done either in try case or catch case
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try
    End Sub

    Private Sub btnDeleteUnit_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeleteUnit.Click

        If IsEditing() = True Then Exit Sub

        If Form.DataCenter.VehicleConfig.VehicleDisPlaySeq = 0 Or Form.DataCenter.VehicleConfig.VehiclePe02 = 0 Or Form.DataCenter.VehicleConfig.VehiclePe03 = 0 Then 'B
            System.Windows.Forms.MessageBox.Show("Sorry, this is not a valid selection. Please select a vehicle by clicking on the vehicle row and try again.", "Error in Delete Unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Dim obj As New Form.DataCenter.GlobalFunctions
        _GlobalFunctions.GetResetFilter()
        'Dim _deleteUnit As frmDeleteVehicle = Nothing

        Dim _frmObject As Object

        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmDeleteVehicle
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmDeleteVehicle_Rig
        Else
            Exit Sub
        End If

        Try
            'disable screen updating
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual
            _frmObject.ShowDialog()
            '-------------------------------------------------------------------
            ' Update undo button state
            '-------------------------------------------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

        Catch ex As Exception
            'if an error ocurres here it must be shown direct
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Delete Unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            'This part would be done either in try case or catch case
            _GlobalFunctions.ReApplyFilter()
            _frmObject.Dispose()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try
    End Sub

    Private Sub btnChangeSequence_Click(sender As Object, e As RibbonControlEventArgs) Handles btnChangeSequence.Click
        If IsEditing() = True Then Exit Sub

        Try
            If Globals.ThisAddIn.Application.Interactive = True Then
                Globals.ThisAddIn.Application.Interactive = False
                Globals.ThisAddIn.Application.Interactive = True
            End If
        Catch ex As Exception
        End Try

        If Form.DataCenter.VehicleConfig.VehicleDisPlaySeq = 0 Or Form.DataCenter.VehicleConfig.VehiclePe02 = 0 Or Form.DataCenter.VehicleConfig.VehiclePe03 = 0 Then 'B
            System.Windows.Forms.MessageBox.Show("Sorry, this is not a valid selection. Please select a vehicle by clicking on the vehicle row and try again.", "Error in Change sequence.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Exit Sub
        End If
        'Dim _frmMoveVehiclePosition As frmMoveVehiclePosition = Nothing
        Dim _frmObject As Object

        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmMoveVehiclePosition
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmMoveVehiclePosition_Rig
        Else
            Exit Sub
        End If
        Try
            'disable screen updating
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            ''show the form and select a HCID
            If _frmObject.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()

            End If
        Catch ex As Exception
            'if an error ocurres here it must be shown direct
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Change Unit Sequence", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            'This part would be done either in try case or catch case
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try
    End Sub

    Private Sub btnUpdateMRD_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdateMRD.Click
        'Dim _frmAddDates As frmAddDates = Nothing

        Dim _frmObject As Object

        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmAddDates
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmAddDates_Rig
        Else
            Exit Sub
        End If

        If IsEditing() = True Then Exit Sub

        Try

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            _frmObject = New frmAddDates()
            'show the form and select a HCID
            If _frmObject.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()

            End If

        Catch ex As Exception

            'if an error ocurres here it must be shown direct
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Update MRD", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)

        Finally

            'This part would be done either in try case or catch case
            '_variable.Dispose()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try
    End Sub

    Private Sub btnRefreshUnit_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRefreshUnit.Click

        If IsEditing() = True Then Exit Sub

        Dim intRow As Integer = Form.DataCenter.GlobalValues.WS.Application.Selection.row

        If intRow < 5 Or intRow > Form.DataCenter.GlobalValues.TotalRow + 4 Then
            System.Windows.Forms.MessageBox.Show("The row you selected is invalid. Please select a unit to refresh", "Refresh unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        Dim strUnitID As String = Form.DataCenter.GlobalValues.WS.Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value2
        'Dim cls As New Form.DataCenter.GlobalFunctions
        Dim intColumn As Integer = Form.DataCenter.GlobalValues.WS.Application.Selection.column
        Dim rngRow As Excel.Range = Nothing
        Dim lngPS As Long = Form.DataCenter.ProcessStepConfig.ProcessStepPe26
        Try

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            If Form.DataCenter.GlobalValues.WS.Name <> Form.DataCenter.GlobalValues.WS.Application.ActiveWorkbook.ActiveSheet.name.ToString Then
                Form.DataCenter.GlobalValues.WS.Activate()
            End If
            If Form.DataCenter.GlobalValues.WS.Application.Selection.cells.count > 1 Then
                strUnitID = ""
                _GlobalFunctions.GetResetFilter()
                For Each rngRow In Form.DataCenter.GlobalValues.WS.Application.Selection.rows
                    strUnitID = strUnitID & "," & Form.DataCenter.GlobalValues.WS.Cells(rngRow.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value2
                    Try
                        _GlobalFunctions.UpdateSection(rngRow.Row, rngRow.Row, False, True)
                    Catch ex As Exception
                    End Try
                Next
                _GlobalFunctions.ReApplyFilter()
            Else
                _GlobalFunctions.UpdateSection(intRow, intRow)
            End If
            'Form.DataCenter.GlobalValues.WS.Range("W5:W2000").NumberFormat = "dd-MM-yyyy"
            Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column), Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.WS.UsedRange.Address.ToString().Split("$")(4)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).NumberFormat = "@"
            'Form.DataCenter.GlobalValues.WS.ScrollArea = Form.DataCenter.GlobalValues.WS.UsedRange.Address
        Catch ex As Exception

            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            _GlobalFunctions.SelectAfterRefresh(lngPS)
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error refreshing unit ID '" & strUnitID & "'.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Exit Sub
        End Try

        Globals.ThisAddIn.Application.ScreenUpdating = True
        Globals.ThisAddIn.Application.EnableEvents = True
        Globals.ThisAddIn.Application.DisplayAlerts = True
        Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        _GlobalFunctions.SelectAfterRefresh(lngPS)
        System.Windows.Forms.MessageBox.Show("Unit ID '" & strUnitID & "' refreshed!", "Refresh unit completed.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information, System.Windows.Forms.MessageBoxDefaultButton.Button1)
    End Sub

    'Private Sub XX_btnRefreshUnit_Click(sender As Object, e As RibbonControlEventArgs) 'Handles btnRefreshUnit.Click

    '    If IsEditing() = True Then Exit Sub

    '    Dim intRow As Integer = Form.DataCenter.GlobalValues.WS.Application.Selection.row

    '    If intRow < 5 Or intRow > Form.DataCenter.GlobalValues.TotalRow + 4 Then
    '        System.Windows.Forms.MessageBox.Show("The row you selected is invalid. Please select a unit to refresh", "Refresh unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
    '        Exit Sub
    '    End If

    '    Dim strUnitID As String = Form.DataCenter.GlobalValues.WS.Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value2
    '    'Dim cls As New Form.DataCenter.GlobalFunctions
    '    Dim intColumn As Integer = Form.DataCenter.GlobalValues.WS.Application.Selection.column
    '    Dim rngRow As Excel.Range = Nothing
    '    Dim lngPS As Long = Form.DataCenter.ProcessStepConfig.ProcessStepPe26
    '    Try

    '        Globals.ThisAddIn.Application.ScreenUpdating = False
    '        Globals.ThisAddIn.Application.EnableEvents = False
    '        Globals.ThisAddIn.Application.DisplayAlerts = False
    '        If Form.DataCenter.GlobalValues.WS.Name <> Form.DataCenter.GlobalValues.WS.Application.ActiveWorkbook.ActiveSheet.name.ToString Then
    '            Form.DataCenter.GlobalValues.WS.Activate()
    '        End If
    '        If Form.DataCenter.GlobalValues.WS.Application.Selection.cells.count > 1 Then
    '            strUnitID = ""
    '            _GlobalFunctions.GetResetFilter()
    '            For Each rngRow In Form.DataCenter.GlobalValues.WS.Application.Selection.rows
    '                strUnitID = strUnitID & "," & Form.DataCenter.GlobalValues.WS.Cells(rngRow.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value2
    '                Try
    '                    _GlobalFunctions.UpdateSection(rngRow.Row, rngRow.Row, False, True)
    '                Catch ex As Exception
    '                End Try
    '            Next
    '            _GlobalFunctions.ReApplyFilter()
    '        Else
    '            _GlobalFunctions.UpdateSection(intRow, intRow)
    '        End If
    '        'Form.DataCenter.GlobalValues.WS.Range("W5:W2000").NumberFormat = "dd-MM-yyyy"
    '        Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column), Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.WS.UsedRange.Address.ToString().Split("$")(4)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).NumberFormat = "@"
    '        'Form.DataCenter.GlobalValues.WS.ScrollArea = Form.DataCenter.GlobalValues.WS.UsedRange.Address
    '    Catch ex As Exception

    '        Globals.ThisAddIn.Application.ScreenUpdating = True
    '        Globals.ThisAddIn.Application.EnableEvents = True
    '        Globals.ThisAddIn.Application.DisplayAlerts = True
    '        Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
    '        _GlobalFunctions.SelectAfterRefresh(lngPS)
    '        System.Windows.Forms.MessageBox.Show(ex.Message, "Error refreshing unit ID '" & strUnitID & "'.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
    '        Exit Sub
    '    End Try

    '    Globals.ThisAddIn.Application.ScreenUpdating = True
    '    Globals.ThisAddIn.Application.EnableEvents = True
    '    Globals.ThisAddIn.Application.DisplayAlerts = True
    '    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
    '    _GlobalFunctions.SelectAfterRefresh(lngPS)
    '    System.Windows.Forms.MessageBox.Show("Unit ID '" & strUnitID & "' refreshed!", "Refresh unit completed.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information, System.Windows.Forms.MessageBoxDefaultButton.Button1)
    'End Sub

    Private Sub btnConvertToSpecific_Click(sender As Object, e As RibbonControlEventArgs) Handles btnConvertToSpecific.Click

        '-------------------------------------------------
        ' validate the excel preadsheet to not to be in edit mode
        '-------------------------------------------------
        If IsEditing() = True Then Exit Sub

        Try

            '-----------------------------------------------
            ' disable screen updating
            '-----------------------------------------------
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual


            '-----------------------------------------------
            ' disable screen updating
            '-----------------------------------------------
            Select Case Form.DataCenter.ProgramConfig.BuildType
                Case CT.Data.DataCenter.BuildType.Vehicle.ToString()

                    '-----------------------------------------------
                    ' Call the function from Ribbon logic layer
                    '-----------------------------------------------
                    Dim _VehiclePlan As CT.RbnTnDControlPanelLogic.VehiclePlan = New RbnTnDControlPanelLogic.VehiclePlan()

                    Select Case _VehiclePlan.CovertGeneric2Spesific()
                        Case DialogResult.OK
                            System.Windows.Forms.MessageBox.Show("The Plan has been refreshed successfuly.", "Loading Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                        Case DialogResult.Cancel
                            ' DO NOTHING
                        Case DialogResult.None
                            Throw New Exception(_VehiclePlan.ErrorMessage)
                    End Select

                Case CT.Data.DataCenter.BuildType.Rig.ToString()

                    '-----------------------------------------------
                    ' Call the function from Ribbon logic layer
                    '-----------------------------------------------
                    Dim _RigPlan As CT.RbnTnDControlPanelLogic.RigPlan = New RbnTnDControlPanelLogic.RigPlan()

                    Select Case _RigPlan.CovertGeneric2Spesific()
                        Case DialogResult.OK
                            System.Windows.Forms.MessageBox.Show("The Plan has been refreshed successfuly.", "Loading Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                        Case DialogResult.Cancel
                            ' DO NOTHING
                        Case DialogResult.None
                            Throw New Exception(_RigPlan.ErrorMessage)
                    End Select


                Case CT.Data.DataCenter.BuildType.Buck.ToString()

                    '-----------------------------------------------
                    ' Call the function from Ribbon logic layer
                    '-----------------------------------------------
                    'Dim _BuckPlan As CT.RbnTnDControlPanelLogic.BuckPlan = New RbnTnDControlPanelLogic.BuckPlan()
                    ' FOR LATER IMPLEMENTATION


                Case Else
                    Throw New Exception("This build type logic has not been developed yet.")
            End Select



        Catch ex As Exception
            '-----------------------------------------------
            ' if an error ocurres here it must be shown direct
            '-----------------------------------------------
            System.Windows.Forms.MessageBox.Show(ex.Message, "Convert Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            '--------------------------------------------------------------------------
            ' activate ribbon
            '--------------------------------------------------------------------------
            ActivateRibbon()
            '-----------------------------------------------
            ' This part would be done either in try case or catch case
            '-----------------------------------------------
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
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
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Tnd Ribbon control", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                Finally
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        objRestrictUser.DisableRibbonButtonsForViewer()
                    Else
                        Dim clsobj As New Form.DataCenter.ModuleFunction
                        clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                    End If

                End Try
            End If
        End Try
    End Sub


    Public Property btnSearchFilter_Enable As Boolean
        Set
            Me.btnSearchFilter.Enabled = False
        End Set
        Get
            Return Me.btnSearchFilter.Enabled
        End Get
    End Property

    ''' <summary>
    ''' This method activate the ribbon
    ''' </summary>
    Public Sub ActivateRibbon()
        '-------------------------------------------------------------------
        ' This function has beed centralized to make the maintenance easier 
        '-------------------------------------------------------------------
        Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
        _RibbonUtilitis.UpdateRibbonButtonsState()
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
                System.Windows.Forms.MessageBox.Show(ex.Message, "Application window activate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Finally
                If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                    objRestrictUser.DisableRibbonButtonsForViewer()
                Else
                    Dim clsobj As New Form.DataCenter.ModuleFunction
                    clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                End If

            End Try
        End If

    End Sub
    Public Sub DeActivateRibbon()

        '-------------------------------------------------------------------
        ' This function has beed centralized to make the maintenance easier 
        '-------------------------------------------------------------------
        Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
        _RibbonUtilities.DeactiveRibbonButtonsState()
        'Call_DeActRibbon()
    End Sub


    'Check all toggle button - to check/uncheck the 'ShowAll Toggle Button'
    Private Sub checkall_Togglebuttonstate()
        If togInstrumentation.Checked = True And
            togNonMFCSpecification.Checked = True And
            togMfcSpecification.Checked = True And
            togProgramInformation.Checked = True And
            togFurtherBasicSpecification.Checked = True And
            togUserShipping.Checked = True And
            togUpdatePack.Checked = True And
            togTiming.Checked = True Then
            togShowAll.Checked = True
        Else
            togShowAll.Checked = False
        End If
    End Sub

    Private Sub togInstrumentation_Click(sender As Object, e As RibbonControlEventArgs) Handles togInstrumentation.Click
        If togInstrumentation.Checked = True Then
            If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togFurtherBasicSpecification_Click(sender As Object, e As RibbonControlEventArgs) Handles togFurtherBasicSpecification.Click
        If togFurtherBasicSpecification.Checked = True Then
            If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togNonMFCSpecification_Click(sender As Object, e As RibbonControlEventArgs) Handles togNonMFCSpecification.Click
        If togNonMFCSpecification.Checked = True Then
            If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togMfcSpecification_Click(sender As Object, e As RibbonControlEventArgs) Handles togMfcSpecification.Click
        If togMfcSpecification.Checked = True Then
            If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togProgramInformation_Click(sender As Object, e As RibbonControlEventArgs) Handles togProgramInformation.Click
        If togProgramInformation.Checked = True Then
            If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togUserShipping_Click(sender As Object, e As RibbonControlEventArgs) Handles togUserShipping.Click
        If togUserShipping.Checked = True Then
            If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togUpdatePack_Click(sender As Object, e As RibbonControlEventArgs) Handles togUpdatePack.Click
        If togUpdatePack.Checked = True Then
            If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False
        Else
            If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = True
        End If
        checkall_Togglebuttonstate()
    End Sub

    Private Sub togShowAll_Click(sender As Object, e As RibbonControlEventArgs) Handles togShowAll.Click
        If togShowAll.Checked = True Then

            If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False
            If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False
            If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False
            If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False
            If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False
            If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False
            If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False


            togInstrumentation.Checked = True
            togNonMFCSpecification.Checked = True
            togMfcSpecification.Checked = True
            togProgramInformation.Checked = True
            togFurtherBasicSpecification.Checked = True
            togUserShipping.Checked = True
            togUpdatePack.Checked = True
            togShowAll.Checked = True
            togTiming.Checked = True

        Else


            If Form.DataCenter.GlobalSections.InstrumentationSection IsNot Nothing Then Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = True
            If Form.DataCenter.GlobalSections.MfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = True
            If Form.DataCenter.GlobalSections.NonMfcSpecificationSection IsNot Nothing Then Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = True
            If Form.DataCenter.GlobalSections.ProgramInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = True
            If Form.DataCenter.GlobalSections.FurtherBasicInformationSection IsNot Nothing Then Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = True
            If Form.DataCenter.GlobalSections.UpdatePackSection IsNot Nothing Then Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = True
            If Form.DataCenter.GlobalSections.UserShippingDetailsSection IsNot Nothing Then Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = True


            togInstrumentation.Checked = False
            togNonMFCSpecification.Checked = False
            togMfcSpecification.Checked = False
            togProgramInformation.Checked = False
            togFurtherBasicSpecification.Checked = False
            togUserShipping.Checked = False
            togUpdatePack.Checked = False
            togShowAll.Checked = False
            togTiming.Checked = True

        End If

    End Sub

    Private Sub btnUpdateClimate_Click(sender As Object, e As RibbonControlEventArgs)
        Dim frmMe As New frmNewVehicle
        frmMe.Show()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Private Sub btnSearchFilter_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSearchFilter.Click

        If IsEditing() = True Then Exit Sub

        Dim _frmSearch As frmSearch = Nothing
        _frmSearch = New frmSearch()
        _frmSearch.ShowDialog()
    End Sub

    Private Sub btnExportToExcel_Click(sender As Object, e As RibbonControlEventArgs) Handles btnExportToExcel.Click

        If IsEditing() = True Then Exit Sub

        Dim _frmObject As Object

        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmExporttoexcel
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmExporttoexcel_Rig
        Else
            Exit Sub
        End If

        'Dim _frmExporttoexcel As frmExporttoexcel = Nothing
        _frmObject.ShowDialog()
    End Sub

    Private Sub btnUnitReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUnitReport.Click

        If IsEditing() = True Then Exit Sub
        Try
            'Globals.ThisAddIn.Application.EnableEvents = False
            Dim strInput As String
            strInput = InputBox("Please enter the ID of unit to generate the Unit Gantt chart report.", "Unit Report", Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.Application.Selection.row, 5).Value)

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.CopyObjectsWithCells = True


            Dim res As Double
            'Form.DataCenter.GlobalValues.WS.Cells(
            res = Form.DataCenter.GlobalValues.WS.Application.WorksheetFunction.Match(strInput, Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count + 10, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column)), 0) ', Form.DataCenter.GlobalValues.WS.Range("E5:E" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count + 10).Column)
            'Form.DataCenter.GlobalValues.WS.Cells(5, Form.DataCenter.Vehicle_ID_Column), Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count + 10, Form.DataCenter.Vehicle_ID_Column).entirecolumn

            Dim intRow As Integer
            If res > 0 Then intRow = res + 4
            If intRow <= 4 Then
                System.Windows.Forms.MessageBox.Show("The given vehicle ID is not exist", "Unit Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Exit Sub
            End If

            If Val(strInput) = 0 Then
                System.Windows.Forms.MessageBox.Show("Sorry, your input Is Not valid.", "Unit Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                Exit Try
            End If

            Dim _unitreport As Form.Reports.UnitReport = New Form.Reports.UnitReport()
            _unitreport.VehicleReport(Form.DataCenter.GlobalValues.WS, Form.DataCenter.GlobalValues.VehicleReportWs, intRow)
        Catch ex As Exception
            If InStr(ex.Message, "Match method", CompareMethod.Text) > 0 Then
                System.Windows.Forms.MessageBox.Show("The given vehicle ID is not exist", "Unit Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
            Else
                System.Windows.Forms.MessageBox.Show(ex.Message, "Unit Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End If
        Finally
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.CopyObjectsWithCells = False
        End Try
    End Sub

    Private Sub btnUndo_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUndo.Click
        Dim _DataTable As System.Data.DataTable
        Dim _ChangeLog As New CT.Data.ChangeLog

        If IsEditing() = True Then Exit Sub

        Try

            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait


            If _ChangeLog.IsUndoCommandAvailable(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType) = True Then
                _DataTable = _ChangeLog.GetTnDLastUndo(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                If _DataTable.Rows.Count = 1 Then

                    If System.Windows.Forms.MessageBox.Show(String.Format("Action Name: {0} ", _DataTable.Rows(0)("ActionName").ToString) + vbNewLine +
                                                            String.Format("Action Nr.: {0} ", _DataTable.Rows(0)("ActionId").ToString) + vbNewLine +
                                                            String.Format("Description: {0} ", _DataTable.Rows(0)("Remark").ToString) + vbNewLine, "Undo", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then

                        Form.DisplayUtilities.Ribbon.UndoButton.Click(Val(_DataTable.Rows(0)("Pe61")), String.Format("Action Name: {0} ", _DataTable.Rows(0)("ActionName").ToString) + vbNewLine +
                                                            String.Format("Action Nr.: {0} ", _DataTable.Rows(0)("ActionId").ToString) + vbNewLine)

                    End If

                Else
                    btnUndo.Enabled = False
                End If
            Else
                btnUndo.Enabled = False
            End If

        Catch ex As Exception

            System.Windows.Forms.MessageBox.Show(ex.Message, "Error while undoing", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try
    End Sub

    Private Sub btnGenerateDraft_Click(sender As Object, e As RibbonControlEventArgs)



        '-----------------------------------------------------------------------------
        ' Generate Draft
        '-----------------------------------------------------------------------------
        Try
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                System.Windows.Forms.MessageBox.Show("Draft option is only for 'Specific' plans.", "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            Dim resultDataTable As New System.Data.DataTable
            resultDataTable = _PlanInterface.SelectAllTndDraftPlans(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID)
            If resultDataTable.Rows.Count >= 3 Then
                System.Windows.Forms.MessageBox.Show("3 Draft versions are already created for this HC ID : " & Form.DataCenter.ProgramConfig.HCID, "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            If _PlanInterface.GenerateDraftOrCheckout(Form.DataCenter.ProgramConfig.HCID, CT.Data.DataCenter.FileStatus.Draft, Form.DataCenter.ProgramConfig.BuildType) = True Then
                System.Windows.Forms.MessageBox.Show("Draft completed successfully.", "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            Else
                System.Windows.Forms.MessageBox.Show(Data.DataCenter.GlobalValues.message, "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try


    End Sub

    Private Sub togTiming_Click(sender As Object, e As RibbonControlEventArgs) Handles togTiming.Click
        checkall_Togglebuttonstate()
    End Sub

    Private Sub btnEngineTransmissionReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btnEngineTransmissionReport.Click

        If IsEditing() = True Then Exit Sub

        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            Dim _enginetransmission As Form.Reports.EngineTransmissionReport = New Form.Reports.EngineTransmissionReport()
            _enginetransmission.EngineTransmissionReport(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.Region)
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            Dim _enginetransmission As Form.Reports.EngineTransmissionReport_Rig = New Form.Reports.EngineTransmissionReport_Rig()
            _enginetransmission.EngineTransmissionReport(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.Region)
        End If
    End Sub

    'Private Sub tgOverlapping_Click(sender As Object, e As RibbonControlEventArgs)

    '    Dim rng As Excel.Range
    '    Dim Row As System.Data.DataRow
    '    Dim Dt As New Data.Plan
    '    Dim rstOL As System.Data.DataTable

    '    If IsEditing() = True Then
    '        System.Windows.Forms.MessageBox.Show("Excel interface is in Edit-Mode please remove the focuse from formula bar by selecting another cell.",
    '                                             "Error in detect overlapping", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
    '        Exit Sub
    '    End If


    '    Try
    '        With Form.DataCenter.GlobalValues.WS
    '            If tgOverlapping.Checked = True Then
    '                rstOL = Dt.DetectOverlapping(Form.DataCenter.ProgramConfig.HCID)
    '                For Each Row In rstOL.Rows
    '                    rng = Nothing
    '                    rng = .Cells(1, Form.DataCenter.Vehicle_P_0_Column).entirecolumn.Find("*;*;" & Row("pe03_TnDProgramVehicles_FK").ToString & ";" & Row("pe45_AllocatedPowerPack_FK").ToString & ";*", .Cells(4, "C"), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
    '                    If Not rng Is Nothing Then
    '                        .Range(.Cells(rng.Row, Form.DataCenter.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.Vehicle_Ship_to_Customer_Column)).Interior.Color = 13487615 'rose
    '                    End If
    '                Next
    '            Else
    '                .Range(.Cells(5, Form.DataCenter.Vehicle_Phase_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.Vehicle_Ship_to_Customer_Column)).Interior.Color = CT.My.Resources.EmptyColor
    '                .Range(.Cells(5, Form.DataCenter.Vehicle_Phase_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.Vehicle_Ship_to_Customer_Column)).Style = Style.Styles.TnsStyleName.ProcessStepStyle.ToString()
    '            End If
    '        End With
    '    Catch ex As Exception
    '    End Try
    'End Sub

    Public Sub btnRefreshPlan_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRefreshPlan.Click


        Dim intHCID As Integer = Form.DataCenter.ProgramConfig.HCID
        Dim bolIsGeneric As Boolean = Form.DataCenter.ProgramConfig.IsGeneric
        Dim bolWithIndiFormatting As Boolean = Form.DataCenter.ProgramConfig.IsWithCustomFormatting
        Dim strBuildType As String = Form.DataCenter.ProgramConfig.BuildType
        Dim Answer As String = String.Empty
        Dim _Plan As New Form.DisplayUtilities.Plan()
        Dim bolFinished As Boolean = False
        Dim BuildType As String = Form.DataCenter.ProgramConfig.BuildType

        If IsEditing() = True Then Exit Sub


        Try

            If Form.DataCenter.ProgramConfig.IsMainPlan = False And Form.DataCenter.ProgramConfig.HCID <> 0 Then

                Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)

                '------------------------------------------------------------------------------------------
                ' implement Load Draft with validation 
                '------------------------------------------------------------------------------------------
                Answer = _Plan.LoadDraftPlan(intHCID, bolWithIndiFormatting, Form.DisplayUtilities.Plan.LoadType.Refreshing, BuildType)
                If Answer <> String.Empty Then Throw New Exception(Answer)

                System.Windows.Forms.MessageBox.Show("The Draft Plan has been refreshed successfuly.", "Loading Draft Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

            ElseIf Form.DataCenter.ProgramConfig.IsMainPlan = True And Form.DataCenter.ProgramConfig.HCID <> 0 Then

                If _Plan.RefreshPlan(intHCID, bolIsGeneric, bolWithIndiFormatting, strBuildType) = False Then Throw New Exception(_Plan.ErrorMessage)
                System.Windows.Forms.MessageBox.Show("The Plan has been refreshed successfuly.", "Loading Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                bolFinished = True

            End If


        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("", "Error in Refresh plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            Dim objGlobal As New Form.DataCenter.ModuleFunction
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            'Form.DataCenter.GlobalSections.AddTodayMarker()
            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString And bolFinished Then
                Dim objfrm As New frmHCIDSelect
                objfrm.NotifyToCheckout.ShowBalloonTip(10000)
                objfrm.NotifyToCheckout.Visible = False
                Dim obj As New Form.DataCenter.ModuleFunction
                obj.DisplayMasterMessage()
            End If
            objGlobal.sbProtectPlan()
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
                System.Windows.Forms.MessageBox.Show(ex.Message, "Application window activate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

            Finally
                If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                    objRestrictUser.DisableRibbonButtonsForViewer()
                Else
                    Dim clsobj As New Form.DataCenter.ModuleFunction
                    clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                End If
            End Try

        End Try

    End Sub



    Public Sub OldVersion_btnRefreshPlan_Click(sender As Object, e As RibbonControlEventArgs) 'Handles btnRefreshPlan.Click


        Dim intHCID As Integer = Form.DataCenter.ProgramConfig.HCID
        Dim bolIsGeneric As Boolean = Form.DataCenter.ProgramConfig.IsGeneric
        Dim bolWithIndiFormatting As Boolean = Form.DataCenter.ProgramConfig.IsWithCustomFormatting
        Dim strBuildType As String = Form.DataCenter.ProgramConfig.BuildType
        Dim Answer As String = String.Empty
        Dim _Plan As New Form.DisplayUtilities.Plan()
        Dim bolFinished As Boolean = False

        If IsEditing() = True Then Exit Sub


        Try

            If Form.DataCenter.ProgramConfig.IsMainPlan = False And Form.DataCenter.ProgramConfig.HCID <> 0 Then

                Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)

                '------------------------------------------------------------------------------------------
                ' implement Load Draft with validation 
                '------------------------------------------------------------------------------------------
                'Answer = _Plan.LoadDraftPlan(intHCID, bolWithIndiFormatting, Form.DisplayUtilities.Plan.LoadType.Refreshing)
                If Answer <> String.Empty Then Throw New Exception(Answer)

                System.Windows.Forms.MessageBox.Show("The Draft Plan has been refreshed successfuly.", "Loading Draft Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

            ElseIf Form.DataCenter.ProgramConfig.IsMainPlan = True And Form.DataCenter.ProgramConfig.HCID <> 0 Then

                If _Plan.RefreshPlan(intHCID, bolIsGeneric, bolWithIndiFormatting, strBuildType) = False Then Throw New Exception(_Plan.ErrorMessage)
                System.Windows.Forms.MessageBox.Show("The Plan has been refreshed successfuly.", "Loading Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                bolFinished = True

            End If


        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("", "Error in Refresh plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            Dim objGlobal As New Form.DataCenter.ModuleFunction
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            'Form.DataCenter.GlobalSections.AddTodayMarker()
            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString And bolFinished Then
                Dim objfrm As New frmHCIDSelect
                objfrm.NotifyToCheckout.ShowBalloonTip(10000)
                objfrm.NotifyToCheckout.Visible = False
                Dim obj As New Form.DataCenter.ModuleFunction
                obj.DisplayMasterMessage()
            End If
            objGlobal.sbProtectPlan()
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
                System.Windows.Forms.MessageBox.Show(ex.Message, "Application window activate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

            Finally
                If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                    objRestrictUser.DisableRibbonButtonsForViewer()
                Else
                    Dim clsobj As New Form.DataCenter.ModuleFunction
                    clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                End If
            End Try

        End Try

    End Sub



    Private Sub btnUpdateHoliday_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdateHoliday.Click
        'Dim _frmHolidayPlan As frmHolidayPlan = Nothing

        If IsEditing() = True Then Exit Sub

        Dim _frmObject As Object
        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmHolidayPlan
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmHolidayPlan_Rig
        Else
            Exit Sub
        End If


        Try
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            If _frmObject.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()

            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error in Holiday Plan.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
        End Try
    End Sub

    Private Sub DropDown1_SelectionChanged(sender As Object, e As RibbonControlEventArgs)

    End Sub
    Private Function CreateRibbonDropDownItem() As RibbonDropDownItem
        Return Me.Factory.CreateRibbonDropDownItem()
    End Function

    Private Function CreateRibbonMenu() As RibbonMenu
        Return Me.Factory.CreateRibbonMenu()
    End Function

    Private Function CreateRibbonButton() As RibbonButton
        Dim button As RibbonButton = Me.Factory.CreateRibbonButton()
        AddHandler(button.Click), AddressOf MenuItem_Click
        Return button
    End Function
    Private Function CreateRibbonToggleButton() As RibbonToggleButton
        Dim button As RibbonToggleButton = Me.Factory.CreateRibbonToggleButton()
        AddHandler(button.Click), AddressOf MenuItem_Click
        Return button
    End Function
    Public Sub loadMenubutton()

        Try
            mnuGenerateDraft.Items.Clear()
            If Form.DataCenter.ProgramConfig.IsGeneric = False Then

                Dim dtDraft As System.Data.DataTable
                Dim _PlanInterface As Data.Interfaces.PlanInterface

                If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _PlanInterface = New Data.VehiclePlan.Plan
                ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _PlanInterface = New Data.RigPlan.Plan
                Else
                    Exit Try
                End If

                dtDraft = _PlanInterface.SelectAllTndDraftPlans(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID)

                If dtDraft Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                If dtDraft.Rows.Count > 0 Then                    '
                    Dim RB As RibbonButton = Nothing
                    For Each rows In dtDraft.Rows
                        RB = CreateRibbonButton()
                        RB.OfficeImageId = "ViewDraftView"
                        mnuGenerateDraft.Items.Add(RB)
                        CType(mnuGenerateDraft.Items.Last(), RibbonButton).Label = rows(2).ToString & " - " & rows(4).ToString
                        CType(mnuGenerateDraft.Items.Last(), RibbonButton).Tag = rows(2).ToString
                    Next
                    mnuGenerateDraft.Enabled = True
                Else
                    mnuGenerateDraft.Enabled = False
                End If

                If Form.DataCenter.ProgramConfig.IsMainPlan = True Then
                    Globals.Ribbons.RbnTnDControlPanel.btnGenerateDraft.Enabled = True
                    Globals.Ribbons.RbnTnDControlPanel.btnDeleteDraft.Enabled = False
                    Globals.Ribbons.RbnTnDControlPanel.btnReplacePlanWithDraft.Enabled = False
                Else
                    Globals.Ribbons.RbnTnDControlPanel.btnGenerateDraft.Enabled = False
                    Globals.Ribbons.RbnTnDControlPanel.btnDeleteDraft.Enabled = True
                    Globals.Ribbons.RbnTnDControlPanel.btnReplacePlanWithDraft.Enabled = True
                End If
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Load Draft Menu", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

        End Try
    End Sub

    Public Sub loadActiveUsers()

        Try
            mnuActiveUsers.Items.Clear()
            If Form.DataCenter.ProgramConfig.IsGeneric = False Then

                Dim dtUsers As System.Data.DataTable
                Dim _PlanActiveUsers As New Data.PlanActiveUsers

                If CT.Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString Then
                    dtUsers = _PlanActiveUsers.SelectAll(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

                    If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    If dtUsers.Rows.Count > 0 Then                    '
                        Dim RB As RibbonButton = Nothing
                        For Each rows In dtUsers.Rows
                            RB = CreateRibbonButton()
                            'AddHandler(Button.Click), AddressOf MenuItem_Click
                            RemoveHandler(RB.Click), AddressOf MenuItem_Click
                            RB.OfficeImageId = "AccessListContacts"
                            mnuActiveUsers.Items.Add(RB)
                            CType(mnuActiveUsers.Items.Last(), RibbonButton).Label = rows(0).ToString '& " - " & rows(4).ToString
                            'CType(mnuActiveUsers.Items.Last(), RibbonButton).Tag = rows(2).ToString
                        Next
                        mnuActiveUsers.Enabled = True
                    Else
                        mnuActiveUsers.Enabled = False
                    End If
                Else
                    mnuActiveUsers.Enabled = False
                End If
                'If Form.DataCenter.ProgramConfig.IsMainPlan = True Then
                '    Globals.Ribbons.RbnTnDControlPanel.btnGenerateDraft.Enabled = True
                '    Globals.Ribbons.RbnTnDControlPanel.btnDeleteDraft.Enabled = False
                '    Globals.Ribbons.RbnTnDControlPanel.btnReplacePlanWithDraft.Enabled = False
                'Else
                '    Globals.Ribbons.RbnTnDControlPanel.btnGenerateDraft.Enabled = False
                '    Globals.Ribbons.RbnTnDControlPanel.btnDeleteDraft.Enabled = True
                '    Globals.Ribbons.RbnTnDControlPanel.btnReplacePlanWithDraft.Enabled = True
                'End If
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Load Active Users", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MenuItem_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Try
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

            Dim _Plan As New Form.DisplayUtilities.Plan()
            Dim Answer As String = String.Empty

            Answer = _Plan.LoadDraftPlan(sender.tag, Nothing, Form.DisplayUtilities.Plan.LoadType.Loading, Form.DataCenter.ProgramConfig.BuildType)
            If Answer <> String.Empty Then Throw New Exception(Answer)
            loadMenubutton()
            System.Windows.Forms.MessageBox.Show("The Draft Plan has been Loaded successfuly.", "Loading Draft Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)
        Finally
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try

    End Sub
    Private Sub btnRedo_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRedo.Click

        If IsEditing() = True Then Exit Sub

        Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
        _RibbonUtilitis.UpdateUndoButtonsState()
    End Sub

    Private Sub btnCDSIDtoDvpTeam_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCDSIDtoDvpTeam.Click
        'Dim _frmCDSIDtoDVPName As frmCDSIDtoDVPName = Nothing

        If IsEditing() = True Then Exit Sub

        Dim _frmObject As Object
        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            _frmObject = New frmCDSIDtoDVPName
        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            _frmObject = New frmCDSIDtoDVPName_Rig
        Else
            Exit Sub
        End If

        Try
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            _frmObject.ShowDialog()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Assign CDSID to DVPname.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try
    End Sub


    Private Function IsEditing() As Boolean
        Dim cBars As Office.CommandBars = Globals.ThisAddIn.Application.CommandBars
        Dim result As Boolean = Not cBars.GetEnabledMso("FileNewDefault")
        Marshal.ReleaseComObject(cBars)
        If result = True Then
            System.Windows.Forms.MessageBox.Show("Excel interface is in Edit-Mode please remove the focus from formula bar by selecting another cell.",
                                                     "CT", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
        End If
        Return result
    End Function

    Private Sub btnClearFilter_Click(sender As Object, e As RibbonControlEventArgs) Handles btnClearFilter.Click
        Globals.ThisAddIn.Application.ScreenUpdating = False
        Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

        Form.DataCenter.GlobalValues.WS.Cells.FormatConditions.Delete()
        Form.DataCenter.GlobalValues.WS.Range("5:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).EntireRow.Hidden = False

        Form.DataCenter.GlobalValues.WS.AutoFilterMode = False
        Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)

        Dim _obj As New Form.DataCenter.ModuleFunction
        _obj.sbProtectPlan()
        Globals.ThisAddIn.Application.ScreenUpdating = True
    End Sub

    Private Sub btnPrecheckF4T_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPrecheckF4T.Click
        Try

            Dim clsReport As New Form.Reports.PrecheckF4TestReport
            If clsReport.Gen_PrecheckF4TestReport = True Then
                MessageBox.Show("Report generation is completed.", "F4Test Precheck Report", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.ExporToExcelReport, clsReport.ErrorMessage), "F4Test Precheck Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnCheckIn_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCheckIn.Click

        'Dim result As DialogResult = MessageBox.Show("Do you really want to checkin the plan?", "checkin plan", MessageBoxButtons.YesNoCancel)
        'If result = DialogResult.Cancel Or result = DialogResult.No Then
        '    Exit Sub
        'End If


        Dim strTitle As String = "Checked in plan"
        Try

            'Check the Activeusers in this plan
            Dim dtUsers As DataTable
            Dim _PlanActiveUsers As New Data.PlanActiveUsers
            Dim strActiveUsers As String = ""
            dtUsers = _PlanActiveUsers.SelectAll(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

            If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            If dtUsers.Rows.Count > 0 Then                    '
                For Each rows In dtUsers.Rows
                    strActiveUsers = strActiveUsers & rows(0).ToString() & ","
                Next
                System.Windows.Forms.MessageBox.Show("The Plan cannot be checked in now as the below users are active in this plan." & vbNewLine & strActiveUsers & "", "Active users in Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                'Cursor = Cursors.Default
                'Me.DialogResult = DialogResult.Retry
                Exit Sub
            End If

            '----------------------------------------------------------
            ' Request user to change the issue version
            '----------------------------------------------------------
            Dim vbResult As MsgBoxResult
            'vbResult = MessageBox.Show("Do you want to change the version?", strTitle, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            vbResult = MessageBox.Show("Do you want to check-in the plan with updating the version?" & vbNewLine & vbNewLine & "Please select Yes to check-in and update version" & vbNewLine &
                                       "Please select No to check-in without version update" & vbNewLine & "Please select Cancel to cancel check-in",
                                       "Check-in", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            If vbResult = DialogResult.Yes Then
                'Dim _frmHeaderEdit As frmHeaderEdit = New frmHeaderEdit
                Dim _frmObject As Object
                If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _frmObject = New frmHeaderEdit
                ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _frmObject = New frmHeaderEdit_Rig
                Else
                    Exit Sub
                End If
                _frmObject.ShowDialog()
            ElseIf vbResult = DialogResult.Cancel Then
                Exit Sub
            End If

            Dim _PlanDisplay As Form.DisplayUtilities.Plan = New Form.DisplayUtilities.Plan

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            '----------------------------------------------------------
            ' Replace checkedout version on master version
            '----------------------------------------------------------
            If _PlanInterface.ConvertCheckedouttToLife(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.HCID, CT.Data.DataCenter.FileStatus.Checkedout, Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            '----------------------------------------------------------
            ' refresh with master version or checked in version
            ' After callingChecked in function the primary keys will be changed
            ' therefore the plan must be refreshed and loaded again with main HCID.
            '-----------------------------------------------------------------------------------------------
            If Form.DataCenter.ProgramConfig.IsMainPlan = True And Form.DataCenter.ProgramConfig.HCID <> 0 And Form.DataCenter.ProgramConfig.MainPlanHCID <> 0 Then

                If _PlanDisplay.RefreshPlan(Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.IsGeneric, Form.DataCenter.ProgramConfig.IsWithCustomFormatting, Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(_PlanDisplay.ErrorMessage)
                If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then
                    Dim obj As New Form.DataCenter.ModuleFunction
                    obj.DisplayMasterMessage()
                End If
                System.Windows.Forms.MessageBox.Show("The Plan has been checked in successfuly.", "Checked in Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, strTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
        End Try



    End Sub

    Private Sub btnDiscard_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDiscard.Click


        Dim strTitle As String = "Discard Checked out plan"
        Dim _PlanDisplay As Form.DisplayUtilities.Plan = New Form.DisplayUtilities.Plan

        Try

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            '----------------------------------------------------------
            ' Take confirmation from user
            '----------------------------------------------------------
            If MessageBox.Show("Do you really want to discard the changes? ", strTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then

                '----------------------------------------------------------
                ' Keep the main HCID, IsGeneric and withCustomFormatting & Discard Plan in DB
                '----------------------------------------------------------
                If _PlanInterface.DeleteDraftOrCheckedout(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, DirectCast([Enum].Parse(GetType(CT.Data.DataCenter.FileStatus), Form.DataCenter.ProgramConfig.FileStatus), CT.Data.DataCenter.FileStatus), Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


                '----------------------------------------------------------
                ' Load Master file of the HCDID
                '----------------------------------------------------------
                If Form.DataCenter.ProgramConfig.IsMainPlan = True And Form.DataCenter.ProgramConfig.MainPlanHCID <> 0 And Form.DataCenter.ProgramConfig.HCID <> 0 Then
                    If _PlanDisplay.RefreshPlan(Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.IsGeneric, Form.DataCenter.ProgramConfig.IsWithCustomFormatting, Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(_PlanDisplay.ErrorMessage)
                    System.Windows.Forms.MessageBox.Show("The changes in plan have been discarded successfuly.", strTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                    If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then
                        Dim obj As New Form.DataCenter.ModuleFunction
                        obj.DisplayMasterMessage()
                    End If
                End If


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, strTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnCheckOut_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCheckOut.Click
        Dim result As DialogResult = MessageBox.Show("Do you really want to checkout?", "Checkout plan", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If

        Try
            Dim Answer As String = String.Empty
            Dim _PlanDisplay As New Form.DisplayUtilities.Plan()

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            '----------------------------------------------------------
            ' checkout and load checkedout plan
            '----------------------------------------------------------
            Answer = _PlanDisplay.CheckOutPlan()
            If Answer <> String.Empty Then Throw New Exception(Answer)

            '----------------------------------------------------------
            ' Activatte the Tnd Controlpanel Ribbon
            '----------------------------------------------------------
            System.Windows.Forms.MessageBox.Show("Plan has been checked out - ready to be modified.", "Check-out Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

        Catch ex As Exception
            If ex.Message.IndexOf("000:") > 0 Then
                System.Windows.Forms.MessageBox.Show(ex.Message, "Check-out", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            Else
                System.Windows.Forms.MessageBox.Show(ex.Message, "Check-out", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End If
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
        End Try
    End Sub

    Private Sub btnGenerateDraft_Click_1(sender As Object, e As RibbonControlEventArgs) Handles btnGenerateDraft.Click
        Try
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                System.Windows.Forms.MessageBox.Show("Draft option is only for 'Specific' plans.", "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim resultDataTable As New System.Data.DataTable

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            resultDataTable = _PlanInterface.SelectAllTndDraftPlans(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID)
            If resultDataTable.Rows.Count >= 3 Then Throw New Exception("3 Draft versions are already created for this HC ID : " & Form.DataCenter.ProgramConfig.HCID)

            If _PlanInterface.GenerateDraftOrCheckout(Form.DataCenter.ProgramConfig.HCID, Data.DataCenter.FileStatus.Draft, Form.DataCenter.ProgramConfig.BuildType) = True Then
                loadMenubutton()
                System.Windows.Forms.MessageBox.Show("Draft completed successfully.", "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                'Refesh the Load Draft submenu Subitem
            Else
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Generate Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try
    End Sub

    'Private Sub btnLoadDraft_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLoadDraft.Click
    '    Try
    '        Globals.ThisAddIn.Application.ScreenUpdating = False
    '        Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

    '        Dim _Plan As New Form.DisplayUtilities.Plan()
    '        Dim Answer As String = String.Empty

    '        Answer = _Plan.LoadDraftPlan(sender.label, Nothing, Form.DisplayUtilities.Plan.LoadType.Loading)
    '        If Answer <> String.Empty Then Throw New Exception(Answer)
    '        loadMenubutton()
    '        System.Windows.Forms.MessageBox.Show("The Draft Plan has been Loaded successfuly.", "Loading Draft Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
    '    Catch ex As Exception
    '        System.Windows.Forms.MessageBox.Show(ex.Message, "Load Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
    '    Finally
    '        Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
    '        Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    '    End Try
    'End Sub

    Private Sub btnDeleteDraft_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeleteDraft.Click
        Try
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

            Dim dtDraft As System.Data.DataTable

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            Dim _HCID As Integer = Form.DataCenter.ProgramConfig.HCID
            Dim _MainPlanHCID As Integer = Form.DataCenter.ProgramConfig.MainPlanHCID
            '--------------------------------------------------------------------------------
            ' Close currect not valid Draft
            '--------------------------------------------------------------------------------

            If System.Windows.Forms.MessageBox.Show("Please confirm if you really want delete the draft plan?", "Delete draft plan", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) = DialogResult.No Then
                Exit Sub
            End If


            Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)


            dtDraft = _PlanInterface.SelectTndDraftPlanDedicated(_HCID, Form.DataCenter.ProgramConfig.BuildType)
            If _PlanInterface.DeleteDraftOrCheckedout(dtDraft.Rows(0).Item("Pe01"), dtDraft.Rows(0).Item("HealthChartID"), DirectCast([Enum].Parse(GetType(CT.Data.DataCenter.FileStatus), dtDraft.Rows(0).Item("FileStatus").ToString), CT.Data.DataCenter.FileStatus), Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


            '--------------------------------------------------------------------------------
            ' Search in open workbooks to find main plan if exist then Refresh main plan
            '--------------------------------------------------------------------------------
            Dim _Worksheet As Excel.Workbook
            For i As Integer = 1 To Globals.ThisAddIn.Application.Workbooks.Count
                Try
                    _Worksheet = Globals.ThisAddIn.Application.Workbooks(i)
                    If _Worksheet.Name Like "TndTemlate*" Then

                        '--------------------------------------------------------------------------------
                        ' Check HCID 
                        '--------------------------------------------------------------------------------
                        _Worksheet.Activate()
                        _Worksheet.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()

                        If Form.DataCenter.ProgramConfig.HCID = _MainPlanHCID Then
                            Exit For
                        End If
                    End If
                Catch
                End Try
            Next

            '--------------------------------------------------------------------------------
            ' refresh draft button
            '--------------------------------------------------------------------------------
            loadMenubutton()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Delete Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try

    End Sub

    Private Sub btnReplacePlanWithDraft_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReplacePlanWithDraft.Click
        Try
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

            If System.Windows.Forms.MessageBox.Show("Please confirm if you really want the original plan to be replaced?", "Replace original with draft plan", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) = DialogResult.No Then
                Exit Sub
            End If

            Dim _MainPlanIsGeneric As Integer = Form.DataCenter.ProgramConfig.IsGeneric
            Dim _HCID As Integer = Form.DataCenter.ProgramConfig.HCID
            Dim _BuildType As String = Form.DataCenter.ProgramConfig.BuildType
            Dim _IsWithCustomFormatting As Integer = Form.DataCenter.ProgramConfig.IsWithCustomFormatting
            Dim dtDraft As System.Data.DataTable

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If _BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf _BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            Dim IsFound As Boolean = False
            Dim _ws As Excel.Workbook
            Dim ActivePlanHCID As Integer = 0


            dtDraft = _PlanInterface.SelectTndDraftPlanDedicated(Form.DataCenter.ProgramConfig.HCID, _BuildType)
            If _PlanInterface.ConvertDraftToLife(dtDraft.Rows(0).Item("Pe01"), dtDraft.Rows(0).Item("MainPlanHCID"), dtDraft.Rows(0).Item("HealthChartID"), CT.Data.DataCenter.FileStatus.Draft, _BuildType, ActivePlanHCID) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
            ''--------------------------------------------------------------------------------
            '' Close currect not valid Draft
            ''--------------------------------------------------------------------------------
            _ws = Globals.ThisAddIn.Application.ActiveWorkbook

            '--------------------------------------------------------------------------------
            ' Search in open workbooks to find Active main plan if exist then Refresh active main plan
            '--------------------------------------------------------------------------------
            Dim _Worksheet As Excel.Workbook
            For i As Integer = 1 To Globals.ThisAddIn.Application.Workbooks.Count
                Try
                    _Worksheet = Globals.ThisAddIn.Application.Workbooks(i)
                    If _Worksheet.Name Like "TndTemlate*" Then

                        '--------------------------------------------------------------------------------
                        ' Check HCID 
                        '--------------------------------------------------------------------------------
                        _Worksheet.Activate()
                        _Worksheet.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString).activate()

                        If Form.DataCenter.ProgramConfig.HCID = ActivePlanHCID Then
                            IsFound = True
                            Exit For
                        End If
                    End If
                Catch
                End Try
            Next
            '--------------------------------------------------------------------------------
            ' Refresh main plan
            '--------------------------------------------------------------------------------
            If IsFound = True Then

                Dim DisplayPlan As Form.DisplayUtilities.Plan = New Form.DisplayUtilities.Plan
                If DisplayPlan.RefreshPlan(ActivePlanHCID, _MainPlanIsGeneric, _IsWithCustomFormatting, _BuildType) = False Then Throw New Exception(DisplayPlan.ErrorMessage)

                '--------------------------------------------------------------------------------
                ' Close currect not valid Draft
                '--------------------------------------------------------------------------------
                _ws.Close(SaveChanges:=False)
            Else

                Dim DisplayPlan As Form.DisplayUtilities.Plan = New Form.DisplayUtilities.Plan
                If DisplayPlan.RefreshPlan(ActivePlanHCID, _MainPlanIsGeneric, _IsWithCustomFormatting, _BuildType) = False Then Throw New Exception(DisplayPlan.ErrorMessage)

            End If

            loadMenubutton()

            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString And Form.DataCenter.ProgramConfig.IsGeneric = False Then
                Dim obj As New Form.DataCenter.ModuleFunction
                obj.DisplayMasterMessage()
            End If

            System.Windows.Forms.MessageBox.Show("The Plan has been refreshed successfuly.", "Loading main Plan after replace", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Replace Plan With Draft", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")

            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try
    End Sub

    Private Sub btnCountReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCountReport.Click
        Try
            Dim TCReport As New TotalCountReport
            TCReport.WriteTotCntReport()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Total count report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnPustFit4Test_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPustFit4Test.Click
        Dim objDat As New CT.Data.VehiclePlan.Plan
        Try


            Dim _dt As DataTable = Nothing

            If Form.DataCenter.ProgramConfig.BuildType = "" Or Val(Form.DataCenter.ProgramConfig.HCID) = 0 Then
                Exit Sub
            End If

            Dim _PlanInterface As Data.Interfaces.PlanInterface = Nothing
            'Dim _frmFit4TestRequest As frmFit4TestRequest = New frmFit4TestRequest
            Dim _frmFit4TestRequest As Object

            '----------------------------------------------------------------------------------
            ' For the time being only the vehicle plan has Fit4Test functionality
            ' - Push 2 Fir4Test
            ' - Fit4Test validation
            ' - Fit4Test userform
            '----------------------------------------------------------------------------------
            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
                _frmFit4TestRequest = New frmFit4TestRequest
                'ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                '    _PlanInterface = New Data.RigPlan.Plan
                '    _frmFit4TestRequest = New frmFit4TestRequest_Rig
                'ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Buck.ToString() Then
                '    _PlanInterface = New Data.BuckPlan.Plan
                '    _frmFit4TestRequest = New frmFit4TestRequest_Rig 'To be modified to Buck once developed
            End If

            Dim _Result As DialogResult
            _Result = System.Windows.Forms.MessageBox.Show("Push to Fit4Test with plan validation ?", "Push to fit 4 test", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question)

            If _Result = DialogResult.Yes Then
                _dt = _PlanInterface.ValidatePlan(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus)

                If _dt.Rows.Count <> 0 Then
                    '_Result = System.Windows.Forms.MessageBox.Show("Sorry! duplicate " & Form.DataCenter.ProgramConfig.BuildType & " exists in the plan. Please remove duplicate VIN's, Vehicle number's & Prefix's. Do you still want to push this plan for fit 4 test?", "Push to fit 4 test", System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Question)

                    Dim rng As Excel.Range
                    rng = Nothing

                    With Form.DataCenter.GlobalValues.WS
                        For Each _row In _dt.Rows
                            rng = .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn.Find(_row("DisplaySeq").ToString, .Cells(4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
                            If Not rng Is Nothing Then
                                .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Interior.Color = 13487615 'rose
                            End If
                        Next
                    End With

                    System.Windows.Forms.MessageBox.Show("Sorry! duplicate " & Form.DataCenter.ProgramConfig.BuildType & " exists in the plan. Please remove duplicate VIN's, Vehicle number's & Prefix's highlighted in sheet.", "Push to fit 4 test", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

                    Exit Sub
                    'If _Result = DialogResult.Cancel Or _Result = DialogResult.No Then Exit Sub
                End If
            End If

            If objDat.PushToF4Test(Form.DataCenter.ProgramConfig.HCID) = True Then


                _frmFit4TestRequest.ShowDialog()


                'System.Windows.Forms.MessageBox.Show("Push to fit 4 test was successful!", "Push to fit 4 test", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            Else
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Push to fit 4 test", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub tmrDisplay_Tick(sender As Object, e As EventArgs) Handles tmrDisplay.Tick
        tmrDisplay.Enabled = False
        Dim objShp As Excel.Shape
        With Form.DataCenter.GlobalValues.WS
            .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Try
                For Each objShp In .Shapes
                    If objShp.Name Like "RectMasterModeDisplay*" Then
                        objShp.Delete()
                    End If
                Next
                Dim objMod As New Form.DataCenter.ModuleFunction
                objMod.sbProtectPlan()
            Catch ex As Exception
            End Try
        End With
    End Sub


    Public Sub TGMessages_Click(Optional sender As Object = Nothing, Optional e As RibbonControlEventArgs = Nothing) Handles TGMessages.Click
        Try

            If Globals.ThisAddIn.Application.ActiveWorkbook.Name.ToString Like "TndTemplate*" = False Or Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.name.ToString.ToLower <> Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                Exit Sub
            Else
                If Form.DataCenter.ProgramConfig.FileStatus <> CT.Data.DataCenter.FileStatus.Checkedout.ToString Or Form.DataCenter.GlobalValues.strUserPermissionLevel = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString Then
                    TGMessages.Enabled = False
                Else
                    TGMessages.Enabled = True
                End If
            End If
            If TGMessages.Checked Then
                Form.DataCenter.GlobalValues.wsEve.ShowMessageTaskPane(True)
            Else
                Form.DataCenter.GlobalValues.wsEve.ShowMessageTaskPane(False)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnTodayIndicator_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTodayIndicator.Click
        Dim objGlobal As New Form.DataCenter.ModuleFunction
        Try

            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            '------------------------------------------------------
            ' This code is transfered here to not to be considered in loading time.
            ' User can use it as a setting option in plan.
            '------------------------------------------------------
            '.Name = "Todaylineshape"
            'If CT.Form.DisplayUtilities.Utilities.DisplayTodayMarker() = False Then Throw New Exception(CT.Form.DisplayUtilities.Utilities.ErrorMbessage)

            Dim shp As Excel.Shape = Nothing
            For Each shp In Form.DataCenter.GlobalValues.WS.Shapes
                If shp.Name.ToString = "Todaylineshape" Then
                    shp.Delete()
                    Exit For
                End If
            Next
            If btnTodayIndicator.Checked Then
                If CT.Form.DisplayUtilities.Utilities.DisplayTodayMarker() = False Then Throw New Exception(CT.Form.DisplayUtilities.Utilities.ErrorMbessage)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Display today indicator", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            objGlobal.sbProtectPlan()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        End Try

    End Sub

    Private Sub tglBtnValidatePlan_Click(sender As Object, e As RibbonControlEventArgs) Handles tglBtnValidatePlan.Click
        Try


            Dim _dt As DataTable = Nothing
            Dim _row As DataRow, intCnt As Integer = 3
            Dim WB As Excel.Workbook = Nothing
            Dim WS As Excel.Worksheet = Nothing

            Dim _PlanInterface As Data.Interfaces.PlanInterface = Nothing

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Buck.ToString() Then
                _PlanInterface = New Data.BuckPlan.Plan
            End If

            If tglBtnValidatePlan.Checked = True Then

                If Form.DataCenter.ProgramConfig.BuildType = "" Or Val(Form.DataCenter.ProgramConfig.HCID) = 0 Then
                    Exit Sub
                End If

                _dt = _PlanInterface.ValidatePlan(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus)

                If _dt.Rows.Count = 0 Then
                    MessageBox.Show("The plan has been validated successful. No duplicate Prefix + TBNo + VIN No in CT database", "Validate Plan", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim rng As Excel.Range
                rng = Nothing

                With Form.DataCenter.GlobalValues.WS
                    For Each _row In _dt.Rows
                        rng = .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn.Find(_row("DisplaySeq").ToString, .Cells(4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
                        If Not rng Is Nothing Then
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Interior.Color = 13487615 'rose
                        End If
                    Next
                End With


            ElseIf tglBtnValidatePlan.Checked = False Then

                If Form.DataCenter.ProgramConfig.BuildType = "" Or Val(Form.DataCenter.ProgramConfig.HCID) = 0 Then
                    Exit Sub
                End If

                'Dim _PlanInterface As Data.Interfaces.PlanInterface = Nothing

                'If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                '    _PlanInterface = New Data.VehiclePlan.Plan
                'ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                '    _PlanInterface = New Data.RigPlan.Plan
                'ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Buck.ToString() Then
                '    _PlanInterface = New Data.BuckPlan.Plan
                'End If

                '_dt = _PlanInterface.ValidatePlan(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus)

                Dim rng As Excel.Range
                rng = Nothing

                With Form.DataCenter.GlobalValues.WS
                    'For Each _row In _dt.Rows
                    'rng = .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn.Find(_row("DisplaySeq").ToString, .Cells(4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
                    'If Not rng Is Nothing Then
                    For i As Int16 = 5 To Form.DataCenter.GlobalValues.TotalRow + 4
                        rng = .Cells(i, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column)
                        If rng.Interior.Color = 13487615 Then
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Interior.Color = System.Drawing.Color.White
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline)
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            .Range(.Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(rng.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
                        End If
                    Next
                    'End If
                    'Next
                End With

            End If

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Validate plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCTHelp.Click
        Dim source
        Dim HelpFilePath As String
        credentials = System.Net.CredentialCache.DefaultNetworkCredentials
        Try
            '-------- Define destination path global
            Dim DestinationPath As String = String.Format("C:\Users\{0}\ct-tool", Environment.UserName)
            HelpFilePath = DestinationPath + "\iDV CT QuickReferenceGuide 2018.pdf"

            'Check file open/exists & close
            If Check_FileExists_Open(HelpFilePath) = False Then Exit Sub

            '---------------- Download Help file from sharepoint --------------------------
            source = New Uri("https://pd3.spt.ford.com/sites/PPEteam/SiteCollectionDocuments/PPEteam/Documents/3_PROTO_TEST_PLANNING/9_TnD_Process_Documentation/iDV-ConnectedTesting/" & CT.My.Resources.CtGuideName & ".pdf")

            My.Computer.Network.DownloadFile(source, HelpFilePath, credentials, True, 60000I, True)

            '---------------- Open file --------------------------
            Process.Start(HelpFilePath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Help Document", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Checks the file is available in the destination locaiton , if exists delete 
    'If file open - then inform user the file is open
    Function Check_FileExists_Open(Filename As String) As Boolean
        Check_FileExists_Open = False
        If System.IO.File.Exists(Filename) Then
            Try
                Dim fOpen As IO.FileStream = System.IO.File.Open(Filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
                fOpen.Close()
                fOpen.Dispose()
                fOpen = Nothing
                My.Computer.FileSystem.DeleteFile(Filename, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently)
                Check_FileExists_Open = True
            Catch ex As Exception
                MessageBox.Show("Help document is already open!" & vbNewLine & CT.My.Resources.CtGuideName & ".pdf", "Help Document", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'AppActivate(CT.My.Resources.CtGuideName)
                Exit Function
            End Try
        Else
            Check_FileExists_Open = True
        End If
    End Function


End Class
