Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Public Class frmHeaderEdit

    Dim colData As New Collection

    Private deletedUserAccessLevels As List(Of clsSaveData) = New List(Of clsSaveData)
    Private UpdatedUserAccessLevels As List(Of clsSaveData) = New List(Of clsSaveData)
    Private InsertedUserAccessLevels As List(Of clsSaveData) = New List(Of clsSaveData)
    Private TndPlannerUserAceessLevel As clsSaveData = Nothing  ' Here we keep the object of ugser access level of the tndPlanner
    Private Pe10_Owner As Integer = 0
    Private Pe27_Owner As Integer = 0

    Private AllUserAceessLevels As DataTable = Nothing

    Private _ID As Long = 0


    Private Sub btnCancel_Click(sender As Object, e As EventArgs)

        Globals.ThisAddIn.Application.EnableEvents = True
        Globals.ThisAddIn.Application.ScreenUpdating = True
        Globals.ThisAddIn.Application.DisplayAlerts = True
        Close()

    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs)

        Dim _obj As New Form.DataCenter.ModuleFunction

        Try


            If _ID <> 0 Then

                Dim _Program As New Data.ProgramConfiguration
                If _Program.Delete(_ID) = False Then Throw New Exception("Sorry, could not reset the data! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)

            Else
                MessageBox.Show("Sorry, the header of this program is already in default mode. Nothing to reset.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

            UpdateHeaderDisplay()

            _obj.sbProtectPlan()

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), "Reset Header", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Me.Close()

        End Try


    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click

        Dim _obj As New Form.DataCenter.ModuleFunction
        Try
            Me.Cursor = Cursors.WaitCursor
            '-------------------------------------------------------------
            ' Tnd Planner validation
            '-------------------------------------------------------------
            txtTndPlanner.Text = txtTndPlanner.Text.Replace(" ", "")
            If Trim(txtTndPlanner.Text) = "" Then
                MsgBox("TnD Planner cannot be blank.", vbInformation + vbOKOnly, Me.Text)
                Me.DialogResult = DialogResult.None
                Exit Sub
            End If


            Dim ErrorMessage As String = String.Empty
            ErrorMessage = PrepareDataForUserAccessLevel()
            If ErrorMessage <> String.Empty Then
                Throw New Exception(ErrorMessage)
            End If


            '-----------------------------------------------------
            ' For deleting from grid
            ' The visitor and Executor and Owner are allowed to get deleted.
            '-----------------------------------------------------
            Dim _Authorization As New CT.Data.Authorization

            For Each item In deletedUserAccessLevels
                If item.pe04_TnDProgramAuthorization_PK > 0 Then
                    If _Authorization.Delete(item.pe04_TnDProgramAuthorization_PK) = False Then
                        Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                    End If

                End If
            Next
            deletedUserAccessLevels.Clear() ' after sucessful deletion the list must get empty.

            '-----------------------------------------------------
            ' For updated user access levels in grid
            '-----------------------------------------------------
            For Each UpdatedItem In UpdatedUserAccessLevels

                If _Authorization.Update(UpdatedItem.pe04_TnDProgramAuthorization_PK, UpdatedItem.HealthChartId, UpdatedItem.pe10_SecurityLevel_FK,
                              UpdatedItem.pe27_Regions_FK, UpdatedItem.Cdsid, UpdatedItem.ProgramFunction) = False Then
                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                    Exit Sub
                End If

            Next
            UpdatedUserAccessLevels.Clear()

            '-----------------------------------------------------
            ' For Inserting new user access levels in grid
            '-----------------------------------------------------
            For Each NewItem In InsertedUserAccessLevels

                If _Authorization.Add(Form.DataCenter.ProgramConfig.BuildType, NewItem.HealthChartId, NewItem.pe10_SecurityLevel_FK, NewItem.pe27_Regions_FK, NewItem.Cdsid, NewItem.ProgramFunction) = False Then
                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                    Exit Sub
                End If

            Next
            InsertedUserAccessLevels.Clear()



            '-------------------------------------------------------------
            ' Save Changes in program config
            '-------------------------------------------------------------
            Dim _Program As New Data.ProgramConfiguration
            If _ID > 0 Then
                If _Program.Update(_ID, Form.DataCenter.ProgramConfig.HCID, txtProgDesc.Text, "Confidential", nudIssueM.Text & "." & nudIssueMin.Text, nudIssueM.Text & "." & nudIssueMin.Text, txtBuildPhase.Text, txtHardwareType.Text, txtTndPlanner.Text) = False Then

                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)

                End If
            Else
                If _Program.Add(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, txtProgDesc.Text, "Confidential", nudIssueM.Text & "." & nudIssueMin.Text, nudIssueM.Text & "." & nudIssueMin.Text, txtBuildPhase.Text, txtHardwareType.Text, txtTndPlanner.Text) = False Then

                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                End If
            End If

            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

            UpdateHeaderDisplay()


        Catch ex As Exception

            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), "Edit Header", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            Me.Cursor = Cursors.Default
            _obj.sbProtectPlan()

            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.DisplayAlerts = True

            If Me.DialogResult = DialogResult.OK Then Me.Close()

        End Try
    End Sub



    Public Sub UpdateHeaderDisplay()
        Dim _TndPlanTitle As Form.DisplayUtilities.TndPlanTitle = New Form.DisplayUtilities.TndPlanTitle
        _TndPlanTitle.LoadAndFormatLabel()
        _TndPlanTitle.FillMismatchedQty()
    End Sub

    Private Sub frmHeaderEdit_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Trim.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or
                Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                Me.Close()
                Exit Sub
            End If

            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then
                'btnReset.Enabled = False
                btnOk.Enabled = False
            End If

            Dim myDatatable As System.Data.DataTable
            Dim _Program As CT.Data.ProgramConfiguration = New CT.Data.ProgramConfiguration
            myDatatable = _Program.SelectProgramConfigs(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

            '-------------------------------------------------------------
            ' Validation for output
            '-------------------------------------------------------------
            If myDatatable Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


            If myDatatable.Rows.Count > 0 Then

                lblHCID.Text = If(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.PairedHealthChartId.ToString).ToString = String.Empty, myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.HealthChartId.ToString), myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.PairedHealthChartId.ToString))

                txtProgDesc.Text = myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.ProgramDescription.ToString)

                txtBuildPhase.Text = If(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.BuildPhases.ToString).ToString = String.Empty, Form.DataCenter.ProgramConfig.BuildPhase, myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.BuildPhases.ToString).ToString)

                nudIssueM.Text = If(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDReleaseStatus.ToString).ToString.Split(".")(0) = String.Empty, "1", myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDReleaseStatus.ToString).ToString.Split(".")(0))

                If myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDReleaseStatus.ToString).ToString.Split(".").Length > 1 Then
                    nudIssueMin.Text = myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDReleaseStatus.ToString).ToString.Split(".")(1)
                Else
                    nudIssueMin.Text = "0"
                End If

                txtHardwareType.Text = If(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.BuildTypes.ToString).ToString = String.Empty, Form.DataCenter.ProgramConfig.BuildType, myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.BuildTypes.ToString).ToString)

                _ID = Long.Parse(Val(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.pe78_TnDProgramConfig_PK.ToString).ToString))

                txtTndPlanner.Text = myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDPlanner.ToString).ToString

                '-------------------------------------------------------------
                ' The user access level information of Owner/TnDPlanner is saved as an object
                '-------------------------------------------------------------
                If IsDBNull(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.pe04_TnDProgramAuthorization_PK.ToString)) = False Then
                    TndPlannerUserAceessLevel = New clsSaveData(CInt(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.pe04_TnDProgramAuthorization_PK.ToString)), Form.DataCenter.ProgramConfig.HCID, CInt(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.pe10_SecurityLevel_FK.ToString)), "Plan Owner", CInt(myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.pe27_Regions_FK.ToString)), myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDPlanner.ToString).ToString)
                End If

            Else

                _ID = 0

                lblHCID.Text = Form.DataCenter.ProgramConfig.HCID.ToString()

                txtProgDesc.Text = Form.DataCenter.ProgramConfig.HCIDName

                txtBuildPhase.Text = Form.DataCenter.ProgramConfig.BuildPhase

                nudIssueM.Text = "1"

                nudIssueMin.Text = "0"

                txtHardwareType.Text = Form.DataCenter.ProgramConfig.BuildType

                '-------------------------------------------------------------
                'If TnDPlanner doesn't exist then the object is set to Nothing
                '-------------------------------------------------------------
                TndPlannerUserAceessLevel = Nothing

            End If

            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                btnOk.Enabled = False
            End If

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            gridRefresh()
        End Try

    End Sub
    Private Sub gridRefresh()

        Try


            btnSave.BackColor = System.Drawing.SystemColors.ButtonFace
            Dim clsAutho As New CT.Data.Authorization
            AllUserAceessLevels = clsAutho.SelectAll(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID)
            grdPermissions.AutoGenerateColumns = False

            '-------------------------------------------------------------
            'The Owner is not displayed in the list
            '-------------------------------------------------------------
            Dim Viewobj As DataView
            Viewobj = New System.Data.DataView(AllUserAceessLevels)
            Viewobj.RowFilter = "SecurityLevel <> 'Owner'"

            Dim dt As DataTable
            dt = Viewobj.ToTable

            grdPermissions.DataSource = dt
            'grdPermissions.Refresh()

            Dim clsSecLevel As New CT.Data.SecurityLevel
            Dim dt2 As System.Data.DataTable = clsSecLevel.SelectAll()

            '-------------------------------------------------------------
            ' Keep Id of the Owner for further usage
            '-------------------------------------------------------------
            Dim results As System.Data.DataRow() = dt2.Select("SecurityLevel = 'Owner'")
            If results.Length = 1 Then
                Pe10_Owner = results(0)("pe10_SecurityLevel_PK")
            End If
            If Pe10_Owner = 0 Then Throw New Exception("The Owner doesn't exist in DB.")


            '-------------------------------------------------------------
            'The Owner item is not displayed and will be controlled with TnDPlanner
            '-------------------------------------------------------------
            Dim SecurityLevelView As System.Data.DataView = New System.Data.DataView(dt2)
            SecurityLevelView.RowFilter = "SecurityLevel <> 'Owner'"

            Dim GridCmbColumn As DataGridViewComboBoxColumn = CType(grdPermissions.Columns("cmbSecurityLevel"), DataGridViewComboBoxColumn)
            GridCmbColumn.DataSource = SecurityLevelView
            GridCmbColumn.DisplayMember = "SecurityLevel"
            GridCmbColumn.ValueMember = "pe10_SecurityLevel_PK"


            Dim clsRegion As New CT.Data.Region
            Dim dt3 As System.Data.DataTable = clsRegion.SelectAll

            '-------------------------------------------------------------
            ' Keep Id of the region of plan for further usage
            '-------------------------------------------------------------
            results = Nothing
            results = dt3.Select("Regions = '" + CT.Form.DataCenter.ProgramConfig.Region + "'")
            If results.Length = 1 Then
                Pe27_Owner = results(0)("pe27_Regions_PK")
            End If
            If Pe27_Owner = 0 Then Throw New Exception("The Region doesn't exist in DB.")


            Dim GridCmbColumn2 As DataGridViewComboBoxColumn = CType(grdPermissions.Columns("Regions"), DataGridViewComboBoxColumn)
            GridCmbColumn2.DataSource = dt3
            GridCmbColumn2.DisplayMember = "Regions"
            GridCmbColumn2.ValueMember = "pe27_Regions_PK"


            Dim GridBtnColumn As DataGridViewButtonColumn = CType(grdPermissions.Columns("btnDelete"), DataGridViewButtonColumn)
            GridBtnColumn.Text = "X"
            GridBtnColumn.UseColumnTextForButtonValue = True

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub frmHeaderEdit_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            btnOk_Click(sender, e)
        ElseIf e.KeyCode = Keys.F8 Then
            btnReset_Click(sender, e)
        End If
    End Sub

    Private Sub btnCancel2_Click(sender As Object, e As EventArgs) Handles btnCancel2.Click
        Me.Close()
    End Sub




    Private Function PrepareDataForUserAccessLevel() As String
        Try
            Dim intCnt As Integer = 0
            Dim colDuplicate As New Collection

            '-------------------------------------------------------------
            ' TnDPlanner Validation
            '-------------------------------------------------------------
            If ReplaceDBNUll(txtTndPlanner.Text, False) = "" Then

                MessageBox.Show("Plan should have an Owner.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Throw New Exception("Plan should have an Owner.")

            Else
                '-------------------------------------------------------------
                ' Correct the TnDpLanner
                '-------------------------------------------------------------
                txtTndPlanner.Text = txtTndPlanner.Text.Replace(" ", "").Trim()
            End If

            '-------------------------------------------------------------
            ' Manage owners
            '-------------------------------------------------------------
            Dim results As List(Of DataRow) = AllUserAceessLevels.Select(" SecurityLevel = 'Owner'").ToList
            If results.Count > 1 Then
                '-------------------------------------------------------------
                ' Remove all except one
                '-------------------------------------------------------------
                Dim i As Integer = 0
                While results.Count > 1
                    deletedUserAccessLevels.Add(New clsSaveData(CInt(results(i)("pe04_TnDProgramAuthorization_PK"))))
                    results.RemoveAt(i)
                    i = i - 1
                End While

                If TndPlannerUserAceessLevel Is Nothing Then
                    TndPlannerUserAceessLevel = New clsSaveData(CInt(results(0)("pe04_TnDProgramAuthorization_PK").ToString), CT.Form.DataCenter.ProgramConfig.HCID, Pe10_Owner, "New Owner", Pe27_Owner, txtTndPlanner.Text)
                Else
                    TndPlannerUserAceessLevel.Cdsid = txtTndPlanner.Text
                End If

                UpdatedUserAccessLevels.Add(TndPlannerUserAceessLevel)
            ElseIf results.Count = 1 Then
                '-------------------------------------------------------------
                ' the object is available
                '-------------------------------------------------------------
                If TndPlannerUserAceessLevel Is Nothing Then

                    TndPlannerUserAceessLevel = New clsSaveData(0, CT.Form.DataCenter.ProgramConfig.HCID, Pe10_Owner, "New Owner", Pe27_Owner, txtTndPlanner.Text)
                    InsertedUserAccessLevels.Add(TndPlannerUserAceessLevel)

                Else
                    TndPlannerUserAceessLevel.Cdsid = txtTndPlanner.Text
                    TndPlannerUserAceessLevel.pe10_SecurityLevel_FK = Pe10_Owner
                    TndPlannerUserAceessLevel.ProgramFunction = "New Owner"
                    UpdatedUserAccessLevels.Add(TndPlannerUserAceessLevel)
                End If
            ElseIf results.Count = 0 Then
                '-------------------------------------------------------------
                ' the object is not available
                '-------------------------------------------------------------
                TndPlannerUserAceessLevel = New clsSaveData(0, CT.Form.DataCenter.ProgramConfig.HCID, Pe10_Owner, "New Owner", Pe27_Owner, txtTndPlanner.Text)
                InsertedUserAccessLevels.Add(TndPlannerUserAceessLevel)
            End If


            With grdPermissions
                '-----------------------------------------------------
                ' Set color to white because the error is displayed with red
                '-----------------------------------------------------
                For intCnt = 0 To .Rows.Count - 2
                    .Rows(intCnt).Cells("CDSID").Style.BackColor = System.Drawing.Color.White
                    .Rows(intCnt).Cells("cmbSecurityLevel").Style.BackColor = System.Drawing.Color.White
                    .Rows(intCnt).Cells("Regions").Style.BackColor = System.Drawing.Color.White
                    .Rows(intCnt).Cells("ProgramFunction").Style.BackColor = System.Drawing.Color.White
                Next

                '-----------------------------------------------------
                ' validation
                '-----------------------------------------------------
                intCnt = .Rows.Count - 2
                While intCnt >= 0

                    If ReplaceDBNUll(.Rows(intCnt).Cells("CDSID").Value, False) = "" Then
                        .Rows(intCnt).Cells("CDSID").Selected = True
                        .Rows(intCnt).Cells("CDSID").Style.BackColor = System.Drawing.Color.Red
                        MessageBox.Show("The CDSID cannot be empty. Please enter the CDSID", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Throw New Exception("The CDSID cannot be empty. Please enter the CDSID")
                    End If

                    If Not colDuplicate.Contains(ReplaceDBNUll(.Rows(intCnt).Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", "")) Then
                        colDuplicate.Add("", ReplaceDBNUll(.Rows(intCnt).Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", ""))
                    Else
                        .Rows(intCnt).Cells("CDSID").Selected = True
                        .Rows(intCnt).Cells("CDSID").Style.BackColor = System.Drawing.Color.Red
                        MessageBox.Show("There are duplicate CDSID's. Please enter unique CDSID's", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Throw New Exception("There are duplicate CDSID's. Please enter unique CDSID's")
                    End If

                    If Val(ReplaceDBNUll(.Rows(intCnt).Cells("cmbSecurityLevel").Value, True)) = 0 Then
                        .Rows(intCnt).Cells("cmbSecurityLevel").Selected = True
                        .Rows(intCnt).Cells("cmbSecurityLevel").Style.BackColor = System.Drawing.Color.Red
                        MessageBox.Show("The security level cannot be empty. Please enter the security level", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Throw New Exception("The security level cannot be empty. Please enter the security level")
                    End If

                    If Val(ReplaceDBNUll(.Rows(intCnt).Cells("Regions").Value, True)) = 0 Then
                        .Rows(intCnt).Cells("Regions").Selected = True
                        .Rows(intCnt).Cells("Regions").Style.BackColor = System.Drawing.Color.Red
                        MessageBox.Show("The region cannot be empty. Please enter the region", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Throw New Exception("The region cannot be empty. Please enter the region")
                    End If

                    If ReplaceDBNUll(.Rows(intCnt).Cells("ProgramFunction").Value, False) = "" Then
                        .Rows(intCnt).Cells("ProgramFunction").Selected = True
                        .Rows(intCnt).Cells("ProgramFunction").Style.BackColor = System.Drawing.Color.Red
                        MessageBox.Show("The Program Function cannot be empty. Please enter the Program Function", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Throw New Exception("The Program Function cannot be empty. Please enter the Program Function")
                    End If

                    intCnt -= 1
                End While

                '-----------------------------------------------------
                ' List of updated values and inserted values
                '-----------------------------------------------------

                For Each dr As DataGridViewRow In .Rows

                    If dr.IsNewRow Then Exit For


                    '-------------------------------------------------------------
                    ' Duplicate CDSID validation
                    '-------------------------------------------------------------
                    Dim UpdateResult = UpdatedUserAccessLevels.Where(Function(x) x.Cdsid.ToString.ToLower = ReplaceDBNUll(dr.Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", ""))
                    Dim InsertResult = InsertedUserAccessLevels.Where(Function(x) x.Cdsid.ToString.ToLower = ReplaceDBNUll(dr.Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", ""))

                    If ReplaceDBNUll(dr.Cells("pe04_TnDProgramAuthorization_PK").Value, True) > 0 Then ' If pe04 exists then this is an item for updating


                        If ReplaceDBNUll(dr.Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", "") = txtTndPlanner.Text.ToLower Then
                            '-------------------------------------------------------------
                            ' Manage TndPlaner CDSID in grid
                            '-------------------------------------------------------------
                            deletedUserAccessLevels.Add(New clsSaveData(CInt(dr.Cells("pe04_TnDProgramAuthorization_PK").Value)))


                        ElseIf UpdateResult.ToList().Count = 0 And InsertResult.ToList.Count = 0 And ReplaceDBNUll(dr.Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", "") <> "" Then
                            '-------------------------------------------------------------
                            ' If value is not a duplicate value then add it to list
                            '-------------------------------------------------------------
                            UpdatedUserAccessLevels.Add(New clsSaveData(ReplaceDBNUll(dr.Cells("pe04_TnDProgramAuthorization_PK").Value, True),
                                                                    Form.DataCenter.ProgramConfig.HCID,
                                                                    CInt(ReplaceDBNUll(dr.Cells("cmbSecurityLevel").Value, True)),
                                                                    dr.Cells("ProgramFunction").Value,
                                                                    CInt(ReplaceDBNUll(dr.Cells("Regions").Value, True)),
                                                                    dr.Cells("CDSID").Value.ToString()))
                        End If

                    ElseIf ReplaceDBNUll(dr.Cells("pe04_TnDProgramAuthorization_PK").Value, True) = 0 Then ' if pe04 is 0 then it's new item for inserting



                        If ReplaceDBNUll(dr.Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", "") = txtTndPlanner.Text.ToLower Then
                            '-------------------------------------------------------------
                            ' Manage TndPlaner CDSID in grid
                            '-------------------------------------------------------------
                            'DO NOTHING -> this must not be considered


                        ElseIf UpdateResult.ToList().Count = 0 And InsertResult.ToList.Count = 0 And ReplaceDBNUll(dr.Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", "") <> "" And dr.Cells("CDSID").Value.ToString.ToLower.Replace(" ", "") <> txtTndPlanner.Text.ToLower Then
                            '-------------------------------------------------------------
                            ' If value is not a duplicate value then add it to list
                            '-------------------------------------------------------------
                            InsertedUserAccessLevels.Add(New clsSaveData(0,
                                                                         Form.DataCenter.ProgramConfig.HCID,
                                                                          CInt(ReplaceDBNUll(dr.Cells("cmbSecurityLevel").Value, True)),
                                                                         dr.Cells("ProgramFunction").Value,
                                                                         CInt(ReplaceDBNUll(dr.Cells("Regions").Value, True)),
                                                                         dr.Cells("CDSID").Value.ToString))
                        End If


                    End If
                    intCnt = intCnt - 1

                Next

            End With
            PrepareDataForUserAccessLevel = String.Empty
        Catch ex As Exception
            PrepareDataForUserAccessLevel = ex.Message
        End Try

    End Function



    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim ErrorMessage As String = String.Empty
            ErrorMessage = PrepareDataForUserAccessLevel()
            If ErrorMessage <> String.Empty Then
                Throw New Exception(ErrorMessage)
            End If


            '-----------------------------------------------------
            ' For deleting from grid
            ' The visitor and Executor and Owner are allowed to get deleted.
            '-----------------------------------------------------
            Dim _Authorization As New CT.Data.Authorization

            For Each item In deletedUserAccessLevels
                If item.pe04_TnDProgramAuthorization_PK > 0 Then
                    If _Authorization.Delete(item.pe04_TnDProgramAuthorization_PK) = False Then
                        Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                    End If

                End If
            Next
            deletedUserAccessLevels.Clear() ' after sucessful deletion the list must get empty.

            '-----------------------------------------------------
            ' For updated user access levels in grid
            '-----------------------------------------------------
            For Each UpdatedItem In UpdatedUserAccessLevels

                If _Authorization.Update(UpdatedItem.pe04_TnDProgramAuthorization_PK, UpdatedItem.HealthChartId, UpdatedItem.pe10_SecurityLevel_FK,
                              UpdatedItem.pe27_Regions_FK, UpdatedItem.Cdsid, UpdatedItem.ProgramFunction) = False Then
                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                    Exit Sub
                End If

            Next
            UpdatedUserAccessLevels.Clear()

            '-----------------------------------------------------
            ' For Inserting new user access levels in grid
            '-----------------------------------------------------
            For Each NewItem In InsertedUserAccessLevels

                If _Authorization.Add(Form.DataCenter.ProgramConfig.BuildType, NewItem.HealthChartId, NewItem.pe10_SecurityLevel_FK, NewItem.pe27_Regions_FK, NewItem.Cdsid, NewItem.ProgramFunction) = False Then
                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                    Exit Sub
                End If

            Next
            InsertedUserAccessLevels.Clear()

            '-------------------------------------------------------------
            ' Save Changes in program config
            '-------------------------------------------------------------
            Dim _Program As New Data.ProgramConfiguration
            If _ID > 0 Then
                If _Program.Update(_ID, Form.DataCenter.ProgramConfig.HCID, txtProgDesc.Text, "Confidential", nudIssueM.Text & "." & nudIssueMin.Text, nudIssueM.Text & "." & nudIssueMin.Text, txtBuildPhase.Text, txtHardwareType.Text, txtTndPlanner.Text) = False Then

                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)

                End If
            Else
                If _Program.Add(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, txtProgDesc.Text, "Confidential", nudIssueM.Text & "." & nudIssueMin.Text, nudIssueM.Text & "." & nudIssueMin.Text, txtBuildPhase.Text, txtHardwareType.Text, txtTndPlanner.Text) = False Then

                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                End If
            End If

            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

            UpdateHeaderDisplay()


            gridRefresh()
            MessageBox.Show("Data saved sucessfully!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            gridRefresh()
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            Me.Cursor = Cursors.Default

            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub grdPermissions_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles grdPermissions.CellEndEdit

        Try
            Dim colDuplicate As New Collection
            Dim intCnt As Integer

            With grdPermissions

                For intCnt = 0 To .Rows.Count - 2
                    .Rows(intCnt).Cells("CDSID").Style.BackColor = System.Drawing.Color.White
                Next

                For intCnt = 0 To .Rows.Count - 2
                    If Not colDuplicate.Contains(ReplaceDBNUll(.Rows(intCnt).Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", "")) Then
                        colDuplicate.Add("", ReplaceDBNUll(.Rows(intCnt).Cells("CDSID").Value, False).ToString.ToLower.Replace(" ", ""))
                    Else
                        .Rows(intCnt).Cells("CDSID").Selected = True
                        .Rows(intCnt).Cells("CDSID").Style.BackColor = System.Drawing.Color.Red
                        .Rows(intCnt).Cells("CDSID").Value = ""
                        MessageBox.Show("There are duplicate CDSID's. Please enter unique CDSID's", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Next
            End With

            Dim objcls As New clsSaveData(Val(ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("pe04_TnDProgramAuthorization_PK").Value, True)))
            With objcls
                .HealthChartId = Val(Form.DataCenter.ProgramConfig.HCID)
                .pe10_SecurityLevel_FK = Val(ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("cmbSecurityLevel").Value, True))
                .pe27_Regions_FK = Val(ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("Regions").Value, True))
                .ProgramFunction = ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("ProgramFunction").Value, False)
                .Cdsid = ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("CDSID").Value, False)
            End With

            If ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("CDSID").Value, False).ToString.Trim <> "" Then
                If colData.Contains(ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("CDSID").Value, False).ToString) Then
                    colData.Remove(ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("CDSID").Value, False).ToString)
                    colData.Add(objcls, ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("CDSID").Value, False).ToString)
                    btnSave.BackColor = System.Drawing.Color.Red
                Else
                    colData.Add(objcls, ReplaceDBNUll(grdPermissions.Rows(e.RowIndex).Cells("CDSID").Value, False).ToString)
                    btnSave.BackColor = System.Drawing.Color.Red
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Class clsSaveData

        Public Sub New(pe04_PK As Integer)
            pe04_TnDProgramAuthorization_PK = pe04_PK
        End Sub

        Public Sub New(pe04_PK As Integer, IntHCID As Integer, IntPe10_SecurityLevel_FK As Integer, StrProgramFunction As String, IntPe27_Regions_FK As Integer, StrCDSID As String)
            pe04_TnDProgramAuthorization_PK = pe04_PK
            HealthChartId = IntHCID
            pe10_SecurityLevel_FK = IntPe10_SecurityLevel_FK
            ProgramFunction = StrProgramFunction
            pe27_Regions_FK = IntPe27_Regions_FK
            Cdsid = StrCDSID
        End Sub


        Public pe04_TnDProgramAuthorization_PK As Integer = 0
        Public HealthChartId As Integer = 0
        Public pe10_SecurityLevel_FK As Integer = 0
        Public ProgramFunction As String = ""
        Public pe27_Regions_FK As Integer = 0
        Public Cdsid As String = ""
    End Class
    Private Sub btnReset2_Click(sender As Object, e As EventArgs) Handles btnReset2.Click
        gridRefresh()
        MessageBox.Show("Reset complete!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Function ReplaceDBNUll(ByRef objValue As Object, IsInteger As Boolean) As Object
        If IsDBNull(objValue) Then
            If IsInteger Then
                Return 0
            ElseIf IsDate(objValue) Then
                Return Nothing
            Else
                Return ""
            End If
        Else
            Return objValue
        End If
    End Function

    Private Sub grdPermissions_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles grdPermissions.UserDeletingRow
        'Try
        '    Dim objcls As New clsSaveData
        '    With objcls
        '        .pe04_TnDProgramAuthorization_PK = ReplaceDBNUll(e.Row.Cells("pe04_TnDProgramAuthorization_PK").Value, True)
        '        If .pe04_TnDProgramAuthorization_PK > 0 Then
        '            colData_Delete.Add(objcls)
        '            btnSave.BackColor = System.Drawing.Color.Red
        '        ElseIf colData.Contains(ReplaceDBNUll(e.Row.Cells("CDSID").Value, False).ToString) Then
        '            colData.Remove(ReplaceDBNUll(e.Row.Cells("CDSID").Value, False).ToString)
        '            btnSave.BackColor = System.Drawing.Color.Red
        '        End If
        '    End With
        'Catch ex As Exception
        '    MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        DeleteGridRow(e.Row.Index)
    End Sub

    Sub DeleteGridRow(RowNo As Integer)
        '-----------------------------------------------------------------------
        ' Check the pe04 in grid and if it's not without HCID then 
        ' it's saved to deleted item collections
        '-----------------------------------------------------------------------
        Try
            If ReplaceDBNUll(grdPermissions.Rows(RowNo).Cells("pe04_TnDProgramAuthorization_PK").Value, True) > 0 Then

                deletedUserAccessLevels.Add(New clsSaveData(ReplaceDBNUll(grdPermissions.Rows(RowNo).Cells("pe04_TnDProgramAuthorization_PK").Value, True)))
                btnSave.BackColor = System.Drawing.Color.Red

            ElseIf colData.Contains(ReplaceDBNUll(grdPermissions.Rows(RowNo).Cells("CDSID").Value, False).ToString) Then

                colData.Remove(ReplaceDBNUll(grdPermissions.Rows(RowNo).Cells("CDSID").Value, False).ToString)
                btnSave.BackColor = System.Drawing.Color.Red

            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHeaderEdit, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grdPermissions_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdPermissions.CellContentClick
        Dim colName As String = grdPermissions.Columns(e.ColumnIndex).Name
        If colName = "btnDelete" And e.RowIndex <> grdPermissions.Rows.Count - 1 Then
            DeleteGridRow(e.RowIndex)
            grdPermissions.Rows.RemoveAt(e.RowIndex)
        End If
    End Sub

    Private Sub grdPermissions_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles grdPermissions.DataError
        MsgBox(sender.ToString)
    End Sub

    Private Sub frmHeaderEdit_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Globals.ThisAddIn.Application.EnableEvents = True
    End Sub

    Private Sub btnCancel_Click_1(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim _obj As New Form.DataCenter.ModuleFunction
        Try
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

            UpdateHeaderDisplay()

        Catch ex As Exception
        Finally
            _obj.sbProtectPlan()
        End Try
    End Sub
End Class