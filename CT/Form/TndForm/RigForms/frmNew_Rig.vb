Imports System.ComponentModel
Imports System.Windows.Forms

Public Class frmNew_Rig

    Private bolResize As Boolean
    Dim bolInFilter As Boolean

    Dim mydataTable, mybindTable As New System.Data.DataTable
    Dim mydataView As New System.Data.DataView

    Dim _Modfunc As New Form.DataCenter.ModuleFunction
    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions

    'Dim DontClose As Boolean = False
    Private Sub frmNew_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            Dim clsStored As CT.Data.RigPlan.Plan = New CT.Data.RigPlan.Plan
            Dim _Usercase As New CT.Data.Usercase

            cboBuildType.DataSource = [Enum].GetNames(GetType(CT.Data.DataCenter.BuildType))

            lblPlatform.Text = Form.DataCenter.ProgramConfig.Platform & "-Platform" 'to be modified to B/C/D platform (shows only pe01 id not name)
            lblBP.Text = Form.DataCenter.ProgramConfig.BuildPhase

            mydataTable = _Usercase.GetAllUsercases(BuildTypes:=Form.DataCenter.VehicleConfig.VehicleBuildType, BuildPhase:=Form.DataCenter.ProgramConfig.BuildPhase, Carline:=Form.DataCenter.ProgramConfig.Carline, Region:=Form.DataCenter.ProgramConfig.Region)

            cboBuildType.Text = Form.DataCenter.VehicleConfig.VehicleBuildType


            'Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep()
            ''_ProcessStep.SelectProcessStepDedicated(_pe26, Form.DataCenter.ProgramConfig.IsGeneric)
            ''_UsercaseSeq = _ProcessStep.AllocatedUsercaseSeq
            ''_ProcessStepSeq = _ProcessStep.ProcessStepSeq

            'Dim dtCDSID As System.Data.DataTable
            'dtCDSID = _ProcessStep.GetAllCdsids(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, _ProcessStep.GlobalDVP)

            'If Not dtCDSID Is Nothing Then
            '    If dtCDSID.Rows.Count > 0 Then
            '        For i As Int16 = 0 To dtCDSID.Rows.Count - 1
            '            cboCDSID.Items.Add(dtCDSID.Rows(i).Item(0).ToString())
            '        Next
            '    End If
            'End If
            btnFilterReset_Click(sender, e)

            cboBuildType.Text = Form.DataCenter.VehicleConfig.VehicleBuildType

            If Form.DataCenter.ProgramConfig.IsGeneric Then
                btnAdd.Enabled = False
            End If

            '---------------------------------------------------------For Build Process step - display only 'Delay' process step
            If Form.DataCenter.ProcessStepConfig.ProcessStepUserCase = "Build" Or Form.DataCenter.ProcessStepConfig.ProcessStepUserCase = "Build Cologne" Then
                For I As Integer = 0 To lstUsercase.Items.Count - 1
                    If lstUsercase.Items(I).Text = "Delay" Then
                        lstUsercase.Items(I).Selected = True
                        lstUsercase.EnsureVisible(I)
                    End If
                Next
                cboBuildType.Enabled = False
                cboCGEA.Enabled = False
                cboFuelType.Enabled = False
                cboGenSpec.Enabled = False
                cboGlobal.Enabled = False
                cboStopStart.Enabled = False
                cboTnDUser.Enabled = False
                cboTransmission.Enabled = False

                btnFilterReset.Enabled = False
                lstUsercase.Enabled = False
            End If
            '---------------------------------------------------------


            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                btnAdd.Enabled = False
            End If

            If Form.DataCenter.ProcessStepConfig.ProcessStepUserCase = "" Then chkInsertBefore.Enabled = False

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmNew, ex.Message), "Add Unit", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub sbFillUserCases()
        Try
            lblTotalDur.Text = "Total duration"
            bolInFilter = True
            Dim i As Double

            dgvProcessSteps.DataSource = Nothing
            txtRemarks.Text = ""

            Dim strOldVal(9) As String

            strOldVal(1) = cboTnDUser.Text
            strOldVal(2) = cboGlobal.Text
            strOldVal(3) = cboTransmission.Text
            strOldVal(4) = cboFuelType.Text
            strOldVal(5) = cboCGEA.Text
            strOldVal(6) = cboStopStart.Text
            strOldVal(7) = cboGenSpec.Text

            strOldVal(9) = cboBuildType.Text

            cboTnDUser.Items.Clear()
            cboGlobal.Items.Clear()
            cboTransmission.Items.Clear()
            cboFuelType.Items.Clear()
            cboCGEA.Items.Clear()
            cboStopStart.Items.Clear()
            cboGenSpec.Items.Clear()
            lstUsercase.Items.Clear()

            cboTnDUser.Items.Add("All")
            cboGlobal.Items.Add("All")
            cboTransmission.Items.Add("All")
            cboFuelType.Items.Add("All")
            cboCGEA.Items.Add("All")
            cboStopStart.Items.Add("All")
            cboGenSpec.Items.Add("All")

            mydataView = New System.Data.DataView(mydataTable)

            Dim filtercriteria As String

            filtercriteria = "BuildTypes = '" & cboBuildType.Text & "' "

            If strOldVal(1) <> "" And strOldVal(1) <> "All" Then
                filtercriteria = filtercriteria & " AND XCCTranslation='" & strOldVal(1) & "'"
            End If

            If strOldVal(2) <> "" And strOldVal(2) <> "All" Then
                filtercriteria = filtercriteria & " AND DvpTeamName='" & strOldVal(2) & "'"
            End If

            If strOldVal(3) <> "" And strOldVal(3) <> "All" Then
                filtercriteria = filtercriteria & " AND TransmissionType='" & strOldVal(3) & "'"
            End If

            If strOldVal(4) <> "" And strOldVal(4) <> "All" Then
                filtercriteria = filtercriteria & " AND FuelType='" & strOldVal(4) & "'"
            End If

            If strOldVal(5) <> "" And strOldVal(5) <> "All" Then
                filtercriteria = filtercriteria & " AND Cgea='" & strOldVal(5) & "'"
            End If

            If strOldVal(6) <> "" And strOldVal(6) <> "All" Then
                filtercriteria = filtercriteria & " AND StartStop='" & strOldVal(6) & "'"
            End If

            If strOldVal(7) <> "" And strOldVal(7) <> "All" Then
                filtercriteria = filtercriteria & " AND GenricSpecific='" & strOldVal(7) & "'"
            End If

            mydataView.RowFilter = filtercriteria

            mybindTable = mydataView.ToTable(True, "XCCTranslation")
            For i = 0 To mybindTable.Rows.Count - 1
                cboTnDUser.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "DvpTeamName")
            For i = 0 To mybindTable.Rows.Count - 1
                cboGlobal.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "TransmissionType")
            For i = 0 To mybindTable.Rows.Count - 1
                cboTransmission.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "FuelType")
            For i = 0 To mybindTable.Rows.Count - 1
                cboFuelType.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "Cgea")
            For i = 0 To mybindTable.Rows.Count - 1
                cboCGEA.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "StartStop")
            For i = 0 To mybindTable.Rows.Count - 1
                cboStopStart.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "GenricSpecific")
            For i = 0 To mybindTable.Rows.Count - 1
                cboGenSpec.Items.Add(mybindTable.Rows(i).Item(0))
            Next

            mybindTable = mydataView.ToTable(True, "Usercase", "UPMin", "UPMax", "UNMin", "UNMax", "PTMin", "PTMax")
            Dim newitem As ListViewItem
            For i = 0 To mybindTable.Rows.Count - 1
                newitem = New ListViewItem(mybindTable.Rows(i).Item(0).ToString())
                newitem.SubItems.Add(mybindTable.Rows(i).Item(1).ToString())
                newitem.SubItems.Add(mybindTable.Rows(i).Item(2).ToString())
                newitem.SubItems.Add(mybindTable.Rows(i).Item(3).ToString())
                newitem.SubItems.Add(mybindTable.Rows(i).Item(4).ToString())
                newitem.SubItems.Add(mybindTable.Rows(i).Item(5).ToString())
                newitem.SubItems.Add(mybindTable.Rows(i).Item(6).ToString())
                lstUsercase.Items.Add(newitem)
            Next

            cboTnDUser.Text = strOldVal(1)
            cboGlobal.Text = strOldVal(2)
            cboTransmission.Text = strOldVal(3)
            cboFuelType.Text = strOldVal(4)
            cboCGEA.Text = strOldVal(5)
            cboStopStart.Text = strOldVal(6)
            cboGenSpec.Text = strOldVal(7)
            cboBuildType.Text = strOldVal(9)

            If cboTransmission.Text = "" Then cboTransmission.Text = "All"
            If cboFuelType.Text = "" Then cboFuelType.Text = "All"

            btnAdd.Enabled = False

            bolInFilter = False
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmNew, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#Region "ComboBox changes"
    Private Sub cboBuildType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboBuildType.SelectedIndexChanged

    End Sub

    Private Sub cboTnDUser_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTnDUser.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub

    Private Sub cboGlobal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGlobal.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub

    Private Sub cboTransmission_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTransmission.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub

    Private Sub cboCGEA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCGEA.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub

    Private Sub cboFuelType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFuelType.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub

    Private Sub cboStopStart_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStopStart.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub
    Private Sub cboGenSpec_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGenSpec.SelectedIndexChanged
        If bolInFilter = False Then sbFillUserCases()
    End Sub

    Private Sub cboUserCase_SelectedIndexChanged(sender As Object, e As EventArgs)
        lblUPMin.Text = 0
        lblUPMax.Text = 0
        lblUNMin.Text = 0
        lblUNMax.Text = 0
        lblPTMin.Text = 0
        lblPTMax.Text = 0
        'cboUserCase.AccessibleDescription = cboUserCase.Text
        'Load UC to Listview
        Analyze_PS_Duration()
    End Sub
#End Region

#Region "Support Subs"
    Private Sub sbClearComboBoxes()
        cboCGEA.Items.Clear()
        cboFuelType.Items.Clear()
        cboGenSpec.Items.Clear()
        cboGlobal.Items.Clear()
        cboStopStart.Items.Clear()
        cboTnDUser.Items.Clear()
        cboTransmission.Items.Clear()
        lstUsercase.Items.Clear()

        txtRemarks.Text = vbNullString

        dgvProcessSteps.DataSource = Nothing
    End Sub

    Private Sub sbHandleCheckboxes(bolACtivate As Boolean)
        If bolACtivate Then
            AddHandler Me.cboTnDUser.SelectedIndexChanged, AddressOf cboTnDUser_SelectedIndexChanged
            AddHandler Me.cboGlobal.SelectedIndexChanged, AddressOf cboGlobal_SelectedIndexChanged
            AddHandler Me.cboTransmission.SelectedIndexChanged, AddressOf cboTransmission_SelectedIndexChanged
            AddHandler Me.cboCGEA.SelectedIndexChanged, AddressOf cboCGEA_SelectedIndexChanged
            AddHandler Me.cboFuelType.SelectedIndexChanged, AddressOf cboFuelType_SelectedIndexChanged
            AddHandler Me.cboStopStart.SelectedIndexChanged, AddressOf cboStopStart_SelectedIndexChanged
            AddHandler Me.lstUsercase.SelectedIndexChanged, AddressOf lstUsercase_SelectedIndexChanged
            AddHandler Me.cboGenSpec.SelectedIndexChanged, AddressOf cboGenSpec_SelectedIndexChanged
        Else
            RemoveHandler Me.cboTnDUser.SelectedIndexChanged, AddressOf cboTnDUser_SelectedIndexChanged
            RemoveHandler Me.cboGlobal.SelectedIndexChanged, AddressOf cboGlobal_SelectedIndexChanged
            RemoveHandler Me.cboTransmission.SelectedIndexChanged, AddressOf cboTransmission_SelectedIndexChanged
            RemoveHandler Me.cboCGEA.SelectedIndexChanged, AddressOf cboCGEA_SelectedIndexChanged
            RemoveHandler Me.cboFuelType.SelectedIndexChanged, AddressOf cboFuelType_SelectedIndexChanged
            RemoveHandler Me.cboStopStart.SelectedIndexChanged, AddressOf cboStopStart_SelectedIndexChanged
            RemoveHandler Me.lstUsercase.SelectedIndexChanged, AddressOf lstUsercase_SelectedIndexChanged
            RemoveHandler Me.cboGenSpec.SelectedIndexChanged, AddressOf cboGenSpec_SelectedIndexChanged
        End If
    End Sub

    Private Sub Analyze_PS_Duration()
        Dim Item As ListViewItem
        Dim lngDuration As Long

        For Each Item In dgvProcessSteps.Rows
            lngDuration = lngDuration + Convert.ToInt32(Item.SubItems(1).Text)
        Next

        lblTotalDur.Text = lngDuration & " Working Days"
    End Sub
#End Region

#Region "Button Commands"
    Public Sub ShowMe()
        Me.ShowDialog()
    End Sub
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Try
            dgvProcessSteps.EndEdit()

            If lstUsercase.SelectedItems(0).Text = "" Or dgvProcessSteps.Rows.Count <= 0 Then ' Or Strings.Trim(cboCDSID.Text) = "" Then
                DialogResult = DialogResult.None
                Throw New Exception("Please enter all data")
            End If
            If optExistingusercase.Checked = False And optSeparateusercase.Checked = False Then
                DialogResult = DialogResult.None
                Throw New Exception("Please choose new usercase option as 'separate' or 'add to existing'. ")
            End If
            If _GlobalFunctions.ContainsInvalidChar(txtRemarks.Text) Then
                DialogResult = DialogResult.None
                Throw New Exception("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ; "" ' & ; ~ ` < >.")
            End If
            If _GlobalFunctions.ContainsInvalidChar(cboCDSID.Text) Then
                DialogResult = DialogResult.None
                Throw New Exception("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ; "" ' & ; ~ ` < >.")
            End If

            'If MessageBox.Show(Me, "Do you really want to add this Usercase to database?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            '    DontClose = True
            '    Exit Sub
            'End If

            Me.Cursor = Cursors.AppStarting
            'Parisa -> This variable is not used here
            'Dim intStCol As Integer
            'intStCol = Form.DataCenter.GlobalValues.WS.Application.Selection.Column + Form.DataCenter.GlobalValues.WS.Application.Selection.Columns.Count
            'If chkInsertBefore.Checked = True Then intStCol = Form.DataCenter.GlobalValues.WS.Application.Selection.Column


            'Parisa -> This variable is not used here
            'Dim intRow As Integer
            'intRow = Form.DataCenter.GlobalValues.WS.Application.Selection.Row


            If chkInsertBefore.Checked = False Then
                Dim dblCnt As Double
                For dblCnt = 0 To dgvProcessSteps.Rows.Count - 1
                    If dgvProcessSteps.Rows(dblCnt).Cells(8).Value = True Then
                        dgvProcessSteps.Rows(dblCnt).Cells(6).Value = Form.DataCenter.ProcessStepConfig.ProcessStepEndDate
                        Exit For
                    End If
                Next
            End If



            Dim bolWasProtected As Boolean
            Dim bolIndependentusercase As Boolean = False
            If optSeparateusercase.Checked = True Then bolIndependentusercase = True
            bolWasProtected = Form.DataCenter.GlobalValues.WS.ProtectContents
            If bolWasProtected Then Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            'implemented as per Marcel request (PS = PS +1)
            Dim _Process As New Data.ProcessStep
            If _Process.Add(pe02:=Form.DataCenter.VehicleConfig.VehiclePe02, pe45:=Form.DataCenter.VehicleConfig.VehiclePe45,
                         HCID:=Form.DataCenter.VehicleConfig.VehicleHCID, AllocatedUsercaseSeq:=Form.DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase, ProcessStepSequence:=IIf(chkInsertBefore.Checked, Form.DataCenter.ProcessStepConfig.ProcessStepSequence, Form.DataCenter.ProcessStepConfig.ProcessStepSequence + 1), ProcessStepList:=dgvProcessSteps.DataSource, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType, InsertAsIndependentUsercase:=bolIndependentusercase) = False Then
                Throw New Exception("Failed: " & Data.DataCenter.GlobalValues.message)
            End If

            'Dim Cls As New Form.DataCenter.GlobalFunctions
            _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, Form.DataCenter.ProcessStepConfig.ProcessStepStartDate)

            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmNew, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnFilterReset_Click(sender As Object, e As EventArgs) Handles btnFilterReset.Click
        Try
            sbHandleCheckboxes(False)

            sbClearComboBoxes()

            lblUPMin.Text = 0
            lblUPMax.Text = 0
            lblUNMin.Text = 0
            lblUNMax.Text = 0
            lblPTMin.Text = 0
            lblPTMax.Text = 0

            cboCDSID.Text = "CDSID"

            lblTotalDur.Text = "Total duration"

            cboCGEA.Items.Add("All")
            cboFuelType.Items.Add("All")
            cboGenSpec.Items.Add("All")
            cboGlobal.Items.Add("All")
            cboStopStart.Items.Add("All")
            cboTnDUser.Items.Add("All")
            cboTransmission.Items.Add("All")

            cboCGEA.SelectedIndex = 0
            cboFuelType.SelectedIndex = 0
            cboGenSpec.SelectedIndex = 0
            cboGlobal.SelectedIndex = 0
            cboStopStart.SelectedIndex = 0
            cboTnDUser.SelectedIndex = 0
            cboTransmission.SelectedIndex = 0

            sbHandleCheckboxes(True)

            sbFillUserCases()
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmNew, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmNew_Validating(sender As Object, e As CancelEventArgs) Handles Me.Validating
        'Dim clsTest As New Form.DataCenter.GlobalFunctions
        Try
            If dgvProcessSteps.Rows.Count = 0 Then Throw New Exception("No Process Steps in Usercase.")
            'If cboCDSID.Text = vbNullString Then Throw New Exception("Missing CDSID input.")
            If _GlobalFunctions.ContainsInvalidChar(txtRemarks.Text) Then Throw New Exception("Invalid Characters ind Remark box. ["";"", ""'"", ""&"", "";"", ""~"", ""`"", ""<"", "">""]")
            If _GlobalFunctions.ContainsInvalidChar(cboCDSID.Text) Then Throw New Exception("Invalid Characters ind Remark box. ["";"", ""'"", ""&"", "";"", ""~"", ""`"", ""<"", "">""]")
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmNew, ex.Message), "Add new usercase", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            e.Cancel = True
        End Try
    End Sub

    Private Sub frmNew_Validated(sender As Object, e As EventArgs) Handles Me.Validated
        Dim bolSuccess As Boolean
        Dim clsStored As CT.Data.VehiclePlan.Plan = New CT.Data.VehiclePlan.Plan

        '@Ramesh implement updateprocedure

        bolSuccess = False 'clsStored.
    End Sub

    Private Sub lstUsercase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstUsercase.SelectedIndexChanged
        Me.Cursor = Cursors.AppStarting
        If lstUsercase.SelectedItems.Count <= 0 Then Exit Sub

        Dim dblCnt As Double
        Dim intDuration As Integer = 0

        lblUPMin.Text = Val(lstUsercase.SelectedItems(0).SubItems(1).Text)
        lblUPMax.Text = Val(lstUsercase.SelectedItems(0).SubItems(2).Text)
        lblUNMin.Text = Val(lstUsercase.SelectedItems(0).SubItems(3).Text)
        lblUNMax.Text = Val(lstUsercase.SelectedItems(0).SubItems(4).Text)
        lblPTMin.Text = Val(lstUsercase.SelectedItems(0).SubItems(5).Text)
        lblPTMax.Text = Val(lstUsercase.SelectedItems(0).SubItems(6).Text)

        mydataView = New System.Data.DataView(mydataTable)
        mydataView.RowFilter = "Usercase = '" + lstUsercase.SelectedItems(0).Text + "'"

        mybindTable = mydataView.ToTable(False, "pe39_SlotFacilityMatching_PK", "ProcessStepSequence", "ProcessStepName", "Duration", "WorkingDays")
        mybindTable.Columns.Add("CDSID", Type.GetType("System.String"))
        mybindTable.Columns.Add("PlannedStart", Type.GetType("System.String"))
        mybindTable.Columns.Add("PlannedEnd", Type.GetType("System.String"))
        mybindTable.Columns.Add("Select", Type.GetType("System.Boolean"))
        mybindTable.Columns.Add("Remarks", Type.GetType("System.String"))

        dgvProcessSteps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgvProcessSteps.DataSource = mybindTable

        dgvProcessSteps.Columns(8).DisplayIndex = 0

        dgvProcessSteps.Columns(0).Visible = False   'pe39_SlotFacilityMatching_PK
        dgvProcessSteps.Columns(6).Visible = False   'PlannedStart
        dgvProcessSteps.Columns(7).Visible = False   'PlannedEnd
        dgvProcessSteps.Columns(4).Visible = False   'WorkingDays
        dgvProcessSteps.Columns(9).Visible = False   'Remarks

        dgvProcessSteps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgvProcessSteps.Columns(8).Width = 60 'Select (Checkbox column)
        dgvProcessSteps.Columns(1).Width = 70 'ProcessStepSequence
        dgvProcessSteps.Columns(1).ReadOnly = True
        dgvProcessSteps.Columns(2).Width = 420 'ProcessStepName
        dgvProcessSteps.Columns(2).ReadOnly = True
        dgvProcessSteps.Columns(3).Width = 70 'Duration

        For dblCnt = 0 To dgvProcessSteps.Rows.Count - 1
            intDuration = intDuration + Val(dgvProcessSteps.Rows(dblCnt).DataBoundItem(3))
            dgvProcessSteps.Rows(dblCnt).Cells(6).Value = Form.DataCenter.ProcessStepConfig.ProcessStepStartDate
            dgvProcessSteps.Rows(dblCnt).Cells(5).Value = cboCDSID.Text
            dgvProcessSteps.Rows(dblCnt).Cells(9).Value = txtRemarks.Text
            dgvProcessSteps.Rows(dblCnt).Cells(8).Value = True
            'Try
            '    dgvProcessSteps.Rows(dblCnt).DataBoundItem("Select") = "True"
            'Catch ex As Exception
            'End Try
        Next
        lblTotalDur.Text = intDuration & " Working Days"
        If intDuration > 0 Then
            btnAdd.Enabled = True
        Else
            btnAdd.Enabled = False
        End If

        'load cds id's for selected usercase
        Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep()
        '_ProcessStep.SelectProcessStepDedicated(_pe26, Form.DataCenter.ProgramConfig.IsGeneric)
        '_UsercaseSeq = _ProcessStep.AllocatedUsercaseSeq
        '_ProcessStepSeq = _ProcessStep.ProcessStepSeq

        Dim dtCDSID As System.Data.DataTable
        dtCDSID = _ProcessStep.GetAllCdsids(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, mydataView.ToTable.Rows(0).Item("Dvpteamname").ToString(), Form.DataCenter.ProgramConfig.BuildType)

        With cboCDSID
            .Text = ""
            .Items.Clear()
            .Items.Add("CDSID")
            If Not dtCDSID Is Nothing Then
                If dtCDSID.Rows.Count > 0 Then
                    For i As Int16 = 0 To dtCDSID.Rows.Count - 1
                        .Items.Add(dtCDSID.Rows(i).Item(0).ToString())
                    Next
                End If
            End If
        End With

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub frmNew_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 And btnAdd.Enabled = True Then
            btnAdd_Click(sender, e)
        ElseIf e.KeyCode = Keys.F8 Then
            btnCancel_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            cboBuildType.Focus()
        End If
    End Sub

    Private Sub dgvProcessSteps_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvProcessSteps.DataError
        MessageBox.Show(e.Exception.InnerException.Message(), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        e.Cancel = True
    End Sub

    Private Sub dgvProcessSteps_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProcessSteps.CellContentClick
        dgvProcessSteps.CommitEdit(DataGridViewDataErrorContexts.Commit)
    End Sub

    'Private Sub txtCDSID_TextChanged(sender As Object, e As EventArgs)
    '    For dblCnt = 0 To dgvProcessSteps.Rows.Count - 1
    '        dgvProcessSteps.Rows(dblCnt).Cells(5).Value = txtCDSID.Text
    '    Next
    'End Sub

    Private Sub txtRemarks_TextChanged(sender As Object, e As EventArgs) Handles txtRemarks.TextChanged
        For dblCnt = 0 To dgvProcessSteps.Rows.Count - 1
            dgvProcessSteps.Rows(dblCnt).Cells(9).Value = txtRemarks.Text
        Next
    End Sub


    Private Sub dgvProcessSteps_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProcessSteps.CellValueChanged
        Dim intDuration As Integer, dblCnt As Integer
        If dgvProcessSteps.Rows.Count > 0 Then
            For dblCnt = 0 To dgvProcessSteps.Rows.Count - 1
                If dgvProcessSteps.Rows(dblCnt).DataBoundItem("Select").ToString = "True" Then
                    intDuration = intDuration + Val(dgvProcessSteps.Rows(dblCnt).DataBoundItem(3))
                End If
            Next
            lblTotalDur.Text = intDuration & " Working Days"
            If intDuration > 0 Then
                btnAdd.Enabled = True
            Else
                btnAdd.Enabled = False
            End If
        End If

        If e.ColumnIndex = 8 And e.RowIndex >= 0 Then
            For i As Int16 = 0 To dgvProcessSteps.RowCount - 1
                If dgvProcessSteps.Rows(i).DataBoundItem("Select").ToString = "True" Then
                    btnAdd.Enabled = True
                    Exit Sub
                End If
            Next
            btnAdd.Enabled = False
        End If
    End Sub

    Private Sub cboCDSID_TextChanged(sender As Object, e As EventArgs) Handles cboCDSID.TextChanged
        For dblCnt = 0 To dgvProcessSteps.Rows.Count - 1
            dgvProcessSteps.Rows(dblCnt).Cells(5).Value = cboCDSID.Text
        Next
    End Sub

    'Private Sub frmNew_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    '    If DontClose = True Then
    '        e.Cancel = True
    '        DontClose = False
    '    End If
    'End Sub
#End Region
End Class