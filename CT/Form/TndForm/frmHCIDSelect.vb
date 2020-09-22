Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.Threading
Imports System.Globalization
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmHCIDSelect
    Dim myDataTable As System.Data.DataTable
    Dim mySource As New BindingSource()
    Dim mySourceGeneric As New BindingSource()
    Dim frmProgress As New frmProgressbar
    Private WithEvents _myLibrary As New Form.DisplayUtilities.Plan
    Dim bol_HeaderCol_Clicked As Boolean
    Public PrerequisitesFulfilled As Boolean = False 'For controlling the return value of precheck button.


    Private _ErrorMessage As String = String.Empty    'For validating the functions and return value ot Ribbon
    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _ErrorMessage
        End Get
    End Property

    Private _CurrentUserStatus As String = String.Empty




    Dim strMainBuildType As String = String.Empty 'For keeping the status of selected MainBuildType


    Private Function Fill_grdPlans() As String
        Dim _Plan As CT.Data.Interfaces.PlanInterface = Nothing  ' using interface instead of defining different variables

        Try

            Fill_grdPlans = String.Empty

            '--------------------------------------------------------------------------
            ' Setting for generic plans
            '--------------------------------------------------------------------------
            grdPlans.AutoGenerateColumns = False
            grdPlans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


            '-------------------------------------------------------------------
            ' assigning the correspondence classes
            '-------------------------------------------------------------------
            If strMainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                _Plan = New Data.VehiclePlan.Plan()
            ElseIf strMainBuildType = CT.Data.DataCenter.BuildType.Buck.ToString Then
                _Plan = New Data.BuckPlan.Plan()
            ElseIf strMainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                _Plan = New Data.RigPlan.Plan()
            End If

            '-------------------------------------------------------------------
            ' Fetch value and validation
            '-------------------------------------------------------------------
            myDataTable = _Plan.SelectAllSpecificTndPlans()
            If myDataTable Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            Dim myView As System.Data.DataView = New System.Data.DataView(myDataTable)
            mySource.DataSource = myView  ' Keep the source for filtering later
            grdPlans.DataSource = myView  ' Keep the source for filtering later
            grdPlans.Refresh()

            '-------------------------------------------------------------------
            ' In specific grid all the rows are red
            '-------------------------------------------------------------------
            grdPlans.Columns("GenOrSpec").DefaultCellStyle.ForeColor = System.Drawing.Color.Red

        Catch ex As Exception
            grdPlans.Enabled = False
            Fill_grdPlans = ex.Message
        End Try
    End Function

    Private Function Fill_grdPlansGeneric() As String
        Dim _Plan As CT.Data.Interfaces.PlanInterface = Nothing  ' using interface instead of defining different variables

        Try

            Fill_grdPlansGeneric = String.Empty

            '--------------------------------------------------------------------------
            ' Setting for generic plans
            '--------------------------------------------------------------------------
            grdPlansGeneric.AutoGenerateColumns = False
            grdPlansGeneric.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


            '-------------------------------------------------------------------
            ' assigning the correspondence classes
            '-------------------------------------------------------------------
            If strMainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                _Plan = New Data.VehiclePlan.Plan()
            ElseIf strMainBuildType = CT.Data.DataCenter.BuildType.Buck.ToString Then
                _Plan = New Data.BuckPlan.Plan()
            ElseIf strMainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                _Plan = New Data.RigPlan.Plan()
            End If

            '-------------------------------------------------------------------
            ' Fetch and validate
            '-------------------------------------------------------------------
            myDataTable = _Plan.SelectAllGenericTndPlan
            If myDataTable Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            Dim myView As System.Data.DataView = New System.Data.DataView(myDataTable)
            mySourceGeneric.DataSource = myView ' Keep for filtering later
            grdPlansGeneric.DataSource = myView ' Keep for filtering later
            grdPlansGeneric.Refresh()

            '-------------------------------------------------------------------
            ' All the rows are blue in generuic
            '-------------------------------------------------------------------
            grdPlansGeneric.Columns("GIsGeneric").DefaultCellStyle.ForeColor = System.Drawing.Color.Blue

        Catch ex As Exception
            grdPlansGeneric.Enabled = False
            lblXCCDBStatus.Visible = True
            Fill_grdPlansGeneric = ex.Message
        End Try

    End Function


    'ProgressBar update
    Public Sub UpdateProgressBar(intprogressvalue As Double) Handles _myLibrary.EventUpdateProgress
        Try
            Form.DataCenter.GlobalValues.WS.Parent.activate
        Catch ex As Exception
        End Try
        If Me.Visible = False Then
            frmProgress.UpdateProgressBar(intprogressvalue)
            If btnOpenLoad.Tag = "" Then
                frmProgress.Text = "Refreshing plan : " & CInt(frmProgress.SmoothProgressBar2.Value) & "% completed."
            Else
                frmProgress.Text = "Loading plan : " & CInt(frmProgress.SmoothProgressBar2.Value) & "% completed."
            End If
        Else
            If (Me.SmoothProgressBar1.Value > 0) Then
                Me.SmoothProgressBar1.Value -= intprogressvalue
                Me.SmoothProgressBar2.Value += intprogressvalue
                Me.lblProgress.Text = CInt(Me.SmoothProgressBar2.Value) & "% loading"
                Me.lblProgress.Refresh()
                Me.SmoothProgressBar1.Refresh()
                Me.SmoothProgressBar2.Refresh()
            End If
        End If
    End Sub


    ''' <summary>
    ''' 'Validation 'Generic Build phase, MR Date/PEC/FEC/Build scle qty
    ''' </summary>
    ''' <returns></returns>
    Function ValidateGridData() As Boolean
        Try
            If TabController1.SelectedIndex = 1 Then
                If grdPlansGeneric.SelectedRows(0).Cells(2).Value = "Generic" Then

                    If grdPlansGeneric.SelectedRows(0).Cells(6).Value.ToString = "" Then Throw New Exception("MR Data cannot be blank.")
                    If Val(grdPlansGeneric.SelectedRows(0).Cells(7).Value) = "0" Then Throw New Exception("Qty cannot be blank.")
                    'If Val(grdPlansGeneric.SelectedRows(0).Cells(8).Value) = "0" Then Throw New Exception("Assy build scale cannot be blank.")
                    Select Case grdPlansGeneric.SelectedRows(0).Cells(5).Value
                        Case CT.Data.DataCenter.BuildPhase.VP.ToString, CT.Data.DataCenter.BuildPhase.DCV.ToString
                            If grdPlansGeneric.SelectedRows(0).Cells(10).Value.ToString = "" Then Throw New Exception("PEC cannot be blank.")
                            If grdPlansGeneric.SelectedRows(0).Cells(11).Value.ToString = "" Then Throw New Exception("FEC cannot be blank.")
                        Case CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.X0.ToString, CT.Data.DataCenter.BuildPhase.X1.ToString, CT.Data.DataCenter.BuildPhase.XM.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString
                            If grdPlansGeneric.SelectedRows(0).Cells(9).Value.ToString = "" Then Throw New Exception("M1DC cannot be blank.")
                    End Select
                End If
                ValidateGridData = True
            Else
                ValidateGridData = False
            End If
        Catch ex As Exception
            ValidateGridData = False
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHCIDSelect, ex.Message) & " Plan cannot be loaded.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


    Private Sub DeactiveControls()
        '-----------------------------------------------------
        ' Deactive buttons
        '-----------------------------------------------------
        btnOpenLoad.Enabled = False
        btnCheckout.Enabled = False
        btnDraft.Enabled = False

        grdPlans.Enabled = False
        grdPlans.ReadOnly = True

        grdPlansGeneric.Enabled = False
        grdPlansGeneric.ReadOnly = True

        txtHcid.ReadOnly = True
        txtHCName.ReadOnly = True
        btnCancel.Enabled = False
        btnPreCheck.Enabled = False
        chkLoadIndFormat.Enabled = False

    End Sub


    'Plan load on button load click event
    Public Sub btnOpenLoad_Click(sender As Object, e As EventArgs) Handles btnOpenLoad.Click

        Dim BuildPhase As String = String.Empty
        Dim BuildScale As String = String.Empty

        Dim HCID As Integer
        Dim IsGeneric As Boolean
        Dim _PlanForLoading As Form.DisplayUtilities.Plan = New Form.DisplayUtilities.Plan(Me)
        Dim bolWasOn As Boolean

        '------------------------------------------------------------
        ' To force user to press pre-check first for generic plans and don't load generic plan if data missed.
        '------------------------------------------------------------
        If TabController1.SelectedIndex = 1 Then
            If PrerequisitesFulfilled = False Then
                btnOpenLoad.Enabled = PrerequisitesFulfilled
                Exit Sub
            End If
        End If


        ''-----------------------------------------------------------
        '' Pre-check the current user if he has opened a plan already
        ''-----------------------------------------------------------
        'If _CurrentUserStatus = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString Then

        '    MessageBox.Show("User this plan has been loaded with your CDSID.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        'End If



        Try
            '------------------------------------------------------------
            ' set button activation according to grids
            '------------------------------------------------------------
            If TabController1.SelectedIndex = 0 Then
                If grdPlans.SelectedRows.Count > 0 Then
                    btnOpenLoad.Enabled = PrerequisitesFulfilled
                    btnCheckout.Enabled = True
                    btnDraft.Enabled = True
                Else
                    MessageBox.Show("Please select a plan to load.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    btnOpenLoad.Enabled = False
                    btnCheckout.Enabled = False
                    btnDraft.Enabled = False
                    Exit Sub
                End If
                btnPreCheck.Enabled = False
                chkLoadIndFormat.Enabled = False

                ' values for validation
                BuildPhase = grdPlans.SelectedRows(0).Cells("BuildPhase").Value

                BuildScale = grdPlans.SelectedRows(0).Cells("AssyBuildScale").Value
                ' values for loading
                HCID = Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value)
                IsGeneric = If(grdPlans.SelectedRows(0).Cells("GenOrSpec").Value.ToString = "Generic", True, False)

            ElseIf TabController1.SelectedIndex = 1 Then
                If grdPlansGeneric.SelectedRows.Count > 0 Then
                    btnOpenLoad.Enabled = PrerequisitesFulfilled ' To force user to press pre-check first and don't load plan if data missed.
                    btnPreCheck.Enabled = True
                Else
                    MessageBox.Show("Please select a plan to load.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    btnOpenLoad.Enabled = False
                    btnPreCheck.Enabled = False
                    chkLoadIndFormat.Enabled = False
                    Exit Sub
                End If
                btnCheckout.Enabled = False
                btnDraft.Enabled = False
                If ValidateGridData() = False Then
                    Me.DialogResult = DialogResult.None
                    Exit Sub
                End If

                ' values for validation
                BuildPhase = grdPlansGeneric.SelectedRows(0).Cells("GBuildPhase").Value.ToString
                BuildScale = Integer.Parse(grdPlansGeneric.SelectedRows(0).Cells("GAssyBuildScale").Value.ToString)
                ' values for loading
                HCID = Integer.Parse(grdPlansGeneric.SelectedRows(0).Cells("GHCID").Value.ToString)
                IsGeneric = If(grdPlansGeneric.SelectedRows(0).Cells("GIsGeneric").Value.ToString = "Generic", True, False)

            End If


            '------------------------------------------------------------
            ' Validate build scale for plans except PP and TT
            '------------------------------------------------------------
            If BuildPhase <> CT.Data.DataCenter.BuildPhase.TT.ToString And BuildPhase <> CT.Data.DataCenter.BuildPhase.PP.ToString Then
                If BuildScale = "0" Then
                    MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHCIDSelect, "Build scale cannot be '0' for Buildphase 'PP/TT'. ") & " Plan cannot be loaded.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Me.DialogResult = DialogResult.None
                    Exit Sub
                End If
            End If


            bol_HeaderCol_Clicked = False
            If txtHcid.Text <> "" Then
                If TabController1.SelectedIndex = 0 Then
                    mySource.Filter = String.Format("Convert([HealthChartID],'System.String') LIKE '%{0}%'", txtHcid.Text)
                    grdPlans.Refresh()

                ElseIf TabController1.SelectedIndex = 1 Then
                    mySourceGeneric.Filter = String.Format("Convert([HealthChartID],'System.String') LIKE '%{0}%'", txtHcid.Text)
                    grdPlansGeneric.Refresh()


                End If

            End If
            Me.Activate()
            If Me.Visible = False Then
                frmProgress.Show()
            End If



            Me.SmoothProgressBar1.Value = 100
            Me.SmoothProgressBar2.Value = 0

            UpdateProgressBar(2)


            '-----------------------------------------------------
            ' Deactive buttons
            '-----------------------------------------------------
            DeactiveControls()

            '------------------------------------------------------------
            ' Main build Type validation & Load plan
            '------------------------------------------------------------
            If strMainBuildType = String.Empty Then Throw New Exception("The Plan type is not defined.")
            If _PlanForLoading.LoadPlan(HCID, IsGeneric, chkLoadIndFormat.Checked, strMainBuildType) <> String.Empty Then
                Me.DialogResult = DialogResult.Retry
                Exit Sub
            End If

            bolWasOn = Globals.ThisAddIn.Application.EnableEvents

            Globals.ThisAddIn.Application.EnableEvents = False

            Try
                Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.ProgramConfig.LastRow + 5 & ":" & Form.DataCenter.GlobalValues.WS.Rows.Count).EntireRow.Delete()
            Catch ex As Exception
            End Try
            Me.DialogResult = DialogResult.OK

        Catch ex1 As Exception
            _ErrorMessage = ex1.Message
            '----------------------------------------------------------------
            ' Because Cancel button has the Cancel DialogResukt the No DialogResult 
            ' is considered as Error
            ' Me.DialogResult = DialogResult.No
            '----------------------------------------------------------------
            Me.DialogResult = DialogResult.No
        Finally
            Globals.ThisAddIn.Application.EnableEvents = bolWasOn
            Form.DataCenter.GlobalValues.bolRefreshCompleted = True
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
            Form.DataCenter.GlobalValues.intProgValue = 0
            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString And Me.DialogResult = DialogResult.OK And Form.DataCenter.ProgramConfig.IsGeneric = False Then
                NotifyToCheckout.ShowBalloonTip(10000)
                NotifyToCheckout.Visible = False
                Dim obj As New Form.DataCenter.ModuleFunction
                obj.DisplayMasterMessage()
            End If
            frmProgress.Close()
            Form.DataCenter.GlobalValues.WS.Activate()
            Me.Close()

        End Try
    End Sub

    'Button cancel click event - to close the form
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        'Dim intcnt As Integer
        'For intcnt = 1 To 10
        '    Globals.ThisAddIn.Application.ScreenUpdating = True
        'Next
        Close()
    End Sub

    'Mouse double click on Gridview click event
    Private Sub grdPlans_DoubleClick(sender As Object, e As EventArgs) Handles grdPlans.DoubleClick

        Try


            Dim dtUsers As System.Data.DataTable

            Dim _PlanActiveUsers As New Data.PlanActiveUsers

            dtUsers = _PlanActiveUsers.SelectAll(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value), grdPlans.SelectedRows(0).Cells("BuildType").Value)

            If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            _CurrentUserStatus = String.Empty
            btnOpenLoad.Enabled = True

            If dtUsers.Rows.Count > 0 Then                    '
                For Each rows In dtUsers.Rows

                    If rows(1).ToString = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString Then
                        _CurrentUserStatus = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString
                    End If

                Next
            End If

            If _CurrentUserStatus = String.Empty Then
                If bol_HeaderCol_Clicked = False Then btnOpenLoad_Click(sender, e)
            Else
                MessageBox.Show("User this plan has been loaded with your CDSID.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                btnOpenLoad.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'HCID textbox key press event
    'To filter & load grid plans
    Private Sub txtHcid_KeyUp(sender As Object, e As KeyEventArgs) Handles txtHcid.KeyUp
        'Dim row As DataGridViewRow = Nothing
        If txtHcid.Text <> "" Then
            mySource.Filter = String.Format("Convert([HealthChartID],'System.String') LIKE '%{0}%'", txtHcid.Text)
            grdPlans.Refresh()

            mySourceGeneric.Filter = String.Format("Convert([HealthChartID],'System.String') LIKE '%{0}%'", txtHcid.Text)
            grdPlansGeneric.Refresh()

        ElseIf txtHCName.Text = "" And txtHcid.Text = "" Then
            mySource.Filter = ""
            mySourceGeneric.Filter = ""
        End If
        If e.KeyCode = Keys.Enter Then
            btnOpenLoad_Click(sender, e)
        End If


        If TabController1.SelectedIndex = 0 Then
            grdPlans.Focus()
            If grdPlans.SelectedRows.Count > 0 Then
                btnOpenLoad.Enabled = True
                btnCheckout.Enabled = True
                btnDraft.Enabled = True
            Else
                btnOpenLoad.Enabled = False
                btnCheckout.Enabled = False
                btnDraft.Enabled = False
            End If
            btnPreCheck.Enabled = False
        ElseIf TabController1.SelectedIndex = 1 Then
            grdPlansGeneric.Focus()
            btnCheckout.Enabled = False
            btnDraft.Enabled = False
            If grdPlansGeneric.SelectedRows.Count > 0 Then
                btnOpenLoad.Enabled = PrerequisitesFulfilled ' To force user to press pre-check first and don't load plan if data missed.
                btnPreCheck.Enabled = True
            Else
                btnOpenLoad.Enabled = False
                btnPreCheck.Enabled = False
            End If
            chkLoadIndFormat.Enabled = False
        End If
        txtHcid.Focus()
    End Sub

    'Program description key press event
    'To filter grid plans
    Private Sub txtHCName_KeyUp(sender As Object, e As KeyEventArgs) Handles txtHCName.KeyUp

        chkLoadIndFormat.Enabled = False

        If txtHCName.Text <> "" Then
            mySource.Filter = String.Format("Convert([HealthChartName],'System.String') LIKE '*{0}*'", txtHCName.Text)
            grdPlans.Refresh()

            mySourceGeneric.Filter = String.Format("Convert([HealthChartName],'System.String') LIKE '*{0}*'", txtHCName.Text)
            grdPlansGeneric.Refresh()

        ElseIf txtHCName.Text = "" And txtHcid.Text = "" Then
            mySource.Filter = ""
            mySourceGeneric.Filter = ""
        End If

        'If TabController1.SelectedIndex = 0 Then
        '    CustomFormatCheckUnCheck()
        'End If
    End Sub

    'For keydown event
    'Shortcut keys for buttons
    Private Sub frmHCIDSelect_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F4 Then
            e.Handled = True
            txtHcid.Focus()
        ElseIf e.KeyCode = Keys.F5 Then
            e.Handled = True
            btnOpenLoad_Click(sender, e)
        ElseIf e.KeyCode = Keys.F9 And btnPreCheck.Enabled = True Then
            e.Handled = True
            btnPreCheck_Click(sender, e)
        ElseIf e.KeyCode = Keys.Escape AndAlso (btnOpenLoad.Enabled = True Or btnPreCheck.Enabled = True) Then
            e.Handled = True
            Me.Close()
        End If
    End Sub


    Private Sub frmHCIDSelect_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If e.Button.ToString = "Right" Then Exit Sub
    End Sub


    Private Sub grdPlans_MouseDown(sender As Object, e As MouseEventArgs) Handles grdPlans.MouseDown
        If e.Button.ToString = "Right" Then Exit Sub
    End Sub


    'Grid 'Enter' button keydown event
    Private Sub grdPlans_KeyDown(sender As Object, e As KeyEventArgs) Handles grdPlans.KeyDown
        Try

            bol_HeaderCol_Clicked = False

            Dim dtUsers As System.Data.DataTable

            Dim _PlanActiveUsers As New Data.PlanActiveUsers

            dtUsers = _PlanActiveUsers.SelectAll(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value), grdPlans.SelectedRows(0).Cells("BuildType").Value)

            If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            _CurrentUserStatus = String.Empty
            btnOpenLoad.Enabled = True

            If dtUsers.Rows.Count > 0 Then                    '
                For Each rows In dtUsers.Rows

                    If rows(1).ToString = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString Then
                        _CurrentUserStatus = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString
                    End If

                Next
            End If


            If _CurrentUserStatus = String.Empty Then
                If e.KeyCode = Keys.Enter Then

                    btnOpenLoad_Click(sender, e)
                End If
            Else
                MessageBox.Show("User this plan has been loaded with your CDSID.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                btnOpenLoad.Enabled = False
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub btnPreCheck_Click(sender As Object, e As EventArgs) Handles btnPreCheck.Click
        Try
            ' pe01- 0, hcid 3, type 13, phase 5
            Dim _frmPlanValidation As frmBase = Nothing

            If strMainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _frmPlanValidation = New frmPlanValidation()
            ElseIf strMainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _frmPlanValidation = New frmPlanValidation_Rig()
            Else
                Exit Sub
            End If


            If TabController1.SelectedIndex = 1 Then
                If grdPlansGeneric.SelectedRows.Count > 0 Then

                    Form.DataCenter.ProgramConfig.XccPe01 = grdPlansGeneric.Item(14, grdPlansGeneric.CurrentRow.Index).Value
                    Form.DataCenter.ProgramConfig.XccPe26 = grdPlansGeneric.Item(15, grdPlansGeneric.CurrentRow.Index).Value
                    Form.DataCenter.ProgramConfig.HCID = grdPlansGeneric.Item(3, grdPlansGeneric.CurrentRow.Index).Value
                    Form.DataCenter.ProgramConfig.BuildType = grdPlansGeneric.Item(13, grdPlansGeneric.CurrentRow.Index).Value
                    Form.DataCenter.ProgramConfig.BuildPhase = grdPlansGeneric.Item(5, grdPlansGeneric.CurrentRow.Index).Value

                    _frmPlanValidation.frmOwner = Me
                    _frmPlanValidation.ShowDialog()

                    btnOpenLoad.Enabled = PrerequisitesFulfilled ' To force user to press pre-check first and don't load plan if data missed.

                End If
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHCIDSelect, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub TabController1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabController1.SelectedIndexChanged
        If TabController1.SelectedIndex = 0 Then
            grdPlans.Focus()
            If grdPlans.SelectedRows.Count > 0 Then
                btnOpenLoad.Enabled = True
                btnCheckout.Enabled = True
                btnDraft.Enabled = True
            Else
                btnOpenLoad.Enabled = False
                btnCheckout.Enabled = False
                btnDraft.Enabled = False
            End If
            btnPreCheck.Enabled = False
        ElseIf TabController1.SelectedIndex = 1 Then
            grdPlansGeneric.Focus()
            btnCheckout.Enabled = False
            btnDraft.Enabled = False
            chkLoadIndFormat.Enabled = False
            If grdPlansGeneric.SelectedRows.Count > 0 Then
                btnOpenLoad.Enabled = PrerequisitesFulfilled ' To force user to press pre-check first and don't load plan if data missed.
                btnPreCheck.Enabled = True
            Else
                btnOpenLoad.Enabled = False
                btnPreCheck.Enabled = False
            End If
        End If

    End Sub

    Private Sub grdPlansGeneric_DoubleClick(sender As Object, e As EventArgs) Handles grdPlansGeneric.DoubleClick

        '------------------------------------------------------------------------------------------------------
        ' To force user to press pre-check first for generic plans and don't load generic plan if data missed.
        '------------------------------------------------------------------------------------------------------
        If PrerequisitesFulfilled = False Then btnOpenLoad.Enabled = PrerequisitesFulfilled

        If bol_HeaderCol_Clicked = False Then btnOpenLoad_Click(sender, e)
    End Sub


    Private Sub grdPlansGeneric_KeyDown(sender As Object, e As KeyEventArgs) Handles grdPlansGeneric.KeyDown
        bol_HeaderCol_Clicked = False
        'switchchkLoadIndFormat(grdPlansGeneric.Item(2, grdPlansGeneric.CurrentRow.Index).Value)
        If e.KeyCode = Keys.Enter Then
            btnOpenLoad_Click(sender, e)
        End If
    End Sub

    Private Sub frmHCIDSelect_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Globals.ThisAddIn.Application.EnableEvents = True
        If Me.DialogResult = DialogResult.Retry Then
            'After any alert message - in button checkout / load plan click
            'Display the form - bby refreshing the buttons/grids
            e.Cancel = True

            If TabController1.SelectedIndex = 1 Then
                btnPreCheck.Enabled = True
                btnOpenLoad.Enabled = False
            End If

            If TabController1.SelectedIndex = 0 Then
                btnCheckout.Enabled = True
                btnDraft.Enabled = True
                btnOpenLoad.Enabled = True
                btnPreCheck.Enabled = False
            End If

            grdPlans.Enabled = True
            grdPlansGeneric.Enabled = True
            txtHcid.ReadOnly = False
            txtHCName.ReadOnly = False
            btnCancel.Enabled = True
            chkLoadIndFormat.Enabled = False

            Me.SmoothProgressBar1.Value = 0
            Me.SmoothProgressBar2.Value = 0
            Me.lblProgress.Text = ""
            Me.lblProgress.Refresh()
            Me.SmoothProgressBar1.Refresh()
            Me.SmoothProgressBar2.Refresh()
        ElseIf Me.DialogResult = DialogResult.No Then
            e.Cancel = True
        ElseIf Me.DialogResult = DialogResult.Yes Then
            'Reload the grid - After Check-in click
            e.Cancel = True
            'frmHCIDSelect_Load(sender, e)

        End If
    End Sub


    Private Sub btnCheckout_Click(sender As Object, e As EventArgs) Handles btnCheckout.Click
        Try
            Dim objPer As New CT.Data.Authorization
            Dim _strUserPermissionLevel As String = String.Empty

            '--------------------------------------------------------
            ' Set value for later usage
            '--------------------------------------------------------
            If TabController1.SelectedIndex = 0 Then
                Form.DataCenter.ProgramConfig.pe01 = Long.Parse(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value)
                Form.DataCenter.ProgramConfig.HCID = Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value)
                Form.DataCenter.ProgramConfig.IsGeneric = If(grdPlans.SelectedRows(0).Cells("GenOrSpec").Value = "Generic", True, False)
                Form.DataCenter.ProgramConfig.pe02 = Long.Parse(grdPlans.SelectedRows(0).Cells("pe02").Value)
                Form.DataCenter.ProgramConfig.XccPe26 = Long.Parse(grdPlans.SelectedRows(0).Cells("XCCpe26").Value)
                Form.DataCenter.ProgramConfig.XccPe01 = Long.Parse(grdPlans.SelectedRows(0).Cells("XCCpe01").Value)
                Form.DataCenter.ProgramConfig.AssyBuildScale = Long.Parse(grdPlans.SelectedRows(0).Cells("AssyBuildScale").Value)
                Form.DataCenter.ProgramConfig.BuildType = grdPlans.SelectedRows(0).Cells("BuildType").Value.ToString
                Form.DataCenter.ProgramConfig.BuildPhase = grdPlans.SelectedRows(0).Cells("BuildPhase").Value.ToString
                Form.DataCenter.ProgramConfig.Carline = grdPlans.SelectedRows(0).Cells("Carline").Value.ToString
                Form.DataCenter.ProgramConfig.Platform = grdPlans.SelectedRows(0).Cells("Platform").Value.ToString
                Form.DataCenter.ProgramConfig.HCIDName = grdPlans.SelectedRows(0).Cells("ProgramDescription").Value.ToString
                Form.DataCenter.ProgramConfig.IsWithCustomFormatting = chkLoadIndFormat.Checked
                Form.DataCenter.ProgramConfig.IsMainPlan = True
                Form.DataCenter.ProgramConfig.Region = Trim(grdPlans.SelectedRows(0).Cells("Region").Value.ToString)
                Form.DataCenter.ProgramConfig.FileStatus = Trim(grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString) ' for check-Out/-In logic
                Form.DataCenter.ProgramConfig.MainPlanHCID = If(Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Checkedout.ToString, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value.ToString.Substring(3)), Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value)) ' for check-Out/-In logic
                Try
                    Dim objDat As New CT.Data.MessagePassing
                    Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                    Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
                Catch ex As Exception

                End Try

                loadActiveusers()

            End If

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
                System.Windows.Forms.MessageBox.Show(ex.Message, Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

            End Try

            If IsNothing(Form.DataCenter.GlobalValues.strUserPermissionLevel) = False Then
                If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower = CT.Data.DataCenter.UserPermissionLevel.Executor.ToString.ToLower Or
                    Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower = CT.Data.DataCenter.UserPermissionLevel.Owner.ToString.ToLower Then

                    Select Case Trim(grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString)
                        Case CT.Data.DataCenter.FileStatus.Master.ToString
                            CheckoutToolStripMenuItem.Visible = True
                            CheckinToolStripMenuItem.Visible = False
                            DiscardToolStripMenuItem.Visible = False
                        Case CT.Data.DataCenter.FileStatus.Checkedout.ToString
                            CheckoutToolStripMenuItem.Visible = False
                            CheckinToolStripMenuItem.Visible = True
                            DiscardToolStripMenuItem.Visible = True

                    End Select

                    ContextMenuStrip1.Show(btnCheckout, 0, btnCheckout.Height)
                    Me.DialogResult = DialogResult.No
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub CheckoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CheckoutToolStripMenuItem.Click
        Try
            '----------------------------------------------------------------
            ' The permission check is done under main button 
            '----------------------------------------------------------------

            Dim Answer As String = String.Empty
            Dim _PlanDisplay As New Form.DisplayUtilities.Plan()

            If TabController1.SelectedIndex = 0 Then
                If grdPlans.SelectedRows.Count <= 0 Then
                    MessageBox.Show("Please select a plan to load.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            End If

            Dim result As DialogResult = MessageBox.Show("Do you really want to checkout?", "Checkout plan", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                Exit Sub
            End If
            '-----------------------------------------------------
            ' Deactive buttons
            '-----------------------------------------------------
            DeactiveControls()


            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            Form.DataCenter.GlobalValues.Clear()


            If Trim(grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString) <> "Master" Then
                Me.DialogResult = DialogResult.Retry
                System.Windows.Forms.MessageBox.Show("Only Master plan can be checked-out.", "Check-out Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Exit Try
            End If

            Answer = _PlanDisplay.CheckOutPlan()
            If Answer <> String.Empty Then Throw New Exception(Answer)


            Me.DialogResult = DialogResult.OK

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Check-out", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True
        End Try
    End Sub

    Private Sub CheckinToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CheckinToolStripMenuItem.Click
        Try
            '----------------------------------------------------------------
            ' The permission check is done under main button 
            '----------------------------------------------------------------

            Cursor = Cursors.AppStarting
            Dim objPer As New CT.Data.Authorization

            If grdPlans.SelectedRows(0).Cells("GenOrSpec").Value.ToString = "Generic" Then Exit Sub

            If Trim(grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString) = "Master" Then
                System.Windows.Forms.MessageBox.Show("The Plan is not checked out yet.", "Checked out Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                Cursor = Cursors.Default
                Me.DialogResult = DialogResult.Retry
                Exit Sub
            End If

            'Check the Activeusers in this plan
            Dim dtUsers As DataTable
            Dim _PlanActiveUsers As New Data.PlanActiveUsers
            Dim strActiveUsers As String = ""
            dtUsers = _PlanActiveUsers.SelectAll(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value), grdPlans.SelectedRows(0).Cells("BuildType").Value)

            If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            If dtUsers.Rows.Count > 0 Then                    '
                For Each rows In dtUsers.Rows
                    strActiveUsers = strActiveUsers & rows(0).ToString() & ","

                Next
                System.Windows.Forms.MessageBox.Show("The Plan cannot be checked in now as the below users are active in this plan." & vbNewLine & strActiveUsers & "", "Active users in Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                Cursor = Cursors.Default
                Me.DialogResult = DialogResult.Retry
                Exit Sub
            End If

            Form.DataCenter.ProgramConfig.HCID = Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value)
            Form.DataCenter.ProgramConfig.IsGeneric = If(grdPlans.SelectedRows(0).Cells("GenOrSpec").Value = "Generic", True, False)
            Form.DataCenter.ProgramConfig.BuildType = grdPlans.SelectedRows(0).Cells("BuildType").Value.ToString
            Try
                Dim objDat As New CT.Data.MessagePassing
                Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
            Catch ex As Exception

            End Try
            Try
                Dim objDat As New CT.Data.MessagePassing
                Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
            Catch ex As Exception

            End Try
            Dim _PlanForCheckin As Data.VehiclePlan.Plan = New Data.VehiclePlan.Plan

            '----------------------------------------------------------
            ' Request user to change the issue version
            '----------------------------------------------------------
            Dim vbResult As MsgBoxResult
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

                '----------------------------------------------------------
                ' If user cancel the form
                '----------------------------------------------------------
                If _frmObject.ShowDialog() = DialogResult.Cancel Then Exit Sub
            ElseIf vbResult = DialogResult.Cancel Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            '----------------------------------------------------------
            ' Replace checkedout version on master version
            '----------------------------------------------------------
            If _PlanForCheckin.ConvertCheckedouttToLife(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.HCID, CT.Data.DataCenter.FileStatus.Checkedout, Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            '--------------------------------------------------------------------------
            ' Setting for specific plans
            '--------------------------------------------------------------------------
            grdPlans.AutoGenerateColumns = False
            grdPlans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            '-----------------------------------------------------------------------------------------
            ' Fill specific plans
            '-----------------------------------------------------------------------------------------
            _ErrorMessage = Fill_grdPlans()
            If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)

            System.Windows.Forms.MessageBox.Show("The Plan has been checked in successfuly.", "Checked in Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DiscardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DiscardToolStripMenuItem.Click
        Dim _Plan As Data.VehiclePlan.Plan = New Data.VehiclePlan.Plan

        Try
            '----------------------------------------------------------------
            ' The permission check is done under main button 
            '----------------------------------------------------------------

            Cursor = Cursors.WaitCursor

            If grdPlans.SelectedRows(0).Cells("GenOrSpec").Value.ToString = "Generic" Then Exit Sub
            If Trim(grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString) = "Master" Then
                System.Windows.Forms.MessageBox.Show("The Plan is not checked out yet.", "Checked out Plan", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                Me.DialogResult = DialogResult.Retry
                Exit Sub
            End If


            '----------------------------------------------------------
            ' Take confirmation from user
            '----------------------------------------------------------
            If MessageBox.Show("Do you really want to discard the changes? ", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then

                '----------------------------------------------------------
                ' Keep the main HCID, IsGeneric and withCustomFormatting & Discard Plan in DB
                '----------------------------------------------------------
                If _Plan.DeleteDraftOrCheckedout(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, DirectCast([Enum].Parse(GetType(CT.Data.DataCenter.FileStatus), Form.DataCenter.ProgramConfig.FileStatus), CT.Data.DataCenter.FileStatus), grdPlans.SelectedRows(0).Cells("BuildType").Value.ToString) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                '----------------------------------------------------------
                ' Refresh the grid 
                '----------------------------------------------------------
                '--------------------------------------------------------------------------
                ' Setting for specific plans
                '--------------------------------------------------------------------------
                grdPlans.AutoGenerateColumns = False
                grdPlans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

                '-----------------------------------------------------------------------------------------
                ' Fill specific plans
                '-----------------------------------------------------------------------------------------
                _ErrorMessage = Fill_grdPlans()
                If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub GenerateDraftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateDraftToolStripMenuItem.Click
        Try

            Cursor = Cursors.WaitCursor
            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                System.Windows.Forms.MessageBox.Show("Draft option is only for 'Specific' plans.", Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Exit Sub
            End If

            'Dim _Plan As New Data.VehiclePlan.Plan
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
            If resultDataTable.Rows.Count >= 3 Then Throw New Exception("3 Draft versions are already created for this HC ID : " & Form.DataCenter.ProgramConfig.HCID)

            If _PlanInterface.GenerateDraftOrCheckout(Form.DataCenter.ProgramConfig.HCID, Data.DataCenter.FileStatus.Draft, Form.DataCenter.ProgramConfig.BuildType) = True Then
                loadMenubutton()
                System.Windows.Forms.MessageBox.Show("Draft completed successfully.", Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                'Refesh the Load Draft submenu Subitem
            Else
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub loadMenubutton()

        Try


            LoadDraftToolStripMenuItem.DropDownItems.Clear()
            If Form.DataCenter.ProgramConfig.IsGeneric = False Then

                Dim dtDraft As System.Data.DataTable
                'Dim _Plan As New Data.VehiclePlan.Plan()
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
                    For Each rows In dtDraft.Rows



                        Dim newitem As ToolStripItem = LoadDraftToolStripMenuItem.DropDownItems.Add(rows(2).ToString & " - " & rows(4).ToString)
                        newitem.Tag = rows(2).ToString
                        AddHandler newitem.Click, AddressOf Me.newitem_click

                    Next
                    LoadDraftToolStripMenuItem.Enabled = True
                    LoadDraftToolStripMenuItem.Visible = True
                Else
                    LoadDraftToolStripMenuItem.Enabled = False
                    LoadDraftToolStripMenuItem.Visible = False
                End If

            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub loadActiveusers()

        Try

            ActiveusersToolStripMenuItem.DropDownItems.Clear()
            If Form.DataCenter.ProgramConfig.IsGeneric = False Then

                Dim dtUsers As System.Data.DataTable

                Dim _PlanActiveUsers As New Data.PlanActiveUsers

                dtUsers = _PlanActiveUsers.SelectAll(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value), grdPlans.SelectedRows(0).Cells("BuildType").Value)

                If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                If dtUsers.Rows.Count > 0 Then                    '
                    For Each rows In dtUsers.Rows

                        Dim newitem As ToolStripItem = ActiveusersToolStripMenuItem.DropDownItems.Add(rows(0).ToString) ' & " - " & rows(4).ToString)
                        'newitem.Tag = rows(2).ToString
                        'AddHandler newitem.Click, AddressOf Me.newitem_click
                    Next
                    ActiveusersToolStripMenuItem.Enabled = True
                    ActiveusersToolStripMenuItem.Visible = True
                Else
                    ActiveusersToolStripMenuItem.Enabled = False
                    ActiveusersToolStripMenuItem.Visible = False
                End If


            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally

        End Try
    End Sub


    Public Sub newitem_click(sender As Object, e As EventArgs)
        Dim bolWasOn As Boolean
        Dim _Plan As New Form.DisplayUtilities.Plan()
        Dim Answer As String = String.Empty
        Try

            Cursor = Cursors.WaitCursor
            '-----------------------------------------------------
            ' Deactive buttons
            '-----------------------------------------------------
            DeactiveControls()

            '-------------------------------------------------
            ' deactivate screen updating
            '-------------------------------------------------
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False

            '-------------------------------------------------
            ' close the previous template
            '-------------------------------------------------
            Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)

            '-------------------------------------------------
            ' Load draft plan
            '-------------------------------------------------
            Dim BuildType As String
            If btnVehicleList.Checked = True Then
                BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString()
            ElseIf btnRigList.Checked = True Then
                BuildType = CT.Data.DataCenter.BuildType.Rig.ToString()
            ElseIf btnBuckList.Checked = True Then
                BuildType = CT.Data.DataCenter.BuildType.Buck.ToString()
            End If
            Answer = _Plan.LoadDraftPlan(sender.tag, chkLoadIndFormat.Checked, Form.DisplayUtilities.Plan.LoadType.Loading, BuildType)
            If Answer <> String.Empty Then Throw New Exception(Answer)

            bolWasOn = Globals.ThisAddIn.Application.EnableEvents

            Globals.ThisAddIn.Application.EnableEvents = False

            Me.DialogResult = DialogResult.OK
        Catch ex As Exception

            _ErrorMessage = "Loading draft plan: " + ex.Message
            '----------------------------------------------------------------
            ' Because Cancel button has the Cancel DialogResukt the No DialogResult 
            ' is considered as Error
            ' Me.DialogResult = DialogResult.No
            '----------------------------------------------------------------
            Me.DialogResult = DialogResult.No
        Finally
            Cursor = Cursors.Default
            Globals.ThisAddIn.Application.EnableEvents = bolWasOn
            Form.DataCenter.GlobalValues.bolRefreshCompleted = True
            Globals.Ribbons.RbnTnDControlPanel.Tabs(0).RibbonUI.ActivateTab("tabTndPlanControlPanel")
            Form.DataCenter.GlobalValues.intProgValue = 0

            frmProgress.Close()
            Form.DataCenter.GlobalValues.WS.Activate()

            Me.Close()
        End Try
    End Sub

    Private Sub btnDraft_Click(sender As Object, e As EventArgs) Handles btnDraft.Click
        Try
            Dim objPer As New CT.Data.Authorization
            Dim _strUserPermissionLevel As String = String.Empty
            '----------------------------------------------
            ' These functions are only for specific plans
            '----------------------------------------------
            If TabController1.SelectedIndex <> 0 Then Exit Sub

            '----------------------------------------------
            ' Check the permission availability
            '----------------------------------------------

            Form.DataCenter.ProgramConfig.HCID = Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value)
            Form.DataCenter.ProgramConfig.IsGeneric = If(grdPlans.SelectedRows(0).Cells("GenOrSpec").Value = "Generic", True, False)
            Form.DataCenter.ProgramConfig.BuildType = grdPlans.SelectedRows(0).Cells("BuildType").Value
            Form.DataCenter.ProgramConfig.FileStatus = grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString
            Try
                Dim objDat As New CT.Data.MessagePassing
                Dim DT As System.Data.DataTable = objDat.SelectAll(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                Form.DataCenter.GlobalValues.CurrentTotalMessages = DT.Rows.Count
            Catch ex As Exception

            End Try
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
                System.Windows.Forms.MessageBox.Show(ex.Message, Me.Text, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
            '----------------------------------------------
            ' These functions are only for Executor and owner available
            '----------------------------------------------
            If IsNothing(Form.DataCenter.GlobalValues.strUserPermissionLevel) = False Then
                If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower = CT.Data.DataCenter.UserPermissionLevel.Executor.ToString.ToLower Or
                    Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower = CT.Data.DataCenter.UserPermissionLevel.Owner.ToString.ToLower Then


                    Form.DataCenter.ProgramConfig.HCID = Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value)
                    Form.DataCenter.ProgramConfig.IsGeneric = If(grdPlans.SelectedRows(0).Cells("GenOrSpec").Value = "Generic", True, False)


                    Select Case Trim(grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString)
                        Case CT.Data.DataCenter.FileStatus.Master.ToString, CT.Data.DataCenter.FileStatus.Checkedout.ToString
                            GenerateDraftToolStripMenuItem.Enabled = True
                            LoadDraftToolStripMenuItem.Enabled = True
                    End Select

                    '----------------------------------------------
                    ' Load the draft versions under button
                    '----------------------------------------------
                    loadMenubutton()

                    '----------------------------------------------
                    ' Display context menu
                    '----------------------------------------------
                    ContextMenuStripDraft.Show(btnDraft, 0, btnDraft.Height)

                    Me.DialogResult = DialogResult.No

                End If
            Else
                Throw New Exception("Access denied! Please contact 'AEREN8@FORD.COM' or OMEIGEN@FORD.COM or PNEZHAD@FORD.COM OR MAGES@FORD.COM for permission!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CustomFormatCheckUnCheck()

        Dim objPer As New CT.Data.Authorization
        Dim _strUserPermissionLevel As String = String.Empty

        Try
            _strUserPermissionLevel = objPer.GetPermissionLevel(grdPlans.SelectedRows(0).Cells("BuildType").Value.ToString, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value.ToString), False)
        Catch ex As Exception
        End Try
        If grdPlans.SelectedRows.Count > 0 Then
            If (_strUserPermissionLevel.ToLower = CT.Data.DataCenter.UserPermissionLevel.Executor.ToString.ToLower Or
            _strUserPermissionLevel.ToLower = CT.Data.DataCenter.UserPermissionLevel.Owner.ToString.ToLower) And
            grdPlans.SelectedRows(0).Cells("FileStatus").Value.ToString.ToLower = CT.Data.DataCenter.FileStatus.Checkedout.ToString.ToLower And
             TabController1.SelectedIndex = 0 Then
                chkLoadIndFormat.Enabled = True
            Else
                chkLoadIndFormat.Checked = False
                chkLoadIndFormat.Enabled = False
                MsgBox("Custom formatting access not enabled for the selected plan.", MsgBoxStyle.Exclamation, "Custom Formatting")
            End If
        Else
            chkLoadIndFormat.Checked = False
            chkLoadIndFormat.Enabled = False
        End If

    End Sub

    Private Sub grdPlans_MouseClick(sender As Object, e As MouseEventArgs) Handles grdPlans.MouseClick

        Try
            'CustomFormatCheckUnCheck()
            chkLoadIndFormat.Checked = False
            chkLoadIndFormat.Enabled = True


            Dim dtUsers As System.Data.DataTable

            Dim _PlanActiveUsers As New Data.PlanActiveUsers

            dtUsers = _PlanActiveUsers.SelectAll(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value), grdPlans.SelectedRows(0).Cells("BuildType").Value)

            If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            _CurrentUserStatus = String.Empty
            btnOpenLoad.Enabled = True

            If dtUsers.Rows.Count > 0 Then                    '
                For Each rows In dtUsers.Rows

                    If rows(1).ToString = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString Then
                        _CurrentUserStatus = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString
                        btnOpenLoad.Enabled = False
                    End If

                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub grdPlans_KeyUp(sender As Object, e As KeyEventArgs) Handles grdPlans.KeyUp
        'CustomFormatCheckUnCheck()
        chkLoadIndFormat.Checked = False
        chkLoadIndFormat.Enabled = True
        Try

            Dim dtUsers As System.Data.DataTable

            Dim _PlanActiveUsers As New Data.PlanActiveUsers

            dtUsers = _PlanActiveUsers.SelectAll(grdPlans.SelectedRows(0).Cells("pe01_TnDBasicProgram_ID").Value, Integer.Parse(grdPlans.SelectedRows(0).Cells("HCID").Value), grdPlans.SelectedRows(0).Cells("BuildType").Value)

            If dtUsers Is Nothing And CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            _CurrentUserStatus = String.Empty
            btnOpenLoad.Enabled = True

            If dtUsers.Rows.Count > 0 Then                    '
                For Each rows In dtUsers.Rows

                    If rows(1).ToString = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString Then
                        _CurrentUserStatus = CT.Data.DataCenter.CurrentUserStatus.CurrentUser.ToString
                        btnOpenLoad.Enabled = False
                    End If

                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub grdPlans_GotFocus(sender As Object, e As EventArgs) Handles grdPlans.GotFocus
        'CustomFormatCheckUnCheck()
        'chkLoadIndFormat.Enabled = True
    End Sub

    Private Sub txtHCName_TextChanged(sender As Object, e As EventArgs) Handles txtHCName.TextChanged

    End Sub

    Public Function FillForm() As String

        Try

            PrerequisitesFulfilled = False 'To monitor user if he used prechecked button or not.
            Form.DataCenter.GlobalValues.intProgValue = 0
            bol_HeaderCol_Clicked = False


            _ErrorMessage = Fill_grdPlans()
            If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)
            _ErrorMessage = Fill_grdPlansGeneric()
            If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)

            grdPlans.Refresh()
            grdPlansGeneric.Refresh()

            '------------------------------------------------------------
            ' set button activation according to grids
            '------------------------------------------------------------
            If TabController1.SelectedIndex = 0 Then
                If grdPlans.SelectedRows.Count > 0 Then
                    btnOpenLoad.Enabled = True
                    btnCheckout.Enabled = True
                    btnDraft.Enabled = True
                Else
                    btnOpenLoad.Enabled = False
                    btnCheckout.Enabled = False
                    btnDraft.Enabled = False
                End If
                btnPreCheck.Enabled = False
                chkLoadIndFormat.Checked = False
                chkLoadIndFormat.Enabled = True
            ElseIf TabController1.SelectedIndex = 1 Then
                btnCheckout.Enabled = False
                btnDraft.Enabled = False
                If grdPlansGeneric.SelectedRows.Count > 0 Then
                    btnOpenLoad.Enabled = PrerequisitesFulfilled ' To force user to press pre-check first and don't load plan if data missed.
                    btnPreCheck.Enabled = True

                Else
                    btnOpenLoad.Enabled = False
                    btnPreCheck.Enabled = False
                    chkLoadIndFormat.Enabled = False
                End If
            End If
            '------------------------------------------------------------
            TabController1.SelectedIndex = 0

            txtHcid.Focus()

            FillForm = String.Empty
        Catch ex As Exception
            FillForm = ex.Message
        End Try


    End Function

    Private Sub btnVehicleList_CheckedChanged(sender As Object, e As EventArgs) Handles btnVehicleList.CheckedChanged

        Try


            Try

            Catch ex As Exception
                MessageBox.Show("There was a problem with stablishing connection to DB!", "Select & Load Plan", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            _ErrorMessage = String.Empty
            If btnVehicleList.Checked Then
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait


                btnBuckList.Checked = False
                btnRigList.Checked = False
                strMainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString

                _ErrorMessage = FillForm()
                If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

        End Try

    End Sub
    Private Sub btnBuckList_CheckedChanged(sender As Object, e As EventArgs) Handles btnBuckList.CheckedChanged

        Try
            _ErrorMessage = String.Empty
            If btnBuckList.Checked Then
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

                btnVehicleList.Checked = False
                btnRigList.Checked = False
                strMainBuildType = CT.Data.DataCenter.BuildType.Buck.ToString

                _ErrorMessage = FillForm()
                If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

        End Try
    End Sub
    Private Sub btnRigList_CheckedChanged(sender As Object, e As EventArgs) Handles btnRigList.CheckedChanged
        Try
            _ErrorMessage = String.Empty
            If btnRigList.Checked Then
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

                btnVehicleList.Checked = False
                btnBuckList.Checked = False
                strMainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString

                _ErrorMessage = FillForm()
                If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

        End Try
    End Sub

    Private Sub frmHCIDSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            '-----------------------------------------------------------------------------------------
            ' Check active database
            '-----------------------------------------------------------------------------------------
            Dim _user_confi As CT.Data.UserLevelConfiguration = New Data.UserLevelConfiguration()
            _user_confi.CT_ConnectionString = "Set"
            '-----------------------------------------------------------------------------------------

            '-----------------------------------------------------------------------------------------
            ' These two values are together if the vehicle button is selected per default the
            ' StartMainBuildType must be Vehicle per default.
            '-----------------------------------------------------------------------------------------
            btnVehicleList.Checked = True
            strMainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString()
            '-----------------------------------------------------------------------------------------


            btnBuckList.Checked = False
            btnRigList.Checked = False
            strMainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString

            _ErrorMessage = FillForm()
            If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)



        Catch ex As Exception

            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

        End Try

    End Sub

    Private Sub chkLoadIndFormat_Enter(sender As Object, e As EventArgs) Handles chkLoadIndFormat.Enter
        CustomFormatCheckUnCheck()
    End Sub

End Class
