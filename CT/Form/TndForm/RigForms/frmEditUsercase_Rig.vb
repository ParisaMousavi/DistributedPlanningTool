Imports System.Windows.Forms
Imports System.Data
Imports System.ComponentModel

Public Class frmEditUsercase_Rig

    Dim dbl_pe26_SpecificVehicleUsercases_PK As Double

    Private _SelectedProcessstep As CT.Data.ProcessStep = New CT.Data.ProcessStep

    Dim bolRecordChanged As Boolean
    Dim bolSkipEvent As Boolean
    Dim bolChangeFlag As Boolean
    Public bolWasUpdated As Boolean = False
    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
    Dim _Editprocess As Data.Facility = New Data.Facility()

    Dim myDatatable As New System.Data.DataTable
    Dim dtCBG As DataTable = Nothing
    Dim myView As New System.Data.DataView
    Dim dtInitEndDate As Date
    Dim dtInitStartDate As Date
    Private _pe03 As Long = 60914
    Private _AllocatedUsercaseSequence As Integer = 3

    Public Property AllocatedUsercaseSequence As Long
        Get
            AllocatedUsercaseSequence = _AllocatedUsercaseSequence
        End Get
        Set(value As Long)
            _AllocatedUsercaseSequence = value
        End Set
    End Property


    Private Sub ClearForm()
        lblProcessStep.Text = ""
        lblUserCase.Text = ""
        cboLocation_CBG.DataSource = Nothing
        cboProcessStepLocation.DataSource = Nothing
        cboMatchedFacility.DataSource = Nothing
        cboSubFacility.DataSource = Nothing

        txtRemarks.Text = ""

        txtDuration.Text = ""
        txtDuration.Enabled = False

        lblGlobal.Text = ""
        lblUser.Text = ""

        chkStart.Checked = False
        chkEnd.Checked = False
        chkStart.Enabled = True
        chkEnd.Enabled = True
        chkStartAndEnd.Enabled = True
        chkStartAndEnd.Checked = False

        opt5Days.Checked = True

        optWorkingDays.Checked = True
        optWorkingDays.Enabled = False '    

        optWeeks.Enabled = False '

    End Sub

    Private Function Fill_dgvUsercase() As String
        Try
            Dim dtUsercaseProcessSteps As DataTable
            Dim _Usercase As CT.Data.Usercase = New Data.Usercase
            Fill_dgvUsercase = String.Empty
            '------------------------------------------
            ' Fetch all the processsteps by loading
            '------------------------------------------
            dtUsercaseProcessSteps = _Usercase.SelectUsercaseDedicated(Form.DataCenter.VehicleConfig.VehiclePe03, _AllocatedUsercaseSequence)


            '------------------------------------------
            ' Validate the return value from DAL
            '------------------------------------------
            If dtUsercaseProcessSteps Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message + "The processsteps pf the usercase were not found.")

            '------------------------------------------
            ' List the usercase processsteps in grid
            '------------------------------------------
            myView = New System.Data.DataView(dtUsercaseProcessSteps)
            dtUsercaseProcessSteps = myView.ToTable(False, "ProcessStepName", "Duration", "PlannedStart", "PlannedEnd", "Workingdays", "Cdsid", "FacilityCbg", "FacilityLocation", "FacilityName", "SubFacilityName", "Remarks", "pe26_SpecificVehicleUsercases_PK")
            dgvUsercase.DataSource = dtUsercaseProcessSteps
        Catch ex As Exception
            Fill_dgvUsercase = ex.Message
        End Try
    End Function


    Private Sub frmEditUsercase_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim ErrorMessage As String = String.Empty

        Try
            '------------------------------------------
            ' Selected usercase validation
            '------------------------------------------
            If _AllocatedUsercaseSequence = -1 Then Throw New Exception("Allocated Usercase Sequence is not valid.")

            '------------------------------------------
            ' Set the cursor
            '------------------------------------------
            Me.Cursor = Cursors.WaitCursor

            '------------------------------------------
            ' Clear the userform
            '------------------------------------------
            ClearForm()

            '------------------------------------------
            ' Fetch CBG only one time
            '------------------------------------------
            dtCBG = _Editprocess.GetCbg(FacilityCbg:=Nothing, FacilityLocation:=Nothing, FacilityName:=Nothing, SubFacilityName:=Nothing)
            '---------------------------------------------------
            ' Validation for CBG
            '---------------------------------------------------
            If dtCBG Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message + " The CBG data source is empty.")


            '---------------------------------------------------
            ' Fill CBG comobobox only one time
            '---------------------------------------------------
            Fill_cboLocation_CBG()


            '---------------------------------------------------
            ' Fill usercase process steps grid
            '---------------------------------------------------
            ErrorMessage = Fill_dgvUsercase()
            If ErrorMessage <> String.Empty Then Throw New Exception(ErrorMessage)

            '------------------------------------------
            ' Apply generic plan Settings
            '------------------------------------------
            cmdUpdate.Enabled = False
            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                cmdReset.Enabled = False
                cmdUpdate.Enabled = False
                cmdUpdateAllCDSID.Enabled = False
            Else
                cmdReset.Enabled = True
            End If


            '------------------------------------------
            ' Fill the form with selected Processstep
            '------------------------------------------
            PopulateData(0)


        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEditUsercase, ex.Message), "Edit Usercase", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Close() ' if the selected allocated usercase sequence is not valid the form can be closed.
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvUsercase_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvUsercase.CellClick
        Try
            Static bolInLoop As Boolean
            If bolInLoop = True Then Exit Sub
            If e.RowIndex < 0 Then Exit Sub

            Me.Cursor = Cursors.WaitCursor

            cmdUpdate.Enabled = False

            PopulateData(e.RowIndex)

            bolInLoop = False
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEditUsercase, ex.Message))
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub Fill_cboLocation_CBG()
        '---------------------------------------------------
        ' Set CBG Combobox
        '---------------------------------------------------
        myView = New System.Data.DataView(dtCBG)
        myDatatable = myView.ToTable(False, "FacilityCbg")
        cboLocation_CBG.DataSource = myDatatable
        cboLocation_CBG.DisplayMember = "FacilityCbg"
        cboLocation_CBG.ValueMember = "FacilityCbg"

    End Sub


    ''' <summary>
    ''' This sub displays the selected processstep information in user form.
    ''' </summary>
    ''' <param name="datarow"></param>
    Private Sub PopulateData(datarow As Integer)
        Try
            Cursor = Cursors.WaitCursor
            '---------------------------------------------------
            ' Validation for CBG
            '---------------------------------------------------
            'If dtCBG Is Nothing Then Throw New Exception("The CBG data source is empty.")


            ''---------------------------------------------------
            '' Set CBG Combobox
            ''---------------------------------------------------
            'myView = New System.Data.DataView(dtCBG)
            'myDatatable = myView.ToTable(False, "FacilityCbg")
            'cboLocation_CBG.DataSource = myDatatable
            'cboLocation_CBG.DisplayMember = "FacilityCbg"
            'cboLocation_CBG.ValueMember = "FacilityCbg"

            'txtDuration.Enabled = True
            'txtDuration.Text = ""
            'txtDuration.Enabled = False

            optWorkingDays.Checked = True

            chkStart.Checked = False
            chkEnd.Checked = False
            chkStart.Enabled = True
            chkEnd.Enabled = True
            chkStartAndEnd.Checked = False
            chkStartAndEnd.Enabled = True


            '---------------------------------------------------
            ' Fech the dedicated Process step
            '---------------------------------------------------
            dbl_pe26_SpecificVehicleUsercases_PK = Val(dgvUsercase.Rows(datarow).Cells("pe26").Value)
            _SelectedProcessstep.SelectProcessStepDedicated(dbl_pe26_SpecificVehicleUsercases_PK, Form.DataCenter.ProgramConfig.IsGeneric)
            If Not _SelectedProcessstep.pe26 > 0 Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message + " Processstep was not found.")


            '---------------------------------------------------
            ' Display Process step information
            '---------------------------------------------------
            lblProcessStep.Text = _SelectedProcessstep.ProcessStepName
            lblUserCase.Text = _SelectedProcessstep.Usercase

            Dim dtCDSID As System.Data.DataTable
            dtCDSID = _SelectedProcessstep.GetAllCdsids(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, _SelectedProcessstep.GlobalDVP, Form.DataCenter.ProgramConfig.BuildType)
            If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            cboCDSID.Items.Clear()
            cboCDSID.Items.Add("CDSID")
            If dtCDSID.Rows.Count > 0 Then
                For i As Int16 = 0 To dtCDSID.Rows.Count - 1
                    cboCDSID.Items.Add(dtCDSID.Rows(i).Item(0).ToString())
                Next

            End If

            If _SelectedProcessstep.ProcessStepName <> "Gap" And _SelectedProcessstep.ProcessStepName <> "Shipping" Then
                cboLocation_CBG.Enabled = True
                cboProcessStepLocation.Enabled = True
                cboMatchedFacility.Enabled = True
                cboSubFacility.Enabled = True
                cboCDSID.Enabled = True
                cboLocation_CBG.Text = _SelectedProcessstep.FacilityCbg
                cboProcessStepLocation.Text = _SelectedProcessstep.FacilityLocation
                cboMatchedFacility.Text = _SelectedProcessstep.FacilityName
                cboSubFacility.Text = _SelectedProcessstep.SubFacilityName
                cboCDSID.Text = _SelectedProcessstep.Cdsid
            Else
                cboLocation_CBG.SelectedIndex = -1
                cboProcessStepLocation.SelectedIndex = -1
                cboMatchedFacility.SelectedIndex = -1
                cboSubFacility.SelectedIndex = -1
                cboCDSID.SelectedIndex = -1
                cboCDSID.Enabled = False
                cboLocation_CBG.Enabled = False
                cboProcessStepLocation.Enabled = False
                cboMatchedFacility.Enabled = False
                cboSubFacility.Enabled = False
            End If

            txtRemarks.Text = _SelectedProcessstep.Remarks

            dtStart.Text = _SelectedProcessstep.PlannedStart

            dtEnd.Text = _SelectedProcessstep.PlannedEnd

            dtInitEndDate = dtStart.Text
            dtInitEndDate = dtEnd.Text

            txtDuration.Text = _SelectedProcessstep.Duration.ToString()
            lblGlobal.Text = _SelectedProcessstep.GlobalDVP
            lblUser.Text = _SelectedProcessstep.TeamName

            If _SelectedProcessstep.WorkingDays = 5 Then
                opt5Days.Checked = True
            ElseIf _SelectedProcessstep.WorkingDays = 6 Then
                opt6Days.Checked = True
            ElseIf _SelectedProcessstep.WorkingDays = 7 Then
                opt7Days.Checked = True
            Else
                opt5Days.Checked = True
            End If

            chkHolidays.Checked = _SelectedProcessstep.IsWithHoliday

            bolRecordChanged = False
            cmdUpdate.Enabled = True
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEditUsercase, ex.Message))
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboLocation_CBG_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLocation_CBG.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.AppStarting
            cboProcessStepLocation.DataSource = Nothing
            cboMatchedFacility.DataSource = Nothing
            cboSubFacility.DataSource = Nothing

            If cboLocation_CBG.Text <> "" Then
                myDatatable = _Editprocess.GetLocation(FacilityCbg:=cboLocation_CBG.Text, FacilityLocation:=Nothing, FacilityName:=Nothing, SubFacilityName:=Nothing)
                myView = New System.Data.DataView(myDatatable)
                myDatatable = myView.ToTable(False, "FacilityLocation")

                Dim dr As System.Data.DataRow = myDatatable.NewRow
                dr("FacilityLocation") = "Select Location"
                myDatatable.Rows.InsertAt(dr, 0)

                cboProcessStepLocation.DataSource = myDatatable
                cboProcessStepLocation.DisplayMember = "FacilityLocation"
                cboProcessStepLocation.ValueMember = "FacilityLocation"
            End If
            bolRecordChanged = True
        Catch ex As Exception
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboProcessStepLocation_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboProcessStepLocation.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.AppStarting
            cboMatchedFacility.DataSource = Nothing
            cboSubFacility.DataSource = Nothing
            If cboLocation_CBG.Text <> "" And cboProcessStepLocation.Text <> "" Then
                myDatatable = _Editprocess.GetName(FacilityCbg:=cboLocation_CBG.Text, FacilityLocation:=cboProcessStepLocation.Text, FacilityName:=Nothing, SubFacilityName:=Nothing)
                myView = New System.Data.DataView(myDatatable)
                myDatatable = myView.ToTable(False, "FacilityName")

                Dim dr As System.Data.DataRow = myDatatable.NewRow
                dr("FacilityName") = "Select Facility"
                myDatatable.Rows.InsertAt(dr, 0)

                cboMatchedFacility.DataSource = myDatatable
                cboMatchedFacility.DisplayMember = "FacilityName"
                cboMatchedFacility.ValueMember = "FacilityName"
            End If
            bolRecordChanged = True
        Catch ex As Exception

        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboMatchedFacility_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMatchedFacility.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.AppStarting
            cboSubFacility.DataSource = Nothing
            If cboLocation_CBG.Text <> "" And cboProcessStepLocation.Text <> "" And cboMatchedFacility.Text <> "" Then
                myDatatable = _Editprocess.GetSubName(FacilityCbg:=cboLocation_CBG.Text, FacilityLocation:=cboProcessStepLocation.Text, FacilityName:=cboMatchedFacility.Text, SubFacilityName:=Nothing)
                myView = New System.Data.DataView(myDatatable)
                myDatatable = myView.ToTable(False, "SubFacilityName")

                Dim dr As System.Data.DataRow = myDatatable.NewRow
                dr("SubFacilityName") = "Select Sub Facility"
                myDatatable.Rows.InsertAt(dr, 0)

                cboSubFacility.DataSource = myDatatable
                cboSubFacility.DisplayMember = "SubFacilityName"
                cboSubFacility.ValueMember = "SubFacilityName"
            End If
            bolRecordChanged = True
        Catch ex As Exception

        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboSubFacility_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSubFacility.SelectedIndexChanged
        bolRecordChanged = True
    End Sub

    Private Sub cboCDSID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCDSID.SelectedIndexChanged
        bolRecordChanged = True
    End Sub

    Private Sub txtRemarks_TextChanged(sender As Object, e As EventArgs) Handles txtRemarks.TextChanged
        bolRecordChanged = True
    End Sub

    Private Sub chkStart_CheckedChanged(sender As Object, e As EventArgs) Handles chkStart.CheckedChanged
        If bolSkipEvent Then Exit Sub

        If chkStart.Checked = True Then

            bolSkipEvent = True

            dtStart.Enabled = True
            chkStart.Enabled = True

            dtEnd.Visible = False
            chkEnd.Visible = False
            lblEnddate.Visible = False
            lblKWEnd.Visible = False

            txtDuration.Visible = True
            LblWorkingdays.Visible = True
            lblDuration.Visible = True
            'optWeeks.Visible = True
            'optWorkingDays.Visible = True

            chkEnd.Checked = False

            txtDuration.Enabled = True
            optWeeks.Enabled = True

            optWorkingDays.Enabled = True
            optWorkingDays.Checked = True


            If chkStart.Checked = True Or chkEnd.Checked = True Then
                opt5Days.Enabled = optWorkingDays.Checked
                opt6Days.Enabled = optWorkingDays.Checked
                opt7Days.Enabled = optWorkingDays.Checked
                chkHolidays.Enabled = optWorkingDays.Checked
            Else
                opt5Days.Enabled = False
                opt6Days.Enabled = False
                opt7Days.Enabled = False
                chkHolidays.Enabled = False
            End If

            bolSkipEvent = False
        Else
            dtEnd.Visible = True
            chkEnd.Visible = True
            lblEnddate.Visible = True
            lblKWEnd.Visible = True

            bolSkipEvent = True
            chkStartAndEnd.Checked = False
            bolSkipEvent = False

            dtStart.Enabled = False

            If chkEnd.Checked = True Then
                txtDuration.Visible = False
                LblWorkingdays.Visible = False
                lblDuration.Visible = False
                'optWeeks.Visible = False
                'optWorkingDays.Visible = False
            Else
                txtDuration.Visible = True
                LblWorkingdays.Visible = True
                lblDuration.Visible = True
                'optWeeks.Visible = True
                'optWorkingDays.Visible = True
            End If
        End If

        If chkEnd.Checked = False And chkStart.Checked = False Then
            bolSkipEvent = True

            chkEnd.Enabled = True '
            chkStart.Enabled = True '

            dtStart.Enabled = False
            dtEnd.Enabled = False

            txtDuration.Enabled = False
            optWeeks.Enabled = False

            optWorkingDays.Enabled = False
            optWorkingDays.Checked = False

            If chkStart.Checked = True Or chkEnd.Checked = True Then
                opt5Days.Enabled = optWorkingDays.Checked
                opt6Days.Enabled = optWorkingDays.Checked
                opt7Days.Enabled = optWorkingDays.Checked
                chkHolidays.Enabled = optWorkingDays.Checked
            Else
                opt5Days.Enabled = False
                opt6Days.Enabled = False
                opt7Days.Enabled = False
                chkHolidays.Enabled = False
            End If

            bolSkipEvent = False
        End If
    End Sub

    Private Sub chkEnd_CheckedChanged(sender As Object, e As EventArgs) Handles chkEnd.CheckedChanged
        Try
            If bolSkipEvent Then Exit Sub

            If chkEnd.Checked = True Then

                txtDuration.Visible = False
                LblWorkingdays.Visible = False
                lblDuration.Visible = False
                'optWeeks.Visible = False
                'optWorkingDays.Visible = False

                dtEnd.Enabled = True
                chkEnd.Enabled = True

                bolSkipEvent = True
                chkStart.Checked = False

                bolSkipEvent = False

                dtStart.Enabled = False
                chkStart.Enabled = False

                txtDuration.Enabled = True
                optWeeks.Enabled = True

                optWorkingDays.Enabled = True
                optWorkingDays.Checked = True

                If chkStart.Checked = True Or chkEnd.Checked = True Then
                    opt5Days.Enabled = optWorkingDays.Checked
                    opt6Days.Enabled = optWorkingDays.Checked
                    opt7Days.Enabled = optWorkingDays.Checked
                    chkHolidays.Enabled = optWorkingDays.Checked
                Else
                    opt5Days.Enabled = False
                    opt6Days.Enabled = False
                    opt7Days.Enabled = False
                    chkHolidays.Enabled = False
                End If
            Else
                txtDuration.Visible = True
                LblWorkingdays.Visible = True
                lblDuration.Visible = True
                txtDuration.Enabled = True

                'optWeeks.Visible = True
                optWeeks.Enabled = True
                'optWorkingDays.Visible = True
                optWorkingDays.Enabled = True

                chkStart.Enabled = True
                bolSkipEvent = True
                chkStartAndEnd.Checked = False
                bolSkipEvent = False

                dtEnd.Enabled = False

                If chkStart.Checked = True Then
                    dtStart.Enabled = True

                    dtEnd.Visible = False
                    chkEnd.Visible = False
                    lblEnddate.Visible = False
                    lblKWEnd.Visible = False
                End If
            End If

            If chkEnd.Checked = False And chkStart.Checked = False Then
                bolSkipEvent = True

                dtStart.Enabled = False
                dtEnd.Enabled = False

                txtDuration.Enabled = False
                optWeeks.Enabled = False

                optWorkingDays.Enabled = False
                optWorkingDays.Checked = False

                If chkStart.Checked = True Or chkEnd.Checked = True Then
                    opt5Days.Enabled = optWorkingDays.Checked
                    opt6Days.Enabled = optWorkingDays.Checked
                    opt7Days.Enabled = optWorkingDays.Checked
                    chkHolidays.Enabled = optWorkingDays.Checked
                Else
                    opt5Days.Enabled = False
                    opt6Days.Enabled = False
                    opt7Days.Enabled = False
                    chkHolidays.Enabled = False
                End If

                bolSkipEvent = False
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEditUsercase, ex.Message))
        End Try
    End Sub

    Private Sub txtDuration_TextChanged(sender As Object, e As EventArgs) Handles txtDuration.TextChanged
        bolRecordChanged = True
    End Sub

    Private Sub txtDuration_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDuration.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtDuration_KeyUp(sender As Object, e As KeyEventArgs) Handles txtDuration.KeyUp
        sbChangeDates()
    End Sub

    Private Sub optWeeks_CheckedChanged(sender As Object, e As EventArgs) Handles optWeeks.CheckedChanged
        bolRecordChanged = True
    End Sub

    Private Sub optWeeks_Click(sender As Object, e As EventArgs) Handles optWeeks.Click
        If optWeeks.Checked = True Then
            txtDuration.Text = DateDiff("ww", CDate(dtStart.Text), CDate(dtEnd.Text), vbMonday, vbFirstFourDays)
        Else
            optWorkingDays_Click(sender, e)
        End If
    End Sub

    Private Sub optWorkingDays_Click(sender As Object, e As EventArgs) Handles optWorkingDays.Click
        'On Error Resume Next
        'If opt5Days.Checked = True Then
        '    txtDuration.Text = _GlobalFunctions.fnGetDuration(CDate(txtStartDate.Text), CDate(txtEndDate.Text), 5)
        'ElseIf opt6Days.Checked = True Then
        '    txtDuration.Text = _GlobalFunctions.fnGetDuration(CDate(txtStartDate.Text), CDate(txtEndDate.Text), 6)
        'Else
        '    txtDuration.Text = _GlobalFunctions.fnGetDuration(CDate(txtStartDate.Text), CDate(txtEndDate.Text), 7)
        'End If
    End Sub

    Private Sub chkStartAndEnd_CheckedChanged(sender As Object, e As EventArgs) Handles chkStartAndEnd.CheckedChanged
        If bolSkipEvent Then Exit Sub
        If chkStartAndEnd.Checked = True Then
            bolSkipEvent = True

            txtDuration.Visible = False
            LblWorkingdays.Visible = False
            lblDuration.Visible = False
            'optWeeks.Visible = False
            'optWorkingDays.Visible = False

            dtEnd.Visible = True
            chkEnd.Visible = True
            lblEnddate.Visible = True
            lblKWEnd.Visible = True

            chkStart.Checked = True
            chkEnd.Checked = True
            chkStart.Enabled = True
            chkEnd.Enabled = True

            dtStart.Enabled = True
            dtEnd.Enabled = True

            txtDuration.Enabled = False
            optWeeks.Enabled = False

            optWorkingDays.Enabled = True
            optWorkingDays.Checked = True


            If chkStart.Checked = True Or chkEnd.Checked = True Then
                opt5Days.Enabled = optWorkingDays.Checked
                opt6Days.Enabled = optWorkingDays.Checked
                opt7Days.Enabled = optWorkingDays.Checked
                chkHolidays.Enabled = optWorkingDays.Checked
            Else
                opt5Days.Enabled = False
                opt6Days.Enabled = False
                opt7Days.Enabled = False
                chkHolidays.Enabled = False
            End If

            txtDuration.Enabled = False
            optWeeks.Enabled = False

            bolSkipEvent = False
        Else
            txtDuration.Visible = True
            LblWorkingdays.Visible = True
            lblDuration.Visible = True
            'optWeeks.Visible = True
            'optWorkingDays.Visible = True

            bolSkipEvent = True

            chkStart.Checked = False
            chkEnd.Checked = False
            chkStart.Enabled = True
            chkEnd.Enabled = True

            dtEnd.Enabled = False
            dtStart.Enabled = False

            txtDuration.Enabled = False
            optWeeks.Enabled = False

            optWorkingDays.Enabled = False
            optWorkingDays.Checked = False


            If chkStart.Checked = True Or chkEnd.Checked = True Then
                opt5Days.Enabled = optWorkingDays.Checked
                opt6Days.Enabled = optWorkingDays.Checked
                opt7Days.Enabled = optWorkingDays.Checked
                chkHolidays.Enabled = optWorkingDays.Checked
            Else
                opt5Days.Enabled = False
                opt6Days.Enabled = False
                opt7Days.Enabled = False
                chkHolidays.Enabled = False
            End If

            bolSkipEvent = False
        End If
    End Sub

    Private Sub opt5Days_CheckedChanged(sender As Object, e As EventArgs)
        bolRecordChanged = True
    End Sub

    Private Sub opt6Days_CheckedChanged(sender As Object, e As EventArgs)
        bolRecordChanged = True
    End Sub

    Private Sub opt7Days_CheckedChanged(sender As Object, e As EventArgs)
        bolRecordChanged = True
    End Sub

    Private Sub opt7Days_Click(sender As Object, e As EventArgs)
        sbChangeDates()
    End Sub

    Private Sub opt6Days_Click(sender As Object, e As EventArgs)
        sbChangeDates()
    End Sub

    Private Sub opt5Days_Click(sender As Object, e As EventArgs)
        sbChangeDates()
    End Sub

    Private Sub sbChangeDates()
        bolChangeFlag = True
        If chkEnd.Checked = False Then
            If optWeeks.Checked = True Then
                dtEnd.Text = Strings.Format(DateAdd("d", -1, DateAdd("ww", Val(txtDuration.Text), CDate(dtStart.Text))), "dd-MMM-yyyy")
            Else
                If opt5Days.Checked = True Then
                    dtEnd.Text = Strings.Format(Date.FromOADate(Globals.ThisAddIn.Application.WorksheetFunction.WorkDay_Intl(CDate(dtStart.Text), Val(txtDuration.Text) - 1, 1)), "dd-MMM-yyyy")
                ElseIf opt6Days.Checked = True Then

                    dtEnd.Text = Strings.Format(Date.FromOADate(Globals.ThisAddIn.Application.WorksheetFunction.WorkDay_Intl(CDate(dtStart.Text), Val(txtDuration.Text) - 1, 11)), "dd-MMM-yyyy")
                Else
                    dtEnd.Text = Strings.Format(DateAdd("d", Val(txtDuration.Text) - 1, CDate(dtStart.Text)), "dd-MMM-yyyy")
                End If
            End If
        ElseIf chkStart.Checked = False Then
            If optWeeks.Checked = True Then
                dtStart.Text = Strings.Format(DateAdd("d", 1, DateAdd("ww", Val(txtDuration.Text) * -1, CDate(dtEnd.Text))), "dd-MMM-yyyy")
            Else
                If opt5Days.Checked = True Then
                    dtStart.Text = Strings.Format(Date.FromOADate(Globals.ThisAddIn.Application.WorksheetFunction.WorkDay_Intl(CDate(dtEnd.Text), (Val(txtDuration.Text) - 1) * -1, 1)), "dd-MMM-yyyy")
                ElseIf opt6Days.Checked = True Then
                    dtStart.Text = Strings.Format(Date.FromOADate(Globals.ThisAddIn.Application.WorksheetFunction.WorkDay_Intl(CDate(dtEnd.Text), (Val(txtDuration.Text) - 1) * -1, 11)), "dd-MMM-yyyy")
                Else
                    dtStart.Text = Strings.Format(DateAdd("d", Val(txtDuration.Text - 1) * -1, CDate(dtEnd.Text)), "dd-MMM-yyyy")
                End If
            End If
        End If
        bolChangeFlag = False
    End Sub

    Private Sub cmdReset_Click(sender As Object, e As EventArgs) Handles cmdReset.Click
        ClearForm()
        Fill_dgvUsercase()
        PopulateData(0)
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub


    'Checks invalid characters (", ', ;) in a string 
    Private Function ContainsInvalidChar(strValue As String) As Boolean
        ContainsInvalidChar = False
        If Strings.InStr(1, strValue, "'") > 0 Or Strings.InStr(1, strValue, """") > 0 Or Strings.InStr(1, strValue, ";") > 0 Then
            ContainsInvalidChar = True
        End If
    End Function

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        Try
            Me.Cursor = Cursors.WaitCursor


            '------------------------------
            ' Validaion
            '------------------------------
            If ContainsInvalidChar(cboCDSID.Text) Then
                DialogResult = DialogResult.None
                MessageBox.Show("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If ContainsInvalidChar(txtRemarks.Text) Then
                DialogResult = DialogResult.None
                MessageBox.Show("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If cboProcessStepLocation.SelectedIndex = 0 Then
                MessageBox.Show("Please select process step location.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If cboMatchedFacility.SelectedIndex = 0 Then
                MessageBox.Show("Please select matched facility.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If cboSubFacility.SelectedIndex = 0 Then
                MessageBox.Show("Please select sub facility.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If CDate(dtStart.Value) > CDate(dtEnd.Value) Then Throw New Exception("End date cannot be earlier than start date")
            If CDate(dtStart.Value) > dtInitEndDate Then Throw New Exception("Sorry, the start date is out of scope from the previous end date. Please use Copy/Cut/Paste functionality.")
            If CDate(dtEnd.Value) < dtInitStartDate Then Throw New Exception("Sorry, the end date is out of scope from the previous start date. Please use Copy/Cut/Paste functionality.")

            If CDate(dtStart.Text) < CDate(Form.DataCenter.GlobalValues.WS.Range(_GlobalFunctions.ColumnLetter(Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn + 1).ToString & 4).Value) Then
                Throw New Exception("Start date should be in the timeline section dates.")
            End If
            If CDate(dtEnd.Text) > CDate(Form.DataCenter.GlobalValues.WS.Range(_GlobalFunctions.ColumnLetter(Form.DataCenter.GlobalSections.TimeLineSectionLastColumn - 1).ToString & 4).Value) Then
                Throw New Exception("End date should be in the timeline section dates.")
            End If

            Dim intWorkDays As Integer = 0
            If opt5Days.Checked = True Then
                intWorkDays = 5
            ElseIf opt6Days.Checked = True Then
                intWorkDays = 6
            ElseIf opt7Days.Checked = True Then
                intWorkDays = 7
            End If

            Dim strMatchedFacility As String = cboMatchedFacility.Text
            Dim strProcessStepLocation As String = cboProcessStepLocation.Text
            Dim strLocation_CBG As String = cboLocation_CBG.Text
            Dim strSubFacility As String = cboSubFacility.Text

            If Form.DataCenter.ProcessStepConfig.PSIsGapOrDelay Then
                strMatchedFacility = Nothing
                strProcessStepLocation = Nothing
                strLocation_CBG = Nothing
                strSubFacility = Nothing
            End If

            '------------------------------
            ' apply changes to DB
            '------------------------------
            If chkStart.Checked = True And chkEnd.Checked = True And txtDuration.Visible = False Then
                If (_SelectedProcessstep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, dbl_pe26_SpecificVehicleUsercases_PK, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _SelectedProcessstep.AllocatedUsercaseSeq, _SelectedProcessstep.ProcessStepSeq, CDate(dtStart.Text), CDate(dtEnd.Text), Nothing, intWorkDays, RemoveSPChars(cboCDSID.Text), RemoveSPChars(txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf chkStart.Checked = False And chkEnd.Checked = True And txtDuration.Visible = False Then
                If (_SelectedProcessstep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, dbl_pe26_SpecificVehicleUsercases_PK, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _SelectedProcessstep.AllocatedUsercaseSeq, _SelectedProcessstep.ProcessStepSeq, Nothing, CDate(dtEnd.Text), Nothing, intWorkDays, RemoveSPChars(cboCDSID.Text), RemoveSPChars(txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf chkStart.Checked = True And chkEnd.Visible = False And txtDuration.Enabled = True Then
                If (_SelectedProcessstep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, dbl_pe26_SpecificVehicleUsercases_PK, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _SelectedProcessstep.AllocatedUsercaseSeq, _SelectedProcessstep.ProcessStepSeq, CDate(dtStart.Text), Nothing, Integer.Parse(txtDuration.Text), intWorkDays, RemoveSPChars(cboCDSID.Text), RemoveSPChars(txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf chkStart.Checked = False And chkEnd.Checked = False Then
                If (_SelectedProcessstep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, dbl_pe26_SpecificVehicleUsercases_PK, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _SelectedProcessstep.AllocatedUsercaseSeq, _SelectedProcessstep.ProcessStepSeq, Nothing, Nothing, Nothing, intWorkDays, RemoveSPChars(cboCDSID.Text), RemoveSPChars(txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
            End If



            bolWasUpdated = True

            '------------------------------
            ' Keep current selected
            '------------------------------
            Dim SelectedRowsIndex = dgvUsercase.SelectedRows(0).Index

            '------------------------------
            ' Fill grid with db values
            '------------------------------
            Fill_dgvUsercase()

            '------------------------------
            ' Go to the selected row.
            '------------------------------
            dgvUsercase.Rows(SelectedRowsIndex).Selected = True

            MessageBox.Show("Data successfully updated...", "Edit usercase", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEditUsercase, ex.Message), "Edit ProcessStep", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub frmEditUsercase_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            cmdUpdate_Click(sender, e)
        ElseIf e.KeyCode = Keys.F5 Then
            cmdReset_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            cboLocation_CBG.Focus()
        End If
    End Sub

    Private Sub dtStart_ValueChanged(sender As Object, e As EventArgs) Handles dtStart.ValueChanged
        If bolChangeFlag Then Exit Sub
        If (dtStart.Text <> "" And dtEnd.Text <> "") Then
            lblKWSt.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtStart.Text, Form.DataCenter.GlobalValues.myCWR, vbMonday)
            If optWorkingDays.Checked = True Then
                If opt5Days.Checked = True Then
                    txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 5)
                ElseIf opt6Days.Checked = True Then
                    txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 6)
                Else
                    txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 7)
                End If
            Else
                txtDuration.Text = DateDiff("ww", CDate(dtStart.Text), CDate(dtEnd.Text), vbMonday, vbFirstFourDays)
            End If
        End If
    End Sub

    Private Sub dtEnd_ValueChanged(sender As Object, e As EventArgs) Handles dtEnd.ValueChanged
        If bolChangeFlag Then Exit Sub
        lblKWEnd.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtEnd.Text, Form.DataCenter.GlobalValues.myCWR, vbMonday)
        If optWorkingDays.Checked = True Then
            If opt5Days.Checked = True Then
                txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 5)
            ElseIf opt6Days.Checked = True Then
                txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 6)
            Else
                txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 7)
            End If
        Else
            txtDuration.Text = DateDiff("ww", CDate(dtStart.Text), CDate(dtEnd.Text), vbMonday, vbFirstFourDays)
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub frmEditUsercase_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'Dim Cls As New Form.DataCenter.GlobalFunctions
        If bolWasUpdated Then _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,, Form.DataCenter.ProcessStepConfig.ProcessStepPe26)

    End Sub

    'Private Sub cmdUpdateAllCDSID_Click(sender As Object, e As EventArgs)
    '    If ContainsInvalidChar(cboCDSID.Text) = True Then
    '        MessageBox.Show("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        DialogResult = DialogResult.None
    '        Exit Sub
    '    End If
    '    If ContainsInvalidChar(txtRemarks.Text) = True Then
    '        MessageBox.Show("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        DialogResult = DialogResult.None
    '        Exit Sub
    '    End If
    '    If MessageBox.Show("Do you want to update all CDSID for this Process", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
    '        Exit Sub
    '    End If
    '    If Strings.Trim(cboCDSIDAll.Text) = "" And Strings.Trim(txtRemarksAll.Text) = "" Then
    '        If MessageBox.Show("CDSID and Remarks is empty.Do you want to update anyway?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
    '            Exit Sub
    '        End If
    '    ElseIf Strings.Trim(cboCDSIDAll.Text) = "" Then
    '        If MessageBox.Show("CDSID is empty.Do you want to update anyway CDSID as Empty?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
    '            Exit Sub
    '        End If
    '    ElseIf Strings.Trim(txtRemarksAll.Text) = "" Then
    '        If MessageBox.Show("Remarks is empty.Do you want to update anyway Remarks as Empty?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
    '            Exit Sub
    '        End If
    '    End If

    '    Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    '    MessageBox.Show("Data sucessfully updated...", "Edit usercase", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    cboCDSIDAll.Text = ""
    '    txtRemarksAll.Text = ""

    'End Sub

    Private Function fnCheckIsNull(rstField As String) As String
        If Not IsDBNull(rstField) Then
            fnCheckIsNull = rstField
        Else
            fnCheckIsNull = "" ' vbNullString
        End If
    End Function

    Private Function RemoveSPChars(strString As String) As String
        Dim strValue As String
        strValue = fnCheckIsNull(strString)
        strValue = Strings.Replace(Strings.Replace(Strings.Replace(strValue, "'", ""), """", ""), ";", "")
        RemoveSPChars = strValue
    End Function

    Private Sub UPdateAllCDSIDandRemarks(bCDSID As Boolean, bRemarks As Boolean)

        Dim _Usercase As CT.Data.Usercase = New Data.Usercase

        If bCDSID = False And bRemarks = False Then
            Exit Sub
        End If

        Try

            If bCDSID = True Then
                If (_Usercase.EditCDSID(Form.DataCenter.VehicleConfig.VehiclePe02, Form.DataCenter.VehicleConfig.VehiclePe45, _AllocatedUsercaseSequence, RemoveSPChars(Strings.Trim(cboCDSIDAll.Text)), Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                bolWasUpdated = True
            End If
            If bRemarks = True Then
                If (_Usercase.EditRemarks(Form.DataCenter.VehicleConfig.VehiclePe02, Form.DataCenter.VehicleConfig.VehiclePe45, _AllocatedUsercaseSequence, RemoveSPChars(Strings.Trim(txtRemarksAll.Text)), Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                bolWasUpdated = True
            End If
        Catch
            MessageBox.Show("Update CDSID/ Remark", CT.Data.DataCenter.GlobalValues.message, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdResetAllCDSID_Click(sender As Object, e As EventArgs)
        cboCDSIDAll.Text = ""
        txtRemarksAll.Text = ""
    End Sub

    Private Sub cmdUpdateAllCDSID_Click(sender As Object, e As EventArgs) Handles cmdUpdateAllCDSID.Click
        Try
            cmdUpdateAllCDSID.Enabled = False
            If ContainsInvalidChar(cboCDSIDAll.Text) = True Then
                MessageBox.Show("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                DialogResult = DialogResult.None
                cmdUpdateAllCDSID.Enabled = True
                Exit Sub
            End If
            If ContainsInvalidChar(txtRemarksAll.Text) = True Then
                MessageBox.Show("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                DialogResult = DialogResult.None
                cmdUpdateAllCDSID.Enabled = True
                Exit Sub
            End If

            If Strings.Trim(cboCDSIDAll.Text) = "" And Strings.Trim(txtRemarksAll.Text) = "" Then
                If MessageBox.Show("CDSID and Remarks is empty.Do you want to update anyway as Empty?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    UPdateAllCDSIDandRemarks(True, True)
                Else
                    cmdUpdateAllCDSID.Enabled = True
                    Exit Sub
                End If
            ElseIf Strings.Trim(cboCDSIDAll.Text) = "" Then
                DialogResult = MessageBox.Show("CDSID is empty.Do you want to update anyway CDSID as Empty?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If DialogResult = DialogResult.Yes Then
                    UPdateAllCDSIDandRemarks(True, True)
                    DialogResult = DialogResult.None
                ElseIf DialogResult = DialogResult.No Then
                    UPdateAllCDSIDandRemarks(False, True)
                    DialogResult = DialogResult.None
                Else
                    DialogResult = DialogResult.None
                    cmdUpdateAllCDSID.Enabled = True
                    Exit Sub
                End If
            ElseIf Strings.Trim(txtRemarksAll.Text) = "" Then
                DialogResult = MessageBox.Show("Remarks is empty.Do you want to update anyway Remarks as Empty?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If DialogResult = DialogResult.Yes Then
                    DialogResult = DialogResult.None
                    UPdateAllCDSIDandRemarks(True, True)
                ElseIf DialogResult = DialogResult.No Then
                    UPdateAllCDSIDandRemarks(True, False)
                    DialogResult = DialogResult.None
                Else
                    DialogResult = DialogResult.None
                    cmdUpdateAllCDSID.Enabled = True
                    Exit Sub
                End If
            Else
                If MessageBox.Show("Do you want to update all CDSID and Remarks for this Process", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    UPdateAllCDSIDandRemarks(True, True)
                Else
                    cmdUpdateAllCDSID.Enabled = True
                    Exit Sub
                End If
            End If
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MessageBox.Show("Data sucessfully updated...", "Edit usercase", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmdUpdateAllCDSID.Enabled = True
            cboCDSIDAll.Text = ""
            txtRemarksAll.Text = ""

            '------------------------------
            ' Keep current selected
            '------------------------------
            Dim SelectedRowsIndex = dgvUsercase.SelectedRows(0).Index

            '------------------------------
            ' Fill grid with db values
            '------------------------------
            Fill_dgvUsercase()

            '------------------------------
            ' Go to the selected row.
            '------------------------------
            dgvUsercase.Rows(SelectedRowsIndex).Selected = True
            PopulateData(SelectedRowsIndex)


        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEditUsercase, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            Me.Cursor = Cursors.Default

        End Try
    End Sub
End Class