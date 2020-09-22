Imports System.Windows.Forms

Public Class frmEdit_Rig
    'Dim DontClose As Boolean = False
    Dim _Facility As Data.Facility = New Data.Facility()
    Dim myDatatable As New System.Data.DataTable
    Dim myView As System.Data.DataView

    Dim IsRecordChanged As Boolean = False
    Dim bolSkipEvent As Boolean
    Dim bolChangeFlag As Boolean

    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions

    Dim Facility As CT.Data.Facility
    Private _pe26 As Long

    Public Property pe26 As Long
        Get
            pe26 = _pe26
        End Get
        Set(value As Long)
            _pe26 = value
        End Set
    End Property

    Private _UsercaseSeq As Integer
    Private _ProcessStepSeq As Integer

    Private Sub SetCboLocation_CBG()
        myView = New System.Data.DataView(_Facility.GetCbg(FacilityCbg:=Nothing, FacilityLocation:=Nothing, FacilityName:=Nothing, SubFacilityName:=Nothing))
        myDatatable = myView.ToTable(False, "FacilityCbg")
        cboLocation_CBG.DataSource = myDatatable
        cboLocation_CBG.DisplayMember = "FacilityCbg"
        cboLocation_CBG.ValueMember = "FacilityCbg"
    End Sub

    Private Sub frmEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If _pe26 = 0 Then
                Throw New Exception("ProcessStep must be selected.")
            End If

            Me.Cursor = Cursors.WaitCursor

            _Facility = New Data.Facility()


            If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                cmdOk.Enabled = False
            End If


            Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep()

            _ProcessStep.SelectProcessStepDedicated(_pe26, Form.DataCenter.ProgramConfig.IsGeneric)
            _UsercaseSeq = _ProcessStep.AllocatedUsercaseSeq
            _ProcessStepSeq = _ProcessStep.ProcessStepSeq

            Dim dtCDSID As System.Data.DataTable
            dtCDSID = _ProcessStep.GetAllCdsids(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, _ProcessStep.GlobalDVP, Form.DataCenter.ProgramConfig.BuildType)
            If dtCDSID Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


            If dtCDSID.Rows.Count > 0 Then
                For i As Int16 = 0 To dtCDSID.Rows.Count - 1
                    cboCDSID.Items.Add(dtCDSID.Rows(i).Item(0).ToString())
                Next
            End If

            If Form.DataCenter.ProcessStepConfig.ProcessStepUserCase = "Gap" Then

            End If

            If Form.DataCenter.ProcessStepConfig.PSIsGapOrDelay = False Then
                SetCboLocation_CBG()
                cboLocation_CBG.Text = _ProcessStep.FacilityCbg
                cboProcessStepLocation.Text = _ProcessStep.FacilityLocation
                cboMatchedFacility.Text = _ProcessStep.FacilityName
                cboSubFacility.Text = _ProcessStep.SubFacilityName
                cboCDSID.Text = _ProcessStep.Cdsid
            End If

            lblGlobal.Text = _ProcessStep.GlobalDVP
            lblUser.Text = _ProcessStep.TeamName
            lblUserCase.Text = _ProcessStep.Usercase
            lblProcessStep.Text = _ProcessStep.ProcessStepName



            txtRemarks.Text = _ProcessStep.Remarks

            dtStart.CustomFormat = "dd-MMM-yyyy"
            dtStart.Value = _ProcessStep.PlannedStart

            dtEnd.CustomFormat = "dd-MMM-yyyy"
            dtEnd.Value = _ProcessStep.PlannedEnd

            txtDuration.Text = _ProcessStep.Duration.ToString()

            If _ProcessStep.WorkingDays = 5 Then
                opt5Days.Checked = True
            ElseIf _ProcessStep.WorkingDays = 6 Then
                opt6Days.Checked = True
            ElseIf _ProcessStep.WorkingDays = 7 Then
                opt7Days.Checked = True
            Else
                opt5Days.Checked = True
            End If


            chkHolidays.Checked = _ProcessStep.IsWithHoliday

            If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                cmdOk.Enabled = True

                If Form.DataCenter.ProcessStepConfig.PSIsGapOrDelay = False Then
                    cboLocation_CBG.Enabled = True
                    cboProcessStepLocation.Enabled = True
                    cboMatchedFacility.Enabled = True
                    cboSubFacility.Enabled = True
                    cboCDSID.Enabled = True
                Else
                    cboLocation_CBG.Enabled = False
                    cboProcessStepLocation.Enabled = False
                    cboMatchedFacility.Enabled = False
                    cboSubFacility.Enabled = False
                    cboCDSID.Enabled = False
                End If

                '-----------------------------------------------------------------------------
                ' They must be disabled only the items, which are allowed to change must be enable.
                '-----------------------------------------------------------------------------
                txtDuration.Enabled = False
                dtStart.Enabled = False
                dtEnd.Enabled = False
                '-----------------------------------------------------------------------------
            Else
                cmdOk.Enabled = False

                cboLocation_CBG.Enabled = False
                cboProcessStepLocation.Enabled = False
                cboMatchedFacility.Enabled = False
                cboSubFacility.Enabled = False
                cboCDSID.Enabled = False

                txtDuration.Enabled = False
                dtStart.Enabled = False
                dtEnd.Enabled = False
            End If
            '--------------------------------------------------------------------------
            ' Checked and Enabled are together because 5,6,7 days are working 
            ' according to  optWorkingDays.Checked
            '--------------------------------------------------------------------------
            optWorkingDays.Checked = True
            optWorkingDays.Enabled = False

            optWeeks.Enabled = False

            IsRecordChanged = False
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEdit, ex.Message), "Edit ProcessStep", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboLocation_CBG_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLocation_CBG.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.WaitCursor
            cboProcessStepLocation.DataSource = Nothing
            cboMatchedFacility.DataSource = Nothing
            cboSubFacility.DataSource = Nothing

            If cboLocation_CBG.Text <> "" Then
                myDatatable = _Facility.GetLocation(FacilityCbg:=cboLocation_CBG.Text, FacilityLocation:=Nothing, FacilityName:=Nothing, SubFacilityName:=Nothing)

                myView = New System.Data.DataView(myDatatable)
                myDatatable = myView.ToTable(False, "FacilityLocation")

                Dim dr As System.Data.DataRow = myDatatable.NewRow
                dr("FacilityLocation") = "Select Location"
                myDatatable.Rows.InsertAt(dr, 0)

                cboProcessStepLocation.DataSource = myDatatable
                cboProcessStepLocation.DisplayMember = "FacilityLocation"
                cboProcessStepLocation.ValueMember = "FacilityLocation"
            End If
            IsRecordChanged = True
        Catch ex As Exception

        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboProcessStepLocation_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboProcessStepLocation.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.WaitCursor
            cboMatchedFacility.DataSource = Nothing
            cboSubFacility.DataSource = Nothing
            If cboLocation_CBG.Text <> "" And cboProcessStepLocation.Text <> "" Then
                myDatatable = _Facility.GetName(FacilityCbg:=cboLocation_CBG.Text, FacilityLocation:=cboProcessStepLocation.Text, FacilityName:=Nothing, SubFacilityName:=Nothing)

                myView = New System.Data.DataView(myDatatable)
                myDatatable = myView.ToTable(False, "FacilityName")

                Dim dr As System.Data.DataRow = myDatatable.NewRow
                dr("FacilityName") = "Select Facility"
                myDatatable.Rows.InsertAt(dr, 0)

                cboMatchedFacility.DataSource = myDatatable
                cboMatchedFacility.DisplayMember = "FacilityName"
                cboMatchedFacility.ValueMember = "FacilityName"
            End If
            IsRecordChanged = True
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
                myDatatable = _Facility.GetSubName(FacilityCbg:=cboLocation_CBG.Text, FacilityLocation:=cboProcessStepLocation.Text, FacilityName:=cboMatchedFacility.Text, SubFacilityName:=Nothing)

                myView = New System.Data.DataView(myDatatable)
                myDatatable = myView.ToTable(False, "SubFacilityName")

                Dim dr As System.Data.DataRow = myDatatable.NewRow
                dr("SubFacilityName") = "Select Sub Facility"
                myDatatable.Rows.InsertAt(dr, 0)

                cboSubFacility.DataSource = myDatatable
                cboSubFacility.DisplayMember = "SubFacilityName"
                cboSubFacility.ValueMember = "SubFacilityName"
            End If
            IsRecordChanged = True
        Catch ex As Exception
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboSubFacility_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSubFacility.SelectedIndexChanged
        IsRecordChanged = True
    End Sub

    Private Sub cboCDSID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCDSID.SelectedIndexChanged
        IsRecordChanged = True
    End Sub

    Private Sub txtRemarks_TextChanged(sender As Object, e As EventArgs) Handles txtRemarks.TextChanged
        IsRecordChanged = True
    End Sub

    Private Sub chkStart_CheckedChanged(sender As Object, e As EventArgs) Handles chkStart.CheckedChanged
        If bolSkipEvent Then Exit Sub

        If chkStart.Checked = True Then

            bolSkipEvent = True

            dtStart.Enabled = True
            chkStart.Enabled = True

            dtEnd.Visible = False
            chkEnd.Visible = False
            lblEndDate.Visible = False
            lblKWEnd.Visible = False

            txtDuration.Visible = True
            LblWorkingdays.Visible = True
            lblDuration.Visible = True
            'optWeeks.Visible = True
            'optWorkingDays.Visible = True


            chkEnd.Checked = False

            dtEnd.Enabled = True
            chkEnd.Enabled = True

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
            lblEndDate.Visible = True
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

    Private Sub chkEnd_CheckedChanged(sender As Object, e As EventArgs) Handles chkEnd.CheckedChanged
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
                lblEndDate.Visible = False
                lblKWEnd.Visible = False
            End If
        End If

        If chkEnd.Checked = False And chkStart.Checked = False Then
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
        End If
    End Sub

    Private Sub txtDuration_TextChanged(sender As Object, e As EventArgs) Handles txtDuration.TextChanged
        IsRecordChanged = True
    End Sub

    Private Sub optWeeks_CheckedChanged(sender As Object, e As EventArgs) Handles optWeeks.CheckedChanged
        IsRecordChanged = True
    End Sub

    Private Sub chkStartAndEnd_CheckedChanged(sender As Object, e As EventArgs) Handles chkStartAndEnd.CheckedChanged
        If bolSkipEvent Then Exit Sub
        If chkStartAndEnd.Checked = True Then
            bolSkipEvent = True

            txtDuration.Visible = False
            LblWorkingdays.Visible = False
            lblDuration.Visible = False

            dtEnd.Visible = True
            chkEnd.Visible = True
            lblEndDate.Visible = True
            lblKWEnd.Visible = True

            chkStart.Checked = True
            chkEnd.Checked = True
            chkStart.Enabled = True
            chkEnd.Enabled = True

            dtEnd.Enabled = True
            dtStart.Enabled = True

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

    Private Sub opt5Days_CheckedChanged(sender As Object, e As EventArgs) Handles opt5Days.CheckedChanged
        IsRecordChanged = True
    End Sub

    Private Sub opt6Days_CheckedChanged(sender As Object, e As EventArgs) Handles opt6Days.CheckedChanged
        IsRecordChanged = True
    End Sub

    Private Sub opt7Days_CheckedChanged(sender As Object, e As EventArgs) Handles opt7Days.CheckedChanged
        IsRecordChanged = True
    End Sub

    Private Sub opt5Days_Click(sender As Object, e As EventArgs) Handles opt5Days.Click
        ChangeDates()
    End Sub

    Private Sub opt6Days_Click(sender As Object, e As EventArgs) Handles opt6Days.Click
        ChangeDates()
    End Sub

    Private Sub opt7Days_Click(sender As Object, e As EventArgs) Handles opt7Days.Click
        ChangeDates()
    End Sub

    Private Sub ChangeDates()
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

    Private Sub txtDuration_KeyUp(sender As Object, e As KeyEventArgs) Handles txtDuration.KeyUp
        ChangeDates()
    End Sub

    Private Sub txtDuration_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDuration.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub cmdCalSt_ValueChanged(sender As Object, e As EventArgs)
        dtStart.Text = dtStart.Value.ToString("dd-MMM-yyyy")
    End Sub

    Private Sub cmdCalSt_Click(sender As Object, e As EventArgs)
        If IsDate(dtStart.Text) Then
            dtStart.Value = CDate(dtStart.Text)
        Else
            dtStart.Value = DateTime.Now
        End If
    End Sub

    Private Sub CmdCalEnd_ValueChanged(sender As Object, e As EventArgs)
        dtEnd.Text = dtEnd.Value.ToString("dd-MMM-yyyy")
    End Sub

    Private Sub CmdCalEnd_Click(sender As Object, e As EventArgs)
        If IsDate(dtEnd.Text) Then
            dtEnd.Value = CDate(dtEnd.Text)
        Else
            dtEnd.Value = DateTime.Now
        End If
    End Sub

    Private Sub optWeeks_Click(sender As Object, e As EventArgs) Handles optWeeks.Click
        If dtStart.Text <> "" And dtEnd.Text <> "" Then
            If optWeeks.Checked = True Then
                txtDuration.Text = DateDiff("ww", CDate(dtStart.Text), CDate(dtEnd.Text), vbMonday, vbFirstFourDays)
            Else
                optWorkingDays_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub optWorkingDays_Click(sender As Object, e As EventArgs) Handles optWorkingDays.Click
        If opt5Days.Checked = True Then
            txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 5)
        ElseIf opt6Days.Checked = True Then
            txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 6)
        Else
            txtDuration.Text = _GlobalFunctions.CalculateDuration(CDate(dtStart.Text), CDate(dtEnd.Text), 7)
        End If
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Function Validation() As Boolean
        Try
            If ContainsInvalidChar(cboCDSID.Text) Then Throw New Exception("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ' "" ;")
            If CDate(dtStart.Text) > CDate(dtEnd.Text) Then Throw New Exception("End date cannot be earlier than start date")

            If cboProcessStepLocation.SelectedIndex = 0 Then Throw New Exception("Please select process step location.")
            If cboMatchedFacility.SelectedIndex = 0 Then Throw New Exception("Please select matched facility.")
            If cboSubFacility.SelectedIndex = 0 Then Throw New Exception("Please select sub facility.")


            Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
            If CDate(dtStart.Text) < CDate(Form.DataCenter.GlobalValues.WS.Range(_GlobalFunctions.ColumnLetter(Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn + 1).ToString & 4).Value) Then
                Throw New Exception("Start date should be in the timeline section dates.")
            End If
            If CDate(dtEnd.Text) > CDate(Form.DataCenter.GlobalValues.WS.Range(_GlobalFunctions.ColumnLetter(Form.DataCenter.GlobalSections.TimeLineSectionLastColumn - 1).ToString & 4).Value) Then
                Throw New Exception("End date should be in the timeline section dates.")
            End If

            Validation = True
        Catch ex As Exception
            Validation = False
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEdit, ex.Message), "Edit ProcessStep Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Function ContainsInvalidChar(strValue As String) As Boolean
        ContainsInvalidChar = False
        If Strings.InStr(1, strValue, "'") > 0 Or Strings.InStr(1, strValue, """") > 0 Or Strings.InStr(1, strValue, ";") > 0 Then
            ContainsInvalidChar = True
        End If
    End Function

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

    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click

        Dim wd As Integer
        Try
            Me.Cursor = Cursors.WaitCursor
            If Validation() = False Then
                DialogResult = DialogResult.None
                Exit Sub
            End If


            If opt5Days.Checked = True Then
                wd = 5
            ElseIf opt6Days.Checked = True Then
                wd = 6
            Else
                wd = 7
            End If

            Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep
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

            If chkStart.Checked = True And chkEnd.Checked = True And txtDuration.Visible = False Then
                '--------------------------------------------------------------------------------------------
                '   Start checked & end checked & Duration unvisible
                '   input : Startvalue , endvalue , nothing ?

                '   Start checked & End unvisible & duration visible
                '   input : Startvalue , nothing ? , duration value

                '   Start unchecked & End checked & Duration unvisible
                '   Input :  nothing , end value  , Nothing ?
                '--------------------------------------------------------------------------------------------
                If (_ProcessStep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, _pe26, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _UsercaseSeq, _ProcessStepSeq,
                                      CDate(dtStart.Text),
                                      CDate(dtEnd.Text),
                                      Nothing, wd, RemoveSPChars(cboCDSID.Text), If(txtRemarks.Text = String.Empty, Nothing, txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf chkStart.Checked = False And chkEnd.Checked = True And txtDuration.Visible = False Then
                If (_ProcessStep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, _pe26, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _UsercaseSeq, _ProcessStepSeq,
                                      Nothing,
                                      CDate(dtEnd.Text),
                                      Nothing, wd, RemoveSPChars(cboCDSID.Text), If(txtRemarks.Text = String.Empty, Nothing, txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf chkStart.Checked = True And chkEnd.Visible = False And txtDuration.Enabled = True Then
                If (_ProcessStep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, _pe26, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _UsercaseSeq, _ProcessStepSeq,
                                      CDate(dtStart.Text),
                                      Nothing,
                                      Integer.Parse(txtDuration.Text), wd, RemoveSPChars(cboCDSID.Text), If(txtRemarks.Text = String.Empty, Nothing, txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf chkStart.Checked = False And chkEnd.Checked = False Then
                If (_ProcessStep.Edit(Form.DataCenter.VehicleConfig.VehiclePe02, _pe26, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, _UsercaseSeq, _ProcessStepSeq,
                                      Nothing,
                                      Nothing,
                                      Nothing, wd, RemoveSPChars(cboCDSID.Text), If(txtRemarks.Text = String.Empty, Nothing, txtRemarks.Text), strMatchedFacility, strProcessStepLocation, strLocation_CBG, strSubFacility, chkHolidays.Checked, Form.DataCenter.ProgramConfig.BuildType) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If CT.Data.DataCenter.GlobalValues.message <> String.Empty Then MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "Edit Process Step", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            'Dim Cls As New Form.DataCenter.GlobalFunctions
            _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,, Form.DataCenter.ProcessStepConfig.ProcessStepPe26)
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            MessageBox.Show("Data sucessfully updated...", "Edit ProcessStep", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEdit, ex.Message), "Edit ProcessStep", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub frmEdit_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F4 Then
            cboLocation_CBG.Focus()
        ElseIf e.KeyCode = Keys.F7 Then
            cmdOk_Click(sender, e)
        End If
    End Sub

    Private Sub dtStart_ValueChanged(sender As Object, e As EventArgs) Handles dtStart.ValueChanged
        If bolChangeFlag Then Exit Sub

        If (dtStart.Text <> "" And dtEnd.Text <> "") Then
            lblKWSt.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtStart.Text, Form.DataCenter.GlobalValues.myCWR, vbMonday)

            If dtEnd.Visible = True Then
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
        End If
    End Sub

    Private Sub dtEnd_ValueChanged(sender As Object, e As EventArgs) Handles dtEnd.ValueChanged
        If bolChangeFlag Then Exit Sub

        If dtEnd.Value < dtStart.Value Then
            MessageBox.Show("End date cannot be greater than start date.", "Process step edit", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

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


End Class