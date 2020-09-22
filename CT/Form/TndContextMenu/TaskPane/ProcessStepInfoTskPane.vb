Imports System.Windows.Forms
Public Class ProcessStepInfoTskPane
    Public Sub loadData()

        Try
            If Form.DataCenter.ProcessStepConfig.ProcessStepPe26 = 0 Then
                Me.Visible = False
                Exit Sub
            End If
            If Me.Visible = False Then Exit Sub
            Me.Cursor = Cursors.WaitCursor

            Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep()
            _ProcessStep.SelectProcessStepDedicated(Form.DataCenter.ProcessStepConfig.ProcessStepPe26, Form.DataCenter.ProgramConfig.IsGeneric)

            If Form.DataCenter.ProcessStepConfig.PSIsGapOrDelay = False Then
                txtLocationCBG.Text = _ProcessStep.FacilityCbg
                txtProcessstepLocation.Text = _ProcessStep.FacilityLocation
                txtMatchedFacility.Text = _ProcessStep.FacilityName
                txtSubFacility.Text = _ProcessStep.SubFacilityName
                txtCDSID.Text = _ProcessStep.Cdsid
            Else
                txtLocationCBG.Text = "-"
                txtProcessstepLocation.Text = "-"
                txtMatchedFacility.Text = "-"
                txtSubFacility.Text = "-"
                txtCDSID.Text = "-"
            End If

            txtGlobal.Text = _ProcessStep.GlobalDVP
            txtUser.Text = _ProcessStep.TeamName
            txtUserCase.Text = _ProcessStep.Usercase
            txtProcessStep.Text = _ProcessStep.ProcessStepName
            txtRemarks.Text = _ProcessStep.Remarks

            txtStartDate.Text = _ProcessStep.PlannedStart & " ( " & DatePart(DateInterval.WeekOfYear, _ProcessStep.PlannedStart, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) & " )"
            txtEdDate.Text = _ProcessStep.PlannedEnd & " ( " & DatePart(DateInterval.WeekOfYear, _ProcessStep.PlannedEnd, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays) & " )"
            txtDurationValue.Text = _ProcessStep.Duration.ToString() & " days"
            txtWorkDays.Text = _ProcessStep.WorkingDays & " days/week"
            txtVacatonDays.Text = IIf(_ProcessStep.IsWithHoliday, "Yes", "No")

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmEdit, ex.Message), "Show ProcessStep info", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub


End Class



