Imports System.ComponentModel
Imports System.Globalization.DateTimeFormatInfo
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmPick1stVP
    'Dim clsStored As CT.Data.VehiclePlan.Plan = New CT.Data.VehiclePlan.Plan


    Dim _Program As CT.Data.ProgramConfiguration = New CT.Data.ProgramConfiguration
    Dim myDatatable As System.Data.DataTable

    Sub Set_DateFormat(ctl As DateTimePicker, format As Boolean)
        If format = True Then ctl.CustomFormat = "dd-MMM-yyyy" : Exit Sub
        ctl.CustomFormat = " "
    End Sub

    Private Sub frmPick1stVP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dtTable As System.Data.DataTable
        Dim drRow As System.Data.DataRow
        Try
            Set_DateFormat(dtMRD, False)
            Set_DateFormat(dtM1, False)
            Set_DateFormat(dtM1DC, False)
            Set_DateFormat(dtVP, False)
            Set_DateFormat(dtPEC, False)
            Set_DateFormat(dtFEC, False)

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            dtTable = _PlanInterface.SelectDateInformation(Form.DataCenter.ProgramConfig.pe02)

            If dtTable IsNot Nothing And dtTable.Rows.Count > 0 Then
                drRow = dtTable.Rows(0)

                If Not IsDBNull(drRow("AssyMRD")) Then
                    Set_DateFormat(dtMRD, True)
                    dtMRD.Text = drRow("AssyMRD").ToString()
                End If

                If Not IsDBNull(drRow("Firstm1")) Then
                    Set_DateFormat(dtM1, True)
                    dtM1.Text = drRow("Firstm1").ToString()
                End If

                If Not IsDBNull(drRow("M1DC")) Then
                    Set_DateFormat(dtM1DC, True)
                    dtM1DC.Text = drRow("M1DC").ToString()
                End If


                If Not IsDBNull(drRow("FirstVP")) Then
                    Set_DateFormat(dtVP, True)
                    dtVP.Text = drRow("FirstVP").ToString()
                End If


                If Not IsDBNull(drRow("PEC")) Then
                    Set_DateFormat(dtPEC, True)
                    dtPEC.Text = drRow("PEC").ToString()
                End If


                If Not IsDBNull(drRow("FEC")) Then
                    Set_DateFormat(dtFEC, True)
                    dtFEC.Text = drRow("FEC").ToString()
                End If
            End If

            Calculate_CW_For_Label(dtMRD.Text, lblMRD)

            If IsPlanM1() Then
                With dtVP
                    .Enabled = False
                End With

                With dtPEC
                    .Enabled = False
                End With

                With dtFEC
                    .Enabled = False
                End With

                With lblVP
                    .Enabled = False
                    .Text = "()"
                End With

                With lblPEC
                    .Enabled = False
                    .Text = "()"
                End With

                With lblFEC
                    .Enabled = False
                    .Text = "()"
                End With

                If dtM1.Text <> " " Then Calculate_CW_For_Label(dtM1.Value.ToString, lblM1)
                If dtM1DC.Text <> " " Then Calculate_CW_For_Label(dtM1DC.Value.ToString, lblM1DC)

            ElseIf IsPlanVP() Then
                With dtM1
                    .Enabled = False
                    '.Text = " "
                End With

                With dtM1DC
                    .Enabled = False
                    '.Text = " "
                End With

                With lblM1
                    .Enabled = False
                    .Text = "()"
                End With

                With lblM1DC
                    .Enabled = False
                    .Text = "()"
                End With

                If dtVP.Text <> " " Then Calculate_CW_For_Label(dtVP.Value.ToString, lblVP)
                If dtPEC.Text <> " " Then Calculate_CW_For_Label(dtPEC.Value.ToString, lblPEC)
                If dtFEC.Text <> " " Then Calculate_CW_For_Label(dtFEC.Value.ToString, lblFEC)
            End If
            If Form.DataCenter.ProgramConfig.BuildPhase = String.Empty Then Throw New Exception("Plan BuildPhase is empty!")
            If Form.DataCenter.ProgramConfig.pe02 = 0 Then Throw New Exception("Plan ID is empty!")

            'Source for pe02?
            myDatatable = _Program.SelectProgramConfigs(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
            txtTnDPlanner.Text = myDatatable.Rows(0)(CT.Data.ProgramConfiguration.SelectProgramConfigsColumns.TnDPlanner.ToString).ToString

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "Load Gateways", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Function IsPlanM1() As Boolean
        Select Case (Form.DataCenter.ProgramConfig.BuildPhase)
            Case CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString, CT.Data.DataCenter.BuildPhase.X0.ToString, CT.Data.DataCenter.BuildPhase.X1.ToString, CT.Data.DataCenter.BuildPhase.XM.ToString
                IsPlanM1 = True
            Case Else
                IsPlanM1 = False
        End Select
    End Function

    Private Function IsPlanVP() As Boolean
        Select Case (Form.DataCenter.ProgramConfig.BuildPhase)
            Case CT.Data.DataCenter.BuildPhase.VP.ToString, CT.Data.DataCenter.BuildPhase.DCV.ToString, CT.Data.DataCenter.BuildPhase.PP.ToString, CT.Data.DataCenter.BuildPhase.TT.ToString
                IsPlanVP = True
            Case Else
                IsPlanVP = False
        End Select
    End Function


    Private Sub frmPick1stVP_Validating(sender As Object, e As CancelEventArgs) Handles Me.Validating
        Try
            If IsPlanVP() Then
                If Not (IsDate(dtMRD.Text) And IsDate(dtVP.Text) And IsDate(dtPEC.Text) And IsDate(dtFEC.Text)) Then _
                    Throw New Exception("One of the entered dates is invalid.")
            ElseIf IsPlanM1() Then
                If Not (IsDate(dtMRD.Text) And IsDate(dtM1.Text) And IsDate(dtM1DC.Text)) Then _
                    Throw New Exception("One of the entered dates is invalid.")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "Change gateway dates", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            e.Cancel = True
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub brnOk_Click(sender As Object, e As EventArgs) Handles brnOk.Click
        Try

            Dim _Modfunc As New Form.DataCenter.ModuleFunction
            '---------------------------------------------------------------
            ' Validation before converting
            '---------------------------------------------------------------
            If IsDate(dtMRD.Text) = False Then
                Throw New Exception("Sorry, the MRD date you entered are invalid or incomplete. Please enter valid dates and try again.")
            End If

            Select Case (Form.DataCenter.ProgramConfig.BuildPhase)
                Case CT.Data.DataCenter.BuildPhase.VP.ToString, CT.Data.DataCenter.BuildPhase.DCV.ToString
                    ', CT.Data.DataCenter.BuildPhase.PP.ToString, CT.Data.DataCenter.BuildPhase.TT.ToString
                    If IsDate(dtVP.Text) = False Then
                        If MessageBox.Show("The VP date is blank." & vbNewLine & "Do you want to continue.", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                    If IsDate(dtPEC.Text) = False Then
                        If MessageBox.Show("The PEC date is blank." & vbNewLine & "Do you want to continue.", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                    If IsDate(dtFEC.Text) = False Then
                        If MessageBox.Show("The FEC date is blank." & vbNewLine & "Do you want to continue.", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                Case CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString, CT.Data.DataCenter.BuildPhase.X0.ToString, CT.Data.DataCenter.BuildPhase.X1.ToString, CT.Data.DataCenter.BuildPhase.XM.ToString
                    If IsDate(dtM1.Text) = False Then
                        If MessageBox.Show("The M1 date is blank." & vbNewLine & "Do you want to continue.", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            Exit Sub
                        End If
                    End If


                    If IsDate(dtM1DC.Text) = False Then
                        If MessageBox.Show("The M1DC date is blank." & vbNewLine & "Do you want to continue.", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            Exit Sub
                        End If
                    End If
            End Select

            If txtTnDPlanner.Text = "" Then
                MessageBox.Show("TndPlanner cannot be blank.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtTnDPlanner.Focus()
                Exit Sub
            End If
            'End If

            Me.Cursor = Cursors.AppStarting

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Try
            End If

            '--------------------------------------------------------------
            ' Convert generic plan 2 specific
            '--------------------------------------------------------------
            If _PlanInterface.ConvertGenericToSpecific(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.XccPe26, Form.DataCenter.ProgramConfig.XccPe01, Form.DataCenter.ProgramConfig.AssyBuildScale, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType, DirectCast([Enum].Parse(GetType(CT.Data.DataCenter.FileStatus), Form.DataCenter.ProgramConfig.FileStatus), CT.Data.DataCenter.FileStatus), chkInduFormatting.Checked) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

            '----------------------------------------------------------------
            ' keep withCustom formating checkbox in program config 
            '----------------------------------------------------------------
            Form.DataCenter.ProgramConfig.IsWithCustomFormatting = chkInduFormatting.Checked

            MessageBox.Show("Dates are saved and applied on interface." & vbNewLine & "Plan has been converted from generic to Specific.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            '----------------------------------------------------------------
            ' After Ok the messagebox we go to refresh plan from ribbon 
            '----------------------------------------------------------------
            Me.DialogResult = DialogResult.OK
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.DialogResult = DialogResult.Cancel
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub



    'Update labels to reflect chosen calendarweek
    Private Sub Calculate_CW_For_Label(txtDate As String, lblCW As System.Windows.Forms.Label)
        'Consolidated Sub to Write the corresponding calendarweek to the labels
        Dim lngCWISO As Long
        lngCWISO = CurrentInfo.Calendar.GetWeekOfYear(Convert.ToDateTime(txtDate),
                           CurrentInfo.CalendarWeekRule, CurrentInfo.FirstDayOfWeek)
        lblCW.Text = "(" & lngCWISO & ")"
    End Sub

    Private Sub dtMRD_TextChanged(sender As Object, e As EventArgs) Handles dtMRD.ValueChanged
        Set_DateFormat(dtMRD, True)
        Calculate_CW_For_Label(dtMRD.Text, lblMRD)
        Try
            If Form.DataCenter.GlobalSections.IsDateValid(dtMRD.Value) = False Then
                ErrorProvider.SetError(dtMRD, "Selected MRD is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtMRD, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "MRD Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtM1_TextChanged(sender As Object, e As EventArgs) Handles dtM1.ValueChanged
        Set_DateFormat(dtM1, True)
        Calculate_CW_For_Label(dtM1.Text, lblM1)
        Try
            If Form.DataCenter.GlobalSections.IsDateValid(dtM1.Value) = False Then
                ErrorProvider.SetError(dtM1, "Selected M1 is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtM1, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "M1 Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtM1DC_TextChanged(sender As Object, e As EventArgs) Handles dtM1DC.ValueChanged
        Set_DateFormat(dtM1DC, True)
        Calculate_CW_For_Label(dtM1DC.Text, lblM1DC)
        Try
            If Form.DataCenter.GlobalSections.IsDateValid(dtM1DC.Value) = False Then
                ErrorProvider.SetError(dtM1DC, "Selected M1DC is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtM1DC, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "M1DC Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtVP_TextChanged(sender As Object, e As EventArgs) Handles dtVP.ValueChanged
        Set_DateFormat(dtVP, True)
        Calculate_CW_For_Label(dtVP.Text, lblVP)
        Try
            If Form.DataCenter.GlobalSections.IsDateValid(dtVP.Value) = False Then
                ErrorProvider.SetError(dtVP, "Selected VP is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtVP, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "VP Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtPEC_TextChanged(sender As Object, e As EventArgs) Handles dtPEC.ValueChanged
        Set_DateFormat(dtPEC, True)
        Calculate_CW_For_Label(dtPEC.Text, lblPEC)
        Try
            If Form.DataCenter.GlobalSections.IsDateValid(dtPEC.Value) = False Then
                ErrorProvider.SetError(dtPEC, "Selected PEC is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtPEC, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "PEC Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtFEC_TextChanged(sender As Object, e As EventArgs) Handles dtFEC.ValueChanged
        Set_DateFormat(dtFEC, True)
        Calculate_CW_For_Label(dtFEC.Text, lblFEC)
        Try
            If Form.DataCenter.GlobalSections.IsDateValid(dtFEC.Value) = False Then
                ErrorProvider.SetError(dtFEC, "Selected FEC is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtFEC, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPick1stVP, ex.Message), "FEC Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClearMRD_Click(sender As Object, e As EventArgs) Handles btnClearMRD.Click
        Set_DateFormat(dtMRD, False)
    End Sub

    Private Sub btnClearM1_Click(sender As Object, e As EventArgs) Handles btnClearM1.Click
        Set_DateFormat(dtM1, False)
    End Sub

    Private Sub btnClearM1DC_Click(sender As Object, e As EventArgs) Handles btnClearM1DC.Click
        Set_DateFormat(dtM1DC, False)
    End Sub

    Private Sub btnClearVP_Click(sender As Object, e As EventArgs) Handles btnClearVP.Click
        Set_DateFormat(dtVP, False)
    End Sub

    Private Sub btnClearPEC_Click(sender As Object, e As EventArgs) Handles btnClearPEC.Click
        Set_DateFormat(dtPEC, False)
    End Sub

    Private Sub btnClearFEC_Click(sender As Object, e As EventArgs) Handles btnClearFEC.Click
        Set_DateFormat(dtFEC, False)
    End Sub

    Private Sub frmPick1stVP_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            brnOk_Click(sender, e)
        End If
    End Sub

End Class