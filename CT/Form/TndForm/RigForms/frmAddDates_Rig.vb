Imports System.Windows.Forms
Imports System.Drawing

Public Class frmAddDates_Rig
    Dim myFirstDOW As DayOfWeek = vbSunday

    'Fontcolor button click event
    Private Sub cmdFontcolor_Click(sender As Object, e As EventArgs) Handles cmdFontcolor.Click
        If ColorDialog1.ShowDialog <> System.Windows.Forms.DialogResult.Cancel Then
            cmdFontcolor.BackColor = ColorDialog1.Color
            cmdFontcolor.Tag = IIf(Val(cmdFontcolor.BackColor.R) < 100 And Val(cmdFontcolor.BackColor.R) >= 10, "0" & cmdFontcolor.BackColor.R, IIf(Val(cmdFontcolor.BackColor.R) < 10, "00" & cmdFontcolor.BackColor.R, cmdFontcolor.BackColor.R)) &
                    IIf(Val(cmdFontcolor.BackColor.G) < 100 And Val(cmdFontcolor.BackColor.G) >= 10, "0" & cmdFontcolor.BackColor.G, IIf(Val(cmdFontcolor.BackColor.G) < 10, "00" & cmdFontcolor.BackColor.G, cmdFontcolor.BackColor.G)) &
                    IIf(Val(cmdFontcolor.BackColor.B) < 100 And Val(cmdFontcolor.BackColor.B) >= 10, "0" & cmdFontcolor.BackColor.B, IIf(Val(cmdFontcolor.BackColor.B) < 10, "00" & cmdFontcolor.BackColor.B, cmdFontcolor.BackColor.B))
        End If
    End Sub

    'Backcolor button click event
    Private Sub cmdBackcolor_Click(sender As Object, e As EventArgs) Handles cmdBackcolor.Click
        If ColorDialog1.ShowDialog <> System.Windows.Forms.DialogResult.Cancel Then
            cmdBackcolor.BackColor = ColorDialog1.Color
            cmdBackcolor.Tag = IIf(Val(cmdBackcolor.BackColor.R) < 100 And Val(cmdBackcolor.BackColor.R) >= 10, "0" & cmdBackcolor.BackColor.R, IIf(Val(cmdBackcolor.BackColor.R) < 10, "00" & cmdBackcolor.BackColor.R, cmdBackcolor.BackColor.R)) &
                    IIf(Val(cmdBackcolor.BackColor.G) < 100 And Val(cmdBackcolor.BackColor.G) >= 10, "0" & cmdBackcolor.BackColor.G, IIf(Val(cmdBackcolor.BackColor.G) < 10, "00" & cmdBackcolor.BackColor.G, cmdBackcolor.BackColor.G)) &
                    IIf(Val(cmdBackcolor.BackColor.B) < 100 And Val(cmdBackcolor.BackColor.B) >= 10, "0" & cmdBackcolor.BackColor.B, IIf(Val(cmdBackcolor.BackColor.B) < 10, "00" & cmdBackcolor.BackColor.B, cmdBackcolor.BackColor.B))
        End If
    End Sub

    'Form controls 'Reset'
    Private Sub ResetForm()
        Set_DateFormat(dtMRD, False)
        Set_DateFormat(dtM1, False)
        Set_DateFormat(dtM1dc, False)
        Set_DateFormat(dtVP, False)
        Set_DateFormat(dtPec, False)
        Set_DateFormat(dtFec, False)

        lblFec.Text = "( )"
        lblM1.Text = "( )"
        lblM1dc.Text = "( )"
        lblMRD.Text = "( )"
        lblPec.Text = "( )"
        lblVp.Text = "( )"

        cmdBackcolor.BackColor = Color.FromArgb(220, 220, 220)
        cmdFontcolor.BackColor = Color.FromArgb(220, 220, 220)
        cmdBackcolor.Tag = ""
        cmdFontcolor.Tag = ""
        lblID.Text = ""
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        cmdAdd.Enabled = True
        txtHCid.Text = ""

        '-------------------------------------------------------------------
        ' Clean notification text
        ErrorProvider.SetError(cmdFontcolor, "")
        ErrorProvider.SetError(cmdBackcolor, "")

        ErrorProvider.SetError(txtHCid, "")
        ErrorProvider.SetError(dtMRD, "")

        ErrorProvider.SetError(dtM1, "")
        ErrorProvider.SetError(dtM1dc, "")

        ErrorProvider.SetError(dtVP, "")
        ErrorProvider.SetError(dtPec, "")
        ErrorProvider.SetError(dtFec, "")
    End Sub

    'Subroutine : sbFillData
    'Purpose    : To retreive & fill MRD date data in datagrid view
    'Parameters : @pe02_TnDProgramDetails_FK
    'Notes      : @pe02_TnDProgramDetails_FK to be passed dynamically for sp [Report_ListAddtionalDateInformation]***  
    Sub FillData()
        Try
            Dim _Plan As New Data.RigPlan.Plan
            Dim _AddtionalDateInformation As New Data.AddtionalDateInformation

            Dim myDataTable As System.Data.DataTable
            Dim myView As System.Data.DataView

            Select Case (Form.DataCenter.ProgramConfig.BuildPhase)
                Case CT.Data.DataCenter.BuildPhase.VP.ToString, CT.Data.DataCenter.BuildPhase.DCV.ToString, CT.Data.DataCenter.BuildPhase.PP.ToString, CT.Data.DataCenter.BuildPhase.TT.ToString
                    grpM1.Enabled = False
                Case CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString, CT.Data.DataCenter.BuildPhase.X0.ToString, CT.Data.DataCenter.BuildPhase.X1.ToString, CT.Data.DataCenter.BuildPhase.XM.ToString
                    grpVP.Enabled = False
            End Select

            '-------------------------------------------------------------------------------------------------
            ' 1- Fetching a list of HCIDs from DB
            myDataTable = _AddtionalDateInformation.SelectAllDateInformation(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.BuildType)
            If myDataTable Is Nothing Then
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If

            '-------------------------------------------------------------------------------------------------
            ' 2- Setting view columns
            myView = New System.Data.DataView(myDataTable)
            myDataTable = myView.ToTable(False, "HealthChartId", "AssyMRD", "Job#1", "Firstm1", "M1DC", "FirstVP", "PEC", "FEC", "DateBackRGB", "DateFontRGB", "pe67_AddtionalDateInformation_PK")

            dgvDates.DataSource = myDataTable
            dgvDates.Columns(2).Visible = False
            dgvDates.Columns(8).Visible = False
            dgvDates.Columns(9).Visible = False
            dgvDates.Columns(10).Visible = False

            '-------------------------------------------------------------------------------------------------
            ' 3- Checking of there is a HCID in DB ot not
            '-------------------------------------------------------------------------------------------------
            If myDataTable.Rows.Count < 1 Then
                myDataTable = Nothing
                myDataTable = _Plan.SelectDateInformation(Form.DataCenter.ProgramConfig.pe02)
                If myDataTable.Rows.Count > 0 Then
                    With myDataTable.Rows(0)
                        txtHCid.Text = .Item("HealthChartId").ToString
                        AssignDatevalue(.Item("AssyMrd").ToString, dtMRD)
                        If dtMRD.Text <> " " Then lblMRD.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtMRD.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                        If Form.DataCenter.ProgramConfig.BuildPhase = "M1" Then
                            AssignDatevalue(.Item("Firstm1").ToString, dtM1)
                            If dtM1.Text <> " " Then lblM1.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtM1.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                            AssignDatevalue(.Item("M1DC").ToString, dtM1dc)
                            If dtM1dc.Text <> " " Then lblM1dc.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtM1dc.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                        Else
                            AssignDatevalue(.Item("FirstVP").ToString, dtVP)
                            If dtVP.Text <> " " Then lblVp.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtVP.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                            AssignDatevalue(.Item("PEC").ToString, dtPec)
                            If dtPec.Text <> " " Then lblPec.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtPec.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                            AssignDatevalue(.Item("FEC").ToString, dtFec)
                            If dtFec.Text <> " " Then lblFec.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtFec.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                        End If
                        If .Item("DateBackRGB").ToString <> "" Then
                            cmdBackcolor.Tag = .Item("DateBackRGB").ToString
                            cmdBackcolor.BackColor = Color.FromArgb(Val(Strings.Mid(.Item("DateBackRGB").ToString, 1, 3)), Val(Strings.Mid(.Item("DateBackRGB").ToString, 4, 3)), Val(Strings.Mid(.Item("DateBackRGB").ToString, 7, 3)))
                        End If
                        If .Item("DateFontRGB").ToString <> "" Then
                            cmdFontcolor.Tag = .Item("DateFontRGB").ToString
                            cmdFontcolor.BackColor = Color.FromArgb(Val(Strings.Mid(.Item("DateFontRGB").ToString, 1, 3)), Val(Strings.Mid(.Item("DateFontRGB").ToString, 4, 3)), Val(Strings.Mid(.Item("DateFontRGB").ToString, 7, 3)))
                        End If
                    End With
                End If
                cmdAdd.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub AssignDatevalue(source As String, ctl As DateTimePicker)
        If Not source Is Nothing Then
            If source <> "" Then
                Set_DateFormat(ctl, True)
                ctl.Value = source.ToString()
            End If
        End If
    End Sub

    'Grid cell click event
    'To load dates from grid to datetimepickers
    Private Sub dgvDates_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDates.CellClick
        If e.RowIndex >= 0 Then

            ResetForm()

            txtHCid.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(0)
            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(1)) = False Then
                If Year(dgvDates.Rows(e.RowIndex).DataBoundItem(1)) <> 1 Then
                    Set_DateFormat(dtMRD, True)
                    dtMRD.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(1)
                    lblMRD.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtMRD.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                End If
            End If

            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(3)) = False Then
                If Year(dgvDates.Rows(e.RowIndex).DataBoundItem(3)) <> 1 Then
                    Set_DateFormat(dtM1, True)
                    dtM1.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(3)
                    lblM1.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtM1.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                End If
            End If

            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(4)) = False Then
                If Year(dgvDates.Rows(e.RowIndex).DataBoundItem(4)) <> 1 Then
                    Set_DateFormat(dtM1dc, True)
                    dtM1dc.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(4)
                    lblM1dc.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtM1dc.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                End If
            End If

            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(5)) = False Then
                If Year(dgvDates.Rows(e.RowIndex).DataBoundItem(5)) <> 1 Then
                    Set_DateFormat(dtVP, True)
                    dtVP.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(5)
                    lblVp.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtVP.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                End If
            End If

            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(6)) = False Then
                If Year(dgvDates.Rows(e.RowIndex).DataBoundItem(6)) <> 1 Then
                    Set_DateFormat(dtPec, True)
                    dtPec.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(6)
                    lblPec.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtPec.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                End If
            End If

            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(7)) = False Then
                If Year(dgvDates.Rows(e.RowIndex).DataBoundItem(7)) <> 1 Then
                    Set_DateFormat(dtFec, True)
                    dtFec.Text = dgvDates.Rows(e.RowIndex).DataBoundItem(7)
                    lblFec.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtFec.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
                End If
            End If

            cmdBackcolor.Tag = IIf(IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(8)), "", dgvDates.Rows(e.RowIndex).DataBoundItem(8))
            cmdFontcolor.Tag = IIf(IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(9)), "", dgvDates.Rows(e.RowIndex).DataBoundItem(9))
            lblID.Text = IIf(IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(10)), "", dgvDates.Rows(e.RowIndex).DataBoundItem(10))

            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(8)) = False Then
                cmdBackcolor.BackColor = Color.FromArgb(Val(Strings.Mid(dgvDates.Rows(e.RowIndex).DataBoundItem(8), 1, 3)), Val(Strings.Mid(dgvDates.Rows(e.RowIndex).DataBoundItem(8), 4, 3)), Val(Strings.Mid(dgvDates.Rows(e.RowIndex).DataBoundItem(8), 7, 3)))
            End If
            If IsDBNull(dgvDates.Rows(e.RowIndex).DataBoundItem(9)) = False Then
                cmdFontcolor.BackColor = Color.FromArgb(Val(Strings.Mid(dgvDates.Rows(e.RowIndex).DataBoundItem(9), 1, 3)), Val(Strings.Mid(dgvDates.Rows(e.RowIndex).DataBoundItem(9), 4, 3)), Val(Strings.Mid(dgvDates.Rows(e.RowIndex).DataBoundItem(9), 7, 3)))
            End If

            cmdUpdate.Enabled = True
            cmdDelete.Enabled = True
            cmdAdd.Enabled = False
        End If
    End Sub

    'To set/reset datetimepicker blank
    Sub Set_DateFormat(ctl As DateTimePicker, format As Boolean)
        If format = True Then ctl.CustomFormat = "dd-MMM-yyyy" : Exit Sub
        ctl.CustomFormat = " "
        ErrorProvider.Clear()
    End Sub

    'Delete Button click event
    'To Delete the selected date from grid
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        Try
            Application.UseWaitCursor = True
            If Val(lblID.Text) = 0 Then
                Throw New Exception("Sorry, invalid selection. Please select a HCID from the list and try again.")
            End If

            'Dim _PlanInterface As Data.Interfaces.PlanInterface

            'If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            '    _PlanInterface = New Data.VehiclePlan.Plan
            'ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            '    _PlanInterface = New Data.RigPlan.Plan
            'Else
            '    Exit Try
            'End If
            Dim _Plan As New Data.AddtionalDateInformation



            If _Plan.DeleteHealthChartID(lblID.Text, Form.DataCenter.ProgramConfig.pe02, txtHCid.Text) = False Then
                Throw New Exception("Sorry, your data could not be deleted! database error. " & CT.Data.DataCenter.GlobalValues.message)
            End If
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            MessageBox.Show("Data deleted sucessfully!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            ResetForm()
            FillData()

            '--------------------------------------------------------------
            ' Apply chnages after update
            '--------------------------------------------------------------
            Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader
            _DrawTndPlanHeader.ApplyHolidaysFlags()
            _DrawTndPlanHeader.ApplyGatewayFlags()

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.UseWaitCursor = False
        End Try
    End Sub

    'Update button click event
    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        Try
            Application.UseWaitCursor = True
            '------------------------------------------------
            ' By updating a HCID the user is not allowed to add new
            ' But after updating the add button is active
            '------------------------------------------------
            cmdAdd.Enabled = False


            If FormValiding() = False Then Throw New Exception("000")
            If lblID.Text = "" Then Throw New Exception("000")

            Dim _AddtionalDateInformation As New Data.AddtionalDateInformation


            If _AddtionalDateInformation.UpdateHealthChartID(lblID.Text, txtHCid.Text, IIf(dtMRD.Text = " ", Nothing, dtMRD.Text), IIf(dtM1.Text = " ", Nothing, dtM1.Text), IIf(dtM1dc.Text = " ", Nothing, dtM1dc.Text), IIf(dtVP.Text = " ", Nothing, dtVP.Text), IIf(dtPec.Text = " ", Nothing, dtPec.Text),
                                                      IIf(dtFec.Text = " ", Nothing, dtFec.Text), Nothing, cmdBackcolor.Tag, cmdFontcolor.Tag) = False Then
                Throw New Exception("Sorry, you changes could not be saved to database! database error. " & CT.Data.DataCenter.GlobalValues.message)
            End If
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            MessageBox.Show("Data saved sucessfully!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            ResetForm()
            FillData()

            '--------------------------------------------------------------
            ' Apply chnages after update
            '--------------------------------------------------------------
            Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader
            _DrawTndPlanHeader.ApplyHolidaysFlags()
            _DrawTndPlanHeader.ApplyGatewayFlags()

            Dim _tndTitle As New Form.DisplayUtilities.TndPlanTitle
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            _tndTitle.LoadAndFormatLabel()
            Dim _obj As New Form.DataCenter.ModuleFunction
            _obj.sbProtectPlan()

            '------------------------------------------------
            ' By updating a HCID the user is not allowed to add new
            ' But after updating the add button is active
            '------------------------------------------------
            cmdAdd.Enabled = True
        Catch ex As Exception
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.UseWaitCursor = False
            txtHCid.Focus()
        End Try
    End Sub

    'Add button click event
    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Try
            Application.UseWaitCursor = True
            If FormValiding() = False Then Throw New Exception("000")

            'Dim _PlanInterface As Data.Interfaces.PlanInterface

            'If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
            '    _PlanInterface = New Data.VehiclePlan.Plan
            'ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
            '    _PlanInterface = New Data.RigPlan.Plan
            'Else
            '    Exit Try
            'End If
            Dim _Plan As New Data.AddtionalDateInformation

            '---------------------------------------------------------------------------------
            ' Validating for duplicate HCID
            '---------------------------------------------------------------------------------
            Using dt As System.Data.DataTable = _Plan.SelectAllDateInformation(Form.DataCenter.ProgramConfig.pe02, Integer.Parse(txtHCid.Text))
                If dt Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If dt.Rows.Count >= 1 Then Throw New Exception("HCID already existed!")
            End Using


            If _Plan.AddHealthChartID(Form.DataCenter.ProgramConfig.pe02, txtHCid.Text, IIf(dtMRD.Text = " ", Nothing, dtMRD.Text), IIf(dtM1.Text = " ", Nothing, dtM1.Text), IIf(dtM1dc.Text = " ", Nothing, dtM1dc.Text), IIf(dtVP.Text = " ", Nothing, dtVP.Text), IIf(dtPec.Text = " ", Nothing, dtPec.Text),
                                                      IIf(dtFec.Text = " ", Nothing, dtFec.Text), Nothing, cmdBackcolor.Tag, cmdFontcolor.Tag, Form.DataCenter.ProgramConfig.BuildType) = False Then
                If InStr(CT.Data.DataCenter.GlobalValues.message.ToString, "duplicate", CompareMethod.Text) > 0 Then
                    Throw New Exception("The given date entry already exists.")
                Else
                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                End If
            End If
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            MessageBox.Show("Data saved sucessfully!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            ResetForm()
            FillData()
            '--------------------------------------------------------------
            ' Apply chnages after update
            '--------------------------------------------------------------
            Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader
            _DrawTndPlanHeader.ApplyHolidaysFlags()
            _DrawTndPlanHeader.ApplyGatewayFlags()

            Dim _tndTitle As New Form.DisplayUtilities.TndPlanTitle
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            _tndTitle.LoadAndFormatLabel()
            Dim _obj As New Form.DataCenter.ModuleFunction
            _obj.sbProtectPlan()
            txtHCid.Focus()
        Catch ex As Exception
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.UseWaitCursor = False
        End Try
    End Sub

    Private Sub frmAddDates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cmdAdd.Enabled = True
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        FillData()
        txtHCid.Focus()

        If Form.DataCenter.ProgramConfig.IsGeneric = True Then
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = False
        End If
    End Sub

    Private Sub dtM1_ValueChanged(sender As Object, e As EventArgs) Handles dtM1.ValueChanged
        dtM1.CustomFormat = "dd-MMM-yyyy"
        lblM1.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtM1.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
        Try
            Dim FindColumn As Excel.Range = Nothing
            FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtM1.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
            If FindColumn Is Nothing Then
                ErrorProvider.SetError(dtM1, "Selected M1 date is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtM1, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "M1 Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtM1dc_ValueChanged(sender As Object, e As EventArgs) Handles dtM1dc.ValueChanged
        dtM1dc.CustomFormat = "dd-MMM-yyyy"
        lblM1dc.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtM1dc.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)

        Try
            Dim FindColumn As Excel.Range = Nothing

            FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtM1dc.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
            If FindColumn Is Nothing Then
                ErrorProvider.SetError(dtM1dc, "Selected M1DC is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtM1dc, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "M1DC Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtVP_ValueChanged(sender As Object, e As EventArgs) Handles dtVP.ValueChanged
        dtVP.CustomFormat = "dd-MMM-yyyy"
        lblVp.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtVP.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)
        Try
            Dim FindColumn As Excel.Range = Nothing

            FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtVP.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
            If FindColumn Is Nothing Then
                ErrorProvider.SetError(dtVP, "Selected VP is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtVP, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "VP Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtPec_ValueChanged(sender As Object, e As EventArgs) Handles dtPec.ValueChanged
        Set_DateFormat(dtPec, True)
        lblPec.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtPec.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)

        Try
            Dim FindColumn As Excel.Range = Nothing

            FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtPec.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
            If FindColumn Is Nothing Then
                ErrorProvider.SetError(dtPec, "Selected PEC is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtPec, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "PEC Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtFec_ValueChanged(sender As Object, e As EventArgs) Handles dtFec.ValueChanged
        dtFec.CustomFormat = "dd-MMM-yyyy"
        lblFec.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtFec.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)

        Try
            Dim FindColumn As Excel.Range = Nothing

            FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtFec.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
            If FindColumn Is Nothing Then
                ErrorProvider.SetError(dtFec, "Selected FEC is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtFec, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "FEC Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Dim _obj As New Form.DataCenter.ModuleFunction
        Try
            Me.Cursor = Cursors.AppStarting
            '----------------------------------------------------------------------
            ' update header label of the plan after changes on this form
            '----------------------------------------------------------------------
            Dim myDatatable As System.Data.DataTable
            Dim _Program As CT.Data.ProgramConfiguration = New CT.Data.ProgramConfiguration
            myDatatable = _Program.SelectProgramConfigs(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
            If myDatatable.Rows.Count = 0 Then

                '--------------------------------------
                ' Only if title row is 0 we can insert 
                ' new row
                '--------------------------------------
                If _Program.Add(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.HCIDName, "Confidential", "1" & "." & "0", "1" & "." & "0", Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType, Nothing) = False Then
                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)
                End If

            ElseIf myDatatable.Rows.Count > 0 Then


                Dim nudIssueMin As String = String.Empty
                If myDatatable.Rows(0).Item(5).ToString.Split(".").Length > 1 Then
                    nudIssueMin = myDatatable.Rows(0).Item(4).ToString.Split(".")(1)
                End If

                If _Program.Update(Long.Parse(myDatatable.Rows(0).Item(9)),
                                   Form.DataCenter.ProgramConfig.HCID,
                                   myDatatable.Rows(0).Item(2),
                                   "Confidential",
                                   myDatatable.Rows(0).Item(5).ToString.Split(".")(0) & "." & nudIssueMin,
                                   myDatatable.Rows(0).Item(5).ToString.Split(".")(0) & "." & nudIssueMin,
                                    myDatatable.Rows(0).Item(6),
                                   myDatatable.Rows(0).Item(7), Nothing) = False Then

                    Throw New Exception("Sorry your changes could not be saved! Database error. Error:-" + CT.Data.DataCenter.GlobalValues.message)

                End If


            End If

            '---------------------------------------
            ' update excel interface
            '---------------------------------------
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Dim _TndPlanTitle As Form.DisplayUtilities.TndPlanTitle = New Form.DisplayUtilities.TndPlanTitle
            _TndPlanTitle.LoadAndFormatLabel()
            _TndPlanTitle.FillMismatchedQty()
            _obj.sbProtectPlan()

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            Me.Close()

        End Try

    End Sub

    Private Sub cmdReset_Click(sender As Object, e As EventArgs) Handles cmdReset.Click
        Try
            Application.UseWaitCursor = True
            ResetForm()
            FillData()
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "Reset", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.UseWaitCursor = False
        End Try
    End Sub

    Private Sub frmAddDates_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            txtHCid.Focus()
        ElseIf e.KeyCode = Keys.F5 Then
            cmdReset_Click(sender, e)
        ElseIf e.KeyCode = Keys.F7 And cmdAdd.Enabled = True Then
            cmdAdd_Click(sender, e)
        ElseIf e.KeyCode = Keys.F8 And cmdUpdate.Enabled = True Then
            cmdUpdate_Click(sender, e)
        ElseIf e.KeyCode = Keys.F9 And cmdDelete.Enabled = True Then
            cmdDelete_Click(sender, e)
        End If
    End Sub

    Private Function FormValiding() As Boolean
        Dim FindColumn As Excel.Range = Nothing
        '------------------------------------------------------------------------------
        If Val(txtHCid.Text) = 0 Then
            ErrorProvider.SetError(txtHCid, "Please enter valid HCID and try again.")
            FormValiding = False
            Exit Function
        Else
            ErrorProvider.SetError(txtHCid, "")
            FormValiding = True
        End If

        '------------------------------------------------------------------------------
        If cmdFontcolor.Tag = "" Then
            ErrorProvider.SetError(cmdFontcolor, "Please select font color.")
            FormValiding = False
            Exit Function
        Else
            ErrorProvider.SetError(cmdFontcolor, "")
            FormValiding = True
        End If

        '------------------------------------------------------------------------------
        If cmdBackcolor.Tag = "" Then
            ErrorProvider.SetError(cmdBackcolor, "Please select back color.")
            FormValiding = False
            Exit Function
        Else
            ErrorProvider.SetError(cmdBackcolor, "")
            FormValiding = True
        End If

        '------------------------------------------------------------------------------
        FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtMRD.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
        If FindColumn Is Nothing Or dtMRD.Text = " " Then

            ErrorProvider.SetError(dtMRD, "Selected MRD is out of time-Line range.")
            FormValiding = False
            Exit Function
        Else
            ErrorProvider.SetError(dtMRD, "")
            FormValiding = True
        End If

        Select Case (Form.DataCenter.ProgramConfig.BuildPhase)
            Case CT.Data.DataCenter.BuildPhase.VP.ToString, CT.Data.DataCenter.BuildPhase.DCV.ToString, CT.Data.DataCenter.BuildPhase.PP.ToString, CT.Data.DataCenter.BuildPhase.TT.ToString

                '    If Form.DataCenter.ProgramConfig.BuildPhase = Data.DataCenter.BuildPhase.VP.ToString Or
                'Form.DataCenter.ProgramConfig.BuildPhase = Data.DataCenter.BuildPhase.DCV.ToString Then ' rdVP.Checked = True Then
                '-----------------------------------------------------------------------------------
                ' VP validation
                '-----------------------------------------------------------------------------------
                If dtPec.Text = " " Then
                    ErrorProvider.SetError(dtPec, "PEC field cannot be blank.")
                    FormValiding = False
                    Exit Function
                Else
                    FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtPec.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    If FindColumn Is Nothing Then
                        ErrorProvider.SetError(dtPec, "Selected PEC is out of time-Line range.")
                        FormValiding = False
                        Exit Function
                    Else
                        ErrorProvider.SetError(dtPec, "")
                        FormValiding = True
                    End If
                End If

                If dtFec.Text = " " Then
                    ErrorProvider.SetError(dtFec, "FEC field cannot be blank.")
                    FormValiding = False
                    Exit Function
                Else
                    FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtFec.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    If FindColumn Is Nothing Then
                        ErrorProvider.SetError(dtFec, "Selected FEC is out of time-Line range.")
                        FormValiding = False
                        Exit Function
                    Else
                        ErrorProvider.SetError(dtFec, "")
                        FormValiding = True
                    End If
                End If

                If dtVP.Text <> " " Then
                    FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtVP.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    If FindColumn Is Nothing Then
                        ErrorProvider.SetError(dtVP, "Selected VP is out of time-Line range.")
                        FormValiding = False
                        Exit Function
                    Else
                        ErrorProvider.SetError(dtVP, "")
                        FormValiding = True
                    End If
                End If

            Case CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString, CT.Data.DataCenter.BuildPhase.X0.ToString, CT.Data.DataCenter.BuildPhase.X1.ToString, CT.Data.DataCenter.BuildPhase.XM.ToString

                '    ElseIf Form.DataCenter.ProgramConfig.BuildPhase = Data.DataCenter.BuildPhase.M1.ToString Or
                'Form.DataCenter.ProgramConfig.BuildPhase = Data.DataCenter.BuildPhase.TPV.ToString Then '  rdM1.Checked = True Then
                '-----------------------------------------------------------------------------------
                ' M1 validation
                '-----------------------------------------------------------------------------------
                If dtM1.Text <> " " Then
                    FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtM1.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    If FindColumn Is Nothing Then

                        ErrorProvider.SetError(dtM1, "Selected M1 date is out of time-Line range.")
                        FormValiding = False
                        Exit Function
                    Else
                        ErrorProvider.SetError(dtM1, "")
                        FormValiding = True
                    End If
                End If

                If dtM1dc.Text = " " Then
                    ErrorProvider.SetError(dtM1dc, "M1DC field cannot be blank.")
                    FormValiding = False
                    Exit Function
                Else
                    FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtM1dc.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    If FindColumn Is Nothing Then
                        ErrorProvider.SetError(dtM1dc, "Selected M1DC is out of time-Line range.")
                        FormValiding = False
                        Exit Function
                    Else
                        ErrorProvider.SetError(dtM1dc, "")
                        FormValiding = True
                    End If
                End If
                'Else
                '    MessageBox.Show("Only VP/M1/TPV/DCV Buildphase dates can be added.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                '    FormValiding = False
                '    Exit Function
                'End If
        End Select
    End Function

    Private Sub btnClearMRD_Click(sender As Object, e As EventArgs) Handles btnClearMRD.Click
        Set_DateFormat(dtMRD, False)
        lblMRD.Text = "( )"
    End Sub

    Private Sub btnClearM1_Click(sender As Object, e As EventArgs) Handles btnClearM1.Click
        Set_DateFormat(dtM1, False)
        lblM1.Text = "( )"
    End Sub

    Private Sub btnClearM1dc_Click(sender As Object, e As EventArgs) Handles btnClearM1dc.Click
        Set_DateFormat(dtM1dc, False)
        lblM1dc.Text = "( )"
    End Sub

    Private Sub btnClearVP_Click(sender As Object, e As EventArgs) Handles btnClearVP.Click
        Set_DateFormat(dtVP, False)
        lblVp.Text = "( )"
    End Sub

    Private Sub btnClearPEC_Click(sender As Object, e As EventArgs) Handles btnClearPEC.Click
        Set_DateFormat(dtPec, False)
        lblPec.Text = "( )"
    End Sub

    Private Sub btnClearFEC_Click(sender As Object, e As EventArgs) Handles btnClearFEC.Click
        Set_DateFormat(dtFec, False)
        lblFec.Text = "( )"
    End Sub

    Private Sub dtMRD_ValueChanged(sender As Object, e As EventArgs) Handles dtMRD.ValueChanged
        dtMRD.CustomFormat = "dd-MMM-yyyy"
        lblMRD.Text = Form.DataCenter.GlobalValues.cal.GetWeekOfYear(dtMRD.Value, Form.DataCenter.GlobalValues.myCWR, myFirstDOW)

        Try
            Dim FindColumn As Excel.Range = Nothing
            FindColumn = Form.DataCenter.GlobalSections.DescriptionSection.Find(dtMRD.Value.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
            If FindColumn Is Nothing Then
                ErrorProvider.SetError(dtMRD, "Selected MRD is out of time-Line range.")
            Else
                ErrorProvider.SetError(dtMRD, "")
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddDates, ex.Message), "MRD Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
End Class