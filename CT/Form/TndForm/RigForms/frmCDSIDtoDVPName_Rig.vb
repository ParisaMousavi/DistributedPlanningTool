Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmCDSIDtoDVPName_Rig
    Dim dgv_change_flag As Boolean
    Dim boolPhonebookload As Boolean
    Dim dtPhoneBook As System.Data.DataTable
    Dim boolPlanRefreshflag As Boolean

    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions

    'Sub: Load phonebook
    'To load phone book data from database to combobox and checkedlistbox
    Sub Load_Phonebook()
        Dim _Phonebook As New Data.Phonebook()
        dtPhoneBook = _Phonebook.SelectAll()
        If boolPhonebookload = True Then
            cbPMTLevel.DataSource = dtPhoneBook
            cbPMTLevel.DisplayMember = "CDSID"
            cbPMTLevel.ValueMember = "CDSID"
            cbPMTLevel.SelectedIndex = -1
            ChkListPhonebook.DataSource = dtPhoneBook
            ChkListPhonebook.DisplayMember = "CDSID"
            ChkListPhonebook.ValueMember = "CDSID"
        End If
    End Sub

    'Form load event - To load 
    Private Sub frmCDSIDtoDVPName_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            boolPhonebookload = False
            boolPlanRefreshflag = False

            cbPMTLevel.Hide()

            RemoveHandler cbPMTLevel.SelectedIndexChanged, AddressOf cbPMTLevel_SelectedIndexChanged

            dgvAssignCDS.Controls.Add(cbPMTLevel)
            dgv_change_flag = False
            lblHCID.Text = If(Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString, Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.HCID) ' It's in this case only for displaying
            Dim _Plan As New Data.RigPlan.Plan()

            Load_Phonebook() 'Now loads only data to datatable 'dtPhoneBook'

            cbPMTLevel.DataSource = dtPhoneBook
            cbPMTLevel.DisplayMember = "CDSID"
            cbPMTLevel.ValueMember = "CDSID"
            cbPMTLevel.Enabled = False

            ChkListPhonebook.DataSource = dtPhoneBook
            ChkListPhonebook.DisplayMember = "CDSID"
            ChkListPhonebook.ValueMember = "CDSID"

            Dim myDataTable As System.Data.DataTable
            myDataTable = _Plan.GetDvpTeamAndCdsid(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

            If myDataTable Is Nothing Then
                btnSave.Enabled = False
                Throw New Exception("No Data to display error: " & Data.DataCenter.GlobalValues.message)
            End If
            myDataTable.Columns.Add("Edited")
            dgvAssignCDS.Columns(0).DataPropertyName = "PmtGroupName"
            dgvAssignCDS.Columns(1).DataPropertyName = "DvpTeamName"
            dgvAssignCDS.Columns(2).DataPropertyName = "PMTLevel"
            dgvAssignCDS.Columns(3).DataPropertyName = "DNRLevel"
            dgvAssignCDS.Columns(4).DataPropertyName = "Edited"
            dgvAssignCDS.DataSource = myDataTable
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmCDSIDtoDVPName, ex.Message), "CT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            dgv_change_flag = True
            AddHandler cbPMTLevel.SelectedIndexChanged, AddressOf cbPMTLevel_SelectedIndexChanged
        End Try
    End Sub

    'Button save click event
    'To update new PMT/DNR level values to database
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            dgv_change_flag = False
            Panel1.Visible = False
            Me.Cursor = Cursors.AppStarting

            Dim _Plan As New Data.RigPlan.Plan()
            With dgvAssignCDS
                Dim strPMTlevel As String = ""
                Dim pe01 = Form.DataCenter.ProgramConfig.pe01
                Dim HCID = Form.DataCenter.ProgramConfig.HCID

                For Each row In dgvAssignCDS.Rows
                    If row.Cells("Edited").FormattedValue = "Y" Then
                        If IsDBNull(row.Cells(2).Value) = False Then strPMTlevel = row.Cells(2).Value
                        If _Plan.AssignCdsid2DvpTeam(pe01:=pe01, HCID:=HCID, MainBUildType:=Form.DataCenter.ProgramConfig.BuildType, PmtGroup:=row.cells(0).Value, DvpTeamName:=row.Cells(1).Value, PMTLevel:=strPMTlevel,
                                                     DNRLevel:=row.Cells(3).Value) = True Then
                            row.Cells("Edited").Value = ""
                            boolPlanRefreshflag = True
                        Else
                            MessageBox.Show("Error in Saving: " & Data.DataCenter.GlobalValues.message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        strPMTlevel = ""
                    End If
                Next
            End With
            MessageBox.Show("Records saved.", "CT", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmCDSIDtoDVPName, ex.Message), "CT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            dgv_change_flag = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Datagridview cell value change event
    'To track the changes in the gridview for save event
    Private Sub dgvAssignCDS_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignCDS.CellValueChanged
        If e.RowIndex < 0 Or dgv_change_flag = False Then Exit Sub
        dgv_change_flag = False
        With dgvAssignCDS.Rows(e.RowIndex)
            .Cells("Edited").Value = "Y"
            If e.ColumnIndex = 2 Then
                For Each row In dgvAssignCDS.Rows 'Based on PMT Group - All PMT Level cells will be updated with same value
                    If row.cells(0).Value = .Cells(0).Value Then
                        row.cells(2).Value = .Cells(2).Value
                        row.cells("Edited").Value = "Y"
                        If IsDBNull(row.Cells(3).Value) = True Then row.Cells(3).Value = ""
                    End If
                Next
            End If
        End With
        dgv_change_flag = True
    End Sub

    'Form keydown event
    'For shortcut keys
    Private Sub frmCDSIDtoDVPName_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                If Panel1.Visible = True Then
                    Panel1.Visible = False
                ElseIf cbPMTLevel.Visible = True Then
                    cbPMTLevel.Hide()
                Else
                    btnCancel_Click(sender, e)
                End If
            Case Keys.F7
                If btnSave.Enabled = True Then btnSave_Click(sender, e)
        End Select
    End Sub

    'Button cancel event - to close the form
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    'Datagridview cell enter event
    'To display CDSIS's in combobox  for PMT Level field
    'To display CDSIS's in checkedlistbox  for DNR Level field
    'To hide the combobox/checkedlistbox for other field selection 
    Private Sub dgvAssignCDS_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignCDS.CellEnter
        If e.ColumnIndex = 2 Then
            Panel1.Visible = False
            cbPMTLevel.Width = dgvAssignCDS.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True).Width
            cbPMTLevel.Height = dgvAssignCDS.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True).Height
            cbPMTLevel.Location = dgvAssignCDS.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True).Location
            cbPMTLevel.SelectedValue = dgvAssignCDS.CurrentCell.Value
            dgvAssignCDS.Columns(2).ReadOnly = True
            cbPMTLevel.Enabled = True
            cbPMTLevel.Show()
        ElseIf e.ColumnIndex = 3 Then
            dgvAssignCDS.Columns(3).ReadOnly = True
            cbPMTLevel.Enabled = False
            cbPMTLevel.Hide()
            Panel1.Visible = True

            For i As Int16 = 0 To ChkListPhonebook.Items.Count - 1
                ChkListPhonebook.SetItemChecked(i, False)
                If dgvAssignCDS.CurrentCell.Value.ToString.IndexOf(ChkListPhonebook.DataSource.Rows(i).item(1).ToString) >= 0 Then ' ChkListPhonebook.DataSource.Rows(i).item(1).ToString = dgvAssignCDS.CurrentCell.Value.ToString Then
                    ChkListPhonebook.SetItemChecked(i, True)
                End If
            Next
        Else
            cbPMTLevel.Hide()
            cbPMTLevel.Enabled = False
            dgvAssignCDS.Columns(2).ReadOnly = False
            dgvAssignCDS.Columns(3).ReadOnly = False
            Panel1.Visible = False
        End If
    End Sub

    'PMT level combobox index change event
    'To update the selection in gridview cell
    Private Sub cbPMTLevel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbPMTLevel.SelectedIndexChanged
        If IsDBNull(cbPMTLevel.SelectedValue) = False And cbPMTLevel.Enabled = True Then
            dgvAssignCDS.Columns(2).ReadOnly = False
            dgvAssignCDS.CurrentCell.Value = cbPMTLevel.SelectedValue
            dgvAssignCDS.Columns(2).ReadOnly = True
            cbPMTLevel.Enabled = False
            cbPMTLevel.Hide()
        End If
    End Sub

    'PMT level combobox text change event
    'To delete the old value in the cell/combobox
    Private Sub cbPMTLevel_TextChanged(sender As Object, e As EventArgs) Handles cbPMTLevel.TextChanged
        If dgv_change_flag = True And cbPMTLevel.Text = "" And cbPMTLevel.Enabled = True Then
            dgvAssignCDS.Columns(2).ReadOnly = False
            dgvAssignCDS.CurrentCell.Value = ""
            dgvAssignCDS.Columns(2).ReadOnly = True
            cbPMTLevel.Enabled = False
            cbPMTLevel.Hide()
        End If
    End Sub

    'Button Insert click event
    'To load the selected CDSIDS in checkedlistbox to DNRLevel cells in datagridview
    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Dim strSelection As String = ""

        For Each itemChecked In ChkListPhonebook.CheckedItems
            strSelection += itemChecked.Row.ItemArray(1).ToString & ";"
        Next

        If strSelection <> "" Then strSelection = strSelection.Substring(0, strSelection.Length - 1)

        dgvAssignCDS.CurrentCell.Value = strSelection

        For i As Int16 = 0 To ChkListPhonebook.Items.Count - 1
            ChkListPhonebook.SetItemCheckState(i, False)
        Next

        Panel1.Visible = False
        ''dgvAssignCDS.CurrentCell.Selected = True
        'dgvAssignCDS.Focus()
    End Sub

    'Datagridview scroll event - to hide the PMT level combobox
    Private Sub dgvAssignCDS_Scroll(sender As Object, e As ScrollEventArgs) Handles dgvAssignCDS.Scroll
        cbPMTLevel.Hide()
    End Sub

    'Button phone book click event - to show the phone book form
    'After the phonebook form close - load the new CDSIDs back to PMT level combobox and DNR level checkedlistbox
    Private Sub btnPhonebook_Click(sender As Object, e As EventArgs) Handles btnPhonebook.Click
        Panel1.Visible = False
        cbPMTLevel.Enabled = False
        cbPMTLevel.Hide()
        boolPhonebookload = True
        Dim obj As New frmPhonebook_Rig
        obj.ShowDialog()
        Load_Phonebook()
        boolPhonebookload = False
    End Sub

    'Form closing event - to refresh the plan area in excel
    Private Sub frmCDSIDtoDVPName_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            '-----------------------------------------------------Loading/Refreshing Plan Area
            If boolPlanRefreshflag = True Then

                NotifyIcon1.BalloonTipText = "Refreshing Plan area"
                NotifyIcon1.BalloonTipTitle = "CT"
                NotifyIcon1.Icon = SystemIcons.Information
                NotifyIcon1.ShowBalloonTip(300)
                NotifyIcon1.Visible = True

                If Form.DataCenter.GlobalValues.WS.AutoFilterMode = True Then
                    _GlobalFunctions.GetSearchFilter()
                    Form.DataCenter.GlobalValues.WS.AutoFilterMode = False
                End If

                Me.Cursor = Cursors.AppStarting
                Dim strMessage As String = String.Empty
                Dim _DrawTndPlanArea As Form.DisplayUtilities.DrawTndPlanArea = New Form.DisplayUtilities.DrawTndPlanArea()
                strMessage = _DrawTndPlanArea.LoadTndPlanAreaToWorkSheet(Nothing, Nothing)

                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                _DrawTndPlanArea.ApplyFormattingAfterLoading(Nothing, Nothing)

                'timeline section border color 
                With Form.DataCenter.GlobalValues.WS
                    .Range(.Cells(4, Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column),
                           .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column)).Select()
                    .Application.Selection.FillDown()

                    .Range(.Cells(4, Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column +
                                  Form.DataCenter.GlobalSections.TimeLineSection.Columns.Count - 1),
                           .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column +
                           Form.DataCenter.GlobalSections.TimeLineSection.Columns.Count - 1)).Select()
                    .Application.Selection.FillDown()

                    .Cells(1, 1).Select
                End With

                _GlobalFunctions.ReApplyFilter()
                If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)

                Dim _obj As New Form.DataCenter.ModuleFunction
                _obj.sbProtectPlan()
                Me.Cursor = Cursors.Default
            End If
            '-----------------------------------------------------Loading/Refreshing Plan Area
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmCDSIDtoDVPName, ex.Message), "CT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            NotifyIcon1.Visible = False
        End Try
    End Sub

    'DNR Level checkedlistbox - selection validation (only 10 entries can be selected)
    Private Sub ChkListPhonebook_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles ChkListPhonebook.ItemCheck
        If ChkListPhonebook.CheckedItems.Count > 9 And e.NewValue = CheckState.Checked Then
            e.NewValue = CheckState.Unchecked
            MessageBox.Show("Maximum 10 CDSID only can be selected", "CT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

End Class