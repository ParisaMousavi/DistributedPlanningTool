Imports System.Diagnostics
Imports System.Windows.Forms

Public Class frmHolidayPlan_Rig

    Dim dvHolidayTypes As System.Data.DataView
    Dim bolCheck As Boolean

    Dim DtRegion As System.Data.DataTable
    Dim DvRegion As System.Data.DataView
    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions

    'Dim bolFormLoad As Boolean

    'To Add/Update holiday information in database
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Me.Cursor = Cursors.AppStarting
            If CheckBlankRow() = False Then Exit Sub 'For manual entry row created

            Dim _PublicHoliday As New Data.PublicHoliday()

            For i As Int16 = 0 To dgvSpecific.Rows.Count - 1
                If dgvSpecific.Rows(i).Cells(10).Value = "A" Or dgvSpecific.Rows(i).Cells(10).Value = "" Then 'Add New Record
                    If _PublicHoliday.Add(HCID:=Form.DataCenter.ProgramConfig.HCID, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType, Regions:=dgvSpecific.Rows(i).Cells(6).Value, Country:=dgvSpecific.Rows(i).Cells(7).Value, State:=dgvSpecific.Rows(i).Cells(8).Value, CityName:=dgvSpecific.Rows(i).Cells(9).Value, PublicHolidayName:=dgvSpecific.Rows(i).Cells(0).Value, PublicHolidayType:=dgvSpecific.Rows(i).Cells(2).Value, PublicHolidayStart:=CType(Date.ParseExact(dgvSpecific.Rows(i).Cells(4).FormattedValue, "dd.MM.yyyy", Nothing), DateTime), PublicHolidayEnd:=CType(Date.ParseExact(dgvSpecific.Rows(i).Cells(5).FormattedValue, "dd.MM.yyyy", Nothing), DateTime), pe83:=dgvSpecific.Rows(i).Cells(1).Value) = True Then
                        dgvSpecific.Rows(i).Cells(10).Value = "U"
                    Else
                        MessageBox.Show("Error in Saving: " & Data.DataCenter.GlobalValues.message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                ElseIf dgvSpecific.Rows(i).Cells(10).Value = "Modified" Then 'Update Record
                    If _PublicHoliday.Update(pe85:=dgvSpecific.Rows(i).Cells(13).Value, Regions:=dgvSpecific.Rows(i).Cells(6).Value, Country:=dgvSpecific.Rows(i).Cells(7).Value, State:=dgvSpecific.Rows(i).Cells(8).Value, CityName:=dgvSpecific.Rows(i).Cells(9).Value, PublicHolidayName:=dgvSpecific.Rows(i).Cells(0).Value, PublicHolidayType:=dgvSpecific.Rows(i).Cells(2).Value, PublicHolidayStart:=CType(Date.ParseExact(dgvSpecific.Rows(i).Cells(4).FormattedValue, "dd.MM.yyyy", Nothing), DateTime), PublicHolidayEnd:=CType(Date.ParseExact(dgvSpecific.Rows(i).Cells(5).FormattedValue, "dd.MM.yyyy", Nothing), DateTime), MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = True Then
                        dgvSpecific.Rows(i).Cells(10).Value = "U"
                    Else
                        MessageBox.Show("Error in Update: " & Data.DataCenter.GlobalValues.message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            Next

            Globals.ThisAddIn.Application.ScreenUpdating = False

            _PublicHoliday.Populate(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

            'Dim Cls As New Form.DataCenter.GlobalFunctions
            'Dim frmProgress As New frmProgressbar
            'frmProgress.Show()
            'Dim dblProgress As Double
            'dblProgress = 100 / Form.DataCenter.GlobalValues.TotalRow
            'For i As Int16 = Form.DataCenter.ProgramConfig.FirstRow To Form.DataCenter.ProgramConfig.LastRow
            '    Cls.UpdateSection(i, i)
            '    frmProgress.UpdateProgressBar(dblProgress)
            '    frmProgress.Text = "Refreshing plan : " & CInt(frmProgress.SmoothProgressBar2.Value) & "% completed."
            'Next
            'frmProgress.Close()

            If Form.DataCenter.GlobalValues.WS.AutoFilterMode = True Then
                _GlobalFunctions.GetSearchFilter()
                Form.DataCenter.GlobalValues.WS.AutoFilterMode = False
            End If

            '-----------------------------------------------------Loading Plan Area
            Dim strMessage As String = String.Empty
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Dim _DrawTndPlanArea As Form.DisplayUtilities.DrawTndPlanArea = New Form.DisplayUtilities.DrawTndPlanArea()
            strMessage = _DrawTndPlanArea.LoadTndPlanAreaToWorkSheet(Nothing, Nothing)
            _DrawTndPlanArea.ApplyFormattingAfterLoading(Nothing, Nothing)
            '-----------------------------------------------------Loading Plan Area

            'Header area refresh
            Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader
            _DrawTndPlanHeader.ApplyHolidaysFlags()
            _DrawTndPlanHeader.ApplyGatewayFlags()

            '-----------------------------------------------------Formatting Timing flag column
            Dim currentcell As Excel.Range
            currentcell = Globals.ThisAddIn.Application.ActiveCell

            Dim intFirstCol As Integer = Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column
            Dim intLastCol As Integer = Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column + Form.DataCenter.GlobalSections.TimeLineSection.Columns.Count
            'intFirstCol -= 1
            intLastCol -= 1
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Cells(4, intFirstCol).Copy
            Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(5, intFirstCol), Form.DataCenter.GlobalValues.WS.Cells(5000, intFirstCol)).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

            Form.DataCenter.GlobalValues.WS.Cells(4, intLastCol).Copy
            Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(5, intLastCol), Form.DataCenter.GlobalValues.WS.Cells(5000, intLastCol)).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

            currentcell.Select()
            '-----------------------------------------------------Formatting Timing flag column

            _GlobalFunctions.ReApplyFilter()
            If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)

            Dim _obj As New Form.DataCenter.ModuleFunction
            Globals.ThisAddIn.Application.ScreenUpdating = False
            _obj.sbProtectPlan()
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            MessageBox.Show("Records saved.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHolidayPlan, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            Globals.ThisAddIn.Application.ScreenUpdating = True
        End Try
    End Sub

    'To close the form
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    'To Add generic holiday to specific list
    'Checking already added or not before adding based on pe83 key
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If CheckBlankRow() = False Then Exit Sub
        Dim bolExist As Boolean
        For Each drr As DataGridViewRow In dgvDefault.SelectedRows
            bolExist = False
            If dgvSpecific.Rows.Count > 0 Then
                For i As Int16 = 0 To dgvSpecific.Rows.Count - 1
                    If drr.Cells("pe83_PublicHolidays_PK").Value = dgvSpecific.Rows(i).Cells(1).Value Then
                        bolExist = True
                        MessageBox.Show("Already added to specific list.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit For
                    End If
                Next
            End If
            If bolExist = False Then
                dgvSpecific.Rows.Add()
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(0).Value = drr.Cells("HolidayName").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(3).Value = drr.Cells("HolidayType").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(4).Value = drr.Cells("StartDate").Value 'DateValue("5/12/2010") 
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(5).Value = drr.Cells("EndDate").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(6).Value = drr.Cells("Region").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(7).Value = drr.Cells("Country").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(8).Value = drr.Cells("State").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(9).Value = drr.Cells("City").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(10).Value = "A" 'Add - new record - used in button save click to call Add/Update

                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(1).Value = drr.Cells("pe83_PublicHolidays_PK").Value
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(2).Value = drr.Cells("PublicHolidayType").Value
            End If
        Next
        Lbl_Specific_Total.Text = "Total : " & dgvSpecific.Rows.Count
    End Sub

    'To Remove selected rows in Specific grid
    'Set as active=0 in database if already stored else removes from grid only
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        Dim bolRemoved As Boolean
        bolCheck = True
        bolRemoved = False
        With dgvSpecific
            Dim i As Int16
            Dim rowcount As Int16
            rowcount = dgvSpecific.Rows.Count - 1
            For i = 0 To rowcount
                If i > rowcount Then Exit For
                If .Rows(i).Cells(11).Value = True And (.Rows(i).Cells(10).Value = "U" Or .Rows(i).Cells(10).Value = "Modified") Then
                    Dim _PublicHoliday As New Data.PublicHoliday()
                    If _PublicHoliday.Delete(Pe85:= .Rows(i).Cells(13).Value) = True Then
                        dgvSpecific.Rows.Remove(dgvSpecific.Rows(i))
                        dgvSpecific.Refresh()
                        bolRemoved = True
                        If i > dgvSpecific.Rows.Count - 1 Then Exit For
                        i = i - 1
                        rowcount -= 1
                    Else
                        MessageBox.Show("Error in Removing: " & Data.DataCenter.GlobalValues.message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                ElseIf .Rows(i).Cells(11).Value = True Then
                    dgvSpecific.Rows.Remove(dgvSpecific.Rows(i))
                    dgvSpecific.Refresh()
                    bolRemoved = True
                    If i > dgvSpecific.Rows.Count - 1 Then Exit For
                    i = i - 1
                    rowcount -= 1
                End If
            Next
        End With

        If bolRemoved = True Then
            Lbl_Specific_Total.Text = "Total : " & dgvSpecific.Rows.Count
            MessageBox.Show("Records removed.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("No Record(s) removed. Please select row to remove.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        bolCheck = False
    End Sub

    'To display exceptional errors in datagridview dgvSpecific
    Private Sub dgvSpecific_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles dgvSpecific.DataError

        'MessageBox.Show("Error happened " _
        '& e.Context.ToString())

        'If (e.Context = DataGridViewDataErrorContexts.Commit) _
        'Then
        '    MessageBox.Show("Commit error")
        'End If
        'If (e.Context = DataGridViewDataErrorContexts _
        '.CurrentCellChange) Then
        '    MessageBox.Show("Cell change")
        'End If
        'If (e.Context = DataGridViewDataErrorContexts.Parsing) _
        'Then
        '    MessageBox.Show("parsing error")
        'End If
        'If (e.Context =
        'DataGridViewDataErrorContexts.LeaveControl) Then
        '    MessageBox.Show("leave control error")
        'End If

        'If (TypeOf (e.Exception) Is System.Data.ConstraintException) Then
        '    Dim view As DataGridView = CType(sender, DataGridView)
        '    view.Rows(e.RowIndex).ErrorText = "an error"
        '    view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
        '    .ErrorText = "an error"

        '    e.ThrowException = False
        'End If
    End Sub

    'Form load - To load holiday information from database to grid/combobox
    Private Sub frmHolidayPlan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'bolFormLoad = True
        FillDefaultHolidays() 'Common holidays from DB
        FillHolidayType() 'Holiday Type, Start/End Date
        FillRegionDetails() 'Region, Country, State, City, RecordState, Select(checkbox), Duration(textbox)
        FillSpecificHolidays() 'Specific holidays chosen for the plan

        Lbl_Generic_Total.Text = "Total : " & dgvDefault.Rows.Count
        Lbl_Specific_Total.Text = "Total : " & dgvSpecific.Rows.Count
        'bolFormLoad = False
    End Sub

    'Load Region,Country,State & City in Specific holidays grid dropdown boxes
    Sub FillRegionDetails()
        Dim _PublicHoliday As New Data.PublicHoliday()

        DtRegion = _PublicHoliday.GetAllLocations

        If DtRegion Is Nothing Then
            Throw New Exception(Data.DataCenter.GlobalValues.message)
        End If

        DvRegion = New System.Data.DataView(DtRegion)
        DtRegion = DvRegion.ToTable(True, "Regions")

        'dtRegions = myView.ToTable(True, "Regions")

        Dim cmb As New DataGridViewComboBoxColumn() 'Region (Dropdown field)
        cmb.HeaderText = "Region"
        cmb.Name = "cmbRegion"
        cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        'myDataTable.Rows.Add("")
        cmb.DataSource = DtRegion
        'cmb.Items.Add("")
        cmb.DisplayMember = "Regions"
        cmb.FillWeight = 60
        cmb.DisplayIndex = 4
        dgvSpecific.Columns.Add(cmb)

        DtRegion = DvRegion.ToTable(True, "Country")

        Dim cmbCountry As New DataGridViewComboBoxColumn() 'Region (Dropdown field)
        cmbCountry.HeaderText = "Country"
        cmbCountry.Name = "cmbCountry"
        cmbCountry.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox

        cmbCountry.DataSource = DtRegion
        cmbCountry.DisplayMember = "Country"
        cmbCountry.DisplayIndex = 5
        dgvSpecific.Columns.Add(cmbCountry)

        DtRegion = DvRegion.ToTable(True, "State")

        Dim cmbState As New DataGridViewComboBoxColumn() 'State (Dropdown field)
        cmbState.HeaderText = "State"
        cmbState.Name = "cmbState"
        cmbState.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        cmbState.DataSource = DtRegion
        cmbState.DisplayMember = "State"
        cmbState.DisplayIndex = 6
        dgvSpecific.Columns.Add(cmbState)

        DtRegion = DvRegion.ToTable(True, "CityName")

        Dim cmbCity As New DataGridViewComboBoxColumn() 'City (Dropdown field)
        cmbCity.HeaderText = "City"
        cmbCity.Name = "cmbCity"
        cmbCity.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        cmbCity.DataSource = DtRegion
        cmbCity.DisplayMember = "CityName"
        cmbCity.DisplayIndex = 7
        dgvSpecific.Columns.Add(cmbCity)

        Dim cmbRecordState As New DataGridViewTextBoxColumn() 'RecordState (text field)
        cmbRecordState.HeaderText = "RecordState"
        cmbRecordState.Name = "cmbRecordState"
        cmbRecordState.DisplayIndex = 8
        cmbRecordState.Visible = False
        dgvSpecific.Columns.Add(cmbRecordState)

        Dim cmbSelect As New DataGridViewCheckBoxColumn()
        cmbSelect.HeaderText = "Select"
        cmbSelect.Name = "cmbSelect"
        cmbSelect.DisplayIndex = 0
        cmbSelect.Visible = True
        cmbSelect.FillWeight = 40
        dgvSpecific.Columns.Add(cmbSelect)

        Dim txtDuration As New DataGridViewTextBoxColumn()
        txtDuration.HeaderText = "Duration"
        txtDuration.Name = "txtDuration"
        txtDuration.DisplayIndex = 9
        txtDuration.Visible = True
        txtDuration.FillWeight = 54
        dgvSpecific.Columns.Add(txtDuration)
    End Sub

    'Loading Holiday Types to Specific DataGrid column as DropdownList
    'Creating Calendar Control in Specific DataGrid for State/End Date
    Sub FillHolidayType()
        Dim _PublicHolidayType As New Data.PublicHolidayType()

        Dim myDataTable As System.Data.DataTable

        'Retrieve holiday type 
        myDataTable = _PublicHolidayType.GetAll()

        If myDataTable Is Nothing Then
            Throw New Exception(Data.DataCenter.GlobalValues.message)
        End If

        dvHolidayTypes = Nothing
        dvHolidayTypes = New System.Data.DataView(myDataTable)

        Dim cmb As New DataGridViewComboBoxColumn() 'Holiday Type creation (Dropdown field)
        cmb.HeaderText = "Holiday Type"
        cmb.Name = "cmbHolidayType"
        cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        cmb.DataSource = dvHolidayTypes
        cmb.DisplayMember = "PublicHolidayTypeFullName"
        cmb.DisplayIndex = 1
        dgvSpecific.Columns.Add(cmb)

        Dim dt_StartCol = New CalendarColumn() 'StartDate creation (DateTime field)
        dt_StartCol.DisplayIndex = 2
        dt_StartCol.Name = "StartDate"
        dt_StartCol.HeaderText = "Start Date"
        dt_StartCol.FillWeight = 80
        'Columns[2].DefaultCellStyle.Format = "MM/dd/yyyy HH:mm:ss";
        dt_StartCol.DefaultCellStyle.Format = "dd.MM.yyyy"

        dgvSpecific.Columns.Add(dt_StartCol)

        Dim dt_EndCol = New CalendarColumn() 'EndDate creation (DateTime field)
        dt_EndCol.DisplayIndex = 3
        dt_EndCol.Name = "EndDate"
        dt_EndCol.HeaderText = "End Date"
        dt_EndCol.FillWeight = 80
        dgvSpecific.Columns.Add(dt_EndCol)
    End Sub

    'To fill specific holidays (for the current hcid) from database
    Sub FillSpecificHolidays()
        Try
            Dim _PublicHoliday As New Data.PublicHoliday()

            Dim myDataTable As System.Data.DataTable

            myDataTable = _PublicHoliday.GetPlanPublicHolidays(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
            If myDataTable Is Nothing Then
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If

            Dim txt85 As New DataGridViewTextBoxColumn()
            txt85.HeaderText = "pe85"
            txt85.Name = "txtPe85"
            txt85.Visible = False
            dgvSpecific.Columns.Add(txt85)

            For i As Int16 = 0 To myDataTable.Rows.Count - 1
                dgvSpecific.Rows.Add()
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(0).Value = myDataTable.Rows(i).Item("PublicHolidayName").ToString
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(3).Value = myDataTable.Rows(i).Item("PublicHolidayTypeFullName").ToString '----------------
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(4).Value = DateValue(myDataTable.Rows(i).Item("PublicHolidayStart")).ToString("dd.MM.yyyy")
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(5).Value = DateValue(myDataTable.Rows(i).Item("PublicHolidayEnd")).ToString("dd.MM.yyyy")
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(6).Value = myDataTable.Rows(i).Item("Cbg").ToString
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(7).Value = myDataTable.Rows(i).Item("Country").ToString
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(8).Value = myDataTable.Rows(i).Item("State").ToString
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(9).Value = myDataTable.Rows(i).Item("City").ToString
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(10).Value = "U" 'Update - record
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(13).Value = myDataTable.Rows(i).Item("pe85_TnDProgramHolidays_PK").ToString 'For update 85, add 83 key 'Conflicts with 83 & 85 so removed
                'dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(1).Value = myDataTable.Rows(i).Item("pe85_TnDProgramHolidays_PK").ToString 'For update 85, add 83 key 'Conflicts with 83 & 85 so removed
                dgvSpecific.Rows(dgvSpecific.Rows.Count - 1).Cells(2).Value = myDataTable.Rows(i).Item("PublicHolidayType").ToString
            Next

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHolidayPlan, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'To load default (generic) holidays from database to grid dgvDefault
    Sub FillDefaultHolidays()
        Try
            Dim _PublicHoliday As New Data.PublicHoliday()

            Dim myDataTable As System.Data.DataTable
            Dim myView As System.Data.DataView

            myDataTable = _PublicHoliday.GetGenericPublicHolidays("FOE")
            If myDataTable Is Nothing Then
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If

            myView = New System.Data.DataView(myDataTable)
            myDataTable = myView.ToTable(False, "HolidayName", "HolidayType", "StartDate", "EndDate", "Region", "Country", "State", "City", "pe83_PublicHolidays_PK", "PublicHolidayType", "Duration")

            dgvDefault.DataSource = myDataTable

            lblBuildPhase.Text = Form.DataCenter.ProgramConfig.BuildPhase
            lblHCID.Text = If(Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString, Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.HCID)  ' In this case it's only for displaying
            lblHCName.Text = Form.DataCenter.ProgramConfig.HCIDName

            dgvDefault.Columns(8).Visible = False 'pe83_PublicHolidays_PK
            dgvDefault.Columns(9).Visible = False 'PublicHolidayType
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHolidayPlan, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Form keydown event - for keyboard shortcut keys
    Private Sub frmHolidayPlan_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                btnCancel_Click(sender, e)
            Case Keys.F4
                btnAdd_Click(sender, e)
            Case Keys.F7
                btnSave_Click(sender, e)
            Case Keys.F5
                btnRemove_Click(sender, e)
            Case Keys.F6
                btnAddRow_Click(sender, e)
            Case Keys.F2
                dgvDefault.Focus()
            Case Keys.F3
                dgvSpecific.Focus()
        End Select
    End Sub

    'To check blank cells 
    'Start/end date validation (greater check)
    'Invalid character check
    Function CheckBlankRow() As Boolean
        If dgvSpecific.Rows.Count <= 0 Then
            CheckBlankRow = True
            Exit Function
        End If
        CheckBlankRow = False

        'Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
        'Blank check for all rows
        For j As Int16 = 0 To dgvSpecific.Rows.Count - 1
            For i As Int16 = 0 To dgvSpecific.Columns.Count - 1
                If i <> 1 And i <> 2 And i <> 5 And i <> 6 And i <> 10 And i <> 11 And i <> 13 Then
                    If IsNothing(dgvSpecific.Rows(j).Cells(i).Value) = True Then '1 is pe83 2 is holiday type key 5 is pe83 6 is holidaytypekey (hidden columns)
                        dgvSpecific.Rows(j).Selected = True
                        MessageBox.Show("Please fill all columns in the selected row.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Function
                    ElseIf dgvSpecific.Rows(j).Cells(i).FormattedValue = "" Then
                        dgvSpecific.Rows(j).Selected = True
                        MessageBox.Show("Please fill all columns in the selected row.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Function
                    End If
                End If
            Next
            If dgvSpecific.Rows(j).Cells(12).Value <= 0 Then
                dgvSpecific.Rows(j).Selected = True
                MessageBox.Show("Start date cannot be greater than End date.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
            End If
            If _GlobalFunctions.ContainsInvalidChar(dgvSpecific.Rows(j).Cells(0).Value) = True Then
                dgvSpecific.Rows(j).Selected = True
                MessageBox.Show(("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ; "" ' & ; ~ ` < >."), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
            End If
        Next
        CheckBlankRow = True
    End Function

    'To Add rows for manual entry
    Private Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRow.Click
        If CheckBlankRow() = False Then Exit Sub
        dgvSpecific.Rows.Add()
        Lbl_Specific_Total.Text = "Total : " & dgvSpecific.Rows.Count
    End Sub

    'For deleting rows in grid - to remove in database also
    Private Sub dgvSpecific_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles dgvSpecific.UserDeletingRow
        If e.Row.Cells(10).Value = "U" Or e.Row.Cells(10).Value = "Modified" Then
            Dim _PublicHoliday As New Data.PublicHoliday()
            If _PublicHoliday.Delete(Pe85:=e.Row.Cells(13).Value) = True Then
                Lbl_Specific_Total.Text = "Total : " & dgvSpecific.Rows.Count
            Else
                MessageBox.Show("Error in Deleting: " & Data.DataCenter.GlobalValues.message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub dgvSpecific_CurrentCellChanged(sender As Object, e As EventArgs) Handles dgvSpecific.CurrentCellChanged
        If IsNothing(dgvSpecific.CurrentCell) = False And bolCheck = False Then
            If dgvSpecific.Rows(dgvSpecific.CurrentCell.RowIndex).Cells(10).Value = "U" Then
                dgvSpecific.Rows(dgvSpecific.CurrentCell.RowIndex).Cells(10).Value = "Modified"
            End If
        End If
    End Sub

    Private Sub dgvSpecific_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSpecific.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Try
            Select Case e.ColumnIndex
                Case 3
                    dvHolidayTypes.Sort = "PublicHolidayTypeFullName"
                    dgvSpecific.Rows(e.RowIndex).Cells(2).Value = dvHolidayTypes(dvHolidayTypes.Find(dgvSpecific.Rows(e.RowIndex).Cells(3).Value))(1)
                Case 6, 7, 8 'Region cell
                    DvRegion.RowFilter = ""
                    Dim cell As DataGridViewComboBoxCell = CType(dgvSpecific.Rows(e.RowIndex).Cells(7), DataGridViewComboBoxCell)
                    DvRegion.RowFilter = "Regions = '" & Trim(dgvSpecific.Rows(e.RowIndex).Cells(6).Value) & "'"
                    DtRegion = DvRegion.ToTable(True, "Country")
                    cell.DataSource = DtRegion
                    'Case 7 'Country cell
                    DvRegion.RowFilter = ""
                    cell = CType(dgvSpecific.Rows(e.RowIndex).Cells(8), DataGridViewComboBoxCell)
                    DvRegion.RowFilter = "Country = '" & Trim(dgvSpecific.Rows(e.RowIndex).Cells(7).Value) & "' and Regions = '" & Trim(dgvSpecific.Rows(e.RowIndex).Cells(6).Value) & "' "
                    DtRegion = DvRegion.ToTable(True, "State")
                    cell.DataSource = DtRegion
                    'Case 8 'State cell
                    DvRegion.RowFilter = ""
                    cell = CType(dgvSpecific.Rows(e.RowIndex).Cells(9), DataGridViewComboBoxCell)
                    DvRegion.RowFilter = "State = '" & Trim(dgvSpecific.Rows(e.RowIndex).Cells(8).Value) & "' and Country = '" & Trim(dgvSpecific.Rows(e.RowIndex).Cells(7).Value) & "' and Regions = '" & Trim(dgvSpecific.Rows(e.RowIndex).Cells(6).Value) & "' "
                    DtRegion = DvRegion.ToTable(True, "CityName")
                    cell.DataSource = DtRegion
            End Select

            If dgvSpecific.Rows(e.RowIndex).Cells(4).FormattedValue <> "" And dgvSpecific.Rows(e.RowIndex).Cells(5).FormattedValue <> "" Then
                dgvSpecific.Rows(e.RowIndex).Cells(12).Value = DateDiff(DateInterval.Day, Date.ParseExact(dgvSpecific.Rows(e.RowIndex).Cells(4).FormattedValue, "dd.MM.yyyy", Nothing), Date.ParseExact(dgvSpecific.Rows(e.RowIndex).Cells(5).FormattedValue, "dd.MM.yyyy", Nothing)) + 1
                If IsNothing(dgvSpecific.Rows(e.RowIndex).Cells(4).Value) = False And IsNothing(dgvSpecific.Rows(e.RowIndex).Cells(5).Value) = False Then
                    If e.ColumnIndex <> 12 And e.ColumnIndex <> 10 Then
                        If dgvSpecific.Rows(e.RowIndex).Cells(12).Value <= 0 Then
                            MessageBox.Show("End date should be equal to or greater than Start date.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmHolidayPlan, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvSpecific_Leave(sender As Object, e As EventArgs) Handles dgvSpecific.Leave
        For Each r As DataGridViewRow In Me.dgvSpecific.Rows
            If r.Cells(4).FormattedValue <> "" And r.Cells(5).FormattedValue <> "" Then
                r.Cells(12).Value = DateDiff(DateInterval.Day, Date.ParseExact(r.Cells(4).FormattedValue, "dd.MM.yyyy", Nothing), Date.ParseExact(r.Cells(5).FormattedValue, "dd.MM.yyyy", Nothing)) + 1
                If IsNothing(r.Cells(4).Value) = False And IsNothing(r.Cells(5).Value) = False Then
                    If r.Cells(12).Value <= 0 Then
                        MessageBox.Show("End date should be equal to or greater than Start date.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If
            End If
        Next
    End Sub
End Class

'Public Class CalendarColumn
'    Inherits DataGridViewColumn

'    Public Sub New()
'        MyBase.New(New CalendarCell())
'    End Sub

'    Public Overrides Property CellTemplate() As DataGridViewCell
'        Get
'            Return MyBase.CellTemplate
'        End Get
'        Set(ByVal value As DataGridViewCell)

'            ' Ensure that the cell used for the template is a CalendarCell.
'            If (value IsNot Nothing) AndAlso
'                Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) _
'                Then
'                Throw New InvalidCastException("Must be a CalendarCell")
'            End If
'            MyBase.CellTemplate = value

'        End Set
'    End Property

'End Class

'Public Class CalendarCell
'    Inherits DataGridViewTextBoxCell

'    Public Sub New()
'        ' Use the short date format.
'        'Me.Style.Format = "dd/MM/yyyy"
'        'Me.Style.Format = "yyyy/MM/dd h:mm:ss tt"
'        Me.Style.Format = "dd.MM.yyyy"
'    End Sub

'    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer,
'        ByVal initialFormattedValue As Object,
'        ByVal dataGridViewCellStyle As DataGridViewCellStyle)

'        ' Set the value of the editing control to the current cell value.
'        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue,
'            dataGridViewCellStyle)

'        Dim ctl As CalendarEditingControl =
'            CType(DataGridView.EditingControl, CalendarEditingControl)

'        ' Use the default row value when Value property is null.
'        If (Me.Value Is Nothing OrElse IsDBNull(Me.Value)) Then
'            ctl.Value = CType(Me.DefaultNewRowValue, DateTime)
'            'ctl.Value = CType(Date.ParseExact(Me.DefaultNewRowValue, "dd.MM.yyyy", Nothing), DateTime)
'        Else
'            If Date.TryParseExact(Value, "dd.MM.yyyy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, Nothing) = False Then
'                ctl.Value = CType(Me.Value, DateTime)
'            Else
'                ctl.Value = CType(Date.ParseExact(Value, "dd.MM.yyyy", Nothing), DateTime)
'            End If

'            'CType(Date.Parse(Value, System.Globalization.CultureInfo.InvariantCulture), DateTime)
'            'CreateSpecificCulture("it-IT")), DateTime) 'CType(Me.Value, DateTime)
'        End If
'    End Sub

'    Public Overrides ReadOnly Property EditType() As Type
'        Get
'            ' Return the type of the editing control that CalendarCell uses.
'            Return GetType(CalendarEditingControl)
'        End Get
'    End Property

'    Public Overrides ReadOnly Property ValueType() As Type
'        Get
'            ' Return the type of the value that CalendarCell contains.
'            Return GetType(DateTime)
'        End Get
'    End Property

'    Public Overrides ReadOnly Property DefaultNewRowValue() As Object
'        Get
'            ' Use the current date and time as the default value.
'            Return DateTime.Now
'        End Get
'    End Property

'End Class

'Class CalendarEditingControl
'    Inherits DateTimePicker
'    Implements IDataGridViewEditingControl

'    Private dataGridViewControl As DataGridView
'    Private valueIsChanged As Boolean = False
'    Private rowIndexNum As Integer

'    Public Sub New()
'        Me.Format = DateTimePickerFormat.Custom
'        Me.CustomFormat = "dd.MM.yyyy"
'    End Sub

'    Public Property EditingControlFormattedValue() As Object _
'        Implements IDataGridViewEditingControl.EditingControlFormattedValue

'        Get
'            Return Me.Value.ToString("dd.MM.yyyy")
'        End Get

'        Set(ByVal value As Object)
'            Try
'                ' This will throw an exception of the string is 
'                ' null, empty, or not in the format of a date.
'                Me.Value = DateTime.Parse(CStr(value))
'            Catch
'                ' In the case of an exception, just use the default
'                ' value so we're not left with a null value.
'                Me.Value = DateTime.Now
'            End Try
'        End Set

'    End Property

'    Public Function GetEditingControlFormattedValue(ByVal context _
'        As DataGridViewDataErrorContexts) As Object _
'        Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

'        'If System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd.MM.yyyy" Or System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy" Then
'        If Mid(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern, 1, 1).ToLower = "d" Then
'            Return Me.Value.ToString("dd.MM.yyyy") '("MM/dd/yyyy") '
'        Else
'            Return Me.Value.ToString("MM/dd/yyyy")
'        End If

'        'If Date.TryParseExact(Value, "dd.MM.yyyy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, Nothing) = False Then
'        '    Return CType(Me.Value, DateTime)
'        'Else
'        '    Return CType(Date.ParseExact(Value, "dd.MM.yyyy", Nothing), DateTime)
'        'End If
'    End Function

'    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As _
'        DataGridViewCellStyle) _
'        Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

'        Me.Font = dataGridViewCellStyle.Font
'        Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
'        Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

'    End Sub

'    Public Property EditingControlRowIndex() As Integer _
'        Implements IDataGridViewEditingControl.EditingControlRowIndex

'        Get
'            Return rowIndexNum
'        End Get
'        Set(ByVal value As Integer)
'            rowIndexNum = value
'        End Set

'    End Property

'    Public Function EditingControlWantsInputKey(ByVal key As Keys,
'        ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
'        Implements IDataGridViewEditingControl.EditingControlWantsInputKey

'        ' Let the DateTimePicker handle the keys listed.
'        Select Case key And Keys.KeyCode
'            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right,
'                Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

'                Return True

'            Case Else
'                Return Not dataGridViewWantsInputKey
'        End Select

'    End Function

'    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
'        Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

'        ' No preparation needs to be done.

'    End Sub

'    Public ReadOnly Property RepositionEditingControlOnValueChange() _
'        As Boolean Implements _
'        IDataGridViewEditingControl.RepositionEditingControlOnValueChange

'        Get
'            Return False
'        End Get

'    End Property

'    Public Property EditingControlDataGridView() As DataGridView _
'        Implements IDataGridViewEditingControl.EditingControlDataGridView

'        Get
'            Return dataGridViewControl
'        End Get
'        Set(ByVal value As DataGridView)
'            dataGridViewControl = value
'        End Set

'    End Property

'    Public Property EditingControlValueChanged() As Boolean _
'        Implements IDataGridViewEditingControl.EditingControlValueChanged

'        Get
'            Return valueIsChanged
'        End Get
'        Set(ByVal value As Boolean)
'            valueIsChanged = value
'        End Set

'    End Property

'    Public ReadOnly Property EditingControlCursor() As Cursor _
'        Implements IDataGridViewEditingControl.EditingPanelCursor

'        Get
'            Return MyBase.Cursor
'        End Get

'    End Property

'    Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)

'        ' Notify the DataGridView that the contents of the cell have changed.
'        valueIsChanged = True
'        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
'        MyBase.OnValueChanged(eventargs)

'    End Sub

'End Class

'Public Class CalendarColumn
'    Inherits DataGridViewColumn
'    Private mFormat As String = ""

'    <System.ComponentModel.Category("Behavior"),
'    System.ComponentModel.Description("Date time format"),
'    System.ComponentModel.DefaultValue(GetType(String), "d")>
'    Public Property DateFormat() As String
'        Get
'            Return mFormat
'        End Get
'        Set(ByVal value As String)
'            mFormat = value
'        End Set
'    End Property
'    ''' <summary>
'    ''' </summary>
'    ''' <returns></returns>
'    ''' <remarks>
'    ''' kevininstructor
'    ''' This is needed to persist our custom property DateFormat
'    ''' </remarks>
'    Public Overrides Function Clone() As Object
'        Dim TheCopy As CalendarColumn = DirectCast(MyBase.Clone(), CalendarColumn)
'        TheCopy.DateFormat = Me.DateFormat
'        Return TheCopy
'    End Function
'    Public Sub New()
'        MyBase.New(New CalendarCell())
'    End Sub
'    Public Overrides Property CellTemplate() As DataGridViewCell
'        Get
'            Return MyBase.CellTemplate
'        End Get
'        Set(ByVal value As DataGridViewCell)
'            ' Ensure that the cell used for the template is a CalendarCell.
'            If Not (value Is Nothing) AndAlso Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) Then
'                Throw New InvalidCastException("Must be a CalendarCell")
'            End If
'            MyBase.CellTemplate = value
'        End Set
'    End Property
'End Class
'Public Class CalendarCell
'    Inherits DataGridViewTextBoxCell
'    Public Sub New()
'    End Sub
'    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As DataGridViewCellStyle)
'        ' Set the value of the editing control to the current cell value.
'        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

'        Dim TheControl As CalendarEditingControl = CType(DataGridView.EditingControl, CalendarEditingControl)
'        If Not Me.Value.GetType Is GetType(DateTime) Then
'            Me.Value = Now
'        End If

'        TheControl.Value = CType(Me.Value, DateTime)
'        Dim MyOwner As CalendarColumn = CType(OwningColumn, CalendarColumn)
'        Me.Style.Format = MyOwner.DateFormat
'        TheControl.Format = DateTimePickerFormat.Custom
'        TheControl.CustomFormat = MyOwner.DateFormat
'    End Sub
'    Public Overrides ReadOnly Property EditType() As Type
'        Get
'            ' Return the type of the editing contol that CalendarCell uses.
'            Return GetType(CalendarEditingControl)
'        End Get
'    End Property
'    Public Overrides ReadOnly Property ValueType() As Type
'        Get
'            ' Return the type of the value that CalendarCell contains.
'            Return GetType(DateTime)
'        End Get
'    End Property
'    Public Overrides ReadOnly Property DefaultNewRowValue() As Object
'        Get
'            ' Kevininstructor changed from  "Use the current date and time as the default value" to DbNullValue
'            Return DBNull.Value
'        End Get
'    End Property
'End Class
'''' <summary>
'''' Provides Calendar popup within the GridView.
'''' </summary>
'''' <remarks></remarks>
'Class CalendarEditingControl
'    Inherits DateTimePicker
'    Implements IDataGridViewEditingControl

'    Private dataGridViewControl As DataGridView
'    Private valueIsChanged As Boolean = False
'    Private rowIndexNumber As Integer

'    Public Sub New()
'        Me.Format = DateTimePickerFormat.Custom
'    End Sub
'    Public Property EditingControlFormattedValue() As Object Implements IDataGridViewEditingControl.EditingControlFormattedValue
'        Get
'            Return Me.Value.ToString(Me.CustomFormat)
'        End Get
'        Set(ByVal value As Object)
'            If TypeOf value Is [String] Then
'                Me.Value = DateTime.Parse(CStr(value))
'            End If
'        End Set
'    End Property
'    Public Function GetEditingControlFormattedValue(ByVal context As DataGridViewDataErrorContexts) As Object _
'        Implements IDataGridViewEditingControl.GetEditingControlFormattedValue
'        Return Me.Value.ToString(Me.CustomFormat)
'    End Function
'    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As DataGridViewCellStyle) Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl
'        Me.Font = dataGridViewCellStyle.Font
'        Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
'        Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor
'    End Sub
'    Public Property EditingControlRowIndex() As Integer Implements IDataGridViewEditingControl.EditingControlRowIndex
'        Get
'            Return rowIndexNumber
'        End Get
'        Set(ByVal value As Integer)
'            rowIndexNumber = value
'        End Set
'    End Property
'    Public Function EditingControlWantsInputKey(ByVal key As Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean Implements IDataGridViewEditingControl.EditingControlWantsInputKey
'        ' Let the DateTimePicker handle the keys listed.
'        Select Case key And Keys.KeyCode
'            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp
'                Return True
'            Case Else
'                Return False
'        End Select
'    End Function
'    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) Implements IDataGridViewEditingControl.PrepareEditingControlForEdit
'        ' No preparation needs to be done.
'    End Sub
'    Public ReadOnly Property RepositionEditingControlOnValueChange() As Boolean Implements IDataGridViewEditingControl.RepositionEditingControlOnValueChange
'        Get
'            Return False
'        End Get
'    End Property
'    Public Property EditingControlDataGridView() As DataGridView Implements IDataGridViewEditingControl.EditingControlDataGridView
'        Get
'            Return dataGridViewControl
'        End Get
'        Set(ByVal value As DataGridView)
'            dataGridViewControl = value
'        End Set
'    End Property
'    Public Property EditingControlValueChanged() As Boolean Implements IDataGridViewEditingControl.EditingControlValueChanged
'        Get
'            Return valueIsChanged
'        End Get
'        Set(ByVal value As Boolean)
'            valueIsChanged = value
'        End Set
'    End Property
'    Public ReadOnly Property EditingControlCursor() As Cursor Implements IDataGridViewEditingControl.EditingPanelCursor
'        Get
'            Return MyBase.Cursor
'        End Get
'    End Property
'    Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)
'        ' Notify the DataGridView that the contents of the cell have changed.
'        valueIsChanged = True
'        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
'        MyBase.OnValueChanged(eventargs)
'    End Sub
'End Class
