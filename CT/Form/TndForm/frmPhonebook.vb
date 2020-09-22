Imports System.Windows.Forms

Public Class frmPhonebook
    Dim dgv_change_flag As Boolean 'Flag to handle combobox events triggering at form load

    Dim DtRegion As System.Data.DataTable
    Dim DvRegion As System.Data.DataView

    'Datagridview cell value change event
    'To track the changes in the grid for save event.
    Private Sub dgvPhonebook_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPhonebook.CellValueChanged
        If e.RowIndex < 0 Or dgv_change_flag = False Then Exit Sub
        dgv_change_flag = False
        With dgvPhonebook.Rows(e.RowIndex)
            If .Cells(0).Value.ToString <> "" Then
                .Cells("Edited").Value = "Update"
            Else
                .Cells("Edited").Value = "Add"
            End If
        End With
        dgv_change_flag = True
    End Sub

    'Phone book form load event
    'To load the phone book data from database to gridview
    'And to restrict fields length
    Private Sub frmPhonebook_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            dgv_change_flag = False
            Dim dtPhoneBook As System.Data.DataTable
            Dim _Phonebook As New Data.Phonebook()
            dtPhoneBook = _Phonebook.SelectAll()
            If dtPhoneBook Is Nothing Then
                Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If
            dtPhoneBook.Columns.Add("Edited")
            dgvPhonebook.DataSource = dtPhoneBook
            dgvPhonebook.Columns(0).Visible = False 'pe90 field
            dgvPhonebook.Columns(4).Visible = False 'region id
            dgvPhonebook.Columns(6).Visible = False 'edited flag field
            dgvPhonebook.Columns(5).Visible = False 'region 

            DirectCast(dgvPhonebook.Columns("CDSID"), DataGridViewTextBoxColumn).MaxInputLength = 12
            DirectCast(dgvPhonebook.Columns("Fullname"), DataGridViewTextBoxColumn).MaxInputLength = 100
            DirectCast(dgvPhonebook.Columns("Tel"), DataGridViewTextBoxColumn).MaxInputLength = 50

            Dim _Region As New Data.Region()
            DtRegion = _Region.SelectAll()
            Dim cmb As New DataGridViewComboBoxColumn() 'Region (Dropdown field)
            cmb.HeaderText = "Region"
            cmb.Name = "cmbRegion"
            cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
            cmb.DataSource = DtRegion
            cmb.DisplayMember = "Regions"
            cmb.ValueMember = "pe27_Regions_PK"
            cmb.DataPropertyName = "pe27_Regions_PK"
            cmb.FillWeight = 60
            cmb.DisplayIndex = 7
            dgvPhonebook.Columns.Add(cmb)

            For i As Int16 = 0 To dgvPhonebook.Rows.Count - 2
                If dgvPhonebook.Rows(i).Cells(5).Value.ToString() <> "" Then
                    CType(Me.dgvPhonebook("cmbRegion", i), DataGridViewComboBoxCell).Value = CInt(Trim(dgvPhonebook.Rows(i).Cells(4).Value))
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPhonebook, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            dgv_change_flag = True
        End Try
    End Sub

    'Button Add click event
    'To store new/modified data in the grid
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Try
            dgv_change_flag = False
            Me.Cursor = Cursors.AppStarting
            Dim _Phonebook As New Data.Phonebook
            Dim _regionid As Integer
            For Each row In dgvPhonebook.Rows
                If row.cells(7).value IsNot Nothing Then
                    _regionid = row.cells(7).value
                Else
                    _regionid = Nothing
                End If
                If row.Cells("Edited").FormattedValue = "Update" Then
                    If _Phonebook.Update(pe90:=row.Cells(0).Value.ToString, CDSID:=row.cells(1).Value.ToString, FullName:=row.cells(2).value.ToString, Tel:=row.cells(3).value.ToString, pe27:=_regionid) = True Then
                    row.Cells("Edited").Value = ""
                Else
                    Throw New Exception("Error in Saving: " & Data.DataCenter.GlobalValues.message)
                    End If
                ElseIf row.cells("Edited").FormattedValue = "Add" Then
                    If _Phonebook.AddNew(CDSID:=row.cells(1).Value.ToString, FullName:=row.cells(2).value.ToString, Tel:=row.cells(3).value.ToString, pe27:=_regionid) = True Then
                        row.Cells("Edited").Value = ""
                    Else
                        Throw New Exception("Error in Saving: " & Data.DataCenter.GlobalValues.message)
                    End If
                End If
            Next
            Me.Cursor = Cursors.Default
            MessageBox.Show("Records saved.", "Phone book", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmPhonebook, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Button close click event - To close the form
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    'Form keydown event
    'For shortcut keys
    Private Sub frmPhonebook_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            btnClose_Click(sender, e)
        ElseIf e.KeyCode = Keys.F7 Then
            btnAdd.Focus()
            btnAdd_Click(sender, e)
        End If
    End Sub

    'Gridview cell leave event
    'To track the cell modified rows for save event
    Private Sub dgvPhonebook_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPhonebook.CellLeave
        If e.RowIndex < 0 Or dgv_change_flag = False Then Exit Sub
        dgv_change_flag = False
        With dgvPhonebook.Rows(e.RowIndex)
            If .Cells(0).Value.ToString <> "" Then
                .Cells("Edited").Value = "Update"
            Else
                .Cells("Edited").Value = "Add"
            End If
        End With
        dgv_change_flag = True
    End Sub

    Private Sub dgvPhonebook_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvPhonebook.DataError

    End Sub
End Class