Imports System.Windows.Forms
Imports System.Drawing
Public Class frmSearch

    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
    Dim _obj As New Form.DataCenter.ModuleFunction

    'Purpose : To start search by 'Enter' click on search text box
    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnSearch_Click(sender, e)
        End If
    End Sub

    'Purpose : To apply filter for selected records (find results)
    Private Sub sbDoFilter(findAddress As String)
        Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
        Dim rngFind As Excel.Range, intCnt As Integer

        Dim colFindRows As New Collection, firstAddress As String

        With Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalSections.TimeLineSection.Address.ToString.Split(":")(0) & ":" & (Form.DataCenter.GlobalValues.WS.UsedRange.Address).ToString.Split(":")(1))
            rngFind = .Find(txtSearch.Text, , Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
            If Not rngFind Is Nothing Then
                firstAddress = rngFind.Address
                Do
                    If Not _GlobalFunctions.colContains(colFindRows, CStr(rngFind.Row)) Then
                        colFindRows.Add(rngFind.Row, CStr(rngFind.Row))
                    End If
                    rngFind = .FindNext(rngFind)
                Loop While Not rngFind Is Nothing And rngFind.Address <> firstAddress
            End If
            Form.DataCenter.GlobalValues.WS.Range("5:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).EntireRow.Hidden = True
            For intCnt = 1 To colFindRows.Count
                Form.DataCenter.GlobalValues.WS.Rows(colFindRows.Item(intCnt)).EntireRow.Hidden = False
            Next
        End With

        txtSearch.Focus()
        _obj.sbProtectPlan()
    End Sub

    'Purpose : Button search click event - To find the key text in the excel sheet and filter the results
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Dim FoundCells As Excel.Range
            Dim FoundCell As Excel.Range

            If Strings.Len(Strings.Trim(txtSearch.Text)) > 0 Then
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                Form.DataCenter.GlobalValues.WS.Cells.FormatConditions.Delete()

                Form.DataCenter.GlobalValues.WS.Range("5:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).EntireRow.Hidden = False

                Dim objFC As Microsoft.Office.Interop.Excel.FormatCondition

                FoundCells = FindAll(SearchRange:=Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalSections.TimeLineSection.Address.ToString.Split(":")(0) & ":" & (Form.DataCenter.GlobalValues.WS.UsedRange.Address).ToString.Split(":")(1)),
                                FindWhat:=txtSearch.Text,
                                LookIn:=Excel.XlFindLookIn.xlFormulas,
                                LookAt:=Excel.XlLookAt.xlPart,
                                SearchOrder:=Excel.XlSearchOrder.xlByRows,
                                MatchCase:=False,
                                BeginsWith:=vbNullString,
                                EndsWith:=vbNullString,
                                BeginEndCompare:=vbTextCompare)
                If FoundCells Is Nothing Then
                    MessageBox.Show("Search text not found.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtSearch.Focus()
                    Exit Sub
                Else
                    Dim numRows, numColumns As Integer
                    For Each FoundCell In FoundCells
                        If FoundCell.Row > 4 Then
                            Do Until FoundCell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                                numRows = FoundCell.Rows.Count
                                numColumns = FoundCell.Columns.Count
                                FoundCell = FoundCell.Resize(numRows, numColumns + 1)
                            Loop

                            With FoundCell
                                objFC = .FormatConditions.Add(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression, , "=NOT(ISERROR(" & FoundCell.Address & "))", ,,)
                                objFC.SetFirstPriority()
                                objFC.StopIfTrue = False
                                objFC.Interior.Color = Color.White ' vbWhite
                                objFC.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Color = Color.Red 'vbRed
                                objFC.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = Color.Red
                            End With
                        End If
                    Next FoundCell
                End If

                If Form.DataCenter.GlobalValues.WS.AutoFilterMode = True Then
                    _GlobalFunctions.GetSearchFilter()
                    If chkDoFilter.Checked = True Then sbDoFilter(FoundCells.Address)
                End If
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Me.Close()
            Else
                MessageBox.Show("Please enter a text to search.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSearch.Focus()
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Exit Sub
            End If
            txtSearch.Focus()
        Catch ex As Exception
            Me.DialogResult = DialogResult.Cancel
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmSearch, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.ScreenUpdating = True
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

        Form.DataCenter.GlobalValues.WS.Cells.FormatConditions.Delete()
        Form.DataCenter.GlobalValues.WS.Range("5:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).EntireRow.Hidden = False

        If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then
            Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)
        Else
            Form.DataCenter.GlobalValues.WS.AutoFilter.ApplyFilter()
        End If

        _GlobalFunctions.ReApplyFilter()
        _obj.sbProtectPlan()
        txtSearch.Focus()
    End Sub

    Private Sub frmSearch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form.DataCenter.ProgramConfig.ISSearchActive = True
        If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)
        txtSearch.Focus()
        Show()
    End Sub

    Private Sub frmSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            btnSearch_Click(sender, e)
        ElseIf e.KeyCode = Keys.F8 Then
            btnReset_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            txtSearch.Focus()
        ElseIf e.KeyCode = Keys.Alt AndAlso Keys.T Then
            txtSearch.Focus()
        End If
    End Sub

    Private Sub frmSearch_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form.DataCenter.ProgramConfig.ISSearchActive = False
    End Sub

    Function FindAll(SearchRange As Excel.Range,
                FindWhat As String,
               Optional LookIn As Excel.XlFindLookIn = Excel.XlFindLookIn.xlValues,
                Optional LookAt As Excel.XlLookAt = Excel.XlLookAt.xlWhole,
                Optional SearchOrder As Excel.XlSearchOrder = Excel.XlSearchOrder.xlByRows,
                Optional MatchCase As Boolean = False,
                Optional BeginsWith As String = vbNullString,
                Optional EndsWith As String = vbNullString,
                Optional BeginEndCompare As CompareMethod = vbTextCompare) As Excel.Range
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' FindAll
        ' This searches the range specified by SearchRange and returns a Range object
        ' that contains all the cells in which FindWhat was found. The search parameters to
        ' this function have the same meaning and effect as they do with the
        ' Range.Find method. If the value was not found, the function return Nothing. If
        ' BeginsWith is not an empty string, only those cells that begin with BeginWith
        ' are included in the result. If EndsWith is not an empty string, only those cells
        ' that end with EndsWith are included in the result. Note that if a cell contains
        ' a single word that matches either BeginsWith or EndsWith, it is included in the
        ' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
        ' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
        ' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
        ' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
        ' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
        ' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
        ' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim FoundCell As Excel.Range
        Dim FirstFound As Excel.Range
        Dim LastCell As Excel.Range
        Dim ResultRange As Excel.Range
        Dim XLookAt As Excel.XlLookAt
        Dim Include As Boolean
        Dim CompMode As CompareMethod
        Dim Area As Excel.Range
        Dim MaxRow As Long
        Dim MaxCol As Long

        FindAll = Nothing
        ResultRange = Nothing

        CompMode = BeginEndCompare
        If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
            XLookAt = Microsoft.Office.Interop.Excel.XlLookAt.xlPart
        Else
            XLookAt = LookAt
        End If

        ' this loop in Areas is to find the last cell
        ' of all the areas. That is, the cell whose row
        ' and column are greater than or equal to any cell
        ' in any Area.

        For Each Area In SearchRange.Areas
            With Area
                If .Cells(.Cells.Count).Row > MaxRow Then
                    MaxRow = .Cells(.Cells.Count).Row
                End If
                If .Cells(.Cells.Count).Column > MaxCol Then
                    MaxCol = .Cells(.Cells.Count).Column
                End If
            End With
        Next Area
        LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)

        Try
            '        After:=LastCell,
            FoundCell = SearchRange.Find(What:=FindWhat,
            LookIn:=LookIn,
            LookAt:=XLookAt,
            SearchOrder:=SearchOrder,
            MatchCase:=MatchCase)

            If Not FoundCell Is Nothing Then
                FirstFound = FoundCell
                ResultRange = Nothing '
                Do Until False ' Loop forever. We'll "Exit Do" when necessary.
                    Include = False
                    If BeginsWith = vbNullString And EndsWith = vbNullString Then
                        Include = True
                    Else
                        If BeginsWith <> vbNullString Then
                            If StrComp(Microsoft.VisualBasic.Left(FoundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                                Include = True
                            End If
                        End If
                        If EndsWith <> vbNullString Then
                            If StrComp(Microsoft.VisualBasic.Right(FoundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                                Include = True
                            End If
                        End If
                    End If
                    '
                    If Include = True Then
                        If ResultRange Is Nothing Then
                            ResultRange = FoundCell
                        Else
                            ResultRange = Form.DataCenter.GlobalValues.WS.Application.Union(ResultRange, FoundCell)
                        End If
                    End If
                    FoundCell = SearchRange.FindNext(After:=FoundCell)
                    If (FoundCell Is Nothing) Then
                        Exit Do
                    End If
                    If (FoundCell.Address = FirstFound.Address) Then
                        Exit Do
                    End If

                Loop
            End If

            If Not (ResultRange Is Nothing) Then FindAll = ResultRange
        Catch ex As Exception
            FindAll = Nothing
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmSearch, ex.Message))
        End Try
    End Function

    'Purpose: To close the form
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class
