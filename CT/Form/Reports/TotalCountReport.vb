Public Class TotalCountReport
    Dim WB As Excel.Workbook = Nothing
    Dim WS1 As Excel.Worksheet = Nothing
    Dim WS2 As Excel.Worksheet = Nothing
    Dim _arrayDT1(,) As String
    Dim _arrayDT2(,) As String
    Dim _arrayDT3(,) As String
    Dim _arrayDT4(,) As String
    Dim _arrayDT5(,) As String
    Public Sub WriteTotCntReport()
        Try
            Dim DS As New CT.Data.VehiclePlan.Report.TotalTestDays
            Dim strHCIDName As String

            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            strHCIDName = Form.DataCenter.ProgramConfig.HCID '& " - " & Form.DataCenter.ProgramConfig.HCIDName
            _arrayDT1 = DS.TotalCountReport1(Val(Form.DataCenter.ProgramConfig.HCID), Form.DataCenter.ProgramConfig.BuildType)
            _arrayDT2 = DS.TotalCountReport2(Val(Form.DataCenter.ProgramConfig.HCID), Form.DataCenter.ProgramConfig.BuildType)
            _arrayDT3 = DS.TotalCountReport3(Val(Form.DataCenter.ProgramConfig.HCID), Form.DataCenter.ProgramConfig.BuildType)
            _arrayDT4 = DS.TotalCountReport4(Val(Form.DataCenter.ProgramConfig.HCID), Form.DataCenter.ProgramConfig.BuildType)
            _arrayDT5 = DS.TotalCountReport5(Val(Form.DataCenter.ProgramConfig.HCID), Form.DataCenter.ProgramConfig.BuildType)
            WB = Globals.ThisAddIn.Application.Workbooks.Add
            Globals.ThisAddIn.Application.DisplayDocumentActionTaskPane = False
            WB.Worksheets(1).Name = "HCID-" & strHCIDName & "-" & "report 1"
            WB.Worksheets.Add(Type.Missing, WB.Worksheets(1))
            WB.Worksheets(2).Name = "HCID-" & strHCIDName & "-" & "reports 2 to 5"
            WS1 = WB.Worksheets(1) : WS2 = WB.Worksheets(2)
            WriteReport_1()
            WriteReport_2()
            WB.Activate()
            WS1.Activate()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Total count report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.CopyObjectsWithCells = False
        End Try
    End Sub
    Private Sub WriteReport_1()
        Try
            Dim rng As Excel.Range = WS1.Range("A1").Resize(_arrayDT1.GetUpperBound(0) + 1, _arrayDT1.GetUpperBound(1) + 1)
            With rng
                .Value2 = _arrayDT1
                .Columns.AutoFit()
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .MergeCells = False
            End With
            WS1.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes).Name = "Table_CountRep1"
            WS1.ListObjects("Table_CountRep1").TableStyle = "TableStyleMedium4"
            rng.EntireRow.RowHeight = 20
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Total count report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub WriteReport_2()
        Try
            Dim rng As Excel.Range = WS2.Range("A1").Resize(_arrayDT2.GetUpperBound(0) + 1, _arrayDT2.GetUpperBound(1) + 1)

            With rng
                .Value2 = _arrayDT2
                .Columns.AutoFit()
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .MergeCells = False
            End With
            WS2.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes).Name = "Table_CountRep2"
            WS2.ListObjects("Table_CountRep2").TableStyle = "TableStyleMedium4"
            rng.EntireRow.RowHeight = 20

            rng = WS2.Range(WS2.Cells(WS2.UsedRange.Rows.Count + 2, 1).address).Resize(_arrayDT3.GetUpperBound(0) + 1, _arrayDT3.GetUpperBound(1) + 1)
            With rng
                .Value2 = _arrayDT3
                .Columns.AutoFit()
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .MergeCells = False
            End With
            WS2.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes).Name = "Table_CountRep3"
            WS2.ListObjects("Table_CountRep3").TableStyle = "TableStyleMedium4"
            rng.EntireRow.RowHeight = 20

            rng = WS2.Range(WS2.Cells(WS2.UsedRange.Rows.Count + 2, 1).address).Resize(_arrayDT4.GetUpperBound(0) + 1, _arrayDT4.GetUpperBound(1) + 1)
            With rng
                .Value2 = _arrayDT4
                .Columns.AutoFit()
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .MergeCells = False
            End With
            WS2.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes).Name = "Table_CountRep4"
            WS2.ListObjects("Table_CountRep4").TableStyle = "TableStyleMedium4"
            rng.EntireRow.RowHeight = 20

            rng = WS2.Range(WS2.Cells(WS2.UsedRange.Rows.Count + 2, 1).address).Resize(_arrayDT5.GetUpperBound(0) + 1, _arrayDT5.GetUpperBound(1) + 1)
            With rng
                .Value2 = _arrayDT5
                .Columns.AutoFit()
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .MergeCells = False
            End With
            WS2.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes).Name = "Table_CountRep5"
            WS2.ListObjects("Table_CountRep5").TableStyle = "TableStyleMedium4"
            rng.EntireRow.RowHeight = 20
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Total count report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub
End Class
