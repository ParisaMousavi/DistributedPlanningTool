Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Diagnostics

Namespace Form.Reports
    Public Class ExporToExcelReport

        Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions

        Private _tbAnswer As System.Data.DataTable = Nothing
        Private _arrayDT As String(,) = Nothing

        Private ReadOnly ActiveWindow As Object

        Dim WS = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(2)
        Dim shtGenChangeLog As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets("ChangeLogs")
        Dim Fcol = Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn
        Dim Lcol = Form.DataCenter.GlobalSections.TimeLineSectionLastColumn
        Dim rng1 = Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn
        Dim rng2 = Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn
        Dim rng3 = Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn
        Dim rng4 = Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn
        Dim rng5 = Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn
        Dim rng6 = Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn
        Dim rng7 = Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn
        Dim rng8 = Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn
        Dim rng9 = Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn
        Dim rng10 = Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn
        Dim rng11 = Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn
        Dim rng12 = Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn
        Dim rng13 = Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn
        Dim rng14 = Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn
        Dim _pe01 As Long = 0
        Dim _HCID As Integer = 0





        Public Sub New(pe01 As Long, HCID As Integer)
            MyClass._pe01 = pe01
            MyClass._HCID = HCID
        End Sub


        Public Sub exporttoexcel(bolTndplan As Boolean, bolChangelogs As Boolean, bolDvpteam As Boolean)
            'Dim remarks As CT.Data.VehiclePlan.Plan
            Dim _PlanInterface As Data.Interfaces.PlanInterface

            Dim BuildType As string = Form.DataCenter.ProgramConfig.BuildType
            If BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Sub
            End If

            Dim remarksdt As System.Data.DataTable
            Dim strRemarks As String
            Dim strProcessStepLocation As String
            Dim result As DataRow()
            Dim strPlansheetname As String = Form.DataCenter.WorkSheet.TnDPlan.ToString()
            Try
                'remarks = New Data.VehiclePlan.Plan()
                remarksdt = _PlanInterface.GetPlanRemarks(_pe01, _HCID, BuildType)
                If remarksdt Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Globals.ThisAddIn.Application.CopyObjectsWithCells = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.CutCopyMode = False
                Form.DataCenter.GlobalValues.bolCopy = False
                Form.DataCenter.GlobalValues.bolCut = False
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait

                ''
                If bolChangelogs = True Then
                    Dim obj As New AddInUtilities
                    obj.RefreshChangeLog()
                    Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString()).Activate
                End If

                Dim Wb As Workbook
                If bolTndplan = True Then
                    Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString()).Activate
                    Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.TnDPlan.ToString()).Cells(3, 4).Activate 'D3

                    Dim shp As Microsoft.Office.Interop.Excel.Shape = Nothing

                    Dim WS = Globals.Factory.GetVstoObject(CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet))
                    WS.Activate()

                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    _GlobalFunctions.GetSearchFilter()
                    Form.DataCenter.GlobalValues.WS.AutoFilterMode = False


                    Wb = Globals.ThisAddIn.Application.Workbooks.Add
                    Globals.ThisAddIn.Application.DisplayDocumentActionTaskPane = False
                    Wb.Worksheets(1).Name = "T&D plan report"
                    WS.Cells.Copy(Wb.Worksheets(1).Cells(1, 1))

                    WS.Activate()
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then WS.Range("4:" & WS.UsedRange.Rows.Count).AutoFilter(Field:=1)
                    _GlobalFunctions.ReApplyFilter()

                    Wb.Application.ScreenUpdating = False
                    Wb.Activate()

                    Wb.Worksheets(1).Activate()
                    Wb.Worksheets(1).Rows(3).Hidden = False
                    WS.Shapes(0).Copy()
                    Threading.Thread.Sleep(100)
                    Wb.Worksheets(1).Range("E2").PasteSpecial(XlPasteType.xlPasteAll)
                    Threading.Thread.Sleep(100)
                    WS.Shapes(1).Copy()
                    Threading.Thread.Sleep(100)
                    Wb.Worksheets(1).Range("M2").PasteSpecial(XlPasteType.xlPasteAll)
                    Threading.Thread.Sleep(100)
                    Wb.Worksheets(1).Shapes(2).left = 378.75
                    Globals.ThisAddIn.Application.ActiveWindow.Zoom = 67
                    ' ActiveWindow.Zoom = 67
                    'shtGenChangeLog.Unprotect(Form.DataCenter.ConstPwd)

                    ''


                    Wb.Activate()
                    Dim strContents(14) As String

                    With Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, Fcol), Wb.Worksheets(1).Cells(1, Lcol)).Interior
                        .ColorIndex = Constants.xlAutomatic
                        .TintAndShade = 0
                    End With
                    Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(8, Fcol), Wb.Worksheets(1).Cells(Wb.Worksheets(1).UsedRange.Rows.Count, Lcol)).NumberFormat = "@"
                    'Dim rng1 As Range, rng2 As Range
                    'Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng1), Wb.Worksheets(1).Cells(1, rng2)).EntireColumn.Hidden = False
                    If rng1 <> 0 Then
                        Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng1), Wb.Worksheets(1).Cells(1, rng14)).EntireColumn.Hidden = False
                    End If

                    Dim intCnt As Integer
                    Dim intFCol As Integer, intLCol As Integer
                    Dim rng As Range, TestArray() As String = Nothing, strTemp As String
                    Dim colSearch As Dictionary(Of String, String) = New Dictionary(Of String, String), strSearch As String
                    Dim rngFind As Excel.Range = Nothing
                    Dim m_stAddress As String = ""
                    Dim colPS As New Collection
                    rng = Nothing
                    For intCnt = 5 To Form.DataCenter.GlobalValues.TotalRow + 4 'WS.UsedRange.Rows.Count
                        Globals.ThisAddIn.Application.ScreenUpdating = False
                        strSearch = ""
                        FindFLCols(intCnt, intFCol, intLCol)
                        If intFCol > 0 Then
                            With WS
                                With .Range(.Cells(intCnt, intFCol - 1), .Cells(intCnt, intLCol))
                                    rngFind = .Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                    If rngFind IsNot Nothing Then
                                        m_stAddress = rngFind.Address
                                        Do
                                            Globals.ThisAddIn.Application.ScreenUpdating = False
                                            rng = Wb.Worksheets(1).Cells(intCnt, rngFind.Column)
                                            colPS.Add(rngFind.Formula & "~" & CDate(WS.Cells(4, rngFind.Column).value2))
                                            strTemp = Convert.ToString(rngFind.Formula)
                                            TestArray = Split(strTemp, ";")
                                            If TestArray.Length = 17 Then

                                                If TestArray(12).StartsWith("PMT-") Or TestArray(12).StartsWith("DNR-") Then
                                                    TestArray(12) = TestArray(12).Substring(4, TestArray(12).Length - 4)
                                                End If

                                                If TestArray(4) = "Gap" Then
                                                    strTemp = TestArray(4) & ";" & TestArray(5)
                                                Else
                                                    strTemp = TestArray(3) &
                                                          ";" & TestArray(4) &
                                                          ";" & TestArray(5) &
                                                          ";" & TestArray(7) &
                                                          ";" & TestArray(9) &
                                                          ";" & TestArray(12)

                                                    '----------------------------------------------------------
                                                    ' Add remarks to report

                                                    If remarksdt.Rows.Count > 0 Then
                                                        result = remarksdt.Select("pe26_SpecificVehicleUsercases_PK = " & TestArray(0).Split("""")(1))
                                                        strRemarks = result(0)(1).ToString()
                                                        strProcessStepLocation = result(0)(2).ToString()
                                                        strTemp = strTemp + ";" & strRemarks + ";" & strProcessStepLocation
                                                    End If

                                                    '----------------------------------------------------------
                                                End If


                                                rng.Merge()
                                                rng.Font.Color = Color.Black
                                                rng.WrapText = False
                                                rng.VerticalAlignment = Constants.xlCenter
                                                rng.HorizontalAlignment = Constants.xlLeft
                                                rng.Value = strTemp
                                                rng.NumberFormat = "@"
                                                strSearch += strTemp & ";"
                                            End If

                                            If intLCol < intFCol Then Exit Do
                                            rngFind = .FindNext(rngFind)
                                        Loop While Not rngFind Is Nothing And rngFind.Address <> m_stAddress
                                    End If
                                End With
                            End With
                            colSearch.Add(intCnt, strSearch)
                        End If
                        rng = Nothing
                        Next
                        If rng1 <> 0 Then
                            strContents(1) = rng1
                            strContents(2) = rng2
                            strContents(3) = rng3
                            strContents(4) = rng4
                            strContents(5) = rng5
                            strContents(6) = rng6
                            strContents(7) = rng7
                            strContents(8) = rng8
                            strContents(9) = rng9
                            strContents(10) = rng10
                            strContents(11) = rng11
                            strContents(12) = rng12
                            strContents(13) = rng13
                            strContents(14) = rng14
                            Wb.Worksheets(1).Cells(1, Val(strContents(1))).EntireColumn.Hidden = False
                            Wb.Worksheets(1).Cells(1, Val(strContents(14))).EntireColumn.Hidden = False
                            'Wb.Worksheets(1).Rows(3).Hidden = False
                            Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, Val(strContents(1))), Wb.Worksheets(1).Cells(1, Val(strContents(14)))).EntireColumn.Group
                            'Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, strContents(1) + 1), Wb.Worksheets(1).Cells(1, strContents(14) - 1)).EntireColumn.Group

                            For intCnt = 1 To 14 Step 2
                                Globals.ThisAddIn.Application.ScreenUpdating = False
                                Dim rng15, rng16 As Integer
                                With WS
                                    rng15 = Nothing
                                    rng16 = Nothing
                                    rng15 = Val(strContents(intCnt))
                                    rng16 = Val(strContents(intCnt + 1))
                                    Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng15 + 1), Wb.Worksheets(1).Cells(1, rng16 - 1)).EntireColumn.Group
                                End With
                            Next
                        End If

                        rng2 = Fcol
                        Wb.Worksheets(1).Cells(1, rng2).EntireColumn.Insert
                        Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng2), Wb.Worksheets(1).Cells(7, rng2)).UnMerge
                        Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng2), Wb.Worksheets(1).Cells(4, rng2)).Merge
                        'Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng2), Wb.Worksheets(1).Cells(Form.DataCenter.GlobalValues.TotalRow, rng2)).Interior.Color = Color.Green
                        Wb.Worksheets(1).Range(Wb.Worksheets(1).Cells(1, rng2), Wb.Worksheets(1).Cells(1, rng2)).EntireColumn.Interior.Color = Color.Green
                        Wb.Worksheets(1).Cells(1, rng2) = "Search"
                        Wb.Worksheets(1).Cells(1, rng2).VerticalAlignment = Constants.xlBottom
                        Wb.Worksheets(1).Cells(1, rng2).HorizontalAlignment = Constants.xlLeft
                        Wb.Worksheets(1).Cells(1, rng2).Orientation = 90
                        Wb.Worksheets(1).Cells(1, rng2).ReadingOrder = Constants.xlContext

                        For intCnt = 5 To WS.UsedRange.Rows.Count
                            Globals.ThisAddIn.Application.ScreenUpdating = False
                            'If Globals.ThisAddIn.Application.DoEvents.Cells(intCnt, "B") = "" Then Exit For
                            If colSearch.ContainsKey(intCnt.ToString()) Then
                                Wb.Worksheets(1).Cells(intCnt, rng2) = colSearch(intCnt.ToString())
                                Wb.Worksheets(1).Cells(intCnt, rng2).Interior.Color = Color.Green
                                Wb.Worksheets(1).Cells(intCnt, rng2).Font.Color = Color.Green
                                Wb.Worksheets(1).Cells(intCnt, rng2).WrapText = True
                                Wb.Worksheets(1).Cells(intCnt, rng2).EntireRow.RowHeight = 20
                            End If
                        Next

                        'Form.DataCenter.GlobalValues.WS.AutoFilterMode = False
                        'Form.DataCenter.GlobalValues.WS.Range("4:" & Wb.Worksheets(1).UsedRange.Rows.Count).AutoFilter()

                        Wb.Worksheets(1).Cells(2, rng2).EntireRow.RowHeight = 75
                        Wb.Worksheets(1).Cells(3, rng2).EntireRow.RowHeight = 23
                        Wb.Worksheets(1).Cells(4, rng2).EntireRow.RowHeight = 85
                        Wb.Worksheets(1).Shapes(1).left = Wb.Worksheets(1).Cells(2, 2).left
                        Wb.Worksheets(1).Activate()
                        'Wb.Worksheets(1).Range("A1").Select
                        Wb.Worksheets(1).Range("4:" & Wb.Worksheets(1).UsedRange.Rows.Count).AutoFilter
                        Wb.Worksheets(1).Outline.ShowLevels(RowLevels:=0, ColumnLevels:=1)

                        'Dim astrLinks As Object
                        'Dim intCnt3 As Integer
                        'astrLinks = Wb.LinkSources(Type:=XlLinkType.xlLinkTypeExcelLinks)

                        'For intCnt3 = LBound(astrLinks) To UBound(astrLinks)
                        '    Wb.Activate()
                        '    Wb.BreakLink(Name:=astrLinks(intCnt3), Type:=XlLinkType.xlLinkTypeExcelLinks)
                        'Next

                        'For Each astrLinks In Wb.Names
                        '    astrLinks.Delete
                        'Next

                        Wb.Worksheets(1).Cells.Validation.Delete
                    End If

                    If bolChangelogs = True Then
                    'Dim obj As New AddInUtilities
                    'obj.RefreshChangeLog()
                    If Wb Is Nothing Then Wb = Globals.ThisAddIn.Application.Workbooks.Add
                    shtGenChangeLog.Copy(Wb.Worksheets(1))
                    Wb.Worksheets(1).Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                End If

                If bolDvpteam = True Then
                    If Wb Is Nothing Then Wb = Globals.ThisAddIn.Application.Workbooks.Add
                    Wb.Worksheets.Add(Wb.Worksheets(1))
                    Wb.Worksheets(1).Name = "DVP Team & CDSIDs"
                    _tbAnswer = _PlanInterface.GetAssignedCDSIDs(_pe01, _HCID, BuildType)
                    If Not _tbAnswer Is Nothing Then
                        ConvertDataTableToStingArray()
                        Wb.Worksheets(1).Cells(1, 1).Value2 = "Col Header"
                        Wb.Worksheets(1).Cells(1, 2).Value2 = "Col Header"
                        Dim top As Excel.Range = Wb.Worksheets(1).Cells(2, 1)
                        Dim bottom As Excel.Range = Wb.Worksheets(1).Cells(_arrayDT.GetUpperBound(0) + 1, _arrayDT.GetUpperBound(1))
                        Dim all As Excel.Range
                        all = Wb.Worksheets(1).Range(top, bottom)
                        all.Value2 = _arrayDT
                        all.Columns.AutoFit()
                    End If
                End If

                If bolTndplan = True Then
                    Wb.Worksheets("T&D plan report").Activate
                Else
                    Wb.Worksheets("Sheet1").Delete()
                End If


                'Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                '_RibbonUtilities.DeactiveRibbonButtonsState()

                Wb.Activate()

                'Wb.SaveAs("C:\temp\Exported_TnD", FileFormat:=51)
                'Wb.Application.WindowState = XlWindowState.xlMaximized
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
                'Globals.ThisAddIn.Application.Goto(Reference:=Wb.Worksheets(1).Range("B8"), Scroll:=True)
                Try
                    Dim SHP1 As Excel.Shape, SHP2 As Excel.Shape
                    SHP1 = Wb.Worksheets("T&D plan report").shapes("Picture 1")
                    SHP2 = Wb.Worksheets("T&D plan report").shapes("Picture 2")
                    SHP2.Left = SHP1.Left + SHP1.Width + 30
                Catch ex As Exception
                End Try
                Globals.ThisAddIn.Application.CutCopyMode = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.ScreenUpdating = True
                MessageBox.Show("Completed.", "Export to Excel report", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "Export to Excel report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.ExporToExcelReport, ex.Message), "Export to Excel report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
            Finally
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
                Globals.ThisAddIn.Application.CutCopyMode = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.ScreenUpdating = True
            End Try
        End Sub

        'Public Function IsDarkColor(lngRGB As Long) As Boolean
        '    Dim HexValue As String
        '    HexValue = Hex$(lngRGB)
        '    If ((HexToDec(Left(HexValue, 2)) + HexToDec(Mid(HexValue, 3, 2)) + HexToDec(Right(HexValue, 2))) / 3) > 128 Then
        '        IsDarkColor = False
        '    Else
        '        IsDarkColor = True
        '    End If
        'End Function

        'Public Function HexToDec(ByVal xHex As String) As Integer
        '    Dim Num1 As Integer, Num2 As Integer
        '    If Len(xHex) = 1 Then xHex = "0" & xHex
        '    Num1 = nVal(Left(xHex, 1))
        '    Num2 = nVal(Right(xHex, 1))
        '    HexToDec = (Num1 * 16) + Num2
        'End Function
        'Public Function nVal(ByVal Char1 As String) As Integer
        '    nVal = IIf(IsNumeric(Char1) = False, InStr(1, "ABCDEF", Char1) + 9, Char1)
        'End Function
        ''' <summary>
        ''' The method converts the DataTable which is returned from Database to string Array 
        ''' because only array string can be written in a Excel Range.
        ''' </summary>
        Private Sub ConvertDataTableToStingArray()


            'aaa aaaa
            Dim i, j As Integer
            If _tbAnswer IsNot Nothing Then

                ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count)
                For i = 0 To _tbAnswer.Rows.Count - 1
                    For j = 0 To _tbAnswer.Columns.Count - 1
                        _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                    Next j
                Next i
            End If
        End Sub

        Public Sub FindFLCols(intRow As Integer, ByRef intFCol As Integer, ByRef intLCol As Integer)
            Try
                Dim rng1 As Range, rng2 As Range
                With WS
                    rng1 = .Range(.Cells(intRow, Fcol), .Cells(intRow, Lcol)).Find("*", , XlFindLookIn.xlFormulas, XlLookAt.xlWhole, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext)
                    rng2 = .Range(.Cells(intRow, Fcol), .Cells(intRow, Lcol)).Find("*", , XlFindLookIn.xlFormulas, XlLookAt.xlWhole, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious)
                End With
                If Not rng1 Is Nothing Then
                    intFCol = rng1.Column
                Else
                    intFCol = Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn
                End If
                If Not rng2 Is Nothing Then
                    intLCol = rng2.Column - 1
                Else
                    intLCol = Form.DataCenter.GlobalSections.TimeLineSectionLastColumn
                End If
            Catch ex As Exception

            End Try

        End Sub
    End Class


End Namespace
