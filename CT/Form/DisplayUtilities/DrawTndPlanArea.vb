
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports System
Imports System.ComponentModel
Imports System.Reflection

Namespace Form.DisplayUtilities


    Public Class DrawTndPlanArea

        Event EventUpdateProgress(progressvalue As Double)

        Private Sub UpdateProgressbar(progressvalue As Double)
            RaiseEvent EventUpdateProgress(progressvalue)
        End Sub

        'Event EventUpdateProgress()
        'Public intPer As Double = 0

        'Private Sub UpdateProgressbar()
        '    RaiseEvent EventUpdateProgress()
        'End Sub

        Public Function LoadTndPlanAreaToWorkSheet(UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, Optional intStRow As Integer = 0, Optional intEndRow As Integer = 0) As String
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            'Dim _TndPlanArea As CT.Data.VehiclePlan.Segment.TestsArea = New Data.VehiclePlan.Segment.TestsArea()
            Dim _TndPlanAreaInterface As Data.Interfaces.TestAreaInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _TndPlanAreaInterface = New Data.VehiclePlan.Segment.TestsArea
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _TndPlanAreaInterface = New Data.RigPlan.Segment.TestsArea
            Else
                Exit Function
            End If


            Dim _TndPlanAreaArray As String(,) = Nothing

            '-------------------------------------------------------
            ' This function uses String to return error 
            LoadTndPlanAreaToWorkSheet = String.Empty
            '-------------------------------------------------------


            Try

                '-----------------------------------------------------------------------------
                'Consider Generic & Specific plan
                If Form.DataCenter.ProgramConfig.IsGeneric = True Then

                    _TndPlanAreaArray = _TndPlanAreaInterface.GetTndAreaDataGeneric(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, UpperBoundDisplaySeq, LowerBoundDisplaySeq)

                ElseIf Form.DataCenter.ProgramConfig.IsGeneric = False Then

                    _TndPlanAreaArray = _TndPlanAreaInterface.GetTndAreaDataSpecific(Form.DataCenter.ProgramConfig.HCID, UpperBoundDisplaySeq, LowerBoundDisplaySeq, Form.DataCenter.ProgramConfig.BuildType)

                End If
                '-----------------------------------------------------------------------------
                UpdateProgressbar(5)

                '-----------------------------------------------------------------------------
                'validating the return value from DB
                If _TndPlanAreaArray Is Nothing Then Throw New Exception("The return Array from DB in LoadTndPlanAreaToWorkSheet is empty." + CT.Data.DataCenter.GlobalValues.message)
                '-----------------------------------------------------------------------------

                Dim top As Excel.Range = Nothing
                Dim bottom As Excel.Range = Nothing
                Dim all As Excel.Range
                If UpperBoundDisplaySeq IsNot Nothing AndAlso LowerBoundDisplaySeq IsNot Nothing Then
                    top = Form.DataCenter.GlobalValues.WS.Cells(intStRow, DataCenter.GlobalSections.TimeLineSectionFirstColumn)
                    bottom = Form.DataCenter.GlobalValues.WS.Cells(intEndRow, _TndPlanAreaArray.GetUpperBound(1) + DataCenter.GlobalSections.TimeLineSectionFirstColumn)
                Else
                    top = Form.DataCenter.GlobalValues.WS.Cells(4 + (If(Convert.ToInt32(UpperBoundDisplaySeq) = 0, 1, Convert.ToInt32(UpperBoundDisplaySeq))), DataCenter.GlobalSections.TimeLineSectionFirstColumn)
                    bottom = Form.DataCenter.GlobalValues.WS.Cells(4 + (If(Convert.ToInt32(LowerBoundDisplaySeq) = 0, (_TndPlanAreaArray.GetUpperBound(0)), Convert.ToInt32(LowerBoundDisplaySeq))), _TndPlanAreaArray.GetUpperBound(1) + DataCenter.GlobalSections.TimeLineSectionFirstColumn)
                    Form.DataCenter.GlobalValues.TotalRow = _TndPlanAreaArray.GetUpperBound(0)
                End If

                all = DirectCast(Form.DataCenter.GlobalValues.WS.Range(top, bottom), Excel.Range)
                all.Style = Style.Styles.TnsStyleName.ProcessStepStyle.ToString()
                all.NumberFormat = "General"
                all.Value2 = _TndPlanAreaArray
                all.RowHeight = 20.0
            Catch ex As Exception
                LoadTndPlanAreaToWorkSheet = ex.Message
            End Try

        End Function
        Public Function ApplyFormattingAfterLoading(UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, Optional intStRow As Integer = 0, Optional intEndRow As Integer = 0) As String

            Dim borders As Excel.Borders

            Dim strCellInfo As String, strFirstCellInfo As String = [String].Empty
            Dim rgbs As String()

            Dim PsStart As Excel.Range = Nothing
            Dim PSend As Excel.Range = Nothing

            Dim top As Excel.Range
            Dim bottom As Excel.Range
            Dim all As Excel.Range

            ApplyFormattingAfterLoading = String.Empty
            Try

                If UpperBoundDisplaySeq IsNot Nothing AndAlso LowerBoundDisplaySeq IsNot Nothing Then
                    top = Form.DataCenter.GlobalValues.WS.Cells(intStRow, Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column)
                    bottom = Form.DataCenter.GlobalValues.WS.Cells(intEndRow, Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column + DataCenter.GlobalSections.TimeLineSection.Columns.Count)
                Else
                    top = Form.DataCenter.GlobalValues.WS.Cells(4 + (If(Convert.ToInt32(UpperBoundDisplaySeq) = 0, 1, Convert.ToInt32(UpperBoundDisplaySeq))), Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column)
                    bottom = Form.DataCenter.GlobalValues.WS.Cells(4 + (If(Convert.ToInt32(LowerBoundDisplaySeq) = 0, Form.DataCenter.GlobalValues.TotalRow, Convert.ToInt32(LowerBoundDisplaySeq))), Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column + DataCenter.GlobalSections.TimeLineSection.Columns.Count)
                End If

                Try
                    all = Form.DataCenter.GlobalValues.WS.Range(top, bottom)
                    borders = all.Borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    borders.Weight = Excel.XlBorderWeight.xlHairline
                Catch ex As Exception
                End Try

                Dim intFirstCol As Integer = Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column
                Dim intLastCol As Integer = Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column + Form.DataCenter.GlobalSections.TimeLineSection.Columns.Count
                Dim intFCol As Integer, intLCol As Integer
                Dim intCnt As Integer = 0
                Dim rngM As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim strPrevID As String = ""
                Dim intPivot As Integer, rngMatch As Microsoft.Office.Interop.Excel.Range = Form.DataCenter.GlobalValues.WS.Cells(1, 1)
                Dim intPivotPrevSt As Integer = 0, intPivotPrevEd As Integer = 0
                Dim col As New Collection
                Dim rngNext As Excel.Range, intNextUC As Integer = 0
                Dim colTemp As Collection
                Dim intStR As Integer = 0
                Dim intEdR As Integer = 0

                If UpperBoundDisplaySeq IsNot Nothing AndAlso LowerBoundDisplaySeq IsNot Nothing Then
                    intStR = intStRow
                    intEdR = intEndRow
                Else
                    intStR = 4 + (If(Convert.ToInt32(UpperBoundDisplaySeq) = 0, 1, Convert.ToInt32(UpperBoundDisplaySeq)))
                    intEdR = 4 + (If(Convert.ToInt32(LowerBoundDisplaySeq) = 0, Form.DataCenter.GlobalValues.TotalRow, Convert.ToInt32(LowerBoundDisplaySeq)))
                End If

                Try
                    Dim dblProgress As Double
                    Dim dbl_TotalProgress As Double = 0
                    dblProgress = 25 / (intEdR - 4)
                    For row As Integer = intStR To intEdR
                        dbl_TotalProgress += dblProgress
                        If dbl_TotalProgress > 3 Then 'update progress bar only for values crossing 3
                            UpdateProgressbar(dbl_TotalProgress)
                            dbl_TotalProgress = 0
                        End If
                        FindFLCols(row, intFCol, intLCol, intFirstCol, intLastCol)
                        intPivot = intFCol
                        With Form.DataCenter.GlobalValues.WS
                            'Globals.ThisAddIn.Application.ScreenUpdating = False
                            .Cells(row, intLCol + 1).numberformat = ";;;"
                            .Cells(row, intLCol + 1).value2 = "-"
                            strPrevID = ""
                            If intFCol <> intFirstCol And intLCol <> intLastCol Then
                                rngMatch = .Cells(row, intFCol)
                                col = New Collection
                                Do Until rngMatch Is Nothing
                                    'Globals.ThisAddIn.Application.ScreenUpdating = False
                                    rngMatch = Nothing
                                    Try
                                        rngMatch = .Evaluate("=INDEX(" & .Range(.Cells(row, intPivot), .Cells(row, intLCol + 1)).Address(False, False) & ",MATCH(TRUE," & .Range(.Cells(row, intPivot), .Cells(row, intLCol + 1)).Address(False, False) & "<>" & .Cells(row, intPivot).Address(False, False) & ",0))")
                                    Catch ex As Exception
                                    End Try

                                    If rngMatch IsNot Nothing Then
                                        intPivotPrevSt = intPivot
                                        intPivot = rngMatch.Column
                                        intPivotPrevEd = intPivot - 1
                                        strCellInfo = ""
                                        Try
                                            strCellInfo = .Cells(row, intPivotPrevSt).Value2.ToString
                                        Catch ex As Exception
                                        End Try

                                        .Range(.Cells(row, intPivotPrevSt), .Cells(row, intPivotPrevEd)).ClearContents()
                                        With .Range(.Cells(row, intPivotPrevSt), .Cells(row, intPivotPrevEd))
                                            'Globals.ThisAddIn.Application.ScreenUpdating = False
                                            If strPrevID.Split(";")(0) <> strCellInfo.Split(";")(0) And strCellInfo <> "" Then
                                                .Cells(1, 1).Formula = Convert.ToString("=CellFace(""") & strCellInfo + """," & .Cells(1, 1).Address & ")"
                                            ElseIf strCellInfo = "" Then
                                                .Cells(1, 1).value2 = "-"
                                                .Cells(1, 1).numberformat = ";;;"
                                            End If
                                            Try
                                                Globals.ThisAddIn.Application.ScreenUpdating = False
                                                .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                                .BorderAround2(Excel.XlLineStyle.xlContinuous, 2)
                                            Catch ex As Exception
                                            End Try
                                            If (Form.DataCenter.GlobalValues.WS.Cells(row, intPivotPrevSt - 1).interior.color = RGB(192, 192, 192) Or Form.DataCenter.GlobalValues.WS.Cells(row, intPivotPrevSt - 1).interior.color = RGB(183, 63, 183)) And strPrevID.Split(";")(0) = strCellInfo.Split(";")(0) Then
                                                .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                            End If
                                            rgbs = strCellInfo.Split(";"c)
                                            Try
                                                If Integer.Parse(rgbs.GetUpperBound(0)) > 2 Then
                                                    If ((Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(0, 3)) = 192 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(3, 3)) = 192 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(6, 3)) = 192) Or
                                                    (Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(0, 3)) = 183 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(3, 3)) = 63 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(6, 3)) = 183) AndAlso strPrevID.Split(";")(0) = strCellInfo.Split(";")(0)) Then
                                                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                                    End If
                                                    .Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(0, 3)), Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(3, 3)), Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(6, 3))))
                                                    .Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(UBound(rgbs)).Substring(0, 3)), Integer.Parse(rgbs(UBound(rgbs)).Substring(3, 3)), Integer.Parse(rgbs(UBound(rgbs)).Substring(6, 3))))
                                                End If
                                            Catch ex As Exception
                                            End Try

                                            Try
                                                If intPivot = intLCol + 1 Then
                                                    If intPivotPrevSt = intFCol Then
                                                        col.Add(Form.DataCenter.GlobalValues.WS.Cells(row, intFCol))
                                                    ElseIf intPivotPrevEd = intLCol And Form.DataCenter.GlobalValues.WS.Cells(row, intLCol).formula <> "" And Form.DataCenter.GlobalValues.WS.Cells(row, intLCol - 1).formula <> "" Then
                                                        If Form.DataCenter.GlobalValues.WS.Cells(row, intLCol).formula.ToString.Split(";")(1) <> Form.DataCenter.GlobalValues.WS.Cells(row, intLCol - 1).formula.ToString.Split(";")(1) Then
                                                            col.Add(Form.DataCenter.GlobalValues.WS.Cells(row, intLCol))
                                                        End If
                                                    End If
                                                    col.Add(Form.DataCenter.GlobalValues.WS.Cells(row, intLCol))
                                                    If col.Count > 2 Then
                                                        Do Until col.Count = 2
                                                            col.Remove(1)
                                                        Loop
                                                    End If
                                                    Form.DataCenter.GlobalValues.WS.Range(CType(col.Item(1), Excel.Range), CType(col.Item(2), Excel.Range)).BorderAround2(Excel.XlLineStyle.xlContinuous, 3)

                                                ElseIf strCellInfo <> "" Then
                                                    If col.Contains(strCellInfo.Split(";")(1)) = False Then
                                                        col.Add(.Cells(1, 1), strCellInfo.Split(";")(1))
                                                        If col.Count > 2 Then
                                                            Do Until col.Count = 2
                                                                col.Remove(1)
                                                            Loop
                                                        End If
                                                        Form.DataCenter.GlobalValues.WS.Range(CType(col.Item(1), Excel.Range), Form.DataCenter.GlobalValues.WS.Cells(row, CType(col.Item(2), Excel.Range).Column - 1)).BorderAround2(Excel.XlLineStyle.xlContinuous, 3)
                                                    End If
                                                ElseIf strCellInfo = "" Then
                                                    rngNext = Nothing
                                                    intNextUC = -1
                                                    Try
                                                        rngNext = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(row, .Cells(1, 1).column), Form.DataCenter.GlobalValues.WS.Cells(row, intLCol + 1)).Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                                        If Not rngNext Is Nothing Then
                                                            If rngNext.Value2.ToString <> "" And rngNext.Value2.ToString <> "-" Then intNextUC = Integer.Parse(rngNext.Value2.ToString.Split(";")(1))
                                                        End If
                                                    Catch ex As Exception
                                                    End Try
                                                    If intNextUC <> -1 Then
                                                        If Integer.Parse(strPrevID.ToString.Split(";")(1)) <> intNextUC Then
                                                            colTemp = New Collection
                                                            colTemp.Add(.Cells(1, 1))
                                                            col.Add(.Cells(1, 1))
                                                            If col.Count > 2 Then
                                                                Do Until col.Count = 2
                                                                    col.Remove(1)
                                                                Loop
                                                            End If
                                                            Form.DataCenter.GlobalValues.WS.Range(CType(col.Item(1), Excel.Range), Form.DataCenter.GlobalValues.WS.Cells(row, CType(col.Item(2), Excel.Range).Column - 1)).BorderAround2(Excel.XlLineStyle.xlContinuous, 3)
                                                            colTemp.Add(Form.DataCenter.GlobalValues.WS.Cells(row, rngNext.Column))
                                                            Form.DataCenter.GlobalValues.WS.Range(CType(colTemp.Item(1), Excel.Range), Form.DataCenter.GlobalValues.WS.Cells(row, CType(colTemp.Item(2), Excel.Range).Column - 1)).BorderAround2(Excel.XlLineStyle.xlContinuous, 3)
                                                        End If
                                                    End If
                                                End If
                                            Catch ex As Exception
                                            End Try
                                        End With
                                        strPrevID = strCellInfo
                                    End If
                                Loop
                                '-------------------------------------------------------------------------------
                                ' Place holder for Buck
                                '-------------------------------------------------------------------------------
                                If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                                    Globals.ThisAddIn.Application.ScreenUpdating = False

                                    If Form.DataCenter.VehicleConfig.VehicleBuildType(row) = CT.Data.DataCenter.BuildType.Buck.ToString Then
                                        '255317;0;1;Build;Build;15;-;;FOE;Testfacility;-;MEC;;Standard build location;N;141141141;000000000
                                        strCellInfo = "0;0;0;Buck;Buck;-;-;;;Buck;-;;;;N;141141141;000000000"

                                        Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(row, intPivotPrevEd), Form.DataCenter.GlobalValues.WS.Cells(row, intPivotPrevEd)).ClearContents()

                                        With Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.WS.Cells(row, intPivotPrevEd + 1), Form.DataCenter.GlobalValues.WS.Cells(row, DataCenter.GlobalSections.TimeLineSectionLastColumn - 1))
                                            Globals.ThisAddIn.Application.ScreenUpdating = False
                                            .ClearContents()

                                            rgbs = strCellInfo.Split(";"c)
                                            If rgbs.GetUpperBound(0) > 2 Then
                                                Try
                                                    If (Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(0, 3)) = 192 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(3, 3)) = 192 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(6, 3)) = 192) Or
                                                    (Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(0, 3)) = 183 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(3, 3)) = 63 AndAlso Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(6, 3)) = 183) Then
                                                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                                        .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                                    End If
                                                    .Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(0, 3)), Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(3, 3)), Integer.Parse(rgbs(UBound(rgbs) - 1).Substring(6, 3))))
                                                    .Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Integer.Parse(rgbs(UBound(rgbs)).Substring(0, 3)), Integer.Parse(rgbs(UBound(rgbs)).Substring(3, 3)), Integer.Parse(rgbs(UBound(rgbs)).Substring(6, 3))))
                                                Catch ex As Exception
                                                End Try
                                            End If


                                            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                            .BorderAround2(Excel.XlLineStyle.xlContinuous, 3)

                                            .Cells(1, 2).Formula = Convert.ToString("=CellFace(""") & strCellInfo + """," & .Cells(1, 1).Address & ")"

                                        End With
                                    End If
                                End If
                                '-------------------------------------------------------------------------------
                            End If
                        End With
                    Next
                    UpdateProgressbar(dbl_TotalProgress) 'For Balance
                Catch ex As Exception
                End Try

                col = Nothing

                Dim RngI As Integer, DtTemp As Date

                Globals.ThisAddIn.Application.FindFormat.Interior.Color = Integer.Parse(CT.My.Resources.EmptyColor) 'Form.DataCenter.vbEmptyColor
                Globals.ThisAddIn.Application.ReplaceFormat.Interior.Color = RGB(192, 192, 192)

                For RngI = Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column To Form.DataCenter.GlobalSections.TimeLineSection.Cells(1, 1).Column + Form.DataCenter.GlobalSections.TimeLineSection.Columns.Count
                    With Form.DataCenter.GlobalValues.WS
                        Try
                            DtTemp = CDate(.Cells(4, RngI).value2.ToString)
                            If DtTemp.DayOfWeek = DayOfWeek.Saturday Or DtTemp.DayOfWeek = DayOfWeek.Sunday Then
                                .Range(.Cells(3, RngI), .Cells(4, RngI)).Interior.Color = RGB(192, 192, 192)
                                .Range(.Cells(intStR, RngI), .Cells(intEdR, RngI)).Interior.Pattern = Excel.XlPattern.xlPatternGray25
                            End If
                        Catch ex As Exception
                        End Try
                    End With
                Next

                Dim tblProto As System.Data.DataTable
                Dim tblRow As System.Data.DataRow
                Dim tblRows() As System.Data.DataRow
                Dim colUnique1 As New Collection
                Dim colUnique2 As New Collection
                Dim _DataCls As New Data.VehiclePlan.Plan

                Dim strXCC_V As String = ""
                Dim strTeamName_V As String = ""
                Dim strXCC_B As String = ""
                Dim strTeamName_B As String = ""
                Dim strXCC_Rg As String = ""
                Dim strTeamName_Rg As String = ""
                Dim strXCC_R As String = ""
                Dim strTeamName_R As String = ""

                tblProto = _DataCls.GetXCCUserTeamNameTranslation(Form.DataCenter.ProgramConfig.BuildType)
                If tblProto Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                tblRows = tblProto.Select("BuildTypes='Vehicle'")

                For Each tblRow In tblRows
                    If Not colUnique1.Contains(tblRow("XCCPrototypeUser")) Then
                        strXCC_V = strXCC_V & "," & tblRow("XCCPrototypeUser")
                        colUnique1.Add(Nothing, tblRow("XCCPrototypeUser").ToString)
                    End If
                    If Not colUnique2.Contains(tblRow("XCCTranslation")) Then
                        strTeamName_V = strTeamName_V & "," & tblRow("XCCTranslation")
                        colUnique2.Add(Nothing, tblRow("XCCTranslation"))
                    End If
                Next

                colUnique1 = New Collection
                colUnique2 = New Collection

                tblRows = tblProto.Select("BuildTypes='Buck'")
                If tblRows Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                For Each tblRow In tblRows
                    If Not colUnique1.Contains(tblRow("XCCPrototypeUser")) Then
                        strXCC_B = strXCC_B & "," & tblRow("XCCPrototypeUser")
                        colUnique1.Add(Nothing, tblRow("XCCPrototypeUser").ToString)
                    End If
                    If Not colUnique2.Contains(tblRow("XCCTranslation")) Then
                        strTeamName_B = strTeamName_B & "," & tblRow("XCCTranslation")
                        colUnique2.Add(Nothing, tblRow("XCCTranslation"))
                    End If
                Next

                colUnique1 = New Collection
                colUnique2 = New Collection

                tblRows = tblProto.Select("BuildTypes='Rig'")
                If tblRows Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                For Each tblRow In tblRows
                    If Not colUnique1.Contains(tblRow("XCCPrototypeUser")) Then
                        strXCC_Rg = strXCC_Rg & "," & tblRow("XCCPrototypeUser")
                        colUnique1.Add(Nothing, tblRow("XCCPrototypeUser").ToString)
                    End If
                    If Not colUnique2.Contains(tblRow("XCCTranslation")) Then
                        strTeamName_Rg = strTeamName_Rg & "," & tblRow("XCCTranslation")
                        colUnique2.Add(Nothing, tblRow("XCCTranslation"))
                    End If
                Next

                colUnique1 = New Collection
                colUnique2 = New Collection

                tblRows = tblProto.Select("BuildTypes='Rebuild'")
                If tblRows Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                For Each tblRow In tblRows
                    If Not colUnique1.Contains(tblRow("XCCPrototypeUser")) Then
                        strXCC_R = strXCC_R & "," & tblRow("XCCPrototypeUser")
                        colUnique1.Add(Nothing, tblRow("XCCPrototypeUser").ToString)
                    End If
                    If Not colUnique2.Contains(tblRow("XCCTranslation")) Then
                        strTeamName_R = strTeamName_R & "," & tblRow("XCCTranslation")
                        colUnique2.Add(Nothing, tblRow("XCCTranslation"))
                    End If
                Next
                Dim strPaintFacility As String = ""

                Dim objData As New CT.Data.PaintFacility
                Dim _DT As System.Data.DataTable = objData.SelectAll
                Dim _DR As System.Data.DataRow = Nothing

                For Each _DR In _DT.Rows
                    strPaintFacility = strPaintFacility & "," & _DR("PaintCode").ToString
                Next

                If strPaintFacility <> "" Then strPaintFacility = Strings.Right(strPaintFacility, Strings.Len(strPaintFacility) - 1)


                If strTeamName_V <> "" Then strTeamName_V = Strings.Right(strTeamName_V, Strings.Len(strTeamName_V) - 1)
                If strXCC_V <> "" Then strXCC_V = Strings.Right(strXCC_V, Strings.Len(strXCC_V) - 1)

                If strTeamName_B <> "" Then strTeamName_B = Strings.Right(strTeamName_B, Strings.Len(strTeamName_B) - 1)
                If strXCC_B <> "" Then strXCC_B = Strings.Right(strXCC_B, Strings.Len(strXCC_B) - 1)

                If strTeamName_Rg <> "" Then strTeamName_Rg = Strings.Right(strTeamName_Rg, Strings.Len(strTeamName_Rg) - 1)
                If strXCC_Rg <> "" Then strXCC_Rg = Strings.Right(strXCC_Rg, Strings.Len(strXCC_Rg) - 1)

                If strTeamName_R <> "" Then strTeamName_R = Strings.Right(strTeamName_R, Strings.Len(strTeamName_R) - 1)
                If strXCC_R <> "" Then strXCC_R = Strings.Right(strXCC_R, Strings.Len(strXCC_R) - 1)

                Dim strList3() As String, strList4() As String
                Dim strList5() As String, strList6() As String
                Dim strList7() As String, strList8() As String
                Dim strList9() As String, strList10() As String, strList11() As String

                Dim ShtProgConfig_ As Object = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString)
                Dim shtTnDPlan As Object = Form.DataCenter.GlobalValues.WS

                ShtProgConfig_.Range("C2:L" & ShtProgConfig_.UsedRange.Rows.Count).ClearContents

                strList3 = Split(strXCC_V, ",")
                ShtProgConfig_.Range("E1").Resize(UBound(strList3) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList3)

                strList4 = Split(strTeamName_V, ",")
                ShtProgConfig_.Range("F1").Resize(UBound(strList4) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList4)

                strList5 = Split(strXCC_B, ",")
                ShtProgConfig_.Range("G1").Resize(UBound(strList5) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList5)

                strList6 = Split(strTeamName_B, ",")
                ShtProgConfig_.Range("H1").Resize(UBound(strList6) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList6)

                strList7 = Split(strXCC_Rg, ",")
                ShtProgConfig_.Range("I1").Resize(UBound(strList7) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList7)

                strList8 = Split(strTeamName_Rg, ",")
                ShtProgConfig_.Range("J1").Resize(UBound(strList8) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList8)

                strList9 = Split(strXCC_R, ",")
                ShtProgConfig_.Range("K1").Resize(UBound(strList9) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList9)

                strList10 = Split(strTeamName_R, ",")
                ShtProgConfig_.Range("L1").Resize(UBound(strList10) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList10)

                strList11 = Split(strPaintFacility, ",")
                ShtProgConfig_.Range("M1").Resize(UBound(strList11) + 1, 1).Value2 = ShtProgConfig_.Application.Transpose(strList11)

                For intCnt = intStR To intEdR
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                    Globals.ThisAddIn.Application.EnableEvents = False
                    If shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 = "Vehicle" Then
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!E1:E" & UBound(strList3) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!F1:F" & UBound(strList4) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With
                    ElseIf shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 = "Buck" Then
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!G1:G" & UBound(strList5) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!H1:H" & UBound(strList6) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With

                    ElseIf shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 = "Rig" Then
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!I1:I" & UBound(strList7) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!J1:J" & UBound(strList8) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With

                    ElseIf shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 = "Rebuild" Then
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!K1:K" & UBound(strList9) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With
                        With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Validation
                            Try
                                .Delete
                            Catch ex As Exception
                            End Try
                            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                        Formula1:="='" & ShtProgConfig_.Name & "'!L1:L" & UBound(strList10) + 1)
                            .IgnoreBlank = False
                            .InCellDropdown = False
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
                        End With
                    End If
                    With shtTnDPlan.Cells(intCnt, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).Validation
                        Try
                            .Delete
                        Catch ex As Exception
                        End Try
                        .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                    Formula1:="='" & ShtProgConfig_.Name & "'!M1:M" & UBound(strList11) + 1)
                        .IgnoreBlank = False
                        .InCellDropdown = False
                        .InputTitle = ""
                        .ErrorTitle = ""
                        .InputMessage = ""
                        .ErrorMessage = ""
                        .ShowInput = True
                        .ShowError = True
                    End With
                Next

                ' Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False

                Dim clsFunc As New Form.DataCenter.GlobalFunctions
                clsFunc.MaxLengthValidationsSections()
                'If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                '    With Form.DataCenter.GlobalValues.WS
                '        Dim rngFnd2 As Excel.Range = Nothing, rngFnd3 As Excel.Range = Nothing, rngFnd4 As Excel.Range = Nothing, rngFnd5 As Excel.Range = Nothing

                '        rngFnd2 = .Range(.Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn)).Find("Engine Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                '        rngFnd4 = .Range(.Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn)).Find("Engine Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)


                '        rngFnd3 = .Range(.Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn)).Find("Transmission Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                '        rngFnd5 = .Range(.Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn)).Find("Transmission Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                                                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                '        If rngFnd2 IsNot Nothing And rngFnd4 IsNot Nothing Then
                '            .Range(.Cells(5, rngFnd2.Column), .Cells(.UsedRange.Rows.Count, rngFnd2.Column)).Value2 = .Range(.Cells(5, rngFnd4.Column), .Cells(.UsedRange.Rows.Count, rngFnd4.Column)).Value2
                '        End If
                '        If rngFnd3 IsNot Nothing And rngFnd5 IsNot Nothing Then
                '            .Range(.Cells(5, rngFnd3.Column), .Cells(.UsedRange.Rows.Count, rngFnd3.Column)).Value2 = .Range(.Cells(5, rngFnd5.Column), .Cells(.UsedRange.Rows.Count, rngFnd5.Column)).Value2
                '        End If
                '    End With
                'End If
                Form.DataCenter.GlobalValues.WS.Application.CalculateFullRebuild()
            Catch ex As Exception
                ApplyFormattingAfterLoading = "ApplyFormattingAfterLoading: " + ex.Message
            End Try

        End Function
        Public Sub FindFLCols(intRow As Integer, ByRef intFCol As Integer, ByRef intLCol As Integer, intFirstCol As Integer, intLastCol As Integer)
            Dim rng1 As Microsoft.Office.Interop.Excel.Range, rng2 As Microsoft.Office.Interop.Excel.Range
            With Form.DataCenter.GlobalValues.WS
                rng1 = .Range(.Cells(intRow, intFirstCol), .Cells(intRow, intLastCol)).Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                rng2 = .Range(.Cells(intRow, intFirstCol), .Cells(intRow, intLastCol)).Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, False, Type.Missing, Type.Missing)
            End With
            If Not rng1 Is Nothing Then
                intFCol = rng1.Column
            Else
                intFCol = intFirstCol
            End If
            If Not rng2 Is Nothing Then
                intLCol = rng2.Column
            Else
                intLCol = intLastCol
            End If
        End Sub

    End Class
End Namespace
