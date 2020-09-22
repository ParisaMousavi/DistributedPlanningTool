'Imports Microsoft.Office.Interop.Excel

Namespace Form.Reports
    Public Class UnitReport

        Dim vbNullDate As Date
        Dim ColDateLibs As Collection
        Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions

        Public Sub VehicleReport(Tndworksheet As Microsoft.Office.Tools.Excel.Worksheet, shtVehicleReport As Microsoft.Office.Tools.Excel.Worksheet, intRow As Integer)
            Try
                'Dim strInput As String
                'strInput = InputBox("Please enter the ID of unit to generate the Unit Gantt chart report.", "Unit Report", Form.DataCenter.GlobalValues.WS.Cells(Form.DataCenter.GlobalValues.WS.Application.Selection.row, 5).Value)

                'If Val(strInput) = 0 Then
                '    shtVehicleReport.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden
                '    Form.DataCenter.GlobalValues.WS.Activate()
                '    Globals.ThisAddIn.Application_SheetActivate(Form.DataCenter.GlobalValues.WS)
                '    System.Windows.Forms.MessageBox.Show("Sorry, your input Is Not valid.", "Unit Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                '    Exit Sub
                'End If


                '_GlobalFunctions.GetSearchFilter()
                'Form.DataCenter.GlobalValues.WS.AutoFilterMode = False



                'rng = Form.DataCenter.GlobalValues.WS.Range("E5:E" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count + 10).Find(strInput, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, False)

                'ws.Activate()
                'If Form.DataCenter.GlobalValues.WS.AutoFilterMode = False Then Form.DataCenter.GlobalValues.WS.Range("4:" & Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count).AutoFilter(Field:=1)
                '_GlobalFunctions.ReApplyFilter()

                Dim Wb As Excel.Workbook
                Dim ws As Excel.Worksheet
                Dim ColDatesSorted As New Collection

                ColDateLibs = New Collection

                Dim timelinearea As String

                'Dim shtVehicleReport As Excel.Worksheet = Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.VehicleReportTemplate.ToString())

                With Tndworksheet ' Form.DataCenter.GlobalValues.WS

                    Dim TimeLineSectionFirstColumn, TimeLineSectionLastColumn As Integer
                    TimeLineSectionFirstColumn = Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn
                    TimeLineSectionLastColumn = Form.DataCenter.GlobalSections.TimeLineSectionLastColumn
                    timelinearea = Form.DataCenter.GlobalSections.TimeLineSection.Address


                    Dim intCnt As Integer, intColCnt As Integer, dtStart As Date
                    'Dim _GlobalFunc As New Form.DataCenter.GlobalFunctions
                    intColCnt = 13
                    Dim dt As Date
                    'dt = .Cells(7, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn)
                    dt = .Range(_GlobalFunctions.ColumnLetter(TimeLineSectionFirstColumn + 1).ToString & 4).Value
                    dt = DateAdd(DateInterval.Day, -(Weekday(dt, FirstDayOfWeek.Monday)) + 1, dt)
                    dtStart = Strings.Format(dt, "dd-MMM-yyyy")


                    Dim strColLetter As String
                    'intval = Weekday(dt, FirstDayOfWeek.Monday) + 1
                    'dtStart = DateAdd(DateInterval.Day, intval, dt)
                    'intval = _GlobalFunc.ColumnLetter()

                    'Dim FindColumn As Excel.Range = Nothing
                    'FindColumn = Tndworksheet.Range(Form.DataCenter.GlobalSections.TimeLineSection.Address).Find(dtStart.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                    'intStartCol = FindColumn.Column
                    'intStartCol = 179

                    'Get column letter from date'***************************************
                    Dim myDataTable As New System.Data.DataTable
                    Dim planspecific As Boolean
                    If Form.DataCenter.ProgramConfig.IsGeneric = False Then ' fnbolIsPlanSpecific Then
                        myDataTable = fnGetSlotTimings_DB(intRow)
                        If myDataTable Is Nothing Then
                            System.Windows.Forms.MessageBox.Show("Data error: No data to display", "Unit Report - Data Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                            Exit Try
                        End If
                        planspecific = True
                    Else
                        'colData = fnGetSlotTimings(intRow)
                        planspecific = False
                    End If


                    Wb = Globals.ThisAddIn.Application.Workbooks.Add
                    shtVehicleReport.Visible = Excel.XlSheetVisibility.xlSheetVisible
                    shtVehicleReport.Copy(Wb.Worksheets(1))
                    shtVehicleReport.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
                    Wb.Worksheets(1).Visible = Excel.XlSheetVisibility.xlSheetVisible
                    Wb.Worksheets(1).Name = "Vehicle ID - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).Value

                    ws = Wb.Worksheets(1)
                    'For Each shp In Tndworksheet.Shapes
                    '    MsgBox(shp.TextFrame.Characters(Type.Missing, Type.Missing).Text)
                    '    ws.Shapes(0).textframe.characters(Type.Missing, Type.Missing).text = shp.TextFrame.Characters(Type.Missing, Type.Missing).Text
                    'Next
                    'ws.Shapes("TxtProgDetails1").OLEFormat.Object.Text = .Shapes("txtPlanHeader").OLEFormat.Object.Text
                    'ws.Shapes("TxtProgDetails1").TextFrame.Characters(Type.Missing, Type.Missing).Text = .Shapes("txtPlanHeader").OLEFormat.Object.Text
                    ws.Shapes(0).textframe.characters(Type.Missing, Type.Missing).text = .Shapes(0).TextFrame.Characters(Type.Missing, Type.Missing).Text
                    ws.Shapes(1).textframe.characters(Type.Missing, Type.Missing).text = "Vehicle ID - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).Value & vbLf & "Phase - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column).Value & vbLf & "Hardware Type - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value & vbLf & "Vehicle Number - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).Value & vbLf &
                                                                        "Engine - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column).Value & vbLf & "Transmission - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column).Value & vbLf & "Team Name - " & .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).Value
                    'ws.Shapes("TxtProgDetails2").OLEFormat.Object.Text = "Vehicle ID - " & .Range("B" & intRow).Value & vbLf & "Phase - " & .Range("D" & intRow).Value & vbLf & "Hardware Type - " & .Range("G" & intRow).Value & vbLf & "Vehicle Number - " & .Range("H" & intRow).Value & vbLf &
                    '                                                    "Engine - " & .Range("J" & intRow).Value & vbLf & "Transmission - " & .Range("K" & intRow).Value & vbLf & "Team Name - " & .Range("R" & intRow).Value
                    'Wb.Worksheets(2).Delete
                    'Wb.Worksheets(2).Delete
                    'Wb.Worksheets(2).Delete
                    'sbUpdateFirstLastTimingColumns
                    Dim strAddress As String
                    'For intval = intStartCol To TimeLineSectionLastColumn + 1
                    strAddress = "$M$4"
                    Do 'Until
                        strColLetter = dtStart.ToString("yyyy-MM-dd") ' .Range(_GlobalFunc.ColumnLetter(intval).ToString & 4).Value
                        ws.Cells(2, intColCnt).value = "'" & DateValue(strColLetter).ToString("dd-MM-yyyy").ToString() ' strColLetter
                        ws.Cells(3, intColCnt).value = DatePart(DateInterval.WeekOfYear, dtStart) 'To Do Match with VBA week number 'Strings.Format(dtStart, "ww") ', vbMonday, vbFirstFourDays)
                        ws.Cells(4, intColCnt).value = "'" & DateValue(strColLetter).ToString("dd-MM-yyyy").ToString()
                        'ColDateLibs.Add(intColCnt, Strings.Format(.Range(_GlobalFunc.ColumnLetter(intval).ToString & 4).Value, "dd-MMM-yyyy"))
                        intColCnt = intColCnt + 1
                        dtStart = DateAdd(DateInterval.Day, 1, dtStart)
                    Loop While dtStart <= .Range(_GlobalFunctions.ColumnLetter(TimeLineSectionLastColumn - 1).ToString & 4).Value
                    strAddress = strAddress & ":$" & _GlobalFunctions.ColumnLetter(intColCnt - 1).ToString & "$4"
                    'Next

                    Dim strTemp() As String, rngFind As Excel.Range
                    Dim intRowCnt As Integer
                    'Dim colData As Collection, clsDat As clsSlotTimings



                    intRowCnt = 7
                    Dim int_dtStartCol, int_dtEndCol As Integer
                    'For intCnt = 1 To colData.Count
                    For intCnt = 0 To myDataTable.Rows.Count - 1
                        'clsDat = colData.Item(intCnt)
                        strTemp = Strings.Split(myDataTable.Rows(intCnt).Item("SlotUniqueName").ToString, ";")
                        ws.Cells(intRowCnt, "A").Value = strTemp(0) & ";" & strTemp(1) & ";" & strTemp(2) & ";" & strTemp(3) & ";" & strTemp(4) & ";" & strTemp(5) & ";" & strTemp(6)
                        ws.Cells(intRowCnt, "D").Value = "'" & DateValue(myDataTable.Rows(intCnt).Item("dtStart").ToString).ToString("dd-MM-yyyy").ToString() 'clsDat.dtStart
                        ws.Cells(intRowCnt, "E").Value = "'" & DateValue(myDataTable.Rows(intCnt).Item("dtEnd").ToString).ToString("dd-MM-yyyy").ToString() 'clsDat.dtEnd
                        ws.Cells(intRowCnt, "G").Value = myDataTable.Rows(intCnt).Item("Duration").ToString 'ws.Application.WorksheetFunction.NetworkDays(DateValue(myDataTable.Rows(intCnt).Item("dtStart").ToString), DateValue(myDataTable.Rows(intCnt).Item("dtEnd").ToString))
                        'ws.Cells(intRowCnt, "G").Value = Application.WorksheetFunction.NetworkDays(clsDat.dtStart, clsDat.dtEnd)
                        If planspecific = True Then ' fnbolIsPlanSpecific Then
                            'Dim objDat As clsSpecificVehicleUsercases
                            'objDat = Nothing
                            'objDat = colSpecificDataLib(Strings.Split(Strings.Split(myDataTable.Rows(intCnt).Item("SlotUniqueName").ToString, ";")(7), "~")(1))
                            'ws.Cells(intRowCnt, "G") = objDat.int_Duration
                        End If

                        'ws.Range(ws.Cells(intRowCnt, ColDateLibs.Item(Strings.Format(myDataTable.Rows(intCnt).Item("dtStart").ToString, "dd-MMM-yyyy"))), ws.Cells(intRowCnt, ColDateLibs.Item(Strings.Format(myDataTable.Rows(intCnt).Item("dtEnd").ToString, "dd-MMM-yyyy")))).Interior.Color = myDataTable.Rows(intCnt).Item("BackColor").ToString
                        'Tndworksheet.Range(Form.DataCenter.GlobalSections.TimeLineSection.Address).Find(dtStart.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                        'ws.Range(ws.Cells(intRowCnt, Tndworksheet.Range(timelinearea).Find(DateValue(myDataTable.Rows(intCnt).Item("dtStart").ToString).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column), ws.Cells(intRowCnt, Tndworksheet.Range(timelinearea).Find(DateValue(myDataTable.Rows(intCnt).Item("dtEnd").ToString).ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column)).Interior.Color = myDataTable.Rows(intCnt).Item("BackColor").ToString

                        'int_dtStartCol = ws.Range(strAddress).Find("'" & DateValue(myDataTable.Rows(intCnt).Item("dtStart").ToString).ToString("M/d/yyyy"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column
                        int_dtStartCol = ws.Range(strAddress).Find(DateValue(myDataTable.Rows(intCnt).Item("dtStart").ToString).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column

                        int_dtEndCol = ws.Range(strAddress).Find(DateValue(myDataTable.Rows(intCnt).Item("dtEnd").ToString).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column

                        If int_dtStartCol <> 0 And int_dtEndCol <> 0 Then
                            ws.Range(ws.Cells(intRowCnt, int_dtStartCol), ws.Cells(intRowCnt, int_dtEndCol)).Interior.Color = myDataTable.Rows(intCnt).Item("BackColor").ToString
                        End If

                        'ws.Range(ws.Cells(intRowCnt, ws.Range(strAddress).Find(DateValue(myDataTable.Rows(intCnt).Item("dtStart").ToString).ToString("M/d/yyyy"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column), ws.Cells(intRowCnt, ws.Range(strAddress).Find(DateValue(myDataTable.Rows(intCnt).Item("dtEnd").ToString).ToString("M/d/yyyy"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column)).Interior.Color = myDataTable.Rows(intCnt).Item("BackColor").ToString
                        intRowCnt = intRowCnt + 2
                    Next

                    'ws.Range(ws.Cells(1, intColCnt), ws.Cells(1, ws.Columns.Count)).EntireColumn.ClearFormats()
                    ws.Range(ws.Cells(intRowCnt, 1), ws.Cells(ws.Rows.Count, 1)).EntireRow.ClearFormats()

                    ws.Range(ws.Cells(1, intColCnt), ws.Cells(1, ws.Columns.Count)).EntireColumn.Hidden = True
                    ws.Range(ws.Cells(intRowCnt, 1), ws.Cells(ws.Rows.Count, 1)).EntireRow.Hidden = True

                    intRowCnt = 7

                    Dim intVeh As Integer

                    If .Range("G" & intRow).Value = "Vehicle" Then 'task id 14
                        intVeh = 3
                    Else
                        intVeh = 2
                    End If

                    For intCnt = 1 To intVeh
                        If intCnt = 1 Then
                            ws.Range("A" & intRowCnt).Value = "Build;" & ws.Range("A" & intRowCnt).Value
                        ElseIf intCnt = 2 Then
                            ws.Range("A" & intRowCnt).Value = "Sign-off;" & ws.Range("A" & intRowCnt).Value
                        ElseIf intCnt = 3 Then
                            ws.Range("A" & intRowCnt).Value = "Fit 4 Test;" & ws.Range("A" & intRowCnt).Value
                        End If
                        intRowCnt = intRowCnt + 2
                    Next

                    'If colContains(ColDateLibs, Strings.Format(Of Date, "dd-MMM-yyyy")) Then
                    Dim col As Integer = 0

                    rngFind = ws.Range(strAddress).Find(DateValue(Date.Now).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                    If Not rngFind Is Nothing Then
                        col = rngFind.Column
                    End If

                    If col > 0 Then
                        ws.Shapes(2).left = ws.Cells(1, ws.Range(strAddress).Find(DateValue(Date.Now).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column).left + (ws.Cells(1, ws.Range(strAddress).Find(DateValue(Date.Now).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column).Width / 2) - (ws.Shapes(2).Width / 2) + 1
                        ws.Shapes(3).left = ws.Cells(1, ws.Range(strAddress).Find(DateValue(Date.Now).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column).left + (ws.Cells(1, ws.Range(strAddress).Find(DateValue(Date.Now).ToString("dd-MM-yyyy").ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing).Column).Width / 2) - (ws.Shapes(3).Width / 2) + 1
                    Else
                        ws.Shapes(2).Visible = False
                        ws.Shapes(3).Visible = False
                    End If
                End With
                Globals.ThisAddIn.Application.DisplayAlerts = True
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.EnableEvents = True
                Globals.ThisAddIn.Application.CopyObjectsWithCells = False
                Wb.Activate()
                Globals.ThisAddIn.Application_WorkbookActivate(Wb)
            Catch ex As Exception
                If InStr(ex.Message, "Match method", CompareMethod.Text) > 0 Then
                    System.Windows.Forms.MessageBox.Show("The given vehicle ID is not exist", "Unit Report", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Else
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End If
                'Finally


            End Try
        End Sub

        Public Function fnGetSlotTimings_DB(intRow As Integer) As System.Data.DataTable

            Dim _Unit As New Data.VehiclePlan.Unit
            'Dim objST As clsSlotTimings
            Dim colReturn As New Collection
            Dim intFrom As Excel.Range

            fnGetSlotTimings_DB = Nothing


            Static intSlotNumber As Integer, intRows As Integer
            If intRow <> intRows Then
                intRows = intRow
                intSlotNumber = 0
            End If

            Dim mydataTable, copyTable As New System.Data.DataTable
            Dim mydataView As System.Data.DataView
            'strSql = "EXEC Report_VehiclesUsercasesDisplayDedicated " & Strings.Split(shtTnDPlan.Cells(intRow, "A"), ";")(1)
            mydataTable = _Unit.GetVehiclesUsercasesDedicated(Strings.Split(Form.DataCenter.GlobalValues.WS.Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column).Value, ";")(3), Form.DataCenter.ProgramConfig.BuildType)
            'form.datacenter.VehicleConfig.VehiclePe45
            mydataView = New System.Data.DataView(mydataTable)

            'rst = objCon.Execute(strSql, , adCmdText)
            'strSql = "SELECT * FROM pe26_SpecificVehicleUsercases WHERE pe45_AllocatedPowerPack_FK=" & Strings.Split(Form.DataCenter.WS.Cells(intRow, "A"), ";")(1)
            'rst2 = objCon.Execute(strSql, , adCmdText)

            'Dim dblSlotBackColor As Double
            'Dim dtEnd As Date
            'Dim dtStart As Date
            'Dim strSlotUniqueName As String

            copyTable.Columns.Add("BackColor")
            copyTable.Columns.Add("dtEnd")
            copyTable.Columns.Add("dtStart")
            copyTable.Columns.Add("SlotUniqueName")
            copyTable.Columns.Add("Duration")

            Dim R As System.Data.DataRow

            If mydataTable.Rows.Count > 0 Then
                'If Not (rst.BOF And rst.EOF) Then
                'With mydataTable
                Dim i As Integer
                For i = 0 To mydataTable.Rows.Count - 1
                    R = copyTable.NewRow
                    If IsDBNull(mydataTable.Rows(i).Item("Usercase").ToString()) = False Then
                        'objST = New clsSlotTimings
                        intSlotNumber = intSlotNumber + 1
                        'rst2.Filter = adFilterNone
                        mydataView.RowFilter = Nothing
                        mydataView.RowFilter = "pe26_SpecificVehicleUsercases_PK=" & mydataTable.Rows(i).Item("pe26_SpecificVehicleUsercases_PK").ToString()
                        'rst2.Filter = "pe26_SpecificVehicleUsercases_PK=" & !pe26_SpecificVehicleUsercases_PK
                        'intFrom = colTSDict(Strings.Format(rst2!DisplayPlannedStart, "dd-MMM-yyyy"))
                        Dim dt As Date
                        If mydataTable.Rows(i).Item("PlannedStart").ToString() = "" Then Exit Function
                        dt = mydataTable.Rows(i).Item("PlannedStart").ToString()
                        'intFrom = Form.DataCenter.GlobalSections.TimeLineSection.Find(dt.ToString("dd-MMM-yyyy"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                        intFrom = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalSections.TimeLineSection.Address).Find(dt.ToString("yyyy-MM-dd"), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)


                        'Dim _GlobalFunc As New Form.DataCenter.GlobalFunctions
                        'intTo = colTSDict(Strings.Format(rst2!DisplayPlannedEnd, "dd-MMM-yyyy"))
                        'If Form.DataCenter.WS.Cells(intRow, intFrom.Column - 1) = "" And intSlotNumber <> 1 Then intSlotNumber = intSlotNumber + 1
                        If Form.DataCenter.GlobalValues.WS.Range(_GlobalFunctions.ColumnLetter(intFrom.Column - 1).ToString & intRow).Value = "" And intSlotNumber <> 1 Then intSlotNumber = intSlotNumber + 1
                        With Form.DataCenter.GlobalValues.WS
                            R("BackColor") = RGB(mydataTable.Rows(i).Item("ProcessStepBackRGB").ToString.Substring(0, 3), mydataTable.Rows(i).Item("ProcessStepBackRGB").ToString.Substring(3, 3), mydataTable.Rows(i).Item("ProcessStepBackRGB").ToString.Substring(6, 3)) ' fnConvertToRGB(rst2!ProcessStepBackRGB)
                            R("dtEnd") = mydataTable.Rows(i).Item("PlannedEnd").ToString()
                            R("dtStart") = mydataTable.Rows(i).Item("PlannedStart").ToString()
                            R("Duration") = mydataTable.Rows(i).Item("Duration").ToString()
                            R("SlotUniqueName") = mydataTable.Rows(i).Item("DvpTeamName").ToString & ";" & mydataTable.Rows(i).Item("Usercase").ToString & ";" &
                                mydataTable.Rows(i).Item("ProcessStepName").ToString & ";" & _GlobalFunctions.CalculateDuration(mydataTable.Rows(i).Item("PlannedStart").ToString, mydataTable.Rows(i).Item("PlannedEnd").ToString, mydataTable.Rows(i).Item("WorkingDays").ToString) & ";" &
                                mydataTable.Rows(i).Item("FacilityLocation").ToString & ";" & _GlobalFunctions.RemoveSPChars(mydataTable.Rows(i).Item("CDSID").ToString) & ";" & _GlobalFunctions.RemoveSPChars(mydataTable.Rows(i).Item("Remarks").ToString) & ";" & intSlotNumber & "~" &
                                mydataTable.Rows(i).Item("pe26_SpecificVehicleUsercases_PK").ToString & "~" & IIf(mydataTable.Rows(i).Item("DvpTeamDisplay").ToString = "", "N", mydataTable.Rows(i).Item("DvpTeamDisplay").ToString)
                            copyTable.Rows.Add(R)
                            'objST.dblSlotBackColor = fnConvertToRGB(rst2!ProcessStepBackRGB)
                            'objST.dtEnd = rst2!DisplayPlannedEnd
                            'objST.dtStart = rst2!DisplayPlannedStart
                            'objST.strSlotUniqueName = fnCheckIsNull(rst2!DvpTeamName) & ";" & fnCheckIsNull(rst2!Usercase) & ";" &
                            '    fnCheckIsNull(rst2!ProcessStepName) & ";" & fnGetDuration(rst2!DisplayPlannedStart, rst2!DisplayPlannedEnd, fnCheckIsNull(rst2!WorkingDays)) & ";" &
                            '    fnCheckIsNull(rst2!FacilityLocation) & ";" & fnRemoveSPChars(fnCheckIsNull(rst2!CDSID)) & ";" & fnRemoveSPChars(fnCheckIsNull(rst2!Remarks)) & ";" & intSlotNumber & "~" &
                            '    rst!pe26_SpecificVehicleUsercases_PK & "~" & IIf(fnCheckIsNull(rst2!DvpTeamDisplay) = "", "N", fnCheckIsNull(rst2!DvpTeamDisplay))
                        End With
                        'colReturn.Add objST
                    End If
                    ' .MoveNext
                    'If fnbolIsValidRecordset(rst) = False Then Exit Do
                Next i
                'End With
                'End If
            End If
            fnGetSlotTimings_DB = copyTable
        End Function

    End Class
End Namespace
