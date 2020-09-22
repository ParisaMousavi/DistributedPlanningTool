Imports System.Windows.Forms
Namespace Form.DataCenter

    Public Class GlobalFunctions

        Public filterArray(,) As Object = Nothing
        Public currentFiltRange As String
        Public col As Integer
        Public bolWasOn As Boolean
        Public Function RemoveSPChars(strString As String) As String
            Dim strValue As String
            strValue = fnCheckIsNull(strString)
            strValue = Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(strValue, "&", ""), "<", ""), ">", ""), "'", ""), """", ""), ";", ""), "`", ""), "~", "")
            RemoveSPChars = strValue
        End Function

        Public Function fnCheckIsNull(rstField As String) As String
            If Not IsDBNull(rstField) Then
                fnCheckIsNull = rstField
            Else
                fnCheckIsNull = vbNullString
            End If
        End Function

        Public Function ColumnLetter(ColumnNumber As Long) As String
            Dim n As Long
            Dim c As Byte
            Dim s As String

            n = ColumnNumber
            Do
                c = ((n - 1) Mod 26)
                s = Chr(c + 65) & s
                n = (n - c) \ 26
            Loop While n > 0
            ColumnLetter = s
        End Function
        Public Sub sbToggleCutCopyAndPaste(bolAllow As Boolean)

            On Error Resume Next

            With Globals.ThisAddIn.Application
                .OnKey("%{F10}", "sbCreateReport_Engine_Trans")
                If Not bolAllow Then
                    .OnKey("^c", "CutCopyPasteDisabled")
                    .OnKey("^v", "CutCopyPasteDisabled")
                    .OnKey("^x", "CutCopyPasteDisabled")
                    .OnKey("{DELETE}", "CutCopyPasteDisabled")
                    .OnKey("{INSERT}", "CutCopyPasteDisabled")
                    .OnKey("{BACKSPACE}", "CutCopyPasteDisabled")
                    .OnKey("{ESC}", "sbCancel")
                Else
                    .OnKey("^c")
                    .OnKey("^v")
                    .OnKey("^x")
                    .OnKey("{DELETE}")
                    .OnKey("{INSERT}")
                    .OnKey("{BACKSPACE}")
                    .OnKey("{ESC}")
                End If
            End With

        End Sub

        ''' <summary>
        ''' This function checked only the following characters
        ''' ' " ;
        ''' </summary>
        ''' <param name="strValue"></param>
        ''' <returns></returns>
        Public Function ContainsInvalidChar(strValue As String) As Boolean
            ContainsInvalidChar = False
            If Strings.InStr(1, strValue, "'") > 0 Or Strings.InStr(1, strValue, """") > 0 Or Strings.InStr(1, strValue, ";") > 0 Or
                    Strings.InStr(1, strValue, ":") > 0 Or Strings.InStr(1, strValue, "&") > 0 Or Strings.InStr(1, strValue, "<") > 0 Or
                    Strings.InStr(1, strValue, ">") > 0 Or Strings.InStr(1, strValue, "~") > 0 Then
                ContainsInvalidChar = True
            End If
        End Function

        '''' <summary>
        '''' This function checked only the following characters
        '''' '
        '''' </summary>
        '''' <param name="strValue"></param>
        '''' <returns></returns>
        'Public Function ContainsInvalidCharLight(strValue As String) As Boolean
        '    ContainsInvalidCharLight = False
        '    If Strings.InStr(1, strValue, "'") > 0 Then
        '        ContainsInvalidCharLight = True
        '    End If
        'End Function

        ''' <summary>
        ''' Calculating the duration in UserForms where we need to calculate duration
        ''' </summary>
        ''' <param name="dtPlannedStart"></param>
        ''' <param name="dtPlannedEnd"></param>
        ''' <param name="intWorkDays"></param>
        ''' <remarks>+++</remarks> 
        ''' <returns></returns>
        Public Function CalculateDuration(dtPlannedStart As Date, dtPlannedEnd As Date, intWorkDays As Integer) As Integer
            If intWorkDays = 5 Then
                'fnGetDuration = Form.DataCenter.GlobalValues.WS.Application.WorksheetFunction.NetworkDays(dtPlannedStart, dtPlannedEnd)
                'fnGetDuration = Form.DataCenter.GlobalValues.wrkFunc.NetworkDays(dtPlannedStart, dtPlannedEnd)
                'fnGetDuration = Application.WorksheetFunction.NetworkDays(dtPlannedStart, dtPlannedEnd)
                CalculateDuration = Globals.ThisAddIn.Application.WorksheetFunction.NetworkDays(dtPlannedStart, dtPlannedEnd)
            ElseIf intWorkDays = 6 Then
                CalculateDuration = Globals.ThisAddIn.Application.WorksheetFunction.NetworkDays_Intl(dtPlannedStart, dtPlannedEnd, 11)
            Else
                CalculateDuration = DateDiff("d", dtPlannedStart, dtPlannedEnd, vbMonday, vbFirstFourDays) + 1
            End If
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="WS"></param>

        Function colContains(col As Collection, strKey As String) As Boolean
            Try
                col.Add(Nothing, strKey)
                col.Remove(col.Count)
                colContains = False
                Exit Function
            Catch ex As Exception
                colContains = True
                Data.DataCenter.GlobalValues.message = String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.GlobalFunction, ex.Message)
            End Try
        End Function
        Public Function UpdateSection(intStartRow As Integer, intEndRow As Integer, Optional bolApplyFilter As Boolean = True, Optional bolFromRefreshUnit As Boolean = False, Optional lngPSID As Long = 0, Optional PSStartDate As Date = Nothing) As Boolean

            If intStartRow < 5 Or intEndRow < 5 Then Exit Function
            Dim intColumn As Integer = Globals.ThisAddIn.Application.Selection.column
            Dim objGlobal As New Form.DataCenter.ModuleFunction
            Dim objCls As New Form.DataCenter.GlobalFunctions
            Dim bolScreenUpdating As Boolean, bolEnableEvents As Boolean, bolDisplayAlerts As Boolean
            Dim lngPS As Long = Form.DataCenter.ProcessStepConfig.ProcessStepPe26
            Dim vbNullDate As Date

            If lngPSID <> 0 Then
                lngPS = lngPSID
            End If

            bolScreenUpdating = Globals.ThisAddIn.Application.ScreenUpdating
            bolEnableEvents = Globals.ThisAddIn.Application.EnableEvents
            bolDisplayAlerts = Globals.ThisAddIn.Application.DisplayAlerts

            Try
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = True

                '------------------------------------------------------------
                'Return message from function to control the correctness
                Dim strMessage As String = String.Empty
                '------------------------------------------------------------

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Globals.ThisAddIn.Application.CutCopyMode = False
                Form.DataCenter.GlobalValues.bolCopy = False
                Form.DataCenter.GlobalValues.bolCut = False
                If bolApplyFilter Then objCls.GetResetFilter()


                If Form.DataCenter.GlobalValues.WS.Name <> Form.DataCenter.GlobalValues.WS.Application.ActiveWorkbook.ActiveSheet.name.ToString Then
                    Form.DataCenter.GlobalValues.WS.Activate()
                End If

                Dim intDisSeqSt As Integer = CInt(Form.DataCenter.GlobalValues.WS.Cells(intStartRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value2.ToString)
                Dim intDisSeqEnd As Integer = CInt(Form.DataCenter.GlobalValues.WS.Cells(intEndRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value2.ToString)

                With Form.DataCenter.GlobalValues.WS
                    .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    Form.DisplayUtilities.TndSection.CleanRange(.Range(.Cells(intStartRow, 1), .Cells(intEndRow, 1)).EntireRow)
                End With

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                Dim _DrawTndPlanInformation As Form.DisplayUtilities.DrawTndPlanInformation = New Form.DisplayUtilities.DrawTndPlanInformation()
                Dim _DrawTndPlanArea As Form.DisplayUtilities.DrawTndPlanArea = New Form.DisplayUtilities.DrawTndPlanArea()

                _DrawTndPlanInformation.LoadTndPlanInformationToWorkSheet(intDisSeqSt, intDisSeqEnd, intStartRow, intEndRow)

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                strMessage = _DrawTndPlanArea.LoadTndPlanAreaToWorkSheet(intDisSeqSt, intDisSeqEnd, intStartRow, intEndRow)
                If strMessage <> String.Empty Then
                    Throw New Exception(strMessage)
                End If
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                _DrawTndPlanArea.ApplyFormattingAfterLoading(intDisSeqSt, intDisSeqEnd, intStartRow, intEndRow)
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                Try
                    If vbNullDate <> PSStartDate And intStartRow = intEndRow Then
                        Dim intCnt As Integer
                        With Form.DataCenter.GlobalValues.WS
                            For intCnt = Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn + 1 To Form.DataCenter.GlobalSections.TimeLineSectionLastColumn - 1
                                If .Cells(intStartRow, intCnt).Formula.ToString <> "" And .Cells(intStartRow, intCnt).Formula.ToString <> "-" Then
                                    lngPS = CLng(.Cells(intStartRow, intCnt).Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim())
                                End If
                                If CDate(.Cells(4, intCnt).value2) = PSStartDate Then
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                Catch ex As Exception
                End Try

                Try
                    With Form.DataCenter.GlobalValues.WS
                        .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                        'Form.DataCenter.GlobalValues.WS.Range("W5:W2000").NumberFormat = "dd-MM-yyyy"
                        ' Form.DataCenter.GlobalValues.WS.Range("W5:W" & Form.DataCenter.GlobalValues.WS.UsedRange.Address.ToString().Split("$")(4)).NumberFormat = "@"
                        Form.DataCenter.GlobalValues.WS.Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).NumberFormat = "@"
                    End With
                    Globals.ThisAddIn.Application.CutCopyMode = False
                    Form.DataCenter.GlobalValues.bolCopy = False
                    Form.DataCenter.GlobalValues.bolCut = False
                Catch ex As Exception
                    'MsgBox(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.GlobalFunction, ex.Message))
                    MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.GlobalFunction, ex.Message), "Error while refreshing unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
                End Try

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Catch ex As Exception
                With Form.DataCenter.GlobalValues.WS
                    Form.DisplayUtilities.TndSection.CleanNumbersAfterRefreshSection(.Range(.Cells(intStartRow, 1), .Cells(intEndRow, 1)).EntireRow)
                End With
                If Not bolFromRefreshUnit Then
                    'If intColumn >= Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn Then
                    '    Form.DataCenter.GlobalValues.WS.Cells(3, intColumn).select
                    'Else
                    '    Form.DataCenter.GlobalValues.WS.Cells(4, intColumn).select
                    'End If
                End If
                If bolApplyFilter Then objCls.ReApplyFilter()
                objGlobal.sbProtectPlan()

                Globals.ThisAddIn.Application.CutCopyMode = False
                Form.DataCenter.GlobalValues.bolCopy = False
                Form.DataCenter.GlobalValues.bolCut = False
                Globals.ThisAddIn.Application.ScreenUpdating = bolScreenUpdating
                Globals.ThisAddIn.Application.EnableEvents = bolEnableEvents
                Globals.ThisAddIn.Application.DisplayAlerts = bolDisplayAlerts
                SelectAfterRefresh(lngPS)
                Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
                System.Windows.Forms.MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.GlobalFunction, ex.Message), "Error while refreshing unit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1)
                Return False
                Exit Function
            End Try
            With Form.DataCenter.GlobalValues.WS
                Form.DisplayUtilities.TndSection.CleanNumbersAfterRefreshSection(.Range(.Cells(intStartRow, 1), .Cells(intEndRow, 1)).EntireRow)
            End With

            If bolApplyFilter Then objCls.ReApplyFilter()
            objGlobal.sbProtectPlan()

            Form.DataCenter.GlobalValues.bolPlanDrawInProgress = False
            Globals.ThisAddIn.Application.CutCopyMode = False
            Form.DataCenter.GlobalValues.bolCopy = False
            Form.DataCenter.GlobalValues.bolCut = False
            Globals.ThisAddIn.Application.EnableEvents = bolEnableEvents
            Globals.ThisAddIn.Application.ScreenUpdating = bolScreenUpdating
            Globals.ThisAddIn.Application.DisplayAlerts = bolDisplayAlerts
            SelectAfterRefresh(lngPS)
            Return True
        End Function
        Public Sub SelectAfterRefresh(lngPS As Long)
            Try

                Dim rng As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                With Form.DataCenter.GlobalValues.WS
                    With .Range(.Cells(Globals.ThisAddIn.Application.Selection.cells(1, 1).ROW, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), .Cells(Globals.ThisAddIn.Application.Selection.cells(1, 1).ROW, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn))
                        rng = .Find("=CellFace(""" & lngPS & ";*", , Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False, False)
                        If Not rng Is Nothing Then
                            rng2 = .Find("*", rng.Cells(1, 1), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False, False)
                            If Not rng2 Is Nothing Then
                                rng2 = Form.DataCenter.GlobalValues.WS.Range(rng, Form.DataCenter.GlobalValues.WS.Cells(rng.Row, rng2.Column - 1))
                                If Not Form.DataCenter.ProgramConfig.ISSearchActive Then
                                    Dim objFC As Excel.FormatCondition
                                    With rng2
                                        objFC = .FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,, "=True")
                                        objFC.Interior.Color = System.Drawing.Color.White
                                    End With
                                End If
                                rng2.Activate()
                                rng2.Select()
                            End If
                        End If
                    End With
                End With
            Catch ex As Exception
                If Globals.ThisAddIn.Application.Selection.column >= Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn Then
                    Form.DataCenter.GlobalValues.WS.Cells(3, Globals.ThisAddIn.Application.Selection.column).select
                Else
                    Form.DataCenter.GlobalValues.WS.Cells(4, Globals.ThisAddIn.Application.Selection.column).select
                End If
            End Try
        End Sub
        Public Sub GetResetFilter()
            Try

                currentFiltRange = ""
                col = 0
                bolWasOn = False

                With Form.DataCenter.GlobalValues.WS.AutoFilter
                    currentFiltRange = .Range.Address
                    With .Filters
                        ReDim filterArray(.Count - 1, 2)
                        For f = 1 To .Count
                            With .Item(f)
                                If .On = True Then
                                    bolWasOn = True
                                    filterArray(f - 1, 0) = .Criteria1
                                    Try
                                        filterArray(f - 1, 1) = .Operator
                                        filterArray(f - 1, 2) = .Criteria2
                                    Catch ex As Exception
                                    End Try
                                End If
                            End With
                        Next f
                    End With
                End With
                If bolWasOn = True Then Form.DataCenter.GlobalValues.WS.AutoFilterMode = False
            Catch ex As Exception
                'MsgBox(ex.Message,, "GetResetFilter")
            End Try

        End Sub
        Public Sub ReApplyFilter()
            Try
                For col = 0 To UBound(filterArray)
                    If filterArray(col, 0) IsNot Nothing Then
                        Try
                            Form.DataCenter.GlobalValues.WS.Range(currentFiltRange).AutoFilter(Field:=col + 1,
                                Criteria1:=filterArray(col, 0),
                                Operator:=filterArray(col, 1), Criteria2:=filterArray(col, 2))
                        Catch ex1 As Exception
                            Try
                                Form.DataCenter.GlobalValues.WS.Range(currentFiltRange).AutoFilter(Field:=col + 1,
                                    Criteria1:=filterArray(col, 0))
                            Catch ex As Exception
                            End Try
                        End Try
                    End If
                Next col
            Catch ex As Exception
                'MsgBox(ex.Message,, "ReApplyFilter")
            End Try
        End Sub
        Public Function ConvertObjToStrArray(Obj As Object) As String()
            Dim IntCnt As Integer = 0, strReturn() As String = Nothing, obj2 As Object = Nothing
            ReDim strReturn(UBound(Obj) - 1)
            For Each obj2 In Obj
                strReturn(IntCnt) = obj2.ToString.Replace("=", "")
                IntCnt = IntCnt + 1
            Next
            Return strReturn
        End Function

        Public Sub GetSearchFilter()
            Try

                currentFiltRange = ""
                col = 0
                bolWasOn = False

                With Form.DataCenter.GlobalValues.WS.AutoFilter
                    If Form.DataCenter.GlobalValues.WS.AutoFilter Is Nothing Then
                        Exit Sub
                    End If
                    currentFiltRange = .Range.Address
                    With .Filters
                        ReDim filterArray(.Count - 1, 2)
                        For f = 1 To .Count
                            With .Item(f)
                                If .On = True Then
                                    bolWasOn = True
                                    filterArray(f - 1, 0) = .Criteria1
                                    If .Operator Then
                                        Try
                                            filterArray(f - 1, 1) = .Operator
                                            filterArray(f - 1, 2) = .Criteria2
                                        Catch ex As Exception
                                        End Try
                                    End If
                                End If
                            End With
                        Next f
                    End With
                End With
                'If bolWasOn = True Then Form.DataCenter.GlobalValues.WS.AutoFilterMode = False
            Catch ex As Exception
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.GlobalFunction, ex.Message), "GetResetFilter", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                'MsgBox(ex.Message,, "GetResetFilter")
            End Try

        End Sub
        Public Sub MaxLengthValidationsSections()
            Try
                If Form.DataCenter.ProgramConfig.IsGeneric = True Then Exit Sub

                Dim rng As Excel.Range = Globals.ThisAddIn.Application.Union(Form.DataCenter.GlobalSections.FurtherBasicInformationSection,
                                                                    Form.DataCenter.GlobalSections.InstrumentationSection,
                                                                    Form.DataCenter.GlobalSections.MfcSpecificationSection,
                                                                    Form.DataCenter.GlobalSections.NonMfcSpecificationSection,
                                                                    Form.DataCenter.GlobalSections.ProgramInformationSection,
                                                                    Form.DataCenter.GlobalSections.UpdatePackSection)
                rng = Form.DataCenter.GlobalValues.WS.Range(Strings.Replace(Strings.Replace(rng.Address, "2", "5"), "4", Form.DataCenter.GlobalValues.WS.UsedRange.Rows.Count))

                With rng.Validation
                    If rng.Validation IsNot Nothing Then .Delete()
                    .Add(Excel.XlDVType.xlValidateTextLength, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, [Operator]:=Excel.XlFormatConditionOperator.xlLessEqual, Formula1:="200")
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ErrorTitle = "TnD Plan"
                    .ErrorMessage = "Sorry! only up to 200 characters are allowed to enter."
                    .ShowInput = False
                    .ShowError = True
                End With
            Catch ex As Exception
            End Try
        End Sub
    End Class
End Namespace

