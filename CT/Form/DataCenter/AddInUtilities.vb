Imports System.Data
Imports System.Runtime.InteropServices
Imports CT
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms

<ComVisible(True)>
Public Interface IAddInUtilities
    Sub RefreshChangeLog()
    Sub Deleterow()

    Sub OpenEditForm()

    Sub Addupdatechangelog(strVBAParam As Integer)

    Sub Cancel()

    Sub putEngineData(strEngine As String, strEngineType As String)

    Sub putTransData(strTrans As String, strTransType As String)

    Sub putOtherData(strValue As String)

    Sub Delete_PS_UC(rng As Excel.Range)

    Sub DeleteVehicle(rng As Excel.Range)

    Sub putPaintData(strPaintFacility As String, strColorCode As String)

End Interface

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.None)>
Public Class AddInUtilities
    Implements IAddInUtilities
    Public Sub Delete_PS_UC(rng As Excel.Range) Implements IAddInUtilities.Delete_PS_UC
        Try
            If Form.DataCenter.GlobalValues.bolUserCaseSelected = True Or Form.DataCenter.GlobalValues.bolSelAll = True Or
                Strings.InStr(rng.Cells(1, 1).formula, "=CELLFACE", CompareMethod.Text) > 0 Then
                Form.TndContextMenu.DeleteButton.Click()
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Sub DeleteVehicle(rng As Excel.Range) Implements IAddInUtilities.DeleteVehicle

        Try
            If rng.Column <> 5 Then Exit Sub
            Dim frmUnitDelete As New frmDeleteVehicle
            frmUnitDelete.frmDeleteVehicle_Shown 
            frmUnitDelete.cmdDelete_Click()
        Catch ex As Exception

        End Try
    End Sub
    'Purpose : To delete selected row data in sheet 'ChangeLogs'
    Public Sub Deleterow() Implements IAddInUtilities.Deleterow

        Dim shtGenChangeLog As Excel.Worksheet = Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString())
        Dim rng As Excel.Range

        Try

            shtGenChangeLog.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            rng = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            If rng.Rows.Count > 1 Then Exit Sub
            'If rng.Column < 2 Or rng.Column > 10 Then Exit Sub
            If rng.Row < 2 Then Exit Sub

            With shtGenChangeLog
                If .Cells(rng.Row, "A").Value <> "" Then
                    'objCon.Execute "EXEC Specific_DeleteChangeLogentry " & .Cells(Selection.Row, "A")
                    Dim _DataLog As Data.DataLog = New Data.DataLog()
                    If _DataLog.DeleteChangeLogentry(.Cells(rng.Row, "A").Value) = False Then Throw New Exception(Data.DataCenter.GlobalValues.message)
                    'MessageBox.Show("Sorry, your changes could not be saved to the database. Error description :- " & Data.DataCenter.GlobalValues.message, "Changelog - Delete row", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    'End If
                    RefreshChangeLog()
                    MessageBox.Show("Completed.", "Change Log", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Please select valid row.", "Change Log", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End With

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.AddInUtilities_DeleterowChangeLog, ex.Message), "Changelog - Delete row", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        Finally
            shtGenChangeLog.Protect(Form.DataCenter.GlobalValues.ConstPwd,,,,,, True)
        End Try

    End Sub

    'Purpose : To refresh data in sheet 'ChangeLogs'
    Public Sub RefreshChangeLog() Implements IAddInUtilities.RefreshChangeLog
        Dim shtGenChangeLog As Excel.Worksheet = Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString())
        Try
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString()).Activate
            shtGenChangeLog.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            If shtGenChangeLog.UsedRange.Rows.Count > 1 Then shtGenChangeLog.Range("A2:J" & shtGenChangeLog.UsedRange.Rows.Count).EntireRow.Delete()

            Dim _DataLog As Data.DataLog = New Data.DataLog()

            Dim _DataLogStringArray As String(,) = _DataLog.GetDataLog(Form.DataCenter.ProgramConfig.pe02)
            Dim top As Excel.Range = shtGenChangeLog.Cells(2, 1)
            Dim bottom As Excel.Range = shtGenChangeLog.Cells(_DataLogStringArray.GetUpperBound(0) + 1, _DataLogStringArray.GetUpperBound(1))
            Dim all As Excel.Range

            If bottom.Row > 1 Then
                all = shtGenChangeLog.Range(top, bottom)
                all.Value2 = _DataLogStringArray
                all.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                all.Borders.Weight = Excel.XlBorderWeight.xlThin
                For i As Int16 = 2 To bottom.Row
                    shtGenChangeLog.Cells(i, 2).value = i - 1
                Next
            End If

        Catch ex As Exception

        Finally
            Globals.ThisAddIn.Application.EnableEvents = True
            shtGenChangeLog.Protect(Form.DataCenter.GlobalValues.ConstPwd,,,,,, True)
        End Try
    End Sub

    'Purpose : To Add new row data into DB from sheet 'ChangeLogs'
    Public Sub Addupdatechangelog(strVBAParam As Integer) Implements IAddInUtilities.Addupdatechangelog
        sbAddUpdateChangeLog(strVBAParam)
    End Sub

    Function CheckLength(input As String, columnname As String) As Boolean
        CheckLength = False
        If columnname <> "Change Description" And input.Length > 100 Then
            MessageBox.Show(columnname & " column length should not exceed 100 characters. Data not saved.", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Exit Function
        ElseIf columnname = "Change Description" And input.Length > 500 Then
            MessageBox.Show(columnname & " column length should not exceed 500 characters. Data not saved.", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Exit Function
        End If
        CheckLength = True
    End Function

    'Purpose : To update row data into DB from sheet 'ChangeLogs'
    Public Sub sbAddUpdateChangeLog(intRow As Integer)
        Dim shtGenChangeLog As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString())
        Try
            Dim objRng As Excel.Range
            objRng = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)

            Dim _DataLog As Data.DataLog = New Data.DataLog()
            Dim bolResult As Boolean

            Dim pe62 As Long
            Dim pe02 As Long, dtChangedate As String, HCID As String, TnDIssue As String, BuildType As String, UnitId As String, ChangeDescription As String, Requestor As String, TnDResponsible As String

            shtGenChangeLog.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            With shtGenChangeLog
                Globals.ThisAddIn.Application.EnableEvents = False
                ''If .Cells(intRow, 3).text <> "" Then
                ''    .Cells(intRow, 3).NumberFormat = "@"
                ''    If DateTime.TryParseExact(.Cells(intRow, 3).Text, "d-M-yyyy", Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, Nothing) = False Then
                ''        MessageBox.Show("Please enter date value in format d-M-yyyy", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                ''        RefreshChangeLog()
                ''        .Cells(intRow, 3).Select
                ''        dtChangedate = Nothing
                ''        Exit Sub
                ''    Else
                ''        Dim intSplit(2) As Integer
                ''        intSplit(0) = Val(.Cells(intRow, "C").value2.ToString.Split("-")(0))
                ''        intSplit(1) = Val(.Cells(intRow, "C").value2.ToString.Split("-")(1))
                ''        intSplit(2) = Val(.Cells(intRow, "C").value2.ToString.Split("-")(2))
                ''        dtChangedate = DateSerial(intSplit(2), intSplit(1), intSplit(0))
                ''    End If
                ''Else
                ''    MessageBox.Show("Date filed cannot be blank.", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                ''    Exit Sub
                ''End If
                ''.Cells(intRow, 2).Value = intRow - 1
                ''Globals.ThisAddIn.Application.EnableEvents = True
                ''If .Cells(intRow, 4).text <> "" Then
                ''    If IsNumeric(shtGenChangeLog.Cells(intRow, 4).Text) = False Then
                ''        .Cells(intRow, 4).Select
                ''        MessageBox.Show("HC ID should be numeric.", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                ''        RefreshChangeLog()
                ''        .Cells(intRow, 4).Select
                ''        Exit Sub
                ''    End If
                ''End If
                ''If .Cells(intRow, 7).text <> "" Then
                ''    If IsNumeric(shtGenChangeLog.Cells(intRow, 7).Text) = False Then
                ''        .Cells(intRow, 7).Select
                ''        MessageBox.Show("Unit ID should be numeric.", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                ''        RefreshChangeLog()
                ''        .Cells(intRow, 7).Select
                ''        Exit Sub
                ''    End If
                ''End If
                'If IsDate(.Cells(intRow, "C").Value) = True Then
                'dtChangedate = Convert.ToDateTime(.Cells(intRow, "C").Value).ToString("yyyy-MM-dd")
                'End If
                pe02 = Form.DataCenter.ProgramConfig.pe02 '.Cells(intRow, "A").Value
                pe62 = .Cells(intRow, "A").Value
                dtChangedate = IIf(IsNothing(.Cells(intRow, "C").Value), "", .Cells(intRow, "C").Value)
                HCID = Val(.Cells(intRow, "D").Value)
                TnDIssue = IIf(IsNothing(.Cells(intRow, "E").Value), "", .Cells(intRow, "E").Value)
                BuildType = IIf(IsNothing(.Cells(intRow, "F").Value), "", .Cells(intRow, "F").Value)
                UnitId = IIf(IsNothing(.Cells(intRow, "G").Value), "", .Cells(intRow, "G").Value)
                ChangeDescription = IIf(IsNothing(.Cells(intRow, "H").Value), "", .Cells(intRow, "H").Value)
                Requestor = IIf(IsNothing(.Cells(intRow, "I").Value), "", .Cells(intRow, "I").Value)
                TnDResponsible = IIf(IsNothing(.Cells(intRow, "J").Value), "", .Cells(intRow, "J").Value)

                If CheckLength(dtChangedate, "Date") = False Then
                    .Cells(intRow, 3).Select
                    Exit Try
                End If
                If CheckLength(HCID, "HCID") = False Then
                    .Cells(intRow, 4).Select
                    Exit Try
                End If
                If CheckLength(TnDIssue, "Tnd Issue") = False Then
                    .Cells(intRow, 5).Select
                    Exit Try
                End If
                If CheckLength(BuildType, "Hardware") = False Then
                    .Cells(intRow, 6).Select
                    Exit Try
                End If
                If CheckLength(UnitId, "UnitId") = False Then
                    .Cells(intRow, 7).Select
                    Exit Try
                End If
                If CheckLength(ChangeDescription, "Change Description") = False Then
                    .Cells(intRow, 8).Select
                    Exit Try
                End If
                If CheckLength(Requestor, "Requestor") = False Then
                    .Cells(intRow, 9).Select
                    Exit Try
                End If
                If CheckLength(TnDResponsible, "TnDResponsible") = False Then
                    .Cells(intRow, 10).Select
                    Exit Try
                End If


                If .Cells(intRow, "A").Value <> "" Then
                    'objCon.Execute "EXEC Specific_UpdateChangeLogentry " & .Cells(intRow, "A") & ",'" & Strings.Format(.Cells(intRow, "C"), "yyyy-mm-dd") & "'," & Val(.Cells(intRow, "D")) & ",'" &
                    '            .Cells(intRow, "E") & "','" & .Cells(intRow, "F") & "','" & .Cells(intRow, "G") & "','" & .Cells(intRow, "H") & "','" & .Cells(intRow, "I") & "','" & .Cells(intRow, "J") & "'"
                    If _DataLog.UpdateChangeLog(pe62:=pe62, ChangeDate:=dtChangedate, HCID:=HCID, TnDIssue:=TnDIssue, BuildType:=BuildType, UnitId:=UnitId, ChangeDescription:=ChangeDescription, Requestor:=Requestor, TnDResponsible:=TnDResponsible) = False Then Throw New Exception(Data.DataCenter.GlobalValues.message)
                    'If bolResult = False Then
                    '    MessageBox.Show("Sorry, your changes could not be saved to the database. Error description :- " & Data.DataCenter.GlobalValues.message, "Change Logs", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    'End If
                Else
                    'Set rst = objCon.Execute("EXEC Specific_AddChangeLogentry " & Val(ShtProgConfig.Cells(13, "B")) & ",'" & Strings.Format(.Cells(intRow, "C"), "yyyy-mm-dd") & "'," & Val(.Cells(intRow, "D")) & ",'" & _
                    '            .Cells(intRow, "E") & "','" & .Cells(intRow, "F") & "','" & .Cells(intRow, "G") & "','" & .Cells(intRow, "H") & "','" & .Cells(intRow, "I") & "','" & .Cells(intRow, "J") & "'")
                    'bolResult = _DataLog.AddChangeLog(pe02:= .Cells(intRow, "A").Value, ChangeDate:=vbNullString, HCID:=Val(.Cells(intRow, "D").Value), TnDIssue:= .Cells(intRow, "E").Value, BuildType:= .Cells(intRow, "F").Value, UnitId:= .Cells(intRow, "G").Value, ChangeDescription:= .Cells(intRow, "H").Value, Requestor:= .Cells(intRow, "I").Value, TnDResponsible:= .Cells(intRow, "J").Value)
                    If _DataLog.AddChangeLog(pe02:=pe02, ChangeDate:=dtChangedate, HCID:=HCID, TnDIssue:=TnDIssue, BuildType:=BuildType, UnitId:=UnitId, ChangeDescription:=ChangeDescription, Requestor:=Requestor, TnDResponsible:=TnDResponsible) = False Then Throw New Exception(Data.DataCenter.GlobalValues.message)
                    'If bolResult = False Then
                    '    MessageBox.Show("Sorry, your changes could not be saved to the database. Error description :- " & Data.DataCenter.GlobalValues.message, "Change Logs", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    'Else
                    ''''''.Cells(intRow, "A") = rst.Fields(0) '''To do pending
                    RefreshChangeLog()
                    'End If
                End If
            End With
            'shtGenChangeLog.Protect(Form.DataCenter.GlobalValues.ConstPwd,,,,,, True)
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.AddInUtilities_AddUpdateChangeLog, ex.Message), "Add Update Changelog", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            'MessageBox.Show(ex.Message, "Add Update Changelog", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            shtGenChangeLog.Protect(Form.DataCenter.GlobalValues.ConstPwd,,,,,, True)
            Globals.ThisAddIn.Application.EnableEvents = True
        End Try
    End Sub

    Public Sub OpenEditForm() Implements IAddInUtilities.OpenEditForm
        Try
            '---------------------------------------------------------------------------
            ' For generic plan and Master Mode this window must not get open.
            '---------------------------------------------------------------------------
            'If Form.DataCenter.ProgramConfig.IsGeneric = True Or Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then Exit Sub
            If Form.DataCenter.ProgramConfig.IsGeneric = True Then Exit Sub

            '---------------------------------------------------------------------------
            ' For specific plan and Checkedout & Draft Modes this window can get open but with permission.
            '---------------------------------------------------------------------------
            If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Trim.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or
                Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                Exit Sub
            End If

            '------------------------------------------------
            ' open window if user granted.
            '------------------------------------------------
            'Dim _editform As New frmHeaderEdit
            Dim _frmObject As Object
            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _frmObject = New frmHeaderEdit
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _frmObject = New frmHeaderEdit_Rig
            Else
                Exit Sub
            End If
            _frmObject.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Open Header Editor", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub Cancel() Implements IAddInUtilities.Cancel
        Try
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
        Catch ex As Exception
        End Try
        Globals.ThisAddIn.Application.CutCopyMode = False
        Form.DataCenter.GlobalValues.bolCutCopyMode = False
        Form.DataCenter.GlobalValues.strCutAddress = ""
        Form.DataCenter.GlobalValues.strCopyAddress = ""
        Form.DataCenter.GlobalValues.bolUserCaseSelected = False
        Form.DataCenter.GlobalValues.strUserCaseSelected = ""
        Form.DataCenter.GlobalValues.bolSelAll = False
        Form.DataCenter.GlobalValues.strSelAllAddress = ""
        Form.DataCenter.GlobalValues.bolCopy = False
        Form.DataCenter.GlobalValues.bolCut = False
        If Globals.ThisAddIn.Application.Selection.column >= Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn Then
            Form.DataCenter.GlobalValues.WS.Cells(3, Globals.ThisAddIn.Application.Selection.column).select
        Else
            Form.DataCenter.GlobalValues.WS.Cells(4, Globals.ThisAddIn.Application.Selection.column).select
        End If
        Form.DataCenter.GlobalValues.WS.Cells.FormatConditions.Delete()
        Dim _obj As New Form.DataCenter.ModuleFunction
        _obj.sbProtectPlan()

    End Sub
    Sub putEngineData(strEngine As String, strEngineType As String) Implements IAddInUtilities.putEngineData

        Globals.ThisAddIn.Application.EnableEvents = False

        Dim _Unit As New CT.Data.VehiclePlan.Unit

        If _Unit.ChangeEngine(Form.DataCenter.VehicleConfig.VehiclePe02, Form.DataCenter.VehicleConfig.VehiclePe45, strEngine, Form.DataCenter.ProgramConfig.BuildType) = True Then
            Try
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Catch ex As Exception
            End Try
            Dim Cls As New Form.DataCenter.GlobalFunctions
            With Form.DataCenter.GlobalValues.WS

                .Cells(.Application.Selection.row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column) = strEngine
                .Cells(.Application.Selection.row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Type_Column) = strEngineType
                Cls.UpdateSection(.Application.Selection.row, .Application.Selection.row)

                'Dim rngFnd2 As Excel.Range = Nothing, rngFnd4 As Excel.Range = Nothing

                'rngFnd2 = .Range(.Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn)).Find("Engine Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                'rngFnd4 = .Range(.Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn)).Find("Engine Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                'If rngFnd2 IsNot Nothing And rngFnd4 IsNot Nothing Then
                '    ' Globals.ThisAddIn.Application.EnableEvents = True
                '    .Cells(.Application.Selection.row, rngFnd2.Column).value2 = .Cells(.Application.Selection.row, rngFnd4.Column).value2
                '    'Globals.ThisAddIn.Application.EnableEvents = False
                'End If
            End With


            Dim _obj As New Form.DataCenter.ModuleFunction
            _obj.sbProtectPlan()
        Else
            System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
        End If

        Globals.ThisAddIn.Application.EnableEvents = True
        Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
        _RibbonUtilitis.UpdateUndoButtonsState()
        DeleteContextMenu()
        Globals.ThisAddIn.Application.CommandBars("Cell").Reset()
        Globals.ThisAddIn.Application.Selection.select

    End Sub

    Sub putTransData(strTrans As String, strTransType As String) Implements IAddInUtilities.putTransData

        Globals.ThisAddIn.Application.EnableEvents = False

        Dim _Unit As New CT.Data.VehiclePlan.Unit

        If _Unit.ChangeTransmission(Form.DataCenter.VehicleConfig.VehiclePe02, Form.DataCenter.VehicleConfig.VehiclePe45, strTrans, Form.DataCenter.ProgramConfig.BuildType) = True Then
            With Form.DataCenter.GlobalValues.WS
                Try
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                Catch ex As Exception
                End Try
                Dim Cls As New Form.DataCenter.GlobalFunctions

                .Cells(.Application.Selection.row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column) = strTrans
                .Cells(.Application.Selection.row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Type_Column) = strTransType

                ''Updating Further Basic Information section 'Transmission Type column based on Column "Q" value
                'Dim FindColumn As Excel.Range = Nothing

                'FindColumn = Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Find(.Cells(4, "Q").Value, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                'If Not (FindColumn Is Nothing) Then
                '    If .Cells(.Application.Selection.row, FindColumn.Column).Value <> "" Then
                '        .Cells(.Application.Selection.row, FindColumn.Column) = strTransType
                '    End If
                'End If

                Cls.UpdateSection(.Application.Selection.row, .Application.Selection.row)

                'Dim rngFnd3 As Excel.Range = Nothing, rngFnd5 As Excel.Range = Nothing

                'rngFnd3 = .Range(.Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn)).Find("Transmission Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                'rngFnd5 = .Range(.Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn)).Find("Transmission Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                '                                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)

                'If rngFnd3 IsNot Nothing And rngFnd5 IsNot Nothing Then
                '    'Globals.ThisAddIn.Application.EnableEvents = True
                '    .Cells(.Application.Selection.row, rngFnd3.Column).value2 = .Cells(.Application.Selection.row, rngFnd5.Column).value2
                '    'Globals.ThisAddIn.Application.EnableEvents = False
                'End If

                Dim _obj As New Form.DataCenter.ModuleFunction
                _obj.sbProtectPlan()

            End With
        Else
            System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
        End If

        Globals.ThisAddIn.Application.EnableEvents = True
        Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
        _RibbonUtilitis.UpdateUndoButtonsState()
        DeleteContextMenu()
        Globals.ThisAddIn.Application.CommandBars("Cell").Reset()
        Globals.ThisAddIn.Application.Selection.select
    End Sub

    Sub putOtherData(strValue As String) Implements IAddInUtilities.putOtherData
        Try
            Try
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Catch ex As Exception
            End Try
            Globals.ThisAddIn.Application.EnableEvents = True

            Globals.ThisAddIn.Application.Selection.value = strValue

            Dim _obj As New Form.DataCenter.ModuleFunction
            _obj.sbProtectPlan()

            DeleteContextMenu()
            Globals.ThisAddIn.Application.CommandBars("Cell").Reset()
            Globals.ThisAddIn.Application.Selection.select


        Catch ex As Exception
        End Try
    End Sub
    Sub putPaintData(strPaintFacility As String, strColorCode As String) Implements IAddInUtilities.putPaintData
        Try
            Try
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            Catch ex As Exception
            End Try
            With Form.DataCenter.GlobalValues.WS
                Globals.ThisAddIn.Application.EnableEvents = True
                .Cells(Globals.ThisAddIn.Application.Selection.row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).value = strPaintFacility

                'Try
                '    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                'Catch ex As Exception
                'End Try

                'Globals.ThisAddIn.Application.EnableEvents = True
                '.Cells(Globals.ThisAddIn.Application.Selection.row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).value = strColorCode.Split("~")(1)

            End With

            Dim _obj As New Form.DataCenter.ModuleFunction
            _obj.sbProtectPlan()

            DeleteContextMenu()
            Globals.ThisAddIn.Application.CommandBars("Cell").Reset()
            Globals.ThisAddIn.Application.Selection.select

        Catch ex As Exception
            Globals.ThisAddIn.Application.EnableEvents = True
        End Try
    End Sub
    Public Sub DeleteContextMenu()
        Try
            For Each btn As Office.CommandBarControl In Globals.ThisAddIn.Application.CommandBars("Cell").Controls
                btn.Delete()
            Next

        Catch ex As Exception

        End Try

    End Sub
End Class
