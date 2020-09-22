
Imports System.Data
Imports System.Windows.Forms
Namespace Form.DisplayUtilities.Ribbon
    Friend NotInheritable Class UndoButton
        ''' <summary>
        ''' The undo function is called here and the refresh must be done here.
        ''' </summary>
        ''' <param name="row"></param>
        Public Shared Sub Click(pe61 As Integer, MessageForDisplaying As String)
            Dim dtAnswer As DataTable = Nothing
            Dim _ChangeLog As CT.Data.ChangeLog = New Data.ChangeLog
            Try
                Globals.ThisAddIn.Application.ScreenUpdating = False
                dtAnswer = _ChangeLog.UndoPreviousOperation(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                If dtAnswer Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                If dtAnswer IsNot Nothing Then
                    '--------------------------------------------------------------
                    ' Refresh desktop
                    'Note: All Undo Actionnames from storedprocedure should return 'Pe03, pe45, pe57 - to perform refresh operation in excel
                    '--------------------------------------------------------------
                    Dim strValues() As String
                    Dim Cls As New Form.DataCenter.GlobalFunctions
                    Dim rngFind As Excel.Range = Nothing
                    Select Case pe61
                        Case 1222 'For other than return values "pe03pe45pe57" or delete unit --- scenarious to be written
                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedFurtherBasic
                            '39
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedInstrumentation
                            '36
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedMFC
                            '8
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedUpdatePack
                            '38
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedProgramInfo
                            '40
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedUserShipping
                            '41
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_DeletedNonMFC
                            '37
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewInstrumentation
                            '------------------------------------------- UNDO NEW COLUMN IN 7TABS -------------------------------------------------------
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewMFC
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewNonMFC
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewProgramInfo
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewFurtherBasic
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewUserShipping
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_NewUpdatepack
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColInstrumentation
                            '------------------------------------------- UNDO EDIT COLUMN IN 7TABS -------------------------------------------------------
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColMFC
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColNonMFC
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColProgramInfo
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColFurtherBasic
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColUserShipping
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection)
                        '-----------------------------------------------------------------------------------------------------------------------------------

                        Case CT.Data.DataCenter.ActionName.Tnd_EditedColUpdatePack
                            Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection)
                            '-----------------------------------------------------------------------------------------------------------------------------------


                        Case CT.Data.DataCenter.ActionName.Tnd_Timing
                            '--------------------------------------------------------------------------------
                            ' Only refresh timing area
                            ' Apply chnages after update
                            '--------------------------------------------------------------
                            Dim _DrawTndPlanHeader As Form.DisplayUtilities.DrawTndPlanHeader = New Form.DisplayUtilities.DrawTndPlanHeader
                            _DrawTndPlanHeader.ApplyHolidaysFlags()
                            _DrawTndPlanHeader.ApplyGatewayFlags()

                        Case CT.Data.DataCenter.ActionName.Tnd_DeleteVehicle
                            '----------------------------------------------------------------------
                            ' INSERT DELETED VEHICLE IN THIS POSITION
                            '----------------------------------------------------------------------
                            'Cls.GetResetFilter()
                            Dim obj As Object
                            Dim NewRow As Integer = -1
                            For rowCounter As Integer = 5 To Form.DataCenter.ProgramConfig.LastRow

                                obj = Form.DataCenter.GlobalValues.WS.Cells(rowCounter, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).Value2
                                If (obj IsNot Nothing) Then

                                    If Val(obj) > Val(dtAnswer.Rows(0)("DisplaySeq")) And NewRow = -1 Then
                                        NewRow = rowCounter
                                        Exit For
                                    End If
                                End If

                            Next

                            If NewRow > -1 Then
                                With Form.DataCenter.GlobalValues.WS

                                    .Range("B" & NewRow).EntireRow.Copy()
                                    .Rows(NewRow).Insert()
                                    .Range("B" & NewRow).EntireRow.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats)
                                    .Range(.Cells(NewRow, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), .Cells(NewRow, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Interior.Color = 16777215 'xlNone
                                    .Cells(NewRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column) = dtAnswer.Rows(0)("DisplaySeq")
                                End With
                                Cls.UpdateSection(NewRow, NewRow, False)
                            Else

                                With Form.DataCenter.GlobalValues.WS
                                    NewRow = Form.DataCenter.ProgramConfig.LastRow
                                    .Range("B" & NewRow).EntireRow.Copy()
                                    .Range("B" & NewRow + 1).EntireRow.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats)
                                    .Range(.Cells(NewRow + 1, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), .Cells(NewRow + 1, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Interior.Color = 16777215 'xlNone
                                    .Cells(NewRow + 1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column) = dtAnswer.Rows(0)("DisplaySeq")
                                End With
                                Cls.UpdateSection(NewRow + 1, NewRow + 1, False)
                            End If
                            Form.DataCenter.GlobalValues.TotalRow = Form.DataCenter.GlobalValues.TotalRow + 1
                        Case CT.Data.DataCenter.ActionName.Tnd_AddNewVehicle
                            '----------------------------------------------------------------------
                            ' DELETE INSERTED VEHICLE IN THIS POSITION
                            '----------------------------------------------------------------------
                            For Each dr As DataRow In dtAnswer.Rows 'Loop in Database Undo operation returned rows to refresh --- return must be 'dr("pe03pe45pe57")
                                With Form.DataCenter.GlobalValues.WS
                                    rngFind = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column), .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column)).Find("*" & dr(0) & "*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                End With


                                If Not rngFind Is Nothing Then
                                    With Form.DataCenter.GlobalValues.WS

                                        .Rows(rngFind.Row).Delete

                                    End With
                                End If
                            Next

                        Case CT.Data.DataCenter.ActionName.Tnd_ChangeVehicleSequence
                            '----------------------------------------------------------------------
                            ' Change Sequence of VEHICLE IN THIS POSITION
                            '----------------------------------------------------------------------
                            'For Each dr As DataRow In dtAnswer.Rows 'Loop in Database Undo operation returned rows to refresh --- return must be 'dr("pe03pe45pe57")
                            '    rngFind = Form.DataCenter.GlobalValues.WS.Range("E5:E" & Form.DataCenter.ProgramConfig.LastRow).Find(dr(1), Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole,
                            '                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                            '    If Not rngFind Is Nothing Then
                            '        Globals.ThisAddIn.Application.ScreenUpdating = False
                            '        Cls.UpdateSection(rngFind.Row, rngFind.Row) 'Refresh matching row
                            '        Globals.ThisAddIn.Application.ScreenUpdating = True
                            '    End If
                            'Next
                            Dim dr1 As DataRow = dtAnswer.Rows(0)
                            Dim dr2 As DataRow = dtAnswer.Rows(dtAnswer.Rows.Count - 1)
                            Dim rngFind1 As Excel.Range = Nothing, rngFind2 As Excel.Range = Nothing
                            With Form.DataCenter.GlobalValues.WS
                                rngFind1 = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column), .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column)).Find("*" & dr1(0) & "*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                                            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                rngFind2 = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column), .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column)).Find("*" & dr2(0) & "*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                                            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                            End With

                            If rngFind1 IsNot Nothing And rngFind2 IsNot Nothing Then
                                Globals.ThisAddIn.Application.ScreenUpdating = False
                                If rngFind2.Row >= rngFind1.Row Then
                                    Cls.UpdateSection(rngFind1.Row, rngFind2.Row)
                                Else
                                    Cls.UpdateSection(rngFind2.Row, rngFind1.Row)
                                End If
                                Globals.ThisAddIn.Application.ScreenUpdating = True
                            End If
                        Case CT.Data.DataCenter.ActionName.Tnd_CutPasteProcessStep, CT.Data.DataCenter.ActionName.Tnd_CutPasteUsercase
                            Dim strMessage As String = String.Empty
                            Dim _DrawTndPlanArea As Form.DisplayUtilities.DrawTndPlanArea = New Form.DisplayUtilities.DrawTndPlanArea()
                            For Each dr As DataRow In dtAnswer.Rows 'Loop in Database Undo operation returned rows to refresh --- return must be 'dr("pe03pe45pe57")
                                With Form.DataCenter.GlobalValues.WS
                                    rngFind = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column), .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column)).Find("*" & dr(0) & "*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                                            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                End With

                                If Not rngFind Is Nothing Then
                                    Globals.ThisAddIn.Application.ScreenUpdating = False
                                    Cls.UpdateSection(rngFind.Row, rngFind.Row) 'Refresh matching row
                                    Globals.ThisAddIn.Application.ScreenUpdating = True
                                End If
                            Next

                        Case Else

                            With Form.DataCenter.GlobalValues.WS
                                'Dim rngFnd2 As Excel.Range = Nothing, rngFnd3 As Excel.Range = Nothing, rngFnd4 As Excel.Range = Nothing, rngFnd5 As Excel.Range = Nothing
                                'If CT.Data.DataCenter.ActionName.Tnd_EngineInfo = pe61 Then
                                '    rngFnd2 = .Range(.Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn)).Find("Engine Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                '                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                '    rngFnd4 = .Range(.Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn)).Find("Engine Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                '                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                'End If
                                'If CT.Data.DataCenter.ActionName.Tnd_TransInfo = pe61 Then
                                '    rngFnd3 = .Range(.Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn)).Find("Transmission Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                '                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                '    rngFnd5 = .Range(.Cells(4, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn), .Cells(2, Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn)).Find("Transmission Type", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                '                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                'End If
                                For Each dr As DataRow In dtAnswer.Rows 'Loop in Database Undo operation returned rows to refresh --- return must be 'dr("pe03pe45pe57")
                                    With Form.DataCenter.GlobalValues.WS
                                        rngFind = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column), .Cells(Form.DataCenter.ProgramConfig.LastRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_P_0_Column)).Find("*" & dr(0) & "*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                                    End With

                                    If Not rngFind Is Nothing Then
                                        Globals.ThisAddIn.Application.ScreenUpdating = False
                                        Cls.UpdateSection(rngFind.Row, rngFind.Row) 'Refresh matching row
                                        'If rngFnd2 IsNot Nothing And rngFnd4 IsNot Nothing Then
                                        '    If CT.Data.DataCenter.ActionName.Tnd_EngineInfo = pe61 Then
                                        '        .Cells(rngFind.Row, rngFnd2.Column).value2 = .Cells(rngFind.Row, rngFnd4.Column).value2
                                        '    End If
                                        'End If
                                        'If rngFnd3 IsNot Nothing And rngFnd5 IsNot Nothing Then
                                        '    If CT.Data.DataCenter.ActionName.Tnd_TransInfo = pe61 Then
                                        '        .Cells(rngFind.Row, rngFnd3.Column).value2 = .Cells(rngFind.Row, rngFnd5.Column).value2
                                        '    End If
                                        'End If
                                        Globals.ThisAddIn.Application.ScreenUpdating = True
                                    End If
                                Next
                            End With
                    End Select
                    Globals.ThisAddIn.Application.EnableEvents = True
                    Globals.ThisAddIn.Application.ScreenUpdating = True
                    MessageBox.Show("Undo the following action " + vbNewLine + MessageForDisplaying + "is Done.", "Undo button", MessageBoxButtons.OK)
                End If
            Catch ex As Exception
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.EnableEvents = True
                MessageBox.Show(ex.Message, "Undo Click function", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Globals.ThisAddIn.Application.ScreenUpdating = True
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
                Globals.ThisAddIn.Application.EnableEvents = True
            End Try

        End Sub
    End Class
End Namespace


'For Each dr As DataRow In dtAnswer.Rows 'Loop in Database Undo operation returned rows to refresh
'    For i As Int16 = 5 To Form.DataCenter.ProgramConfig.LastRow 'Loop in excel column "C" values
'        strValues = Form.DataCenter.GlobalValues.WS.Cells(i, 3).ToString().Split(";") 'Split excel sheet column 'C' values
'        If dr("pe03").ToString() = strValues(2) And dr("pe45").ToString() = strValues(3) And dr("pe57").ToString() = strValues(4) Then 'Compare with DB to Excel Col 'C'
'            Cls.UpdateSection(i, i) 'Refresh matching row
'            Exit For
'        End If
'    Next
'Next