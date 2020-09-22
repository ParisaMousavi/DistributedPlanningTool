
Namespace Form.TndContextMenu

    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class CopyButton
        Public Shared Sub Click(strAddress As String)
            Try
                If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                    Throw New Exception("Sorry, this operation is prohibited in 'Generic' plan. If you have permission, you can convert to 'Specific' using the tool button and start editing. Please contact Eren, Ali (A.) (aeren8@ford.com) for further support.")
                End If

                If Form.DataCenter.GlobalValues.bolUserCaseSelected = True And Form.DataCenter.GlobalValues.strUserCaseSelected <> "" Then
                    Form.DataCenter.GlobalValues.strCopyAddress = Form.DataCenter.GlobalValues.strUserCaseSelected
                ElseIf Form.DataCenter.GlobalValues.bolSelAll And Form.DataCenter.GlobalValues.strSelAllAddress <> "" Then
                    Form.DataCenter.GlobalValues.strCopyAddress = Form.DataCenter.GlobalValues.strSelAllAddress
                Else
                    Form.DataCenter.GlobalValues.strCopyAddress = strAddress
                End If

                Dim rng As Excel.Range
                rng = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strCopyAddress)
                rng.Copy()
                Form.DataCenter.GlobalValues.bolCopy = True
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Copy Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Finally
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                'Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                '_RibbonUtilitis.UpdateUndoButtonsState()
            End Try
        End Sub
        Public Shared Sub InsertBefore()

            Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep
            Dim lstPS As New List(Of Long)
            Dim rngT As Excel.Range
            Dim rng As Excel.Range
            Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
            Dim intCutRow As Integer = 0

            rng = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strCopyAddress)
            intCutRow = rng.Row
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            Try
                For Each rngT In rng.Cells
                    If rngT.Formula IsNot Nothing And rngT.Formula <> "" And rngT.Formula <> "-" Then
                        lstPS.Add(CLng(rngT.Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim()))
                    End If
                Next
                If lstPS.Count > 0 Then
                    If Form.DataCenter.GlobalValues.bolCopy Then
                        If _ProcessStep.CopyPaste(DataCenter.VehicleConfig.VehiclePe02(rng.Worksheet.Application.Selection.row),
                                                  DataCenter.VehicleConfig.VehiclePe45(rng.Worksheet.Application.Selection.row),
                                                  DataCenter.VehicleConfig.VehicleHCID(rng.Worksheet.Application.Selection.row),
                                                  DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase,
                                                  DataCenter.ProcessStepConfig.ProcessStepSequence,
                                                  DataCenter.ProcessStepConfig.ProcessStepStartDate,
                                                  lstPS,
                                                  Form.DataCenter.ProgramConfig.BuildType,
                                                  Nothing,
                                                  False) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                    ElseIf Form.DataCenter.GlobalValues.bolCut Then
                        If _ProcessStep.CutPaste(DataCenter.VehicleConfig.VehiclePe02(intCutRow),
                                                 DataCenter.VehicleConfig.VehiclePe45(intCutRow),
                                                 DataCenter.VehicleConfig.VehiclePe45(rng.Worksheet.Application.Selection.row),
                                                 DataCenter.VehicleConfig.VehicleHCID(rng.Worksheet.Application.Selection.row),
                                                 DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase,
                                                 DataCenter.ProcessStepConfig.ProcessStepSequence,
                                                 DataCenter.ProcessStepConfig.ProcessStepStartDate,
                                                 lstPS,
                                                 Form.DataCenter.ProgramConfig.BuildType,
                                                 Nothing,
                                                 False) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                    End If
                End If
            Catch ex As Exception
                If Form.DataCenter.GlobalValues.bolCopy Then
                    _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepStartDate)
                    Cancel()
                    System.Windows.Forms.MessageBox.Show(ex.Message, "InsertBefore Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepStartDate)
                    _GlobalFunctions.UpdateSection(intCutRow, intCutRow)
                    Cancel()
                    System.Windows.Forms.MessageBox.Show(ex.Message, "InsertBefore Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                End If
                _RibbonUtilitis.UpdateUndoButtonsState()
                'Exit Sub
            Finally
                If Form.DataCenter.GlobalValues.bolCopy Then
                    _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepStartDate)
                    Cancel()
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepStartDate)
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    _GlobalFunctions.UpdateSection(intCutRow, intCutRow)
                    Cancel()
                End If

                _RibbonUtilitis.UpdateUndoButtonsState()

            End Try
            'If Form.DataCenter.GlobalValues.bolCopy Then
            '    _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepStartDate)
            '    Cancel()
            'ElseIf Form.DataCenter.GlobalValues.bolCut Then
            '    _GlobalFunctions.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepStartDate)
            '    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            '    _GlobalFunctions.UpdateSection(intCutRow, intCutRow)
            '    Cancel()
            'End If

            '_RibbonUtilitis.UpdateUndoButtonsState()
        End Sub

        Public Shared Sub Cancel()
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
            'If Globals.ThisAddIn.Application.Selection.column >= Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn Then
            '    Form.DataCenter.GlobalValues.WS.Cells(3, Globals.ThisAddIn.Application.Selection.column).select
            'Else
            '    Form.DataCenter.GlobalValues.WS.Cells(4, Globals.ThisAddIn.Application.Selection.column).select
            'End If
            'Form.DataCenter.WS.Cells.FormatConditions.Delete()
            Dim _obj As New Form.DataCenter.ModuleFunction
            _obj.sbProtectPlan()
        End Sub
        Public Shared Sub InsertAfter()

            Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep
            Dim lstPS As New List(Of Long)
            Dim rngT As Excel.Range
            Dim rng As Excel.Range
            Dim Cls As New Form.DataCenter.GlobalFunctions
            Dim intCutRow As Integer = 0
            Dim CurUC As String = ""

            rng = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strCopyAddress)
            intCutRow = rng.Row
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            Try
                For Each rngT In rng.Cells
                    If rngT.Formula IsNot Nothing And rngT.Formula <> "" And rngT.Formula <> "-" Then
                        lstPS.Add(CLng(rngT.Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim()))
                    End If
                Next

                CurUC = rng.Worksheet.Application.Selection.cells(1, 1).Formula.ToString.Split(";")(3).Trim()

                If lstPS.Count > 0 Then
                    If Form.DataCenter.GlobalValues.bolCopy = True Then
                        If _ProcessStep.CopyPaste(DataCenter.VehicleConfig.VehiclePe02(rng.Worksheet.Application.Selection.row),
                                                  DataCenter.VehicleConfig.VehiclePe45(rng.Worksheet.Application.Selection.row),
                                                  DataCenter.VehicleConfig.VehicleHCID(rng.Worksheet.Application.Selection.row),
                                                  IIf(CurUC = DataCenter.ProcessStepConfig.ProcessStepUserCase, DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase, DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase + 1),
                                                  IIf(CurUC = DataCenter.ProcessStepConfig.ProcessStepUserCase, DataCenter.ProcessStepConfig.ProcessStepSequence + 1, 1),
                                                  DataCenter.ProcessStepConfig.ProcessStepEndDate,
                                                  lstPS,
                                                  Form.DataCenter.ProgramConfig.BuildType,
                                                  Nothing,
                                                  False) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                    ElseIf Form.DataCenter.GlobalValues.bolCut Then
                        If _ProcessStep.CutPaste(DataCenter.VehicleConfig.VehiclePe02(intCutRow), DataCenter.VehicleConfig.VehiclePe45(intCutRow),
                                                 DataCenter.VehicleConfig.VehiclePe45(rng.Worksheet.Application.Selection.row),
                                                 DataCenter.VehicleConfig.VehicleHCID(rng.Worksheet.Application.Selection.row),
                                                 IIf(CurUC = DataCenter.ProcessStepConfig.ProcessStepUserCase, DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase, DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase + 1),
                                                 IIf(CurUC = DataCenter.ProcessStepConfig.ProcessStepUserCase, DataCenter.ProcessStepConfig.ProcessStepSequence + 1, 1),
                                                 DataCenter.ProcessStepConfig.ProcessStepEndDate,
                                                 lstPS,
                                                 Form.DataCenter.ProgramConfig.BuildType,
                                                 Nothing,
                                                 False) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                    End If
                End If
            Catch ex As Exception
                If Form.DataCenter.GlobalValues.bolCopy Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepEndDate)
                    Cancel()
                    System.Windows.Forms.MessageBox.Show(ex.Message, "InsertAfter Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepEndDate)
                    Cls.UpdateSection(intCutRow, intCutRow)
                    Cancel()
                    System.Windows.Forms.MessageBox.Show(ex.Message, "InsertAfter Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                End If
                _RibbonUtilitis.UpdateUndoButtonsState()
                'Exit Sub
            Finally

                If Form.DataCenter.GlobalValues.bolCopy Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepEndDate)
                    Cancel()
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepEndDate)
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    Cls.UpdateSection(intCutRow, intCutRow)
                    Cancel()
                End If
                _RibbonUtilitis.UpdateUndoButtonsState()

            End Try
            'If Form.DataCenter.GlobalValues.bolCopy Then
            '    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepEndDate)
            '    Cancel()
            'ElseIf Form.DataCenter.GlobalValues.bolCut Then
            '    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, DataCenter.ProcessStepConfig.ProcessStepEndDate)
            '    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            '    Cls.UpdateSection(intCutRow, intCutRow)
            '    Cancel()
            'End If
            '_RibbonUtilitis.UpdateUndoButtonsState()
        End Sub
        Public Shared Sub Insert()
            Dim lstPS As New List(Of Tuple(Of Long, Short))
            Dim lstUCwisePS As New List(Of Long)
            Dim rngT As Excel.Range
            Dim rng As Excel.Range
            Dim Cls As New Form.DataCenter.GlobalFunctions
            Dim intCutRow As Integer = 0
            Dim intCurPS As Integer = 0
            Dim intCurUC As Integer = 0
            Dim CurUC As String = ""
            Dim intFCol As Integer = 0, intLCol As Integer = 0
            Dim NextPlannedStart As Object = Nothing
            Dim InsertAsIndependentUsercase As Boolean = False

            rng = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strCopyAddress)
            intCutRow = rng.Row
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            Try
                Dim int_Ucno As Short 'To store user case number and differentiate process step from next usercase

                int_Ucno = 0

                For Each rngT In rng.Cells
                    If rngT.Formula IsNot Nothing And rngT.Formula <> "" And rngT.Formula <> "-" Then
                        lstPS.Add(Tuple.Create(CLng(rngT.Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim()), CShort(rngT.Formula.ToString.Split(";")(1))))
                    End If
                Next
                '2001, 0 , Build
                '2002, 1
                '2003, 2
                '2004, 2
                '2005, 2

                While lstPS.Count > 0


                    int_Ucno = lstPS.Min(Function(i) i.Item2)
                    For Each a As Tuple(Of Long, Short) In lstPS.FindAll(Function(ps) ps.Item2 = int_Ucno)
                        lstUCwisePS.Add(a.Item1)
                    Next

                    InsertAsIndependentUsercase = True
                    NextPlannedStart = SaveSteps(intCutRow, lstUCwisePS, rng, NextPlannedStart, InsertAsIndependentUsercase)
                    lstPS.RemoveAll(Function(ps) ps.Item2 = int_Ucno)
                    lstUCwisePS.Clear()

                End While


            Catch ex As Exception
                If Form.DataCenter.GlobalValues.bolCopy Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                    Cancel()
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Insert Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                    Cls.UpdateSection(intCutRow, intCutRow)
                    Cancel()
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Insert Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                End If

                _RibbonUtilitis.UpdateUndoButtonsState()
                'Exit Sub
            Finally
                If Form.DataCenter.GlobalValues.bolCopy Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                    Cancel()
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    Cls.UpdateSection(intCutRow, intCutRow)
                    Cancel()
                End If

                _RibbonUtilitis.UpdateUndoButtonsState()

            End Try

            'If Form.DataCenter.GlobalValues.bolCopy Then
            '    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.WS.Cells(4, Form.DataCenter.WS.Application.Selection.Column).value2))
            '    Cancel()
            'ElseIf Form.DataCenter.GlobalValues.bolCut Then
            '    Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.WS.Cells(4, Form.DataCenter.WS.Application.Selection.Column).value2))
            '    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            '    Cls.UpdateSection(intCutRow, intCutRow)
            '    Cancel()
            'End If

            '_RibbonUtilitis.UpdateUndoButtonsState()

        End Sub








        Shared Function SaveSteps(intCutRow As Integer, lstUCwisePS As List(Of Long), rng As Excel.Range, PlannedStart As Object, Optional InsertAsIndependentUsercase As Boolean = False) As Object
            Dim _ProcessStep As CT.Data.ProcessStep = New Data.ProcessStep
            Dim NextPlannedStart As Date
            ' Pe26s
            Try
                If Form.DataCenter.GlobalValues.bolCopy Then
                    If _ProcessStep.CopyPaste(
                        DataCenter.VehicleConfig.VehiclePe02(rng.Worksheet.Application.Selection.row),
                        DataCenter.VehicleConfig.VehiclePe45(rng.Worksheet.Application.Selection.row),
                        DataCenter.VehicleConfig.VehicleHCID(rng.Worksheet.Application.Selection.row),
                        0,
                        0,
                        If(PlannedStart Is Nothing, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2), CDate(PlannedStart)),
                        lstUCwisePS,
                        CT.Form.DataCenter.ProgramConfig.BuildType,
                        NextPlannedStart,
                        InsertAsIndependentUsercase) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                ElseIf Form.DataCenter.GlobalValues.bolCut Then
                    If _ProcessStep.CutPaste(DataCenter.VehicleConfig.VehiclePe02(intCutRow), DataCenter.VehicleConfig.VehiclePe45(intCutRow),
                                            DataCenter.VehicleConfig.VehiclePe45(rng.Worksheet.Application.Selection.row),
                                            DataCenter.VehicleConfig.VehicleHCID(rng.Worksheet.Application.Selection.row),
                                            0, 0,
                                            If(PlannedStart Is Nothing, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2), CDate(PlannedStart)),
                                            lstUCwisePS,
                                            CT.Form.DataCenter.ProgramConfig.BuildType,
                                            NextPlannedStart,
                                            InsertAsIndependentUsercase) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                End If
                _ProcessStep = Nothing
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Cut/Copy & paste", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Finally
                SaveSteps = NextPlannedStart

            End Try
        End Function
    End Class
End Namespace
