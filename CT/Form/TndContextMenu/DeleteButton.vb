
Namespace Form.TndContextMenu
    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class DeleteButton
        Public Shared Sub Click()
            Try
                Dim _processStep As CT.Data.ProcessStep = New CT.Data.ProcessStep()
                Dim Cls As New Form.DataCenter.GlobalFunctions
                Dim rngT As Excel.Range
                Dim rng As Excel.Range
                Dim lstPS As New List(Of Long)
                Dim lngPrevPS As Long = 0

                If Form.DataCenter.GlobalValues.bolUserCaseSelected = True Or Form.DataCenter.GlobalValues.bolSelAll = True Then
                    If Form.DataCenter.GlobalValues.bolUserCaseSelected = True Then
                        rng = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strUserCaseSelected)
                    Else
                        rng = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strSelAllAddress)
                    End If

                    For Each rngT In rng.Cells
                        If rngT.Formula IsNot Nothing And rngT.Formula <> "" And rngT.Formula <> "-" Then
                            lstPS.Add(CLng(rngT.Formula.ToString.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim()))
                        End If
                    Next
                    If _processStep.Delete(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.VehicleConfig.VehicleHCID, Form.DataCenter.VehicleConfig.VehiclePe45, lstPS, Form.DataCenter.ProgramConfig.BuildType) = True Then
                        Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                    Else
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    Cancel()
                Else
                    If _processStep.Delete(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.VehicleConfig.VehicleHCID, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.ProcessStepConfig.ProcessStepPe26, Form.DataCenter.ProgramConfig.BuildType) = True Then
                        Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                    Else
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    Cancel()
                End If
            Catch ex As Exception
                Cancel()
                System.Windows.Forms.MessageBox.Show(ex.Message, "Delete Process Step click", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Finally
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()

            End Try

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
    End Class
End Namespace
