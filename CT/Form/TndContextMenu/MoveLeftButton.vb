
Namespace Form.TndContextMenu
    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class MoveLeftButton
        Public Shared Sub Click()

            Dim myValue As String = InputBox("The count of Calendar days to move: ", "Move Left", "1")
            Dim Cls As New Form.DataCenter.GlobalFunctions
            If CType(If(IsNumeric(myValue), myValue, 0), Integer) > 0 Then
                If Form.DataCenter.ProcessStepConfig.ProcessStepSequence <> 0 Then
                    Dim _PS As CT.Data.ProcessStep = New CT.Data.ProcessStep
                    If _PS.MoveLeft(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.VehicleConfig.VehiclePe45, Form.DataCenter.VehicleConfig.VehicleHCID, Form.DataCenter.ProcessStepConfig.ProcessStepAllocatedUsercase, Form.DataCenter.ProcessStepConfig.ProcessStepSequence, Convert.ToInt16(myValue), Form.DataCenter.ProgramConfig.BuildType) = True Then
                        Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row,,, Form.DataCenter.ProcessStepConfig.ProcessStepPe26)
                    Else
                        System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "MoveLeftButton Click", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                    End If
                Else
                    Dim _ProcessStep As CT.Data.ProcessStep = New CT.Data.ProcessStep
                    If _ProcessStep.MoveLeft(Form.DataCenter.ProgramConfig.pe02, New List(Of Long)(New Long() {Form.DataCenter.VehicleConfig.VehiclePe45}), Form.DataCenter.VehicleConfig.VehicleHCID, CType(myValue, Integer), Form.DataCenter.ProgramConfig.BuildType) = True Then
                        Cls.UpdateSection(Form.DataCenter.GlobalValues.WS.Application.Selection.row, Form.DataCenter.GlobalValues.WS.Application.Selection.row)
                    Else
                        System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "MoveLeftButton Click", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                    End If
                End If
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()

            End If
        End Sub

        Public Shared Sub Click_Multi()

            Dim myValue As String = InputBox("The count of Calendar days to move: ", "Move units left", "1")
            If CType(If(IsNumeric(myValue), myValue, 0), Integer) > 0 Then
                Dim _ProcessStep As CT.Data.ProcessStep = New CT.Data.ProcessStep

                Dim lstVehicles As New List(Of Long), colRows As New Collection
                Dim rng As Excel.Range = Nothing

                For Each rng In Globals.ThisAddIn.Application.Selection.rows
                    Try
                        lstVehicles.Add(Integer.Parse(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.TnDPlan.ToString).Range("C" + rng.Row.ToString()).Value.ToString().Split(";")(3)))
                        colRows.Add(rng.Row)
                    Catch ex As Exception
                    End Try
                Next

                If _ProcessStep.MoveLeft(Form.DataCenter.ProgramConfig.pe02, lstVehicles, Form.DataCenter.VehicleConfig.VehicleHCID, CType(myValue, Integer), Form.DataCenter.ProgramConfig.BuildType) = True Then
                    Dim Cls As New Form.DataCenter.GlobalFunctions
                    Dim intcnt As Integer = 0
                    For intcnt = 1 To colRows.Count
                        Cls.UpdateSection(colRows.Item(intcnt), colRows.Item(intcnt))
                    Next
                Else

                    System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message, "MoveLeftButton Multi click", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                End If
            End If
        End Sub
    End Class
End Namespace