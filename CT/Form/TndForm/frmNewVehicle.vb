Imports System.Windows.Forms

Public Class frmNewVehicle

    Private Sub frmNewVehicle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetCboBuildType()
        lblBuildPhase.Text = Form.DataCenter.ProgramConfig.BuildPhase
        lblHCID.Text = If(Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString, Form.DataCenter.ProgramConfig.MainPlanHCID, Form.DataCenter.ProgramConfig.HCID)  ' In this case it's only for displaying
        lblHCName.Text = Form.DataCenter.ProgramConfig.HCIDName
        If Form.DataCenter.ProgramConfig.IsGeneric = True Then
            btnAddUnit.Enabled = False
        End If
    End Sub

    Private Sub SetCboBuildType()
        cboBuildType.DataSource = [Enum].GetValues(GetType(CT.Data.DataCenter.BuildType))
        cboBuildType.SelectedItem = CT.Data.DataCenter.BuildType.Vehicle
    End Sub

    Private Sub btnAddUnit_Click(sender As Object, e As EventArgs) Handles btnAddUnit.Click
        Dim Cls As New Form.DataCenter.GlobalFunctions
        Try
            Dim _Modfunc As New Form.DataCenter.ModuleFunction
            If Form.DataCenter.ProgramConfig.pe02 = 0 Then Throw New Exception("Operation not allowed! Please open a plan.")

            If Form.DataCenter.ProgramConfig.IsGeneric Then Throw New Exception("This operation is not allowed in 'Generic' plan. Only allowed in 'Specific' plan.")

            Dim _Unit As CT.Data.VehiclePlan.Unit = New Data.VehiclePlan.Unit

            Me.Cursor = Cursors.AppStarting

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

            Dim pe03, pe45 As Int64
            Dim intRow, GenericSplitRowNumber As Integer

            For i As Int16 = 1 To Val(txtCounts.Text)
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                If Not _Unit.AddUnit(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.BuildPhase, Form.DataCenter.ProgramConfig.BuildType, cboBuildType.Text, Form.DataCenter.ProgramConfig.HCID, pe03_ID:=pe03, pe45_ID:=pe45, GenericSplitRowNumber:=GenericSplitRowNumber) Then
                    Throw New Exception("Add Unit not successful." + CT.Data.DataCenter.GlobalValues.message)
                Else
                    Cls.GetResetFilter()

                    With Form.DataCenter.GlobalValues.WS
                        intRow = Form.DataCenter.ProgramConfig.LastRow + 1
                        Form.DataCenter.GlobalValues.TotalRow = Form.DataCenter.GlobalValues.TotalRow + 1
                        .Range("B" & intRow - 1).EntireRow.Copy()
                        .Range("B" & intRow).EntireRow.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats)
                        .Range(.Cells(intRow, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn), .Cells(intRow, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Interior.Color = 16777215 'xlNone
                        .Cells(intRow, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).value = GenericSplitRowNumber
                    End With
                End If
                Cls.UpdateSection(intRow, intRow, False)
            Next

            Form.DataCenter.GlobalValues.WS.Cells(7, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).Select

            Dim _TndPlanTitle As Form.DisplayUtilities.TndPlanTitle = New Form.DisplayUtilities.TndPlanTitle

            Try
                _TndPlanTitle.FillMismatchedQty()
                Cls.ReApplyFilter()
            Catch ex As Exception
            End Try
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            MessageBox.Show("Add Unit successful", "Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information, System.Windows.Forms.MessageBoxDefaultButton.Button1)

        Catch ex As Exception
            Cls.ReApplyFilter()
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmNewVehicle, ex.Message), "Add Unit", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            Me.Cursor = Cursors.Default
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        End Try

        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmNewVehicle_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            btnAddUnit_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            cboBuildType.Focus()
        End If
    End Sub

    Private Sub txtCounts_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCounts.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class