Imports System.ComponentModel
Imports System.Windows.Forms

Public Class frmMoveVehiclePosition


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub frmMoveVehiclePosition_Load(sender As Object, e As EventArgs) Handles Me.Load
        lblVehicleID.Text = Form.DataCenter.VehicleConfig.VehicleDisPlaySeq
        lblPhase.Text = Form.DataCenter.VehicleConfig.VehiclePhase
        lblXCCTeamName.Text = Form.DataCenter.VehicleConfig.VehicleXCCTeamName
        lblType.Text = Form.DataCenter.VehicleConfig.VehicleBuildType
        lblEngine.Text = Form.DataCenter.VehicleConfig.VehicleEngine
        lblEngineType.Text = Form.DataCenter.VehicleConfig.VehicleEngineType
        lblTransmission.Text = Form.DataCenter.VehicleConfig.VehicleTransmission
        lblTransmissionType.Text = Form.DataCenter.VehicleConfig.VehicleTransmissionType
        lblTeamName.Text = Form.DataCenter.VehicleConfig.VehicleTeamName
        txtMovePositionRank.Text = ""
        If Form.DataCenter.ProgramConfig.IsGeneric = True Then
            btnMove.Enabled = False
        End If
    End Sub

    Private Sub btnMove_Click(sender As Object, e As EventArgs) Handles btnMove.Click

        Dim rngFind As Excel.Range, IntCurrentRow As Integer, intFutureRow As Integer
        Dim intStRow As Integer, intEdRow As Integer
        Dim Cls As New Form.DataCenter.GlobalFunctions
        Dim objPro As New Form.DataCenter.ModuleFunction


        Try
            If Val(txtMovePositionRank.Text) = 0 Or Val(txtMovePositionRank.Text) = Val(lblVehicleID.Text) Then Throw New Exception("Please enter a valid position rank to move!")


            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual

            Cls.GetResetFilter()

            With Form.DataCenter.GlobalValues.WS
                .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)

                rngFind = Nothing
                rngFind = .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn.Find(txtMovePositionRank.Text, .Cells(4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
                If rngFind Is Nothing Then Throw New Exception("000Sorry, the position rank you are trying to move the vehicle is not possible because the destination vehicle ID doesn't exist. Please try again with a Vehicle ID that exists in the plan.")
                intFutureRow = rngFind.Row
                rngFind = Nothing
                rngFind = .Cells(1, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).entirecolumn.Find(lblVehicleID.Text, .Cells(4, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column), Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
                If rngFind Is Nothing Then Throw New Exception("000Sorry, the position rank you are trying to move the vehicle is not possible because the source vehicle ID doesn't exist. Please try again with a Vehicle ID that exists in the plan.")

                IntCurrentRow = rngFind.Row

                If intFutureRow > IntCurrentRow Then
                    intStRow = IntCurrentRow
                    intEdRow = intFutureRow
                Else
                    intStRow = intFutureRow
                    intEdRow = IntCurrentRow
                End If

                '-------------------------------------------
                ' Apply changes to DB 
                '-------------------------------------------
                Dim _SequenceResult As System.Data.DataTable = Nothing
                Dim _Unit As New Data.VehiclePlan.Unit
                _SequenceResult = _Unit.ChangeBuildSequence(Form.DataCenter.ProgramConfig.pe02, lblVehicleID.Text, txtMovePositionRank.Text, Form.DataCenter.ProgramConfig.BuildType)
                '-------------------------------------------
                ' Validate DB result
                '-------------------------------------------    
                If _SequenceResult Is Nothing Then Throw New Exception("Sorry, your changes could not be saved to the database! Database error :- " & Data.DataCenter.GlobalValues.message)

                '-------------------------------------------
                ' Apply changes to excel 
                '-------------------------------------------
                .Rows(IntCurrentRow).Cut
                If IntCurrentRow > intFutureRow Then
                    .Rows(intFutureRow).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
                Else
                    .Rows(intFutureRow + 1).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
                End If


                For i As Int16 = 5 To _SequenceResult.Rows.Count + 4
                    .Cells(i, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_ID_Column).Value2 = _SequenceResult.Rows(i - 5).Item(0).ToString()
                Next
            End With

            MessageBox.Show("Unit sequence changed successfully!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)


        Catch ex As Exception
            Cls.ReApplyFilter()
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            objPro.sbProtectPlan()

            If ex.Message.IndexOf("000") = 0 Then
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmMoveVehiclePosition, ex.Message.Substring(3)), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            Else
                MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmMoveVehiclePosition, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Finally
            '-------------------------------
            ' Activate undo button
            '-------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            Cls.ReApplyFilter()
            objPro.sbProtectPlan()
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Me.Close()

        End Try
    End Sub

    Private Sub frmMoveVehiclePosition_Validating(sender As Object, e As CancelEventArgs) Handles Me.Validating
        Try
            If txtMovePositionRank.Text = 0 Then Throw New Exception("Invalid rank: 0")
            If txtMovePositionRank.Text = lblVehicleID.Text Then Throw New Exception("Input identical to current position.")
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmMoveVehiclePosition, ex.Message), "Move unit position", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            e.Cancel = True
        End Try
    End Sub

    Private Sub frmMoveVehiclePosition_Validated(sender As Object, e As EventArgs) Handles Me.Validated
        '@Ramesh: Implement Worksheet interaction
    End Sub

    Private Sub txtMovePositionRank_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMovePositionRank.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtMovePositionRank_KeyDown(sender As Object, e As KeyEventArgs) Handles txtMovePositionRank.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            btnMove_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            txtMovePositionRank.Focus()
        End If
    End Sub
End Class

