Imports System.Windows.Forms

Public Class frmDeleteVehicle_Rig
    '-----------------------------------------------------------
    ' We have defined these variables here for more
    ' consistancy through deleting a unit
    '-----------------------------------------------------------
    Private _pe03 As Long = 0
    Private _pe02 As Long = 0
    Private _pe45 As Long = 0

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close() 'close the form
    End Sub

    Public Sub cmdDelete_Click(Optional sender As Object = Nothing, Optional e As EventArgs = Nothing) Handles cmdDelete.Click
        'Dim objPro As New Form.DataCenter.ModuleFunction
        Dim cls As New Form.DataCenter.GlobalFunctions
        Try
            '------------------------------------------------
            ' Valodatiion
            '------------------------------------------------
            If Val(lblVehicleID.Text) = 0 Then Throw New Exception("Please select a vehicle to delete.")

            '------------------------------------------------
            ' Remove unit in DB
            '------------------------------------------------
            Dim _DeleteVehicle As Data.RigPlan.Unit = New Data.RigPlan.Unit()
            If _DeleteVehicle.Delete(Form.DataCenter.ProgramConfig.HCID, _pe03, _pe02, _pe45, Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception("Sorry, your changes could not be saved to the database! Database error :- " & Data.DataCenter.GlobalValues.message)

            '------------------------------------------------
            ' Start screen updating
            '------------------------------------------------
            Globals.ThisAddIn.Application.ScreenUpdating = False

            cls.GetResetFilter()
            '------------------------------------------------
            ' Apply changes to EXCEL interface
            '------------------------------------------------
            Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            With Form.DataCenter.GlobalValues.WS
                .Cells(.Application.Selection.Row, "A").EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            End With
            Form.DataCenter.GlobalValues.TotalRow = Form.DataCenter.GlobalValues.TotalRow - 1


            Dim _TndPlanTitle As Form.DisplayUtilities.TndPlanTitle = New Form.DisplayUtilities.TndPlanTitle
            _TndPlanTitle.FillMismatchedQty()


            'Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            '_RibbonUtilitis.UpdateUndoButtonsState()
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MessageBox.Show("Unit deleted successfully!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmDeleteVehicle, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            '-------------------------------
            ' Activate undo button
            '-------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            cls.ReApplyFilter()
            'objPro.sbProtectPlan() 'Already protected in FillMismatchQty Function
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.DisplayAlerts = True
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Me.Close()
        End Try
    End Sub

    Private Sub frmDeleteVehicle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        If Form.DataCenter.ProgramConfig.IsGeneric = True Then
            cmdDelete.Enabled = False
        End If
    End Sub

    Public Sub frmDeleteVehicle_Shown(Optional sender As Object = Nothing, Optional e As EventArgs = Nothing) Handles Me.Shown
        Try
            _pe02 = Form.DataCenter.VehicleConfig.VehiclePe02
            _pe03 = Form.DataCenter.VehicleConfig.VehiclePe03
            _pe45 = Form.DataCenter.VehicleConfig.VehiclePe45

            lblVehicleID.Text = Form.DataCenter.VehicleConfig.VehicleDisPlaySeq
            lblPhase.Text = Form.DataCenter.VehicleConfig.VehiclePhase
            lblType.Text = Form.DataCenter.VehicleConfig.VehicleBuildType
            lblEngine.Text = Form.DataCenter.VehicleConfig.VehicleEngine
            lblEngineType.Text = Form.DataCenter.VehicleConfig.VehicleEngineType
            lblTransmission.Text = Form.DataCenter.VehicleConfig.VehicleTransmission
            lblTransmissionType.Text = Form.DataCenter.VehicleConfig.VehicleTransmissionType
            lblTemaName.Text = Form.DataCenter.VehicleConfig.VehicleTeamName
            lblXCCTEamName.Text = Form.DataCenter.VehicleConfig.VehicleXCCTeamName
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmDeleteVehicle, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    Private Sub frmDeleteVehicle_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F7 Then
            cmdDelete_Click(sender, e)
        End If
    End Sub
End Class
