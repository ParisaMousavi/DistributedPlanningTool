

Namespace Form.TndContextMenu


    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class EditUsercaseButton
        Public Shared Sub click()
            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                Dim _frmEditusercase As New frmEditUsercase()
                _frmEditusercase.AllocatedUsercaseSequence = Globals.ThisAddIn.Application.ActiveCell.Formula.Split(";")(1)
                _frmEditusercase.ShowDialog()
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                Dim _frmEditusercase As New frmEditUsercase_Rig()
                _frmEditusercase.AllocatedUsercaseSequence = Globals.ThisAddIn.Application.ActiveCell.Formula.Split(";")(1)
                _frmEditusercase.ShowDialog()
            End If
            '-------------------------------------------------------------------
            ' Update undo button state
            '-------------------------------------------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

        End Sub
    End Class
End Namespace
