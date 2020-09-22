
Namespace Form.TndContextMenu
    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class EditProcessStepButton
        Public Shared Sub Click()

            Dim _frmObject As Object

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _frmObject = New frmEdit
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _frmObject = New frmEdit_Rig
            Else
                Exit Sub
            End If
            'Dim _frmEdit As New frmEdit()
            _frmObject.pe26 = Globals.ThisAddIn.Application.ActiveCell.Formula.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim()
            _frmObject.ShowDialog()
            '-------------------------------------------------------------------
            ' Update undo button state
            '-------------------------------------------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

        End Sub

    End Class
End Namespace