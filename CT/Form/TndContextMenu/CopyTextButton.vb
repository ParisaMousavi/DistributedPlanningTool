
Namespace Form.TndContextMenu


    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class CopyTextButton

        Public Shared Sub Click(strAddress As String)
            Try
                System.Windows.Forms.Clipboard.SetText(Form.DataCenter.GlobalValues.WS.Range(strAddress).Cells(1).formula.ToString.Split("""")(1))
                MsgBox("Formula copied to clipboard.", vbInformation + vbOKOnly)
            Catch ex As Exception
            Finally
                '-------------------------------------------------------------------
                ' Update undo button state
                '-------------------------------------------------------------------
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()

            End Try
        End Sub

    End Class
End Namespace
