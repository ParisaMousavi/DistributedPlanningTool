Namespace Form.TndContextMenu


    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class CutButton
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
                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                rng.Cut()
                Form.DataCenter.GlobalValues.bolCut = True
                'Dim _obj As New Form.DataCenter.ModuleFunction
                '_obj.sbProtectPlan()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Cut Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
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