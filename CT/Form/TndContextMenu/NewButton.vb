

Namespace Form.TndContextMenu

    ''' <summary>
    ''' Each Button of the context menu has a module. The Module should have at least
    ''' one public Click sub
    ''' </summary>
    Friend NotInheritable Class NewButton
        'Public Shared _frmNew As frmNew
        Public Shared _frmNew As Object
        Public Shared Sub ClicK(strAddress As String)


            Dim colPS As New Collection
            Dim intcnt As Integer = 0, rngFind As Excel.Range = Nothing, intRow As Integer = 0
            Dim m_stAddress As String = ""
            Dim intFCol As Integer, intLCol As Integer


            Try


                With Form.DataCenter.GlobalValues.WS
                    Form.DisplayUtilities.Utilities.FindFLCols(.Range(strAddress).Row, intFCol, intLCol)
                    If Globals.ThisAddIn.Application.Selection.Column < intFCol Then Throw New Exception("Sorry, you cannot add process steps before Build, Fit 4 Test & Sign-Off.")

                End With


                If Form.DataCenter.GlobalValues.WS.Range(strAddress).Cells.Count > 1 Then
                    With Form.DataCenter.GlobalValues.WS
                        intRow = .Range(strAddress).Row
                        With .Range(strAddress)
                            rngFind = .Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                            If rngFind IsNot Nothing Then
                                m_stAddress = rngFind.Address
                                Do
                                    colPS.Add(rngFind.Formula & "~" & CDate(Form.DataCenter.GlobalValues.WS.Cells(4, rngFind.Column).value2))
                                    rngFind = .FindNext(rngFind)
                                Loop While Not rngFind Is Nothing And rngFind.Address <> m_stAddress
                            End If
                        End With
                    End With
                End If

                If colPS.Count > 1 Then Throw New Exception("Sorry, you have selected multiple process steps. Please select only one process step to do add operation.")

                If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    _frmNew = New frmNew
                ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    _frmNew = New frmNew_Rig
                Else
                    Exit Sub
                End If
                _frmNew.ShowMe()

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "New Process Step", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
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