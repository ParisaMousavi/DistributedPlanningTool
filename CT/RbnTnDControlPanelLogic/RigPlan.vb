
Namespace RbnTnDControlPanelLogic

    Public Class RigPlan
        Private _ErrorMessage As String
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property

        Public Function CovertGeneric2Spesific() As System.Windows.Forms.DialogResult
            Dim _frmPick1stVP As frmPick1stVP = Nothing
            Dim _Plan As New Form.DisplayUtilities.Plan()

            CovertGeneric2Spesific = System.Windows.Forms.DialogResult.Cancel

            Try
                _frmPick1stVP = New frmPick1stVP()

                If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                    CovertGeneric2Spesific = _frmPick1stVP.ShowDialog()

                    If CovertGeneric2Spesific = System.Windows.Forms.DialogResult.OK Then
                        Form.DataCenter.ProgramConfig.IsGeneric = False

                        '--------------------------------------------------------------------------
                        ' Referesh plan
                        '--------------------------------------------------------------------------
                        If Form.DataCenter.ProgramConfig.IsMainPlan = True And Form.DataCenter.ProgramConfig.HCID <> 0 Then
                            If Not _Plan.RefreshPlan(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.IsGeneric, Form.DataCenter.ProgramConfig.IsWithCustomFormatting, Form.DataCenter.ProgramConfig.BuildType) Then Throw New Exception(_Plan.ErrorMessage)
                        End If

                    End If
                End If
                _ErrorMessage = String.Empty
            Catch ex As Exception
                CovertGeneric2Spesific = System.Windows.Forms.DialogResult.None
                _ErrorMessage = ex.Message
            End Try
        End Function


    End Class
End Namespace
