Namespace RbnTnDControlPanelLogic
    Public Class VehiclePlan

        Private _ErrorMessage As String
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
        End Property


        Public Function CovertGeneric2Spesific() As Boolean
            Dim _frmPick1stVP As frmPick1stVP = Nothing
            Dim _Plan As New Form.DisplayUtilities.Plan()

            Try
                _frmPick1stVP = New frmPick1stVP()

                If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                    If _frmPick1stVP.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
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
                Return True
            Catch ex As Exception
                _ErrorMessage = ex.Message
                Return False
            End Try
        End Function



    End Class
End Namespace
