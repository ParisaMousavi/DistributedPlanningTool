Imports System.Data
Imports System.Data.SqlClient
Imports CT.Data.Interfaces

Namespace RigPlan.Segment
    Public Class Leftside
        Inherits CtBaseClass
        Implements Interfaces.LeftInterface

        Public Function GetPlanDataHcIdSpecific(HcId As Integer, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,) Implements LeftInterface.GetPlanDataHcIdSpecific

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A2_VehicleAnd7Tabs_Rig_Specific.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HcId
                    command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
                    command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq

                    _tbAnswer = Nothing
                    _arrayDT = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetPlanDataHcIdSpecific = _arrayDT
            Catch ex As Exception
                '----------------------------------------------------------------
                ' Error classification mechanism
                '----------------------------------------------------------------
                Dim ErrorId As Integer
                Select Case ex.Message
                    Case ex.Message.IndexOf("Permission") >= 0
                        ErrorId = DataCenter.ErrorCenter.Permission
                    Case ex.Message.IndexOf("could not found") >= 0
                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                    Case Else
                        ErrorId = DataCenter.ErrorCenter.TndPlanInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPlanDataHcIdSpecific = Nothing
            End Try

        End Function


        Public Function GetPlanDataHcIdGeneric(HcId As Integer, MainBuildType As String, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object) As String(,) Implements LeftInterface.GetPlanDataHcIdGeneric

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A2_VehicleAnd7Tabs_Rig_Generic.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HcId
                    command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
                    command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq

                    _tbAnswer = Nothing
                    _arrayDT = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetPlanDataHcIdGeneric = _arrayDT
            Catch ex As Exception
                '----------------------------------------------------------------
                ' Error classification mechanism
                '----------------------------------------------------------------
                Dim ErrorId As Integer
                Select Case ex.Message
                    Case ex.Message.IndexOf("Permission") >= 0
                        ErrorId = DataCenter.ErrorCenter.Permission
                    Case ex.Message.IndexOf("could not found") >= 0
                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                    Case Else
                        ErrorId = DataCenter.ErrorCenter.TndPlanInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPlanDataHcIdGeneric = Nothing
            End Try

        End Function

        Protected Overloads Sub ConvertDataTableToStingArray()

            Dim i, j As Integer
            If _tbAnswer IsNot Nothing Then

                ReDim _arrayDT(_tbAnswer.Rows.Count - 1, _tbAnswer.Columns.Count - 1)
                For i = 0 To _tbAnswer.Rows.Count - 1
                    For j = 0 To _tbAnswer.Columns.Count - 1
                        If IsDate(_tbAnswer.Rows(i)(j).ToString()) = True Then
                            '_arrayDT(i, j) = Convert.ToDateTime(DateValue(_tbAnswer.Rows(i)(j).ToString()).ToShortDateString()).ToString("dd-MM-yyyy") 'DateValue(_tbAnswer.Rows(i)(j).ToString()).ToShortDateString()
                            _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                        Else
                            _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                        End If
                    Next j
                Next i
            End If
        End Sub

    End Class
End Namespace