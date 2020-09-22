Imports System.Data
Imports System.Data.SqlClient

Namespace RigPlan.SevenTabs
    Public Class NonMfcSpecification
        Inherits CtBaseClass
        Implements CT.Data.Interfaces.NonMfcInterface

        Public Function GetPlanData(pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,) Implements Interfaces.NonMfcInterface.GetPlanData

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing
                    command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A2_VehicleAnd7Tabs_Rig_Specific_NonMfcPartial.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = pe02
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
                GetPlanData = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPlanData = Nothing
            End Try

        End Function

        Public Function GetTndPlanHeader(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,) Implements Interfaces.NonMfcInterface.GetTndPlanHeader

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing

                    command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A1_Header_Rig_Specific_NonMfcPartial.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 4).Value = BuildPhase
                    command.Parameters.Add("@BuildTypes", SqlDbType.NVarChar, 10).Value = BuildType

                    _tbAnswer = Nothing
                    _arrayDT = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetTndPlanHeader = _arrayDT

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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetTndPlanHeader = Nothing
            End Try

        End Function


    End Class
End Namespace
