
Imports System.Data
Imports System.Data.SqlClient

Namespace VehiclePlan.Segment
    Public Class Header
        Inherits CtBaseClass
        Implements CT.Data.Interfaces.HeaderInterface

        Public Function GetTndPlanHeaderSpecific(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,) Implements Interfaces.HeaderInterface.GetPlanHeaderSpecific
            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing
                    command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.A1_Header_Vehicle_Specific.ToString())

                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
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

                If _tbAnswer.Rows.Count = 0 Then Throw New Exception("Return value for plan header from DB is Empty")

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetTndPlanHeaderSpecific = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.TndPlanHeader_vehicle
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetTndPlanHeaderSpecific = Nothing
            End Try
        End Function

        Public Function GetPlanHeaderGeneric(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,) Implements Interfaces.HeaderInterface.GetPlanHeaderGeneric

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing
                    command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.A1_Header_Vehicle_Generic.ToString())


                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
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

                If _tbAnswer.Rows.Count = 0 Then Throw New Exception("Return value for plan header from DB is Empty")

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetPlanHeaderGeneric = _arrayDT

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
                        ErrorId = DataCenter.ErrorCenter.TndPlanHeader_vehicle
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPlanHeaderGeneric = Nothing
            End Try

        End Function


    End Class
End Namespace
