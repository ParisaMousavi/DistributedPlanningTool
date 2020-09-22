Imports System.Data
Imports System.Data.SqlClient

Namespace RigPlan.Segment
    Public Class TestsArea
        Inherits CtBaseClass
        Implements Interfaces.TestAreaInterface

        ''' <summary>
        ''' This methode retrieves only data from database. These data don't have format.
        ''' </summary>
        ''' <param name="HCID"></param>
        ''' <param name="UpperBoundDisplaySeq">The UpperBoundDisplaySeq defined the uper bound if the rows
        ''' which are retrieved. If we need all the Vehicles this value must be NULL/Nothing and if a dedicated vehicles are needed we should define it.</param>
        ''' <param name="LowerBoundDisplaySeq">The LowerBoundDisplaySeq defined the uper bound if the rows
        ''' which are retrieved. If we need all the Vehicles this value must be NULL/Nothing and if a dedicated vehicles are needed we should define it.</param>
        ''' <returns></returns>
        Public Function GetTndAreaDataSpecific(HCID As Integer, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,) Implements Interfaces.TestAreaInterface.GetTndAreaDataSpecific

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A3_TimeLineData_Rig_Specific.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
                    command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq


                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetTndAreaDataSpecific = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.TndPlanArea_Rig
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetTndAreaDataSpecific = Nothing
            End Try

        End Function


        Public Function GetTndAreaDataGeneric(HCID As Integer, MainBuildType As String, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object) As String(,) Implements Interfaces.TestAreaInterface.GetTndAreaDataGeneric

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A3_TimeLineData_Rig_Generic.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
                    command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq


                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetTndAreaDataGeneric = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.TndPlanArea_Rig
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetTndAreaDataGeneric = Nothing
            End Try

        End Function

    End Class
End Namespace
