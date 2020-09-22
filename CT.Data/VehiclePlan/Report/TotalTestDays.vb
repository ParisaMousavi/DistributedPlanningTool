Imports System.Data
Imports System.Data.SqlClient

Namespace VehiclePlan.Report
    Public Class TotalTestDays
        Inherits CtBaseClass



        Public Function TotalCountReport1(HCID As Integer, MainBuildType As String) As String(,)
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_VehicleReportI.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableWithColumnToStingArray()

                DataCenter.GlobalValues.message = String.Empty
                TotalCountReport1 = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.Program
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                TotalCountReport1 = Nothing
            End Try

        End Function

        Public Function TotalCountReport2(HCID As Integer, MainBuildType As String) As String(,)
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_VehicleReportII.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableWithColumnToStingArray()

                DataCenter.GlobalValues.message = String.Empty
                TotalCountReport2 = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.Program
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                TotalCountReport2 = Nothing
            End Try

        End Function

        Public Function TotalCountReport3(HCID As Integer, MainBuildType As String) As String(,)
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_VehicleReportIII.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableWithColumnToStingArray()

                DataCenter.GlobalValues.message = String.Empty
                TotalCountReport3 = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.Program
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                TotalCountReport3 = Nothing
            End Try

        End Function

        Public Function TotalCountReport4(HCID As Integer, MainBuildType As String) As String(,)
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_VehicleReportIV.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableWithColumnToStingArray()

                DataCenter.GlobalValues.message = String.Empty
                TotalCountReport4 = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.Program
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                TotalCountReport4 = Nothing
            End Try

        End Function

        Public Function TotalCountReport5(HCID As Integer, MainBuildType As String) As String(,)
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_VehicleReportV.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableWithColumnToStingArray()

                DataCenter.GlobalValues.message = String.Empty
                TotalCountReport5 = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.Program
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                TotalCountReport5 = Nothing
            End Try

        End Function


    End Class
End Namespace
