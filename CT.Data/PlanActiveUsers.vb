
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' table 97
''' </summary>
Public Class PlanActiveUsers
    Inherits CtBaseClass


    Public Function Insert(pe01 As Long, HCID As Integer, MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Try

            If pe01 > 0 And HCID > 0 And MainBuildType <> String.Empty Then

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PlanActiveUserInsert.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_TnDBasicProgram_Fk", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType

                    command.ExecuteNonQuery()
                    transaction.Commit()

                End Using
            End If
            DataCenter.GlobalValues.message = String.Empty
            Insert = True

        Catch ex As Exception
            transaction.Rollback()
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
                    ErrorId = DataCenter.ErrorCenter.PlanActiveUser
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Insert = False

        End Try

    End Function


    Public Function Remove(pe01 As Long, HCID As Integer, MainBuildType As String) As Boolean
        Dim transaction As SqlTransaction = Nothing
        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                conTnd.Open()
                transaction = conTnd.BeginTransaction()

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PlanActiveUserRemove.ToString())
                command.Connection = conTnd
                command.Transaction = transaction
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe01_TnDBasicProgram_Fk", SqlDbType.BigInt, 8).Value = pe01
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType


                command.ExecuteNonQuery()
                transaction.Commit()

            End Using

            DataCenter.GlobalValues.message = String.Empty
            Remove = True

        Catch ex As Exception
            transaction.Rollback()
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
                    ErrorId = DataCenter.ErrorCenter.PlanActiveUser
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Remove = False

        End Try

    End Function


    Public Function SelectAll(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PlanActiveUserSelectAll.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe01_TnDBasicProgram_Fk", SqlDbType.BigInt, 8).Value = pe01
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using


            End Using

            DataCenter.GlobalValues.message = String.Empty
            SelectAll = _tbAnswer

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
                    ErrorId = DataCenter.ErrorCenter.PlanActiveUser
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectAll = Nothing

        End Try

    End Function



End Class
