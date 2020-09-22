Imports System.Data
Imports System.Data.SqlClient

Public Class MessagePassing
    Inherits CtBaseClass



    Public Function SetAsRead(pe94 As Integer) As Boolean
        Try


            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_MessagePassingDeactive.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe94_TnDPlanMessage_ID", SqlDbType.Int, 4).Value = pe94

                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using
            End Using

            DataCenter.GlobalValues.message = String.Empty
            SetAsRead = True
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
                    ErrorId = DataCenter.ErrorCenter.MessagePassing
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SetAsRead = False


        End Try
    End Function


    Private Function Insert(HCID As Integer, BuildType As String, MessageText As String) As Boolean
        Try


            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_MessagePassingInsert.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                command.Parameters.Add("@MessageText", SqlDbType.NVarChar, 250).Value = MessageText

                command.ExecuteNonQuery()

            End Using

            DataCenter.GlobalValues.message = String.Empty
            Insert = True
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
                    ErrorId = DataCenter.ErrorCenter.MessagePassing
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Insert = False


        End Try
    End Function




    Public Function Insert(HCID As Integer, BuildType As String, MessageText As String, Optional transaction As SqlTransaction = Nothing, Optional conTnd As SqlConnection = Nothing) As Boolean
        Dim IsRunningLocal As Boolean = Nothing
        Try

            If conTnd Is Nothing Then
                conTnd = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                conTnd.Open()
                IsRunningLocal = True
            End If

            If transaction Is Nothing And IsRunningLocal = False Then
                transaction = conTnd.BeginTransaction
                IsRunningLocal = True
            End If

            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_MessagePassingInsert.ToString())
            command.Connection = conTnd

            If IsRunningLocal = False Then command.Transaction = transaction

            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
            command.Parameters.Add("@MessageText", SqlDbType.NVarChar, 250).Value = MessageText

            command.ExecuteNonQuery()

            If IsRunningLocal = True Then
                conTnd.Close()
            End If


            DataCenter.GlobalValues.message = String.Empty
            Insert = True
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
                    ErrorId = DataCenter.ErrorCenter.MessagePassing
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Insert = False

        End Try
    End Function




    Public Function SelectAll(HCID As Integer, BuildType As String) As DataTable
        Try


            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_MessagePassingSelectAll.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType

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
                    ErrorId = DataCenter.ErrorCenter.MessagePassing
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectAll = Nothing


        End Try
    End Function




End Class
