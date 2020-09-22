Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Error classification is 300
''' </summary>
Public Class Authorization
    Inherits CtBaseClass



    Public Function SelectAll(MainBuildType As String, HCID As Integer) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Permission_TnDAuthorizationSelectByHCID.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure

                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID


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
                    ErrorId = DataCenter.ErrorCenter.Authorization
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectAll = Nothing
        End Try


    End Function


    Public Function GetPermissionLevel(MainBuildType As String, HCID As Integer, ForGenericPlan As Boolean) As String

        Try
            GetPermissionLevel = String.Empty
            If ForGenericPlan = False Then

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Permission_TnDCurrentUserPermission.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID


                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using
                End Using

                If _tbAnswer.Rows.Count <> 1 Then Throw New Exception("Each user can have only one permission level.")
                GetPermissionLevel = _tbAnswer.Rows(0)("SecurityLevel").ToString

                DataCenter.GlobalValues.message = String.Empty

            ElseIf ForGenericPlan = True Then
                GetPermissionLevel = GetPermissionLevelForGenericPlan(HCID, MainBuildType)
                ' I have deactivated this line
                'If GetPermissionLevel = String.Empty And DataCenter.GlobalValues.message <> String.Empty Then Throw New Exception(Data.DataCenter.GlobalValues.message)
            End If

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
                    ErrorId = DataCenter.ErrorCenter.Authorization
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetPermissionLevel = String.Empty  ' I have chnaged it my self
        End Try

    End Function

    Private Function GetPermissionLevelForGenericPlan(HCID As Integer, BuildType As String) As String

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Permission_TnDCurrentUserPermissionGeneric.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 50).Value = BuildType


                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using
            End Using

            If _tbAnswer.Rows.Count > 1 Then Throw New Exception("Each user can have only one permission level.")
            GetPermissionLevelForGenericPlan = _tbAnswer.Rows(0)("SecurityLevel").ToString


            DataCenter.GlobalValues.message = String.Empty

        Catch ex As Exception
            GetPermissionLevelForGenericPlan = String.Empty
        End Try


    End Function



    Public Function Add(BuildType As String, HCID As Integer, pe10 As Integer, pe27 As Integer, StrCDSID As String, strProgramFunction As String) As Boolean

        Dim transaction As SqlTransaction = Nothing

        DataCenter.GlobalValues.message = String.Empty

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand

                    command = New SqlCommand(DataCenter.StoredProcedures.General.Permission_TnDAuthorizationAdd.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@pe10_SecurityLevel_FK", SqlDbType.Int, 4).Value = pe10
                    command.Parameters.Add("@ProgramFunction", SqlDbType.NVarChar, 50).Value = strProgramFunction ' DBNull.Value
                    command.Parameters.Add("@pe27_Regions_FK", SqlDbType.Int, 4).Value = pe27
                    command.Parameters.Add("@Cdsid", SqlDbType.NVarChar, 16).Value = StrCDSID

                    command.ExecuteNonQuery()


                    transaction.Commit()
                    Add = True

                Catch ex0 As Exception
                    '----------------------------------------------------------------
                    ' Error classification mechanism
                    '----------------------------------------------------------------
                    Dim ErrorId As Integer
                    Select Case ex0.Message
                        Case ex0.Message.IndexOf("Permission") >= 0
                            ErrorId = DataCenter.ErrorCenter.Permission
                        Case ex0.Message.IndexOf("could not found") >= 0
                            ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                        Case Else
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    transaction.Rollback()
                    Add = False

                End Try


            End Using

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
                    ErrorId = DataCenter.ErrorCenter.Authorization
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Add = False

        End Try

    End Function



    Public Function Delete(pe04 As Integer) As Boolean

        Dim transaction As SqlTransaction = Nothing

        DataCenter.GlobalValues.message = String.Empty

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand

                    command = New SqlCommand(DataCenter.StoredProcedures.General.Permission_TnDAuthorizationDelete.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe04_TnDProgramAuthorization_PK", SqlDbType.Int, 4).Value = pe04

                    command.ExecuteNonQuery()


                    transaction.Commit()
                    Delete = True

                Catch ex0 As Exception
                    '----------------------------------------------------------------
                    ' Error classification mechanism
                    '----------------------------------------------------------------
                    Dim ErrorId As Integer
                    Select Case ex0.Message
                        Case ex0.Message.IndexOf("Permission") >= 0
                            ErrorId = DataCenter.ErrorCenter.Permission
                        Case ex0.Message.IndexOf("could not found") >= 0
                            ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                        Case Else
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    transaction.Rollback()
                    Delete = False

                End Try


            End Using

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
                    ErrorId = DataCenter.ErrorCenter.Authorization
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Delete = False

        End Try

    End Function




    Public Function Update(pe04 As Integer, HCID As Integer, pe10 As Integer, pe27 As Integer, StrCDSID As String, strProgramFunction As String) As Boolean

        Dim transaction As SqlTransaction = Nothing

        DataCenter.GlobalValues.message = String.Empty

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand

                    command = New SqlCommand(DataCenter.StoredProcedures.General.Permission_TnDAuthorizationUpdate.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe04_TnDProgramAuthorization_PK", SqlDbType.Int, 4).Value = pe04
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@pe10_SecurityLevel_FK", SqlDbType.Int, 4).Value = pe10
                    command.Parameters.Add("@ProgramFunction", SqlDbType.NVarChar, 50).Value = strProgramFunction 'DBNull.Value
                    command.Parameters.Add("@pe27_Regions_FK", SqlDbType.Int, 4).Value = pe27
                    command.Parameters.Add("@Cdsid", SqlDbType.NVarChar, 16).Value = StrCDSID

                    command.ExecuteNonQuery()

                    transaction.Commit()
                    Update = True

                Catch ex0 As Exception
                    '----------------------------------------------------------------
                    ' Error classification mechanism
                    '----------------------------------------------------------------
                    Dim ErrorId As Integer
                    Select Case ex0.Message
                        Case ex0.Message.IndexOf("Permission") >= 0
                            ErrorId = DataCenter.ErrorCenter.Permission
                        Case ex0.Message.IndexOf("could not found") >= 0
                            ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                        Case Else
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    transaction.Rollback()
                    Update = False

                End Try


            End Using

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
                    ErrorId = DataCenter.ErrorCenter.Authorization
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Update = False

        End Try

    End Function





End Class
