Imports System.Data
Imports System.Data.SqlClient
Public Class ProgramConfiguration
    Inherits CtBaseClass


    Public Function Add(pe01 As Long, HCID As Integer, ProgramDescription As String, ProgramStatus As String, ProgramreleaseStatus As String, TndReleaseStatus As String, BuildPhases As String, BuildTypes As String, TnDPlanner As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProgramConfigAdd.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildTypes
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@ProgramDescription", SqlDbType.NVarChar, 50).Value = ProgramDescription
                    command.Parameters.Add("@ProgramStatus", SqlDbType.NVarChar, 20).Value = ProgramStatus
                    command.Parameters.Add("@ProgramreleaseStatus", SqlDbType.NVarChar, 20).Value = ProgramreleaseStatus
                    command.Parameters.Add("@TndReleaseStatus", SqlDbType.NVarChar, 50).Value = TndReleaseStatus
                    command.Parameters.Add("@BuildPhases", SqlDbType.NVarChar, 25).Value = BuildPhases
                    command.Parameters.Add("@BuildTypes", SqlDbType.NVarChar, 50).Value = BuildTypes
                    command.Parameters.Add("@TndPlanner", SqlDbType.NVarChar, 50).Value = If(TnDPlanner Is Nothing, DBNull.Value, TnDPlanner)

                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    Add = True

                Catch ex0 As Exception

                    transaction.Rollback()
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
                            ErrorId = DataCenter.ErrorCenter.Program
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
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
                    ErrorId = DataCenter.ErrorCenter.Program
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Add = False

        End Try

    End Function

    Public Function Delete(pe78 As Int32) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProgramConfigDelete.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe78_TnDProgramConfig_PK", SqlDbType.Int, 4).Value = pe78

                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    Delete = True

                Catch ex0 As Exception

                    transaction.Rollback()
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
                            ErrorId = DataCenter.ErrorCenter.Program
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
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
                    ErrorId = DataCenter.ErrorCenter.Program
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Delete = False

        End Try

    End Function

    Public Function Update(pe78 As Long, HCID As Int32, ProgramDescription As String, ProgramStatus As String, ProgramreleaseStatus As String, TndReleaseStatus As String, BuildPhases As String, BuildTypes As String, TnDPlanner As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProgramConfigUpdate.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe78_TnDProgramConfig_PK", SqlDbType.Int, 4).Value = pe78
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildTypes
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@ProgramDescription", SqlDbType.NVarChar, 50).Value = ProgramDescription
                    command.Parameters.Add("@ProgramStatus", SqlDbType.NVarChar, 20).Value = ProgramStatus
                    command.Parameters.Add("@ProgramreleaseStatus", SqlDbType.NVarChar, 20).Value = ProgramreleaseStatus
                    command.Parameters.Add("@TndReleaseStatus", SqlDbType.NVarChar, 50).Value = TndReleaseStatus
                    command.Parameters.Add("@BuildPhases", SqlDbType.NVarChar, 25).Value = BuildPhases
                    command.Parameters.Add("@BuildTypes", SqlDbType.NVarChar, 50).Value = BuildTypes
                    command.Parameters.Add("@TnDPlanner", SqlDbType.NVarChar, 50).Value = If(TnDPlanner Is Nothing, DBNull.Value, TnDPlanner)


                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    Update = True

                Catch ex0 As Exception

                    transaction.Rollback()
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
                            ErrorId = DataCenter.ErrorCenter.Program
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
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
                    ErrorId = DataCenter.ErrorCenter.Program
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Update = False

        End Try

    End Function


    ''' <summary>
    ''' User Form : frmHeaderEdit
    ''' This function is used for Title label at left top corner
    ''' Columns List is : 
    ''' pe01_TnDBasicProgram_FK
    ''' HealthChartId
    ''' ProgramDescription
    ''' ProgramStatus
    ''' ProgramReleaseStatus
    ''' TnDReleaseStatus
    ''' BuildPhases
    ''' BuildTypes
    ''' PairedHealthChartId
    ''' pe78_TnDProgramConfig_PK
    ''' pe04_TnDProgramAuthorization_PK
    ''' </summary>
    ''' <param name="Pe02"></param>
    ''' <param name="HCID"></param>
    ''' <returns></returns>
    Public Function SelectProgramConfigs(Pe02 As Long, HCID As Integer, MainBuildType As String) As DataTable
        Dim _tbAnswer As DataTable = Nothing
        Try
            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProgramConfigSelectByPlan.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe02_TnDprogramDetails_FK", SqlDbType.BigInt, 8).Value = Pe02
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            SelectProgramConfigs = _tbAnswer

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
            SelectProgramConfigs = Nothing
        End Try
    End Function
    Public Enum SelectProgramConfigsColumns
        pe01_TnDBasicProgram_FK
        HealthChartId
        ProgramDescription
        ProgramStatus
        ProgramReleaseStatus
        TnDReleaseStatus
        BuildPhases
        BuildTypes
        PairedHealthChartId
        pe78_TnDProgramConfig_PK
        pe04_TnDProgramAuthorization_PK
        TnDPlanner
        HealthChartName
        pe27_Regions_FK
        pe10_SecurityLevel_FK
    End Enum




End Class
