

Imports System.Data
Imports System.Data.SqlClient
Public Class Usercase
    Inherits CtBaseClass


    Public Function GetAllCdsids(pe01 As Long, HCID As Integer, pe21 As Integer) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_CdsidDvpTeamDedicated.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe01", SqlDbType.BigInt, 8).Value = pe01
                command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID
                'command.Parameters.Add("@DvpTeam", SqlDbType.NVarChar, 50).Value = DvpTeam


                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)


                End Using

            End Using



            DataCenter.GlobalValues.message = String.Empty
            GetAllCdsids = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.Usercase
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetAllCdsids = Nothing
        End Try




    End Function


    Public Function SelectUsercaseDedicated(pe03 As Long, AllocatedUsercaseSequence As Integer) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_UsercaseDedicated.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe03_TnDProgramVehicles_FK", SqlDbType.Int, 4).Value = pe03
                command.Parameters.Add("@AllocatedUsercaseSequence", SqlDbType.Int, 4).Value = AllocatedUsercaseSequence

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)


                End Using

            End Using



            DataCenter.GlobalValues.message = String.Empty
            SelectUsercaseDedicated = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.Usercase
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectUsercaseDedicated = Nothing
        End Try


    End Function



    Public Function EditCDSID(pe02 As Long, pe45 As Long, UsercaseSeq As Integer, CDSID As String, MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        DataCenter.GlobalValues.message = String.Empty

        Try



            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand

                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedProcessStep, pe02, pe45, String.Format(".NET Update CDSID of UsercaseSeq {0} in Unit {1} to {2}.", UsercaseSeq, pe45, CDSID), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If


                    command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UsercaseEditCDSID.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                    command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = UsercaseSeq
                    command.Parameters.Add("@Cdsid", SqlDbType.NVarChar, 16).Value = CDSID
                    command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                    command.ExecuteNonQuery()


                    transaction.Commit()
                    EditCDSID = True

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
                            ErrorId = DataCenter.ErrorCenter.Usercase
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    transaction.Rollback()
                    EditCDSID = False

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
                    ErrorId = DataCenter.ErrorCenter.Usercase
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            EditCDSID = False

        End Try

    End Function



    Public Function EditRemarks(pe02 As Long, pe45 As Long, UsercaseSeq As Integer, Remarks As String, MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        DataCenter.GlobalValues.message = String.Empty

        Try



            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand

                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedProcessStep, pe02, pe45, String.Format(".NET Update Remarks of UsercaseSeq {0} in Unit {1} to {2}.", UsercaseSeq, pe45, Remarks), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If


                    command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UsercaseEditRemarks.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                    command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = UsercaseSeq
                    command.Parameters.Add("@Remarks", SqlDbType.NVarChar, 300).Value = Remarks
                    command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                    command.ExecuteNonQuery()


                    transaction.Commit()
                    EditRemarks = True

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
                            ErrorId = DataCenter.ErrorCenter.Usercase
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    transaction.Rollback()
                    EditRemarks = False

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
                    ErrorId = DataCenter.ErrorCenter.Usercase
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            EditRemarks = False

        End Try

    End Function



    Public Function GetAllUsercases(BuildTypes As String, BuildPhase As String, Carline As String, Region As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_UsercasesInformation.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@BuildTypes", SqlDbType.NVarChar, 100).Value = BuildTypes
                command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 100).Value = BuildPhase
                command.Parameters.Add("@Carline", SqlDbType.NVarChar, 100).Value = Carline
                command.Parameters.Add("@Regions", SqlDbType.NVarChar, 10).Value = Region

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)


                End Using

            End Using



            DataCenter.GlobalValues.message = String.Empty
            GetAllUsercases = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.Usercase
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetAllUsercases = Nothing
        End Try


    End Function



End Class
