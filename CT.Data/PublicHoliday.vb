Imports System.Data
Imports System.Data.SqlClient

Public Class PublicHoliday
    Inherits CtBaseClass



    Public Function GetGenericPublicHolidays(RegionName As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Standards_GetGenericPublicHolidays.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            GetGenericPublicHolidays = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetGenericPublicHolidays = Nothing
        End Try




    End Function




    Public Function Populate(pe02 As Long, HCID As Integer, MainBuildType As String) As Boolean


        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()



                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_MoveLeftorRight, pe02, Nothing, String.Format("Populate Pubilic Holiday."), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If

                    Dim Command As SqlCommand

                    Command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepPopulateHolidays.ToString())
                    Command.Connection = conTnd
                    Command.Transaction = transaction
                    Command.CommandType = CommandType.StoredProcedure
                    Command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    Command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    Command.Parameters.Add("@ActionId", SqlDbType.BigInt, 8).Value = ActionId

                    Command.ExecuteNonQuery()



                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    Populate = True

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
                            ErrorId = DataCenter.ErrorCenter.PublicHoliday
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    Populate = False

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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Populate = False

        End Try

    End Function


    Public Function Add(HCID As Integer, MainBuildType As String, Regions As String, Country As String, State As String, CityName As String, PublicHolidayName As String, PublicHolidayType As String, PublicHolidayStart As DateTime, PublicHolidayEnd As DateTime, pe83 As Long) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PublicHolidayAdd.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@Regions", SqlDbType.NVarChar, 5).Value = Regions
                    command.Parameters.Add("@Country", SqlDbType.NVarChar, 50).Value = Country
                    command.Parameters.Add("@State", SqlDbType.NVarChar, 50).Value = State
                    command.Parameters.Add("@CityName", SqlDbType.NVarChar, 50).Value = CityName
                    command.Parameters.Add("@PublicHolidayStart", SqlDbType.Date, 3).Value = PublicHolidayStart
                    command.Parameters.Add("@PublicHolidayEnd", SqlDbType.Date, 3).Value = PublicHolidayEnd
                    command.Parameters.Add("@PublicHolidayName", SqlDbType.NVarChar, 50).Value = PublicHolidayName
                    command.Parameters.Add("@PublicHolidayType", SqlDbType.NVarChar, 4).Value = PublicHolidayType
                    command.Parameters.Add("@pe83", SqlDbType.BigInt, 8).Value = pe83

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
                            ErrorId = DataCenter.ErrorCenter.PublicHoliday
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Add = False

        End Try

    End Function


    Public Function Delete(Pe85 As Long) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PublicHolidayDelete.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe85_TnDProgramHolidays_PK", SqlDbType.Int, 4).Value = Pe85


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
                            ErrorId = DataCenter.ErrorCenter.PublicHoliday
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Delete = False

        End Try


    End Function




    Public Function Update(pe85 As Long, Regions As String, Country As String, State As String, CityName As String, PublicHolidayName As String, PublicHolidayType As String, PublicHolidayStart As DateTime, PublicHolidayEnd As DateTime, MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PublicHolidayUpdate.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe85_TnDProgramHolidays_PK", SqlDbType.Int, 4).Value = pe85
                    command.Parameters.Add("@Regions", SqlDbType.NVarChar, 5).Value = Regions
                    command.Parameters.Add("@Country", SqlDbType.NVarChar, 50).Value = Country
                    command.Parameters.Add("@State", SqlDbType.NVarChar, 50).Value = State
                    command.Parameters.Add("@CityName", SqlDbType.NVarChar, 50).Value = CityName
                    command.Parameters.Add("@PublicHolidayStart", SqlDbType.Date, 3).Value = PublicHolidayStart
                    command.Parameters.Add("@PublicHolidayEnd", SqlDbType.Date, 3).Value = PublicHolidayEnd
                    command.Parameters.Add("@PublicHolidayName", SqlDbType.NVarChar, 50).Value = PublicHolidayName
                    command.Parameters.Add("@PublicHolidayType", SqlDbType.NVarChar, 4).Value = PublicHolidayType

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
                            ErrorId = DataCenter.ErrorCenter.PublicHoliday
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Update = False

        End Try


    End Function



    Public Function GetPlanPublicHolidays(HCID As Integer, MainBuildType As String) As DataTable

        Dim _tbAnswer As DataTable = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_GetPlanPublicHolidays.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            GetPlanPublicHolidays = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetPlanPublicHolidays = Nothing
        End Try



    End Function

    Public Function GetPlanPublicHolidaysForHeader(MainBuildType As String, HCID As Integer) As DataTable

        Dim _tbAnswer As DataTable = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_PlanHolidays.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            GetPlanPublicHolidaysForHeader = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetPlanPublicHolidaysForHeader = Nothing
        End Try



    End Function

    Public Function GetAllLocations() As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Standards_GetPublicHolidayLocations.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            GetAllLocations = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.PublicHoliday
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetAllLocations = Nothing
        End Try




    End Function




End Class
