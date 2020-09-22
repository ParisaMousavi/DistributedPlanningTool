Imports System.Data
Imports System.Data.SqlClient

Public Class PlanIndivitualFormatting
    Inherits CtBaseClass

    'Private _tbAnswer As DataTable = Nothing



    Public Function InitialFormat(pe01 As Long, HCID As Integer, MainBuildType As String, FileStatus As String, Optional transaction As SqlTransaction = Nothing, Optional conTnd As SqlConnection = Nothing) As Boolean
        Dim IsRunningLocal As Boolean = Nothing
        Try


            If conTnd Is Nothing Then
                conTnd = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
            End If

            If conTnd.State <> ConnectionState.Open Then
                conTnd.Open() ' it must be here because of BeginTransaction
                IsRunningLocal = True
            End If

            If transaction Is Nothing Then
                transaction = conTnd.BeginTransaction()
                IsRunningLocal = True
            End If

            'conTnd.Open()

            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_FormatInitialGenerate.ToString())
            command.Connection = conTnd
            command.Transaction = transaction

            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus
            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

            If IsRunningLocal = True Then
                command.ExecuteNonQuery()
                transaction.Commit()
                conTnd.Close()
            Else
                command.ExecuteNonQuery()
            End If


            DataCenter.GlobalValues.message = String.Empty
            InitialFormat = True
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
                    ErrorId = DataCenter.ErrorCenter.PlanIndivitualFormatting
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            transaction.Rollback()
            conTnd.Close()
            InitialFormat = False

        End Try

    End Function

    ''' <summary>
    ''' The output if always ONLY one row.
    ''' The value in each cell are as following : Columnname, ColumnBackRGB , ColumnFontRGB, ColumnFilter, ColumnFilterCriteria, ColumnWidth, ColumnSettingCdsid.
    ''' The values are separated with ;.
    ''' Each column of the interface has a column in output. The order of columns in output is the same as the oder of columns on interface.
    ''' </summary>
    ''' <param name="HCID"></param>
    ''' <param name="LoggedinCDSIS">The current logged in user in windows.</param>
    ''' <returns></returns>
    Public Function GetTndPlanHeaderSettings(HCID As Integer, MainBuildType As String, LoggedinCDSIS As String) As DataTable

        Try
            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)


                Dim command As SqlCommand = Nothing
                If MainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                    command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.A4_Vehicle_ColumnFormatSettings.ToString())
                ElseIf MainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                    command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A4_Rig_ColumnFormatSettings.ToString())
                End If

                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@ColumnSettingCdsid", SqlDbType.NVarChar, 32).Value = LoggedinCDSIS


                _tbAnswer = Nothing
                ' This line must have new
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using


                DataCenter.GlobalValues.message = String.Empty
                GetTndPlanHeaderSettings = _tbAnswer

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
                    ErrorId = DataCenter.ErrorCenter.PlanIndivitualFormatting
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetTndPlanHeaderSettings = Nothing
        End Try

    End Function




    ''' <summary>
    ''' This function must only update the setting of the loggedin user. If we want to update a column, which is located in
    ''' the sections which have Section and Header th input parameter is section and Header of the column is placed in section 
    ''' , which doesn't have section the section name like GlobalSection must be sent to this function.
    ''' </summary>
    ''' <param name="HCID"></param>
    ''' <param name="GroupId"> is the Id of the section the same as GlobalSections.</param>
    ''' <param name="Header"></param>
    ''' <param name="Section">In case 2 , 4 , 8 : section = section name of the column or the rest Nothing can be passed to the function.</param>
    ''' <param name="ColumnSettingCdsid"> is the loggedin user.</param>
    ''' <param name="ColumnWidth"></param>
    ''' <returns></returns>
    Public Function UpdateSettings(HCID As Integer, MainBuildType As String, GroupId As Integer, Header As String, Section As String, ColumnSettingCdsid As String, ColumnWidth As Double) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()


                    'In case 2 , 4 , 8 : section = section name of the column
                    Select Case GroupId
                        Case 1
                            Section = "Program Info Static Part (Left Side)"
                        Case 3
                            Section = "Non MFC"
                        Case 5
                            Section = "Program Information"
                        Case 6
                            Section = "Further Basic Information"
                        Case 7
                            Section = "User Shipping Details"
                        Case 8
                            Section = "Update Pack"

                    End Select


                    If GroupId = 0 Or Section = "" Then Throw New Exception("GroupId and Section are not allowed to be null.")

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_FormatUpdateSettings.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@GroupId", SqlDbType.Int, 4).Value = GroupId
                    command.Parameters.Add("@Header", SqlDbType.NVarChar, 150).Value = Header
                    command.Parameters.Add("@Section", SqlDbType.NVarChar, 50).Value = Section
                    command.Parameters.Add("@ColumnSettingCdsid", SqlDbType.NVarChar, 16).Value = ColumnSettingCdsid
                    command.Parameters.Add("@ColumnWidth", SqlDbType.Float, 8).Value = ColumnWidth


                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    UpdateSettings = True

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
                            ErrorId = DataCenter.ErrorCenter.PlanIndivitualFormatting
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    UpdateSettings = False

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
                    ErrorId = DataCenter.ErrorCenter.PlanIndivitualFormatting
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            UpdateSettings = False

        End Try

    End Function


End Class
