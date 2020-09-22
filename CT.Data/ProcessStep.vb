Imports System.Data
Imports System.Data.SqlClient

Public Class ProcessStep
    Inherits CtBaseClass


    Private Enum ColumnName As Integer
        pe26_SpecificVehicleUsercases_PK = 0
        Usercase = 1
        ProcessStepName = 2
        FacilityName = 3
        FacilityLocation = 4
        FacilityCbg = 5
        SubFacilityName = 6
        PlannedStart = 7
        PlannedEnd = 8
        Cdsid = 9
        Remarks = 10
        Duration = 11
        WorkingDays = 12
        TeamName = 13
        XCCTeamName = 14
        GlobalDVP = 15
        AllocatedUsercaseSeq = 16
        ProcessStepSeq = 17
        PublicHolidayFlag = 18
    End Enum

    Dim _pe26 As Long = Nothing
    Public ReadOnly Property pe26 As Long
        Get
            pe26 = _pe26
        End Get
    End Property


    Dim _FacilityCbg As String
    Public ReadOnly Property FacilityCbg As String
        Get
            FacilityCbg = _FacilityCbg
        End Get
    End Property

    Dim _FacilityLocation As String
    Public ReadOnly Property FacilityLocation As String
        Get
            FacilityLocation = _FacilityLocation
        End Get
    End Property


    Dim _FacilityName As String
    Public ReadOnly Property FacilityName As String
        Get
            FacilityName = _FacilityName
        End Get
    End Property

    Dim _SubFacilityName As String
    Public ReadOnly Property SubFacilityName As String
        Get
            SubFacilityName = _SubFacilityName
        End Get
    End Property

    Dim _PlannedStart As Date
    Public ReadOnly Property PlannedStart As Date
        Get
            PlannedStart = _PlannedStart
        End Get
    End Property

    Dim _PlannedEnd As Date
    Public ReadOnly Property PlannedEnd As Date
        Get
            PlannedEnd = _PlannedEnd
        End Get
    End Property


    Dim _WorkingDays As Integer
    Public ReadOnly Property WorkingDays As Integer
        Get
            WorkingDays = _WorkingDays
        End Get
    End Property

    Dim _Duration As Integer
    Public ReadOnly Property Duration As Integer
        Get
            Duration = _Duration
        End Get
    End Property

    Dim _Usercase As String
    Public ReadOnly Property Usercase() As String
        Get
            Usercase = _Usercase
        End Get
    End Property

    Dim _ProcessStepName As String
    Public ReadOnly Property ProcessStepName() As String
        Get
            ProcessStepName = _ProcessStepName
        End Get
    End Property

    Dim _Cdsid As String
    Public ReadOnly Property Cdsid() As String
        Get
            Cdsid = _Cdsid
        End Get
    End Property

    Dim _Remarks As String
    Public ReadOnly Property Remarks() As String
        Get
            Remarks = _Remarks
        End Get
    End Property


    Dim _TeamName As String
    Public ReadOnly Property TeamName() As String
        Get
            TeamName = _TeamName
        End Get
    End Property

    Dim _XCCTeamName As String
    Public ReadOnly Property XCCTeamName() As String
        Get
            XCCTeamName = _XCCTeamName
        End Get
    End Property

    Dim _GlobalDVP As String
    Public ReadOnly Property GlobalDVP() As String
        Get
            GlobalDVP = _GlobalDVP
        End Get
    End Property


    Dim _AllocatedUsercaseSeq As String
    Public ReadOnly Property AllocatedUsercaseSeq() As String
        Get
            AllocatedUsercaseSeq = _AllocatedUsercaseSeq
        End Get
    End Property

    Dim _ProcessStepSeq As String
    Public ReadOnly Property ProcessStepSeq() As String
        Get
            ProcessStepSeq = _ProcessStepSeq
        End Get
    End Property

    Private _IsWithHoliday As Boolean
    Public ReadOnly Property IsWithHoliday() As Boolean
        Get
            Return _IsWithHoliday
        End Get
    End Property




    Public Function MoveLeft(pe02 As Long, pe45 As Long, HCID As Integer, SelectedAllocatedUsercaseSeq As Integer, SelectedProcessStepSeq As Integer, DayCount As Integer, MainBuildType As String) As Boolean


        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()


                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_MoveLeftorRight, pe02, pe45, String.Format("vehicle pe45s: {0} is moved to left for {1} days.", pe45, DayCount.ToString), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If


                    '-------------------------------------------------------------
                    '  first Process step must be get
                    '-------------------------------------------------------------
                    Dim _ProcessStep As ProcessStep = New ProcessStep
                    If _ProcessStep.SelectProcessStepDedicated(pe45, SelectedAllocatedUsercaseSeq, SelectedProcessStepSeq, transaction, conTnd) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                    Dim Command As SqlCommand

                    Command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                    Command.Connection = conTnd
                    Command.Transaction = transaction
                    Command.CommandType = CommandType.StoredProcedure
                    Command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    Command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    Command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = _ProcessStep.pe26
                    Command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                    Command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = _ProcessStep.AllocatedUsercaseSeq
                    Command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = _ProcessStep.ProcessStepSeq
                    Command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = _ProcessStep.PlannedStart.AddDays(DayCount * -1)
                    'command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = New Object() ' MUST BE NULL
                    Command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = _ProcessStep.Duration
                    Command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = _ProcessStep.WorkingDays
                    Command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(_ProcessStep.IsWithHoliday = True, 1, 0)

                    Command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                    Command.ExecuteNonQuery()


                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    MoveLeft = True

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
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    MoveLeft = False

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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)

            MoveLeft = False

        End Try

    End Function




    Public Function MoveRight(pe02 As Long, pe45 As Long, HCID As Integer, SelectedAllocatedUsercaseSeq As Integer, SelectedProcessStepSeq As Integer, DayCount As Integer, MainBuildType As String) As Boolean


        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()



                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_MoveLeftorRight, pe02, pe45, String.Format("vehicle pe45s: {0} is moved to right for {1} days.", pe45, DayCount.ToString), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If

                    '-------------------------------------------------------------
                    ' Get first Process step must be get
                    '-------------------------------------------------------------
                    Dim _ProcessStep As ProcessStep = New ProcessStep
                    If _ProcessStep.SelectProcessStepDedicated(pe45, SelectedAllocatedUsercaseSeq, SelectedProcessStepSeq, transaction, conTnd) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


                    Dim Command As SqlCommand

                    Command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                    Command.Connection = conTnd
                    Command.Transaction = transaction
                    Command.CommandType = CommandType.StoredProcedure
                    Command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    Command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    Command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = _ProcessStep.pe26
                    Command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                    Command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = _ProcessStep.AllocatedUsercaseSeq
                    Command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = _ProcessStep.ProcessStepSeq
                    Command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = _ProcessStep.PlannedStart.AddDays(DayCount)
                    Command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  ' MUST BE NULL
                    Command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = _ProcessStep.Duration
                    Command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = _ProcessStep.WorkingDays
                    Command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(_ProcessStep.IsWithHoliday = True, 1, 0)
                    Command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                    Command.ExecuteNonQuery()



                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    MoveRight = True

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
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    MoveRight = False

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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            MoveRight = False

        End Try

    End Function



    Public Function Add(pe02 As Long, pe45 As Long, HCID As Integer,
                        AllocatedUsercaseSeq As Integer, ProcessStepSequence As Integer, ProcessStepList As DataTable, MainBuildType As String, Optional InsertAsIndependentUsercase As Boolean = False) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1
        Dim CounterCheck As Boolean = False
        Dim NewAllocatedUsercaseSeq As Integer = -1
        Dim NewProcessStepSequenceMin As Integer = -1
        Dim NewProcessStepSequenceMax As Integer = -1


        Try
            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_NewProcessStep, pe02, pe45, String.Format("Add ProcessStep: {0} to usercase sequence : {1}.", "ProcessStepName", AllocatedUsercaseSeq), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If


                    'pe39 ,
                    'The columns, which must be in input table and table must be ordered by Process step Sequence

                    'Selected_PlannedStart ( only first row must have value ) , 
                    'PlannedEnd ( Can be empty ) , 
                    'Duration ( must have value ) , 
                    'WorkingDays ( must have value ) 


                    Dim command As SqlCommand
                    For i As Int16 = 0 To ProcessStepList.Rows.Count - 1
                        If ProcessStepList.Rows(i).Item("Select") <> "True" Then Continue For

                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepAddNewInsert.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType

                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@pe39_SlotFacilityMatching_FK", SqlDbType.BigInt, 4).Value = ProcessStepList.Rows(i).Item("pe39_SlotFacilityMatching_PK")
                        command.Parameters.Add("@pe26_SpecificvehicleUsercases_ID", SqlDbType.BigInt, 4).Value = DBNull.Value  'This value must be Null in new function
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = If(_tbAnswer Is Nothing, AllocatedUsercaseSeq, Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq")))
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = If(_tbAnswer Is Nothing, ProcessStepSequence, Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence")) + 1)
                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = If(_tbAnswer Is Nothing, Date.Parse(ProcessStepList.Rows(i).Item("PlannedStart").ToString), Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")).AddDays(1))
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(ProcessStepList.Rows(i).Item("Duration").ToString)
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(ProcessStepList.Rows(i).Item("WorkingDays").ToString)
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        '-----------------------------------------------------------------------------
                        ' the sequence is important
                        '-----------------------------------------------------------------------------
                        _tbAnswer = Nothing
                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using


                        If _tbAnswer.Rows.Count < 1 Then Throw New Exception("The Process Step was not inserted.")


                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditData.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe26_SpecificVehicleUsercases_PK", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                        command.Parameters.Add("@Cdsid", SqlDbType.NVarChar, 16).Value = If(Len(ProcessStepList.Rows(i).Item("CDSID")) <= 0, DBNull.Value, ProcessStepList.Rows(i).Item("CDSID"))
                        command.Parameters.Add("@Remarks", SqlDbType.NVarChar, 300).Value = ProcessStepList.Rows(i).Item("Remarks")
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteNonQuery()


                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))

                        '------------------------------------------------------------------------------------
                        ' For inserting new independet Usercase
                        '------------------------------------------------------------------------------------
                        If NewAllocatedUsercaseSeq = -1 Then NewAllocatedUsercaseSeq = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))

                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))

                        '------------------------------------------------------------------------------------
                        ' For inserting new independet Usercase
                        '------------------------------------------------------------------------------------
                        If NewProcessStepSequenceMin = -1 Then NewProcessStepSequenceMin = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))

                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = Date.Parse(_tbAnswer.Rows(0)("PlannedStart"))
                        command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  ' This value must not be passed to DB
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("Duration"))
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("WorkingDays"))
                        command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsDBNull(_tbAnswer.Rows(0)("PublicHolidayFlag")) = True, 0, If(_tbAnswer.Rows(0)("PublicHolidayFlag") = True, 1, 0))
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using
                        CounterCheck = True
                    Next

                    '------------------------------------------------------------------------------------
                    ' For inserting new independet Usercase
                    '------------------------------------------------------------------------------------
                    NewProcessStepSequenceMax = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))


                    If CounterCheck = False Then Throw New Exception("No ProcessStep is added to DB")

                    If InsertAsIndependentUsercase = True Then

                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepAddNewInsertSequence.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = NewAllocatedUsercaseSeq
                        command.Parameters.Add("@ProcessStepSeqMin", SqlDbType.Int, 4).Value = NewProcessStepSequenceMin
                        command.Parameters.Add("@ProcessStepSeqMax", SqlDbType.Int, 4).Value = NewProcessStepSequenceMax
                        command.Parameters.Add("@ActionId", SqlDbType.Int, 4).Value = ActionId

                        command.ExecuteNonQuery()

                    End If

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
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Add = Nothing

        End Try

    End Function

    Public Function Edit(pe02 As Long, pe26 As Long, pe45 As Long, HCID As Integer, UsercaseSeq As Integer, ProcessStepSeq As Integer, PlannedStart As Object, PlannedEnd As Object, Duration As Object, WorkingDays As Integer,
                         Cdsid As String, Remarks As String, FacilityName As String, FacilityLocation As String, FacilityCbg As String, SubFacilityName As String, IsWithHoliday As Boolean, MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        DataCenter.GlobalValues.message = String.Empty

        Try

            SelectProcessStepDedicated(pe26, False)

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    SelectProcessStepDedicated(pe26, False)

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand

                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedProcessStep, pe02, pe45, String.Format(".Net updated Process Step {0} of Usercase {1}.", Me.ProcessStepName, Me.Usercase), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If




                    If PlannedStart IsNot Nothing Or PlannedEnd IsNot Nothing Or Duration IsNot Nothing Then
                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = pe26
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = UsercaseSeq
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = ProcessStepSeq

                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = If(PlannedStart Is Nothing, DBNull.Value, Date.Parse(PlannedStart))
                        command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = If(PlannedEnd Is Nothing, DBNull.Value, Date.Parse(PlannedEnd))
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = If(Duration Is Nothing, DBNull.Value, Integer.Parse(Duration))

                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = WorkingDays
                        command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsWithHoliday = True, 1, 0)
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using

                        CT.Data.DataCenter.GlobalValues.message = String.Empty
                        If _tbAnswer IsNot Nothing Then

                            If _tbAnswer.Rows.Count > 0 Then


                                If PlannedStart IsNot Nothing And PlannedEnd IsNot Nothing Then
                                    If Date.Parse(_tbAnswer.Rows(0)("PlannedStart")) <> PlannedStart Then
                                        CT.Data.DataCenter.GlobalValues.message = CT.Data.DataCenter.GlobalValues.message + String.Format("Selected PlannedStart {0:dd-MM-yyyy} has been changed to {1:dd-MM-yyyy}.", PlannedStart, _tbAnswer.Rows(0)("PlannedStart"))
                                    ElseIf Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")) <> PlannedEnd Then
                                        CT.Data.DataCenter.GlobalValues.message = CT.Data.DataCenter.GlobalValues.message + String.Format("Selected PlannedEnd {0:dd-MM-yyyy} has been changed to {1:dd-MM-yyyy}.", PlannedEnd, _tbAnswer.Rows(0)("PlannedEnd"))
                                    ElseIf Date.Parse(_tbAnswer.Rows(0)("PlannedStart")) <> PlannedStart Or Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")) <> PlannedEnd Then
                                        CT.Data.DataCenter.GlobalValues.message = CT.Data.DataCenter.GlobalValues.message + String.Format("Selected PlannedEnd {0:dd-MM-yyyy} has been changed to {1:dd-MM-yyyy}. Selected PlannedStart {2:dd-MM-yyyy} has been changed to {3:dd-MM-yyyy}.", PlannedEnd, _tbAnswer.Rows(0)("PlannedEnd"), PlannedStart, _tbAnswer.Rows(0)("PlannedStart"))
                                    End If
                                ElseIf PlannedStart IsNot Nothing And Duration IsNot Nothing Then
                                    If Date.Parse(_tbAnswer.Rows(0)("PlannedStart")) <> PlannedStart Then
                                        CT.Data.DataCenter.GlobalValues.message = CT.Data.DataCenter.GlobalValues.message + String.Format("Selected PlannedStart {0:dd-MM-yyyy} has been changed to {1:dd-MM-yyyy}.", PlannedStart, _tbAnswer.Rows(0)("PlannedStart"))
                                    End If
                                End If


                            End If


                        End If



                    End If


                    command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditData.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe26_SpecificVehicleUsercases_PK", SqlDbType.BigInt, 8).Value = pe26
                    command.Parameters.Add("@Cdsid", SqlDbType.NVarChar, 16).Value = IIf(Cdsid Is Nothing, DBNull.Value, Cdsid)
                    command.Parameters.Add("@Remarks", SqlDbType.NVarChar, 300).Value = IIf(Remarks Is Nothing, DBNull.Value, Remarks)

                    command.Parameters.Add("@FacilityName", SqlDbType.NVarChar, 50).Value = If(FacilityName Is Nothing, DBNull.Value, FacilityName)
                    command.Parameters.Add("@FacilityLocation", SqlDbType.NVarChar, 50).Value = If(FacilityLocation Is Nothing, DBNull.Value, FacilityLocation)
                    command.Parameters.Add("@FacilityCbg", SqlDbType.NVarChar, 50).Value = If(FacilityCbg Is Nothing, DBNull.Value, FacilityCbg)
                    command.Parameters.Add("@SubFacilityName", SqlDbType.NVarChar, 50).Value = If(SubFacilityName Is Nothing, DBNull.Value, SubFacilityName)

                    command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                    command.ExecuteNonQuery()


                    transaction.Commit()
                    Edit = True

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
                    Edit = False

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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Edit = False

        End Try

    End Function



    Public Function Delete(pe02 As Long, HCID As Integer, pe45 As Long, Pe26s As List(Of Long), MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeletedUsercase, pe02, pe45, String.Format("Delete Usercase: {0} from usercase: {1} of vehicle: {2} is deleted.", "ProcessStepName", "UsercaseName", "Vehicle ID"), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If


                    For Each pe26 As Long In Pe26s


                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepDelete.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@HealthChartId", SqlDbType.BigInt, 8).Value = HCID
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@pe26_SpecifcVehicleUsercases_PK", SqlDbType.BigInt, 8).Value = pe26
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using


                        If _tbAnswer.Rows.Count < 1 Then
                            Dim _unit As New VehiclePlan.Unit
                            Using dt As DataTable = _unit.GetVehiclesUsercasesDedicated(pe45, MainBuildType, transaction, conTnd)
                                If dt.Rows.Count > 0 Then Throw New Exception("The Process Step was not allowed to delete.")
                            End Using
                        Else

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                            command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                            command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe45_AllocatedPowerPack_Fk"))
                            command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))
                            command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                            command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = Date.Parse(_tbAnswer.Rows(0)("PlannedStart"))
                            command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  'This value must not be passed to DB
                            command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("Duration"))
                            command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("WorkingDays"))
                            command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsDBNull(_tbAnswer.Rows(0)("PublicHolidayFlag")) = True, 0, If(_tbAnswer.Rows(0)("PublicHolidayFlag") = True, 1, 0))
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId


                            command.ExecuteNonQuery()

                        End If

                    Next

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Delete = Nothing

        End Try


    End Function



    Public Function Delete(pe02 As Long, HCID As Integer, pe45 As Long, pe26 As Long, MainBuildType As String) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()
                    Dim command As SqlCommand
                    command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepDelete.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.BigInt, 8).Value = HCID
                    command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                    command.Parameters.Add("@pe26_SpecifcVehicleUsercases_PK", SqlDbType.BigInt, 8).Value = pe26


                    changelog = New ChangeLog()
                    ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeletedProcessStep, pe02, pe45, String.Format("ProcessStep: {0} from usercase: {1} of vehicle: {2} is deleted.", "ProcessStepName", "UsercaseName", "Vehicle ID"), MainBuildType, transaction, conTnd)
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If
                    command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId


                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using


                    If _tbAnswer.Rows.Count < 1 Then
                        Dim _unit As New VehiclePlan.Unit
                        Using dt As DataTable = _unit.GetVehiclesUsercasesDedicated(pe45, MainBuildType, transaction, conTnd)
                            If dt.Rows.Count > 0 Then Throw New Exception("The Process Step was not allowed to delete.")
                        End Using
                    Else

                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe45_AllocatedPowerPack_Fk"))
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = Date.Parse(_tbAnswer.Rows(0)("PlannedStart"))
                        command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  'This value must not be passed to DB
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("Duration"))
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("WorkingDays"))
                        command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsDBNull(_tbAnswer.Rows(0)("PublicHolidayFlag")) = True, 0, If(_tbAnswer.Rows(0)("PublicHolidayFlag") = True, 1, 0))
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId





                        command.ExecuteNonQuery()

                    End If


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
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Delete = Nothing

        End Try


    End Function



    Public Function CutPaste(frompe02 As Long, FromPe45 As Long, ToPe45 As Long, FromHCID As Integer,
                        PasteAllocatedUsercaseSeq As Integer, PasteProcessStepSequence As Integer, PastePlannedStart As Date, SelectedPe26s As List(Of Long), MainBuildType As String, ByRef NextPlannedStart As Object, InsertAsIndependentUsercase As Boolean) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim _tbAnswer2 As DataTable
        Dim ChangeLog As ChangeLog = Nothing
        Dim ActionId As Long = -1
        Dim counter As Integer = 1

        Try
            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    ChangeLog = New ChangeLog()


                    If SelectedPe26s.Count = 1 Then

                        ActionId = ChangeLog.AddChangeLog(DataCenter.ActionName.Tnd_CutPasteProcessStep, frompe02, Nothing, String.Format(".Net Cuts pe26 {0} from  pe45 {1} -> Pastes to pe45 {2}", SelectedPe26s(0), FromPe45, ToPe45), MainBuildType, transaction, conTnd)

                    ElseIf SelectedPe26s.Count > 1 Then

                        ActionId = ChangeLog.AddChangeLog(DataCenter.ActionName.Tnd_CutPasteUsercase, frompe02, Nothing, String.Format(".Net Cuts a set of Process steps from  pe45 {0} -> Pastes to pe45 {1}", FromPe45, ToPe45), MainBuildType, transaction, conTnd)

                    End If
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If

                    NextPlannedStart = Nothing
                    For Each pe26 As Long In SelectedPe26s
                        _tbAnswer = Nothing
                        _tbAnswer = Nothing


                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepAddNewInsert.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = FromHCID
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = ToPe45
                        command.Parameters.Add("@pe39_SlotFacilityMatching_FK", SqlDbType.BigInt, 4).Value = Nothing 'This value must be Null in new function
                        command.Parameters.Add("@pe26_SpecificvehicleUsercases_ID", SqlDbType.BigInt, 4).Value = pe26
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = If(_tbAnswer Is Nothing, PasteAllocatedUsercaseSeq, Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq")))
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = If(_tbAnswer Is Nothing, PasteProcessStepSequence, Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence")) + 1)
                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = If(_tbAnswer Is Nothing, PastePlannedStart, Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")).AddDays(1))
                        'command.Parameters.Add("@Plannedend", SqlDbType.Date, 3).Value = Nothing
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Nothing
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Nothing
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId


                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using


                        If _tbAnswer.Rows.Count < 1 Then Throw New Exception("The Process Step was not inserted.")


                        If counter = 1 And InsertAsIndependentUsercase = True Then

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepAddNewInsertSequence.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = ToPe45
                            command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))
                            command.Parameters.Add("@ProcessStepSeqMin", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                            command.Parameters.Add("@ProcessStepSeqMax", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                            command.Parameters.Add("@ActionId", SqlDbType.Int, 4).Value = ActionId

                            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                                _tbAnswer = New DataTable()
                                dataAdapter.Fill(_tbAnswer)
                            End Using


                        End If




                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = FromHCID
                        command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = ToPe45
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = Date.Parse(_tbAnswer.Rows(0)("PlannedStart"))
                        command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  'This value must not be passed to DB
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("Duration"))
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("WorkingDays"))
                        command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsDBNull(_tbAnswer.Rows(0)("PublicHolidayFlag")) = True, 0, If(_tbAnswer.Rows(0)("PublicHolidayFlag") = True, 1, 0))
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using


                        If _tbAnswer.Rows.Count < 1 Then Throw New Exception("The Process Step was not inserted after Specific_ProcessStepCallEditDate.")

                        PastePlannedStart = Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")).AddDays(1)
                        PasteProcessStepSequence = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence")) + 1
                        PasteAllocatedUsercaseSeq = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))

                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepDelete.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@HealthChartId", SqlDbType.BigInt, 8).Value = FromHCID
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = FromPe45
                        command.Parameters.Add("@pe26_SpecifcVehicleUsercases_PK", SqlDbType.BigInt, 8).Value = pe26
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId


                        _tbAnswer2 = Nothing
                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer2 = New DataTable()
                            dataAdapter.Fill(_tbAnswer2)
                        End Using


                        If _tbAnswer2.Rows.Count < 1 Then
                            Dim _unit As New VehiclePlan.Unit
                            Using dt As DataTable = _unit.GetVehiclesUsercasesDedicated(FromPe45, MainBuildType, transaction, conTnd)
                                If dt.Rows.Count > 0 Then Throw New Exception("The Process Step was not allowed to delete.")
                            End Using
                        Else

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = FromHCID
                            command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer2.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                            command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer2.Rows(0)("pe45_AllocatedPowerPack_FK"))
                            command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer2.Rows(0)("AllocatedUsercaseSeq"))
                            command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer2.Rows(0)("ProcessStepSequence"))
                            command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = Date.Parse(_tbAnswer2.Rows(0)("PlannedStart"))
                            command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  'This value must not be passed to DB
                            command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer2.Rows(0)("Duration"))
                            command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer2.Rows(0)("WorkingDays"))
                            command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsDBNull(_tbAnswer.Rows(0)("PublicHolidayFlag")) = True, 0, If(_tbAnswer.Rows(0)("PublicHolidayFlag") = True, 1, 0))
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteNonQuery()

                        End If

                        counter = counter + 1
                    Next

                    If _tbAnswer IsNot Nothing Then NextPlannedStart = Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")).AddDays(1)

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    CutPaste = True

                Catch ex0 As Exception

                    transaction.Rollback()
                    '----------------------------------------------------------------
                    ' Error classification mechanism check
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
                    CutPaste = False

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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            CutPaste = Nothing

        End Try

    End Function


    Public Function CopyPaste(TOpe02 As Long, ToPe45 As Long, ToHCID As Integer,
                        ToAllocatedUsercaseSeq As Integer, ToProcessStepSequence As Integer, ToPlannedStart As Date, SelectedPe26s As List(Of Long), MainBuildType As String, ByRef NextPlannedStart As Object, InsertAsIndependentUsercase As Boolean) As Boolean

        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1
        Dim counter As Integer = 1

        Try
            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    changelog = New ChangeLog()

                    If SelectedPe26s.Count = 1 Then

                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_CopyPasteProcessStep, TOpe02, Nothing, String.Format(".Net Copies ProcessStep: {0} to pe45: {1}", SelectedPe26s(0), ToPe45), MainBuildType, transaction, conTnd)

                    ElseIf SelectedPe26s.Count > 1 Then

                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_CopyPasteUsercase, TOpe02, Nothing, String.Format(".Net copies PSs to pe45 : {0}", ToPe45), MainBuildType, transaction, conTnd)

                    End If
                    If ActionId = -1 Then
                        Throw New Exception("The ActionID must not be -1.")
                    End If

                    _tbAnswer = Nothing
                    NextPlannedStart = Nothing
                    For Each pe26 As Long In SelectedPe26s



                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepAddNewInsert.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = ToHCID
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = ToPe45
                        command.Parameters.Add("@pe39_SlotFacilityMatching_FK", SqlDbType.BigInt, 4).Value = DBNull.Value  'This value must be Null in new function
                        command.Parameters.Add("@pe26_SpecificvehicleUsercases_ID", SqlDbType.BigInt, 4).Value = pe26
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = If(_tbAnswer Is Nothing, ToAllocatedUsercaseSeq, Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq")))
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = If(_tbAnswer Is Nothing, ToProcessStepSequence, Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence")) + 1)
                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = If(_tbAnswer Is Nothing, ToPlannedStart, Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")).AddDays(1))
                        ' command.Parameters.Add("@Plannedend", SqlDbType.Date, 3).Value = Nothing ' wrong parameter, Commemted by Ramesh 27th jul 2017
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Nothing
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Nothing
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId


                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using

                        If _tbAnswer.Rows.Count < 1 Then Throw New Exception("The Process Step was not inserted.")


                        If counter = 1 And InsertAsIndependentUsercase = True Then

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepAddNewInsertSequence.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = ToPe45
                            command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))
                            command.Parameters.Add("@ProcessStepSeqMin", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                            command.Parameters.Add("@ProcessStepSeqMax", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                            command.Parameters.Add("@ActionId", SqlDbType.Int, 4).Value = ActionId

                            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                                _tbAnswer = New DataTable()
                                dataAdapter.Fill(_tbAnswer)
                            End Using

                        End If



                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = ToHCID
                        command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = Long.Parse(_tbAnswer.Rows(0)("pe26_SpecificVehicleUsercases_PK"))
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = ToPe45
                        command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("AllocatedUsercaseSeq"))
                        command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("ProcessStepSequence"))
                        command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = Date.Parse(_tbAnswer.Rows(0)("PlannedStart"))
                        command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  'This value must not be passed to DB
                        command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("Duration"))
                        command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = Integer.Parse(_tbAnswer.Rows(0)("WorkingDays"))
                        command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = If(IsDBNull(_tbAnswer.Rows(0)("PublicHolidayFlag")) = True, 0, If(_tbAnswer.Rows(0)("PublicHolidayFlag") = True, 1, 0))
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using

                        counter = counter + 1
                    Next

                    NextPlannedStart = Date.Parse(_tbAnswer.Rows(0)("PlannedEnd")).AddDays(1)

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    CopyPaste = True

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
                            ErrorId = DataCenter.ErrorCenter.ProcessStep
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    CopyPaste = False

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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            CopyPaste = False

        End Try

    End Function




    ''' <summary>
    ''' The return value is only a row therefore we can use the object of the class to refer to values
    ''' </summary>
    ''' <param name="pe26"></param>
    ''' <returns></returns>
    Public Function SelectProcessStepDedicated(pe26 As Long, IsGenericPlan As Boolean) As Boolean

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_ProcessStepDedicated.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe26_SpecificVehicleUsercases_PK", SqlDbType.Int, 4).Value = pe26


                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)


                End Using

            End Using


            _Cdsid = _tbAnswer.Rows(0)(ColumnName.Cdsid).ToString()
            _ProcessStepName = _tbAnswer.Rows(0)(ColumnName.ProcessStepName).ToString()
            _Duration = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.Duration))
            _FacilityCbg = _tbAnswer.Rows(0)(ColumnName.FacilityCbg).ToString()
            _FacilityLocation = _tbAnswer.Rows(0)(ColumnName.FacilityLocation).ToString()
            _FacilityName = _tbAnswer.Rows(0)(ColumnName.FacilityName).ToString()
            _pe26 = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.pe26_SpecificVehicleUsercases_PK))
            _PlannedEnd = Date.Parse(_tbAnswer.Rows(0)(ColumnName.PlannedEnd))
            _PlannedStart = Date.Parse(_tbAnswer.Rows(0)(ColumnName.PlannedStart))
            _Remarks = _tbAnswer.Rows(0)(ColumnName.Remarks).ToString()
            _SubFacilityName = _tbAnswer.Rows(0)(ColumnName.SubFacilityName).ToString()
            _Usercase = _tbAnswer.Rows(0)(ColumnName.Usercase).ToString()
            _WorkingDays = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.WorkingDays))
            _TeamName = _tbAnswer.Rows(0)(ColumnName.TeamName).ToString()
            _XCCTeamName = _tbAnswer.Rows(0)(ColumnName.XCCTeamName).ToString()
            _GlobalDVP = _tbAnswer.Rows(0)(ColumnName.GlobalDVP).ToString()
            _AllocatedUsercaseSeq = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.AllocatedUsercaseSeq))
            _ProcessStepSeq = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.ProcessStepSeq))
            _IsWithHoliday = If(IsDBNull(_tbAnswer.Rows(0)(ColumnName.PublicHolidayFlag)) = True, False, Boolean.Parse(_tbAnswer.Rows(0)(ColumnName.PublicHolidayFlag)))


            DataCenter.GlobalValues.message = String.Empty
            SelectProcessStepDedicated = True
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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectProcessStepDedicated = False
        End Try


    End Function


    Public Function GetAllCdsids(pe01 As Long, HCID As Integer, DvpTeam As String, MainBuildType As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_CdsidDvpTeamDedicated.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe01", SqlDbType.BigInt, 8).Value = pe01
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID
                command.Parameters.Add("@DvpTeam", SqlDbType.NVarChar, 50).Value = DvpTeam


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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetAllCdsids = Nothing
        End Try




    End Function




    Public Function SelectProcessStepDedicated(pe45 As Long, intAllocatedUsercaseSeq As Integer, intProcessStepSeq As Integer, transaction As SqlTransaction, conTnd As SqlConnection) As Boolean

        Try


            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Code_ProcessStepDedicated.ToString())
            command.Connection = conTnd
            command.Transaction = transaction
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add("@pe45", SqlDbType.BigInt, 8).Value = pe45
            command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = intAllocatedUsercaseSeq
            command.Parameters.Add("@ProcessStepSeq", SqlDbType.Int, 4).Value = intProcessStepSeq


            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                _tbAnswer = New DataTable()
                dataAdapter.Fill(_tbAnswer)

            End Using



            _Cdsid = _tbAnswer.Rows(0)(ColumnName.Cdsid).ToString()
            _ProcessStepName = _tbAnswer.Rows(0)(ColumnName.ProcessStepName).ToString()
            _Duration = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.Duration))
            _FacilityCbg = _tbAnswer.Rows(0)(ColumnName.FacilityCbg).ToString()
            _FacilityLocation = _tbAnswer.Rows(0)(ColumnName.FacilityLocation).ToString()
            _FacilityName = _tbAnswer.Rows(0)(ColumnName.FacilityName).ToString()
            _pe26 = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.pe26_SpecificVehicleUsercases_PK))
            _PlannedEnd = Date.Parse(_tbAnswer.Rows(0)(ColumnName.PlannedEnd))
            _PlannedStart = Date.Parse(_tbAnswer.Rows(0)(ColumnName.PlannedStart))
            _Remarks = _tbAnswer.Rows(0)(ColumnName.Remarks).ToString()
            _SubFacilityName = _tbAnswer.Rows(0)(ColumnName.SubFacilityName).ToString()
            _Usercase = _tbAnswer.Rows(0)(ColumnName.Usercase).ToString()
            _WorkingDays = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.WorkingDays))
            _TeamName = _tbAnswer.Rows(0)(ColumnName.TeamName).ToString()
            _XCCTeamName = _tbAnswer.Rows(0)(ColumnName.XCCTeamName).ToString()
            _GlobalDVP = _tbAnswer.Rows(0)(ColumnName.GlobalDVP).ToString()
            _AllocatedUsercaseSeq = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.AllocatedUsercaseSeq))
            _ProcessStepSeq = Integer.Parse(_tbAnswer.Rows(0)(ColumnName.ProcessStepSeq))
            _IsWithHoliday = If(IsDBNull(_tbAnswer.Rows(0)(ColumnName.PublicHolidayFlag)) = True, False, Boolean.Parse(_tbAnswer.Rows(0)(ColumnName.PublicHolidayFlag)))

            DataCenter.GlobalValues.message = String.Empty
            SelectProcessStepDedicated = True
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
                    ErrorId = DataCenter.ErrorCenter.ProcessStep
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectProcessStepDedicated = False
        End Try


    End Function






    Public Function MoveRight(pe02 As Long, pe45s As List(Of Long), HCID As Integer, DayCount As Integer, MainBuildType As String) As Boolean


        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()


                    For Each pe45 In pe45s


                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_MoveLeftorRight, pe02, pe45, String.Format("vehicle pe45s: {0} is moved to right for {1} days.", pe45, DayCount.ToString), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        '-------------------------------------------------------------
                        ' Get first Process step must be get
                        '-------------------------------------------------------------
                        Dim _ProcessStep As ProcessStep = New ProcessStep
                        If _ProcessStep.SelectProcessStepDedicated(pe45, 0, 1, transaction, conTnd) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


                        Dim Command As SqlCommand

                        Command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        Command.Connection = conTnd
                        Command.Transaction = transaction
                        Command.CommandType = CommandType.StoredProcedure
                        Command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        Command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        Command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = _ProcessStep.pe26
                        Command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        Command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = _ProcessStep.AllocatedUsercaseSeq
                        Command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = _ProcessStep.ProcessStepSeq
                        Command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = _ProcessStep.PlannedStart.AddDays(DayCount)
                        Command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = DBNull.Value  ' MUST BE NULL
                        Command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = _ProcessStep.Duration
                        Command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = _ProcessStep.WorkingDays
                        Command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = _ProcessStep.IsWithHoliday
                        Command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Command.ExecuteNonQuery()


                    Next



                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    MoveRight = True

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
                            ErrorId = DataCenter.ErrorCenter.Unit
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    MoveRight = False

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
                    ErrorId = DataCenter.ErrorCenter.Unit
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            MoveRight = False

        End Try

    End Function


    Public Function MoveLeft(pe02 As Long, pe45s As List(Of Long), HCID As Integer, DayCount As Integer, MainBuildType As String) As Boolean


        Dim transaction As SqlTransaction = Nothing
        Dim changelog As ChangeLog = Nothing
        Dim ActionId As Long = -1

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()


                    For Each pe45 In pe45s


                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_MoveLeftorRight, pe02, pe45, String.Format("vehicle pe45s: {0} is moved to left for {1} days.", pe45, DayCount.ToString), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If



                        '-------------------------------------------------------------
                        '  first Process step must be get
                        '-------------------------------------------------------------
                        Dim _ProcessStep As ProcessStep = New ProcessStep
                        If _ProcessStep.SelectProcessStepDedicated(pe45, 0, 1, transaction, conTnd) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)


                        Dim Command As SqlCommand

                        Command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ProcessStepCallEditDate.ToString())
                        Command.Connection = conTnd
                        Command.Transaction = transaction
                        Command.CommandType = CommandType.StoredProcedure
                        Command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        Command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        Command.Parameters.Add("@pe26_SpecificVehicleUsercaseSeq", SqlDbType.BigInt, 8).Value = _ProcessStep.pe26
                        Command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        Command.Parameters.Add("@AllocatedUsercaseSeq", SqlDbType.Int, 4).Value = _ProcessStep.AllocatedUsercaseSeq
                        Command.Parameters.Add("@ProcessStepSequence", SqlDbType.Int, 4).Value = _ProcessStep.ProcessStepSeq
                        Command.Parameters.Add("@PlannedStart", SqlDbType.Date, 3).Value = _ProcessStep.PlannedStart.AddDays(DayCount * -1)
                        'command.Parameters.Add("@PlannedEnd", SqlDbType.Date, 3).Value = New Object() ' MUST BE NULL
                        Command.Parameters.Add("@Duration", SqlDbType.Int, 4).Value = _ProcessStep.Duration
                        Command.Parameters.Add("@WorkingDays", SqlDbType.Int, 4).Value = _ProcessStep.WorkingDays
                        Command.Parameters.Add("@PublicHolidayFlag", SqlDbType.Int, 4).Value = _ProcessStep.IsWithHoliday
                        Command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Command.ExecuteNonQuery()


                    Next



                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    MoveLeft = True

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
                            ErrorId = DataCenter.ErrorCenter.Unit
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    MoveLeft = False

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
                    ErrorId = DataCenter.ErrorCenter.Unit
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            MoveLeft = False

        End Try

    End Function




End Class
