
Imports System.Data
Imports System.Data.SqlClient

Namespace VehiclePlan

    ''' <summary>
    ''' Error Code : 60
    ''' </summary>
    Public Class Plan
        Inherits CtBaseClass
        Implements CT.Data.Interfaces.PlanInterface

        Public Enum SelectAllSpecificTndPlansColumns
            pe01_TnDBasicProgram_FK
            GenericSpecific
            HealthChartId
            HealthChartName
            PlanVersion
            FileStatus
            BuildPhase
            Quantity
            AssyMrd
            M1DC
            PEC
            FEC
            Platform
            Carline
            XCCpe01
            XCCpe26
            AssyBuildScale
            BuildType
            pe02
            pe01
            Region
        End Enum

        'Private _tbAnswer As DataTable = Nothing
        'Private _arrayDT As String(,) = Nothing
        ''' <summary>
        ''' This sub generates the list of plans from CT.
        ''' <para/>This table contains this columns:
        ''' <para/>pe01_TnDBasicProgram_FK,	GenericSpecific,	HealthChartID,	HealthChartName,	BuildPhase,	AssyMrd,	Quantity,	M1DC,	PEC,	FEC
        ''' </summary>
        ''' <returns> </returns>
        Public Function SelectAllSpecificTndPlans() As DataTable Implements Interfaces.PlanInterface.SelectAllSpecificTndPlans
            Dim _tbAnswer As DataTable = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing
                    command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Code_Vehicle_AllSpecificTnDPlanListVers115.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using
                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectAllSpecificTndPlans = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectAllSpecificTndPlans = Nothing
            End Try
        End Function

        Public Function SelectAllGenericTndPlan() As DataTable Implements Interfaces.PlanInterface.SelectAllGenericTndPlan
            Dim _tbAnswer As DataTable = Nothing

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Code_Vehicle_AllGenericTnDPlanListVers115.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectAllGenericTndPlan = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectAllGenericTndPlan = Nothing
            End Try

        End Function

        Public Function SelectAllTndDraftPlans(MainBuildType As String, HCID As Integer) As DataTable Implements Interfaces.PlanInterface.SelectAllTndDraftPlans
            Dim _tbAnswer As DataTable = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Code_Vehicle_AllTnDDraftPlanList.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@MainPlanHCID", SqlDbType.Int, 4).Value = HCID


                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectAllTndDraftPlans = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectAllTndDraftPlans = Nothing
            End Try

        End Function

        Public Function SelectTndDraftPlanDedicated(DraftHCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.SelectTndDraftPlanDedicated
            Dim _tbAnswer As DataTable = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Code_Vehicle_TnDDraftPlanDedicated.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@DraftHCID", SqlDbType.Int, 4).Value = DraftHCID


                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectTndDraftPlanDedicated = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectTndDraftPlanDedicated = Nothing
            End Try

        End Function


        ''' <summary>
        ''' pe02_TnDProgramDetails_PK , Quantity, BuildPhase, BuildTypes , Carline, Platform 101192
        ''' </summary>
        ''' <param name="pe01"></param>
        ''' <param name="HCID"></param>
        ''' <returns></returns>
        Public Function SelectPlanDedicated(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable
            Dim _tbAnswer As DataTable = Nothing
            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_SelectPlanDedicated.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HCID", SqlDbType.BigInt, 8).Value = HCID

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectPlanDedicated = _tbAnswer

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectPlanDedicated = Nothing
            End Try

        End Function


        'Public Function GetPlanColumnInputFormat(pe01 As Long, HCID As Integer) As DataTable

        '    Try
        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ColumnInputFormat.ToString())
        '            command.Connection = conTnd
        '            command.CommandType = CommandType.StoredProcedure

        '            _tbAnswer = Nothing
        '            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
        '                _tbAnswer = New DataTable()
        '                dataAdapter.Fill(_tbAnswer)
        '            End Using

        '        End Using

        '        DataCenter.GlobalValues.message = String.Empty
        '        GetPlanColumnInputFormat = _tbAnswer

        '    Catch ex As Exception
        '        '----------------------------------------------------------------
        '        ' Error classification mechanism
        '        '----------------------------------------------------------------
        '        Dim ErrorId As Integer
        '        Select Case ex.Message
        '            Case ex.Message.IndexOf("Permission") >= 0
        '                ErrorId = DataCenter.ErrorCenter.Permission
        '            Case ex.Message.IndexOf("could not found") >= 0
        '                ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
        '            Case Else
        '                ErrorId = DataCenter.ErrorCenter.Plan
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
        '        GetPlanColumnInputFormat = Nothing
        '    End Try

        'End Function



        'Public Function DetectOverlapping(HCID As Integer, MainBuildType As String) As DataTable


        '    Try
        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DetectOverlapping.ToString())
        '            command.Connection = conTnd
        '            command.CommandType = CommandType.StoredProcedure
        '            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
        '            command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID

        '            _tbAnswer = Nothing
        '            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
        '                _tbAnswer = New DataTable()
        '                dataAdapter.Fill(_tbAnswer)
        '            End Using

        '        End Using

        '        DataCenter.GlobalValues.message = String.Empty
        '        DetectOverlapping = _tbAnswer

        '    Catch ex As Exception
        '        '----------------------------------------------------------------
        '        ' Error classification mechanism
        '        '----------------------------------------------------------------
        '        Dim ErrorId As Integer
        '        Select Case ex.Message
        '            Case ex.Message.IndexOf("Permission") >= 0
        '                ErrorId = DataCenter.ErrorCenter.Permission
        '            Case ex.Message.IndexOf("could not found") >= 0
        '                ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
        '            Case Else
        '                ErrorId = DataCenter.ErrorCenter.Plan
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
        '        DetectOverlapping = Nothing
        '    End Try

        'End Function

        Public Function ConvertDraftToLife(pe01 As Long, LiveHCID As Integer, DraftOrCheckedoutHCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String, ByRef ActivePlanHCID As Integer) As Boolean Implements Interfaces.PlanInterface.ConvertDraftToLife
            Dim transaction As SqlTransaction = Nothing
            ActivePlanHCID = -1
            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                    Try
                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()
                        Dim command As SqlCommand
                        command = New SqlCommand(DataCenter.StoredProcedures.General.Draft_SwitchDraftToMaster.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@LiveHealthChartId", SqlDbType.BigInt, 8).Value = LiveHCID
                        command.Parameters.Add("@DraftHealthChartId", SqlDbType.Int, 4).Value = DraftOrCheckedoutHCID
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        _tbAnswer = Nothing
                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)
                        End Using

                        '----------------------------------------------------------------
                        ' Get active HCID 
                        '----------------------------------------------------------------
                        Select Case _tbAnswer.Rows.Count
                            Case <= 0
                                Throw New Exception("The count of returned HCID is less or equal to zero.")
                            Case > 1
                                Throw New Exception("The count of returned HCID is more than one.")
                            Case Else
                                ActivePlanHCID = CInt(_tbAnswer.Rows(0)("HealthChartId"))
                        End Select


                        If ActivePlanHCID <= 0 Then Throw New Exception("The return HCID is not valid")

                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        ConvertDraftToLife = True
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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        ConvertDraftToLife = False
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                ConvertDraftToLife = False
            End Try
        End Function


        Public Function ConvertCheckedouttToLife(pe01 As Long, LiveHCID As Integer, DraftOrCheckedoutHCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String) As Boolean Implements Interfaces.PlanInterface.ConvertCheckedouttToLife
            Dim transaction As SqlTransaction = Nothing
            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                    Try
                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()
                        Dim command As SqlCommand
                        command = New SqlCommand(DataCenter.StoredProcedures.General.Draft_SwitchCheckedoutToMaster.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@LiveHealthChartId", SqlDbType.BigInt, 8).Value = LiveHCID
                        command.Parameters.Add("@DraftHealthChartId", SqlDbType.Int, 4).Value = DraftOrCheckedoutHCID
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        ConvertCheckedouttToLife = True
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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        ConvertCheckedouttToLife = False
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                ConvertCheckedouttToLife = False
            End Try
        End Function


        Public Function Validaion_1(XCCpe01 As Long, HCID As Integer, BuildPhase As String, BuildType As String) As Boolean

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Generic_Check_Vehicle_1_InsertToPe01.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_programUniqueId", SqlDbType.BigInt, 8).Value = XCCpe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = CT.Data.DataCenter.BuildType.Vehicle.ToString
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 10).Value = BuildPhase
                    command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 10).Value = BuildType

                    conTnd.Open()
                    command.ExecuteNonQuery()
                    conTnd.Close()

                End Using

                DataCenter.GlobalValues.message = String.Empty
                Validaion_1 = True

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                Validaion_1 = False
            End Try


        End Function

        Public Function ValidatePlan(HealChartId As Integer, MainBuildType As String, FileStatus As String) As DataTable Implements Interfaces.PlanInterface.ValidatePlan
            Dim _tbAnswer As DataTable = Nothing

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Validation_VehicleOverallValidation.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealChartId
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus
                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                ValidatePlan = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                ValidatePlan = Nothing
            End Try


        End Function



        Public Function Validaion_2(XCCpe01 As Long, XCCpe26 As Long, HCID As Integer, BuildPhase As String, BuildType As String) As Boolean

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Generic_Check_Vehicle_2_InsertToPe02.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_programUniqueId", SqlDbType.BigInt, 8).Value = XCCpe01
                    command.Parameters.Add("@pe26_ProgramBaseInfoUniqueId", SqlDbType.BigInt, 8).Value = XCCpe26
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 10).Value = BuildPhase
                    command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 10).Value = BuildType
                    conTnd.Open()
                    command.ExecuteNonQuery()
                    conTnd.Close()
                End Using

                DataCenter.GlobalValues.message = String.Empty
                Validaion_2 = True

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Validaion_2 = False
            End Try


        End Function


        ''' <summary>
        ''' The output is a dataset. table[0] is engine , table [1] is transmission, table[2] is Prototype user.
        ''' The output columns are as following: 
        ''' <para />
        ''' pe29_BuildTypeCode, BuildTypeQuantity, PowertrainHardware, Quantity, PowerPackQuantity, 
        ''' SpecificationQuantity, PowerPackAllocationQuantity, Section, TotalQuantityValidation, QuantityValidation
        ''' <para />
        ''' The columns which are displayed in grid are the following.
        ''' </summary>
        ''' <param name="HCID"></param>
        ''' <param name="BuildType"></param>
        ''' <returns></returns>
        Public Function Validaion_5(HCID As Integer, BuildType As String) As DataSet
            Dim _dsAnswer As DataSet = New DataSet()

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Generic_Check_Vehicle_5_XCCEngineeTransmissionPrototypeuser.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 10).Value = BuildType



                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        dataAdapter.Fill(_dsAnswer)
                    End Using



                End Using

                DataCenter.GlobalValues.message = String.Empty
                Validaion_5 = _dsAnswer

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Validaion_5 = Nothing
            End Try


        End Function


        Public Function Validaion_6(HCID As Integer, BuildType As String) As DataTable

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Generic_Check_Vehicle_6_XCCRemovedEnginee.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 10).Value = BuildType



                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using



                End Using

                DataCenter.GlobalValues.message = String.Empty
                Validaion_6 = _tbAnswer

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Validaion_6 = Nothing
            End Try


        End Function




        Public Function GetDvpTeamAndCdsid(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_ProgramDvpTeamNameCdsid.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID


                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                GetDvpTeamAndCdsid = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetDvpTeamAndCdsid = Nothing
            End Try

        End Function



        Public Function GetPlanRemarks(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.GetPlanRemarks
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_PlanRemarks.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_TnDBasicProgram_ID", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID


                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                GetPlanRemarks = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPlanRemarks = Nothing
            End Try

        End Function

        Public Function GetAssignedCDSIDs(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.GetAssignedCDSIDs
            Try


                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_PlanAssignedCDSIDs.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_TnDBasicProgram_ID", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID


                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using


                DataCenter.GlobalValues.message = String.Empty
                GetAssignedCDSIDs = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetAssignedCDSIDs = Nothing
            End Try

        End Function



        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pe01"></param>
        ''' <param name="HCID"></param>
        ''' <param name="DvpTeamName"></param>
        ''' <param name="PMTLevel"> </param>
        ''' <param name="DNRLevel"></param>
        ''' <returns></returns>
        Public Function AssignCdsid2DvpTeam(pe01 As Long, HCID As Integer, MainBUildType As String, PmtGroup As String, DvpTeamName As String, PMTLevel As String, DNRLevel As String) As Boolean
            Dim transaction As SqlTransaction = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()
                        Dim command As SqlCommand


                        command = New SqlCommand(DataCenter.StoredProcedures.General.Specific_AssignCdsid2DvpTeamName.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBUildType
                        command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@PmtGroup", SqlDbType.NVarChar, 50).Value = PmtGroup
                        command.Parameters.Add("@DvpTeamName", SqlDbType.NVarChar, 50).Value = DvpTeamName
                        command.Parameters.Add("@PMTLevel", SqlDbType.NVarChar, 100).Value = If(PMTLevel.Length <> 0, PMTLevel, DBNull.Value)
                        command.Parameters.Add("@DNRLevel", SqlDbType.NVarChar, 200).Value = If(DNRLevel.Length <> 0, DNRLevel, DBNull.Value)

                        command.ExecuteNonQuery()

                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        AssignCdsid2DvpTeam = True

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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        AssignCdsid2DvpTeam = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                AssignCdsid2DvpTeam = False

            End Try
        End Function



        ''' <summary>
        ''' This function deletes the draft version or checkedout object or 
        ''' better to say discarrds the wip object in DB and change the state of plan 
        ''' from checkedout or Master.
        ''' </summary>
        ''' <param name="pe01"></param>
        ''' <param name="HCID"></param>
        ''' <param name="FileStatus"></param>
        ''' <returns></returns>
        Public Function DeleteDraftOrCheckedout(pe01 As Long, HCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String) As Boolean Implements Interfaces.PlanInterface.DeleteDraftOrCheckedout


            Dim transaction As SqlTransaction = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()
                        Dim command As SqlCommand

                        If FileStatus = DataCenter.FileStatus.Master Then Throw New Exception("This Plan cannot be deleted.")

                        If FileStatus = DataCenter.FileStatus.Draft Then

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Generic_DeleteXCCProgram.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                            command.ExecuteNonQuery()


                        ElseIf FileStatus = DataCenter.FileStatus.Checkedout Then

                            command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_Discard.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                            command.ExecuteNonQuery()

                        End If


                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        DeleteDraftOrCheckedout = True

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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        DeleteDraftOrCheckedout = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                DeleteDraftOrCheckedout = False

            End Try


        End Function


        Public Function GenerateDraftOrCheckout(ByRef HCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String) As Boolean Implements Interfaces.PlanInterface.GenerateDraftOrCheckout

            Dim transaction As SqlTransaction = Nothing
            Dim DraftScenarioId As Integer = 100

            Try



                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()
                        Dim command As SqlCommand



                        '-------------------------------------------------
                        ' Step 0
                        '-------------------------------------------------

                        If FileStatus = DataCenter.FileStatus.Draft Then

                            command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Code_Vehicle_AllTnDDraftPlanList.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@MainPlanHCID", SqlDbType.Int, 4).Value = HCID

                            _tbAnswer = Nothing
                            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                                _tbAnswer = New DataTable()
                                dataAdapter.Fill(_tbAnswer)
                            End Using

                            Select Case _tbAnswer.Rows.Count
                                Case 0
                            ' do nothing
                                Case 1, 2
                                    Dim MaxDraftVersion As String = _tbAnswer.Rows(_tbAnswer.Rows.Count - 1)("HealthChartID")
                                    If MaxDraftVersion <> String.Empty Then

                                        MaxDraftVersion = MaxDraftVersion.Substring(0, 3)
                                        DraftScenarioId = If(IsNumeric(MaxDraftVersion), Integer.Parse(MaxDraftVersion) + 1, DraftScenarioId)

                                    End If
                                Case 3
                                    Throw New Exception("Three draft version exists already in DB. The Max Draft version is 3.")
                            End Select
                        ElseIf FileStatus = DataCenter.FileStatus.Checkedout Then

                            DraftScenarioId = 999

                        Else
                            Throw New Exception("The passed filestatus value is not valid.")
                        End If

                        '-------------------------------------------------
                        ' Step 1
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_A_pe02.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()



                        '-------------------------------------------------
                        ' Step 2
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_E_pe34.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 3
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_F_pe22.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 4
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_G_pe44.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 5
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_H_pe26.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 6
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_I_pe77.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 7
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_J_pe75.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 8
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_K_pe73.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 9
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_L_pe71.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 10
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_M_pe69.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 11
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_N_pe66.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 12
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_O_pe85.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 13
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_P_pe78.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 14
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_Q_pe87.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 15
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_R_pe67.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 16
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_S_pe88.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 17
                        '-------------------------------------------------

                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_T_pe62.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 18
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.General.DraftGeneration_U_pe04.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@DraftScenarioId", SqlDbType.Int, 4).Value = DraftScenarioId
                        command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                        command.ExecuteNonQuery()

                        transaction.Commit()


                        If FileStatus = DataCenter.FileStatus.Checkedout Then HCID = Integer.Parse(DraftScenarioId.ToString + HCID.ToString())


                        DataCenter.GlobalValues.message = String.Empty
                        GenerateDraftOrCheckout = True

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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        GenerateDraftOrCheckout = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GenerateDraftOrCheckout = False

            End Try

        End Function





        ''' <summary>
        ''' if User want to have Custom formatting from beginning we must pass true value to WithCustomFormat.
        ''' </summary>
        ''' <param name="pe01"></param>
        ''' <param name="pe02"></param>
        ''' <param name="HCID"></param>
        ''' <param name="xccpe26"></param>
        ''' <param name="xccpe01"></param>
        ''' <param name="AssyBuildScale"></param>
        ''' <param name="BuildPhase"></param>
        ''' <param name="BuildType"></param>
        ''' <param name="WithCustomFormat"></param>
        ''' <returns></returns>
        Public Function ConvertGenericToSpecific(ByRef pe01 As Long, ByRef pe02 As Long, HCID As Integer, xccpe26 As Long, xccpe01 As Long, AssyBuildScale As Integer, BuildPhase As String, BuildType As String, FileStatus As DataCenter.FileStatus, WithCustomFormat As Boolean) As Boolean Implements Interfaces.PlanInterface.ConvertGenericToSpecific

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_ConvertToSpecific, pe02, Nothing, String.Format("Convert from generic to specific."), BuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If


                        Dim command As SqlCommand


                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        If pe01 <> 0 Then

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Generic_DeleteXCCProgram.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                            command.ExecuteNonQuery()

                        End If

                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle1_Generic_AddXccSingleProgram.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_ProgramUniqueId", SqlDbType.Int, 4).Value = xccpe01
                        command.Parameters.Add("@pe26_ProgramBaseInfoUniqueId", SqlDbType.Int, 4).Value = xccpe26
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 10).Value = BuildPhase
                        command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 10).Value = BuildType
                        command.Parameters.Add("@pe02_TnDProgramDetails_ID", SqlDbType.BigInt, 8).Value = Nothing
                        command.Parameters.Add("@pe01_TnDBasicProgram_ID", SqlDbType.BigInt, 8).Value = Nothing

                        command.Parameters("@pe02_TnDProgramDetails_ID").Direction = ParameterDirection.Output
                        command.Parameters("@pe01_TnDBasicProgram_ID").Direction = ParameterDirection.Output

                        command.ExecuteNonQuery()

                        pe02 = Long.Parse(command.Parameters("@pe02_TnDProgramDetails_ID").Value)
                        pe01 = Long.Parse(command.Parameters("@pe01_TnDBasicProgram_ID").Value)


                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle2_Generic_AddTnDPlanner.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_ProgramUniqueId", SqlDbType.Int, 4).Value = xccpe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                        command.ExecuteNonQuery()



                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle3_Generic_GenericSplit.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@XCCpe26_ProgramBaseInfoUniqueId_PK", SqlDbType.Int, 4).Value = xccpe26
                        command.Parameters.Add("@AssyBuildScalerate", SqlDbType.Int, 4).Value = AssyBuildScale

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle4_Generic_PowerPackAllocation.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe26_ProgramBaseInfoUniqueId", SqlDbType.Int, 4).Value = xccpe26

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle5_Generic_UsercaseAllocationLogicII.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@AssyBuildScale", SqlDbType.Int, 4).Value = AssyBuildScale

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle6_Generic_InitialBuckLoading.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@Vehpe02_TnDProgramDetails_PK", SqlDbType.BigInt, 8).Value = pe02

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle7_Specific_AllocatedUsercasesGenericToSpecificII.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@pe26_ProgramBaseInfoUniqueId", SqlDbType.BigInt, 8).Value = xccpe26
                        command.Parameters.Add("@ActionId", SqlDbType.Int, 4).Value = ActionId

                        command.ExecuteNonQuery()



                        '-------------------------------------------------
                        '
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle8_Generic_InitialPrototypeTeamnameAssaignment.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Indivitual formatting - pe87_ColumnFormat
                        ' Each row in table pe87_ColumnFormat is a column in interface
                        '-------------------------------------------------
                        If WithCustomFormat = True Then
                            Dim _PlanIndivitualFormatting As CT.Data.PlanIndivitualFormatting = New PlanIndivitualFormatting
                            If _PlanIndivitualFormatting.InitialFormat(pe01, HCID, BuildType, FileStatus.ToString, transaction, conTnd) = False Then
                                Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                            End If
                        End If



                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        ConvertGenericToSpecific = True

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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        ConvertGenericToSpecific = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                ConvertGenericToSpecific = False

            End Try

        End Function


        Public Function GetCTEnginesAndTransmissions(HCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.GetCTEnginesAndTransmissions

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_AssaignedCTEnginesAndTransmissions.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using


                End Using
                DataCenter.GlobalValues.message = String.Empty
                GetCTEnginesAndTransmissions = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetCTEnginesAndTransmissions = Nothing
            End Try


        End Function

        Public Function GetQuantityTableXCC(HCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.GetQuantityTableXCC

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_Vehicle_QuantityTableXCC.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = HCID

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using


                End Using
                DataCenter.GlobalValues.message = String.Empty
                GetQuantityTableXCC = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetQuantityTableXCC = Nothing
            End Try


        End Function



        Public Function GetQuantityTableCT(HCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.GetQuantityTableCT

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_Vehicle_QuantityTableCT.ToString())
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
                GetQuantityTableCT = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetQuantityTableCT = Nothing
            End Try


        End Function



        Public Function GetXCCEnginesAndTransmissions(HCID As Integer, MainBuildType As String) As DataTable Implements Interfaces.PlanInterface.GetXCCEnginesAndTransmissions


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_AssaignedXCCEnginesAndTransmissions.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    'command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using


                End Using
                ' ConvertDataTableToStingArray()
                'It means no error has been occured
                DataCenter.GlobalValues.message = String.Empty
                GetXCCEnginesAndTransmissions = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetXCCEnginesAndTransmissions = Nothing
            End Try


        End Function






        Public Function SelectDateInformation(pe02 As Long) As DataTable Implements Interfaces.PlanInterface.SelectDateInformation


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_ListDateInformation.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_ID", SqlDbType.BigInt, 8).Value = pe02

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                    DataCenter.GlobalValues.message = String.Empty
                    SelectDateInformation = _tbAnswer


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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                SelectDateInformation = Nothing

            End Try

        End Function









        ''' <summary>
        ''' This function must be called before loading a generic plan
        ''' </summary>
        ''' <returns></returns>
        Public Function GenerateGenericPlan(ByRef pe01 As Long, ByRef pe02 As Long, HCID As Integer, xccpe26 As Long, xccpe01 As Long, AssyBuildScale As Integer, BuildPhase As String, BuildType As String, FileStatus As DataCenter.FileStatus) Implements Interfaces.PlanInterface.GenerateGenericPlan

            Dim transaction As SqlTransaction = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()
                        Dim command As SqlCommand

                        '-------------------------------------------------
                        ' Step 1
                        '-------------------------------------------------
                        If pe01 <> 0 Then

                            command = New SqlCommand(DataCenter.StoredProcedures.General.Generic_DeleteXCCProgram.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus.ToString

                            command.ExecuteNonQuery()
                        End If

                        '-------------------------------------------------
                        ' Step 2
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle1_Generic_AddXccSingleProgram.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_ProgramUniqueId", SqlDbType.Int, 4).Value = xccpe01
                        command.Parameters.Add("@pe26_ProgramBaseInfoUniqueId", SqlDbType.Int, 4).Value = xccpe26
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 10).Value = BuildPhase
                        command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 10).Value = BuildType
                        command.Parameters.Add("@pe02_TnDProgramDetails_ID", SqlDbType.BigInt, 8).Value = Nothing
                        command.Parameters.Add("@pe01_TnDBasicProgram_ID", SqlDbType.BigInt, 8).Value = Nothing

                        command.Parameters("@pe02_TnDProgramDetails_ID").Direction = ParameterDirection.Output
                        command.Parameters("@pe01_TnDBasicProgram_ID").Direction = ParameterDirection.Output

                        command.ExecuteNonQuery()
                        pe01 = 0
                        pe02 = 0
                        pe02 = Long.Parse(command.Parameters("@pe02_TnDProgramDetails_ID").Value)
                        pe01 = Long.Parse(command.Parameters("@pe01_TnDBasicProgram_ID").Value)

                        If pe01 = 0 Or pe02 = 0 Then Throw New Exception("The XCC Plan couldn't be saved as Single program.")

                        '-------------------------------------------------
                        ' Step 2.1
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle2_Generic_AddTnDPlanner.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_ProgramUniqueId", SqlDbType.Int, 4).Value = xccpe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 3
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle3_Generic_GenericSplit.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@XCCpe26_ProgramBaseInfoUniqueId_PK", SqlDbType.Int, 4).Value = xccpe26
                        command.Parameters.Add("@AssyBuildScalerate", SqlDbType.Int, 4).Value = AssyBuildScale

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 4
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle4_Generic_PowerPackAllocation.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe26_ProgramBaseInfoUniqueId", SqlDbType.Int, 4).Value = xccpe26

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 5
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle5_Generic_UsercaseAllocationLogicII.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@AssyBuildScale", SqlDbType.Int, 4).Value = AssyBuildScale

                        command.ExecuteNonQuery()


                        '-------------------------------------------------
                        ' Step 6
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle6_Generic_InitialBuckLoading.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@Vehpe02_TnDProgramDetails_PK", SqlDbType.BigInt, 8).Value = pe02

                        command.ExecuteNonQuery()

                        '-------------------------------------------------
                        ' Step 7
                        '-------------------------------------------------
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.LoadingVehicle8_Generic_InitialPrototypeTeamnameAssaignment.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = BuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID

                        command.ExecuteNonQuery()


                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        GenerateGenericPlan = True

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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        GenerateGenericPlan = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GenerateGenericPlan = False

            End Try


        End Function


        Public Function UpdateTnDProgramDetails(pe02 As Long, HealthChartId As Integer, MainBuildType As String, AssyMrd As Object, Firstm1 As Object, M1DC As Object, FirstVP As Object, PEC As Object, FEC As Object) As Boolean

            Dim transaction As SqlTransaction = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Specific_UpdateTnDProgramDetails.ToString())

                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_PK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@AssyMrd", SqlDbType.Date, 3).Value = AssyMrd
                        command.Parameters.Add("@Firstm1", SqlDbType.Date, 3).Value = Firstm1
                        command.Parameters.Add("@M1DC", SqlDbType.Date, 3).Value = M1DC
                        command.Parameters.Add("@FirstVP", SqlDbType.Date, 3).Value = FirstVP
                        command.Parameters.Add("@PEC", SqlDbType.Date, 3).Value = PEC
                        command.Parameters.Add("@FEC", SqlDbType.Date, 3).Value = FEC

                        command.ExecuteNonQuery()

                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        UpdateTnDProgramDetails = True

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
                                ErrorId = DataCenter.ErrorCenter.Plan
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        UpdateTnDProgramDetails = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateTnDProgramDetails = False

            End Try

        End Function

        Public Function GetXCCUserTeamNameTranslation(MainBuildType As String) As DataTable

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_XCCUserTeamNameTranslation.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    ' Code disabled in DAL until Marcel completes the SP code amendments for thsi procedure.
                    'command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                GetXCCUserTeamNameTranslation = _tbAnswer
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
                GetXCCUserTeamNameTranslation = Nothing
            End Try

        End Function



        Public Function PrecheckF4Test(HCID As Integer) As DataTable


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_TnDPlanF4TValidation.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@CallingFromAddin", SqlDbType.Bit, 1).Value = 1

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using


                End Using
                'ConvertDataTableToStingArray()
                'It means no error has been occured
                DataCenter.GlobalValues.message = String.Empty
                PrecheckF4Test = _tbAnswer
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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                PrecheckF4Test = Nothing
            End Try


        End Function


        Public Function PushToF4Test(HCID As Integer) As Boolean
            Dim transaction As SqlTransaction = Nothing

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.Report_TnDPlanF4TTransfer.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID


                        command.ExecuteNonQuery()

                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        PushToF4Test = True

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
                        PushToF4Test = False

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
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                PushToF4Test = False
            End Try


        End Function




    End Class
End Namespace