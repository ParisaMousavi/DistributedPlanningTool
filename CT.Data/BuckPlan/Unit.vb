Imports System.Data
Imports System.Data.SqlClient
Imports CT.Data.Interfaces

Namespace BuckPlan


    Public Class Unit
        Inherits CtBaseClass
        Implements Interfaces.UnitInterface

        'Private _tbAnswer As DataTable = Nothing
        'Private _arrayDT As String(,) = Nothing

        Public Function GetPreviousValueGeneral(pe02 As Long, pe03 As Long, strField As String, MainBuildType As String) As String Implements Interfaces.UnitInterface.GetPreviousValueGeneral

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_GetUnitProgramInformation.ToString())
                    command.Connection = conTnd
                    command.Parameters.Add("@pe03", SqlDbType.BigInt, 8).Value = pe03
                    command.Parameters.Add("@pe02", SqlDbType.BigInt, 8).Value = pe02

                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                If _tbAnswer.Rows.Count <> 1 Then

                    Throw New Exception("the output is not allowed to be more than one row")

                End If

                DataCenter.GlobalValues.message = String.Empty
                GetPreviousValueGeneral = _tbAnswer.Rows(0)(strField).ToString

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
                GetPreviousValueGeneral = String.Empty

            End Try

        End Function

        Public Function GetPreviousValueVin(pe02 As Long, pe03 As Long, MainBuildType As String) As String Implements Interfaces.UnitInterface.GetPreviousValueVin


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_GetUnitProgramInformation.ToString())
                    command.Connection = conTnd
                    command.Parameters.Add("@pe03", SqlDbType.BigInt, 8).Value = pe03
                    command.Parameters.Add("@pe02", SqlDbType.BigInt, 8).Value = pe02

                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                If _tbAnswer.Rows.Count <> 1 Then Throw New Exception("the output is not allowed to be more than one row")

                DataCenter.GlobalValues.message = String.Empty
                GetPreviousValueVin = _tbAnswer.Rows(0)("Vin").ToString
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
                GetPreviousValueVin = String.Empty
            End Try


        End Function

        Public Function GetPreviousValueVehicleNumber(pe02 As Long, pe03 As Long, MainBuildType As String) As String Implements Interfaces.UnitInterface.GetPreviousValueVehicleNumber


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_GetUnitProgramInformation.ToString())
                    command.Connection = conTnd
                    command.Parameters.Add("@pe03", SqlDbType.BigInt, 8).Value = pe03
                    command.Parameters.Add("@pe02", SqlDbType.BigInt, 8).Value = pe02

                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                If _tbAnswer.Rows.Count <> 1 Then Throw New Exception("the output is not allowed to be more than one row")

                DataCenter.GlobalValues.message = String.Empty
                GetPreviousValueVehicleNumber = _tbAnswer.Rows(0)("TBNumber").ToString
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
                GetPreviousValueVehicleNumber = Nothing
            End Try


        End Function

        Public Function GetPreviousValueShippingToCustomer(pe02 As Long, pe03 As Long, MainBuildType As String) As String Implements Interfaces.UnitInterface.GetPreviousValueShippingToCustomer
            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_GetUnitProgramInformation.ToString())
                    command.Connection = conTnd
                    command.Parameters.Add("@pe03", SqlDbType.BigInt, 8).Value = pe03
                    command.Parameters.Add("@pe02", SqlDbType.BigInt, 8).Value = pe02

                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                If _tbAnswer.Rows.Count <> 1 Then Throw New Exception("the output is not allowed to be more than one row")

                DataCenter.GlobalValues.message = String.Empty
                If _tbAnswer.Rows(0)("ShippingToCustomerDate").ToString <> "" Then
                    GetPreviousValueShippingToCustomer = _tbAnswer.Rows(0)("ShippingToCustomerDate")
                Else
                    GetPreviousValueShippingToCustomer = String.Empty
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
                        ErrorId = DataCenter.ErrorCenter.Unit
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPreviousValueShippingToCustomer = String.Empty
            End Try

        End Function

        ''' <summary>
        ''' 
        ''' The return value is a list of vehicles which have been changed with pe03
        ''' </summary>
        ''' <param name="pe02"></param>
        ''' <param name="CurrentDisplaySeq"></param>
        ''' <param name="FutureDisplaySeq"></param>
        ''' <returns></returns>
        Public Function ChangeBuildSequence(pe02 As Long, CurrentDisplaySeq As Integer, FutureDisplaySeq As Integer, MainBuildType As String) As DataTable Implements Interfaces.UnitInterface.ChangeBuildSequence

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeVehicleDisplaySeq.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@CurrentDisplaySeq", SqlDbType.Int, 4).Value = CurrentDisplaySeq
                        command.Parameters.Add("@FutureDisplaySeq", SqlDbType.Int, 4).Value = FutureDisplaySeq

                        'This text 00000 must be replaced in store3d procedure. 
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_ChangeVehicleSequence, pe02, Nothing, String.Format(".Net Vehicle 00000 from position {0} to {1}.", CurrentDisplaySeq, FutureDisplaySeq), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                            _tbAnswer = New DataTable()
                            dataAdapter.Fill(_tbAnswer)

                        End Using

                        transaction.Commit()
                        'ConvertDataTableToStingArray()
                        ChangeBuildSequence = _tbAnswer ' _arrayDT

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
                        ChangeBuildSequence = Nothing

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
                ChangeBuildSequence = Nothing

            End Try

        End Function

        Public Function AddUnit(HCID As Integer, pe01 As Long, pe02 As Long, BuildPhase As String, MainBuildType As String, HardwareBuildType As String, HealthChartId As Integer, ByRef pe03_ID As Long, ByRef pe45_ID As Long, ByRef GenericSplitRowNumber As Integer) As Boolean Implements Interfaces.UnitInterface.AddUnit

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim messagePassing As MessagePassing = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try


                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_AddUnit.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 5).Value = BuildPhase
                        command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 15).Value = HardwareBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealthChartId
                        command.Parameters.Add("@pe03_ID", SqlDbType.BigInt, 8).Value = DBNull.Value
                        command.Parameters.Add("@pe45_ID", SqlDbType.BigInt, 8).Value = DBNull.Value
                        command.Parameters.Add("@GenericSplitRowNumber", SqlDbType.Int, 4).Value = DBNull.Value

                        command.Parameters("@pe03_ID").Direction = ParameterDirection.Output
                        command.Parameters("@pe45_ID").Direction = ParameterDirection.Output
                        command.Parameters("@GenericSplitRowNumber").Direction = ParameterDirection.Output

                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_AddNewVehicle, pe02, Nothing, String.Format(".Net Add new vehicle."), MainBuildType, transaction, conTnd)

                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteScalar()

                        pe03_ID = command.Parameters("@pe03_ID").Value
                        pe45_ID = command.Parameters("@pe45_ID").Value
                        GenericSplitRowNumber = command.Parameters("@GenericSplitRowNumber").Value


                        '----------------------------------------------------------------
                        ' Send a message to other plans
                        '----------------------------------------------------------------
                        messagePassing = New MessagePassing()
                        If (messagePassing.Insert(HCID, MainBuildType, String.Format("New {0} has been inserted in {1} Plan with HCID {2}.", HardwareBuildType, MainBuildType, HCID), transaction, conTnd) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)



                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        AddUnit = True

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
                        pe03_ID = Nothing
                        pe45_ID = Nothing
                        AddUnit = False

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
                pe03_ID = Nothing
                pe45_ID = Nothing
                GenericSplitRowNumber = Nothing
                AddUnit = False

            End Try

        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pe02"></param>
        ''' <param name="pe45"></param>
        ''' <param name="NewEngineName">Max length is 50 character.</param>
        ''' <returns></returns>
        Public Function ChangeEngine(pe02 As Long, pe45 As Long, NewEngineName As String, MainBuildType As String) As Boolean Implements Interfaces.UnitInterface.ChangeEngine

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try


                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeVehicleEngine.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@pe45_AllocatedPowerPack_PK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@EngineName", SqlDbType.NVarChar, 50).Value = NewEngineName

                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EngineInfo, pe02, pe45, String.Format("Change engine in vehicle {0}.", pe45), MainBuildType, transaction, conTnd)

                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteScalar()


                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        ChangeEngine = True

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
                        ChangeEngine = False

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
                ChangeEngine = False

            End Try

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pe02"></param>
        ''' <param name="pe45"></param>
        ''' <param name="NewTransmissionName">Max lenth is 50 characters</param>
        ''' <returns></returns>
        Public Function ChangeTransmission(pe02 As Long, pe45 As Long, NewTransmissionName As String, MainBuildType As String) As Boolean Implements Interfaces.UnitInterface.ChangeTransmission

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try


                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeVehicleTransmission.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@pe45_AllocatedPowerPack_PK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@TransName", SqlDbType.NVarChar, 50).Value = NewTransmissionName

                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_TransInfo, pe02, pe45, String.Format("Change Transmission in vehicle {0}.", pe45), MainBuildType, transaction, conTnd)

                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteScalar()


                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        ChangeTransmission = True

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
                        ChangeTransmission = False

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
                ChangeTransmission = False

            End Try

        End Function






        ''' <summary>
        ''' Return column/s:
        '''  <para/> pe26_SpecificVehicleUsercases_PK,AllocatedUsercaseSeq,pe45_AllocatedPowerPack_FK,RowNumber,GenericSplitRownumber,GenericSplitRank,Gpat,XCCPrototypeUser,Translation,
        '''  <para/> DvpTeamName,Usercase,Distribution,PlanStatus,UsercaseLevel,UsercaseStatus,ProcessStepSequence,ProcessStepName,ProcessStepCode,Duration,PlannedStart,PlannedEnd,DisplayPlannedStart,DisplayPlannedEnd,
        '''  <para/> ProcessStepBackRGB,ProcessStepFontRGB,TwentyFourSeven,WorkingDays,FacilityCbg,FacilityName,FacilityCode,FacilityLocation,FacilityKeyContact,FacilityCostCenter,
        '''  <para/> FacilityGroupName,pe26.SubFacilityName,pe45.EngineName,TransName,pe26_SpecificVehicleUsercases_PK,ProcessStepDisplay,BuildStart,pe03_TnDProgramVehicles_FK,
        '''  <para/> VIN,TBNumber,BuckNumber,ColorCode,BodyStyle,DriveSide,Dedicated,BuildNumber,EmissionStagem,ShippingAdress,ShippingToCustomerDate,CBG,VehiclePrototypeUser,VehicleTeamName,EngineType,TransType
        ''' </summary>
        ''' <param name="pe45"></param>
        ''' <returns></returns>
        Public Function GetVehiclesUsercasesDedicated(pe45 As Long, MainBuildType As String, Optional transaction As SqlTransaction = Nothing, Optional conTnd As SqlConnection = Nothing) As DataTable Implements Interfaces.UnitInterface.GetVehiclesUsercasesDedicated

            Try


                If conTnd Is Nothing Then
                    conTnd = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                End If


                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_VehiclesUsercasesDisplayDedicated.ToString())
                command.Connection = conTnd
                If transaction IsNot Nothing Then command.Transaction = transaction
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe45_AllocatedPowerPack_PK", SqlDbType.BigInt, 8).Value = pe45

                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)

                End Using

                DataCenter.GlobalValues.message = String.Empty
                GetVehiclesUsercasesDedicated = _tbAnswer

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
                GetVehiclesUsercasesDedicated = Nothing

            End Try

        End Function






        Public Function Delete(HCID As Integer, pe03 As Long, pe02 As Long, pe45 As Long, MainBuildType As String) As Boolean Implements Interfaces.UnitInterface.Delete

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim messagePassing As MessagePassing = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DeactivateVehicle.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe03_TnDProgramVehicles_Fk", SqlDbType.BigInt, 8).Value = pe03

                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeleteVehicle, pe02, pe45, String.Format(".Net Delete vehicle pe03 , pe45 : {0} , {1}.", pe03.ToString(), pe45.ToString), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteNonQuery()


                        '----------------------------------------------------------------
                        ' Send a message to other plans
                        '----------------------------------------------------------------
                        messagePassing = New MessagePassing()
                        If (messagePassing.Insert(HCID, MainBuildType, String.Format("Unit has been removed from {0} Plan with HCID {1}.", MainBuildType, HCID), transaction, conTnd) = False) Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)




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
                                ErrorId = DataCenter.ErrorCenter.Unit
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
                        ErrorId = DataCenter.ErrorCenter.Unit
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Delete = False

            End Try

        End Function

        Public Function ChangeInfoII(Pe02 As Long, pe45 As Long, Pe03 As Long, MainBuildType As String, FileStatus As String, HealthChartId As Long, Optional StrCBG As Object = Nothing,
                                     Optional StrXccTeamName As Object = Nothing,
                                     Optional StrDedicated As Object = Nothing,
                                     Optional StrTBNumber As Object = Nothing,
                                     Optional StrVin As Object = Nothing,
                                     Optional StrEmissionStage As Object = Nothing,
                                     Optional StrBodySyle As Object = Nothing,
                                     Optional StrColorCode As Object = Nothing,
                                     Optional StrDriveSide As Object = Nothing,
                                     Optional StrTeamName As Object = Nothing,
                                     Optional StrRemarks As Object = Nothing,
                                     Optional StrShippingToCustomerDate As Object = Nothing,
                                     Optional strCustomerRequiredDate As Object = Nothing,
                                     Optional strRigCustomerPickDate As Object = Nothing,
                                     Optional StrTbNumberPrefix As Object = Nothing,
                                     Optional StrBuildId As Object = Nothing,
                                     Optional StrTagNumber As Object = Nothing,
                                     Optional StrPaintFacility As Object = Nothing,
                                     Optional CustomerRequiredDate As Object = Nothing,
                                     Optional RigCustomerPickDate As Object = Nothing) As Boolean Implements Interfaces.UnitInterface.ChangeInfoII
            Throw New NotImplementedException()
        End Function






        'Private Sub ConvertDataTableToStingArray()

        '    Dim i, j As Integer
        '    If _tbAnswer IsNot Nothing Then

        '        ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count)
        '        For i = 0 To _tbAnswer.Rows.Count - 1
        '            For j = 0 To _tbAnswer.Columns.Count - 1
        '                _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
        '            Next j
        '        Next i
        '    End If
        'End Sub


    End Class

End Namespace
