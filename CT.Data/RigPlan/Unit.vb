Imports System.Data
Imports System.Data.SqlClient

Namespace RigPlan

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

        '<Obsolete("This method is deprecated, use ChangeInfoII instead.", True)>
        'Public Function ChangeInfo(Pe02 As Long, pe45 As Long, Pe03 As Long, CBG As String, XccTeamName As String, Dedicated As String, TBNumber As String, Vin As String, EmissionStage As String, BodySyle As String, ColorCode As String, DriveSide As String, TeamName As String, Remarks As String, ShippingToCustomerDate As Object, TbNumberPrefix As Object, BuildId As Object, TagNumber As Object, PaintFacility As Object, MainBuildType As String) As Boolean

        '    Dim transaction As SqlTransaction = Nothing
        '    Dim changelog As ChangeLog = Nothing
        '    Dim ActionId As Long = -1
        '    changelog = New ChangeLog()


        '    Try

        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Try
        '                conTnd.Open()
        '                transaction = conTnd.BeginTransaction()

        '                Dim strMessage = " ,Vin : " + If(Vin IsNot Nothing, Vin.ToString, "") + " ,TBNumber : " + If(TBNumber IsNot Nothing, TBNumber.ToString, "") + " ,ColorCode : " + If(ColorCode IsNot Nothing, ColorCode.ToString, "") + " ,BodySyle : " + If(BodySyle IsNot Nothing, BodySyle.ToString, "")
        '                strMessage = strMessage + " ,DriveSide : " + If(DriveSide IsNot Nothing, DriveSide.ToString, "") + " ,Dedicated : " + If(Dedicated IsNot Nothing, Dedicated.ToString, "") + " ,EmissionStage : " + If(EmissionStage IsNot Nothing, EmissionStage.ToString, "")
        '                strMessage = strMessage + " ,Remarks : " + If(Remarks IsNot Nothing, Remarks.ToString, "") + " ,XccTeamName : " + If(XccTeamName IsNot Nothing, XccTeamName.ToString, "") + " ,CBG : " + If(CBG IsNot Nothing, CBG.ToString, "") + " ,TeamName : " + If(TeamName IsNot Nothing, TeamName.ToString, "") + " ,ShippingToCustomerDate : " + If(ShippingToCustomerDate IsNot Nothing, ShippingToCustomerDate.ToString, "")

        '                ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_ProgramInfo, Pe02, pe45, String.Format("Change program info in vehicle {0} {1}.", pe45, strMessage), MainBuildType, transaction, conTnd)

        '                If ActionId = -1 Then
        '                    Throw New Exception("The ActionID must not be -1.")
        '                End If


        '                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedure.ooo.ToString())
        '                command.Connection = conTnd
        '                command.Transaction = transaction
        '                command.CommandType = CommandType.StoredProcedure
        '                command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
        '                command.Parameters.Add("@Vin", SqlDbType.NVarChar, 50).Value = Vin

        '                command.Parameters.Add("@TbNumberPrefix", SqlDbType.NVarChar, 4).Value = If(TbNumberPrefix Is Nothing, DBNull.Value, TbNumberPrefix.ToString)
        '                command.Parameters.Add("@TBNumber", SqlDbType.NVarChar, 4).Value = TBNumber

        '                command.Parameters.Add("@BuildId", SqlDbType.NVarChar, 8).Value = If(BuildId Is Nothing, DBNull.Value, BuildId.ToString)
        '                command.Parameters.Add("@TagNumber", SqlDbType.NVarChar, 7).Value = If(TagNumber Is Nothing, DBNull.Value, TagNumber.ToString)

        '                command.Parameters.Add("@PaintFacility", SqlDbType.NVarChar, 100).Value = If(PaintFacility Is Nothing, DBNull.Value, PaintFacility.ToString)
        '                command.Parameters.Add("@ColorCode", SqlDbType.NVarChar, 50).Value = ColorCode

        '                command.Parameters.Add("@BodySyle", SqlDbType.NVarChar, 50).Value = BodySyle
        '                command.Parameters.Add("@DriveSide", SqlDbType.NVarChar, 50).Value = DriveSide
        '                command.Parameters.Add("@Dedicated", SqlDbType.NVarChar, 50).Value = Dedicated
        '                command.Parameters.Add("@EmissionStage", SqlDbType.NVarChar, 50).Value = EmissionStage
        '                command.Parameters.Add("@Remarks", SqlDbType.NVarChar, 100).Value = Remarks
        '                command.Parameters.Add("@ShippingToCustomerDate", SqlDbType.NVarChar, 10).Value = If(ShippingToCustomerDate Is Nothing, DBNull.Value, ShippingToCustomerDate)
        '                command.Parameters.Add("@CBG", SqlDbType.NVarChar, 10).Value = CBG

        '                command.Parameters.Add("@XCCPrototypeUser", SqlDbType.NVarChar, 50).Value = XccTeamName
        '                command.Parameters.Add("@Translation", SqlDbType.NVarChar, 50).Value = TeamName

        '                command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

        '                command.ExecuteScalar()


        '                transaction.Commit()
        '                DataCenter.GlobalValues.message = String.Empty
        '                ChangeInfo = True

        '            Catch ex0 As Exception

        '                transaction.Rollback()
        '                '----------------------------------------------------------------
        '                ' Error classification mechanism
        '                '----------------------------------------------------------------
        '                Dim ErrorId As Integer
        '                Select Case ex0.Message
        '                    Case ex0.Message.IndexOf("Permission") >= 0
        '                        ErrorId = DataCenter.ErrorCenter.Permission
        '                    Case ex0.Message.IndexOf("could not found") >= 0
        '                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
        '                    Case Else
        '                        ErrorId = DataCenter.ErrorCenter.Unit
        '                End Select
        '                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
        '                ChangeInfo = False

        '            End Try

        '        End Using

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
        '                ErrorId = DataCenter.ErrorCenter.Unit
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
        '        ChangeInfo = False

        '    End Try

        'End Function


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

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1
            changelog = New ChangeLog()


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try
                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim strMessage = ""

                        If StrXccTeamName IsNot Nothing Then strMessage = strMessage + String.Format("XCC Team Name : {0} ,", StrXccTeamName.ToString)
                        If StrDedicated IsNot Nothing Then strMessage = strMessage + String.Format("Dedicated : {0} ,", StrDedicated.ToString)
                        If StrTBNumber IsNot Nothing Then strMessage = strMessage + String.Format("TBNumber : {0} ,", StrTBNumber.ToString)
                        If StrVin IsNot Nothing Then strMessage = strMessage + String.Format("Vin : {0} ,", StrVin.ToString)
                        If StrEmissionStage IsNot Nothing Then strMessage = strMessage + String.Format("EmissionStage : {0} ,", StrEmissionStage.ToString)
                        If StrBodySyle IsNot Nothing Then strMessage = strMessage + String.Format("BodyStyle : {0} ,", StrBodySyle.ToString)
                        If StrColorCode IsNot Nothing Then strMessage = strMessage + String.Format("ColorCode : {0} ,", StrColorCode.ToString)
                        If StrDriveSide IsNot Nothing Then strMessage = strMessage + String.Format("DriveSide : {0} ,", StrDriveSide.ToString)
                        If StrTeamName IsNot Nothing Then strMessage = strMessage + String.Format("TeamName : {0} ,", StrTeamName.ToString)
                        If StrRemarks IsNot Nothing Then strMessage = strMessage + String.Format("Remark : {0} ,", StrRemarks.ToString)
                        If StrShippingToCustomerDate IsNot Nothing Then strMessage = strMessage + String.Format("ShippingToCustomer : {0} ,", StrShippingToCustomerDate.ToString)
                        If strCustomerRequiredDate IsNot Nothing Then strMessage = strMessage + String.Format("CustomerRequiredDate : {0} ,", strCustomerRequiredDate.ToString)
                        If strRigCustomerPickDate IsNot Nothing Then strMessage = strMessage + String.Format("RigCustomerPickDate : {0} ,", strRigCustomerPickDate.ToString)
                        If StrTbNumberPrefix IsNot Nothing Then strMessage = strMessage + String.Format("TBNumberPrefix : {0} ,", StrTbNumberPrefix.ToString)
                        If StrBuildId IsNot Nothing Then strMessage = strMessage + String.Format("BuildId : {0} ,", StrBuildId.ToString)
                        If StrTagNumber IsNot Nothing Then strMessage = strMessage + String.Format("TagName : {0} ,", StrTagNumber.ToString)
                        If StrPaintFacility IsNot Nothing Then strMessage = strMessage + String.Format("PaintFacility : {0} ,", StrPaintFacility.ToString)


                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_ProgramInfo, Pe02, pe45, String.Format("Change program info in vehicle {0} {1}.", pe45, strMessage), MainBuildType, transaction, conTnd)

                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        If StrColorCode IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitColorCodeUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@ColorCode", SqlDbType.NVarChar, 50).Value = If(StrColorCode.ToString = String.Empty, DBNull.Value, StrColorCode.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrBodySyle IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitBodySyleUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@BodySyle", SqlDbType.NVarChar, 50).Value = If(StrBodySyle.ToString = String.Empty, DBNull.Value, StrBodySyle.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrDriveSide IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitDriveSideUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@DriveSide", SqlDbType.NVarChar, 50).Value = If(StrDriveSide.ToString = String.Empty, DBNull.Value, StrDriveSide.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrDedicated IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitDedicatedUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@Dedicated", SqlDbType.NVarChar, 50).Value = If(StrDedicated.ToString = String.Empty, DBNull.Value, StrDedicated.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrEmissionStage IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitEmissionStageUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@EmissionStage", SqlDbType.NVarChar, 50).Value = If(StrEmissionStage.ToString = String.Empty, DBNull.Value, StrEmissionStage.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrRemarks IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitRemarksUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@Remarks", SqlDbType.NVarChar, 100).Value = If(StrRemarks.ToString = String.Empty, DBNull.Value, StrRemarks.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrShippingToCustomerDate IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitShippingToCustomerDateUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@ShippingToCustomerDate", SqlDbType.NVarChar, 10).Value = If(StrShippingToCustomerDate.ToString = String.Empty, DBNull.Value, StrShippingToCustomerDate.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()
                        ElseIf strRigCustomerPickDate IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitRigCustomerPickDateUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@RigCustomerPickDate", SqlDbType.NVarChar, 10).Value = If(strRigCustomerPickDate.ToString = String.Empty, DBNull.Value, strRigCustomerPickDate.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()
                        ElseIf strCustomerRequiredDate IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitCustomerRequiredDateUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@CustomerRequiredDate", SqlDbType.NVarChar, 10).Value = If(strCustomerRequiredDate.ToString = String.Empty, DBNull.Value, strCustomerRequiredDate.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()
                        ElseIf StrCBG IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitCBGUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@CBG", SqlDbType.NVarChar, 18).Value = If(StrCBG.ToString = String.Empty, DBNull.Value, StrCBG.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrBuildId IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitBuildIdUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@BuildId", SqlDbType.NVarChar, 8).Value = If(StrBuildId.ToString = String.Empty, DBNull.Value, StrBuildId.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrPaintFacility IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitPaintFacilityUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@PaintFacility", SqlDbType.NVarChar, 100).Value = If(StrPaintFacility.ToString = String.Empty, DBNull.Value, StrPaintFacility.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrTagNumber IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitTagNumberUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@TagNumber", SqlDbType.NVarChar, 7).Value = If(StrTagNumber.ToString = String.Empty, DBNull.Value, StrTagNumber.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()

                        ElseIf StrTbNumberPrefix IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitTbNumberPrefixUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@HealthChartId", SqlDbType.BigInt, 8).Value = HealthChartId
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus

                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@TbNumberPrefix", SqlDbType.NVarChar, 4).Value = If(StrTbNumberPrefix.ToString = String.Empty, DBNull.Value, StrTbNumberPrefix.ToString())
                            '-------------------------------------------------
                            ' This value is passed to DB but It will not be save
                            ' It's only for compairing
                            '-------------------------------------------------

                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId
                            command.Parameters.Add("@ValidationOutput", SqlDbType.Bit, 1).Value = DBNull.Value
                            command.Parameters("@ValidationOutput").Direction = ParameterDirection.InputOutput

                            command.ExecuteScalar()
                            If command.Parameters("@ValidationOutput").Value = True Then
                                Throw New Exception("Sorry! the Vehicle number and prefix you entered is already in use. Please enter unique Vehicle number and prefix")
                            End If

                        ElseIf StrTBNumber IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitTBNumberUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@HealthChartId", SqlDbType.BigInt, 8).Value = HealthChartId
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus

                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            '-------------------------------------------------
                            ' This value is passed to DB but It will not be save
                            ' It's only for compairing
                            '-------------------------------------------------

                            command.Parameters.Add("@TBNumber", SqlDbType.NVarChar, 4).Value = If(StrTBNumber.ToString = String.Empty, DBNull.Value, StrTBNumber.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId
                            command.Parameters.Add("@ValidationOutput", SqlDbType.Bit, 1).Value = DBNull.Value
                            command.Parameters("@ValidationOutput").Direction = ParameterDirection.InputOutput

                            command.ExecuteScalar()
                            If command.Parameters("@ValidationOutput").Value = True Then
                                Throw New Exception("Sorry! the Vehicle number and prefix you entered is already in use. Please enter unique Vehicle number and prefix")
                            End If

                        ElseIf StrVin IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitVINUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@HealthChartId", SqlDbType.BigInt, 8).Value = HealthChartId
                            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@Vin", SqlDbType.NVarChar, 17).Value = If(StrVin.ToString = String.Empty, DBNull.Value, StrVin.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId
                            command.Parameters.Add("@ValidationOutput", SqlDbType.Bit, 1).Value = DBNull.Value
                            command.Parameters("@ValidationOutput").Direction = ParameterDirection.InputOutput

                            command.ExecuteScalar()
                            If command.Parameters("@ValidationOutput").Value = True Then
                                Throw New Exception("Sorry! the VIN number you entered is already in use. Please enter unique VIN number")
                            End If

                        ElseIf StrTeamName IsNot Nothing And StrXccTeamName IsNot Nothing Then

                            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_UnitXCCTeaMAndTeamNameUpdate.ToString())
                            command.Connection = conTnd
                            command.Transaction = transaction
                            command.CommandType = CommandType.StoredProcedure
                            command.Parameters.Add("@pe03_TnDProgramVehicles_PK", SqlDbType.BigInt, 8).Value = Pe03
                            command.Parameters.Add("@XCCPrototypeUser", SqlDbType.NVarChar, 50).Value = If(StrXccTeamName.ToString = String.Empty, DBNull.Value, StrXccTeamName.ToString())
                            command.Parameters.Add("@Translation", SqlDbType.NVarChar, 50).Value = If(StrTeamName.ToString = String.Empty, DBNull.Value, StrTeamName.ToString())
                            command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                            command.ExecuteScalar()
                        End If



                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        ChangeInfoII = True

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
                        ChangeInfoII = False

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
                ChangeInfoII = False

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
                        MessagePassing = New MessagePassing()
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
