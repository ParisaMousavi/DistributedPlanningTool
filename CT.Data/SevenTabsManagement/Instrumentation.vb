

Imports System.Data
Imports System.Data.SqlClient
Namespace SevenTabsManagement
    Public Class Instrumentation
        Inherits CtBaseClass


        'Public Function GetPlanData(pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,)

        '    Try

        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = Nothing
        '            If MainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedure.A2_VehicleAnd7Tabs_Specific_InstrumentationPartial_Vehicle.ToString())
        '            ElseIf MainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A2_VehicleAnd7Tabs_Rig_Specific_InstrumentationPartial.ToString())
        '            End If
        '            command.Connection = conTnd
        '            command.CommandType = CommandType.StoredProcedure
        '            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
        '            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = pe02
        '            command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
        '            command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq

        '            _tbAnswer = Nothing
        '            _arrayDT = Nothing
        '            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
        '                _tbAnswer = New DataTable()
        '                dataAdapter.Fill(_tbAnswer)
        '            End Using

        '        End Using

        '        ConvertDataTableToStingArray()
        '        DataCenter.GlobalValues.message = String.Empty
        '        GetPlanData = _arrayDT
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
        '                ErrorId = DataCenter.ErrorCenter.Instrumentation
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
        '        GetPlanData = Nothing
        '    End Try

        'End Function

        'Public Function GetTndPlanHeader(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,)

        '    Try

        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = Nothing
        '            If BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedure.A1_Header_Specific_InstrumentationPartial_Vehicle.ToString())
        '            ElseIf BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A1_Header_Rig_Specific_InstrumentationPartial.ToString())
        '            End If

        '            command.Connection = conTnd
        '            command.CommandType = CommandType.StoredProcedure
        '            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
        '            command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
        '            command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 4).Value = BuildPhase
        '            command.Parameters.Add("@BuildTypes", SqlDbType.NVarChar, 10).Value = BuildType

        '            _tbAnswer = Nothing
        '            _arrayDT = Nothing
        '            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
        '                _tbAnswer = New DataTable()
        '                dataAdapter.Fill(_tbAnswer)
        '            End Using

        '        End Using

        '        ConvertDataTableToStingArray()
        '        DataCenter.GlobalValues.message = String.Empty
        '        GetTndPlanHeader = _arrayDT

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
        '                ErrorId = DataCenter.ErrorCenter.Instrumentation
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
        '        GetTndPlanHeader = Nothing
        '    End Try

        'End Function




        Public Function Delete(pe01 As Long, pe02 As Long, HCID As Integer, InstrumentationList As String, Section As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicInstrumentationDelete.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_fK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@InstrumentationList", SqlDbType.NVarChar, 75).Value = InstrumentationList
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 25).Value = Section

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeletedInstrumentation, pe02, Nothing, String.Format(".NET DeleteColumn {0} from Instrumentation .", InstrumentationList), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteNonQuery()

                        DataCenter.GlobalValues.message = String.Empty
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
                                ErrorId = DataCenter.ErrorCenter.Instrumentation
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
                        ErrorId = DataCenter.ErrorCenter.Instrumentation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Delete = False
            End Try

        End Function


        Public Function AddColumn(pe01 As Long, pe02 As Long, HCID As Integer, InstrumentationList As String, Section As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicInstrumentationAddColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@InstrumentationList", SqlDbType.NVarChar, 75).Value = InstrumentationList
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 25).Value = Section

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_NewInstrumentation, pe02, Nothing, String.Format(".Net AddColumn to Instrumentation {0}.", InstrumentationList), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If
                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteNonQuery()
                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        AddColumn = True

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
                                ErrorId = DataCenter.ErrorCenter.Instrumentation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        AddColumn = False
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
                        ErrorId = DataCenter.ErrorCenter.Instrumentation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                AddColumn = False
            End Try

        End Function

        Public Function EditColumn(pe01 As Long, pe02 As Long, HCID As Integer, InstrumentationList As String, InstrumentationListNew As String, Section As String, SectionNew As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicInstrumentationEditColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_pK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@InstrumentationList", SqlDbType.NVarChar, 75).Value = InstrumentationList
                        command.Parameters.Add("@InstrumentationListNew", SqlDbType.NVarChar, 75).Value = InstrumentationListNew
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 25).Value = Section
                        command.Parameters.Add("@SectionNew", SqlDbType.NVarChar, 25).Value = SectionNew


                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedColInstrumentation, pe02, Nothing, String.Format(".Net EditColumn of Instrumentation from {0} to {1}.", InstrumentationList, InstrumentationListNew), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteNonQuery()
                        DataCenter.GlobalValues.message = String.Empty
                        transaction.Commit()
                        EditColumn = True

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
                                ErrorId = DataCenter.ErrorCenter.Instrumentation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        EditColumn = False

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
                        ErrorId = DataCenter.ErrorCenter.Instrumentation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                EditColumn = False

            End Try

        End Function


        ''' <summary>
        ''' For multi copy & paste update value.
        ''' </summary>
        ''' <param name="InstrumentationDataList"></param>
        ''' <returns></returns>
        Public Function UpdateData(InstrumentationDataList As List(Of DataCenter.InstrumentationData), MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Dim _InstrumentationDataList As List(Of DataCenter.InstrumentationData)
            Dim _pe02, _pe45 As Long

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        While InstrumentationDataList.Count >= 1
                            _pe02 = InstrumentationDataList(0).pe02
                            _pe45 = InstrumentationDataList(0).pe45
                            _InstrumentationDataList = InstrumentationDataList.FindAll(Function(ins) ins.pe02 = _pe02 And ins.pe45 = _pe45)

                            '------------------------------------------------------------------
                            ' This code portion is very important for Undo Please Deactive
                            ' Parisa
                            '------------------------------------------------------------------
                            changelog = New ChangeLog()
                            ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditInstrumentation, _pe02, _pe45, String.Format(".Net Instrumentation of (pe02,pe45) : ({0}, {1}) had new values.", _pe02, _pe45), MainBuildType, transaction, conTnd)
                            If ActionId = -1 Then
                                Throw New Exception("The ActionID must not be -1.")
                            End If


                            For Each _InstrumentationData As DataCenter.InstrumentationData In _InstrumentationDataList


                                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicInstrumentationUpdateData.ToString())
                                command.Connection = conTnd
                                command.Transaction = transaction
                                command.CommandType = CommandType.StoredProcedure
                                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = _InstrumentationData.pe02
                                command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = _InstrumentationData.pe45
                                command.Parameters.Add("@InstrumentationList", SqlDbType.NVarChar, 75).Value = _InstrumentationData.InstrumentationList
                                command.Parameters.Add("@Section", SqlDbType.NVarChar, 25).Value = _InstrumentationData.Section
                                command.Parameters.Add("@InstrumentationData", SqlDbType.NVarChar, 200).Value = IIf(_InstrumentationData.InstrumentationData = Nothing, DBNull.Value, _InstrumentationData.InstrumentationData)
                                command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                                command.ExecuteNonQuery()

                                InstrumentationDataList.Remove(_InstrumentationData)


                            Next
                        End While
                        DataCenter.GlobalValues.message = String.Empty
                        transaction.Commit()
                        UpdateData = True

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
                                ErrorId = DataCenter.ErrorCenter.Instrumentation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        UpdateData = False

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
                        ErrorId = DataCenter.ErrorCenter.Instrumentation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False
            End Try

        End Function

        ''' <summary>
        ''' Only for one cell value update.
        ''' </summary>
        ''' <param name="pe02"></param>
        ''' <param name="pe45"></param>
        ''' <param name="Section"></param>
        ''' <param name="InstrumentationList"></param>
        ''' <param name="InstrumentationData"></param>
        ''' <returns></returns>
        Public Function UpdateData(pe02 As Long, pe45 As Long, Section As String, InstrumentationList As String, InstrumentationData As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicInstrumentationUpdateData.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@InstrumentationList", SqlDbType.NVarChar, 75).Value = InstrumentationList
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 25).Value = Section
                        command.Parameters.Add("@InstrumentationData", SqlDbType.NVarChar, 200).Value = IIf(InstrumentationData = Nothing, DBNull.Value, InstrumentationData)

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditInstrumentation, pe02, pe45, String.Format(".Net Instrumentation of (pe02,pe45) : ({0}, {1}) had new value in -> Section : {2} -> Header : {3} -> value : {4}.", pe02, pe45, Section, InstrumentationList, InstrumentationData), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                        command.ExecuteNonQuery()

                        DataCenter.GlobalValues.message = String.Empty
                        transaction.Commit()
                        UpdateData = True

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
                                ErrorId = DataCenter.ErrorCenter.Instrumentation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        UpdateData = False

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
                        ErrorId = DataCenter.ErrorCenter.Instrumentation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False
            End Try

        End Function


    End Class

End Namespace
