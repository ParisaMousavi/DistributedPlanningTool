﻿Imports System.Data
Imports System.Data.SqlClient

Namespace VehiclePlan.SevenTabs
    Public Class ProgramInformation
        Inherits CtBaseClass
        Implements CT.Data.Interfaces.ProgramInformationInterface


        Public Function GetPlanData(pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,) Implements Interfaces.ProgramInformationInterface.GetPlanData

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing
                    command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.A2_VehicleAnd7Tabs_Vehicle_Specific_ProgramInformationPartial.ToString())

                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = pe02
                    command.Parameters.Add("@UpperBoundDisplaySeq", SqlDbType.Int, 4).Value = UpperBoundDisplaySeq
                    command.Parameters.Add("@LowerBoundDisplaySeq", SqlDbType.Int, 4).Value = LowerBoundDisplaySeq

                    _tbAnswer = Nothing
                    _arrayDT = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetPlanData = _arrayDT
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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetPlanData = Nothing
            End Try

        End Function


        Public Function GetTndPlanHeader(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,) Implements Interfaces.ProgramInformationInterface.GetTndPlanHeader

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = Nothing

                    If MainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                        command = New SqlCommand(DataCenter.StoredProcedures.VehiclePlan.A1_Header_Vehicle_Specific_ProgramInformationPartial.ToString())
                    ElseIf MainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                        command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A1_Header_Rig_Specific_ProgramInformationPartial.ToString())
                    End If

                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                    command.Parameters.Add("@BuildPhase", SqlDbType.NVarChar, 4).Value = BuildPhase
                    command.Parameters.Add("@BuildTypes", SqlDbType.NVarChar, 10).Value = BuildType

                    _tbAnswer = Nothing
                    _arrayDT = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                ConvertDataTableToStingArray()
                DataCenter.GlobalValues.message = String.Empty
                GetTndPlanHeader = _arrayDT

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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                GetTndPlanHeader = Nothing
            End Try

        End Function





        Public Function Delete(pe01 As Long, pe02 As Long, HCID As Integer, ProgramInformationList As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicProgramInformationDelete.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@ProgramInformationList", SqlDbType.NVarChar, 50).Value = ProgramInformationList


                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeletedProgramInfo, pe02, Nothing, String.Format(".Net Delete from ProgramInformation {0}.", ProgramInformationList), MainBuildType, transaction, conTnd)
                        If ActionId = -1 Then
                            Throw New Exception("The ActionID must not be -1.")
                        End If

                        command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId
                        command.ExecuteNonQuery()

                        transaction.Commit()
                        DataCenter.GlobalValues.message = String.Empty
                        Delete = True

                    Catch ex0 As Exception
                        'error has been ocurred by running command or open transaction

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
                                ErrorId = DataCenter.ErrorCenter.ProgramInformation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        Delete = False

                    End Try

                End Using

            Catch ex As Exception
                'It means an error has been occured by opening connection

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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Delete = False

            End Try


        End Function

        Public Function AddColumn(pe01 As Long, pe02 As Long, HCID As Integer, ProgramInformationList As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicProgramInformationAddColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@ProgramInformationList", SqlDbType.NVarChar, 50).Value = ProgramInformationList

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_NewProgramInfo, pe02, Nothing, String.Format(".Net AddColumn to ProgramInformation {0}.", ProgramInformationList), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.ProgramInformation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        AddColumn = False

                    End Try

                End Using

            Catch ex As Exception
                'It means an error has been occured
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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                AddColumn = False

            End Try

        End Function

        Public Function EditColumn(pe01 As Long, pe02 As Long, HCID As Integer, ProgramInformationList As String, ProgramInformationListNew As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicProgramInformationEditColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.Int, 4).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@ProgramInformationList", SqlDbType.NVarChar, 50).Value = ProgramInformationList
                        command.Parameters.Add("@ProgramInformationListNew", SqlDbType.NVarChar, 50).Value = ProgramInformationListNew

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedColProgramInfo, pe02, Nothing, String.Format(".Net EditColumn of ProgramInformation from {0} to {1}.", ProgramInformationList, ProgramInformationListNew), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.ProgramInformation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        EditColumn = False

                    End Try

                End Using

            Catch ex As Exception
                'It means an error has been occured

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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                EditColumn = False

            End Try

        End Function

        Public Function UpdateData(ProgramInformationDataList As List(Of DataCenter.ProgramInformationData), MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1



            Dim _ProgramInformationList As List(Of DataCenter.ProgramInformationData)
            Dim _pe02, _pe45 As Long

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        While ProgramInformationDataList.Count >= 1
                            _pe02 = ProgramInformationDataList(0).pe02
                            _pe45 = ProgramInformationDataList(0).pe45
                            _ProgramInformationList = ProgramInformationDataList.FindAll(Function(proi) proi.pe02 = _pe02 And proi.pe45 = _pe45)


                            '------------------------------------------------------------------
                            ' This code portion is very important for Undo Please Deactive
                            ' Parisa
                            '------------------------------------------------------------------
                            changelog = New ChangeLog()
                            ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedProgramInfo, _pe02, _pe45, String.Format(".Net ProgramInformation of (pe02,pe45) : ({0}, {1}) had new values.", _pe02, _pe45), MainBuildType, transaction, conTnd)
                            If ActionId = -1 Then
                                Throw New Exception("The ActionID must not be -1.")
                            End If

                            For Each _ProgramInformationData As DataCenter.ProgramInformationData In _ProgramInformationList


                                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicProgramInformationUpdateData.ToString())
                                command.Connection = conTnd
                                command.Transaction = transaction
                                command.CommandType = CommandType.StoredProcedure
                                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.Int, 4).Value = _ProgramInformationData.pe02
                                command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.Int, 4).Value = _ProgramInformationData.pe45
                                command.Parameters.Add("@ProgramInformationList", SqlDbType.NVarChar, 50).Value = _ProgramInformationData.ProgramInformationList
                                '50->200 characters
                                command.Parameters.Add("@ProgramInformationData", SqlDbType.NVarChar, 200).Value = IIf(_ProgramInformationData.ProgramInformationData = Nothing, DBNull.Value, _ProgramInformationData.ProgramInformationData)
                                command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                                command.ExecuteNonQuery()

                                ProgramInformationDataList.Remove(_ProgramInformationData)
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
                                ErrorId = DataCenter.ErrorCenter.ProgramInformation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        UpdateData = False

                    End Try

                End Using

            Catch ex As Exception

                'It means an error has been occured

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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False

            End Try

        End Function


        Public Function UpdateData(pe02 As Long, pe45 As Long, ProgramInformationList As String, ProgramInformationData As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicProgramInformationUpdateData.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.Int, 4).Value = pe02
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.Int, 4).Value = pe45
                        command.Parameters.Add("@ProgramInformationList", SqlDbType.NVarChar, 50).Value = ProgramInformationList
                        '50->200 characters
                        command.Parameters.Add("@ProgramInformationData", SqlDbType.NVarChar, 200).Value = IIf(ProgramInformationData = Nothing, DBNull.Value, ProgramInformationData)

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedProgramInfo, pe02, pe45, String.Format(".Net UpdateColumn {0} of ProgramInformation  = {1}.", ProgramInformationList, ProgramInformationData), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.ProgramInformation
                        End Select
                        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                        transaction.Rollback()
                        UpdateData = False

                    End Try

                End Using

            Catch ex As Exception

                'It means an error has been occured
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
                        ErrorId = DataCenter.ErrorCenter.ProgramInformation
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False

            End Try

        End Function


    End Class
End Namespace
