
Imports System.Data
Imports System.Data.SqlClient

Namespace SevenTabsManagement
    Public Class NonMfcSpecification
        Inherits CtBaseClass


        'Public Function GetPlanData(pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,)

        '    Try

        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = Nothing
        '            If MainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedure.A2_VehicleAnd7Tabs_Specific_NonMfcPartial_Vehicle.ToString())
        '            ElseIf MainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A2_VehicleAnd7Tabs_Rig_Specific_NonMfcPartial.ToString())
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
        '                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
        '        GetPlanData = Nothing
        '    End Try

        'End Function

        'Public Function GetTndPlanHeader(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,)

        '    Try

        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = Nothing

        '            If MainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedure.A1_Header_Specific_NonMfcPartial_Vehicle.ToString())
        '            ElseIf MainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A1_Header_Rig_Specific_NonMfcPartial.ToString())
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
        '                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
        '        GetTndPlanHeader = Nothing
        '    End Try

        'End Function




        Public Function Delete(pe01 As Long, pe02 As Long, HCID As Integer, NonMfcSpecification As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicNonMfcSpecificationDelete.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@NonMfcSpecification", SqlDbType.NVarChar, 100).Value = NonMfcSpecification

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeletedNonMFC, pe02, Nothing, String.Format(".Net Delete from NonMfcSpecification {0}.", NonMfcSpecification), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Delete = False

            End Try


        End Function


        Public Function AddColumn(pe01 As Long, pe02 As Long, HCID As Integer, NonMfcSpecification As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicNonMfcSpecificationAddColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@NonMfcSpecification", SqlDbType.NVarChar, 100).Value = NonMfcSpecification

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_NewNonMFC, pe02, Nothing, String.Format(".Net AddColumn to NonMfcSpecification {0}.", NonMfcSpecification), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                AddColumn = False

            End Try

        End Function

        Public Function EditColumn(pe01 As Long, pe02 As Long, HCID As Integer, NonMfcSpecification As String, NonMfcSpecificationNew As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicNonMfcSpecificationEditColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.Int, 4).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@NonMfcSpecification", SqlDbType.NVarChar, 100).Value = NonMfcSpecification
                        command.Parameters.Add("@NonMfcSpecificationNew", SqlDbType.NVarChar, 100).Value = NonMfcSpecificationNew

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedColNonMFC, pe02, Nothing, String.Format(".Net EditColumn of NonMfcSpecification from {0} to {1}.", NonMfcSpecification, NonMfcSpecificationNew), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                EditColumn = False

            End Try

        End Function
        Public Function UpdateData(NonMfcSpecificationDataList As List(Of DataCenter.NonMfcSpecificationData), MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Dim _NonMfcSpecificationList As List(Of DataCenter.NonMfcSpecificationData)
            Dim _pe02, _pe45 As Long

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        While NonMfcSpecificationDataList.Count >= 1
                            _pe02 = NonMfcSpecificationDataList(0).pe02
                            _pe45 = NonMfcSpecificationDataList(0).pe45
                            _NonMfcSpecificationList = NonMfcSpecificationDataList.FindAll(Function(nmfc) nmfc.pe02 = _pe02 And nmfc.pe45 = _pe45)


                            '------------------------------------------------------------------
                            ' This code portion is very important for Undo Please Deactive
                            ' Parisa
                            '------------------------------------------------------------------
                            changelog = New ChangeLog()
                            ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedNonMFC, _pe02, _pe45, String.Format(".Net NonMfcSpecification of (pe02,pe45) : ({0}, {1}) had new values.", _pe02, _pe45), MainBuildType, transaction, conTnd)
                            If ActionId = -1 Then
                                Throw New Exception("The ActionID must not be -1.")
                            End If

                            For Each _NonMfcSpecificationData As DataCenter.NonMfcSpecificationData In _NonMfcSpecificationList

                                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicNonMfcSpecificationUpdateData.ToString())
                                command.Connection = conTnd
                                command.Transaction = transaction
                                command.CommandType = CommandType.StoredProcedure
                                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = _NonMfcSpecificationData.pe02
                                command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = _NonMfcSpecificationData.pe45
                                command.Parameters.Add("@NonMfcSpecification", SqlDbType.NVarChar, 100).Value = _NonMfcSpecificationData.NonMfcSpecification
                                '75->200 characters
                                command.Parameters.Add("@NonMfcSpecificationData", SqlDbType.NVarChar, 200).Value = IIf(_NonMfcSpecificationData.NonMfcSpecificationData = Nothing, DBNull.Value, _NonMfcSpecificationData.NonMfcSpecificationData)
                                command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                                command.ExecuteNonQuery()

                                NonMfcSpecificationDataList.Remove(_NonMfcSpecificationData)
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
                                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False

            End Try

        End Function


        Public Function UpdateData(pe02 As Long, pe45 As Long, NonMfcSpecification As String, NonMfcSpecificationData As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicNonMfcSpecificationUpdateData.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@NonMfcSpecification", SqlDbType.NVarChar, 100).Value = NonMfcSpecification
                        '75->200 characters
                        command.Parameters.Add("@NonMfcSpecificationData", SqlDbType.NVarChar, 200).Value = IIf(NonMfcSpecificationData = Nothing, DBNull.Value, NonMfcSpecificationData)

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedNonMFC, pe02, pe45, String.Format(".Net Update value of {0} of NonMfcSpecification  = {1}.", NonMfcSpecification, NonMfcSpecificationData), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.NonMfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False

            End Try

        End Function

    End Class

End Namespace