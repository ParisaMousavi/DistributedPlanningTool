Imports System.Data
Imports System.Data.SqlClient

Namespace SevenTabsManagement
    Public Class MfcSpecification
        Inherits CtBaseClass


        'Public Function GetPlanData(pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,)

        '    Try

        '        Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

        '            Dim command As SqlCommand = Nothing
        '            If MainBuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedure.A2_VehicleAnd7Tabs_Specific_MfcPartial_Vehicle.ToString())
        '            ElseIf MainBuildType = CT.Data.DataCenter.BuildType.rig.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A2_VehicleAnd7Tabs_Rig_Specific_MfcPartial.ToString())
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
        '                ErrorId = DataCenter.ErrorCenter.MfcSpecification
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
        '                command = New SqlCommand(DataCenter.StoredProcedure.A1_Header_Specific_MfcPartial_Vehicle.ToString())
        '            ElseIf MainBuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
        '                command = New SqlCommand(DataCenter.StoredProcedures.RigPlan.A1_Header_Rig_Specific_MfcPartial.ToString())
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
        '                ErrorId = DataCenter.ErrorCenter.MfcSpecification
        '        End Select
        '        DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
        '        GetTndPlanHeader = Nothing
        '    End Try

        'End Function





        Public Function Delete(pe01 As Long, pe02 As Long, HCID As Integer, Mfc As String, Section As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try


                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicMfcSpecificationDelete.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@MFC", SqlDbType.NVarChar, 75).Value = Mfc
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 50).Value = Section


                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_DeletedMFC, pe02, Nothing, String.Format(".Net Delete : {0} from Mfc.", Mfc), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.MfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.MfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                Delete = False

            End Try

        End Function


        Public Function AddColumn(pe01 As Long, pe02 As Long, HCID As Integer, Mfc As String, Section As String, Description As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicMfcSpecificationAddColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_PK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@MFC", SqlDbType.NVarChar, 75).Value = Mfc
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 50).Value = Section
                        command.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = Description

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_NewMFC, pe02, Nothing, String.Format(".Net AddColumn : {0} to Mfc.", Mfc), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.MfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.MfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                AddColumn = False
            End Try

        End Function

        Public Function EditColumn(pe01 As Long, pe02 As Long, HCID As Integer, Mfc As String, NewMFC As String, NewDescription As String, Section As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicMfcSpecificationEditColumn.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                        command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                        command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID
                        command.Parameters.Add("@MFC", SqlDbType.NVarChar, 75).Value = Mfc
                        command.Parameters.Add("@NewMFC", SqlDbType.NVarChar, 75).Value = NewMFC
                        command.Parameters.Add("@NewDescription", SqlDbType.NVarChar, 200).Value = NewDescription
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 50).Value = Section


                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedColMFC, pe02, Nothing, String.Format(".Net EditColumn of Mfc from {0} to {1}.", Mfc, NewMFC), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.MfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.MfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                EditColumn = False
            End Try

        End Function


        Public Function UpdateData(MfcSpecificationDataList As List(Of DataCenter.MfcSpecificationData), MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Dim _MfcSpecificationDataList As List(Of DataCenter.MfcSpecificationData)
            Dim _pe02, _pe45 As Long


            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        While MfcSpecificationDataList.Count >= 1
                            _pe02 = MfcSpecificationDataList(0).pe02
                            _pe45 = MfcSpecificationDataList(0).pe45
                            _MfcSpecificationDataList = MfcSpecificationDataList.FindAll(Function(mfc) mfc.pe02 = _pe02 And mfc.pe45 = _pe45)

                            '------------------------------------------------------------------
                            ' This code portion is very important for Undo Please Deactive
                            ' Parisa
                            '------------------------------------------------------------------
                            changelog = New ChangeLog()
                            ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedMFC, _pe02, _pe45, String.Format(".NET MfcSpecification of (pe02,pe45) : ({0}, {1}) had new values.", _pe02, _pe45), MainBuildType, transaction, conTnd)
                            If ActionId = -1 Then
                                Throw New Exception("The ActionID must not be -1.")
                            End If

                            For Each _MfcSpecificationData As DataCenter.MfcSpecificationData In _MfcSpecificationDataList


                                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicMfcSpecificationUpdateData.ToString())
                                command.Connection = conTnd
                                command.Transaction = transaction
                                command.CommandType = CommandType.StoredProcedure
                                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = _MfcSpecificationData.pe02
                                command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = _MfcSpecificationData.pe45
                                command.Parameters.Add("@MFC", SqlDbType.NVarChar, 75).Value = _MfcSpecificationData.Mfc
                                command.Parameters.Add("@Section", SqlDbType.NVarChar, 50).Value = _MfcSpecificationData.Section
                                '50->200 characters
                                command.Parameters.Add("@Data", SqlDbType.NVarChar, 200).Value = IIf(_MfcSpecificationData.Data = Nothing, DBNull.Value, _MfcSpecificationData.Data)
                                command.Parameters.Add("@ActionID", SqlDbType.BigInt, 8).Value = ActionId

                                command.ExecuteNonQuery()

                                MfcSpecificationDataList.Remove(_MfcSpecificationData)

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
                                ErrorId = DataCenter.ErrorCenter.MfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.MfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False
            End Try

        End Function


        Public Function UpdateData(pe02 As Long, pe45 As Long, Mfc As String, Section As String, Data As String, MainBuildType As String) As Boolean

            Dim transaction As SqlTransaction = Nothing
            Dim changelog As ChangeLog = Nothing
            Dim ActionId As Long = -1

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Try

                        conTnd.Open()
                        transaction = conTnd.BeginTransaction()

                        Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_DynamicMfcSpecificationUpdateData.ToString())
                        command.Connection = conTnd
                        command.Transaction = transaction
                        command.CommandType = CommandType.StoredProcedure
                        command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                        command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
                        command.Parameters.Add("@MFC", SqlDbType.NVarChar, 75).Value = Mfc
                        command.Parameters.Add("@Section", SqlDbType.NVarChar, 50).Value = Section
                        '50->200 characters
                        command.Parameters.Add("@Data", SqlDbType.NVarChar, 200).Value = IIf(Data = Nothing, DBNull.Value, Data)

                        '------------------------------------------------------------------
                        ' This code portion is very important for Undo Please Deactive
                        ' Parisa
                        '------------------------------------------------------------------
                        changelog = New ChangeLog()
                        ActionId = changelog.AddChangeLog(DataCenter.ActionName.Tnd_EditedMFC, pe02, pe45, String.Format(".NET MfcSpecification of (pe02,pe45) : ({0}, {1}) had new values.", pe02, pe45), MainBuildType, transaction, conTnd)
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
                                ErrorId = DataCenter.ErrorCenter.MfcSpecification
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
                        ErrorId = DataCenter.ErrorCenter.MfcSpecification
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                UpdateData = False
            End Try

        End Function





    End Class

End Namespace
