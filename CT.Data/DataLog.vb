
Imports System.Data
Imports System.Data.SqlClient


''' <summary>
''' Error code : 20
''' Table : pe62_ProgramChangelog
''' </summary>
Public Class DataLog
    Inherits CtBaseClass

    'Private _tbAnswer As DataTable = Nothing
    'Private _arrayDT As String(,) = Nothing

    Public Function GetDataLog(pe02 As Long) As String(,)
        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeLogentrySelectByPlan.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02

                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)

                End Using

            End Using

            ConvertDataTableToStingArray()
            DataCenter.GlobalValues.message = String.Empty
            GetDataLog = _arrayDT

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
                    ErrorId = DataCenter.ErrorCenter.DataLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetDataLog = Nothing
        End Try

    End Function

    Public Function DeleteChangeLogentry(pe62 As Long) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeLogentryDelete.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe62_ProgramChangelog", SqlDbType.BigInt, 8).Value = pe62

                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    DeleteChangeLogentry = True

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
                            ErrorId = DataCenter.ErrorCenter.DataLog
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    DeleteChangeLogentry = False

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
                    ErrorId = DataCenter.ErrorCenter.DataLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            DeleteChangeLogentry = False

        End Try

    End Function


    Public Function AddChangeLog(pe02 As Long, ChangeDate As String, HCID As String, TnDIssue As String, BuildType As String, UnitId As String, ChangeDescription As String, Requestor As String, TnDResponsible As String) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeLogentryAdd.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    If ChangeDate Is Nothing Then
                        command.Parameters.Add("@Date", SqlDbType.NVarChar, 100).Value = DBNull.Value
                    Else
                        command.Parameters.Add("@Date", SqlDbType.NVarChar, 100).Value = ChangeDate
                    End If
                    command.Parameters.Add("@HCID", SqlDbType.NVarChar, 100).Value = HCID
                    command.Parameters.Add("@TnDIssue", SqlDbType.NVarChar, 100).Value = TnDIssue
                    command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 100).Value = BuildType
                    command.Parameters.Add("@UnitId", SqlDbType.NVarChar, 100).Value = UnitId
                    command.Parameters.Add("@ChangeDescription", SqlDbType.NVarChar, 500).Value = ChangeDescription
                    command.Parameters.Add("@Requestor", SqlDbType.NVarChar, 100).Value = Requestor
                    command.Parameters.Add("@TnDResponsible", SqlDbType.NVarChar, 100).Value = TnDResponsible

                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    AddChangeLog = True

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
                            ErrorId = DataCenter.ErrorCenter.DataLog
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    AddChangeLog = False


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
                    ErrorId = DataCenter.ErrorCenter.DataLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            AddChangeLog = False

        End Try

    End Function

    Public Function UpdateChangeLog(pe62 As Long, ChangeDate As String, HCID As String, TnDIssue As String, BuildType As String, UnitId As String, ChangeDescription As String, Requestor As String, TnDResponsible As String) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_ChangeLogentryUpdate.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe62_ProgramChangelog", SqlDbType.BigInt, 8).Value = pe62
                    If ChangeDate Is Nothing Then
                        command.Parameters.Add("@Date", SqlDbType.NVarChar, 100).Value = DBNull.Value
                    Else
                        command.Parameters.Add("@Date", SqlDbType.NVarChar, 100).Value = ChangeDate 'SqlDbType.Date, 3).Value = ChangeDate
                    End If
                    command.Parameters.Add("@HCID", SqlDbType.NVarChar, 100).Value = HCID
                    command.Parameters.Add("@TnDIssue", SqlDbType.NVarChar, 100).Value = TnDIssue
                    command.Parameters.Add("@BuildType", SqlDbType.NVarChar, 100).Value = BuildType
                    command.Parameters.Add("@UnitId", SqlDbType.NVarChar, 100).Value = UnitId
                    command.Parameters.Add("@ChangeDescription", SqlDbType.NVarChar, 500).Value = ChangeDescription
                    command.Parameters.Add("@Requestor", SqlDbType.NVarChar, 100).Value = Requestor
                    command.Parameters.Add("@TnDResponsible", SqlDbType.NVarChar, 100).Value = TnDResponsible

                    command.ExecuteNonQuery()
                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    UpdateChangeLog = True

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
                            ErrorId = DataCenter.ErrorCenter.DataLog
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    UpdateChangeLog = False

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
                    ErrorId = DataCenter.ErrorCenter.DataLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            UpdateChangeLog = False

        End Try

    End Function


    ''' <summary>
    ''' The method converts the DataTable which is returned from Database to string Array 
    ''' because only array string can be written in a Excel Range.
    ''' </summary>
    Protected Overrides Sub ConvertDataTableToStingArray()

        Dim i, j As Integer
        If _tbAnswer IsNot Nothing Then

            ReDim _arrayDT(_tbAnswer.Rows.Count, _tbAnswer.Columns.Count)
            For i = 0 To _tbAnswer.Rows.Count - 1
                For j = 0 To _tbAnswer.Columns.Count - 1
                    _arrayDT(i, j) = _tbAnswer.Rows(i)(j).ToString()
                Next j
            Next i
        End If
    End Sub

End Class
