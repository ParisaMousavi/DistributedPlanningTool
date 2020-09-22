Imports System.Data
Imports System.Data.SqlClient
''' <summary>
''' Error code : 10 
''' </summary>
Public Class ChangeLog

    Private _tbAnswer As DataTable = Nothing
    Private _arrayDT As String(,) = Nothing
    ''' <summary>
    ''' 
    ''' Create a new ActionID which must be used to save
    ''' a chage step in database for implementing Undo.
    ''' For easier usa we have implemented pe61 as num that the develper uses only from list
    ''' </summary>
    ''' <param name="pe61"></param>
    ''' <param name="pe02"></param>
    ''' <param name="pe45"></param>
    ''' <param name="Remark"></param>
    ''' <returns>The return value is ActionID. It can be a Nothing/Null or a ActionID.</returns>
    Public Function AddChangeLog(pe61 As DataCenter.ActionName, pe02 As Long, pe45 As Long, Remark As String, MainBuildType As String, Optional transaction As SqlTransaction = Nothing, Optional conTnd As SqlConnection = Nothing) As Long

        Dim ActionId As Long = 1
        Dim IsRunningLocal As Boolean = Nothing
        Try
            If conTnd.State <> ConnectionState.Open Then
                conTnd = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                conTnd.Open() ' it must be here because of BeginTransaction
                IsRunningLocal = True
            End If

            If transaction Is Nothing Then
                transaction = conTnd.BeginTransaction()
                IsRunningLocal = True
            End If

            Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Undo_AddChangeLog.ToString())
            command.Connection = conTnd
            command.Transaction = transaction
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add("@pe61_ActionName_FK", SqlDbType.Int, 4).Value = Integer.Parse(pe61)
            command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
            command.Parameters.Add("@TableName", SqlDbType.NVarChar, 255).Value = Nothing
            command.Parameters.Add("@ColumnName", SqlDbType.NVarChar, 255).Value = Nothing
            command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
            command.Parameters.Add("@pe45_AllocatedPowerPack_FK", SqlDbType.BigInt, 8).Value = pe45
            command.Parameters.Add("@Remark", SqlDbType.NVarChar, 255).Value = Remark
            command.Parameters.Add("@ActionId", SqlDbType.BigInt, 8).Value = ActionId
            command.Parameters("@ActionId").Direction = ParameterDirection.InputOutput


            If IsRunningLocal = True Then
                command.ExecuteScalar()
                AddChangeLog = command.Parameters("@ActionId").Value
                transaction.Commit()
                conTnd.Close()
            Else
                command.ExecuteScalar()
                AddChangeLog = command.Parameters("@ActionId").Value
            End If


            DataCenter.GlobalValues.message = String.Empty

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
                    ErrorId = DataCenter.ErrorCenter.ChangeLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            transaction.Rollback()
            conTnd.Close()
            AddChangeLog = -1

        End Try

    End Function

    ''' <summary>
    '''<para>
    '''Return Paramater : Pe60,Pe61,GroupName,ActionName,ActionId,User,CreatedOn,Remark 
    ''' </para>
    ''' </summary>
    ''' <param name="pe01"></param>
    ''' <param name="HealthChartID"></param>
    ''' <returns></returns>

    Public Function GetTnDLastUndo(pe01 As Long, HealthChartID As Integer, MainBuildType As String) As DataTable

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)



                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Undo_GetTnDLastUndo.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe01_TnDBasicProgram_FK", SqlDbType.BigInt, 8).Value = pe01
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HealthChartID", SqlDbType.Int, 4).Value = HealthChartID


                _tbAnswer = Nothing

                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using


                DataCenter.GlobalValues.message = String.Empty
                GetTnDLastUndo = _tbAnswer


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
                    ErrorId = DataCenter.ErrorCenter.ChangeLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetTnDLastUndo = Nothing

        End Try

    End Function


    ''' <summary>
    '''<para>
    '''Return Paramater : Yes / No
    ''' </para>
    ''' </summary>
    ''' <param name="pe02"></param>
    ''' <returns></returns>

    Public Function IsRedoCommandAvailable(pe02 As Long, HealthChartID As Integer, MainBuildType As String) As DataTable

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Undo_IsRedoCommandAvailable.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartID", SqlDbType.Int, 4).Value = HealthChartID
                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    _tbAnswer = Nothing

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    IsRedoCommandAvailable = _tbAnswer

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
                            ErrorId = DataCenter.ErrorCenter.ChangeLog
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    IsRedoCommandAvailable = Nothing

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
                    ErrorId = DataCenter.ErrorCenter.ChangeLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            IsRedoCommandAvailable = Nothing

        End Try

    End Function


    ''' <summary>
    '''<para>
    '''Return Paramater : Yes / No
    ''' </para>
    ''' </summary>
    ''' <param name="pe02"></param>
    ''' <param name="HealthChartID"></param>
    ''' <returns></returns>

    Public Function IsUndoCommandAvailable(pe02 As Long, HealthChartID As Integer, MainBuildType As String) As Boolean


        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)



                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Undo_IsUndoCommandAvailable.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HealthChartID", SqlDbType.Int, 4).Value = HealthChartID


                _tbAnswer = Nothing

                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using


                If _tbAnswer.Rows.Count = 1 Then

                    If _tbAnswer.Rows(0)("Answer").ToString() = "YES" Then
                        IsUndoCommandAvailable = True
                    Else
                        IsUndoCommandAvailable = False
                    End If

                Else

                    Throw New Exception("More than onw row in not allowed in IsUndoCommandAvailable. ")

                End If


                DataCenter.GlobalValues.message = String.Empty

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
                    ErrorId = DataCenter.ErrorCenter.ChangeLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            IsUndoCommandAvailable = False

        End Try

    End Function

    ''' <summary>
    '''<para>
    '''Return Paramater : Table value for Changelog
    ''' </para>
    ''' </summary>
    ''' <param name="pe01"></param>
    ''' <param name="HealthChartID"></param>
    ''' <returns></returns>
    Public Function UndoPreviousOperation(pe01 As Long, HealthChartID As Integer, MainBuildType As String) As DataTable

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Undo_PreviousOperation.ToString())
                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe01_TnDBasicProgram_ID", SqlDbType.BigInt, 8).Value = pe01
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealthChartID

                    _tbAnswer = Nothing

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using



                    If _tbAnswer.Columns(0).ColumnName = "Answer" Then Throw New Exception(_tbAnswer.Rows(0)("Answer").ToString)


                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    UndoPreviousOperation = _tbAnswer

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
                            ErrorId = DataCenter.ErrorCenter.ChangeLog
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    UndoPreviousOperation = Nothing

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
                    ErrorId = DataCenter.ErrorCenter.ChangeLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            UndoPreviousOperation = Nothing

        End Try

    End Function

    ''' <summary>
    '''<para>
    '''Return Paramater : Table Values
    ''' </para> 
    ''' </summary>
    ''' <param name="pe02"></param>
    ''' <returns></returns>

    Public Function RedoPreviousOperation(pe02 As Long, HCID As Integer, MainBuildType As String) As DataTable

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Undo_UndoPreviousOperation.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HCID


                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    _tbAnswer = Nothing

                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    RedoPreviousOperation = _tbAnswer

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
                            ErrorId = DataCenter.ErrorCenter.ChangeLog
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    RedoPreviousOperation = Nothing

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
                    ErrorId = DataCenter.ErrorCenter.ChangeLog
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            RedoPreviousOperation = Nothing

        End Try

    End Function


    ''' <summary>
    ''' The method converts the DataTable which is returned from Database to string Array 
    ''' because only array string can be written in a Excel Range.
    ''' </summary>
    Private Sub ConvertDataTableToStingArray()


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
