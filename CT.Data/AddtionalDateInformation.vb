
Imports System.Data
Imports System.Data.SqlClient

Public Class AddtionalDateInformation
    Inherits CtBaseClass



    ''' <summary>
    ''' The SelectAllDateInformation function fetches all the inserted HCID.
    ''' </summary>
    ''' <param name="pe02"></param>
    ''' <param name="HealthChartId"></param>
    ''' <param name="AssyMrd"></param>
    ''' <param name="Firstm1"></param>
    ''' <param name="M1DC"></param>
    ''' <param name="FirstVP"></param>
    ''' <param name="PEC"></param>
    ''' <param name="FEC"></param>
    ''' <param name="JobOne"></param>
    ''' <param name="DateBackRGB"></param>
    ''' <param name="DateFontRGB"></param>
    ''' <returns></returns>
    Public Function AddHealthChartID(pe02 As Long, HealthChartId As Integer, AssyMrd As Object, Firstm1 As Object, M1DC As Object, FirstVP As Object, PEC As Object, FEC As Object, JobOne As Object, DateBackRGB As String, DateFontRGB As String, MainBuildType As String) As Boolean


        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_AddtionalDateInformationAdd.ToString())

                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealthChartId

                    command.Parameters.Add("@AssyMrd", SqlDbType.Date, 3).Value = If(AssyMrd IsNot Nothing, Date.Parse(AssyMrd), DBNull.Value)
                    command.Parameters.Add("@Firstm1", SqlDbType.Date, 3).Value = If(Firstm1 IsNot Nothing, Date.Parse(Firstm1), DBNull.Value)
                    command.Parameters.Add("@M1DC", SqlDbType.Date, 3).Value = If(M1DC IsNot Nothing, Date.Parse(M1DC), DBNull.Value)
                    command.Parameters.Add("@FirstVP", SqlDbType.Date, 3).Value = If(FirstVP IsNot Nothing, Date.Parse(FirstVP), DBNull.Value)
                    command.Parameters.Add("@PEC", SqlDbType.Date, 3).Value = If(PEC IsNot Nothing, Date.Parse(PEC), DBNull.Value)
                    command.Parameters.Add("@FEC", SqlDbType.Date, 3).Value = If(FEC IsNot Nothing, Date.Parse(FEC), DBNull.Value)
                    command.Parameters.Add("@Job#1", SqlDbType.Date, 3).Value = If(JobOne IsNot Nothing, Date.Parse(JobOne), DBNull.Value)

                    command.Parameters.Add("@DateBackRGB", SqlDbType.NVarChar, 9).Value = DateBackRGB
                    command.Parameters.Add("@DateFontRGB", SqlDbType.NVarChar, 9).Value = DateFontRGB

                    command.ExecuteNonQuery()

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    AddHealthChartID = True

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
                            ErrorId = DataCenter.ErrorCenter.AddtionalDateInformation
                    End Select
                    DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex0.Message)
                    transaction.Rollback()
                    AddHealthChartID = False

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
            AddHealthChartID = False

        End Try

    End Function



    Public Function DeleteHealthChartID(pe67 As Long, pe02 As Long, HealthChartId As Integer) As Boolean



        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_AddtionalDateInformationDelete.ToString())

                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe67_AddtionalDateInformation_PK", SqlDbType.BigInt, 8).Value = pe67
                    command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealthChartId

                    command.ExecuteNonQuery()

                    DataCenter.GlobalValues.message = String.Empty
                    transaction.Commit()
                    DeleteHealthChartID = True

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
                    DeleteHealthChartID = False

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
            DeleteHealthChartID = False

        End Try

    End Function


    ''' <summary>
    ''' The HCIDs, which are inserted with AddHealthChartID function are listed in output. 
    ''' </summary>
    ''' <param name="pe02"></param>
    ''' <param name="HCID">This is optional but as type integer</param>
    ''' <returns></returns>
    Public Function SelectAllDateInformation(pe02 As Long, MainBuildType As String, Optional HCID As Object = Nothing) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_AddtionalDateInformationSelectByPlan.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe02_TnDProgramDetails_FK", SqlDbType.BigInt, 8).Value = pe02
                command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 20).Value = MainBuildType
                command.Parameters.Add("@HCID", SqlDbType.Int, 4).Value = If(HCID Is Nothing, DBNull.Value, HCID)

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)
                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectAllDateInformation = _tbAnswer

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
            SelectAllDateInformation = Nothing

        End Try

    End Function




    Public Function UpdateHealthChartID(pe67 As Long, HealthChartId As Integer, AssyMrd As Date, Firstm1 As Date, M1DC As Date, FirstVP As Date, PEC As Date, FEC As Date, JobOne As Date, DateBackRGB As String, DateFontRGB As String) As Boolean

        Dim transaction As SqlTransaction = Nothing

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Try

                    conTnd.Open()
                    transaction = conTnd.BeginTransaction()

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_AddtionalDateInformationUpdate.ToString())

                    command.Connection = conTnd
                    command.Transaction = transaction
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@pe67_AddtionalDateInformation_PK", SqlDbType.BigInt, 8).Value = pe67
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealthChartId
                    command.Parameters.Add("@AssyMrd", SqlDbType.Date, 3).Value = AssyMrd
                    command.Parameters.Add("@Firstm1", SqlDbType.Date, 3).Value = Firstm1
                    command.Parameters.Add("@M1DC", SqlDbType.Date, 3).Value = M1DC
                    command.Parameters.Add("@FirstVP", SqlDbType.Date, 3).Value = FirstVP
                    command.Parameters.Add("@PEC", SqlDbType.Date, 3).Value = PEC
                    command.Parameters.Add("@FEC", SqlDbType.Date, 3).Value = FEC
                    command.Parameters.Add("@Job#1", SqlDbType.Date, 3).Value = DBNull.Value ' this value must not be inseted from User
                    command.Parameters.Add("@DateBackRGB", SqlDbType.NVarChar, 9).Value = DateBackRGB
                    command.Parameters.Add("@DateFontRGB", SqlDbType.NVarChar, 9).Value = DateFontRGB

                    command.ExecuteNonQuery()

                    transaction.Commit()
                    DataCenter.GlobalValues.message = String.Empty
                    UpdateHealthChartID = True

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
                    UpdateHealthChartID = False

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
            UpdateHealthChartID = False

        End Try

    End Function

End Class
