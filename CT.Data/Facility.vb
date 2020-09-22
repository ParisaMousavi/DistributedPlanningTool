
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Error code : 40
''' </summary>
Public Class Facility
    Inherits CtBaseClass

    'Private _tbAnswer As DataTable = Nothing
    'Private _arrayDT As String(,) = Nothing



    ''' <summary>
    ''' Return column/s:
    '''  <para/>  FacilityCbg
    ''' </summary>
    ''' <param name="FacilityCbg"></param>
    ''' <param name="FacilityLocation"></param>
    ''' <param name="FacilityName"></param>
    ''' <param name="SubFacilityName"></param>
    ''' <returns></returns>
    Public Function GetCbg(FacilityCbg As String, FacilityLocation As String, FacilityName As String, SubFacilityName As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_FacilityCbgs.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@FacilityCbg", SqlDbType.NVarChar, 100).Value = FacilityCbg
                command.Parameters.Add("@FacilityLocation", SqlDbType.NVarChar, 100).Value = FacilityLocation
                command.Parameters.Add("@FacilityName", SqlDbType.NVarChar, 100).Value = FacilityName
                command.Parameters.Add("@SubFacilityName", SqlDbType.NVarChar, 100).Value = SubFacilityName

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)

                End Using

            End Using
            DataCenter.GlobalValues.message = String.Empty
            GetCbg = _tbAnswer
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
                    ErrorId = DataCenter.ErrorCenter.Facility
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetCbg = Nothing

        End Try

    End Function


    ''' <summary>
    ''' Return column/s:
    '''  <para/>  FacilityLocation
    ''' </summary>
    ''' <param name="FacilityCbg"></param>
    ''' <param name="FacilityLocation"></param>
    ''' <param name="FacilityName"></param>
    ''' <param name="SubFacilityName"></param>
    ''' <returns></returns>

    Public Function GetLocation(FacilityCbg As String, FacilityLocation As String, FacilityName As String, SubFacilityName As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_FacilityLocations.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@FacilityCbg", SqlDbType.NVarChar, 100).Value = FacilityCbg
                command.Parameters.Add("@FacilityLocation", SqlDbType.NVarChar, 100).Value = FacilityLocation
                command.Parameters.Add("@FacilityName", SqlDbType.NVarChar, 100).Value = FacilityName
                command.Parameters.Add("@SubFacilityName", SqlDbType.NVarChar, 100).Value = SubFacilityName

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)

                End Using

            End Using
            DataCenter.GlobalValues.message = String.Empty
            GetLocation = _tbAnswer

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
                    ErrorId = DataCenter.ErrorCenter.Facility
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetLocation = Nothing

        End Try

    End Function

    ''' <summary>
    ''' Return column/s:
    '''  <para/>  FacilityName
    ''' </summary>
    ''' <param name="FacilityCbg"></param>
    ''' <param name="FacilityLocation"></param>
    ''' <param name="FacilityName"></param>
    ''' <param name="SubFacilityName"></param>
    ''' <returns></returns>
    Public Function GetName(FacilityCbg As String, FacilityLocation As String, FacilityName As String, SubFacilityName As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_FacilityNames.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@FacilityCbg", SqlDbType.NVarChar, 100).Value = FacilityCbg
                command.Parameters.Add("@FacilityLocation", SqlDbType.NVarChar, 100).Value = FacilityLocation
                command.Parameters.Add("@FacilityName", SqlDbType.NVarChar, 100).Value = FacilityName
                command.Parameters.Add("@SubFacilityName", SqlDbType.NVarChar, 100).Value = SubFacilityName

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)

                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            GetName = _tbAnswer

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
                    ErrorId = DataCenter.ErrorCenter.Facility
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetName = Nothing

        End Try

    End Function


    ''' <summary>
    ''' Return column/s:
    '''  <para/>  SubFacilityName
    ''' </summary>
    ''' <param name="FacilityCbg"></param>
    ''' <param name="FacilityLocation"></param>
    ''' <param name="FacilityName"></param>
    ''' <param name="SubFacilityName"></param>
    ''' <returns></returns>

    Public Function GetSubName(FacilityCbg As String, FacilityLocation As String, FacilityName As String, SubFacilityName As String) As DataTable

        Try

            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_SubFacilityNames.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@FacilityCbg", SqlDbType.NVarChar, 100).Value = FacilityCbg
                command.Parameters.Add("@FacilityLocation", SqlDbType.NVarChar, 100).Value = FacilityLocation
                command.Parameters.Add("@FacilityName", SqlDbType.NVarChar, 100).Value = FacilityName
                command.Parameters.Add("@SubFacilityName", SqlDbType.NVarChar, 100).Value = SubFacilityName

                _tbAnswer = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    _tbAnswer = New DataTable()
                    dataAdapter.Fill(_tbAnswer)

                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            GetSubName = _tbAnswer

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
                    ErrorId = DataCenter.ErrorCenter.Facility
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            GetSubName = Nothing

        End Try

    End Function


End Class
