
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Error code : 50
''' </summary>
Public Class Phonebook
    Inherits CtBaseClass


    'Private _tbAnswer As DataTable = Nothing

    ''' <summary>
    ''' This function is the list of all CDSIDs in phone book and 
    ''' is displayed in CDSID2DVPTeam in Grid and user sees only this list.
    ''' The output table has the following columns: pe90_Phonebook_PK, CDSID, Fullname, Tel, pe27_Regions_Fk, Regions
    ''' </summary>
    ''' <returns></returns>
    Public Function SelectAll() As DataTable

        Try
            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PhonebookSelectAll.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                _tbAnswer = New DataTable
                SelectAll = Nothing
                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                    dataAdapter.Fill(_tbAnswer)
                End Using

            End Using

            DataCenter.GlobalValues.message = String.Empty
            SelectAll = _tbAnswer

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
                    ErrorId = DataCenter.ErrorCenter.Phonebook
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectAll = Nothing
        End Try


    End Function


    Public Function AddNew(CDSID As String, FullName As String, Tel As String, pe27 As Integer) As Boolean

        Try

            If CDSID.Length > 16 Or CDSID.Length = 0 Then Throw New Exception("The value of CDSID is not valid.")


            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                conTnd.Open()

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PhonebookNewEntry.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@CDSID", SqlDbType.NVarChar, 12).Value = CDSID
                command.Parameters.Add("@Fullname", SqlDbType.NVarChar, 100).Value = If(FullName.Length = 0, DBNull.Value, FullName)
                command.Parameters.Add("@Tel", SqlDbType.NVarChar, 50).Value = If(Tel.Length = 0, DBNull.Value, Tel)
                command.Parameters.Add("@pe27_Regions_FK", SqlDbType.Int, 4).Value = If(pe27 = 0, DBNull.Value, pe27)

                command.ExecuteNonQuery()

            End Using


            DataCenter.GlobalValues.message = String.Empty
            AddNew = True
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
                    ErrorId = DataCenter.ErrorCenter.Phonebook
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            AddNew = False
        End Try

    End Function


    Public Function Update(pe90 As Integer, CDSID As String, FullName As String, Tel As String, pe27 As Integer) As Boolean

        Try
            If CDSID.Length > 16 Or CDSID.Length = 0 Then Throw New Exception("The value of CDSID is not valid.")


            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                conTnd.Open()

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Specific_PhonebookEditEntry.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe90", SqlDbType.Int, 4).Value = pe90
                command.Parameters.Add("@CDSID", SqlDbType.NVarChar, 16).Value = CDSID
                command.Parameters.Add("@Fullname", SqlDbType.NVarChar, 100).Value = If(FullName.Length = 0, DBNull.Value, FullName)
                command.Parameters.Add("@Tel", SqlDbType.NVarChar, 50).Value = If(Tel.Length = 0, DBNull.Value, Tel)
                command.Parameters.Add("@pe27_Regions_FK", SqlDbType.Int, 4).Value = If(pe27 = 0, DBNull.Value, pe27)

                command.ExecuteNonQuery()

            End Using



            DataCenter.GlobalValues.message = String.Empty
            Update = True
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
                    ErrorId = DataCenter.ErrorCenter.Phonebook
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            Update = False
        End Try


    End Function

End Class
