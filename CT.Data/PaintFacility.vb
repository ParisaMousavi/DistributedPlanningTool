
Imports System.Data.SqlClient
Imports System.Data

Public Class PaintFacility
        Inherits CtBaseClass
        ''' <summary>
        ''' </summary>
        ''' <returns></returns>
        Public Function SelectAll() As DataTable

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.StandardCatalog_PaintFacilitySelectById.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@pe95_PaintFacilityStandards_PK", SqlDbType.Int, 4).Value = DBNull.Value

                Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
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
                        ErrorId = DataCenter.ErrorCenter.Region
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
                SelectAll = Nothing
            End Try

        End Function
    End Class

