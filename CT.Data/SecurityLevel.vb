Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' The error classification 290
''' </summary>
Public Class SecurityLevel
    Inherits CtBaseClass



    Public Function SelectAll() As DataTable

        Try


            Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.General.Report_SecurityLevels.ToString())
                command.Connection = conTnd
                command.CommandType = CommandType.StoredProcedure

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
                    ErrorId = DataCenter.ErrorCenter.SecurityLevel
            End Select
            DataCenter.GlobalValues.message = String.Format("{0:d}:  {1}", ErrorId, ex.Message)
            SelectAll = Nothing
        End Try




    End Function



End Class
