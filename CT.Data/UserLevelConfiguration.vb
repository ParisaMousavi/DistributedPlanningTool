
Imports System.Data.SqlClient


Public Class UserLevelConfiguration

    Public Property CT_ConnectionString As String

        Get
            Try
                CT_ConnectionString = CT.Data.My.Settings.ConnectionString1
            Catch
                CT_ConnectionString = String.Empty
            End Try

        End Get
        Set(value As String)
            Dim conTnd As SqlConnection = Nothing

            'Try
            '    conTnd = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
            '    conTnd.Open()
            'Catch ex0 As Exception
            Try
                conTnd = New SqlConnection(CT.Data.My.Resources.ConnectionString1)
                conTnd.Open()
                CT.Data.My.Settings.ConnectionString1 = CT.Data.My.Resources.ConnectionString1
                CT.Data.My.Settings.Save()
            Catch ex1 As Exception
                Try
                    conTnd = New SqlConnection(CT.Data.My.Resources.ConnectionString2)
                    conTnd.Open()
                    CT.Data.My.Settings.ConnectionString1 = CT.Data.My.Resources.ConnectionString2
                    CT.Data.My.Settings.Save()
                Catch ex2 As Exception
                    CT.Data.My.Settings.ConnectionString1 = String.Empty
                    CT.Data.My.Settings.Save()
                End Try

            Finally
                If conTnd IsNot Nothing Then
                    conTnd.Close()
                    conTnd.Dispose()
                End If

            End Try
            'Finally
            '    If conTnd IsNot Nothing Then
            '        conTnd.Close()
            '        conTnd.Dispose()
            '    End If
            'End Try
        End Set

    End Property

End Class
