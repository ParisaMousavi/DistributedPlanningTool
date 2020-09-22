
Namespace Form.AutoUpdate
    Public Class Update

        Public Function GetCurrentVersion() As String

            Try
                '--------------------------------------------------
                ' Go to registry and read current version
                '--------------------------------------------------



            Catch ex As Exception
                GetCurrentVersion = String.Empty
            End Try

        End Function


        Public Function GetAvailableVersion() As String
            Try

            Catch ex As Exception

                '---------------------------------------------
                ' Go to Sharepoint and read version file
                '---------------------------------------------

                GetAvailableVersion = String.Empty
            End Try
        End Function

    End Class
End Namespace
