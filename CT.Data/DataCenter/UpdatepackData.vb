Namespace DataCenter
    Public Class UpdatepackData

        Public pe02 As Long
        Public pe45 As Long
        Public UpdatePackList As String
        Public UpdatepackData As String

        Sub New(lpe02 As Long, lpe45 As Long, lUpdatePackList As String, lUpdatepackData As String)

            pe02 = lpe02
            pe45 = lpe45
            UpdatePackList = lUpdatePackList
            UpdatepackData = lUpdatepackData

        End Sub

    End Class
End Namespace