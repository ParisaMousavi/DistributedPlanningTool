
Namespace DataCenter
    Public Class MfcSpecificationData

        Public pe02 As Long
        Public pe45 As Long
        Public Mfc As String
        Public Section As String
        Public Data As String


        Sub New(lpe02 As Long, lpe45 As Long, lMfc As String, lSection As String, lData As String)

            pe02 = lpe02
            pe45 = lpe45
            Mfc = lMfc
            Section = lSection
            Data = lData

        End Sub

    End Class
End Namespace
