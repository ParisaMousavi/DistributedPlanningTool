

Namespace DataCenter
    Public Class NonMfcSpecificationData

        Public pe02 As Long
        Public pe45 As Long
        Public NonMfcSpecification As String
        Public NonMfcSpecificationData As String

        Sub New(lpe02 As Long, lpe45 As Long, lNonMfcSpecification As String, lNonMfcSpecificationData As String)

            pe02 = lpe02
            pe45 = lpe45
            NonMfcSpecification = lNonMfcSpecification
            NonMfcSpecificationData = lNonMfcSpecificationData

        End Sub

    End Class
End Namespace
