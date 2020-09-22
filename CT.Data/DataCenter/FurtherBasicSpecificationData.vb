Namespace DataCenter
    Public Class FurtherBasicSpecificationData
        Public pe02 As Long
        Public pe45 As Long
        Public FurtherBasicSpecificationList As String
        Public FurtherBasicSpecificationData As String
        Sub New(lPe02 As Long, lpe45 As Long, lFurtherBasicSpecificationList As String, lFurtherBasicSpecificationData As String)
            pe02 = lPe02
            pe45 = lpe45
            FurtherBasicSpecificationList = lFurtherBasicSpecificationList
            FurtherBasicSpecificationData = lFurtherBasicSpecificationData
        End Sub
    End Class
End Namespace
