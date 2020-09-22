
Namespace DataCenter
    Public Class UserShippingDetailsData


        Public pe02 As Long
        Public pe45 As Long
        Public UserShippingDetailsList As String
        Public UserShippingDetailsData As String

        Sub New(lpe02 As Long, lpe45 As Long, lUserShippingDetailsList As String, lUserShippingDetailsData As String)

            pe02 = lpe02
            pe45 = lpe45
            UserShippingDetailsList = lUserShippingDetailsList
            UserShippingDetailsData = lUserShippingDetailsData

        End Sub

    End Class
End Namespace