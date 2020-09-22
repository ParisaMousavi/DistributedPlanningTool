Namespace DataCenter

    Public Class ProgramInformationData
        Public pe02 As Long
        Public pe45 As Long
        Public ProgramInformationList As String
        Public ProgramInformationData As String

        Sub New(lpe02 As Long, lpe45 As Long, lProgramInformationList As String, lProgramInformationData As String)
            pe02 = lpe02
            pe45 = lpe45
            ProgramInformationList = lProgramInformationList
            ProgramInformationData = lProgramInformationData
        End Sub
    End Class

    Public Enum ProgramInfoFields
        Cbg
        Dedicated
        TBNumber
        Vin
        Emissionstage
        Bodystyle
        ColorCode
        DriveSide
        ShippingAdress
        ShippingToCustomerDate
        TbNumberPrefix
        BuildId
        PaintFacility
        TagNumber
        RigCustomerPickDate
        CustomerRequiredDate
    End Enum

End Namespace