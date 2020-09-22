
Namespace DataCenter
    Public Class InstrumentationData


        Public pe02 As Long
        Public pe45 As Long
        Public Section As String
        Public InstrumentationList As String
        Public InstrumentationData As String

        Sub New(lPe02 As Long, lpe45 As Long, lSection As String, lInstrumentationList As String, lInstrumentationData As String)

            pe02 = lPe02
            pe45 = lpe45
            Section = lSection
            InstrumentationList = lInstrumentationList
            InstrumentationData = lInstrumentationData

        End Sub

    End Class
End Namespace
