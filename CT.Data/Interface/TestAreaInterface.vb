
Namespace Interfaces
    Public Interface TestAreaInterface
        'Test Area
        Function GetTndAreaDataSpecific(HCID As Integer, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,)
        Function GetTndAreaDataGeneric(HCID As Integer, MainBuildType As String, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object) As String(,)


    End Interface
End Namespace
