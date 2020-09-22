Namespace Interfaces
    Public Interface MfcInterface
        Function GetPlanData(pe02 As Long, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,)

        Function GetTndPlanHeader(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,)

    End Interface
End Namespace