Namespace Interfaces
    Public Interface HeaderInterface
        'Header
        Function GetPlanHeaderSpecific(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,)
        Function GetPlanHeaderGeneric(HCID As Integer, BuildType As String, BuildPhase As String, MainBuildType As String) As String(,)

    End Interface
End Namespace