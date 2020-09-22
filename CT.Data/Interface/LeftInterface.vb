Namespace Interfaces
    Public Interface LeftInterface

        'Left side
        Function GetPlanDataHcIdSpecific(HcId As Integer, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object, MainBuildType As String) As String(,)
        Function GetPlanDataHcIdGeneric(HcId As Integer, MainBuildType As String, UpperBoundDisplaySeq As Object, LowerBoundDisplaySeq As Object) As String(,)

    End Interface

End Namespace