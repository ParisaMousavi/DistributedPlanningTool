Namespace Interfaces

    Public Interface PlanInterface



        Function SelectDateInformation(pe02 As Long) As DataTable

        Function GenerateGenericPlan(ByRef pe01 As Long, ByRef pe02 As Long, HCID As Integer, xccpe26 As Long, xccpe01 As Long, AssyBuildScale As Integer, BuildPhase As String, BuildType As String, FileStatus As DataCenter.FileStatus)

        Function SelectTndDraftPlanDedicated(DraftHCID As Integer, MainBuildType As String) As DataTable

        Function GenerateDraftOrCheckout(ByRef HCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String) As Boolean

        Function SelectAllSpecificTndPlans() As DataTable

        Function SelectAllGenericTndPlan() As DataTable

        Function GetQuantityTableXCC(HCID As Integer, MainBuildType As String) As DataTable

        Function GetQuantityTableCT(HCID As Integer, MainBuildType As String) As DataTable

        Function ConvertGenericToSpecific(ByRef pe01 As Long, ByRef pe02 As Long, HCID As Integer, xccpe26 As Long, xccpe01 As Long, AssyBuildScale As Integer, BuildPhase As String, BuildType As String, FileStatus As DataCenter.FileStatus, WithCustomFormat As Boolean) As Boolean

        Function SelectAllTndDraftPlans(MainBuildType As String, HCID As Integer) As DataTable

        Function ConvertDraftToLife(pe01 As Long, LiveHCID As Integer, DraftOrCheckedoutHCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String, ByRef ActivePlanHCID As Integer) As Boolean

        Function ConvertCheckedouttToLife(pe01 As Long, LiveHCID As Integer, DraftOrCheckedoutHCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String) As Boolean

        Function DeleteDraftOrCheckedout(pe01 As Long, HCID As Integer, FileStatus As DataCenter.FileStatus, MainBuildType As String) As Boolean

        Function GetCTEnginesAndTransmissions(HCID As Integer, MainBuildType As String) As DataTable

        Function GetXCCEnginesAndTransmissions(HCID As Integer, MainBuildType As String) As DataTable

        Function GetPlanRemarks(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable

        Function GetAssignedCDSIDs(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable

        Function ValidatePlan(HealChartId As Integer, MainBuildType As String, FileStatus As String) As DataTable


    End Interface

End Namespace
