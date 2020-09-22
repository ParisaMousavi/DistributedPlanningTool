Imports System.Data.SqlClient

Namespace Interfaces


    Public Interface UnitInterface
        Function GetPreviousValueGeneral(pe02 As Long, pe03 As Long, strField As String, MainBuildType As String) As String
        Function GetPreviousValueVin(pe02 As Long, pe03 As Long, MainBuildType As String) As String
        Function GetPreviousValueVehicleNumber(pe02 As Long, pe03 As Long, MainBuildType As String) As String

        Function GetPreviousValueShippingToCustomer(pe02 As Long, pe03 As Long, MainBuildType As String) As String

        Function ChangeBuildSequence(pe02 As Long, CurrentDisplaySeq As Integer, FutureDisplaySeq As Integer, MainBuildType As String) As DataTable

        Function AddUnit(HCID As Integer, pe01 As Long, pe02 As Long, BuildPhase As String, MainBuildType As String, HardwareBuildType As String, HealthChartId As Integer, ByRef pe03_ID As Long, ByRef pe45_ID As Long, ByRef GenericSplitRowNumber As Integer) As Boolean

        Function ChangeEngine(pe02 As Long, pe45 As Long, NewEngineName As String, MainBuildType As String) As Boolean

        Function ChangeTransmission(pe02 As Long, pe45 As Long, NewTransmissionName As String, MainBuildType As String) As Boolean

        Function ChangeInfoII(Pe02 As Long, pe45 As Long, Pe03 As Long, MainBuildType As String, FileStatus As String, HealthChartId As Long, Optional StrCBG As Object = Nothing,
                                     Optional StrXccTeamName As Object = Nothing,
                                     Optional StrDedicated As Object = Nothing,
                                     Optional StrTBNumber As Object = Nothing,
                                     Optional StrVin As Object = Nothing,
                                     Optional StrEmissionStage As Object = Nothing,
                                     Optional StrBodySyle As Object = Nothing,
                                     Optional StrColorCode As Object = Nothing,
                                     Optional StrDriveSide As Object = Nothing,
                                     Optional StrTeamName As Object = Nothing,
                                     Optional StrRemarks As Object = Nothing,
                                     Optional StrShippingToCustomerDate As Object = Nothing,
                                     Optional strCustomerRequiredDate As Object = Nothing,
                                     Optional strRigCustomerPickDate As Object = Nothing,
                                     Optional StrTbNumberPrefix As Object = Nothing,
                                     Optional StrBuildId As Object = Nothing,
                                     Optional StrTagNumber As Object = Nothing,
                                     Optional StrPaintFacility As Object = Nothing,
                                     Optional CustomerRequiredDate As Object = Nothing,
                                     Optional RigCustomerPickDate As Object = Nothing) As Boolean


        Function GetVehiclesUsercasesDedicated(pe45 As Long, MainBuildType As String, Optional transaction As SqlTransaction = Nothing, Optional conTnd As SqlConnection = Nothing) As DataTable

        Function Delete(HCID As Integer, pe03 As Long, pe02 As Long, pe45 As Long, MainBuildType As String) As Boolean


    End Interface

End Namespace
