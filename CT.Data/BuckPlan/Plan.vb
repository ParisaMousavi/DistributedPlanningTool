
Imports System.Data
Imports System.Data.SqlClient
Imports CT.Data.DataCenter
Imports CT.Data.Interfaces

Namespace BuckPlan

    Public Class Plan
        Inherits CtBaseClass
        Implements Interfaces.PlanInterface

        Public Enum SelectAllSpecificTndPlansColumns
            pe01_TnDBasicProgram_FK
            GenericSpecific
            HealthChartId
            HealthChartName
            PlanVersion
            FileStatus
            BuildPhase
            Quantity
            AssyMrd
            M1DC
            PEC
            FEC
            Platform
            Carline
            XCCpe01
            XCCpe26
            AssyBuildScale
            BuildType
            pe02
            pe01
            Region

        End Enum

        Public Function SelectAllSpecificTndPlans() As DataTable Implements Interfaces.PlanInterface.SelectAllSpecificTndPlans
            Dim _tbAnswer As DataTable = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.BuckPlan.ooo.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectAllSpecificTndPlans = _tbAnswer
            Catch ex As Exception
                '----------------------------------------------------------------
                ' Error classification mechanism
                '----------------------------------------------------------------
                Dim ErrorId As Integer
                Select Case ex.Message
                    Case ex.Message.IndexOf("Permission") >= 0
                        ErrorId = DataCenter.ErrorCenter.Permission
                    Case ex.Message.IndexOf("could not found") >= 0
                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                    Case Else
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectAllSpecificTndPlans = Nothing

            End Try

        End Function

        Public Function ValidatePlan(HealChartId As Integer, MainBuildType As String, FileStatus As String) As DataTable Implements Interfaces.PlanInterface.ValidatePlan
            Dim _tbAnswer As DataTable = Nothing

            Try
                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)
                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.BuckPlan.ooo.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@HealthChartId", SqlDbType.Int, 4).Value = HealChartId
                    command.Parameters.Add("@MainBuildType", SqlDbType.NVarChar, 10).Value = MainBuildType 'CT.Data.DataCenter.BuildType.Buck.ToString
                    command.Parameters.Add("@FileStatus", SqlDbType.NVarChar, 20).Value = FileStatus

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                ValidatePlan = _tbAnswer
            Catch ex As Exception
                '----------------------------------------------------------------
                ' Error classification mechanism
                '----------------------------------------------------------------
                Dim ErrorId As Integer
                Select Case ex.Message
                    Case ex.Message.IndexOf("Permission") >= 0
                        ErrorId = DataCenter.ErrorCenter.Permission
                    Case ex.Message.IndexOf("could not found") >= 0
                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                    Case Else
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                ValidatePlan = Nothing
            End Try


        End Function
        Public Function SelectAllGenericTndPlan() As DataTable Implements Interfaces.PlanInterface.SelectAllGenericTndPlan
            Dim _tbAnswer As DataTable = Nothing

            Try

                Using conTnd As SqlConnection = New SqlConnection(CT.Data.My.Settings.ConnectionString1)

                    Dim command As SqlCommand = New SqlCommand(DataCenter.StoredProcedures.BuckPlan.ooo.ToString())
                    command.Connection = conTnd
                    command.CommandType = CommandType.StoredProcedure

                    _tbAnswer = Nothing
                    Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(command)
                        _tbAnswer = New DataTable()
                        dataAdapter.Fill(_tbAnswer)
                    End Using

                End Using

                DataCenter.GlobalValues.message = String.Empty
                SelectAllGenericTndPlan = _tbAnswer
            Catch ex As Exception
                '----------------------------------------------------------------
                ' Error classification mechanism
                '----------------------------------------------------------------
                Dim ErrorId As Integer
                Select Case ex.Message
                    Case ex.Message.IndexOf("Permission") >= 0
                        ErrorId = DataCenter.ErrorCenter.Permission
                    Case ex.Message.IndexOf("could not found") >= 0
                        ErrorId = DataCenter.ErrorCenter.Could_Not_Find_Sp
                    Case Else
                        ErrorId = DataCenter.ErrorCenter.Plan
                End Select
                DataCenter.GlobalValues.message = String.Format("{0:d} :  {1}", ErrorId, ex.Message)
                SelectAllGenericTndPlan = Nothing
            End Try

        End Function



        Public Function SelectDateInformation(pe02 As Long) As DataTable Implements PlanInterface.SelectDateInformation
            Throw New NotImplementedException()
        End Function


        Public Function GenerateGenericPlan(ByRef pe01 As Long, ByRef pe02 As Long, HCID As Integer, xccpe26 As Long, xccpe01 As Long, AssyBuildScale As Integer, BuildPhase As String, BuildType As String, FileStatus As FileStatus) As Object Implements PlanInterface.GenerateGenericPlan
            Throw New NotImplementedException()
        End Function

        Public Function SelectTndDraftPlanDedicated(DraftHCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.SelectTndDraftPlanDedicated
            Throw New NotImplementedException()
        End Function

        Public Function GenerateDraftOrCheckout(ByRef HCID As Integer, FileStatus As FileStatus, MainBuildType As String) As Boolean Implements PlanInterface.GenerateDraftOrCheckout
            Throw New NotImplementedException()
        End Function

        Public Function GetQuantityTableXCC(HCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.GetQuantityTableXCC
            Throw New NotImplementedException()
        End Function

        Public Function GetQuantityTableCT(HCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.GetQuantityTableCT
            Throw New NotImplementedException()
        End Function

        Public Function ConvertGenericToSpecific(ByRef pe01 As Long, ByRef pe02 As Long, HCID As Integer, xccpe26 As Long, xccpe01 As Long, AssyBuildScale As Integer, BuildPhase As String, BuildType As String, FileStatus As FileStatus, WithCustomFormat As Boolean) As Boolean Implements PlanInterface.ConvertGenericToSpecific
            Throw New NotImplementedException()
        End Function

        Public Function SelectAllTndDraftPlans(MainBuildType As String, HCID As Integer) As DataTable Implements PlanInterface.SelectAllTndDraftPlans
            Throw New NotImplementedException()
        End Function

        Public Function DeleteDraftOrCheckedout(pe01 As Long, HCID As Integer, FileStatus As FileStatus, MainBuildType As String) As Boolean Implements PlanInterface.DeleteDraftOrCheckedout
            Throw New NotImplementedException()
        End Function

        Public Function GetCTEnginesAndTransmissions(HCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.GetCTEnginesAndTransmissions
            Throw New NotImplementedException()
        End Function

        Public Function GetXCCEnginesAndTransmissions(HCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.GetXCCEnginesAndTransmissions
            Throw New NotImplementedException()
        End Function

        Public Function GetPlanRemarks(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.GetPlanRemarks
            Throw New NotImplementedException()
        End Function

        Public Function GetAssignedCDSIDs(pe01 As Long, HCID As Integer, MainBuildType As String) As DataTable Implements PlanInterface.GetAssignedCDSIDs
            Throw New NotImplementedException()
        End Function

        Public Function ConvertCheckedouttToLife(pe01 As Long, LiveHCID As Integer, DraftOrCheckedoutHCID As Integer, FileStatus As FileStatus, MainBuildType As String) As Boolean Implements PlanInterface.ConvertCheckedouttToLife
            Throw New NotImplementedException()
        End Function

        Public Function ConvertDraftToLife(pe01 As Long, LiveHCID As Integer, DraftOrCheckedoutHCID As Integer, FileStatus As FileStatus, MainBuildType As String, ByRef ActivePlanHCID As Integer) As Boolean Implements PlanInterface.ConvertDraftToLife
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace