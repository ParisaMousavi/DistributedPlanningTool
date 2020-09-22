Namespace DataCenter.StoredProcedures
    Public Enum RigPlan
        A1_Header_Rig_Generic

        A4_Rig_ColumnFormatSettings


        Code_Rig_AllGenericTnDPlanListVers115
        Code_Rig_AllSpecificTnDPlanListVers115

        Code_Rig_AllTnDDraftPlanList

        Code_Rig_TnDDraftPlanDedicated

        Report_Rig_QuantityTableXCC ' For count table in top left corner.
        Report_Rig_QuantityTableCT ' For count table in top left corner.


        Validation_RigOverallValidation

        Generic_Check_Rig_1_InsertToPe01
        Generic_Check_Rig_2_InsertToPe02

        ' New amendments to the SP's list included

        A1_Header_Rig_Specific
        A1_Header_Rig_Specific_FurtherBasicPartial
        A1_Header_Rig_Specific_InstrumentationPartial
        A1_Header_Rig_Specific_MfcPartial
        A1_Header_Rig_Specific_NonMfcPartial
        A1_Header_Rig_Specific_ProgramInformationPartial
        A1_Header_Rig_Specific_UpdatePackPartial
        A1_Header_Rig_Specific_UserShippingDetailsPartial

        A2_VehicleAnd7Tabs_Rig_Generic
        A2_VehicleAnd7Tabs_Rig_Specific
        A2_VehicleAnd7Tabs_Rig_Specific_FurtherBasicPartial
        A2_VehicleAnd7Tabs_Rig_Specific_InstrumentationPartial
        A2_VehicleAnd7Tabs_Rig_Specific_MfcPartial
        A2_VehicleAnd7Tabs_Rig_Specific_NonMfcPartial
        A2_VehicleAnd7Tabs_Rig_Specific_ProgramInformationPartial
        A2_VehicleAnd7Tabs_Rig_Specific_UpdatePackPartial
        A2_VehicleAnd7Tabs_Rig_Specific_UserShippingDetailsPartial

        A3_TimeLineData_Rig_Specific
        A3_TimeLineData_Rig_Generic

        '----------------------------------------------------------
        ' For generating generic plan and converting generic to specific plan
        '----------------------------------------------------------
        LoadingRig1_Generic_AddXccSingleProgram
        LoadingRig2_Generic_AddTnDPlanner
        LoadingRig3_Generic_GenericSplit
        LoadingRig4_Generic_PowerPackAllocation
        LoadingRig5_Generic_UsercaseAllocationLogicII
        LoadingRig7_Specific_AllocatedUsercasesGenericToSpecificII
        LoadingRig8_Generic_InitialPrototypeTeamnameAssaignment


        '----------------------------------------------------------
        ' For changeInfo only for Rig
        '----------------------------------------------------------
        Specific_UnitCustomerRequiredDateUpdate
        Specific_UnitRigCustomerPickDateUpdate

    End Enum
End Namespace
