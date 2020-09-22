Namespace DataCenter.StoredProcedures
    Public Enum VehiclePlan

        A1_Header_Vehicle_Generic
        A1_Header_Vehicle_Specific

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.FurtherBasicSpecification Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_FurtherBasicPartial  ' VehiclePlan.Seventabs.FurtherBasic
        A2_VehicleAnd7Tabs_Vehicle_Specific_FurtherBasicPartial    ' VehiclePlan.Seventabs.FurtherBasic

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.Instrumentation Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_InstrumentationPartial
        A2_VehicleAnd7Tabs_Vehicle_Specific_InstrumentationPartial

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.MfcSpecification Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_MfcPartial
        A2_VehicleAnd7Tabs_Vehicle_Specific_MfcPartial

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.NonMfcSpecification Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_NonMfcPartial
        A2_VehicleAnd7Tabs_Vehicle_Specific_NonMfcPartial

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.ProgramInformation Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_ProgramInformationPartial
        A2_VehicleAnd7Tabs_Vehicle_Specific_ProgramInformationPartial

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.Updatepack Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_UpdatePackPartial
        A2_VehicleAnd7Tabs_Vehicle_Specific_UpdatePackPartial

        '---------------------------------------------------------
        ' VehiclePlan.SevenTabs.UserShippingDetails Class
        '---------------------------------------------------------
        A1_Header_Vehicle_Specific_UserShippingDetailsPartial
        A2_VehicleAnd7Tabs_Vehicle_Specific_UserShippingDetailsPartial

        A2_VehicleAnd7Tabs_Vehicle_Generic
        A2_VehicleAnd7Tabs_Vehicle_Specific

        A3_TimeLineData_Vehicle_Specific
        A3_TimeLineData_Vehicle_Generic

        ' New amendments to the SP's list included

        Report_VehicleReportI ' Total count report
        Report_VehicleReportII '
        Report_VehicleReportIII
        Report_VehicleReportIV
        Report_VehicleReportV

        Code_Vehicle_AllGenericTnDPlanListVers115
        Code_Vehicle_AllSpecificTnDPlanListVers115

        Code_Vehicle_AllTnDDraftPlanList
        Code_Vehicle_TnDDraftPlanDedicated

        Report_Vehicle_QuantityTableXCC ' For count table in top left corner.
        Report_Vehicle_QuantityTableCT ' For count table in top left corner.

        '----------------------------------------------------------
        ' For generating generic plan and converting generic to specific plan
        '----------------------------------------------------------
        LoadingVehicle1_Generic_AddXccSingleProgram 'Loading1_Generic_AddXccSingleProgram                    'Loading Generic plan       'Convert G->S      
        LoadingVehicle2_Generic_AddTnDPlanner 'Loading2_Generic_AddTnDPlanner                          'Loading Generic plan       'Convert G->S 
        LoadingVehicle3_Generic_GenericSplit 'Loading3_Generic_GenericSplit                           'Loading Generic plan       'Convert G->S 
        LoadingVehicle4_Generic_PowerPackAllocation 'Loading4_Generic_PowerPackAllocation                    'Loading Generic plan       'Convert G->S 
        LoadingVehicle5_Generic_UsercaseAllocationLogicII 'Loading5_Generic_UsercaseAllocationLogicII              'Loading Generic plan       'Convert G->S 
        LoadingVehicle6_Generic_InitialBuckLoading 'Loading6_Generic_InitialBuckLoading                     'Loading Generic Plan       'Convert G->S 
        LoadingVehicle7_Specific_AllocatedUsercasesGenericToSpecificII 'Loading7_Specific_AllocatedUsercasesGenericToSpecificII '---------------------      'Convert G->S 
        LoadingVehicle8_Generic_InitialPrototypeTeamnameAssaignment 'Loading8_Generic_InitialPrototypeTeamnameAssaignment    'Laoding Generic Plan       'Convert G->S 

        Specific_ProcessStepAddNewInsertII

        '----------------------------------------------------------
        ' frmPlanValidation
        '----------------------------------------------------------
        Generic_Check_Vehicle_1_InsertToPe01
        Generic_Check_Vehicle_2_InsertToPe02                               'frmPlanValidation
        Generic_Check_Vehicle_5_XCCEngineeTransmissionPrototypeuser        'frmPlanValidation
        Generic_Check_Vehicle_6_XCCRemovedEnginee                          'frmPlanValidation

        Specific_UpdateTnDProgramDetails

        A4_Vehicle_ColumnFormatSettings 'A4_ColumnFormatSettings 'PlanIndivitualFormatting

        Report_TnDPlanF4TValidation
        Report_TnDPlanF4TTransfer

        Validation_VehicleOverallValidation

    End Enum
End Namespace
