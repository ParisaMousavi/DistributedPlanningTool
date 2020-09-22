Namespace DataCenter.StoredProcedures
    Public Enum General
        OOO
        '-----------------------------------------------------------
        ' Specifications functionality
        '-----------------------------------------------------------
        Specific_DynamicHeaders                             'frmAddColumn
        Specific_DynamicFurtherBasicSpecificationAddColumn
        Specific_DynamicFurtherBasicSpecificationDelete
        Specific_DynamicFurtherBasicSpecificationEditColumn
        Specific_DynamicFurtherBasicSpecificationUpdateData



        Specific_DynamicInstrumentationAddColumn
        Specific_DynamicInstrumentationDelete
        Specific_DynamicInstrumentationEditColumn
        Specific_DynamicInstrumentationUpdateData



        Specific_DynamicMfcSpecificationAddColumn
        Specific_DynamicMfcSpecificationDelete
        Specific_DynamicMfcSpecificationEditColumn
        Specific_DynamicMfcSpecificationUpdateData



        Specific_DynamicNonMfcSpecificationAddColumn
        Specific_DynamicNonMfcSpecificationDelete
        Specific_DynamicNonMfcSpecificationEditColumn
        Specific_DynamicNonMfcSpecificationUpdateData



        Specific_DynamicProgramInformationAddColumn
        Specific_DynamicProgramInformationDelete
        Specific_DynamicProgramInformationEditColumn
        Specific_DynamicProgramInformationUpdateData



        Specific_DynamicUpdatepackAddColumn
        Specific_DynamicUpdatepackDelete
        Specific_DynamicUpdatepackEditColumn
        Specific_DynamicUpdatepackUpdateData



        Specific_DynamicUserShippingDetailsAddColumn
        Specific_DynamicUserShippingDetailsDelete
        Specific_DynamicUserShippingDetailsEditColumn
        Specific_DynamicUserShippingDetailsUpdateData

        '--------------------------------------------------------------
        ' frmHeaderEdit
        '--------------------------------------------------------------
        Permission_TnDAuthorizationSelectByHCID ' frmHeaderEdit
        Permission_TnDAuthorizationAdd ' frmHeaderEdit
        Permission_TnDAuthorizationDelete ' frmHeaderEdit
        Permission_TnDAuthorizationUpdate ' frmHeaderEdit
        Permission_TnDCurrentUserPermission ' frmHeaderEdit
        Permission_TnDCurrentUserPermissionGeneric

        '--------------------------------------------------------------
        ' ChangeLog class in DAL
        '--------------------------------------------------------------
        Undo_AddChangeLog '--->
        Undo_GetTnDLastUndo
        Undo_IsRedoCommandAvailable
        Undo_IsUndoCommandAvailable
        Undo_PreviousOperation
        Undo_UndoPreviousOperation


        '--------------------------------------------------------------
        ' DataLog class in DAL
        '--------------------------------------------------------------
        Specific_ChangeLogentrySelectByPlan
        Specific_ChangeLogentryDelete
        Specific_ChangeLogentryAdd
        Specific_ChangeLogentryUpdate

        '--------------------------------------------------------------
        ' Engine class in DAL
        '--------------------------------------------------------------
        Specific_ListXCCEngines
        'Specific_ListCTEngines

        '--------------------------------------------------------------
        ' Phonebook class in DAL
        '--------------------------------------------------------------
        Specific_PhonebookSelectAll
        Specific_PhonebookNewEntry
        Specific_PhonebookEditEntry

        '--------------------------------------------------------------
        ' IndividualFormatting class in DAL
        '--------------------------------------------------------------
        Specific_FormatInitialGenerate 'PlanIndivitualFormatting
        Specific_FormatUpdateSettings

        '--------------------------------------------------------------
        ' ProcessStep class in DAL
        '--------------------------------------------------------------
        Specific_ProcessStepCallEditData
        Specific_ProcessStepCallEditDate
        Specific_ProcessStepAddNewInsert
        Specific_ProcessStepAddNewInsertSequence
        Specific_ProcessStepDelete
        Report_ProcessStepDedicated '-> Used for frmEdit Process step
        Code_ProcessStepDedicated '-> Used for Move right

        'DvpTeamName & CDSID
        Report_ProgramDvpTeamNameCdsid
        Specific_AssignCdsid2DvpTeamName
        Report_CdsidDvpTeamDedicated



        '--------------------------------------------------------------
        ' ProgramConfiguration class in DAL
        ' frmHeaderEdit
        '--------------------------------------------------------------
        Specific_ProgramConfigAdd
        Specific_ProgramConfigDelete ' frmHeaderEdit
        Specific_ProgramConfigUpdate ' frmHeaderEdit
        Specific_ProgramConfigSelectByPlan ' frmHeaderEdit


        '--------------------------------------------------------------
        ' Plan class in DAL
        '--------------------------------------------------------------
        Report_XCCUserTeamNameTranslation
        Report_SelectPlanDedicated


        '--------------------------------------------------------------
        ' PublicHoliday class in DAL
        '--------------------------------------------------------------
        Standards_GetGenericPublicHolidays
        Specific_ProcessStepPopulateHolidays
        Specific_PublicHolidayAdd
        Specific_PublicHolidayDelete
        Specific_PublicHolidayUpdate
        Specific_GetPlanPublicHolidays
        Report_PlanHolidays
        Standards_GetPublicHolidayLocations

        '--------------------------------------------------------------
        ' PublicHolidayType class in DAL
        '--------------------------------------------------------------
        Standards_GetPublicHolidayTypes


        '--------------------------------------------------------------
        ' For frmPhonebook & frmHeaderEdit
        '--------------------------------------------------------------
        StandardCatalog_RegionsSelectById

        '--------------------------------------------------------------
        ' SecurityLevel class in DAL
        '--------------------------------------------------------------
        Report_SecurityLevels


        '--------------------------------------------------------------
        ' Transmission class in DAL
        '--------------------------------------------------------------
        Specific_ListXCCTransmissions


        StandardCatalog_PaintFacilitySelectById

        Specific_ColumnInputFormat
        'Specific_DetectOverlapping
        Specific_GetUnitProgramInformation


        '--------------------------------------------------------------
        ' More usage
        '--------------------------------------------------------------
        Generic_DeleteXCCProgram                                'Loading Generic plan       'Convert G->S 


        '-----------------------------------------------------------
        ' Unit functionality
        '-----------------------------------------------------------
        Specific_AddUnit
        Specific_ChangeVehicleDisplaySeq
        Specific_DeactivateVehicle



        '--------------------------------------------------------------
        ' Generate Draft
        ' Delete Draft
        ' Replace Draft
        '
        ' Checkout
        ' Checkin
        ' Discard
        '--------------------------------------------------------------
        DraftGeneration_A_pe02
        DraftGeneration_B_pe03
        DraftGeneration_C_pe45
        DraftGeneration_D_pe57
        DraftGeneration_E_pe34
        DraftGeneration_F_pe22
        DraftGeneration_G_pe44
        DraftGeneration_H_pe26
        DraftGeneration_I_pe77
        DraftGeneration_J_pe75
        DraftGeneration_K_pe73
        DraftGeneration_L_pe71
        DraftGeneration_M_pe69
        DraftGeneration_N_pe66
        DraftGeneration_O_pe85
        DraftGeneration_P_pe78
        DraftGeneration_Q_pe87
        DraftGeneration_R_pe67
        DraftGeneration_S_pe88
        DraftGeneration_T_pe62
        DraftGeneration_U_pe04
        Draft_SwitchDraftToMaster
        Draft_SwitchCheckedoutToMaster
        DraftGeneration_Discard




        Report_ListDateInformation 'Used in frmAddDates & Vehicle/Rig/Buck Plan.vb


        ' developed in AddtionalDateInformation class in DAL
        Specific_AddtionalDateInformationSelectByPlan 'Used in frmAddDates & Vehicle/Rig/Buck Plan.vb
        Specific_AddtionalDateInformationAdd
        Specific_AddtionalDateInformationDelete
        Specific_AddtionalDateInformationUpdate

        Report_AssaignedCTEnginesAndTransmissions
        Report_AssaignedXCCEnginesAndTransmissions

        'Used in frmEdit, frmEditUserCase both Vehicle&Rig
        Report_FacilityNames
        Report_FacilityLocations
        Report_FacilityCbgs
        Report_SubFacilityNames


        'used in PlanActiveUser class
        Specific_PlanActiveUserInsert
        Specific_PlanActiveUserRemove
        Specific_PlanActiveUserSelectAll  ' frmHCIDSelect


        'Used in UnitReport - ContextMenu Delete Processstep
        Report_VehiclesUsercasesDisplayDedicated

        'Used in EditUsercase
        Report_UsercaseDedicated
        Specific_UsercaseEditCDSID
        Specific_UsercaseEditRemarks

        'Used in frmNew
        Report_UsercasesInformation

        'Used in Export to Excel report
        Report_PlanRemarks
        Report_PlanAssignedCDSIDs

        Specific_ChangeVehicleTransmission ' this for combo boxes
        Specific_ChangeVehicleEngine ' this for combo boxes

        'Program Info update
        Specific_UnitColorCodeUpdate
        Specific_UnitBodySyleUpdate
        Specific_UnitDriveSideUpdate
        Specific_UnitDedicatedUpdate
        Specific_UnitEmissionStageUpdate
        Specific_UnitRemarksUpdate
        Specific_UnitShippingToCustomerDateUpdate
        Specific_UnitCBGUpdate
        Specific_UnitBuildIdUpdate
        Specific_UnitPaintFacilityUpdate
        Specific_UnitTagNumberUpdate
        Specific_UnitTbNumberPrefixUpdate
        Specific_UnitTBNumberUpdate
        Specific_UnitVINUpdate
        Specific_UnitXCCTeaMAndTeamNameUpdate
        Specific_UnitCustomerRequiredDateUpdate
        Specific_UnitRigCustomerPickDateUpdate
        'Not used yet
        Report_DynamicBuckAnd7Tabs
        Report_DynamicRebuildAnd7Tabs
        Report_DynamicRigAnd7Tabs
        Report_DynamicVehicleAnd7Tabs

        Specific_ListCTTransmissions



        '-----------------------------------------------------------
        ' Message Passing class
        '-----------------------------------------------------------
        Specific_MessagePassingSelectAll
        Specific_MessagePassingInsert
        Specific_MessagePassingDeactive

    End Enum
End Namespace
