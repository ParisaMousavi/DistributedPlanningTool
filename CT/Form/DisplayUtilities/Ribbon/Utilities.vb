
Namespace Form.DisplayUtilities.Ribbon
    Public Class Utilities


        Public Sub UpdateRibbonButtonsState()
            Try

                Try
                    If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                        Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Visible = False
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Visible = False
                        Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Visible = False
                        'Globals.Ribbons.RbnTnDControlPanel.SepCountReport.Visible = False
                        'Globals.Ribbons.RbnTnDControlPanel.SepPushToFit4Test.Visible = False
                        'Globals.Ribbons.RbnTnDControlPanel.SepPrecheckReport.Visible = False
                    Else
                        Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Visible = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Visible = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Visible = True
                        'Globals.Ribbons.RbnTnDControlPanel.SepCountReport.Visible = True
                        'Globals.Ribbons.RbnTnDControlPanel.SepPushToFit4Test.Visible = True
                        'Globals.Ribbons.RbnTnDControlPanel.SepPrecheckReport.Visible = True
                    End If
                Catch ex As Exception
                End Try


                Dim objPer As New CT.Data.Authorization
                Dim _strUserPermissionLevel As String = String.Empty
                Try
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel = Nothing Then
                        '--------------------------------------------------------------------------
                        ' validation for controlling the result of DAL
                        '--------------------------------------------------------------------------
                        _strUserPermissionLevel = objPer.GetPermissionLevel(Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.IsGeneric)
                        If _strUserPermissionLevel Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        Form.DataCenter.GlobalValues.strUserPermissionLevel = _strUserPermissionLevel
                    End If
                Catch ex As Exception
                    Form.DataCenter.GlobalValues.strUserPermissionLevel = String.Empty
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Update Ribbon Buttons State", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                End Try
                If Form.DataCenter.ProgramConfig.HCID <> 0 And Form.DataCenter.ProgramConfig.IsGeneric = True Then
                    If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then

                        Select Case Form.DataCenter.GlobalValues.strUserPermissionLevel
                            Case CT.Data.DataCenter.UserPermissionLevel.Owner.ToString
                                Generic_Master_OwnerAndExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Executor.ToString
                                Generic_Master_OwnerAndExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString
                                Generic_Master_Visitor()
                            Case Else
                                DeactiveRibbonButtonsState()
                        End Select

                    Else
                        DeactiveRibbonButtonsState()
                    End If
                ElseIf Form.DataCenter.ProgramConfig.HCID <> 0 And Form.DataCenter.ProgramConfig.IsGeneric = False Then


                    If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then

                        Select Case Form.DataCenter.GlobalValues.strUserPermissionLevel
                            Case CT.Data.DataCenter.UserPermissionLevel.Owner.ToString
                                Specific_Master_OwnerAndExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Executor.ToString
                                Specific_Master_OwnerAndExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString
                                Specific_Master_Visitor()
                            Case Else
                                DeactiveRibbonButtonsState()
                        End Select

                    ElseIf Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString Then

                        Select Case Form.DataCenter.GlobalValues.strUserPermissionLevel
                            Case CT.Data.DataCenter.UserPermissionLevel.Owner.ToString
                                Specific_Checkedout_OwnerAndExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Executor.ToString
                                Specific_Checkedout_OwnerAndExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString
                                Specific_Checkedout_Visitor()
                            Case Else
                                DeactiveRibbonButtonsState()
                        End Select


                    ElseIf Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Draft.ToString Then


                        Select Case Form.DataCenter.GlobalValues.strUserPermissionLevel
                            Case CT.Data.DataCenter.UserPermissionLevel.Owner.ToString
                                Specific_Draft_OwnerOrExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Executor.ToString
                                Specific_Draft_OwnerOrExecutor()
                            Case CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString
                                DeactiveRibbonButtonsState()
                            Case Else
                                DeactiveRibbonButtonsState()
                        End Select

                    Else
                        DeactiveRibbonButtonsState()
                    End If


                Else
                    DeactiveRibbonButtonsState()
                End If


            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub


        Private Sub Generic_Master_OwnerAndExecutor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = True
                    .menuCheckInOut.Enabled = False
                    .menuDraft.Enabled = False
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = False
                    .btnUndo.Enabled = False
                    .btnRedo.Enabled = False
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = False
                    .btnDeleteUnit.Enabled = False
                    .btnChangeSequence.Enabled = False
                    .btnUpdateMRD.Enabled = False
                    .togFurtherBasicSpecification.Enabled = False
                    .togInstrumentation.Enabled = False
                    .togMfcSpecification.Enabled = False
                    .togNonMFCSpecification.Enabled = False
                    .togProgramInformation.Enabled = False
                    .togShowAll.Enabled = False
                    .togTiming.Enabled = False
                    .togUpdatePack.Enabled = False
                    .togUserShipping.Enabled = False
                    .btnUpdateColumns.Enabled = False
                    .btnUpdateHoliday.Enabled = False
                    .btnCDSIDtoDvpTeam.Enabled = False
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = False
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = False
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = False



                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub Generic_Master_Visitor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = False
                    .menuCheckInOut.Enabled = False
                    .menuDraft.Enabled = False
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = False
                    .btnUndo.Enabled = False
                    .btnRedo.Enabled = False
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = False
                    .btnDeleteUnit.Enabled = False
                    .btnChangeSequence.Enabled = False
                    .btnUpdateMRD.Enabled = False
                    .togFurtherBasicSpecification.Enabled = False
                    .togInstrumentation.Enabled = False
                    .togMfcSpecification.Enabled = False
                    .togNonMFCSpecification.Enabled = False
                    .togProgramInformation.Enabled = False
                    .togShowAll.Enabled = False
                    .togTiming.Enabled = False
                    .togUpdatePack.Enabled = False
                    .togUserShipping.Enabled = False
                    .btnUpdateColumns.Enabled = False
                    .btnUpdateHoliday.Enabled = False
                    .btnCDSIDtoDvpTeam.Enabled = False
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = False
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = False
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = False
                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub

        Private Sub Specific_Master_OwnerAndExecutor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = False
                    .menuDraft.Enabled = True
                    .menuCheckInOut.Enabled = True
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = False
                    .btnUndo.Enabled = False
                    .btnRedo.Enabled = False
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = False
                    .btnDeleteUnit.Enabled = False
                    .btnChangeSequence.Enabled = False
                    .btnUpdateMRD.Enabled = False
                    .togFurtherBasicSpecification.Enabled = True
                    .togInstrumentation.Enabled = True
                    .togMfcSpecification.Enabled = True
                    .togNonMFCSpecification.Enabled = True
                    .togProgramInformation.Enabled = True
                    .togShowAll.Enabled = True
                    .togTiming.Enabled = True
                    .togUpdatePack.Enabled = True
                    .togUserShipping.Enabled = True
                    .btnUpdateColumns.Enabled = False
                    .btnUpdateHoliday.Enabled = False
                    .btnCDSIDtoDvpTeam.Enabled = False
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = True
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = True
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = True

                    '------------------------------
                    ' Check-In/-Out logic
                    '------------------------------
                    .btnCheckOut.Visible = True
                    .btnCheckOut.Enabled = True

                    .btnCheckIn.Enabled = False
                    .btnCheckIn.Visible = False

                    .btnDiscard.Enabled = False
                    .btnDiscard.Visible = False
                    '-----------------------------


                    '------------------------------
                    ' Draft logic
                    '------------------------------
                    .btnGenerateDraft.Visible = True
                    .btnGenerateDraft.Enabled = True
                    .mnuGenerateDraft.Visible = True
                    .mnuGenerateDraft.Enabled = True

                    .mnuActiveUsers.Visible = False
                    .mnuActiveUsers.Enabled = False

                    .btnDeleteDraft.Visible = False
                    .btnReplacePlanWithDraft.Visible = False
                    '------------------------------


                    If Form.DataCenter.GlobalSections.TimeLineSection.Columns.Hidden = False Then
                        .togTiming.Checked = True
                    Else
                        .togTiming.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False Then
                        .togUpdatePack.Checked = True
                    Else
                        .togUpdatePack.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False Then
                        .togUserShipping.Checked = True
                    Else
                        .togUserShipping.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False Then
                        .togFurtherBasicSpecification.Checked = True
                    Else
                        .togFurtherBasicSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False Then
                        .togProgramInformation.Checked = True
                    Else
                        .togProgramInformation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False Then
                        .togMfcSpecification.Checked = True
                    Else
                        .togMfcSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False Then
                        .togInstrumentation.Checked = True
                    Else
                        .togInstrumentation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False Then
                        .togNonMFCSpecification.Checked = True
                    Else
                        .togNonMFCSpecification.Checked = False
                    End If

                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub

        Private Sub Specific_Master_Visitor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = False
                    .menuDraft.Enabled = False
                    .menuCheckInOut.Enabled = False
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = False
                    .btnUndo.Enabled = False
                    .btnRedo.Enabled = False
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = False
                    .btnDeleteUnit.Enabled = False
                    .btnChangeSequence.Enabled = False
                    .btnUpdateMRD.Enabled = False
                    .togFurtherBasicSpecification.Enabled = True
                    .togInstrumentation.Enabled = True
                    .togMfcSpecification.Enabled = True
                    .togNonMFCSpecification.Enabled = True
                    .togProgramInformation.Enabled = True
                    .togShowAll.Enabled = True
                    .togTiming.Enabled = True
                    .togUpdatePack.Enabled = True
                    .togUserShipping.Enabled = True
                    .btnUpdateColumns.Enabled = False
                    .btnUpdateHoliday.Enabled = False
                    .btnCDSIDtoDvpTeam.Enabled = False
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = True
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = False
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = False
                    If Form.DataCenter.GlobalSections.TimeLineSection.Columns.Hidden = False Then
                        .togTiming.Checked = True
                    Else
                        .togTiming.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False Then
                        .togUpdatePack.Checked = True
                    Else
                        .togUpdatePack.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False Then
                        .togUserShipping.Checked = True
                    Else
                        .togUserShipping.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False Then
                        .togFurtherBasicSpecification.Checked = True
                    Else
                        .togFurtherBasicSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False Then
                        .togProgramInformation.Checked = True
                    Else
                        .togProgramInformation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False Then
                        .togMfcSpecification.Checked = True
                    Else
                        .togMfcSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False Then
                        .togInstrumentation.Checked = True
                    Else
                        .togInstrumentation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False Then
                        .togNonMFCSpecification.Checked = True
                    Else
                        .togNonMFCSpecification.Checked = False
                    End If

                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub

        Private Sub Specific_Checkedout_OwnerAndExecutor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = False
                    .menuDraft.Enabled = True
                    .menuCheckInOut.Enabled = True
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = True
                    .btnUndo.Enabled = True
                    '.btnRedo.Enabled = True
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = True
                    .btnDeleteUnit.Enabled = True
                    .btnChangeSequence.Enabled = True
                    .btnUpdateMRD.Enabled = True
                    .togFurtherBasicSpecification.Enabled = True
                    .togInstrumentation.Enabled = True
                    .togMfcSpecification.Enabled = True
                    .togNonMFCSpecification.Enabled = True
                    .togProgramInformation.Enabled = True
                    .togShowAll.Enabled = True
                    .togTiming.Enabled = True
                    .togUpdatePack.Enabled = True
                    .togUserShipping.Enabled = True
                    .btnUpdateColumns.Enabled = True
                    .btnUpdateHoliday.Enabled = True
                    .btnCDSIDtoDvpTeam.Enabled = True
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = True
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = True
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = True

                    '------------------------------
                    ' Check-In/-Out logic
                    '------------------------------
                    .btnCheckOut.Visible = False
                    .btnCheckOut.Enabled = False

                    .btnCheckIn.Enabled = True
                    .btnCheckIn.Visible = True

                    .btnDiscard.Enabled = True
                    .btnDiscard.Visible = True
                    '-----------------------------

                    '------------------------------
                    ' Draft logic
                    '------------------------------
                    .btnGenerateDraft.Visible = True
                    .btnGenerateDraft.Enabled = True
                    .mnuGenerateDraft.Visible = True
                    .mnuGenerateDraft.Enabled = True

                    .mnuActiveUsers.Visible = True
                    .mnuActiveUsers.Enabled = True

                    .btnDeleteDraft.Visible = False
                    .btnReplacePlanWithDraft.Visible = False
                    '------------------------------

                    If Form.DataCenter.GlobalSections.TimeLineSection.Columns.Hidden = False Then
                        .togTiming.Checked = True
                    Else
                        .togTiming.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False Then
                        .togUpdatePack.Checked = True
                    Else
                        .togUpdatePack.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False Then
                        .togUserShipping.Checked = True
                    Else
                        .togUserShipping.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False Then
                        .togFurtherBasicSpecification.Checked = True
                    Else
                        .togFurtherBasicSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False Then
                        .togProgramInformation.Checked = True
                    Else
                        .togProgramInformation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False Then
                        .togMfcSpecification.Checked = True
                    Else
                        .togMfcSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False Then
                        .togInstrumentation.Checked = True
                    Else
                        .togInstrumentation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False Then
                        .togNonMFCSpecification.Checked = True
                    Else
                        .togNonMFCSpecification.Checked = False
                    End If

                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub

        Private Sub Specific_Checkedout_Visitor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = False
                    .menuDraft.Enabled = False
                    .menuCheckInOut.Enabled = False
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = False
                    .btnUndo.Enabled = False
                    .btnRedo.Enabled = False
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = False
                    .btnDeleteUnit.Enabled = False
                    .btnChangeSequence.Enabled = False
                    .btnUpdateMRD.Enabled = False
                    .togFurtherBasicSpecification.Enabled = True
                    .togInstrumentation.Enabled = True
                    .togMfcSpecification.Enabled = True
                    .togNonMFCSpecification.Enabled = True
                    .togProgramInformation.Enabled = True
                    .togShowAll.Enabled = True
                    .togTiming.Enabled = True
                    .togUpdatePack.Enabled = True
                    .togUserShipping.Enabled = True
                    .btnUpdateColumns.Enabled = False
                    .btnUpdateHoliday.Enabled = False
                    .btnCDSIDtoDvpTeam.Enabled = False
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = True
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = False
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = False
                    If Form.DataCenter.GlobalSections.TimeLineSection.Columns.Hidden = False Then
                        .togTiming.Checked = True
                    Else
                        .togTiming.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False Then
                        .togUpdatePack.Checked = True
                    Else
                        .togUpdatePack.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False Then
                        .togUserShipping.Checked = True
                    Else
                        .togUserShipping.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False Then
                        .togFurtherBasicSpecification.Checked = True
                    Else
                        .togFurtherBasicSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False Then
                        .togProgramInformation.Checked = True
                    Else
                        .togProgramInformation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False Then
                        .togMfcSpecification.Checked = True
                    Else
                        .togMfcSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False Then
                        .togInstrumentation.Checked = True
                    Else
                        .togInstrumentation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False Then
                        .togNonMFCSpecification.Checked = True
                    Else
                        .togNonMFCSpecification.Checked = False
                    End If

                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub

        Private Sub Specific_Draft_OwnerOrExecutor()
            Try
                With Globals.Ribbons.RbnTnDControlPanel
                    .btnLoadOpenTnDPlan.Enabled = True
                    .btnConvertToSpecific.Enabled = False
                    .menuDraft.Enabled = True
                    .menuCheckInOut.Enabled = False
                    .btnRefreshPlan.Enabled = True
                    .btnRefreshUnit.Enabled = True
                    .btnUndo.Enabled = True
                    '.btnRedo.Enabled = True
                    .btnSearchFilter.Enabled = True
                    .btnClearFilter.Enabled = True
                    .btnAddUnit.Enabled = True
                    .btnDeleteUnit.Enabled = True
                    .btnChangeSequence.Enabled = True
                    .btnUpdateMRD.Enabled = True
                    .togFurtherBasicSpecification.Enabled = True
                    .togInstrumentation.Enabled = True
                    .togMfcSpecification.Enabled = True
                    .togNonMFCSpecification.Enabled = True
                    .togProgramInformation.Enabled = True
                    .togShowAll.Enabled = True
                    .togTiming.Enabled = True
                    .togUpdatePack.Enabled = True
                    .togUserShipping.Enabled = True
                    .btnUpdateColumns.Enabled = True
                    .btnUpdateHoliday.Enabled = True
                    .btnCDSIDtoDvpTeam.Enabled = True
                    .btnExportToExcel.Enabled = True
                    .btnUnitReport.Enabled = True
                    .btnEngineTransmissionReport.Enabled = True
                    .btnPrecheckF4T.Enabled = True
                    .btnCountReport.Enabled = True
                    .btnPustFit4Test.Enabled = True
                    '------------------------------
                    ' Draft logic
                    '------------------------------
                    .btnGenerateDraft.Visible = False
                    .mnuGenerateDraft.Visible = False

                    .mnuActiveUsers.Visible = False

                    .btnDeleteDraft.Visible = True
                    .btnDeleteDraft.Enabled = True
                    .btnReplacePlanWithDraft.Visible = True
                    .btnReplacePlanWithDraft.Enabled = True
                    '------------------------------


                    If Form.DataCenter.GlobalSections.TimeLineSection.Columns.Hidden = False Then
                        .togTiming.Checked = True
                    Else
                        .togTiming.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UpdatePackSection.Columns.Hidden = False Then
                        .togUpdatePack.Checked = True
                    Else
                        .togUpdatePack.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.UserShippingDetailsSection.Columns.Hidden = False Then
                        .togUserShipping.Checked = True
                    Else
                        .togUserShipping.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Columns.Hidden = False Then
                        .togFurtherBasicSpecification.Checked = True
                    Else
                        .togFurtherBasicSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.ProgramInformationSection.Columns.Hidden = False Then
                        .togProgramInformation.Checked = True
                    Else
                        .togProgramInformation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.MfcSpecificationSection.Columns.Hidden = False Then
                        .togMfcSpecification.Checked = True
                    Else
                        .togMfcSpecification.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.InstrumentationSection.Columns.Hidden = False Then
                        .togInstrumentation.Checked = True
                    Else
                        .togInstrumentation.Checked = False
                    End If

                    If Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Columns.Hidden = False Then
                        .togNonMFCSpecification.Checked = True
                    Else
                        .togNonMFCSpecification.Checked = False
                    End If

                End With
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Active and deactive Ribbon", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub




        Public Sub DeactiveRibbonButtonsState()
            With Globals.Ribbons.RbnTnDControlPanel


                '------------------------------
                ' Check-In/-Out logic
                '------------------------------
                .menuCheckInOut.Enabled = False
                .menuDraft.Enabled = False

                .btnCheckOut.Visible = False
                .btnCheckOut.Enabled = False

                .btnCheckIn.Enabled = False
                .btnCheckIn.Visible = False

                .btnDiscard.Enabled = False
                .btnDiscard.Visible = False


                .btnSearchFilter.Enabled = False
                .btnClearFilter.Enabled = False

                .btnConvertToSpecific.Enabled = False

                .mnuGenerateDraft.Enabled = False
                .mnuActiveUsers.Enabled = False


                .btnUndo.Enabled = False
                .btnRedo.Enabled = False

                .btnUpdateHoliday.Enabled = False
                .btnRefreshPlan.Enabled = False
                .btnRefreshUnit.Enabled = False
                .btnAddUnit.Enabled = False
                .btnDeleteUnit.Enabled = False
                .btnChangeSequence.Enabled = False
                .btnUpdateMRD.Enabled = False

                .togInstrumentation.Enabled = False
                .togNonMFCSpecification.Enabled = False
                .togMfcSpecification.Enabled = False
                .togProgramInformation.Enabled = False
                .togFurtherBasicSpecification.Enabled = False
                .togUserShipping.Enabled = False
                .togUpdatePack.Enabled = False
                .togShowAll.Enabled = False
                .togTiming.Enabled = False

                .btnUpdateColumns.Enabled = False


                .btnCDSIDtoDvpTeam.Enabled = False

                .btnExportToExcel.Enabled = False
                .btnUnitReport.Enabled = False
                .btnEngineTransmissionReport.Enabled = False
                .btnPrecheckF4T.Enabled = False
                .btnCountReport.Enabled = False
                .btnPustFit4Test.Enabled = False
                .tglBtnValidatePlan.Enabled = False
                Globals.Ribbons.RbnTnDControlPanel.TGMessages.Enabled = False
                Globals.Ribbons.RbnTnDControlPanel.btnTodayIndicator.Enabled = False
                Try
                    If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                        Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Visible = False
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Visible = False
                        Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Visible = False
                        'Globals.Ribbons.RbnTnDControlPanel.SepCountReport.Visible = False
                        'Globals.Ribbons.RbnTnDControlPanel.SepPushToFit4Test.Visible = False
                        'Globals.Ribbons.RbnTnDControlPanel.SepPrecheckReport.Visible = False
                    Else
                        Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Visible = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Visible = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Visible = True
                        'Globals.Ribbons.RbnTnDControlPanel.SepCountReport.Visible = True
                        'Globals.Ribbons.RbnTnDControlPanel.SepPushToFit4Test.Visible = True
                        'Globals.Ribbons.RbnTnDControlPanel.SepPrecheckReport.Visible = True
                    End If

                Catch ex As Exception

                End Try


            End With
        End Sub

        Public Sub UpdateUndoButtonsState()
            Try
                Dim _ChangeLog As New CT.Data.ChangeLog
                If _ChangeLog.IsUndoCommandAvailable(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType) = True Then
                    Globals.Ribbons.RbnTnDControlPanel.btnUndo.Enabled = True
                Else
                    Globals.Ribbons.RbnTnDControlPanel.btnUndo.Enabled = False
                End If
            Catch ex As Exception
            End Try
        End Sub
    End Class
End Namespace