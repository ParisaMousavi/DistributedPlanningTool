Namespace Form.DataCenter
    Public Class ModuleFunction
        Public Sub sbProtectPlan()

            Dim rng As Excel.Range = Nothing

            Try
                With Form.DataCenter.GlobalValues.WS

                    Try
                        .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                        .EnableOutlining = True
                    Catch ex As Exception
                    End Try

                    ' DrawFileStatus()

                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        Try
                            .Range(.Cells(1, 1), .Cells(4, .UsedRange.Columns.Count)).Locked = True
                            .Range(.Cells(5, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).Locked = True
                        Catch ex As Exception
                        End Try
                    Else
                        Try
                            .Range(.Cells(1, 1), .Cells(4, .UsedRange.Columns.Count)).Locked = False
                            .Range(.Cells(5, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).Locked = False
                        Catch ex As Exception
                        End Try

                        Try
                            .Range(.Cells(1, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn),
                               .Cells(.UsedRange.Rows.Count, Form.DataCenter.GlobalSections.TimeLineSectionLastColumn)).Locked = True
                        Catch ex As Exception
                        End Try

                        Try
                            .Range(.Cells(1, "A"),
                            .Cells(4, Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn)).Locked = True
                        Catch ex As Exception
                        End Try

                        .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Locked = True
                        .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column)).Locked = False
                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Locked = False
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column)).Locked = False
                            .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column), .Cells(.UsedRange.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Locked = False
                        End If
                    End If
                    .Protect(Form.DataCenter.GlobalValues.ConstPwd, True, True, False, True, True, False, False, False, False, False, False, False, False, True, True)
                End With
            Catch ex As Exception
            End Try
        End Sub

        Public Sub DisableRibbonButtonsForViewer()
            Try
                Dim RibCon As Object

                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPProgramTDToolbar.Items
                    RibCon.enabled = False
                Next

                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPShowHideSpecificationsections.Items
                    RibCon.enabled = False
                Next

                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPAddDeleteUpdateVehicle.Items
                    RibCon.enabled = False
                Next

                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPPlan.Items
                    RibCon.enabled = False
                Next


                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPReport.Items
                    RibCon.enabled = False
                Next

                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPIndicator.Items
                    RibCon.enabled = False
                Next

                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPHolidayMaster.Items
                    RibCon.enabled = False
                Next
                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPIndicator.Items
                    RibCon.enabled = False
                Next


                Globals.Ribbons.RbnTnDControlPanel.btnLoadOpenTnDPlan.Enabled = True
                Globals.Ribbons.RbnTnDControlPanel.btnRefreshPlan.Enabled = True

                With Form.DataCenter.GlobalValues.WS
                    Try
                        .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                        .EnableOutlining = True
                    Catch ex As Exception
                    End Try
                    Try
                        .Range(.Cells(1, 1), .Cells(4, .UsedRange.Columns.Count)).Locked = True
                        .Range(.Cells(5, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).Locked = True
                        If Form.DataCenter.ProgramConfig.IsGeneric = False And Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Master.ToString Then
                            For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPShowHideSpecificationsections.Items
                                RibCon.enabled = True
                            Next
                            Globals.Ribbons.RbnTnDControlPanel.btnUpdateColumns.Enabled = False
                            Form.DataCenter.GlobalSections.InstrumentationSection.Locked = True
                            Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Locked = True
                            Form.DataCenter.GlobalSections.MfcSpecificationSection.Locked = True
                            Form.DataCenter.GlobalSections.ProgramInformationSection.Locked = True
                            Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Locked = True
                            Form.DataCenter.GlobalSections.UserShippingDetailsSection.Locked = True
                            Form.DataCenter.GlobalSections.UpdatePackSection.Locked = True
                            Globals.Ribbons.RbnTnDControlPanel.btnTodayIndicator.Enabled = True
                        ElseIf Form.DataCenter.ProgramConfig.IsGeneric = False And Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Checkedout.ToString Then

                            If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                                For Each RibCon In Globals.Ribbons.RbnTnDControlPanel.GRPShowHideSpecificationsections.Items
                                    RibCon.enabled = True
                                Next
                                Globals.Ribbons.RbnTnDControlPanel.btnTodayIndicator.Enabled = True
                                Globals.Ribbons.RbnTnDControlPanel.btnUpdateColumns.Enabled = False
                                Form.DataCenter.GlobalSections.InstrumentationSection.Locked = True
                                Form.DataCenter.GlobalSections.NonMfcSpecificationSection.Locked = True
                                Form.DataCenter.GlobalSections.MfcSpecificationSection.Locked = True
                                Form.DataCenter.GlobalSections.ProgramInformationSection.Locked = True
                                Form.DataCenter.GlobalSections.FurtherBasicInformationSection.Locked = True
                                Form.DataCenter.GlobalSections.UserShippingDetailsSection.Locked = True
                                Form.DataCenter.GlobalSections.UpdatePackSection.Locked = True
                            End If
                        End If

                        Globals.Ribbons.RbnTnDControlPanel.btnExportToExcel.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnUnitReport.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnEngineTransmissionReport.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Enabled = True
                        'Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Enabled = True

                        .Protect(Form.DataCenter.GlobalValues.ConstPwd, True, True, False, True, True, False, False, False, False, False, False, False, False, True, True)
                    Catch ex As Exception
                    End Try
                End With

                If Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Master.ToString Or
                    (Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or
                    Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "") Then
                    With Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString)
                        Try
                            .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                            .EnableOutlining = True
                        Catch ex As Exception
                        End Try
                        Try
                            .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).EntireColumn.Locked = True
                            .Protect(Form.DataCenter.GlobalValues.ConstPwd, True, True, False, True, True, False, False, False, False, False, False, False, False, True, True)
                        Catch ex As Exception
                        End Try
                    End With
                End If

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
        End Sub
        Public Sub DisableRibbonButtonsForMaster_Draft_CheckedOut()
            Dim _strUserPermissionLevel As String = String.Empty
            Try
                If Form.DataCenter.ProgramConfig.HCID <> 0 Then
                    Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
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
                        System.Windows.Forms.MessageBox.Show(ex.Message, "Disable Ribbon buttons", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                    End Try
                End If
                If Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Master.ToString Or
                    (Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or
                    Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "") Then
                    With Globals.ThisAddIn.Application.Worksheets(Form.DataCenter.WorkSheet.ChangeLogs.ToString)
                        Try
                            .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                            .EnableOutlining = True
                        Catch ex As Exception
                        End Try
                        Try
                            .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count)).EntireColumn.Locked = True
                            .Protect(Form.DataCenter.GlobalValues.ConstPwd, True, True, False, True, True, False, False, False, False, False, False, False, False, True, True)
                        Catch ex As Exception
                        End Try
                    End With
                End If
                If Form.DataCenter.ProgramConfig.IsGeneric Then
                    DisableRibbonButtonsForViewer()
                    Globals.Ribbons.RbnTnDControlPanel.btnConvertToSpecific.Enabled = True
                    Exit Sub
                End If
                If Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Master.ToString Then
                    DisableRibbonButtonsForViewer()
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") <> CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower And Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim <> "" Then
                        Globals.Ribbons.RbnTnDControlPanel.menuCheckInOut.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnCheckIn.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnCheckOut.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnDiscard.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.menuDraft.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnRefreshUnit.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.tglBtnValidatePlan.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnTodayIndicator.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Enabled = True 'Added 10 Oct 18
                        Globals.Ribbons.RbnTnDControlPanel.TGMessages.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.loadMenubutton()
                    End If
                ElseIf Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Draft.ToString Then
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") <> CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower And Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim <> "" Then
                        Globals.Ribbons.RbnTnDControlPanel.menuCheckInOut.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.btnCheckIn.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.btnCheckOut.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.btnDiscard.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.menuDraft.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnRefreshUnit.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.tglBtnValidatePlan.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.TGMessages.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.btnTodayIndicator.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.loadMenubutton()
                    Else
                        DisableRibbonButtonsForViewer()
                    End If
                ElseIf Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Checkedout.ToString Then
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        DisableRibbonButtonsForViewer()
                    Else
                        Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Enabled = False
                        Globals.Ribbons.RbnTnDControlPanel.menuDraft.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.tglBtnValidatePlan.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.TGMessages.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnTodayIndicator.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Enabled = True
                        Globals.Ribbons.RbnTnDControlPanel.loadMenubutton()
                    End If
                End If
                If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                    Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Visible = False
                    Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Visible = False
                    Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Visible = False
                    ' Globals.Ribbons.RbnTnDControlPanel.SepCountReport.Visible = False
                    ' Globals.Ribbons.RbnTnDControlPanel.SepPushToFit4Test.Visible = False
                    ' Globals.Ribbons.RbnTnDControlPanel.SepPrecheckReport.Visible = False
                Else
                    Globals.Ribbons.RbnTnDControlPanel.btnCountReport.Visible = True
                    Globals.Ribbons.RbnTnDControlPanel.btnPustFit4Test.Visible = True
                    Globals.Ribbons.RbnTnDControlPanel.btnPrecheckF4T.Visible = True
                    'Globals.Ribbons.RbnTnDControlPanel.SepCountReport.Visible = True
                    'Globals.Ribbons.RbnTnDControlPanel.SepPushToFit4Test.Visible = True
                    'Globals.Ribbons.RbnTnDControlPanel.SepPrecheckReport.Visible = True
                End If
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Disable Ribbon buttons", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub
        'Public Sub DrawFileStatus()
        '    Try

        '        Dim strMode As String = ""
        '        If Form.DataCenter.ProgramConfig.IsGeneric Then
        '            strMode = "Mode - Generic plan"
        '        Else
        '            If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString Then
        '                strMode = "Mode - Checked out"
        '            ElseIf Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Draft.ToString Then
        '                strMode = "Mode - Draft plan"
        '            ElseIf Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then
        '                strMode = "Mode - Master plan"
        '            End If
        '        End If

        '        Dim objShp As Excel.Shape = Nothing

        '        With Form.DataCenter.GlobalValues.WS
        '            Try
        '                For Each objShp In .Shapes
        '                    If objShp.Name Like "txtFileStatus*" Then
        '                        objShp.Delete()
        '                    End If
        '                Next
        '            Catch ex As Exception
        '            End Try
        '            objShp = .Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 550, 10, 200, 40)
        '            With objShp
        '                .ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset38
        '                .TextFrame2.TextRange.Characters.Text = strMode
        '                .Name = "txtFileStatus"
        '                .Locked = True
        '            End With
        '            With objShp.TextFrame2.TextRange.Characters(1, objShp.TextFrame2.TextRange.Characters.Count).ParagraphFormat
        '                .FirstLineIndent = 0
        '                .Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter
        '            End With
        '            With objShp.TextFrame2.TextRange.Characters(1, objShp.TextFrame2.TextRange.Characters.Count).Font
        '                .NameComplexScript = "+mn-cs"
        '                .NameFarEast = "+mn-ea"
        '                .Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
        '                .Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorLight1
        '                .Fill.ForeColor.TintAndShade = 0
        '                .Fill.ForeColor.Brightness = 0
        '                .Fill.Transparency = 0
        '                .Fill.Solid()
        '                .Size = 21
        '                .Name = "+mn-lt"
        '            End With
        '        End With
        '    Catch ex As Exception
        '        System.Windows.Forms.MessageBox.Show(ex.Message, "Draw file status display shape", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        '    End Try
        'End Sub
        Public Sub DisplayMasterMessage()
            'Try

            '    Dim objShp As Excel.Shape

            '    With Form.DataCenter.GlobalValues.WS
            '        .Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
            '        Try
            '            For Each objShp In .Shapes
            '                If objShp.Name Like "RectMasterModeDisplay*" Then
            '                    objShp.Delete()
            '                End If
            '            Next
            '        Catch ex As Exception
            '        End Try
            '        objShp = .Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle, 600, 75, 800, 200)
            '        objShp.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlFreeFloating
            '    End With
            '    With objShp
            '        .Name = "RectMasterModeDisplay"
            '        .ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset41
            '        .TextFrame2.TextRange.Font.Size = 32
            '        .TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle
            '        .TextFrame2.TextRange.Characters.Text = "The TnD plan has been opened in """"Master"""" mode. The plan is non editable. If you need to make any changes, please checkout the plan."
            '        With .TextFrame2.TextRange.Characters(1, .TextFrame2.TextRange.Characters.Count).ParagraphFormat
            '            .FirstLineIndent = 0
            '            .Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter
            '        End With
            '        With .TextFrame2.TextRange.Characters(1, .TextFrame2.TextRange.Characters.Count).Font
            '            .BaselineOffset = 0
            '            .NameComplexScript = "+mn-cs"
            '            .NameFarEast = "+mn-ea"
            '            .Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
            '            .Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorLight1
            '            .Fill.ForeColor.TintAndShade = 0
            '            .Fill.ForeColor.Brightness = 0
            '            .Fill.Transparency = 0
            '            .Fill.Solid()
            '            .Size = 32
            '            .Name = "+mn-lt"
            '            .Shadow.Type = Microsoft.Office.Core.MsoShadowType.msoShadow21
            '        End With
            '        With .ThreeD
            '            .SetPresetCamera(Microsoft.Office.Core.MsoPresetCamera.msoCameraOrthographicFront)
            '            .RotationX = 0
            '            .RotationY = 0
            '            .RotationZ = 0
            '            .FieldOfView = 0
            '            .LightAngle = 145
            '            .PresetLighting = Microsoft.Office.Core.MsoLightRigType.msoLightRigBalanced
            '            .PresetMaterial = Microsoft.Office.Core.MsoPresetMaterial.msoMaterialMatte
            '            .Depth = 0
            '            .ContourWidth = 0
            '            .BevelTopType = Microsoft.Office.Core.MsoBevelType.msoBevelCircle
            '            .BevelTopInset = 15
            '            .BevelTopDepth = 3
            '            .BevelBottomType = Microsoft.Office.Core.MsoBevelType.msoBevelNone
            '        End With
            '        With .Shadow
            '            .Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
            '            .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
            '            .Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
            '            .Blur = 3.5
            '            .OffsetX = 0.00000000000000013471114791
            '            .OffsetY = 2.2
            '            .RotateWithShape = Microsoft.Office.Core.MsoTriState.msoTrue
            '            .ForeColor.RGB = RGB(0, 0, 0)
            '            .Transparency = 0.6800000072
            '            .Size = 100
            '        End With
            '        .Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            '    End With
            'Catch ex As Exception
            '    System.Windows.Forms.MessageBox.Show(ex.Message, "Display message for master mode.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            'Finally
            '    sbProtectPlan()
            '    Globals.Ribbons.RbnTnDControlPanel.tmrDisplay.Enabled = True
            'End Try
            Try


                Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                sbPutWatermark()
                If Form.DataCenter.ProgramConfig.HCID <> 0 Then
                    Dim objPer As New CT.Data.Authorization, objRestrictUser As New Form.DataCenter.ModuleFunction
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
                        System.Windows.Forms.MessageBox.Show(ex.Message, "Worksheet events", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                    Finally
                        If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                            objRestrictUser.DisableRibbonButtonsForViewer()
                        Else
                            Dim clsobj As New Form.DataCenter.ModuleFunction
                            clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                        End If
                    End Try
                End If
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "'Read only' watermark.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub
        Sub sbPutWatermark(Optional strText As String = "")
            Try
                Dim strPicFName As String
                Dim shpTxtCDSID As Excel.Shape
                Dim shpGrpCDSID As Excel.Shape
                Dim FSO As Scripting.FileSystemObject
                FSO = New Scripting.FileSystemObject
                If strText = "" Then strText = "Read only!"
                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.CutCopyMode = False

                With Form.DataCenter.GlobalValues.WS
                    .Activate()
                    strPicFName = System.IO.Path.GetTempPath & "\Watermark_Temp.jpg"
                    If FSO.FileExists(strPicFName) Then FSO.DeleteFile(strPicFName)
                    shpGrpCDSID = .Shapes.AddChart(Excel.XlChartType.xlLine, 0, 0, .Range("A1:U50").Width, .Range("A1:U50").Height)
                    shpGrpCDSID.Name = "chWatermarkChart"
                    shpGrpCDSID.Chart.ChartType = Excel.XlChartType.xlLine
                    shpTxtCDSID = .Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect3, strText, "+mn-lt", 100, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, 366.8075590551, 211.7104724409)
                    shpTxtCDSID.IncrementRotation(322)
                    shpTxtCDSID.Name = "txtCDSID"
                    shpTxtCDSID.Cut()
                    .ChartObjects("chWatermarkChart").Activate
                    Globals.ThisAddIn.Application.ActiveChart.Paste()
                    Globals.ThisAddIn.Application.CutCopyMode = False
                    Globals.ThisAddIn.Application.ActiveChart.Shapes(0).Left = .Range("A1:U50").Width / 3.5
                    Globals.ThisAddIn.Application.ActiveChart.Shapes(0).Top = .Range("A1:U50").Height / 2.7
                    .ChartObjects("chWatermarkChart").Activate
                    .ChartObjects("chWatermarkChart").Chart.ChartArea.ClearContents
                    .ChartObjects("chWatermarkChart").Chart.ChartArea.ClearFormats
                    Globals.ThisAddIn.Application.ActiveSheet.Shapes("chWatermarkChart").Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    Globals.ThisAddIn.Application.ActiveSheet.Shapes("chWatermarkChart").Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    Globals.ThisAddIn.Application.ActiveChart.Export(Filename:=strPicFName, FilterName:="JPG")
                    .SetBackgroundPicture(strPicFName)
                    .PageSetup.CenterHeaderPicture.Filename = strPicFName
                    .PageSetup.CenterHeader = "&G"
                    If FSO.FileExists(strPicFName) Then FSO.DeleteFile(strPicFName)
                    .ChartObjects("chWatermarkChart").Delete
                End With

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.CutCopyMode = False
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "'Read only' watermark.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub
    End Class
End Namespace