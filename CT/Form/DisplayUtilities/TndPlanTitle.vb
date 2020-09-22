Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data

Namespace Form.DisplayUtilities
    Public Class TndPlanTitle

        Public Function FillMismatchedQty() As String
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Dim myDataTable, myDataTable_CT, myDataTable_XCC As New DataTable
            Dim _obj As New Form.DataCenter.ModuleFunction

            Dim _PlanInterface As Data.Interfaces.PlanInterface

            If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString() Then
                _PlanInterface = New Data.VehiclePlan.Plan
            ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString() Then
                _PlanInterface = New Data.RigPlan.Plan
            Else
                Exit Function
            End If

            'Result set alignment should be as below
            'DBSOURCE    Vehicle	Buck	Rig	  Rebuild
            'CT             26	    3	    1	    NULL
            'XCC            23	    2	    NULL	NULL

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
            FillMismatchedQty = String.Empty
            Try



                '------------------------------------------------------
                ' The XCC values are dispayed in CT but if XCC DB is down the CT
                ' shuold work properly. Therefore we have validation on CT value 
                ' but not on XCC values.
                '------------------------------------------------------
                myDataTable_XCC = _PlanInterface.GetQuantityTableXCC(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
                myDataTable_CT = _PlanInterface.GetQuantityTableCT(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

                '------------------------------------------------------
                ' Only validation on CT values.
                '------------------------------------------------------
                If myDataTable_CT Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)

                If IsNothing(myDataTable_XCC) = False Then
                    If myDataTable_XCC.Rows.Count > 0 Then
                        myDataTable_CT.Merge(myDataTable_XCC, False)
                    End If
                End If

                myDataTable = myDataTable_CT

                If myDataTable Is Nothing Then  Exit Function
                If myDataTable.Rows.Count <= 0 Then
                    Exit Function
                ElseIf myDataTable.Rows.Count = 2 Then
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    With Form.DataCenter.GlobalValues.WS.Shapes

                        Dim Shape_Counts As Excel.Shape
                        Try
                            Shape_Counts = .Item("txtVehiclesXCC")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(1).Item(1))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(1).Item(1))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtVehiclesCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(1))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(1))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtBucksXCC")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(1).Item(2))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(1).Item(2))
                        Catch ex As Exception
                        End Try
                        Try
                            Shape_Counts = .Item("txtBucksCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(2))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(2))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRigsXCC")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(1).Item(3))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(1).Item(3))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRigsCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(3))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(3))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRebuildsXcc")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(1).Item(4))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(1).Item(4))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRebuildsCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(4))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(4))
                        Catch ex As Exception
                        End Try

                    End With
                ElseIf myDataTable.Rows.Count = 1 Then
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                    With Form.DataCenter.GlobalValues.WS.Shapes
                        Dim Shape_Counts As Excel.Shape

                        Try
                            Shape_Counts = .Item("txtVehiclesXCC")
                            Shape_Counts.OLEFormat.Object.text = ""
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = ""
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtVehiclesCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(1))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(1))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtBucksXCC")
                            Shape_Counts.OLEFormat.Object.text = ""
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = ""
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtBucksCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(2))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(2))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRigsXCC")
                            Shape_Counts.OLEFormat.Object.text = ""
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = ""
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRigsCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(3))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(3))
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRebuildsXcc")
                            Shape_Counts.OLEFormat.Object.text = ""
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = ""
                        Catch ex As Exception
                        End Try

                        Try
                            Shape_Counts = .Item("txtRebuildsCT")
                            Shape_Counts.OLEFormat.Object.text = fnCheckNull(myDataTable.Rows(0).Item(4))
                            Shape_Counts.TextFrame2.TextRange.Characters.Text = fnCheckNull(myDataTable.Rows(0).Item(4))
                        Catch ex As Exception
                        End Try

                    End With
                End If

            Catch ex As Exception
                FillMismatchedQty = "FillMismatchedQty : " + ex.Message
            Finally
                Globals.ThisAddIn.Application.ScreenUpdating = False
                _obj.sbProtectPlan()
            End Try
        End Function
        Private Function fnCheckNull(Value As Object) As String
            If IsDBNull(Value) Then
                Return ""
            Else
                Return Value.ToString()
            End If
        End Function
        Public Function LoadAndFormatLabel() As String
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.EnableEvents = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            Dim strTempTitle1, strTempTitle2 As String
            Dim myDatatable As System.Data.DataTable
            Dim _Program As CT.Data.ProgramConfiguration = New CT.Data.ProgramConfiguration
            myDatatable = _Program.SelectProgramConfigs(Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)

            LoadAndFormatLabel = String.Empty
            Try
                Dim HeaderLabel As Excel.Shape = Form.DataCenter.GlobalValues.WS.Shapes.Item("txtPlanHeader")
                Globals.ThisAddIn.Application.ScreenUpdating = False

                If myDatatable.Rows.Count > 0 Then

                    strTempTitle1 = "HCID's-" & If(myDatatable.Rows(0)("PairedHealthChartId").ToString = String.Empty, myDatatable.Rows(0)("HealthChartId"), myDatatable.Rows(0)("PairedHealthChartId")) & vbLf & myDatatable.Rows(0)("ProgramDescription") & vbLf & myDatatable.Rows(0)("BuildPhases")
                    If Form.DataCenter.ProgramConfig.IsMainPlan = True Then
                        strTempTitle2 = vbLf & "Confidential" & vbLf & "T&D Plan - Issue " & myDatatable.Rows(0)("TnDReleaseStatus") & IIf(Form.DataCenter.ProgramConfig.IsGeneric = False, "(Specific)", "(Generic)") & vbLf & myDatatable.Rows(0)("BuildTypes") & vbLf & myDatatable.Rows(0)("TnDPlanner")
                    Else
                        strTempTitle2 = vbLf & "Confidential" & vbLf & "T&D Plan - Issue " & myDatatable.Rows(0)("TnDReleaseStatus") & IIf(Form.DataCenter.ProgramConfig.IsGeneric = False, "(Specific-Draft)", "(Generic-Draft)") & vbLf & myDatatable.Rows(0)("BuildTypes") & vbLf & myDatatable.Rows(0)("TnDPlanner")
                    End If

                    HeaderLabel.TextFrame2.TextRange.Characters.Text = strTempTitle1 & strTempTitle2

                    With HeaderLabel.TextFrame2.TextRange.Characters(1, Len(strTempTitle1))
                        .Font.Size = 11
                        .Font.Name = "Arial"
                        .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                    End With

                    With HeaderLabel.TextFrame2.TextRange.Characters(Len(strTempTitle1) + 1, 13)
                        .Font.Size = 10
                        .Font.Name = "Arial"
                        .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                        .Font.Fill.ForeColor.RGB = RGB(255, 0, 0)

                    End With

                    With HeaderLabel.TextFrame2.TextRange.Characters(Len(strTempTitle1) + 15, Len(strTempTitle2) - 14)
                        .Font.Size = 9
                        .Font.Name = "Arial"
                        .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                        .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                    End With
                Else
                    strTempTitle1 = "HCID's-" & Form.DataCenter.ProgramConfig.HCID & vbLf & Form.DataCenter.ProgramConfig.HCIDName & vbLf & Form.DataCenter.ProgramConfig.BuildPhase
                    If Form.DataCenter.ProgramConfig.IsMainPlan = True Then
                        strTempTitle2 = vbLf & "Confidential" & vbLf & "T&D Plan - Issue 1.0 " & IIf(Form.DataCenter.ProgramConfig.IsGeneric = False, "(Specific)", "(Generic)") & vbLf & Form.DataCenter.ProgramConfig.BuildType.ToString() '"Vehicle"
                    Else
                        strTempTitle2 = vbLf & "Confidential" & vbLf & "T&D Plan - Issue 1.0 " & IIf(Form.DataCenter.ProgramConfig.IsGeneric = False, "(Specific-Draft)", "(Generic-Draft)") & vbLf & Form.DataCenter.ProgramConfig.BuildType.ToString() '"Vehicle"
                    End If

                    HeaderLabel.TextFrame2.TextRange.Characters.Text = strTempTitle1 & strTempTitle2

                    With HeaderLabel.TextFrame2.TextRange.Characters(1, Len(strTempTitle1))
                        .Font.Size = 11
                        .Font.Name = "Arial"
                        .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                    End With

                    With HeaderLabel.TextFrame2.TextRange.Characters(Len(strTempTitle1) + 1, 13)
                        .Font.Size = 10
                        .Font.Name = "Arial"
                        .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                        .Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
                    End With

                    With HeaderLabel.TextFrame2.TextRange.Characters(Len(strTempTitle1) + 15, Len(strTempTitle2) - 14)
                        .Font.Size = 9
                        .Font.Name = "Arial"
                        .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                        .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                    End With
                End If

                'If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString Then
                '    HeaderLabel.Fill.Solid()
                '    HeaderLabel.Fill.Visible = True
                '    HeaderLabel.Fill.BackColor.RGB = RGB(170, 170, 170)
                '    HeaderLabel.Fill.TwoColorGradient(Microsoft.Office.Core.MsoGradientStyle.msoGradientHorizontal, 1)
                'End If

                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString Then
                        With HeaderLabel.Fill
                            .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                            .ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6 ' green
                            .ForeColor.TintAndShade = 0
                            .ForeColor.Brightness = 0.6000000238
                            .Transparency = 0
                            .Solid()
                        End With
                    ElseIf Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Checkedout.ToString Then 'Or Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Draft.ToString Then
                        With HeaderLabel.Fill
                            .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                            .ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent4 ' yellow
                            .ForeColor.TintAndShade = 0
                            .ForeColor.Brightness = 0.6000000238
                            .Transparency = 0
                            .Solid()
                        End With
                    End If
                End If

            Catch ex As Exception
                LoadAndFormatLabel = "LoadAndFormatLabel: " + ex.Message
            End Try
        End Function



    End Class
End Namespace