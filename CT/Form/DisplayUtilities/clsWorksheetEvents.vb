
Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Drawing
Imports Microsoft.Office.Core
Imports System.Linq
Imports System.Collections
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Namespace Form.DisplayUtilities
    ' References issue reported by Parisa on 13th Mar 2018.
    '1- when he open a plan And then lock the window And come back And login after a while he has no event. I had seen this issue before by Marcel.
    '2- He should work with more than 7 excel file simultaneously And after a while he cannot scroll Or select a cell when he has an open TndPlan.
    Public Class clsWorksheetEvents
        Public objCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane = Nothing
        Public WithEvents objCustomTaskPaneMessages As Microsoft.Office.Tools.CustomTaskPane = Nothing
        Public WithEvents XLApp As Excel.Application = Nothing
        Private Const VK_LBUTTON = &H1
        Private Const VK_RBUTTON = &H2
        Private WithEvents WSOps As Microsoft.Office.Tools.Excel.Worksheet
        Dim PEG, PED, XCCG, XCCD, PTM, PTA, XCCTM, XCCTA As Office.CommandBarButton
        Dim colEngineData As New List(Of String)()
        Dim colEngineDataXCC As New List(Of String)()
        Dim colTransData = New List(Of String)()
        Dim colTransDataXCC = New List(Of String)()
        Dim bolRightClicked As Boolean = False
        Dim bolRightClicked2 As Boolean = False
        Public _CusContMnu As Form.TndContextMenu.CustomContextMenu
        Public bolDisablePopupButtons As Boolean = False
        WithEvents columnwidth As Office.CommandBarButton

        <DllImport("user32.dll")>
        Shared Function GetAsyncKeyState(ByVal vKey As System.Windows.Forms.Keys) As Short
        End Function

        Public Sub New()
            WSOps = Form.DataCenter.GlobalValues.WS
            _CusContMnu = New Form.TndContextMenu.CustomContextMenu()
            XLApp = WSOps.Application
            If objCustomTaskPane Is Nothing Then
                objCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(New ProcessStepInfoTskPane(), "Process step information", WSOps.Application.ActiveWindow)
                objCustomTaskPane.Visible = False
                objCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                objCustomTaskPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal
            End If

            If objCustomTaskPaneMessages Is Nothing Then
                objCustomTaskPaneMessages = Globals.ThisAddIn.CustomTaskPanes.Add(New MessageTaskPaneControl(), "You've got a message!", WSOps.Application.ActiveWindow)
                objCustomTaskPaneMessages.Visible = False
                objCustomTaskPaneMessages.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                objCustomTaskPaneMessages.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal
            End If

        End Sub
        Private Sub WSOps_BeforeRightClick(Target As Range, ByRef Cancel As Boolean) Handles WSOps.BeforeRightClick

            Try
                bolDisablePopupButtons = False
                'WSOps.Application.ScreenUpdating = False
                bolRightClicked2 = True
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                    If Form.DataCenter.ProgramConfig.FileStatus = Data.DataCenter.FileStatus.Master.ToString Or (Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "") Then
                        bolDisablePopupButtons = True
                    End If
                Else
                    Cancel = True
                    bolRightClicked = False
                    Exit Sub
                End If

                bolRightClicked = True

                If WSOps.Parent.Name Like "TndTemplate*" Then
                    DeleteContextMenu()
                    WSOps.Application.CommandBars("Cell").Reset()
                    If (Target.Column >= Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And Target.Column <= Form.DataCenter.GlobalSections.TimeLineSectionLastColumn) Then
                        Cancel = True
                    End If
                    If Target.Row > 4 And Target.Column > Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn Then
                        DisplayCopySpec(Target)
                        Exit Sub
                    End If
                    If Target.Cells.Count > 1 Then
                        If Target.Cells(1, 1).Interior.Color <> Integer.Parse(CT.My.Resources.EmptyColor) Then
                            Target.Cells(1, 1).Select()
                        Else
                            WSOps_SelectionChange(Target, Cancel)
                        End If
                    Else
                        WSOps_SelectionChange(Target, Cancel)
                    End If
                End If
            Catch ex As Exception
                ' MessageBox.Show(ex.Message, "Right click event", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                bolRightClicked2 = False
                WSOps.Application.ScreenUpdating = True
            End Try
        End Sub



        Private Sub WSOps_Change(Target As Excel.Range) Handles WSOps.Change
            If Form.DataCenter.GlobalValues.bolPlanDrawInProgress = True Then Exit Sub
            Dim TGT As Excel.Range = Nothing

            Try
                'Delete excess copied data (more than last row)
                Globals.ThisAddIn.Application.EnableEvents = False
                Dim strAddress As String
                strAddress = Target.Address
                If strAddress.Split("$").Length <= 3 Then
                    If strAddress.Split("$")(2) > Form.DataCenter.ProgramConfig.LastRow Then
                        WSOps.Range("A" & Form.DataCenter.ProgramConfig.LastRow + 1 & ":ZZ" & strAddress.Split("$")(2) & "").ClearContents()
                        Target = WSOps.Range(strAddress.Split("$")(1) & Form.DataCenter.ProgramConfig.LastRow)
                    End If
                Else
                    If strAddress.Split("$")(4) > Form.DataCenter.ProgramConfig.LastRow Then
                        WSOps.Range("A" & Form.DataCenter.ProgramConfig.LastRow + 1 & ":ZZ" & strAddress.Split("$")(4) & "").ClearContents()
                        If strAddress.Split("$")(2).Replace(":", "") > Form.DataCenter.ProgramConfig.LastRow Then
                            'Target = WSOps.Range(strAddress.Split("$")(1) & strAddress.Split("$")(2) & strAddress.Split("$")(3) & strAddress.Split("$")(2))
                            Exit Try
                        Else
                            Target = WSOps.Range(strAddress.Split("$")(1) & strAddress.Split("$")(2) & strAddress.Split("$")(3) & Form.DataCenter.ProgramConfig.LastRow)
                        End If
                    End If
                End If
                Form.DisplayUtilities.PlanSections.SevenTabsFunctions.UpdateData(Target, WSOps)

            Catch ex As Exception
                Dim Cls As New Form.DataCenter.GlobalFunctions
                For Each TGT In Target.Rows
                    If TGT.Row > Form.DataCenter.ProgramConfig.LastRow Then Exit For
                    Cls.UpdateSection(TGT.Row, TGT.Row,,,, CDate(Form.DataCenter.GlobalValues.WS.Cells(4, Form.DataCenter.GlobalValues.WS.Application.Selection.Column).value2))
                Next
                MessageBox.Show(ex.Message, "Edit unit information", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Globals.ThisAddIn.Application.EnableEvents = True
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
                DeleteContextMenu()
                WSOps.Application.CommandBars("Cell").Reset()

            End Try

        End Sub

        Private Sub WSOps_SelectionChange(Target As Range, Optional ByRef Cancel As Boolean = False) Handles WSOps.SelectionChange

            'Try
            '    If (Target.Column > Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And
            '        Target.Column < Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn) And
            '        (Integer.Parse(GetAsyncKeyState(System.Windows.Forms.Keys.Right).ToString) = -32768 Or
            '        Integer.Parse(GetAsyncKeyState(System.Windows.Forms.Keys.Left).ToString) = -32768 Or
            '        Integer.Parse(GetAsyncKeyState(System.Windows.Forms.Keys.Down).ToString) = -32768 Or
            '        Integer.Parse(GetAsyncKeyState(System.Windows.Forms.Keys.Up).ToString) = -32768) Then
            '        Exit Sub
            '    End If
            'Catch ex As Exception
            'End Try

            Dim TimeLineSectionFirstColumn As Integer = Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn

            Try
                If (Target.Column > Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And
                    Target.Column < TimeLineSectionFirstColumn) Then
                    Exit Sub
                End If
            Catch ex As Exception
            End Try




            Try
                If bolRightClicked2 = False And Integer.Parse(GetAsyncKeyState(System.Windows.Forms.Keys.RButton).ToString) = -32768 Then
                    Exit Sub
                End If
            Catch ex As Exception
            End Try

            Dim bolDisTskPn As Boolean = False

            If Form.DataCenter.GlobalValues.bolPlanDrawInProgress = True Then Exit Sub

            Dim sCellValue As String = String.Empty

            If Form.DataCenter.ProgramConfig.IsWithCustomFormatting Then
                ChangeColumnWidth(Target)
                Form.DataCenter.GlobalValues.bolfrmwidth = False
            End If

            Dim intcount As Int32 = 0
            Dim rng2 As Excel.Range = Nothing
            Dim rng3 As Excel.Range = Nothing

            Dim rngSelcolor As Excel.Range = Nothing
            Dim intFCol As Integer = 0, intLCol As Integer = 0
            Dim objPro As New Form.DataCenter.ModuleFunction
            If WSOps.Application.ActiveWorkbook.Name.ToString <> WSOps.Parent.Name.ToString Then Exit Sub
            WSOps.Application.EnableEvents = False

            Try


                If _CusContMnu Is Nothing Then _CusContMnu = New Form.TndContextMenu.CustomContextMenu()

                Dim IsGeneric As Boolean = Form.DataCenter.ProgramConfig.IsGeneric
                Dim ScreenUpdating As Boolean = WSOps.Application.ScreenUpdating
                Dim LastRow As Long = Form.DataCenter.ProgramConfig.LastRow
                Dim FirstRow As Long = Form.DataCenter.ProgramConfig.FirstRow
                Dim FileStatus As String = Form.DataCenter.ProgramConfig.FileStatus
                Dim strUserPermissionLevel As String = Form.DataCenter.GlobalValues.strUserPermissionLevel
                Dim TimeLineSectionLastColumn As Integer = Form.DataCenter.GlobalSections.TimeLineSectionLastColumn
                Dim strUserCaseSelected As String = Form.DataCenter.GlobalValues.strUserCaseSelected

                If IsGeneric = False And ScreenUpdating = True And
                    Target.Row > 4 And Target.Row <= LastRow And Target.Cells.Count = 1 And bolRightClicked = True And FileStatus <> Data.DataCenter.FileStatus.Master.ToString And
                    strUserPermissionLevel.ToLower.Replace(" ", "") <> CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower And strUserPermissionLevel.Trim <> "" Then

                    If (Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column Or Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column) Then
                        DisplayPopup(Target)
                    End If

                    If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column Then
                        DisplayPopupPaintFacility()
                    End If

                    If (Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column Or Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column) Then
                        Try
                            DisplayPopup_Team()
                        Catch ex As Exception
                        End Try
                    End If
                End If

                If Target.Column >= TimeLineSectionFirstColumn And Target.Column <= TimeLineSectionLastColumn And
                         Target.Row <= LastRow Then
                    sbToggleCutCopyAndPaste(False)
                Else
                    sbToggleCutCopyAndPaste(True)
                End If

                If strUserCaseSelected <> "" Then
                    If Globals.ThisAddIn.Application.Intersect(Form.DataCenter.GlobalValues.WS.Range(strUserCaseSelected), Target) Is Nothing Then
                        Form.DataCenter.GlobalValues.strUserCaseSelected = ""
                        Form.DataCenter.GlobalValues.bolUserCaseSelected = False
                    End If
                End If
                If Form.DataCenter.GlobalValues.strSelAllAddress <> "" Then
                    If Globals.ThisAddIn.Application.Intersect(Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strSelAllAddress), Target) Is Nothing Then
                        Form.DataCenter.GlobalValues.strSelAllAddress = ""
                        Form.DataCenter.GlobalValues.bolSelAll = False
                    End If
                End If

                If Not Form.DataCenter.ProgramConfig.ISSearchActive Then
                    WSOps.Cells.FormatConditions.Delete()
                End If

                '-------------------------------------------------------------------------------------------------------------------
                ' For Tnd Plan Area
                '-------------------------------------------------------------------------------------------------------------------
                Form.DisplayUtilities.Utilities.FindFLCols(Target.Row, intFCol, intLCol)

                If Target.Column >= TimeLineSectionFirstColumn And Target.Column <= TimeLineSectionLastColumn And Target.Columns.Count = 1 And
                    Target.Rows.Count = 1 And Target.Interior.Color <> Integer.Parse(CT.My.Resources.EmptyColor) And Target.Row >= FirstRow And
                    Target.Row <= LastRow Then
                    '-------------------------------------------------------------------------------------------------------------------
                    ' Colored area
                    '-------------------------------------------------------------------------------------------------------------------
                    strUserCaseSelected = Form.DataCenter.GlobalValues.strUserCaseSelected
                    If strUserCaseSelected <> "" Or Form.DataCenter.GlobalValues.strSelAllAddress <> "" Then
                        If strUserCaseSelected <> "" Then
                            If Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Range(strUserCaseSelected), Target) Is Nothing Then
                                Form.DataCenter.GlobalValues.strUserCaseSelected = ""
                                Form.DataCenter.GlobalValues.bolUserCaseSelected = False
                            Else
                                rngSelcolor = Globals.ThisAddIn.Application.Range(strUserCaseSelected)
                            End If
                        End If
                        If Form.DataCenter.GlobalValues.strSelAllAddress <> "" Then
                            If Globals.ThisAddIn.Application.Intersect(Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strSelAllAddress), Target) Is Nothing Then
                                Form.DataCenter.GlobalValues.strSelAllAddress = ""
                                Form.DataCenter.GlobalValues.bolSelAll = False
                            Else
                                rngSelcolor = Globals.ThisAddIn.Application.Range(Form.DataCenter.GlobalValues.strSelAllAddress)
                            End If
                        End If
                    Else
                        With WSOps.Range(WSOps.Cells(WSOps.Application.Selection.row, TimeLineSectionFirstColumn), WSOps.Cells(WSOps.Application.Selection.row, WSOps.Application.Selection.column))
                            rng2 = .Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, False, Type.Missing, Type.Missing)
                        End With

                        With WSOps.Range(WSOps.Cells(WSOps.Application.Selection.row, WSOps.Application.Selection.COLUMN), WSOps.Cells(WSOps.Application.Selection.row, TimeLineSectionLastColumn))
                            rng3 = .Find("*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, Type.Missing, Type.Missing)
                        End With

                        If rng2 IsNot Nothing And rng3 IsNot Nothing Then
                            rngSelcolor = WSOps.Range(WSOps.Cells(WSOps.Application.Selection.row, rng2.Column), WSOps.Cells(WSOps.Application.Selection.row, rng3.Column - 1))
                            Try
                                If bolRightClicked = False Then
                                    bolDisTskPn = True
                                End If
                            Catch ex As Exception
                            End Try
                        End If

                        Try
                            Form.DataCenter.GlobalValues.sEditId = WSOps.Cells(WSOps.Application.Selection.row, rng2.Column).Formula.Split(";")(0).Replace("=CellFace(", "").Replace("""", "").Trim()
                        Catch ex As Exception
                        End Try
                    End If

                    If Not Form.DataCenter.ProgramConfig.ISSearchActive Then
                        'WSOps.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                        Dim objFC As FormatCondition
                        With rngSelcolor
                            objFC = .FormatConditions.Add(XlFormatConditionType.xlExpression,, "=True")
                            objFC.Interior.Color = Color.White
                            objFC.Font.Color = Color.Black
                            objFC.Font.Bold = True
                        End With
                        'objPro.sbProtectPlan()
                    End If

                    If rngSelcolor IsNot Nothing Then
                        '  WSOps.Application.EnableEvents = False
                        WSOps.Application.ScreenUpdating = False
                        rngSelcolor.Select()
                        'WSOps.Application.DisplayFormulaBar = False
                        'WSOps.Application.DisplayFormulaBar = True
                        WSOps.Application.ScreenUpdating = True

                    End If

                    Try
                        If bolRightClicked = True Then
                            If Form.DataCenter.GlobalValues.bolCopy = False And Form.DataCenter.GlobalValues.bolCut = False Then
                                _CusContMnu.AddToCellMenu(1, rngSelcolor.Address, bolDisablePopupButtons)
                            Else
                                _CusContMnu.AddToCellMenu(6, rngSelcolor.Address, bolDisablePopupButtons)
                            End If

                        End If


                    Catch ex As Exception

                    End Try

                ElseIf Target.Column > TimeLineSectionFirstColumn And Target.Column < TimeLineSectionLastColumn And Target.Columns.Count = 1 And WSOps.Application.Selection.cells.Count = 1 And
                    Target.Interior.Color = Integer.Parse(CT.My.Resources.EmptyColor) And Target.Row >= FirstRow And Target.Row <= LastRow Then
                    '-------------------------------------------------------------------------------------------------------------------
                    ' Not colored area
                    '-------------------------------------------------------------------------------------------------------------------

                    Try
                        'If WSOps.Application.CutCopyMode <> XlCutCopyMode.xlCopy And WSOps.Application.CutCopyMode <> XlCutCopyMode.xlCut Then
                        If bolRightClicked Then
                            If Form.DataCenter.GlobalValues.bolCopy = False And Form.DataCenter.GlobalValues.bolCut = False Then
                                If Target.Column > intFCol Then
                                    If Target.Column < intLCol Then
                                        _CusContMnu.AddToCellMenu(5, WSOps.Application.Selection.Address, bolDisablePopupButtons)
                                    Else
                                        _CusContMnu.AddToCellMenu(2, WSOps.Application.Selection.Address, bolDisablePopupButtons)
                                    End If
                                Else
                                    _CusContMnu.AddToCellMenu(4, WSOps.Application.Selection.Address, bolDisablePopupButtons)
                                End If
                            Else
                                _CusContMnu.AddToCellMenu(3, WSOps.Application.Selection.Address, bolDisablePopupButtons)
                            End If

                        End If
                    Catch ex As Exception
                    End Try
                ElseIf Target.Interior.Color = Integer.Parse(CT.My.Resources.EmptyColor) And Target.Column >= TimeLineSectionFirstColumn And
                    Target.Columns.Count = 1 And Target.Column <= TimeLineSectionLastColumn And
                        WSOps.Application.Selection.cells.Count > 1 And Target.Row > FirstRow - 1 And Target.Row <= LastRow Then
                    'If WSOps.Application.CutCopyMode <> XlCutCopyMode.xlCopy And WSOps.Application.CutCopyMode <> XlCutCopyMode.xlCut Then
                    If bolRightClicked Then
                        If Form.DataCenter.GlobalValues.bolCopy = False And Form.DataCenter.GlobalValues.bolCut = False Then
                            _CusContMnu.AddToCellMenu(8, WSOps.Application.Selection.Address, bolDisablePopupButtons)
                        End If
                    End If
                End If
            Catch ex As Exception
                WSOps.Application.EnableEvents = True
                WSOps.Application.ScreenUpdating = True
                MsgBox(ex.Message)
            Finally
                Dim bolIsValidObject As Boolean = False
                Try
                    Dim test As Boolean = objCustomTaskPane.Visible
                    bolIsValidObject = True
                Catch ex As Exception
                    bolIsValidObject = False
                End Try


                If bolIsValidObject Then
                    objCustomTaskPane.Control.Visible = bolDisTskPn
                    objCustomTaskPane.Visible = bolDisTskPn
                    If bolDisTskPn Then
                        objCustomTaskPane.Width = 300
                        Dim obj As Object = objCustomTaskPane.Control
                        obj.loadData()
                    End If
                End If
                bolRightClicked = False
                bolRightClicked2 = False
                Cancel = True
                WSOps.Application.EnableEvents = True
                WSOps.Application.ScreenUpdating = True
            End Try
        End Sub
        Private Sub ChangeColumnWidth(Index1 As Range)
            Try
                DeleteContextMenu()
                Dim Index = Index1.Column
                If Index1.Row = 1 Then
                    If Index > Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn And Index < Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                    If Index > Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                    If Index > Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                    If Index > Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If

                    If Index > Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                    If Index > Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                    If Index > Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn And Index < Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                    If Index > Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn And Index < Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn Then
                        showcolumnwidth(Index)
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub
        Private Sub showcolumnwidth(Index As Integer)
            Dim ContextMenu As Office.CommandBar
            ContextMenu = Globals.ThisAddIn.Application.CommandBars("Cell")
            columnwidth = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
            With columnwidth
                .Caption = "Set Column Width"
                .Tag = Index
            End With
            WSOps.Application.EnableEvents = True
            Try
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then ContextMenu.ShowPopup()
            Catch ex As Exception
            End Try

        End Sub


        Public Sub DeleteContextMenu()

            Try
                For Each btn As Office.CommandBarControl In Globals.ThisAddIn.Application.CommandBars("Cell").Controls
                    btn.Delete()
                Next
            Catch ex As Exception
            End Try

        End Sub

        Public Sub ShowMessageTaskPane(bolDisTskPn_Message As Boolean)
            Try
                Dim bolIsValidObject As Boolean = False
                Try
                    Dim test As Boolean = objCustomTaskPaneMessages.Visible
                    bolIsValidObject = True
                Catch ex As Exception
                    bolIsValidObject = False
                End Try

                If bolIsValidObject Then
                    objCustomTaskPaneMessages.Visible = bolDisTskPn_Message
                    If bolDisTskPn_Message Then
                        objCustomTaskPaneMessages.Width = 300
                        Dim obj As Object = objCustomTaskPaneMessages.Control
                        obj.loadData()
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Show hide messages", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        Private Sub ButtonClick(ByVal ctrl As Office.CommandBarButton, ByRef Cancel As Boolean) Handles columnwidth.Click

            Try
                If Form.DataCenter.GlobalValues.bolfrmwidth = False Then
                    Form.DataCenter.GlobalValues.bolfrmwidth = True
                    Dim Index As Integer = Convert.ToUInt32(ctrl.Tag)
                    If Index > Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn And Index < Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn Then
                        Dim strSection As String = ""
                        Dim strHeader = CType(Form.DataCenter.GlobalValues.WS.Cells(4, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.VehicleProgramInfoSection, strHeader, strSection)

                    End If
                    If Index > Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn Then
                        Dim strSection As String = CType(Form.DataCenter.GlobalValues.WS.Cells(2, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        Dim strInstrumentation As String = CType(Form.DataCenter.GlobalValues.WS.Cells(3, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.InstrumentationSection, strInstrumentation, strSection)
                    End If
                    If Index > Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn Then
                        Dim strSection As String = ""
                        Dim strNonMFC As String = CType(Form.DataCenter.GlobalValues.WS.Cells(3, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection, strNonMFC, strSection)
                        Exit Sub
                    End If
                    If Index > Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn Then
                        Dim strMFCSection As String = CType(Form.DataCenter.GlobalValues.WS.Cells(2, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        Dim strMFC As String = CType(Form.DataCenter.GlobalValues.WS.Cells(3, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection, strMFC, strMFCSection)
                        Exit Sub
                    End If

                    If Index > Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn Then
                        Dim strMFCSection As String = ""
                        Dim strProgramInfo As String = CType(Form.DataCenter.GlobalValues.WS.Cells(3, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection, strProgramInfo, strMFCSection)
                        Exit Sub
                    End If
                    If Index > Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn And Index < Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn Then
                        Dim strSection As String = ""
                        Dim strFurtherBasic As String = CType(Form.DataCenter.GlobalValues.WS.Cells(3, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection, strFurtherBasic, strSection)
                        Exit Sub
                    End If
                    If Index > Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn And Index < Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn Then
                        Dim strSection As String = ""
                        Dim strUserShipping As String = CType(Form.DataCenter.GlobalValues.WS.Cells(3, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection, strUserShipping, strSection)
                        Exit Sub
                    End If
                    If Index > Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn And Index < Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn Then
                        Dim strSection As String = ""
                        Dim strUpdatepack As String = CType(Form.DataCenter.GlobalValues.WS.Cells(2, Index), Excel.Range).MergeArea.Cells(1, 1).Value
                        updatecolumnwidth(Index, Form.DataCenter.GlobalSections.SectionName.UpdatePackSection, strUpdatepack, strSection)
                        Exit Sub
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub
        Private Sub updatecolumnwidth(Index As Integer, GroupId As Integer, strHeader As String, strSection As String)
            Dim _planCustomfor As CT.Data.PlanIndivitualFormatting = New CT.Data.PlanIndivitualFormatting()
            Dim dr As DialogResult
            Dim _frmcolwidth As New frmColumnWidth
            Try
                _frmcolwidth.txtColumnwidth.Text = Form.DataCenter.GlobalValues.WS.Columns(Index).entirecolumn.ColumnWidth.ToString()
                _frmcolwidth.StartPosition = FormStartPosition.CenterScreen
                'MessageBox.Show("Dia")
                dr = _frmcolwidth.ShowDialog()
                If dr = DialogResult.OK Then
                    If Not _frmcolwidth.txtColumnwidth.Text = Form.DataCenter.GlobalValues.WS.Columns(Index).entirecolumn.ColumnWidth.ToString() Then
                        _planCustomfor.UpdateSettings(Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType, GroupId, strHeader, strSection, Environment.UserDomainName + "\" + Environment.UserName, _frmcolwidth.txtColumnwidth.Text)
                        Form.DataCenter.GlobalValues.WS.Columns(Index).ColumnWidth = _frmcolwidth.txtColumnwidth.Text
                        ' _frmcolwidth.Close()
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub
        Public Sub sbToggleCutCopyAndPaste(bolAllow As Boolean)

            On Error Resume Next

            With Globals.ThisAddIn.Application
                .OnKey("%{F10}", "sbCreateReport_Engine_Trans")
                If Not bolAllow Then
                    .OnKey("^c", "CutCopyPasteDisabled")
                    .OnKey("^v", "CutCopyPasteDisabled")
                    .OnKey("^x", "CutCopyPasteDisabled")
                    If Form.DataCenter.ProgramConfig.IsGeneric = False Then
                        .OnKey("{DELETE}", "sbDelete_PS_UC")
                    Else
                        .OnKey("{DELETE}", "CutCopyPasteDisabled")
                    End If
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Trim.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        .OnKey("{DELETE}", "CutCopyPasteDisabled")
                        .OnKey("%{F10}")
                    End If

                    If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString And
                    (Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") <> CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower And Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim <> "") Then
                        .OnKey("{DELETE}", "CutCopyPasteDisabled")
                        .OnKey("%{F10}")
                    End If

                    .OnKey("{INSERT}", "CutCopyPasteDisabled")
                    .OnKey("{BACKSPACE}", "CutCopyPasteDisabled")
                    .OnKey("{ESC}", "sbCancel")
                Else
                    .OnKey("^c")
                    .OnKey("^v")
                    .OnKey("^x")
                    .OnKey("{DELETE}")
                    If Form.DataCenter.ProgramConfig.IsGeneric = False And
                            Globals.ThisAddIn.Application.Selection.column = 5 And
                            Globals.ThisAddIn.Application.ActiveSheet.name = WSOps.Name Then
                        .OnKey("{DELETE}", "sbDeleteVehicle")
                    Else
                        .OnKey("{DELETE}")
                    End If
                    .OnKey("{INSERT}")
                    .OnKey("{BACKSPACE}")
                    .OnKey("{ESC}")

                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Trim.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        .OnKey("^c", "CutCopyPasteDisabled")
                        .OnKey("^v", "CutCopyPasteDisabled")
                        .OnKey("^x", "CutCopyPasteDisabled")
                        .OnKey("{DELETE}", "CutCopyPasteDisabled")
                        .OnKey("{INSERT}", "CutCopyPasteDisabled")
                        .OnKey("{BACKSPACE}", "CutCopyPasteDisabled")
                        .OnKey("{ESC}", "sbCancel")
                    End If

                    If Form.DataCenter.ProgramConfig.FileStatus = CT.Data.DataCenter.FileStatus.Master.ToString And
                    (Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") <> CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower And Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim <> "") Then
                        .OnKey("^c", "CutCopyPasteDisabled")
                        .OnKey("^v", "CutCopyPasteDisabled")
                        .OnKey("^x", "CutCopyPasteDisabled")
                        .OnKey("{DELETE}", "CutCopyPasteDisabled")
                        .OnKey("{INSERT}", "CutCopyPasteDisabled")
                        .OnKey("{BACKSPACE}", "CutCopyPasteDisabled")
                        .OnKey("{ESC}", "sbCancel")
                    End If
                End If
            End With

        End Sub
        Public Sub DisplayCopySpec(rngTGT As Range)

            If Not (rngTGT.Column >= Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And rngTGT.Column <= Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn) Then
                Exit Sub
            End If

            Dim _CusContMnu As Form.TndContextMenu.CustomContextMenu = New Form.TndContextMenu.CustomContextMenu()
            _CusContMnu.AddToCellMenu(7, rngTGT.Address, bolDisablePopupButtons)

            Globals.ThisAddIn.Application.ScreenUpdating = True

        End Sub
        Public Sub DisplayPopup_Team()

            Dim ContextMenu As Office.CommandBar
            Dim ShtProgConfig_ As Object = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString)

            Dim _CusContMnu As Form.TndContextMenu.CustomContextMenu = New Form.TndContextMenu.CustomContextMenu()
            _CusContMnu.DeleteContextMenu()

            Dim objValid As Excel.Validation, rngTemp As Excel.Range, rng As Excel.Range
            objValid = WSOps.Application.Selection.Validation
            rng = ShtProgConfig_.Range(Strings.Split(objValid.Formula1, "!")(1))
            ContextMenu = Globals.ThisAddIn.Application.CommandBars("Cell")
            For Each rngTemp In rng.Cells
                With ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                    .Caption = rngTemp.Value
                    .OnAction = "'sbPutValue_Others" & " """ & rngTemp.Value & """'"
                    .FaceId = 327
                End With
            Next

            Try
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then ContextMenu.ShowPopup()
            Catch ex As Exception
            End Try

        End Sub
        Public Sub DisplayPopup(rng As Range)
            Dim ProgramEngines As CommandBarControl
            Dim XCCEngines As CommandBarControl
            Dim ProgramTransmission As CommandBarControl
            Dim XCCTransmissions As CommandBarControl


            Dim ContextMenu As Office.CommandBar

            ' If Form.DataCenter.GlobalValues.KeyPressed(1) Then Exit Sub

            Try
                GetETData(rng)
                ContextMenu = Globals.ThisAddIn.Application.CommandBars("Cell")
                If rng.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column Then
                    Dim _CusContMnu As Form.TndContextMenu.CustomContextMenu = New Form.TndContextMenu.CustomContextMenu()
                    _CusContMnu.DeleteContextMenu()
                    ProgramEngines = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=1)
                    With ProgramEngines
                        .Caption = "Program Engines"
                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=1)
                            .Caption = "Gas"
                            For intCnt = 0 To colEngineData.Count - 1
                                If Strings.Split(colEngineData.Item(intCnt), "~")(0) = "Gas" Then
                                    PEG = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With PEG
                                        .Caption = Strings.Split(colEngineData.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Engine" & " """ & colEngineData.Item(intCnt) & """'"
                                        .FaceId = 548
                                    End With
                                End If
                            Next
                        End With

                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=2)
                            .Caption = "Diesel"
                            For intCnt = 0 To colEngineData.Count - 1
                                If Strings.Split(colEngineData.Item(intCnt), "~")(0) = "Diesel" Then
                                    PED = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With PED
                                        .Caption = Strings.Split(colEngineData.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Engine" & " """ & colEngineData.Item(intCnt) & """'"
                                        .FaceId = 548
                                    End With
                                End If
                            Next

                        End With
                    End With

                    XCCEngines = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=2)

                    With XCCEngines
                        .Caption = "XCC Engines"
                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=1)
                            .Caption = "Gas"
                            For intCnt = 0 To colEngineDataXCC.Count - 1
                                If Strings.Split(colEngineDataXCC.Item(intCnt), "~")(0) = "Gas" Then
                                    XCCG = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With XCCG
                                        .Caption = Strings.Split(colEngineDataXCC.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Engine" & " """ & colEngineDataXCC.Item(intCnt) & """'"
                                        .FaceId = 548
                                    End With
                                End If
                            Next
                        End With
                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=2)
                            .Caption = "Diesel"
                            For intCnt = 0 To colEngineDataXCC.Count - 1
                                If Strings.Split(colEngineDataXCC.Item(intCnt), "~")(0) = "Diesel" Then
                                    XCCD = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With XCCD
                                        .Caption = Strings.Split(colEngineDataXCC.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Engine" & " """ & colEngineDataXCC.Item(intCnt) & """'"
                                        .FaceId = 548
                                    End With
                                End If
                            Next
                        End With
                    End With

                    With ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=3)
                        .Caption = "No Engine"
                        .OnAction = "'sbPutValue_Engine" & " """ & "-~-" & """'"
                        .FaceId = 330
                    End With


                    Try
                        If Form.DataCenter.ProgramConfig.IsGeneric = False Then ContextMenu.ShowPopup()
                    Catch ex As Exception
                    End Try

                    colEngineData.Clear()
                    colEngineDataXCC.Clear()

                ElseIf rng.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column Then
                    Dim _CusContMnu As Form.TndContextMenu.CustomContextMenu = New Form.TndContextMenu.CustomContextMenu()
                    _CusContMnu.DeleteContextMenu()
                    ProgramTransmission = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=1)
                    With ProgramTransmission
                        .Caption = "Program Transmission"
                        .Tag = "ProgramTransmission_Tag"

                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=1)
                            .Caption = "Manual"
                            For intCnt = 0 To colTransData.Count - 1
                                If Strings.Split(colTransData.Item(intCnt), "~")(0) = "Manual" Then
                                    PTM = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With PTM
                                        .Caption = Strings.Split(colTransData.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Trans" & " """ & colTransData.Item(intCnt) & """'"
                                        .FaceId = 3049
                                    End With
                                End If
                            Next

                        End With
                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=2)
                            .Caption = "Automatic"
                            For intCnt = 0 To colTransData.Count - 1
                                If Strings.Split(colTransData.Item(intCnt), "~")(0) = "Automatic" Then
                                    PTA = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With PTA
                                        .Caption = Strings.Split(colTransData.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Trans" & " """ & colTransData.Item(intCnt) & """'"
                                        .FaceId = 3049
                                    End With
                                End If
                            Next
                        End With
                    End With

                    XCCTransmissions = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=2)

                    With XCCTransmissions
                        .Caption = "XCC Transmissions"
                        .Tag = "XCCTransmissions_Tag"

                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=1)
                            .Caption = "Manual"
                            For intCnt = 0 To colTransDataXCC.Count - 1
                                If Strings.Split(colTransDataXCC.Item(intCnt), "~")(0) = "Manual" Then
                                    XCCTM = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With XCCTM
                                        .Caption = Strings.Split(colTransDataXCC.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Trans" & " """ & colTransDataXCC.Item(intCnt) & """'"
                                        .FaceId = 3049
                                    End With
                                End If
                            Next
                        End With
                        With .Controls.Add(Type:=Office.MsoControlType.msoControlPopup, Before:=2)
                            .Caption = "Automatic"
                            For intCnt = 0 To colTransDataXCC.Count - 1
                                If Strings.Split(colTransDataXCC.Item(intCnt), "~")(0) = "Automatic" Then
                                    XCCTA = .Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                                    With XCCTA
                                        .Caption = Strings.Split(colTransDataXCC.Item(intCnt), "~")(1)
                                        .OnAction = "'sbPutValue_Trans" & " """ & colTransDataXCC.Item(intCnt) & """'"
                                        .FaceId = 3049
                                    End With
                                End If
                            Next
                        End With
                    End With

                    With ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=3)
                        .Caption = "No Transmission"
                        .OnAction = "'sbPutValue_Trans" & " """ & "-~-" & """'"
                        .FaceId = 330
                    End With

                    Try

                        If Form.DataCenter.ProgramConfig.IsGeneric = False Then ContextMenu.ShowPopup()
                    Catch ex As Exception
                    End Try

                    colTransData.clear()
                    colTransDataXCC.clear()
                End If
            Catch ex As Exception

            End Try


        End Sub
        Public Sub DisplayPopupPaintFacility()

            Try
                Dim ContextMenu As Office.CommandBar
                Dim ShtProgConfig_ As Object = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(DataCenter.WorkSheet.ProgramConfig.ToString)
                Dim _CusContMnu As Form.TndContextMenu.CustomContextMenu = New Form.TndContextMenu.CustomContextMenu()
                Dim objData As New CT.Data.PaintFacility
                Dim _DT As System.Data.DataTable = objData.SelectAll
                Dim _DR As System.Data.DataRow = Nothing

                _CusContMnu.DeleteContextMenu()
                ContextMenu = Globals.ThisAddIn.Application.CommandBars("Cell")
                For Each _DR In _DT.Rows
                    With ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton)
                        .Caption = _DR("PaintCode").ToString '& ", " & _DR("Color").ToString
                        '.OnAction = "'sbPutValue_Paint" & " """ & _DR("PaintCode").ToString & "~" & _DR("Color").ToString & """'"
                        .OnAction = "'sbPutValue_Paint" & " """ & _DR("PaintCode").ToString & "~" & "" & """'"
                        .FaceId = 108
                    End With
                Next
                If Form.DataCenter.ProgramConfig.IsGeneric = False Then ContextMenu.ShowPopup()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "DisplayPopupPaintFacility", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try

        End Sub
        Private Sub GetETData(rng As Range)
            ' Dim _program As CT.Data.Program
            Dim _Engine As CT.Data.Engine
            Dim EngineListinExl As New ArrayList()
            Dim EngineTypeListinExl As New ArrayList()
            Dim TransmissionListinExl As New ArrayList()
            Dim TransmissionTypeListinExl As New ArrayList()
            Dim lRow As Long
            Dim rst As New System.Data.DataTable
            Dim sTnDRegion As String = String.Empty
            Dim _transmission As CT.Data.Transmission

            ' If Form.DataCenter.GlobalValues.KeyPressed(1) Then Exit Sub

            Try
                If rng.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column Or rng.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column Then
                    ' _program = New CT.Data.Program()
                    'Form.DataCenter.VehicleConfig.VehicleHCID = 192
                    sTnDRegion = Form.DataCenter.ProgramConfig.Region  ' _program.GetXccLead(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.VehicleConfig.VehicleHCID)
                End If
                If rng.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column Then
                    With Form.DataCenter.GlobalValues.WS
                        'lRow = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column), .Cells(.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column)).End(XlDirection.xlUp).Row
                        lRow = Form.DataCenter.ProgramConfig.LastRow
                    End With
                    For index = 5 To lRow
                        Dim xRng As Excel.Range = CType(Form.DataCenter.GlobalValues.WS.Cells(index, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Column), Excel.Range)
                        Dim xEngineTypeRng As Excel.Range = CType(Form.DataCenter.GlobalValues.WS.Cells(index, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Engine_Type_Column), Excel.Range)
                        Dim val As Object = xRng.Value()
                        If val <> "" And val <> "-" Then
                            If Not EngineListinExl.Contains(val) Then
                                EngineListinExl.Add(val)
                                EngineTypeListinExl.Add(xEngineTypeRng.Value())
                            End If
                        End If
                    Next
                    _Engine = New CT.Data.Engine()
                    rst = _Engine.GetXccEngineList(sTnDRegion)
                    If rst IsNot Nothing Then
                        For Each row As DataRow In rst.Rows
                            If EngineListinExl.Contains(CStr(row("EngineName"))) Then
                                colEngineData.Add(row("FuelType") & "~" & row("EngineName"))
                            Else
                                colEngineDataXCC.Add(row("FuelType") & "~" & row("EngineName"))
                            End If
                            If Not fnbolIsValidRecordset(rst) Then Exit For
                        Next
                    Else
                        If EngineListinExl.Count > 0 Then
                            For i As Integer = 0 To EngineListinExl.Count - 1
                                colEngineData.Add(EngineTypeListinExl.Item(i) & "~" & EngineListinExl.Item(i))
                            Next
                        End If
                    End If
                ElseIf rng.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column Then
                    With Form.DataCenter.GlobalValues.WS
                        'lRow = .Range(.Cells(5, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column), .Cells(.Rows.Count, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column)).End(XlDirection.xlUp).Row
                        lRow = Form.DataCenter.ProgramConfig.LastRow
                    End With
                    For index = 5 To lRow
                        Dim xRng As Excel.Range = CType(Form.DataCenter.GlobalValues.WS.Cells(index, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Column), Excel.Range)
                        Dim xTransmissionTypeRng As Excel.Range = CType(Form.DataCenter.GlobalValues.WS.Cells(index, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Transmission_Type_Column), Excel.Range)
                        Dim val As Object = xRng.Value()
                        If val <> "" And val <> "-" Then
                            If Not TransmissionListinExl.Contains(val) Then
                                TransmissionListinExl.Add(val)
                                TransmissionTypeListinExl.Add(xTransmissionTypeRng.Value())
                            End If
                        End If
                    Next
                    _transmission = New CT.Data.Transmission()
                    rst = _transmission.GetXCCTransmissions(sTnDRegion)
                    If rst IsNot Nothing Then

                        For Each row As DataRow In rst.Rows
                            If TransmissionListinExl.Contains(CStr(row("pe07_TransName"))) Then
                                colTransData.Add(row("pe16_TransType") & "~" & row("pe07_TransName"))
                            Else
                                colTransDataXCC.Add(row("pe16_TransType") & "~" & row("pe07_TransName"))
                            End If
                            If Not fnbolIsValidRecordset(rst) Then Exit For

                        Next
                    Else
                        If TransmissionListinExl.Count > 0 Then
                            For i As Integer = 0 To TransmissionListinExl.Count - 1
                                colTransData.Add(TransmissionTypeListinExl.Item(i) & "~" & TransmissionListinExl.Item(i))
                            Next
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        '        ''' <summary>
        '        ''' Convert ArrayList to List.
        '        ''' </summary>
        ''<System.Runtime.CompilerServices.Extension>
        Public Shared Function ToList(Of T)(arrayList As ArrayList) As List(Of T)
            Dim list As New List(Of T)(arrayList.Count)
            For Each instance As T In arrayList
                list.Add(instance)
            Next
            Return list
        End Function

        Public Function fnbolIsValidRecordset(rst As System.Data.DataTable) As Boolean

            If rst.Rows.Count > 0 Then
                fnbolIsValidRecordset = True
            Else
                fnbolIsValidRecordset = False
            End If
        End Function
        Private Sub WSOps_BeforeDoubleClick(Target As Range, ByRef Cancel As Boolean) Handles WSOps.BeforeDoubleClick
            DisplayPopup(Target)
            Cancel = True
        End Sub
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Public Sub WSOps_ActivateEvent() Handles WSOps.ActivateEvent
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
        End Sub
        Public Sub XLApp_WorkbookActivate(Wb As Workbook) Handles XLApp.WorkbookActivate
            Try
                If Wb.Name Like "TndTemplate*" = False Or Wb.ActiveSheet.name.ToString.ToLower <> Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                    Exit Sub
                End If

                Dim ObjCustTsk As Microsoft.Office.Tools.CustomTaskPane = Nothing

                For intCnt As Integer = Globals.ThisAddIn.CustomTaskPanes.Count To 1 Step -1
                    ObjCustTsk = Globals.ThisAddIn.CustomTaskPanes.Item(intCnt - 1)
                    Dim bolIsObjValid As Boolean = False
                    Try
                        If ObjCustTsk.Title = "You've got a message!" Then
                            bolIsObjValid = True
                        End If
                    Catch ex As Exception
                        bolIsObjValid = False
                    End Try
                    If bolIsObjValid = True Then
                        Globals.ThisAddIn.CustomTaskPanes.RemoveAt(intCnt - 1)
                    End If
                Next

                objCustomTaskPaneMessages = Globals.ThisAddIn.CustomTaskPanes.Add(New MessageTaskPaneControl(), "You've got a message!", Globals.ThisAddIn.Application.ActiveWindow)
                objCustomTaskPaneMessages.Visible = False
                objCustomTaskPaneMessages.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                objCustomTaskPaneMessages.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal

                Globals.Ribbons.RbnTnDControlPanel.TGMessages.Checked = False

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "XLApp_WorkbookActivate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub
    End Class
End Namespace

