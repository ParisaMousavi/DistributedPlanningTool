Imports CT.Data.Reports
Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Text.RegularExpressions


Public Class ThisAddIn
    Private utilities As AddInUtilities
    Dim ribbonObj As Microsoft.Office.Core.IRibbonExtensibility

    Protected Overrides Function RequestComAddInAutomationService() As Object
        If utilities Is Nothing Then
            utilities = New AddInUtilities()
        End If
        Return utilities
    End Function




    'Dim objkk As Microsoft.Office.Tools.Ribbon.IRibbonExtension

    'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
    '    ribbonObj = New Ribbon(Me)
    '    Return ribbonObj
    'End Function

    'Sub activateribbon()
    '    ribbonObj = New RbnTnDControlPanel

    'End Sub


    Private Sub Application_SheetBeforeRightClick(Sh As Object, Target As Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeRightClick
        Dim bolWasSU As Boolean
        Dim bolWasEE As Boolean
        Try
            bolWasSU = Sh.Application.ScreenUpdating
            bolWasEE = Sh.Application.EnableEvents

            Sh.Application.ScreenUpdating = False
            Sh.Application.EnableEvents = False

            If Sh.parent.Name Like "TndTemplate*" = False Or Sh.name.ToString.ToLower <> Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                If CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu IsNot Nothing Then
                    CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu.DeleteContextMenu()
                    CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu = Nothing
                End If
                Cancel = False
                Try
                    Sh.Application.CommandBars("Cell").Reset()
                Catch ex As Exception
                End Try
            End If
            Sh.Application.ScreenUpdating = bolWasSU
            Sh.Application.EnableEvents = bolWasEE
        Catch ex As Exception
            ' System.Windows.Forms.MessageBox.Show(ex.Message, "Right Click event.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            Sh.Application.ScreenUpdating = bolWasSU
            Sh.Application.EnableEvents = bolWasEE
        End Try
    End Sub
    Public Sub Application_WorkbookActivate(Wb As Workbook) Handles Application.WorkbookActivate
        ' System.Threading.Thread.Sleep(5000)
        Dim _obj As New Form.DataCenter.GlobalFunctions
        Wb.Application.ScreenUpdating = False
        Wb.Application.EnableEvents = False

        Try
            If Wb.Name Like "TndTemplate*" Then

                '-------------------------------------------------------------------
                ' This function has beed centralized to make the maintenance easier 
                '-------------------------------------------------------------------

                If Wb.ActiveSheet.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                    Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                    _RibbonUtilities.UpdateRibbonButtonsState()
                Else
                    Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                    _RibbonUtilities.DeactiveRibbonButtonsState()
                End If

                'Update_RibbonButtons()
                Form.DataCenter.GlobalValues.objWBCurrent = Wb

                If Wb.Application.ActiveSheet.Name.ToString.ToLower <> CT.Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                    _obj.sbToggleCutCopyAndPaste(True)
                End If
            Else
                _obj.sbToggleCutCopyAndPaste(True)
                If CT.Form.DataCenter.GlobalValues.wsEve IsNot Nothing Then
                    If CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu IsNot Nothing Then
                        CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu.DeleteContextMenu()
                        CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu = Nothing
                    End If
                End If

                Try
                    Wb.Application.CommandBars("Cell").Reset()
                Catch ex As Exception
                End Try
                'CT.Form.DataCenter.GlobalValues.wsEve = Nothing
                Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilities.DeactiveRibbonButtonsState()
                'Deactive_RibbonButtons()

            End If

            If Wb.Name Like "TndTemplate*" And Wb.ActiveSheet.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then

                'Parisa
                'To implement from outside to TndPlan
                'Wb.Application.CutCopyMode = False


                Form.DataCenter.GlobalValues.bolCopy = False
                Form.DataCenter.GlobalValues.bolCut = False
                'Parisa
                'If Wb.ActiveSheet.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                '    If Wb.Application.Selection.locked = False Then
                '        Wb.Application.DisplayFormulaBar = False
                '    Else
                '        Wb.Application.DisplayFormulaBar = True
                '    End If
                'Else
                '    Wb.Application.DisplayFormulaBar = True
                'End If
            Else
                Try
                    If CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu IsNot Nothing Then
                        CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu.DeleteContextMenu()
                        CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu = Nothing
                    End If
                    Wb.Application.CommandBars("Cell").Reset()
                Catch ex As Exception
                End Try
            End If

        Catch ex As Exception
        Finally
            If Form.DataCenter.ProgramConfig.HCID <> 0 And Wb.Name Like "TndTemplate*" And Wb.ActiveSheet.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
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
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Workbook activate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                Finally
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        objRestrictUser.DisableRibbonButtonsForViewer()
                    Else
                        Dim clsobj As New Form.DataCenter.ModuleFunction
                        clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                    End If

                End Try
            End If
        End Try



        Wb.Application.ScreenUpdating = True
        Wb.Application.EnableEvents = True

    End Sub
    Private Sub ThisAddIn_Startup(sender As Object, e As EventArgs) Handles Me.Startup
        Try
            'Globals.ThisAddIn = Me
            Globals.ThisAddIn.Application.EditDirectlyInCell = True


        Catch ex As Exception
            MsgBox("Sorry! error occured! '" & ex.Message & "' . Please contact CDSID 'AEREN8' for user access and any other information.", MsgBoxStyle.Exclamation, "T&D Plan tool")
            Globals.Ribbons.RbnTnDControlPanel.btnLoadOpenTnDPlan.Enabled = False
        End Try
    End Sub

    Private Sub Application_WorkbookOpen(Wb As Workbook) Handles Application.WorkbookOpen
        Try
            Dim objShellWindows As New SHDocVw.ShellWindows
            Dim win As Object
            Dim marker As Integer = 0
            Dim i As Integer = 0
            Dim arrTnDIEVal, arrHC As String()
            If Wb.Name Like "TnDPlanDemo_Vehicles_CT*" Then
                For Each win In objShellWindows
                    If TypeName(win.Document) = "HTMLDocumentClass" Then
                        If win.Document.Title.ToString().Length() > 30 Then
                            If win.Document.Location.pathname = "/sites/PPEteam/Pages/Default.aspx" Then
                                If win.Document.Title.ToString().Substring(0, 30) = "Open TnD Plans Excel For HCID:" Then
                                    Dim strTnDIEVal = win.Document.Title
                                    arrTnDIEVal = Split(strTnDIEVal, ":")
                                    Dim strHCVal = arrTnDIEVal(1)
                                    arrHC = Split(strHCVal, "_")
                                    Dim HCID = arrHC(0)
                                    Dim strHCName = arrHC(1)
                                    With Globals.Ribbons.RbnTnDControlPanel
                                        .btnLoadOpenTnDPlan.Tag = HCID
                                        .btnLoadOpenTnDPlan_Click(.btnLoadOpenTnDPlan, Nothing)
                                    End With
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
                Wb.Close(SaveChanges:=False)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Application_WindowActivate(Wb As Workbook, Wn As Window) Handles Application.WindowActivate
        If Wb.Name Like "TndTemplate*" And Wb.ActiveSheet.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then

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
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Application window activate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Finally
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        objRestrictUser.DisableRibbonButtonsForViewer()
                    Else
                        Dim clsobj As New Form.DataCenter.ModuleFunction
                        clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                    End If

                End Try
            End If

            Form.DataCenter.GlobalValues.bolCopy = False
            Form.DataCenter.GlobalValues.bolCut = False
        End If
    End Sub
    Private Sub Application_WorkbookDeactivate(Wb As Workbook) Handles Application.WorkbookDeactivate
    End Sub
    Private Sub Application_SheetDeactivate(Sh As Object) Handles Application.SheetDeactivate
    End Sub
    Public Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        Try



            If Sh.parent.Name Like "TndTemplate*" And Sh.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilities.UpdateRibbonButtonsState()

            Else
                Dim _RibbonUtilities As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilities.DeactiveRibbonButtonsState()
            End If



        Catch ex As Exception
        Finally
            If Form.DataCenter.ProgramConfig.HCID <> 0 And Sh.parent.Name Like "TndTemplate*" And Sh.name.ToString.ToLower = Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
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
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Sheet activate", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Finally
                    If Form.DataCenter.GlobalValues.strUserPermissionLevel.ToLower.Replace(" ", "") = CT.Data.DataCenter.UserPermissionLevel.Visitor.ToString.ToLower Or Form.DataCenter.GlobalValues.strUserPermissionLevel.Trim = "" Then
                        objRestrictUser.DisableRibbonButtonsForViewer()
                    Else
                        Dim clsobj As New Form.DataCenter.ModuleFunction
                        clsobj.DisableRibbonButtonsForMaster_Draft_CheckedOut()
                    End If

                End Try
            End If
            If Sh.parent.Name Like "TndTemplate*" = False Or Sh.name.ToString.ToLower <> Form.DataCenter.WorkSheet.TnDPlan.ToString.ToLower Then
                Try
                    If CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu IsNot Nothing Then
                        CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu.DeleteContextMenu()
                        CT.Form.DataCenter.GlobalValues.wsEve._CusContMnu = Nothing
                    End If
                    Sh.parent.Application.CommandBars("Cell").Reset()
                Catch ex As Exception
                End Try
            End If
        End Try
    End Sub

    Private Sub Application_WorkbookBeforeSave(Wb As Workbook, SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        If Wb.Name Like "TndTemplate*" Then
            MsgBox("Sorry! saving this file locally, is not allowed!", MsgBoxStyle.Information, "TnD Plan")
            Cancel = True
        End If
    End Sub

    Private Sub Application_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose

        'menuCheckInOut

        'Globals.Ribbons.RbnTnDControlPanel.loadActiveUsers()
        Try
            Dim _PlanActiveUsers As New Data.PlanActiveUsers
            _PlanActiveUsers.Remove(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "WorkBook before close", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub
End Class
