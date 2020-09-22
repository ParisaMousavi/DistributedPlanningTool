Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports System.Windows.Forms

Namespace Form.TndContextMenu
    Public Class CustomContextMenu
        Private commandBar As Office.CommandBar

        Dim ContextMenu As Office.CommandBar
        WithEvents EditusercaseMacro, Moveleft, SelectAll, Selectusercase, Delete, Copy, Cut, NewMacro, Moveright, CopyText, EditMacro, Insert, Insertbefore, Insertafter, Paste, Paste_, Copy_, Delete_, MoveLeftUnits, MoveRightUnits As Office.CommandBarButton
        Dim _modfun As New Form.DataCenter.ModuleFunction()
        Dim scntMnuAdd As String
        Dim icntMnuVal As String
        Public Sub DeleteContextMenu()
            Try
                For Each btn As Office.CommandBarControl In Globals.ThisAddIn.Application.CommandBars("Cell").Controls
                    btn.Delete()
                Next
            Catch ex As Exception

            End Try
            '   Globals.ThisAddIn.Application.CommandBars("Cell").Delete()
        End Sub
        Public Sub AddToCellMenu(Val As Int32, sAddress As String, Optional bolDisableEdit As Boolean = False)
            scntMnuAdd = sAddress
            icntMnuVal = Val

            DeleteContextMenu()
            ContextMenu = Globals.ThisAddIn.Application.CommandBars("Cell")

            If Val = 1 Then
                'Add Custom control to context Menu.
                Moveleft = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With Moveleft
                    .FaceId = 41
                    .Caption = "Move left"
                    .Enabled = Not Form.DataCenter.GlobalValues.bolSelAll
                    .Enabled = Not bolDisableEdit
                End With
                Moveright = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=2)
                With Moveright
                    .FaceId = 39
                    .Caption = "Move right"
                    .Enabled = Not bolDisableEdit
                End With
                NewMacro = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=3)
                With NewMacro
                    .FaceId = 598
                    .Caption = "New"
                    .Enabled = Not bolDisableEdit
                End With
                EditMacro = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=4)
                With EditMacro
                    .FaceId = 162
                    .Caption = "Edit"
                    .Enabled = Not bolDisableEdit
                End With
                EditusercaseMacro = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=5)
                With EditusercaseMacro
                    .FaceId = 592
                    .Caption = "Edit user case"
                    .Enabled = Not bolDisableEdit
                End With
                Cut = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=6)
                With Cut
                    .FaceId = 21
                    .Caption = "Cut"
                    .Enabled = Not bolDisableEdit
                End With
                Copy = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=7)
                With Copy
                    .FaceId = 19
                    .Caption = "Copy"
                    .Enabled = Not bolDisableEdit
                End With
                CopyText = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=8)
                With CopyText
                    .FaceId = 248
                    .Caption = "Copy text"
                End With
                Delete = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=9)
                With Delete
                    .FaceId = 214
                    .Caption = "Delete"
                    .Enabled = Not bolDisableEdit
                End With
                Selectusercase = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=10)
                With Selectusercase
                    .FaceId = 1196
                    .Caption = "Select user case"
                End With
                SelectAll = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=11)
                With SelectAll
                    .FaceId = 1197
                    .Caption = "Select all"
                End With
            End If
            If Val = 2 Then
                Moveleft = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With Moveleft
                    .FaceId = 41
                    .Caption = "Move left"
                    .Enabled = Not bolDisableEdit
                End With
                Moveright = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=2)
                With Moveright
                    .FaceId = 39
                    .Caption = "Move right"
                    .Enabled = Not bolDisableEdit
                End With
                NewMacro = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=3)
                With NewMacro
                    .FaceId = 598
                    .Caption = "New"
                    .Enabled = Not bolDisableEdit
                End With

            End If
            If Val = 3 Then
                Insert = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With Insert
                    .FaceId = 297
                    .Caption = "Insert"
                    .Enabled = Not bolDisableEdit
                End With
            End If
            If Val = 4 Then
                Moveleft = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With Moveleft
                    .FaceId = 41
                    .Caption = "Move left"
                    .Enabled = Not bolDisableEdit
                End With
                Moveright = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=2)
                With Moveright
                    .FaceId = 39
                    .Caption = "Move right"
                    .Enabled = Not bolDisableEdit
                End With
            End If
            If Val = 5 Then
                NewMacro = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With NewMacro
                    .FaceId = 598
                    .Caption = "New"
                    .Enabled = Not bolDisableEdit
                End With
            End If
            If Val = 6 Then
                Insertbefore = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With Insertbefore
                    .FaceId = 154
                    .Caption = "Insert before"
                    .Enabled = Not bolDisableEdit
                End With
                Insertafter = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=2)
                With Insertafter
                    .FaceId = 157
                    .Caption = "Insert after"
                    .Enabled = Not bolDisableEdit
                End With
            End If

            If Val = 7 Then
                If Globals.ThisAddIn.Application.CutCopyMode = Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy Or Globals.ThisAddIn.Application.CutCopyMode = Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCut Then
                    ' If Form.DataCenter.GlobalValues.bolCopy Or Form.DataCenter.GlobalValues.bolCut Then
                    Paste_ = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                    With Paste_
                        .FaceId = 3624
                        .Caption = "Paste "
                        .Enabled = Not bolDisableEdit
                    End With
                Else
                    Copy_ = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                    With Copy_
                        .FaceId = 19
                        .Caption = "Copy "
                        .Enabled = Not bolDisableEdit
                    End With
                    Delete_ = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=2)
                    With Delete_
                        .FaceId = 358
                        .Caption = "Delete "
                        .Enabled = Not bolDisableEdit
                    End With
                End If
            End If

            If Val = 8 Then
                MoveLeftUnits = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=1)
                With MoveLeftUnits
                    .FaceId = 41
                    .Caption = "Move units left"
                    .Enabled = Not bolDisableEdit
                End With
                MoveRightUnits = ContextMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Before:=2)
                With MoveRightUnits
                    .FaceId = 39
                    .Caption = "Move units right"
                    .Enabled = Not bolDisableEdit
                End With
            End If

            Try

                If Form.DataCenter.ProgramConfig.IsGeneric = False Then ContextMenu.ShowPopup()
            Catch ex As Exception
            End Try

        End Sub
        Private Sub ButtonClick(ByVal ctrl As Office.CommandBarButton, ByRef Cancel As Boolean) Handles Moveleft.Click, SelectAll.Click, Selectusercase.Click, Delete.Click, Copy.Click, Cut.Click, EditusercaseMacro.Click, EditMacro.Click, NewMacro.Click, Moveright.Click, CopyText.Click, Insert.Click, Insertbefore.Click, Insertafter.Click, Paste.Click, MoveRightUnits.Click, MoveLeftUnits.Click, Paste_.Click, Copy_.Click, Delete_.Click

            '  Dim _globalvalue As 
            'Dim _unit As CT.Data.Unit = New CT.Data.Unit()

            Dim rng As Excel.Range, rngrow As Excel.Range
            Dim strBuildtype As String

            Select Case ctrl.Caption
                Case "Edit user case"

                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------

                    Form.TndContextMenu.EditUsercaseButton.click()

                Case "Edit"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------

                    Form.TndContextMenu.EditProcessStepButton.Click()


                Case "Move left"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    'If MessageBox.Show("Do you want to move left?", "Move left", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    Form.TndContextMenu.MoveLeftButton.Click()
                    CancelSel()
                    ' _modfun.MoveLeft(scntMnuAdd) ' No


                Case "Move right"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    'If MessageBox.Show("Do you want to move right?", "Move right", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    Form.TndContextMenu.MoveRightButton.Click()
                    '_modfun.MoveRight(scntMnuAdd) ' No
                    CancelSel()

                Case "New"
                    '_modfun.cNew(scntMnuAdd)
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    Form.TndContextMenu.NewButton.ClicK(scntMnuAdd)


                Case "Cut"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    Form.TndContextMenu.CutButton.Click(scntMnuAdd)
                    '_modfun.Cut(scntMnuAdd)


                Case "Copy"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    Form.TndContextMenu.CopyButton.Click(scntMnuAdd)
                    '_modfun.Copy(scntMnuAdd)

                Case "Copy text"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    Form.TndContextMenu.CopyTextButton.Click(scntMnuAdd)
                    '_modfun.CopyText(scntMnuAdd)

                Case "Delete"
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    'Globals.ThisAddIn.Application.ActiveCell.Formula.Split(";")(2)
                    '  _modfun.Delete(scntMnuAdd) ' No
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    If MessageBox.Show("Do you want to delete?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    Form.TndContextMenu.DeleteButton.Click()

                Case "Select user case"
                    '_modfun.UserCaseSelect(scntMnuAdd)
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    Form.TndContextMenu.SelectUsercaseButton.Click(scntMnuAdd)

                Case "Select all"
                    '------------------------------------------------------
                    ' For clean code make a module for each button and all of then 
                    ' with click sub
                    '------------------------------------------------------
                    Form.TndContextMenu.SelectAllButton.Click(scntMnuAdd)

                Case "Insert"
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    If Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$").Length <= 3 Then
                        strBuildtype = Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$")(2)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value
                    Else
                        strBuildtype = Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$")(4)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value
                    End If
                    If Form.DataCenter.VehicleConfig.VehicleBuildType <> strBuildtype Then
                        MessageBox.Show("Sorry, Source and destination Build type is not matching.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    Form.TndContextMenu.CopyButton.Insert()
                    CancelSel()
                Case "Insert before"
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    If Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$").Length <= 3 Then
                        strBuildtype = Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$")(2)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value
                    Else
                        strBuildtype = Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$")(4)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value
                    End If
                    If Form.DataCenter.VehicleConfig.VehicleBuildType <> strBuildtype Then
                        MessageBox.Show("Sorry, Source and destination Build type is not matching.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    Form.TndContextMenu.CopyButton.InsertBefore()
                    CancelSel()
                Case "Insert after"
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    If Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$").Length <= 3 Then
                        strBuildtype = Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$")(2)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value
                    Else
                        strBuildtype = Form.DataCenter.GlobalValues.WS.Cells(CInt(Form.DataCenter.GlobalValues.strCopyAddress.ToString.Split("$")(4)), Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).Value
                    End If
                    If Form.DataCenter.VehicleConfig.VehicleBuildType <> strBuildtype Then
                        MessageBox.Show("Sorry, Source and destination Build type is not matching.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    Form.TndContextMenu.CopyButton.InsertAfter()
                    CancelSel()
                Case "Paste "
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If

                    rng = Form.DataCenter.GlobalValues.WS.Application.Selection
                    For Each rngrow In rng.Rows
                        If rngrow.EntireRow.Hidden = True Then
                            System.Windows.Forms.MessageBox.Show("Sorry this operation is Not allowed when try to paste in hided rows.", "Paste Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    Next
                    Try
                        Dim rngCopy As Excel.Range
                        rngCopy = Form.DataCenter.GlobalValues.WS.Range(Form.DataCenter.GlobalValues.strCopyAddress)
                        rngCopy.Copy()
                    Catch ex As Exception
                    End Try
                    rng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues)
                    Form.DataCenter.GlobalValues.WS.Application.CutCopyMode = False
                    Form.DataCenter.GlobalValues.bolCopy = False
                    Form.DataCenter.GlobalValues.bolCut = False
                    Form.DataCenter.GlobalValues.strCopyAddress = ""
                Case "Copy "
                    rng = Form.DataCenter.GlobalValues.WS.Application.Selection
                    For Each rngrow In rng.Rows
                        If rngrow.EntireRow.Hidden = True Then
                            System.Windows.Forms.MessageBox.Show("Sorry this operation is Not allowed when try to copy from hided rows.", "Copy Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    Next
                    rng.Copy()
                    Form.DataCenter.GlobalValues.bolCopy = True
                    Form.DataCenter.GlobalValues.strCopyAddress = rng.Address
                Case "Delete "
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    If MessageBox.Show("Do you want to clear data?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    rng = Form.DataCenter.GlobalValues.WS.Application.Selection
                    For Each rngrow In rng.Rows
                        If rngrow.EntireRow.Hidden = True Then
                            System.Windows.Forms.MessageBox.Show("Sorry this operation is Not allowed when try to delete from hided rows.", "Delete Function", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    Next
                    Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                    _RibbonUtilitis.UpdateUndoButtonsState()
                    rng.Clear()
                    CancelSel()
                Case "Move units left"
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    'If MessageBox.Show("Do you want to move left?", "Move left", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    Form.TndContextMenu.MoveLeftButton.Click_Multi()
                    CancelSel()
                    Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                    _RibbonUtilitis.UpdateUndoButtonsState()
                Case "Move units right"
                    If Form.DataCenter.ProgramConfig.IsGeneric = True Then
                        MessageBox.Show("Sorry, this operation is not allowed in 'Generic' plan.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                    'If MessageBox.Show("Do you want to move right?", "Move right", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    Form.TndContextMenu.MoveRightButton.Click_Multi()
                    CancelSel()
                    Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                    _RibbonUtilitis.UpdateUndoButtonsState()
            End Select

            ' Globals.ThisAddIn.Application.ActiveWindow.SetFocus()
        End Sub

        Public Sub CancelSel()
            Try
                Try
                    Form.DataCenter.GlobalValues.WS.Unprotect(Form.DataCenter.GlobalValues.ConstPwd)
                Catch ex As Exception
                End Try
                Globals.ThisAddIn.Application.CutCopyMode = False
                Form.DataCenter.GlobalValues.bolCutCopyMode = False
                Form.DataCenter.GlobalValues.strCutAddress = ""
                Form.DataCenter.GlobalValues.strCopyAddress = ""
                Form.DataCenter.GlobalValues.bolUserCaseSelected = False
                Form.DataCenter.GlobalValues.strUserCaseSelected = ""
                Form.DataCenter.GlobalValues.bolSelAll = False
                Form.DataCenter.GlobalValues.strSelAllAddress = ""
                Form.DataCenter.GlobalValues.bolCopy = False
                Form.DataCenter.GlobalValues.bolCut = False
                'If Globals.ThisAddIn.Application.Selection.column >= Form.DataCenter.GlobalSections.TimeLineSectionFirstColumn Then
                '    Form.DataCenter.GlobalValues.WS.Cells(3, Globals.ThisAddIn.Application.Selection.column).select
                'Else
                '    Form.DataCenter.GlobalValues.WS.Cells(4, Globals.ThisAddIn.Application.Selection.column).select
                'End If
                'Form.DataCenter.WS.Cells.FormatConditions.Delete()
                Dim _obj As New Form.DataCenter.ModuleFunction
                _obj.sbProtectPlan()
            Catch ex As Exception

            End Try
        End Sub
    End Class
End Namespace

