
Imports System.Threading
Imports System.Windows.Forms

Namespace Form.DisplayUtilities.PlanSections



    Friend NotInheritable Class SevenTabsFunctions 'Friend = public, NotInheritable = static


        Private Shared _ErrorMessage As String
        Public Shared Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
            Set(ByVal value As String)
                _ErrorMessage = value
            End Set
        End Property

        Private Shared trd As Thread
        Private Shared _GlobalFunctions As New DataCenter.GlobalFunctions
        Private Shared Function datefieldvalidation(colheadertext As String, inputvalue As String, rowno As Integer, colno As Integer, lngMaxLength As Long, Optional bolSkip As Boolean = False) As Boolean
            'Date input/format validation
            datefieldvalidation = False
            'Dim thisDt As DateTime
            inputvalue = Form.DataCenter.GlobalValues.WS.Cells(rowno, colno).Text
            If InStr(colheadertext, "date", CompareMethod.Text) > 0 And InStr(colheadertext, "update", CompareMethod.Text) <= 0 And inputvalue <> "" Then
                '                If IsDate(inputvalue) = False Or DateTime.TryParseExact(inputvalue, "d-M-yyyy",
                'Globalization.CultureInfo.InvariantCulture,
                'System.Globalization.DateTimeStyles.None, thisDt) = False Then

                'Form.DataCenter.GlobalValues.WS.Cells(rowno, colno).NumberFormat = "dd-MM-yyyy"
                'Form.DataCenter.GlobalValues.WS.Cells(rowno, colno).NumberFormat = "@"

                If DateTime.TryParseExact(inputvalue, "d-M-yyyy", Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, Nothing) = False Then
                    'If IsDate(inputvalue) = False Then
                    If inputvalue <> "" Then System.Windows.Forms.MessageBox.Show("Please enter date value in format d-M-yyyy", "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)

                    Form.DataCenter.GlobalValues.WS.Application.EnableEvents = False
                    Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = False
                    Form.DataCenter.GlobalValues.WS.Cells(rowno, colno) = ""
                    'Form.DataCenter.GlobalValues.WS.Cells(rowno, colno).NumberFormat = "@"
                    Form.DataCenter.GlobalValues.WS.Application.EnableEvents = True
                    Form.DataCenter.GlobalValues.WS.Application.ScreenUpdating = True

                    datefieldvalidation = False

                    Exit Function

                    datefieldvalidation = True

                End If
                'End If 'Date input/format validation ends here
            Else

                If bolSkip = False Then
                    If StringLengthValidation(inputvalue, lngMaxLength) Then
                        datefieldvalidation = True
                    Else
                        Form.DataCenter.GlobalValues.WS.Cells(rowno, colno) = ""
                        datefieldvalidation = False
                    End If
                ElseIf _GlobalFunctions.ContainsInvalidChar(inputvalue) = False Then
                    datefieldvalidation = True
                Else
                    datefieldvalidation = False
                    Exit Function
                End If
            End If

            datefieldvalidation = True
        End Function

        Private Shared Function StringLengthValidation(strString As String, lngMaxLength As Long, Optional colName As String = "", Optional bolEnforceExact As Boolean = False)

            'Dim clsGlobalFunc As DataCenter.GlobalFunctions = New DataCenter.GlobalFunctions()

            If _GlobalFunctions.ContainsInvalidChar(strString) = False Then
                If Not bolEnforceExact Then
                    If Len(strString) > lngMaxLength Then
                        System.Windows.Forms.MessageBox.Show(colName & " Text length exceeds max length " & lngMaxLength & ".", "Invalid text", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                        StringLengthValidation = False
                    Else
                        StringLengthValidation = True
                    End If
                Else
                    If Len(strString) <> lngMaxLength Then
                        System.Windows.Forms.MessageBox.Show(colName & " Text length should be exactly " & lngMaxLength & ".", "Invalid text", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                        StringLengthValidation = False
                    Else
                        StringLengthValidation = True
                    End If
                End If

            Else
                System.Windows.Forms.MessageBox.Show("Text contains invalid characters (')", "Invalid text", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                StringLengthValidation = False
            End If

        End Function

        Public Shared Function UpdateData(Target As Excel.Range, WSOps As Microsoft.Office.Tools.Excel.Worksheet) As Boolean

            Dim objCell As Excel.Range = Nothing
            Dim _unit As CT.Data.VehiclePlan.Unit = New CT.Data.VehiclePlan.Unit()
            Dim sPreviousVaue As String = String.Empty
            Dim sPreviousVaue1 As String = String.Empty
            Dim ErrMsg As String = String.Empty
            Dim ErrCnt As Integer = 0

            Try
                Globals.ThisAddIn.Application.EnableEvents = False
                If Target.Column > Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn Then
                    For Each objCell In Target.Cells
                        sPreviousVaue = String.Empty
                        sPreviousVaue1 = String.Empty
                        If WSOps.Cells(objCell.Row, 5).text = "" Then
                            WSOps.Cells(objCell.Row, objCell.Column).Value = ""
                            Continue For
                        End If
                        If objCell.Text <> "" Then
                            If objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column Or objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column Then 'VIN
                                If ChangeLog(objCell, WSOps) = False Then
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column Then 'VIN
                                If objCell.Value.ToString.Length <> 17 Then
                                    sPreviousVaue = _unit.GetPreviousValueVin(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the VIN should be exactly 17 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue = _unit.GetPreviousValueVin(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column Then 'Vehicle Number
                                If StringLengthValidation(objCell.Text, 4, WSOps.Cells(4, objCell.Column).Value, True) = False Then
                                    sPreviousVaue = _unit.GetPreviousValueVehicleNumber(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the Vehicle Number should be exactly 4 characters." & " (Error cell:- " & objCell.Address & ")"
                                ElseIf CStr(WSOps.Cells(objCell.Row, DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).text.ToString.Trim) = String.Empty Then
                                    sPreviousVaue = _unit.GetPreviousValueVehicleNumber(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! please enter the prefix number first and then enter vehicle number."
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue = _unit.GetPreviousValueVehicleNumber(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column Then 'Ship to Customer
                                Try
                                    If datefieldvalidation("date", objCell.Text, objCell.Row, objCell.Column, 0) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueShippingToCustomer(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Data format was not valid"
                                    Else
                                        If ChangeLog(objCell, WSOps) = False Then
                                            sPreviousVaue1 = _unit.GetPreviousValueShippingToCustomer(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                            WSOps.Application.EnableEvents = False
                                            objCell.Value = sPreviousVaue1
                                            objCell.Font.Color = System.Drawing.Color.Blue
                                            WSOps.Application.EnableEvents = True
                                            ErrCnt = ErrCnt + 1
                                            ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                        End If
                                    End If
                                Catch ex As Exception
                                    sPreviousVaue1 = _unit.GetPreviousValueShippingToCustomer(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                End Try
                                '--------------------------------------TBD--------------------------------------------------------------
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column Then
                                Try
                                    If datefieldvalidation("date", objCell.Text, objCell.Row, objCell.Column, 0) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.CustomerRequiredDate.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Data format was not valid"
                                    Else
                                        If ChangeLog(objCell, WSOps) = False Then
                                            sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.CustomerRequiredDate.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                            WSOps.Application.EnableEvents = False
                                            objCell.Value = sPreviousVaue1
                                            objCell.Font.Color = System.Drawing.Color.Blue
                                            WSOps.Application.EnableEvents = True
                                            ErrCnt = ErrCnt + 1
                                            ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                        End If
                                    End If
                                Catch ex As Exception
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.CustomerRequiredDate.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                End Try
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column Then
                                Try
                                    If datefieldvalidation("date", objCell.Text, objCell.Row, objCell.Column, 0) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.RigCustomerPickDate.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Data format was not valid"
                                    Else
                                        If ChangeLog(objCell, WSOps) = False Then
                                            sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.RigCustomerPickDate.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                            WSOps.Application.EnableEvents = False
                                            objCell.Value = sPreviousVaue1
                                            objCell.Font.Color = System.Drawing.Color.Blue
                                            WSOps.Application.EnableEvents = True
                                            ErrCnt = ErrCnt + 1
                                            ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                        End If
                                    End If
                                Catch ex As Exception
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.RigCustomerPickDate.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                End Try
                                '--------------------------------------TBD--------------------------------------------------------------
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column Then 'Specification CBG
                                If StringLengthValidation(objCell.Text, 10) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Cbg.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the CBG should be within 10 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Cbg.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column Then 'dedicated/shared/deleted
                                If StringLengthValidation(objCell.Text, 50) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Dedicated.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the dedicated/Shared/deleted should be within 50 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Dedicated.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column Then 'Emission Stage
                                If StringLengthValidation(objCell.Text, 50) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Emissionstage.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the Emission stage should be within 50 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Emissionstage.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column Then 'Bodystyle
                                If StringLengthValidation(objCell.Text, 50) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Bodystyle.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the Body style should be within 50 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.Bodystyle.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column Then 'Color
                                If StringLengthValidation(objCell.Text, 50) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.ColorCode.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the color should be within 50 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.ColorCode.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column Then 'Driveside
                                If StringLengthValidation(objCell.Text, 50) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.DriveSide.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the drive side should be within 50 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.DriveSide.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column Then 'First User Shipping Address & Contract Name
                                If StringLengthValidation(objCell.Text, 100) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.ShippingAdress.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the remarks should be within 100 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.ShippingAdress.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column Then 'First User Shipping Address & Contract Name
                                If StringLengthValidation(objCell.Text, 4, WSOps.Cells(4, objCell.Column).Value, True) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.TbNumberPrefix.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the prefix should be exactly 4 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.TbNumberPrefix.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column Then 'First User Shipping Address & Contract Name
                                If StringLengthValidation(objCell.Text, 8, Data.DataCenter.ProgramInfoFields.BuildId.ToString, True) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.BuildId.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the Build ID should be exactly 8 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.BuildId.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column Then 'First User Shipping Address & Contract Name
                                If StringLengthValidation(objCell.Text, 7, Data.DataCenter.ProgramInfoFields.TagNumber.ToString, True) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.TagNumber.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the Vehicle tag number should be exactly 7 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.TagNumber.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            ElseIf objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column Then 'First User Shipping Address & Contract Name
                                If StringLengthValidation(objCell.Text, 100) = False Then
                                    sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.PaintFacility.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                    WSOps.Application.EnableEvents = False
                                    objCell.Value = sPreviousVaue1
                                    objCell.Font.Color = System.Drawing.Color.Blue
                                    WSOps.Application.EnableEvents = True
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & "Sorry entry not valid! the paint facility should be within 100 characters." & " (Error cell:- " & objCell.Address & ")"
                                Else
                                    If ChangeLog(objCell, WSOps) = False Then
                                        sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.PaintFacility.ToString, Form.DataCenter.ProgramConfig.BuildType)
                                        WSOps.Application.EnableEvents = False
                                        objCell.Value = sPreviousVaue1
                                        objCell.Font.Color = System.Drawing.Color.Blue
                                        WSOps.Application.EnableEvents = True
                                        ErrCnt = ErrCnt + 1
                                        ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                    End If
                                End If
                            End If
                        ElseIf CStr(WSOps.Cells(objCell.Row, DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).text.ToString.Trim) = String.Empty And
                            CStr(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).text.ToString.Trim) <> String.Empty Then
                            sPreviousVaue1 = _unit.GetPreviousValueGeneral(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row), Form.DataCenter.VehicleConfig.VehiclePe03(objCell.Row), Data.DataCenter.ProgramInfoFields.TbNumberPrefix.ToString, Form.DataCenter.ProgramConfig.BuildType)
                            WSOps.Application.EnableEvents = False
                            objCell.Value = sPreviousVaue1
                            objCell.Font.Color = System.Drawing.Color.Blue
                            WSOps.Application.EnableEvents = True
                            MsgBox("Please clear the vehicle number first and then clear prefix number.", MsgBoxStyle.Exclamation, "Update data")
                        Else
                            If objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column Or
                                    objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column Then
                                If ChangeLog(objCell, WSOps) = False Then
                                    ErrCnt = ErrCnt + 1
                                    ErrMsg = ErrMsg & vbNewLine & ErrCnt & ") " & _ErrorMessage & " (Error cell:- " & objCell.Address & ")"
                                End If
                            End If
                        End If
                        If objCell.Interior.Color = 13487615 Then
                            If objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column Or objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column Or
                                objCell.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column Then
                                WSOps.Range(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Interior.Color = System.Drawing.Color.White
                                WSOps.Range(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline)
                                WSOps.Range(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                WSOps.Range(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
                                WSOps.Range(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                WSOps.Range(WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Phase_Column), WSOps.Cells(objCell.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column)).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
                            End If
                        End If
                    Next
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.InstrumentationSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.InstrumentationSectionLastColumn Then
                    Dim strSection As String
                    Dim strInstrumentation As String
                    Dim strValue As String
                    Dim _list As New List(Of CT.Data.DataCenter.InstrumentationData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strInstrumentation = CType(Form.DataCenter.GlobalValues.WS.Cells(3, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strInstrumentation, "", objCell.Row, objCell.Column, 200) Then
                                strSection = CType(Form.DataCenter.GlobalValues.WS.Cells(2, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                                strValue = objCell.Text
                                _list.Add(New Data.DataCenter.InstrumentationData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row), strSection, strInstrumentation, strValue))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateInstrumentationMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.NonMfSpecificationSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.NonMfSpecificationSectionLastColumn Then
                    Dim strNonMFC As String
                    Dim _list As New List(Of CT.Data.DataCenter.NonMfcSpecificationData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strNonMFC = CType(Form.DataCenter.GlobalValues.WS.Cells(3, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strNonMFC, "", objCell.Row, objCell.Column, 200) Then
                                _list.Add(New CT.Data.DataCenter.NonMfcSpecificationData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                                            Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row),
                                                                            strNonMFC,
                                                                            IIf(objCell.Text = Nothing, "", objCell.Text)))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateNonMFCMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.MfcSpecificationSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.MfcSpecificationSectionLastColumn Then
                    Dim strMFCSection As String
                    Dim strMFC As String
                    Dim strValue As String
                    Dim _list As New List(Of CT.Data.DataCenter.MfcSpecificationData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strMFC = CType(Form.DataCenter.GlobalValues.WS.Cells(3, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strMFC, "", objCell.Row, objCell.Column, 200) Then
                                strMFCSection = CType(Form.DataCenter.GlobalValues.WS.Cells(2, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                                strValue = objCell.Text
                                _list.Add(New CT.Data.DataCenter.MfcSpecificationData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                           Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row),
                                                           strMFC, strMFCSection,
                                                           strValue))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateMFCMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.ProgramInformationSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.ProgramInformationSectionLastColumn Then
                    Dim strProgramInfo As String
                    Dim strValue As String
                    Dim _list As New List(Of CT.Data.DataCenter.ProgramInformationData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strProgramInfo = CType(Form.DataCenter.GlobalValues.WS.Cells(3, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strProgramInfo, "", objCell.Row, objCell.Column, 200) Then
                                strValue = objCell.Text
                                _list.Add(New CT.Data.DataCenter.ProgramInformationData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                       Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row), strProgramInfo, strValue))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateProgramInformationMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.FurtherBasicInformationSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.FurtherBasicInformationSectionLastColumn Then
                    Dim strFurtherBasic As String
                    Dim strValue As String
                    Dim _list As New List(Of CT.Data.DataCenter.FurtherBasicSpecificationData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strFurtherBasic = CType(Form.DataCenter.GlobalValues.WS.Cells(3, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strFurtherBasic, "", objCell.Row, objCell.Column, 200) Then
                                strValue = objCell.Text
                                _list.Add(New CT.Data.DataCenter.FurtherBasicSpecificationData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                  Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row), strFurtherBasic, strValue))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateFurtherBasicSpecificationMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.UserShippingDetailsSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.UserShippingDetailsSectionLastColumn Then
                    Dim strUserShipping As String
                    Dim strValue As String
                    Dim _list As New List(Of CT.Data.DataCenter.UserShippingDetailsData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strUserShipping = CType(Form.DataCenter.GlobalValues.WS.Cells(3, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strUserShipping, "", objCell.Row, objCell.Column, 250) Then
                                strValue = objCell.Text
                                _list.Add(New CT.Data.DataCenter.UserShippingDetailsData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                  Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row), strUserShipping, strValue))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateUserShippingDetailsMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                ElseIf Target.Column > Form.DataCenter.GlobalSections.UpdatePackSectionFirstColumn And Target.Column < Form.DataCenter.GlobalSections.UpdatePackSectionLastColumn Then
                    Dim strUpdatepack As String
                    Dim strValue As String
                    Dim _list As New List(Of CT.Data.DataCenter.UpdatepackData)
                    For Each objCell In Target.Cells
                        If Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row) > 0 And Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row) > 0 Then
                            strUpdatepack = CType(Form.DataCenter.GlobalValues.WS.Cells(2, objCell.Column), Excel.Range).MergeArea.Cells(1, 1).Value
                            If datefieldvalidation(strUpdatepack, "", objCell.Row, objCell.Column, 200) Then
                                strValue = objCell.Text
                                _list.Add(New CT.Data.DataCenter.UpdatepackData(Form.DataCenter.VehicleConfig.VehiclePe02(objCell.Row),
                                                 Form.DataCenter.VehicleConfig.VehiclePe45(objCell.Row), strUpdatepack, strValue))
                            End If
                        End If
                    Next
                    trd = New Thread(AddressOf UpdateUpdatepackMulti)
                    trd.IsBackground = True
                    trd.Start(New Object() {_list})
                    Exit Function
                End If
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Return False
            Finally
                Globals.ThisAddIn.Application.EnableEvents = True
                If ErrMsg <> String.Empty Then
                    System.Windows.Forms.MessageBox.Show(ErrMsg, "Update Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                End If
            End Try
        End Function
        Public Shared Function checkEmptyPassEmpty(Value As Object) As Object
            Try
                If IsDBNull(Value) Or Value = Nothing Then
                    Return String.Empty 'Nothing
                ElseIf Value.ToString = "" Then
                    Return String.Empty
                Else
                    Return Value
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Function ChangeLog(Target As Excel.Range, WSOps As Microsoft.Office.Tools.Excel.Worksheet) As Boolean
            Try
                _ErrorMessage = String.Empty
                If Target.Column >= Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn And Target.Column <= Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn Then
                    With Form.DataCenter.GlobalValues.WS
                        Dim tblProto As System.Data.DataTable = Nothing
                        Dim tblRows() As System.Data.DataRow
                        Dim _DataCls As New Data.VehiclePlan.Plan
                        If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column Or Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column Then
                            tblProto = _DataCls.GetXCCUserTeamNameTranslation(Form.DataCenter.ProgramConfig.BuildType)
                            If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column Then
                                WSOps.Application.EnableEvents = False
                                tblRows = tblProto.Select("XCCPrototypeUser='" & Target.Value2 & "' AND BuildTypes='" & .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 & "'")
                                WSOps.Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column) = tblRows(0)("XCCTranslation")
                                WSOps.Application.EnableEvents = True
                            End If
                            If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column Then
                                WSOps.Application.EnableEvents = False
                                tblRows = tblProto.Select("XCCTranslation='" & Target.Value2 & "' AND BuildTypes='" & .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 & "'")
                                WSOps.Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column) = tblRows(0)("XCCPrototypeUser")
                                WSOps.Application.EnableEvents = True
                            End If
                        End If

                        Dim _UpdateData As Data.Interfaces.UnitInterface

                        If Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Vehicle.ToString Then
                            _UpdateData = New Data.VehiclePlan.Unit
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Rig.ToString Then
                            _UpdateData = New Data.RigPlan.Unit
                        ElseIf Form.DataCenter.ProgramConfig.BuildType = CT.Data.DataCenter.BuildType.Buck.ToString Then
                            _UpdateData = New Data.BuckPlan.Unit
                        Else
                            _UpdateData = Nothing
                        End If

                        Select Case CInt(Target.Column)
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row),
                                                  Form.DataCenter.ProgramConfig.BuildType,
                                                  Form.DataCenter.ProgramConfig.FileStatus,
                                                  Form.DataCenter.ProgramConfig.HCID,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleSpecificationCBG(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column, DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                 DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                 DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType,
                                                 Form.DataCenter.ProgramConfig.FileStatus,
                                                 Form.DataCenter.ProgramConfig.HCID, Nothing,
                                                checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleXCCTeamName(Target.Row)), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                               , checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleTeamName(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row),
                                                   Form.DataCenter.ProgramConfig.BuildType,
                                                   Form.DataCenter.ProgramConfig.FileStatus,
                                                   Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleDedicatedShared(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column '
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row),
                                                  Form.DataCenter.ProgramConfig.BuildType,
                                                  Form.DataCenter.ProgramConfig.FileStatus,
                                                  Form.DataCenter.ProgramConfig.HCID,
                                                  StrTBNumber:=checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleNumber(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleVin(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleEmissionStage(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleBodystyle(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleColor(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleDriveside(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                   checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleRemarks(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleShipToCustomer(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                                '-----------------------------TBD--------------------------------------------
                            Case DataCenter.VehicleProgramInfoColumns.Rig_CustomerRequiredDate_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.Rig_CustomerRequiredDate(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Rig_RigCustomerPickDate_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, Nothing, Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.Rig_RigCustomerPickDate(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                                '-----------------------------TBD--------------------------------------------
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row),
                                                  Form.DataCenter.ProgramConfig.BuildType,
                                                  Form.DataCenter.ProgramConfig.FileStatus,
                                                  Form.DataCenter.ProgramConfig.HCID,
                                                  Nothing,
                                                  Nothing,
                                                  Nothing,
                                                  Nothing,
                                                  Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, Nothing, Nothing,
                                                  checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleNumberPrefix(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, Nothing, Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleBuildId(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, Nothing, Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.VehicleTagNumber(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                            Case DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column
                                If _UpdateData.ChangeInfoII(DataCenter.ProgramConfig.pe02,
                                                  DataCenter.VehicleConfig.VehiclePe45(Target.Row),
                                                  DataCenter.VehicleConfig.VehiclePe03(Target.Row), Form.DataCenter.ProgramConfig.BuildType, Form.DataCenter.ProgramConfig.FileStatus, Form.DataCenter.ProgramConfig.HCID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
                                                  Nothing, Nothing, Nothing, checkEmptyPassEmpty(DataCenter.VehicleConfig.VehiclePaintFacility(Target.Row))) = False Then
                                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                                End If
                        End Select
                    End With
                End If
                Return True
            Catch ex As Exception
                _ErrorMessage = String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.SevenTabsFunctions_ChangeUnitInformation, ex.Message)
                Return False
            End Try
        End Function

        'Private Sub xxChangeLog(Target As Excel.Range, WSOps As Microsoft.Office.Tools.Excel.Worksheet)
        '    Try
        '        If Target.Column >= Form.DataCenter.GlobalSections.VehicleProgramInfoSectionFirstColumn And Target.Column <= Form.DataCenter.GlobalSections.VehicleProgramInfoSectionLastColumn Then
        '            With Form.DataCenter.GlobalValues.WS

        '                Dim tblProto As System.Data.DataTable = Nothing
        '                Dim tblRows() As System.Data.DataRow
        '                Dim _DataCls As New Data.VehiclePlan.Plan

        '                If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column Or Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column Then
        '                    tblProto = _DataCls.GetXCCUserTeamNameTranslation(Form.DataCenter.ProgramConfig.BuildType)
        '                    If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column Then
        '                        WSOps.Application.EnableEvents = False
        '                        tblRows = tblProto.Select("XCCPrototypeUser='" & Target.Value2 & "' AND BuildTypes='" & .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 & "'")
        '                        WSOps.Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column) = tblRows(0)("XCCTranslation")
        '                        WSOps.Application.EnableEvents = True
        '                    End If
        '                    If Target.Column = Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column Then
        '                        WSOps.Application.EnableEvents = False
        '                        tblRows = tblProto.Select("XCCTranslation='" & Target.Value2 & "' AND BuildTypes='" & .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Hardwaretype_Column).value2 & "'")
        '                        WSOps.Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column) = tblRows(0)("XCCPrototypeUser")
        '                        WSOps.Application.EnableEvents = True
        '                    End If

        '                End If
        '                Dim _UpdateData As New Data.VehiclePlan.Unit
        '                'For Each TGT In Target.Rows

        '                If _UpdateData.ChangeInfoII(Form.DataCenter.ProgramConfig.pe02,
        '                                          Form.DataCenter.VehicleConfig.VehiclePe45(Target.Row),
        '                                          Form.DataCenter.VehicleConfig.VehiclePe03(Target.Row),
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column).value2,
        '                                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).value2) = False Then
        '                    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
        '                End If


        '                'If _UpdateData.ChangeInfo (Form.DataCenter.ProgramConfig.pe02,
        '                '                          Form.DataCenter.VehicleConfig.VehiclePe45(Target.Row),
        '                '                          Form.DataCenter.VehicleConfig.VehiclePe03(Target.Row),
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Specification_CBG_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_XCC_Team_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_dedicated_Shared_deleted_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vin_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Emission_Stage_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Bodystyle_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Color_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Driveside_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Team_Names_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Remarks_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Ship_to_Customer_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Vehicle_Number_Prefix_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Build_Id_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Tag_Number_Column).value2,
        '                '                          .Cells(Target.Row, Form.DataCenter.VehicleProgramInfoColumns.Vehicle_Paint_Facility_Column).value2) = False Then
        '                '    Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
        '                'End If
        '                '  Next
        '            End With
        '        End If
        '    Catch ex As Exception
        '        MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.SevenTabsFunctions_ChangeUnitInformation, ex.Message), "Change Unit Information", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        '    End Try
        'End Sub

        Public Shared Sub UpdateInstrumentation(parameters As Object())
            Dim _Instrumentation As New CT.Data.SevenTabsManagement.Instrumentation

            If _Instrumentation.UpdateData(parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub



        Public Shared Sub UpdateInstrumentationMulti(parameters As Object())
            Dim _Instrumentation As New CT.Data.SevenTabsManagement.Instrumentation

            '---------------------------------------------------------------------
            ' Multi Select Copy & Paste
            '---------------------------------------------------------------------
            If _Instrumentation.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub


        Public Shared Sub UpdateNonMFC(parameters As Object())
            Dim _NonMfcSpecification As New CT.Data.SevenTabsManagement.NonMfcSpecification
            If _NonMfcSpecification.UpdateData(parameters(0),
                                        parameters(1),
                                        parameters(2),
                                        parameters(3), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateNonMFCMulti(parameters As Object())
            Dim _NonMfcSpecification As New CT.Data.SevenTabsManagement.NonMfcSpecification
            If _NonMfcSpecification.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub

        Public Shared Sub UpdateMFC(parameters As Object())
            Dim _MfcSpecification As New CT.Data.SevenTabsManagement.MfcSpecification
            If _MfcSpecification.UpdateData(parameters(0),
                                        parameters(1),
                                        parameters(2),
                                        parameters(3),
                                        parameters(4), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateMFCMulti(parameters As Object())
            Dim _MfcSpecification As New CT.Data.SevenTabsManagement.MfcSpecification
            If _MfcSpecification.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub

        Public Shared Sub UpdateProgramInformation(parameters As Object())
            Dim _ProgramInformation As New CT.Data.SevenTabsManagement.ProgramInformation
            If _ProgramInformation.UpdateData(parameters(0),
                                        parameters(1),
                                        parameters(2),
                                        parameters(3), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateProgramInformationMulti(parameters As Object())
            Dim _ProgramInformation As New CT.Data.SevenTabsManagement.ProgramInformation
            If _ProgramInformation.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateFurtherBasicSpecification(parameters As Object())
            Dim _FurtherBasicSpecification As New CT.Data.SevenTabsManagement.FurtherBasicSpecification
            If _FurtherBasicSpecification.UpdateData(parameters(0),
                                        parameters(1),
                                        parameters(2),
                                        parameters(3), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateFurtherBasicSpecificationMulti(parameters As Object())
            Dim _FurtherBasicSpecification As New CT.Data.SevenTabsManagement.FurtherBasicSpecification



            If _FurtherBasicSpecification.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub

        Public Shared Sub UpdateUserShippingDetails(parameters As Object())
            Dim _UserShippingDetails As New CT.Data.SevenTabsManagement.UserShippingDetails
            If _UserShippingDetails.UpdateData(parameters(0),
                                        parameters(1),
                                        parameters(2),
                                        parameters(3), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateUserShippingDetailsMulti(parameters As Object())
            Dim _UserShippingDetails As New CT.Data.SevenTabsManagement.UserShippingDetails
            If _UserShippingDetails.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub

        Public Shared Sub UpdateUpdatepack(parameters As Object())
            Dim _Updatepack As New CT.Data.SevenTabsManagement.Updatepack
            If _Updatepack.UpdateData(parameters(0),
                                        parameters(1),
                                        parameters(2),
                                        parameters(3), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub
        Public Shared Sub UpdateUpdatepackMulti(parameters As Object())
            Dim _Updatepack As New CT.Data.SevenTabsManagement.Updatepack
            If _Updatepack.UpdateData(parameters(0), Form.DataCenter.ProgramConfig.BuildType) = True Then
                Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
                _RibbonUtilitis.UpdateUndoButtonsState()
            Else
                System.Windows.Forms.MessageBox.Show(CT.Data.DataCenter.GlobalValues.message)
            End If

        End Sub

    End Class
End Namespace
