Imports System.Windows.Forms
Imports System.Data


Public Class frmPlanValidation

    Private PrerequisitesFulfilled As Boolean = False
    'Public frmOwner As frmHCIDSelect = Nothing


    Private Sub frmPlanValidation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Try

        '    Cursor = Cursors.WaitCursor
        '    dgvEngine.AutoGenerateColumns = False
        '    dgvTransmission.AutoGenerateColumns = False
        '    dgvPrototypeUser.AutoGenerateColumns = False
        '    dgvXCCRemovedEngine.AutoGenerateColumns = False


        '    Dim _Plan As Data.Plan = New Data.Plan()
        '    PrerequisitesFulfilled = True ' The assumption is that, the  PrerequisitesFulfilled = True at the beginning.

        '    '-------------------------------------------------------
        '    ' This validation is for all the buildPhase
        '    '-------------------------------------------------------
        '    If _Plan.Validaion_1(Form.DataCenter.ProgramConfig.XccPe01.ToString, Form.DataCenter.ProgramConfig.HCID.ToString,
        '                                        Form.DataCenter.ProgramConfig.BuildPhase.ToString,
        '                                        Form.DataCenter.ProgramConfig.BuildType.ToString) = True Then ' 0, 3, type 13, phase 5
        '        lblBasicProgramInformation.Text = "Ok"
        '        lblBasicProgramInformation.ForeColor = Drawing.Color.ForestGreen

        '    Else
        '        lblBasicProgramInformation.Text = CT.Data.DataCenter.GlobalValues.message
        '        lblBasicProgramInformation.ForeColor = Drawing.Color.IndianRed

        '        '---------------------------------------------------------
        '        ' To identify if the prerequisites are fulfilled 
        '        '---------------------------------------------------------
        '        PrerequisitesFulfilled = False

        '    End If

        '    '-------------------------------------------------------
        '    ' This validation is for all the buildPhase
        '    '-------------------------------------------------------
        '    If _Plan.Validaion_2(Form.DataCenter.ProgramConfig.XccPe01.ToString, Form.DataCenter.ProgramConfig.XccPe26.ToString, Form.DataCenter.ProgramConfig.HCID.ToString,
        '                                        Form.DataCenter.ProgramConfig.BuildPhase.ToString,
        '                                        Form.DataCenter.ProgramConfig.BuildType.ToString) = True Then ' 0, 3, type 13, phase 5
        '        lblProgramDetails.Text = "Ok"
        '        lblProgramDetails.ForeColor = Drawing.Color.ForestGreen
        '    Else
        '        lblProgramDetails.Text = CT.Data.DataCenter.GlobalValues.message
        '        lblProgramDetails.ForeColor = Drawing.Color.IndianRed

        '        '---------------------------------------------------------
        '        ' To identify if the prerequisites are fulfilled 
        '        '---------------------------------------------------------
        '        PrerequisitesFulfilled = False

        '    End If


        '    Dim ds As New DataSet
        '    ds = _Plan.Validaion_5(Form.DataCenter.ProgramConfig.HCID.ToString, Form.DataCenter.ProgramConfig.BuildType.ToString)
        '    If IsNothing(ds) = False Then

        '        lblXCCEngineSummary.Text = "Ok"
        '        lblXCCEngineSummary.ForeColor = Drawing.Color.ForestGreen

        '        lblXCCTransmissionSummary.Text = "Ok"
        '        lblXCCTransmissionSummary.ForeColor = Drawing.Color.ForestGreen

        '        lblXCCPrototypeUserSummary.Text = "Ok"
        '        lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.ForestGreen

        '        lblXCCRemovedHardware.Text = "-"
        '        lblXCCRemovedHardware.ForeColor = Drawing.Color.Black



        '        Select Case Form.DataCenter.ProgramConfig.BuildPhase
        '            Case CT.Data.DataCenter.BuildPhase.DCV.ToString, CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString, CT.Data.DataCenter.BuildPhase.VP.ToString

        '                If ds.Tables.Count > 0 Then dgvPrototypeUser.DataSource = ds.Tables(0)
        '                If ds.Tables.Count > 1 Then dgvEngine.DataSource = ds.Tables(1)
        '                If ds.Tables.Count > 2 Then dgvTransmission.DataSource = ds.Tables(2)

        '                '--------- Engine validation --------------------------------------
        '                For Each row In dgvEngine.Rows
        '                    If row.cells("TotalQuantityValidation").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("QuantityValidation").Value.ToString.Substring(0, 3) = "NOK" Then
        '                        lblXCCEngineSummary.Text = "The count mismatch"
        '                        lblXCCEngineSummary.ForeColor = Drawing.Color.IndianRed
        '                        PrerequisitesFulfilled = False
        '                        Exit For
        '                    End If
        '                Next
        '                If dgvEngine.Rows.Count = 0 Then
        '                    lblXCCEngineSummary.Text = "The count mismatch"
        '                    lblXCCEngineSummary.ForeColor = Drawing.Color.IndianRed
        '                    PrerequisitesFulfilled = False
        '                End If
        '                '--------- Transmission validation --------------------------------------
        '                For Each row In dgvTransmission.Rows
        '                    If row.cells("TotalQuantityValidation_Transmission").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("QuantityValidation_Transmission").Value.ToString.Substring(0, 3) = "NOK" Then
        '                        lblXCCTransmissionSummary.Text = "The count mismatch"
        '                        lblXCCTransmissionSummary.ForeColor = Drawing.Color.IndianRed
        '                        PrerequisitesFulfilled = False
        '                        Exit For
        '                    End If
        '                Next
        '                If dgvTransmission.Rows.Count = 0 Then
        '                    lblXCCTransmissionSummary.Text = "The count mismatch"
        '                    lblXCCTransmissionSummary.ForeColor = Drawing.Color.IndianRed
        '                    PrerequisitesFulfilled = False
        '                End If
        '                '--------- Prototypeuser validation --------------------------------------
        '                For Each row In dgvPrototypeUser.Rows
        '                    If row.cells("TotalQuantityValidation_PrototypeuserSummary").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("TotalQuantityValidation_Prototypeuser").Value.ToString.Substring(0, 3) = "NOK" Then
        '                        lblXCCPrototypeUserSummary.Text = "The count mismatch"
        '                        lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
        '                        PrerequisitesFulfilled = False
        '                        Exit For
        '                    End If
        '                Next
        '                If dgvPrototypeUser.Rows.Count = 0 Then
        '                    lblXCCPrototypeUserSummary.Text = "The count mismatch"
        '                    lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
        '                    PrerequisitesFulfilled = False
        '                End If
        '                '--------- Removed engines validation --------------------------------------
        '                Dim dt As New DataTable
        '                dt = _Plan.Validaion_6(Form.DataCenter.ProgramConfig.HCID.ToString, Form.DataCenter.ProgramConfig.BuildType.ToString)
        '                If dt IsNot Nothing Then
        '                    If dt.Rows.Count > 0 Then
        '                        dgvXCCRemovedEngine.DataSource = dt

        '                        For Each row In dgvXCCRemovedEngine.Rows
        '                            If row.cells("Result").Value = "Deleted Transmission but PowerPack assignment" Or row.cells("Result").Value = "Deleted Engine but PowerPack assignment" Then
        '                                lblXCCRemovedHardware.Text = "The Hardware mismatch"
        '                                lblXCCRemovedHardware.ForeColor = Drawing.Color.IndianRed
        '                                PrerequisitesFulfilled = False
        '                                Exit For
        '                            End If
        '                        Next
        '                    End If
        '                End If

        '            Case CT.Data.DataCenter.BuildPhase.TT.ToString, CT.Data.DataCenter.BuildPhase.PP.ToString
        '                If ds.Tables.Count > 0 Then dgvPrototypeUser.DataSource = ds.Tables(0)

        '                tabEngine.Enabled  = False
        '                tabTransmission.Enabled = False
        '                tabXCCRemovedEngine.Enabled = False

        '                lblXCCEngineSummary.Text = "-"
        '                lblXCCEngineSummary.ForeColor = Drawing.Color.Black

        '                lblXCCTransmissionSummary.Text = "-"
        '                lblXCCTransmissionSummary.ForeColor = Drawing.Color.Black


        '                '--------- Prototypeuser validation --------------------------------------
        '                For Each row In dgvPrototypeUser.Rows
        '                    If row.cells("TotalQuantityValidation_PrototypeuserSummary").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("TotalQuantityValidation_Prototypeuser").Value.ToString.Substring(0, 3) = "NOK" Then
        '                        lblXCCPrototypeUserSummary.Text = "The count mismatch"
        '                        lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
        '                        PrerequisitesFulfilled = False
        '                        Exit For
        '                    End If
        '                Next
        '                If dgvPrototypeUser.Rows.Count = 0 Then
        '                    lblXCCPrototypeUserSummary.Text = "The count mismatch"
        '                    lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
        '                    PrerequisitesFulfilled = False
        '                End If

        '            Case Else
        '                ' We don't have logic for this case
        '                ' X0, X1, XM
        '                lblXCCEngineSummary.Text = "-"
        '                lblXCCEngineSummary.ForeColor = Drawing.Color.Black

        '                lblXCCTransmissionSummary.Text = "-"
        '                lblXCCTransmissionSummary.ForeColor = Drawing.Color.Black

        '                lblXCCPrototypeUserSummary.Text = "-"
        '                lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.Black

        '                lblXCCRemovedHardware.Text = "-"
        '                lblXCCRemovedHardware.ForeColor = Drawing.Color.Black

        '        End Select



        '    End If

        '    '---------------------------------------------------------
        '    ' Transfer the prerequisit state to the owner.
        '    '---------------------------------------------------------
        '    frmOwner.PrerequisitesFulfilled = PrerequisitesFulfilled
        'Catch ex As Exception

        'Finally
        '    Cursor = Cursors.Default
        'End Try
    End Sub

    Private Sub frmPlanValidation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

 

    Private Sub frmPlanValidation_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Try

            Cursor = Cursors.WaitCursor
            dgvEngine.AutoGenerateColumns = False
            dgvTransmission.AutoGenerateColumns = False
            dgvPrototypeUser.AutoGenerateColumns = False
            dgvXCCRemovedEngine.AutoGenerateColumns = False


            Dim _Plan As Data.VehiclePlan.Plan = New Data.VehiclePlan.Plan()
            PrerequisitesFulfilled = True ' The assumption is that, the  PrerequisitesFulfilled = True at the beginning.

            '-------------------------------------------------------
            ' This validation is for all the buildPhase
            '-------------------------------------------------------
            If _Plan.Validaion_1(Form.DataCenter.ProgramConfig.XccPe01.ToString, Form.DataCenter.ProgramConfig.HCID.ToString,
                                                Form.DataCenter.ProgramConfig.BuildPhase.ToString,
                                                Form.DataCenter.ProgramConfig.BuildType.ToString) = True Then ' 0, 3, type 13, phase 5
                lblBasicProgramInformation.Text = "Ok"
                lblBasicProgramInformation.ForeColor = Drawing.Color.ForestGreen

            Else
                lblBasicProgramInformation.Text = CT.Data.DataCenter.GlobalValues.message
                lblBasicProgramInformation.ForeColor = Drawing.Color.IndianRed

                '---------------------------------------------------------
                ' To identify if the prerequisites are fulfilled 
                '---------------------------------------------------------
                PrerequisitesFulfilled = False

            End If

            '-------------------------------------------------------
            ' This validation is for all the buildPhase
            '-------------------------------------------------------
            If _Plan.Validaion_2(Form.DataCenter.ProgramConfig.XccPe01.ToString, Form.DataCenter.ProgramConfig.XccPe26.ToString, Form.DataCenter.ProgramConfig.HCID.ToString,
                                                Form.DataCenter.ProgramConfig.BuildPhase.ToString,
                                                Form.DataCenter.ProgramConfig.BuildType.ToString) = True Then ' 0, 3, type 13, phase 5
                lblProgramDetails.Text = "Ok"
                lblProgramDetails.ForeColor = Drawing.Color.ForestGreen
            Else
                lblProgramDetails.Text = CT.Data.DataCenter.GlobalValues.message
                lblProgramDetails.ForeColor = Drawing.Color.IndianRed

                '---------------------------------------------------------
                ' To identify if the prerequisites are fulfilled 
                '---------------------------------------------------------
                PrerequisitesFulfilled = False

            End If


            Dim ds As New DataSet
            ds = _Plan.Validaion_5(Form.DataCenter.ProgramConfig.HCID.ToString, Form.DataCenter.ProgramConfig.BuildType.ToString)
            If IsNothing(ds) = False Then

                lblXCCEngineSummary.Text = "Ok"
                lblXCCEngineSummary.ForeColor = Drawing.Color.ForestGreen

                lblXCCTransmissionSummary.Text = "Ok"
                lblXCCTransmissionSummary.ForeColor = Drawing.Color.ForestGreen

                lblXCCPrototypeUserSummary.Text = "Ok"
                lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.ForestGreen

                lblXCCRemovedHardware.Text = "-"
                lblXCCRemovedHardware.ForeColor = Drawing.Color.Black



                Select Case Form.DataCenter.ProgramConfig.BuildPhase
                    Case CT.Data.DataCenter.BuildPhase.DCV.ToString, CT.Data.DataCenter.BuildPhase.M1.ToString, CT.Data.DataCenter.BuildPhase.TPV.ToString, CT.Data.DataCenter.BuildPhase.VP.ToString

                        If ds.Tables.Count > 0 Then dgvPrototypeUser.DataSource = ds.Tables(0)
                        If ds.Tables.Count > 1 Then dgvEngine.DataSource = ds.Tables(1)
                        If ds.Tables.Count > 2 Then dgvTransmission.DataSource = ds.Tables(2)

                        '--------- Engine validation --------------------------------------
                        For Each row In dgvEngine.Rows
                            If row.cells("TotalQuantityValidation").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("QuantityValidation").Value.ToString.Substring(0, 3) = "NOK" Then
                                lblXCCEngineSummary.Text = "The count mismatch"
                                lblXCCEngineSummary.ForeColor = Drawing.Color.IndianRed
                                PrerequisitesFulfilled = False
                                Exit For
                            End If
                        Next
                        If dgvEngine.Rows.Count = 0 Then
                            lblXCCEngineSummary.Text = "The count mismatch"
                            lblXCCEngineSummary.ForeColor = Drawing.Color.IndianRed
                            PrerequisitesFulfilled = False
                        End If
                        '--------- Transmission validation --------------------------------------
                        For Each row In dgvTransmission.Rows
                            If row.cells("TotalQuantityValidation_Transmission").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("QuantityValidation_Transmission").Value.ToString.Substring(0, 3) = "NOK" Then
                                lblXCCTransmissionSummary.Text = "The count mismatch"
                                lblXCCTransmissionSummary.ForeColor = Drawing.Color.IndianRed
                                PrerequisitesFulfilled = False
                                Exit For
                            End If
                        Next
                        If dgvTransmission.Rows.Count = 0 Then
                            lblXCCTransmissionSummary.Text = "The count mismatch"
                            lblXCCTransmissionSummary.ForeColor = Drawing.Color.IndianRed
                            PrerequisitesFulfilled = False
                        End If
                        '--------- Prototypeuser validation --------------------------------------
                        For Each row In dgvPrototypeUser.Rows
                            If row.cells("TotalQuantityValidation_PrototypeuserSummary").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("TotalQuantityValidation_Prototypeuser").Value.ToString.Substring(0, 3) = "NOK" Then
                                lblXCCPrototypeUserSummary.Text = "The count mismatch"
                                lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
                                PrerequisitesFulfilled = False
                                Exit For
                            End If
                        Next
                        If dgvPrototypeUser.Rows.Count = 0 Then
                            lblXCCPrototypeUserSummary.Text = "The count mismatch"
                            lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
                            PrerequisitesFulfilled = False
                        End If
                        '--------- Removed engines validation --------------------------------------
                        Dim dt As New DataTable
                        dt = _Plan.Validaion_6(Form.DataCenter.ProgramConfig.HCID.ToString, Form.DataCenter.ProgramConfig.BuildType.ToString)
                        If dt IsNot Nothing Then
                            If dt.Rows.Count > 0 Then
                                dgvXCCRemovedEngine.DataSource = dt

                                For Each row In dgvXCCRemovedEngine.Rows
                                    If row.cells("Result").Value = "Deleted Transmission but PowerPack assignment" Or row.cells("Result").Value = "Deleted Engine but PowerPack assignment" Then
                                        lblXCCRemovedHardware.Text = "The Hardware mismatch"
                                        lblXCCRemovedHardware.ForeColor = Drawing.Color.IndianRed
                                        PrerequisitesFulfilled = False
                                        Exit For
                                    End If
                                Next
                            End If
                        End If

                    Case CT.Data.DataCenter.BuildPhase.TT.ToString, CT.Data.DataCenter.BuildPhase.PP.ToString
                        If ds.Tables.Count > 0 Then dgvPrototypeUser.DataSource = ds.Tables(0)

                        tabEngine.Enabled = False
                        tabTransmission.Enabled = False
                        tabXCCRemovedEngine.Enabled = False

                        lblXCCEngineSummary.Text = "-"
                        lblXCCEngineSummary.ForeColor = Drawing.Color.Black

                        lblXCCTransmissionSummary.Text = "-"
                        lblXCCTransmissionSummary.ForeColor = Drawing.Color.Black


                        '--------- Prototypeuser validation --------------------------------------
                        For Each row In dgvPrototypeUser.Rows
                            If row.cells("TotalQuantityValidation_PrototypeuserSummary").Value.ToString.Substring(0, 3) = "NOK" Or row.cells("TotalQuantityValidation_Prototypeuser").Value.ToString.Substring(0, 3) = "NOK" Then
                                lblXCCPrototypeUserSummary.Text = "The count mismatch"
                                lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
                                PrerequisitesFulfilled = False
                                Exit For
                            End If
                        Next
                        If dgvPrototypeUser.Rows.Count = 0 Then
                            lblXCCPrototypeUserSummary.Text = "The count mismatch"
                            lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.IndianRed
                            PrerequisitesFulfilled = False
                        End If

                    Case Else
                        ' We don't have logic for this case
                        ' X0, X1, XM
                        lblXCCEngineSummary.Text = "-"
                        lblXCCEngineSummary.ForeColor = Drawing.Color.Black

                        lblXCCTransmissionSummary.Text = "-"
                        lblXCCTransmissionSummary.ForeColor = Drawing.Color.Black

                        lblXCCPrototypeUserSummary.Text = "-"
                        lblXCCPrototypeUserSummary.ForeColor = Drawing.Color.Black

                        lblXCCRemovedHardware.Text = "-"
                        lblXCCRemovedHardware.ForeColor = Drawing.Color.Black

                End Select



            End If

            '---------------------------------------------------------
            ' Transfer the prerequisit state to the owner.
            '---------------------------------------------------------
            frmOwner.PrerequisitesFulfilled = PrerequisitesFulfilled
        Catch ex As Exception

        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

    End Sub
End Class