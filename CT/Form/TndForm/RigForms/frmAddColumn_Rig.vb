Imports System.Windows.Forms

Public Class frmAddColumn_Rig

    Dim strContents(14) As String
    Dim IsChanged As Boolean = False
    Dim myDataTable, myBindTable As System.Data.DataTable
    Dim myView As System.Data.DataView
    Dim strFeature, strDescription As String
    Dim _GlobalFunctions As New Form.DataCenter.GlobalFunctions
    Dim oldDescription As String

    'Event  : Feature group combo box selectedindex change event
    'Purpose: To load sections in combo box based on the selected feature
    'Notes  :  @pe01_TnDBasicProgram_PK and @HealthChartId to be passed dynamically for [GetDynamicHeaders]***  
    Private Sub cboFeatureGroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFeatureGroup.SelectedIndexChanged
        Try
            RemoveHandler cboSections.SelectedIndexChanged, AddressOf cboSections_SelectedIndexChanged

            lstColumns.DataSource = Nothing
            lstColumns.Items.Clear()

            '----------------------------------------
            ' Get dynamic columns of the plan
            '----------------------------------------
            Dim _ModifyColumns As Data.SevenTabsManagement.General = New Data.SevenTabsManagement.General()
            myDataTable = _ModifyColumns.GetDynamicHeaders(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.HCID, Form.DataCenter.ProgramConfig.BuildType)
            '----------------------------------------
            ' Validate DB output
            '----------------------------------------
            If myDataTable Is Nothing Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
            If myDataTable.Rows.Count > 0 Then
                myView = New System.Data.DataView(myDataTable)
                myView.RowFilter = "GroupId=" & cboFeatureGroup.SelectedIndex + 2
                myBindTable = myView.ToTable(True, "Section")
                myView = myBindTable.DefaultView

                AddHandler cboSections.SelectedIndexChanged, AddressOf cboSections_SelectedIndexChanged


                cboSections.DataSource = myView
                cboSections.DisplayMember = "Section"
                cboSections.ValueMember = "Section"

                If cboSections.Items.Count < 2 Then
                    cboSections_SelectedIndexChanged(sender, e)
                End If
            Else
                MessageBox.Show("No data to display.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Section dropdown selection change event
    Private Sub cboSections_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSections.SelectedIndexChanged
        grbHeaderDescription.Visible = False
        lstColumns.DataSource = Nothing
        lstColumns.Items.Clear()

        myView = myDataTable.DefaultView

        If cboSections.Items.Count > 1 Then
            myView.RowFilter = "Section='" & cboSections.Text & "' AND " & "GroupId=" & cboFeatureGroup.SelectedIndex + 2
            myBindTable = myView.ToTable(True, "Header")
            myView = myBindTable.DefaultView
            lstColumns.DataSource = myView
            lstColumns.DisplayMember = "Header"
            lstColumns.ValueMember = "Header"
        Else
            myView.RowFilter = "GroupId=" & cboFeatureGroup.SelectedIndex + 2
            myBindTable = myView.ToTable(True, "Header")
            myView = myBindTable.DefaultView
            lstColumns.DataSource = myView
            lstColumns.DisplayMember = "Header"
            lstColumns.ValueMember = "Header"
        End If
    End Sub

    'Purpose: To close the form
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    'Purpose: To close the popup group box
    'Button 'Cancel' (In Groupbox) click event
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        grbHeaderDescription.Visible = False
        txtHeader.Text = Nothing
        txtDescription.Text = Nothing
        'GroupBox1.Enabled = True
    End Sub

    'Button 'OK' (In Groupbox) click event
    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
        Try
            If txtHeader.Text.Trim() = txtDescription.Text.Trim() Then
                MessageBox.Show("Header text and description cannot be same")
                Exit Sub
            End If
            Application.UseWaitCursor = True
            If cboFeatureGroup.SelectedIndex = 2 Or cboFeatureGroup.SelectedIndex = 6 Then
                If txtHeader.Text = "" Or txtDescription.Text = "" Then Throw New Exception("Header and Description cannot be blank.")
                If _GlobalFunctions.ContainsInvalidChar(txtHeader.Text) Then Throw New Exception("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ; "" ' & ; ~ ` < >.")
                If _GlobalFunctions.ContainsInvalidChar(txtDescription.Text) Then Throw New Exception("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ; "" ' & ; ~ ` < >.")
            End If

            strFeature = _GlobalFunctions.RemoveSPChars(strFeature)

            If cboFeatureGroup.SelectedIndex = 2 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.MfcSpecification = New Data.SevenTabsManagement.MfcSpecification()
                If cmdOk.Tag = "Update" Then
                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, lstColumns.SelectedValue, txtHeader.Text, txtDescription.Text, cboSections.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection)
                    '--------------------------------------------------------------------------------------------------------

                ElseIf cmdOk.Tag = "Add" Then
                    If _UpdateColumns.AddColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, txtHeader.Text, cboSections.Text, txtDescription.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection)

                End If
            ElseIf cboFeatureGroup.SelectedIndex = 6 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.Updatepack = New Data.SevenTabsManagement.Updatepack()
                If cmdOk.Tag = "Update" Then
                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, lstColumns.SelectedValue, txtHeader.Text, txtDescription.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection)
                    '--------------------------------------------------------------------------------------------------------

                ElseIf cmdOk.Tag = "Add" Then
                    If _UpdateColumns.AddColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, txtHeader.Text, txtDescription.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection)

                End If
            End If

            If cmdOk.Tag = "Update" Then
                MessageBox.Show("Header name '" & lstColumns.SelectedValue & "' was updated to '" & txtHeader.Text & "'." & vbNewLine & "Description name '" & oldDescription & "' was updated to '" & txtDescription.Text & "' sucessfully...", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("New column name '" & txtHeader.Text & "' was added sucessfully...", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            ResetForm()

            grbHeaderDescription.Visible = False
            txtHeader.Text = ""
            txtDescription.Text = ""

        Catch ex As Exception
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddColumn, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.UseWaitCursor = False
        End Try
    End Sub

    'Purpose: To delete the column
    'Input  : Selected column(s) in the gridview
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader

        Try
            If MessageBox.Show("Do you really want to delete the selected column(s)?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Exit Sub
            End If
            If lstColumns.SelectedItems.Count <= 0 Then Throw New Exception("Please select a column to delete.")
            If lstColumns.Items.Count = 1 Then Throw New Exception("The last column is not allowed to be deleted. At least one column should exist.")
            If lstColumns.Items.Count = lstColumns.SelectedItems.Count Then Throw New Exception("The last column is not allowed to be deleted. At least one column should exist.")
            Me.Cursor = Cursors.AppStarting
            Dim strSelectedItem As String = ""
            Dim strSelectedItems As String = ""


            Select Case cboFeatureGroup.SelectedIndex
                Case 0

                    Dim _UpdateColumns As Data.SevenTabsManagement.Instrumentation = New Data.SevenTabsManagement.Instrumentation()

                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, InstrumentationList:=strSelectedItem, Section:=cboSections.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection) '1
                    Next
                Case 2

                    Dim _UpdateColumns As Data.SevenTabsManagement.MfcSpecification = New Data.SevenTabsManagement.MfcSpecification()

                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, Mfc:=strSelectedItem, Section:=cboSections.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.MfcSpecificationSection) '2
                    Next
                Case 1

                    Dim _UpdateColumns As Data.SevenTabsManagement.NonMfcSpecification = New Data.SevenTabsManagement.NonMfcSpecification()

                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, NonMfcSpecification:=strSelectedItem, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection) '3
                    Next
                Case 3

                    Dim _UpdateColumns As Data.SevenTabsManagement.ProgramInformation = New Data.SevenTabsManagement.ProgramInformation()
                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, ProgramInformationList:=strSelectedItem, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection) '7
                    Next
                Case 4

                    Dim _UpdateColumns As Data.SevenTabsManagement.FurtherBasicSpecification = New Data.SevenTabsManagement.FurtherBasicSpecification()

                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, FurtherBasicSpecificationList:=strSelectedItem, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection) '4
                    Next
                Case 5

                    Dim _UpdateColumns As Data.SevenTabsManagement.UserShippingDetails = New Data.SevenTabsManagement.UserShippingDetails()
                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, UserShippingDetailsList:=strSelectedItem, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection) '5
                    Next
                Case 6

                    Dim _UpdateColumns As Data.SevenTabsManagement.Updatepack = New Data.SevenTabsManagement.Updatepack()

                    For Each itm As System.Data.DataRowView In lstColumns.SelectedItems
                        strSelectedItem = itm.Row.ItemArray(0).ToString
                        strSelectedItems &= strSelectedItem & ", "
                        '-----------------------------------------------------------------------------------------
                        ' Validation to see if the Stored Procedure has been done or not and if not
                        ' throw a exception
                        '-----------------------------------------------------------------------------------------
                        If _UpdateColumns.Delete(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, UpdatePackList:=strSelectedItem, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                            Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                        End If
                        '--------------------------------------------------------------------------------------------------------
                        ' Update display after deleting column
                        '--------------------------------------------------------------------------------------------------------
                        _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UpdatePackSection) '6

                    Next
            End Select
            If lstColumns.SelectedItems.Count > 1 Then
                strSelectedItems = Mid(strSelectedItems, 1, Len(strSelectedItems) - 2)
                MessageBox.Show("Column name(s) '" & strSelectedItems & "' were deleted sucessfully...", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Column name '" & strSelectedItem & "' was deleted sucessfully...", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            '----------------------------------------------------------------------
            ' It means Excel inteface can be changed after closing the form
            '----------------------------------------------------------------------
            IsChanged = True

        Catch ex As Exception
            IsChanged = False
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddColumn, "Sorry, your changes could not be saved to the database! Database error :-" + ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            '---------------------------------
            ' Activate undo button
            '---------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            ResetForm()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Purpose: Refresh the form
    Private Sub ResetForm()
        Try
            Me.Cursor = Cursors.AppStarting

            Dim strOldValFeature As String
            Dim strOldValSection As String

            strOldValFeature = cboFeatureGroup.Text
            strOldValSection = cboSections.Text

            cboFeatureGroup_SelectedIndexChanged("", e:=Nothing)

            If strOldValSection <> "" Then cboSections_SelectedIndexChanged("", e:=Nothing)

            cboFeatureGroup.Text = strOldValFeature

            If strOldValSection <> "" Then cboSections.Text = strOldValSection
        Catch ex As Exception
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddColumn, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Purpose: To check the column already exist in the listbox
    Private Function IsDuplicateExists(strFeature As String) As Boolean
        Try
            IsDuplicateExists = False
            Dim i As Integer = 0
            While i < lstColumns.Items.Count
                lstColumns.SelectedIndex = i
                If lstColumns.SelectedValue.ToString.Replace(" ", "") = strFeature.Replace(" ", "") Then
                    i = lstColumns.Items.Count
                    IsDuplicateExists = True
                    i = i - 1
                End If
                i += 1
            End While
        Catch ex As Exception
            IsDuplicateExists = False
            MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddColumn, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    'Button 'Add a feature' click event
    'Purpose: To add new column
    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        cmdOk.Tag = "Add"
        strFeature = ""
        strDescription = ""
        IsChanged = False
        Try
            '--------------------------------------------------------
            ' this line means exit sub -> Throw New Exception("000")
            '--------------------------------------------------------
            If MessageBox.Show("Do you really want to add a column?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Throw New Exception("000")
            If cboFeatureGroup.SelectedIndex = 0 And cboSections.Text = "" Then Throw New Exception("Please Select a section.")
            If cboFeatureGroup.SelectedIndex = 2 Then
                txtHeader.MaxLength = 75
                txtDescription.MaxLength = 75
                grbHeaderDescription.Visible = True
                txtHeader.Focus()
                Exit Sub
            ElseIf cboFeatureGroup.SelectedIndex = 6 Then
                txtHeader.MaxLength = 50
                txtDescription.MaxLength = 50
                grbHeaderDescription.Visible = True
                txtHeader.Focus()
                Exit Sub
            Else
                strFeature = InputBox("Please enter the new column name.", Me.Text)

                Dim intColLength As Integer
                If cboFeatureGroup.SelectedIndex = 0 Then 'Instrumentation
                    intColLength = 75
                ElseIf cboFeatureGroup.SelectedIndex = 1 Then ' Non MFC Specification
                    intColLength = 100
                Else 'Program Information, Further Basic Information, User Shipping Details
                    intColLength = 50
                End If

                Do Until strFeature.Count < intColLength + 1
                    MessageBox.Show("Column name cannot exceed " & intColLength & " characters.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    strFeature = InputBox("Please enter the new column name.", Me.Text, strFeature)
                Loop


                If strFeature = "" Or strFeature = "False" Or strFeature = "Falsch" Then Throw New Exception("You have not entered any text. Please enter column name to continue...")
                If _GlobalFunctions.ContainsInvalidChar(strFeature) Then Throw New Exception("Sorry, the following characters are not allowed to be entered in the plan data. Please remove the special characters and try again. The invalid charaters are ; "" ' & ; ~ ` < >.")
            End If

            Me.Cursor = Cursors.AppStarting
            strFeature = _GlobalFunctions.RemoveSPChars(strFeature)
            If IsDuplicateExists(strFeature) Then Throw New Exception("Sorry, the column you are trying to add, already exists.")
            strDescription = _GlobalFunctions.RemoveSPChars(strDescription)
            'sbAddChangeLog strPe61(cboFeatureGroup.ListIndex), , "", "Specification section column " & strFeature & " was added sucessfully...", dblActionID

            If cboFeatureGroup.SelectedIndex = 0 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.Instrumentation = New Data.SevenTabsManagement.Instrumentation()
                If _UpdateColumns.AddColumn(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, InstrumentationList:=strFeature, Section:=cboSections.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                '--------------------------------------------------------------------------------------------------------
                ' Update display after add column
                '--------------------------------------------------------------------------------------------------------
                Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection)

            ElseIf cboFeatureGroup.SelectedIndex = 1 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.NonMfcSpecification = New Data.SevenTabsManagement.NonMfcSpecification()
                If _UpdateColumns.AddColumn(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, NonMfcSpecification:=strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                '--------------------------------------------------------------------------------------------------------
                ' Update display after add column
                '--------------------------------------------------------------------------------------------------------
                Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection)

            ElseIf cboFeatureGroup.SelectedIndex = 3 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.ProgramInformation = New Data.SevenTabsManagement.ProgramInformation()
                If _UpdateColumns.AddColumn(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, ProgramInformationList:=strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                '--------------------------------------------------------------------------------------------------------
                ' Update display after add column
                '--------------------------------------------------------------------------------------------------------
                Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection)

            ElseIf cboFeatureGroup.SelectedIndex = 4 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.FurtherBasicSpecification = New Data.SevenTabsManagement.FurtherBasicSpecification()
                If _UpdateColumns.AddColumn(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, FurtherBasicSpecificationList:=strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                '--------------------------------------------------------------------------------------------------------
                ' Update display after add column
                '--------------------------------------------------------------------------------------------------------
                Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection)

            ElseIf cboFeatureGroup.SelectedIndex = 5 Then
                Dim _UpdateColumns As Data.SevenTabsManagement.UserShippingDetails = New Data.SevenTabsManagement.UserShippingDetails()
                If _UpdateColumns.AddColumn(pe01:=Form.DataCenter.ProgramConfig.pe01, pe02:=Form.DataCenter.ProgramConfig.pe02, HCID:=Form.DataCenter.ProgramConfig.HCID, UserShippingDetailsList:=strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                '--------------------------------------------------------------------------------------------------------
                ' Update display after add column
                '--------------------------------------------------------------------------------------------------------
                Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection)

            End If

            ResetForm()
            IsChanged = True
            MessageBox.Show("Column name '" & strFeature & "' was added sucessfully...", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            IsChanged = False
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddColumn, "Error in adding column : " + ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            '----------------------------------------------
            ' Activate undo button
            '----------------------------------------------
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Button 'Rename' click event
    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        Try
            If lstColumns.SelectedItems.Count > 1 Then
                Throw New Exception("Please select only one column to rename.")
            End If
            cmdOk.Tag = "Update"
            strFeature = ""
            strDescription = ""
            Dim strSelectedItem As String

            If MessageBox.Show("Do you really want to rename this column?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Throw New Exception("000")
            strSelectedItem = lstColumns.SelectedValue
            If strSelectedItem = "" Then Throw New Exception("Please select a column to edit.")

            Application.UseWaitCursor = True

            If cboFeatureGroup.SelectedIndex = 2 Or cboFeatureGroup.SelectedIndex = 6 Then
                If cboFeatureGroup.SelectedIndex = 2 Then
                    txtHeader.MaxLength = 75
                    txtDescription.MaxLength = 75
                Else
                    txtHeader.MaxLength = 50
                    txtDescription.MaxLength = 25
                End If
                txtHeader.Text = strSelectedItem
                myView = Nothing
                myView = myDataTable.DefaultView
                myView.RowFilter = String.Empty
                myView.RowFilter = "Header='" & strSelectedItem & "' AND GroupId=" & cboFeatureGroup.SelectedIndex + 2
                myBindTable = myView.ToTable(True, "Description")
                myView = myBindTable.DefaultView

                If myBindTable.Rows.Count > 0 Then
                    txtDescription.Text = myView.ToTable().Rows(0)(0).ToString()
                    oldDescription = txtDescription.Text
                End If
                grbHeaderDescription.Visible = True
                txtHeader.Focus()
            Else

                strFeature = InputBox("Please enter the new column name.", Me.Text, strSelectedItem)

                Dim intColLength As Integer
                If cboFeatureGroup.SelectedIndex = 0 Then 'Instrumentation
                    intColLength = 75
                ElseIf cboFeatureGroup.SelectedIndex = 1 Then ' Non MFC Specification
                    intColLength = 100
                Else 'Program Information, Further Basic Information, User Shipping Details
                    intColLength = 50
                End If

                Do Until strFeature.Count < intColLength + 1
                    MessageBox.Show("Column name cannot exceed " & intColLength & " characters.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    strFeature = InputBox("Please enter the new column name.", Me.Text, strFeature)
                Loop

                If strFeature = "" Then Throw New Exception("You have not entered any text. Please enter column name to continue...")

                strFeature = _GlobalFunctions.RemoveSPChars(strFeature)
                If IsDuplicateExists(strFeature) Then Throw New Exception("Sorry, the column you are trying to rename already exits. Please select other unique name.")

                If cboFeatureGroup.SelectedIndex = 0 Then
                    Dim _UpdateColumns As Data.SevenTabsManagement.Instrumentation = New Data.SevenTabsManagement.Instrumentation()
                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, strSelectedItem, strFeature, cboSections.Text, cboSections.Text, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.InstrumentationSection)
                    '--------------------------------------------------------------------------------------------------------
                    '    objCon.Execute strProcs(cboFeatureGroup.ListIndex) & ShtProgConfig.Range("B14") & "," & ShtProgConfig.Range("B11") & ",'" & strSelectedItem & "','" & strFeature & "','" & cboSections.Text & "','" & cboSections.Text & "'," & dblActionID
                ElseIf cboFeatureGroup.SelectedIndex = 1 Then
                    Dim _UpdateColumns As Data.SevenTabsManagement.NonMfcSpecification = New Data.SevenTabsManagement.NonMfcSpecification()
                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, strSelectedItem, strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.NonMfcSpecificationSection)
                ElseIf cboFeatureGroup.SelectedIndex = 3 Then
                    Dim _UpdateColumns As Data.SevenTabsManagement.ProgramInformation = New Data.SevenTabsManagement.ProgramInformation()

                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, strSelectedItem, strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.ProgramInformationSection)
                ElseIf cboFeatureGroup.SelectedIndex = 4 Then
                    Dim _UpdateColumns As Data.SevenTabsManagement.FurtherBasicSpecification = New Data.SevenTabsManagement.FurtherBasicSpecification()

                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, strSelectedItem, strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.FurtherBasicInformationSection)
                ElseIf cboFeatureGroup.SelectedIndex = 5 Then
                    Dim _UpdateColumns As Data.SevenTabsManagement.UserShippingDetails = New Data.SevenTabsManagement.UserShippingDetails()
                    If _UpdateColumns.EditColumn(Form.DataCenter.ProgramConfig.pe01, Form.DataCenter.ProgramConfig.pe02, Form.DataCenter.ProgramConfig.HCID, strSelectedItem, strFeature, MainBuildType:=Form.DataCenter.ProgramConfig.BuildType) = False Then
                        Throw New Exception(CT.Data.DataCenter.GlobalValues.message)
                    End If
                    '--------------------------------------------------------------------------------------------------------
                    ' Update display after add column
                    '--------------------------------------------------------------------------------------------------------
                    Dim _DrawTndPlanHeader As New Form.DisplayUtilities.DrawTndPlanHeader
                    _DrawTndPlanHeader.ApplyColorAndMergeToHeaderSection(Form.DataCenter.GlobalSections.SectionName.UserShippingDetailsSection)
                End If

                ResetForm()
                MessageBox.Show("Column name '" & strSelectedItem & "' was updated to '" & strFeature & "' sucessfully...", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            If ex.Message <> "000" Then MessageBox.Show(String.Format("Error {0:d} :  {1}", CT.Form.DataCenter.ErrorCenter.frmAddColumn, ex.Message), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Dim _RibbonUtilitis As New Form.DisplayUtilities.Ribbon.Utilities
            _RibbonUtilitis.UpdateUndoButtonsState()

            Application.UseWaitCursor = False
        End Try
    End Sub

    Private Sub frmAddColumn_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form.DataCenter.ProgramConfig.IsGeneric = True Then
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = False
        End If
    End Sub

    'Form keydown event
    'For shortcut keys
    Private Sub frmAddColumn_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If grbHeaderDescription.Visible = True Then
                grbHeaderDescription.Visible = False
                Exit Sub
            End If
            cmdClose_Click(sender, e)
        ElseIf e.KeyCode = Keys.F4 Then
            cboFeatureGroup.Focus()
        ElseIf e.KeyCode = Keys.F7 Then
            cmdAdd_Click(sender, e)
        ElseIf e.KeyCode = Keys.F8 Then
            cmdUpdate_Click(sender, e)
        ElseIf e.KeyCode = Keys.F9 Then
            cmdDelete_Click(sender, e)
        End If
    End Sub

    Private Sub grbHeaderDescription_Enter(sender As Object, e As EventArgs) Handles grbHeaderDescription.Enter
        grbHeaderDescription.Visible = True
        grbHeaderDescription.BringToFront()
    End Sub

    Private Sub lstColumns_Click(sender As Object, e As EventArgs) Handles lstColumns.Click
        'grbHeaderDescription.Visible = False
        'txtHeader.Text = ""
        'txtDescription.Text = ""
    End Sub

    Private Sub frmAddColumn_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub grbHeaderDescription_Click(sender As Object, e As EventArgs) Handles grbHeaderDescription.Click
        grbHeaderDescription.Visible = True
        grbHeaderDescription.BringToFront()

    End Sub
End Class