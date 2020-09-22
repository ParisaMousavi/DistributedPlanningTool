Imports System.Windows.Forms

Public Class frmExporttoexcel

    'Cancel button click event
    'To close the form
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    'Export button click event
    'To Export sheets based on selection of checkbox
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If chkChangelogs.Checked = False And chkDvpteam.Checked = False And chkTndplan.Checked = False Then
            MessageBox.Show("Select any one report sheet to export.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim _exportexcel As Form.Reports.ExporToExcelReport = New Form.Reports.ExporToExcelReport(CT.Form.DataCenter.ProgramConfig.pe01, CT.Form.DataCenter.ProgramConfig.HCID)
        _exportexcel.exporttoexcel(chkTndplan.Checked, chkChangelogs.Checked, chkDvpteam.Checked)

        'If System.Windows.Forms.Application.OpenForms.OfType(Of frmExporttoexcel).Any() Then
        '    System.Windows.Forms.Application.OpenForms.OfType(Of frmExporttoexcel).First.Close()
        'End If
        Me.Close()
    End Sub

    'Checkboxes click event - validation to enable export button
    Private Sub chkTndplan_CheckedChanged(sender As Object, e As EventArgs) Handles chkTndplan.CheckedChanged
        If chkChangelogs.Checked = False And chkDvpteam.Checked = False And chkTndplan.Checked = False Then
            btnExport.Enabled = False
        Else
            btnExport.Enabled = True
        End If
    End Sub

    'Checkboxes click event - validation to enable export button
    Private Sub chkChangelogs_CheckedChanged(sender As Object, e As EventArgs) Handles chkChangelogs.CheckedChanged
        If chkChangelogs.Checked = False And chkDvpteam.Checked = False And chkTndplan.Checked = False Then
            btnExport.Enabled = False
        Else
            btnExport.Enabled = True
        End If
    End Sub

    'Checkboxes click event - validation to enable export button
    Private Sub chkDvpteam_CheckedChanged(sender As Object, e As EventArgs) Handles chkDvpteam.CheckedChanged
        If chkChangelogs.Checked = False And chkDvpteam.Checked = False And chkTndplan.Checked = False Then
            btnExport.Enabled = False
        Else
            btnExport.Enabled = True
        End If
    End Sub
End Class