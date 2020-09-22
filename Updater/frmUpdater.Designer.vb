<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmUpdater
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Step7_InstallTheLatestOne = New System.Windows.Forms.Label()
        Me.Step6_UninstallCurrentOne = New System.Windows.Forms.Label()
        Me.Step5_CloseExcelFiles = New System.Windows.Forms.Label()
        Me.Step4_ExtractingZipFile = New System.Windows.Forms.Label()
        Me.Step3_DownloadLatestVersion = New System.Windows.Forms.Label()
        Me.Step2_LatestVersion = New System.Windows.Forms.Label()
        Me.Step1_CurrentVersion = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.btnNextUpdate = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TxtMessage = New System.Windows.Forms.TextBox()
        Me.Step7_Info = New System.Windows.Forms.Label()
        Me.Step6_Info = New System.Windows.Forms.Label()
        Me.Step5_Info = New System.Windows.Forms.Label()
        Me.Step4_Info = New System.Windows.Forms.Label()
        Me.Step3_Info = New System.Windows.Forms.Label()
        Me.Step2_Info = New System.Windows.Forms.Label()
        Me.Step1_Info = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(12, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(289, 25)
        Me.Label7.TabIndex = 1
        Me.Label7.Text = "Update iDV/ Connected Testing"
        '
        'Step7_InstallTheLatestOne
        '
        Me.Step7_InstallTheLatestOne.AutoSize = True
        Me.Step7_InstallTheLatestOne.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step7_InstallTheLatestOne.ForeColor = System.Drawing.Color.Black
        Me.Step7_InstallTheLatestOne.Location = New System.Drawing.Point(34, 170)
        Me.Step7_InstallTheLatestOne.Name = "Step7_InstallTheLatestOne"
        Me.Step7_InstallTheLatestOne.Size = New System.Drawing.Size(137, 18)
        Me.Step7_InstallTheLatestOne.TabIndex = 0
        Me.Step7_InstallTheLatestOne.Text = "Install the latest one"
        '
        'Step6_UninstallCurrentOne
        '
        Me.Step6_UninstallCurrentOne.AutoSize = True
        Me.Step6_UninstallCurrentOne.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step6_UninstallCurrentOne.ForeColor = System.Drawing.Color.Black
        Me.Step6_UninstallCurrentOne.Location = New System.Drawing.Point(4, 143)
        Me.Step6_UninstallCurrentOne.Name = "Step6_UninstallCurrentOne"
        Me.Step6_UninstallCurrentOne.Size = New System.Drawing.Size(167, 18)
        Me.Step6_UninstallCurrentOne.TabIndex = 0
        Me.Step6_UninstallCurrentOne.Text = "Uninstall the current one"
        '
        'Step5_CloseExcelFiles
        '
        Me.Step5_CloseExcelFiles.AutoSize = True
        Me.Step5_CloseExcelFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step5_CloseExcelFiles.ForeColor = System.Drawing.Color.Black
        Me.Step5_CloseExcelFiles.Location = New System.Drawing.Point(10, 116)
        Me.Step5_CloseExcelFiles.Name = "Step5_CloseExcelFiles"
        Me.Step5_CloseExcelFiles.Size = New System.Drawing.Size(161, 18)
        Me.Step5_CloseExcelFiles.TabIndex = 0
        Me.Step5_CloseExcelFiles.Text = "Close Excel application"
        '
        'Step4_ExtractingZipFile
        '
        Me.Step4_ExtractingZipFile.AutoSize = True
        Me.Step4_ExtractingZipFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step4_ExtractingZipFile.ForeColor = System.Drawing.Color.Black
        Me.Step4_ExtractingZipFile.Location = New System.Drawing.Point(7, 89)
        Me.Step4_ExtractingZipFile.Name = "Step4_ExtractingZipFile"
        Me.Step4_ExtractingZipFile.Size = New System.Drawing.Size(164, 18)
        Me.Step4_ExtractingZipFile.TabIndex = 0
        Me.Step4_ExtractingZipFile.Text = "Extracting latest version"
        '
        'Step3_DownloadLatestVersion
        '
        Me.Step3_DownloadLatestVersion.AutoSize = True
        Me.Step3_DownloadLatestVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step3_DownloadLatestVersion.ForeColor = System.Drawing.Color.Black
        Me.Step3_DownloadLatestVersion.Location = New System.Drawing.Point(5, 63)
        Me.Step3_DownloadLatestVersion.Name = "Step3_DownloadLatestVersion"
        Me.Step3_DownloadLatestVersion.Size = New System.Drawing.Size(166, 18)
        Me.Step3_DownloadLatestVersion.TabIndex = 0
        Me.Step3_DownloadLatestVersion.Text = "Download latest version"
        '
        'Step2_LatestVersion
        '
        Me.Step2_LatestVersion.AutoSize = True
        Me.Step2_LatestVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step2_LatestVersion.ForeColor = System.Drawing.Color.Black
        Me.Step2_LatestVersion.Location = New System.Drawing.Point(71, 37)
        Me.Step2_LatestVersion.Name = "Step2_LatestVersion"
        Me.Step2_LatestVersion.Size = New System.Drawing.Size(100, 18)
        Me.Step2_LatestVersion.TabIndex = 0
        Me.Step2_LatestVersion.Text = "Latest version"
        '
        'Step1_CurrentVersion
        '
        Me.Step1_CurrentVersion.AutoSize = True
        Me.Step1_CurrentVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step1_CurrentVersion.ForeColor = System.Drawing.Color.Black
        Me.Step1_CurrentVersion.Location = New System.Drawing.Point(34, 11)
        Me.Step1_CurrentVersion.Name = "Step1_CurrentVersion"
        Me.Step1_CurrentVersion.Size = New System.Drawing.Size(137, 18)
        Me.Step1_CurrentVersion.TabIndex = 0
        Me.Step1_CurrentVersion.Text = "The current Version"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.Step7_InstallTheLatestOne)
        Me.Panel1.Controls.Add(Me.Step6_UninstallCurrentOne)
        Me.Panel1.Controls.Add(Me.Step5_CloseExcelFiles)
        Me.Panel1.Controls.Add(Me.Step4_ExtractingZipFile)
        Me.Panel1.Controls.Add(Me.Step3_DownloadLatestVersion)
        Me.Panel1.Controls.Add(Me.Step2_LatestVersion)
        Me.Panel1.Controls.Add(Me.Step1_CurrentVersion)
        Me.Panel1.Location = New System.Drawing.Point(17, 61)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(175, 337)
        Me.Panel1.TabIndex = 2
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 270)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(378, 23)
        Me.ProgressBar1.TabIndex = 4
        '
        'btnNextUpdate
        '
        Me.btnNextUpdate.Location = New System.Drawing.Point(228, 16)
        Me.btnNextUpdate.Name = "btnNextUpdate"
        Me.btnNextUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnNextUpdate.TabIndex = 4
        Me.btnNextUpdate.Text = "Ok"
        Me.btnNextUpdate.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(309, 16)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnNextUpdate)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Location = New System.Drawing.Point(199, 350)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(390, 48)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TxtMessage)
        Me.GroupBox2.Controls.Add(Me.Step7_Info)
        Me.GroupBox2.Controls.Add(Me.Step6_Info)
        Me.GroupBox2.Controls.Add(Me.Step5_Info)
        Me.GroupBox2.Controls.Add(Me.Step4_Info)
        Me.GroupBox2.Controls.Add(Me.Step3_Info)
        Me.GroupBox2.Controls.Add(Me.Step2_Info)
        Me.GroupBox2.Controls.Add(Me.Step1_Info)
        Me.GroupBox2.Controls.Add(Me.ProgressBar1)
        Me.GroupBox2.Location = New System.Drawing.Point(199, 61)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(390, 299)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        '
        'TxtMessage
        '
        Me.TxtMessage.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TxtMessage.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtMessage.Location = New System.Drawing.Point(6, 202)
        Me.TxtMessage.Multiline = True
        Me.TxtMessage.Name = "TxtMessage"
        Me.TxtMessage.Size = New System.Drawing.Size(377, 62)
        Me.TxtMessage.TabIndex = 12
        Me.TxtMessage.Visible = False
        '
        'Step7_Info
        '
        Me.Step7_Info.AutoSize = True
        Me.Step7_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step7_Info.ForeColor = System.Drawing.Color.Black
        Me.Step7_Info.Location = New System.Drawing.Point(6, 171)
        Me.Step7_Info.Name = "Step7_Info"
        Me.Step7_Info.Size = New System.Drawing.Size(23, 17)
        Me.Step7_Info.TabIndex = 5
        Me.Step7_Info.Text = "---"
        '
        'Step6_Info
        '
        Me.Step6_Info.AutoSize = True
        Me.Step6_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step6_Info.ForeColor = System.Drawing.Color.Black
        Me.Step6_Info.Location = New System.Drawing.Point(6, 144)
        Me.Step6_Info.Name = "Step6_Info"
        Me.Step6_Info.Size = New System.Drawing.Size(23, 17)
        Me.Step6_Info.TabIndex = 6
        Me.Step6_Info.Text = "---"
        '
        'Step5_Info
        '
        Me.Step5_Info.AutoSize = True
        Me.Step5_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step5_Info.ForeColor = System.Drawing.Color.Black
        Me.Step5_Info.Location = New System.Drawing.Point(6, 117)
        Me.Step5_Info.Name = "Step5_Info"
        Me.Step5_Info.Size = New System.Drawing.Size(23, 17)
        Me.Step5_Info.TabIndex = 7
        Me.Step5_Info.Text = "---"
        '
        'Step4_Info
        '
        Me.Step4_Info.AutoSize = True
        Me.Step4_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step4_Info.ForeColor = System.Drawing.Color.Black
        Me.Step4_Info.Location = New System.Drawing.Point(6, 90)
        Me.Step4_Info.Name = "Step4_Info"
        Me.Step4_Info.Size = New System.Drawing.Size(23, 17)
        Me.Step4_Info.TabIndex = 8
        Me.Step4_Info.Text = "---"
        '
        'Step3_Info
        '
        Me.Step3_Info.AutoSize = True
        Me.Step3_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step3_Info.ForeColor = System.Drawing.Color.Black
        Me.Step3_Info.Location = New System.Drawing.Point(6, 64)
        Me.Step3_Info.Name = "Step3_Info"
        Me.Step3_Info.Size = New System.Drawing.Size(23, 17)
        Me.Step3_Info.TabIndex = 9
        Me.Step3_Info.Text = "---"
        '
        'Step2_Info
        '
        Me.Step2_Info.AutoSize = True
        Me.Step2_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step2_Info.ForeColor = System.Drawing.Color.Black
        Me.Step2_Info.Location = New System.Drawing.Point(6, 38)
        Me.Step2_Info.Name = "Step2_Info"
        Me.Step2_Info.Size = New System.Drawing.Size(48, 17)
        Me.Step2_Info.TabIndex = 10
        Me.Step2_Info.Text = "1.14.1"
        '
        'Step1_Info
        '
        Me.Step1_Info.AutoSize = True
        Me.Step1_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Step1_Info.ForeColor = System.Drawing.Color.Black
        Me.Step1_Info.Location = New System.Drawing.Point(6, 12)
        Me.Step1_Info.Name = "Step1_Info"
        Me.Step1_Info.Size = New System.Drawing.Size(48, 17)
        Me.Step1_Info.TabIndex = 11
        Me.Step1_Info.Text = "1.13.1"
        '
        'frmUpdater
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 405)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label7)
        Me.Name = "frmUpdater"
        Me.Text = "iDV/Connected Testing updater"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label7 As Label
    Friend WithEvents Step7_InstallTheLatestOne As Label
    Friend WithEvents Step6_UninstallCurrentOne As Label
    Friend WithEvents Step5_CloseExcelFiles As Label
    Friend WithEvents Step4_ExtractingZipFile As Label
    Friend WithEvents Step3_DownloadLatestVersion As Label
    Friend WithEvents Step2_LatestVersion As Label
    Friend WithEvents Step1_CurrentVersion As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents btnNextUpdate As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Step7_Info As Label
    Friend WithEvents Step6_Info As Label
    Friend WithEvents Step5_Info As Label
    Friend WithEvents Step4_Info As Label
    Friend WithEvents Step3_Info As Label
    Friend WithEvents Step2_Info As Label
    Friend WithEvents Step1_Info As Label
    Friend WithEvents TxtMessage As TextBox
End Class
