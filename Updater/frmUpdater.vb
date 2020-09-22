Imports System.IO
Imports System.IO.Compression
Public Class frmUpdater

    Private RegistryKeyName As String = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins"
    Private RegistryKeyName2 As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\AddInLoadTimes"
    Private SharepointPath As String
    Private TotalSteps As Integer = 7
    Private DestinationPath As String
    Private TextFilePath As String
    Private ZipFilePath As String
    Private SetupFilePath As String
    Private credentials As System.Net.NetworkCredential = System.Net.CredentialCache.DefaultNetworkCredentials
    Private fileReader As String
    Private PreviousPath As String
    Private _CurrentVersion As String = String.Empty

    Private Sub frmUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ClearForm()
            btnNextUpdate.Text = "Next"
            Dim properties As String() = Command().Split(",")
            _CurrentVersion = properties(0).Split(":")(1).ToString
        Catch
            _CurrentVersion = String.Empty
        End Try
    End Sub

    Private Sub ClearForm()
        Step1_Info.Text = String.Empty
        Step2_Info.Text = String.Empty
        Step3_Info.Text = String.Empty
        Step4_Info.Text = String.Empty
        Step5_Info.Text = String.Empty
        Step6_Info.Text = String.Empty
        Step7_Info.Text = String.Empty
        ProgressBar1.Value = 0
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub


    Private Function GetCurrentVersion() As String
        Dim readValue
        Try
            '-------- Clear the return error message text ----------
            GetCurrentVersion = String.Empty

            '-------- Define destination path global
            DestinationPath = String.Format("C:\Users\{0}\ct-tool", Environment.UserName)

            If _CurrentVersion = String.Empty Then

                '-------- Create destination path if it's not exsiting
                If (Not System.IO.Directory.Exists(DestinationPath)) Then
                    System.IO.Directory.CreateDirectory(DestinationPath)
                End If

                '-------- Check the registry for the current value
                readValue = My.Computer.Registry.GetValue(String.Format("{0}\CT", RegistryKeyName), "Version", Nothing)
                If readValue Is Nothing Then
                    Step1_Info.Text = "1.13.1"
                Else
                    Step1_Info.Text = readValue.ToString
                End If
            Else
                Step1_Info.Text = _CurrentVersion
            End If
            '-------- change the color to green as success and the progress bar
            Step1_CurrentVersion.ForeColor = Color.Green
            ProgressBar1.Value = ProgressBar1.Value + If(Math.Ceiling((100 / TotalSteps)) > 100, 100, Math.Ceiling((100 / TotalSteps)))

        Catch ex As Exception
            Step1_CurrentVersion.ForeColor = Color.Red
            GetCurrentVersion = ex.Message
        End Try

    End Function

    Private Function GetLatestVersion() As String
        Dim source
        Try
            '-------- Clear the return error message text ----------
            GetLatestVersion = String.Empty

            '---------------- Download Version file from sharepoint --------------------------
            source = New Uri("https://pd3.spt.ford.com/sites/PPEteam/SiteCollectionDocuments/PPEteam/Documents/3_PROTO_TEST_PLANNING/9_TnD_Process_Documentation/iDV-ConnectedTesting/Version.txt")
            TextFilePath = DestinationPath + "\Version.txt"
            My.Computer.Network.DownloadFile(source, TextFilePath, credentials, True, 60000I, True)

            '---------------- Read file --------------------------
            fileReader = My.Computer.FileSystem.ReadAllText(TextFilePath)
            Step2_Info.Text = fileReader

            '-------- change the color to green as success and the progress bar
            Step2_LatestVersion.ForeColor = Color.Green
            ProgressBar1.Value = ProgressBar1.Value + If(Math.Ceiling((100 / TotalSteps)) > 100, 100, Math.Ceiling((100 / TotalSteps)))

        Catch ex As Exception
            Step2_LatestVersion.ForeColor = Color.Red
            GetLatestVersion = ex.Message
        End Try
    End Function

    Private Function DownloadLatestVersion() As String
        Dim source
        Try
            '-------- Clear the return error message text ----------
            DownloadLatestVersion = String.Empty
            source = New Uri(String.Format("https://pd3.spt.ford.com/sites/PPEteam/SiteCollectionDocuments/PPEteam/Documents/3_PROTO_TEST_PLANNING/9_TnD_Process_Documentation/iDV-ConnectedTesting/CT_{0}.zip", fileReader.ToString))
            ZipFilePath = String.Format("{0}\CT_{1}.zip", DestinationPath, fileReader.ToString)
            SetupFilePath = String.Format("{0}\CT_{1}\CT.VSTO", DestinationPath, fileReader.ToString)
            My.Computer.Network.DownloadFile(source, ZipFilePath, credentials, True, 60000I, True)
            Step3_Info.Text = "Download completet in " + DestinationPath

            '-------- change the color to green as success and the progress bar
            Step3_DownloadLatestVersion.ForeColor = Color.Green
            ProgressBar1.Value = ProgressBar1.Value + If(Math.Ceiling((100 / TotalSteps)) > 100, 100, Math.Ceiling((100 / TotalSteps)))

        Catch ex As Exception
            Step3_DownloadLatestVersion.ForeColor = Color.Red
            DownloadLatestVersion = ex.Message
        End Try

    End Function

    Private Function ExtractZipFile() As String
        Try
            '-------- Clear the return error message text ----------
            ExtractZipFile = String.Empty

            If Directory.Exists(ZipFilePath.Replace(".zip", "")) = True Then
                Directory.Delete(ZipFilePath.Replace(".zip", ""), True)
            End If


            ZipFile.ExtractToDirectory(ZipFilePath, DestinationPath)

            Step4_Info.Text = "Extracting completet in " + DestinationPath
            btnNextUpdate.Text = "Uninstall"

            '-------- change the color to green as success and the progress bar
            Step4_ExtractingZipFile.ForeColor = Color.Green
            ProgressBar1.Value = ProgressBar1.Value + If(Math.Ceiling((100 / TotalSteps)) > 100, 100, Math.Ceiling((100 / TotalSteps)))

        Catch ex As Exception
            Step4_ExtractingZipFile.ForeColor = Color.Red
            ExtractZipFile = ex.Message
        End Try

    End Function

    Private Function ValidateClosedExcelApplication() As String
        Try
            '-------- Clear the return error message text ----------
            ValidateClosedExcelApplication = String.Empty
            Dim IsExcelApplicationOpen As Boolean = False

            For Each P As Process In Process.GetProcessesByName("Excel")
                IsExcelApplicationOpen = True
            Next

            Do While IsExcelApplicationOpen = True

                If IsExcelApplicationOpen = True Then MessageBox.Show("Close Excel files and continue.", "Update iDV/Connected Testing")

                IsExcelApplicationOpen = False
                For Each P As Process In Process.GetProcessesByName("Excel")
                    IsExcelApplicationOpen = True
                Next

            Loop

            Step5_Info.Text = "No open Excel application."

            '-------- change the color to green as success and the progress bar
            Step5_CloseExcelFiles.ForeColor = Color.Green
            ProgressBar1.Value = ProgressBar1.Value + Math.Ceiling((100 / TotalSteps))

        Catch ex As Exception
            Step5_CloseExcelFiles.ForeColor = Color.Red
            ValidateClosedExcelApplication = ex.Message

        End Try

    End Function


    Private Function UninstallCurrentVersion() As String
        Try
            '-------- Clear the return error message text ----------
            UninstallCurrentVersion = String.Empty

            '-------- Keep previous path for uninstalling
            PreviousPath = My.Computer.Registry.GetValue(String.Format("{0}\CT", RegistryKeyName), "Manifest", String.Empty)
            PreviousPath = PreviousPath.Replace("|vstolocal", "")

            '-------- If previous version exists then uninstall it
            If PreviousPath <> String.Empty Then
                Dim basePath As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles)
                Dim subPath As String = "Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
                Dim vstoInstallerPath As String = Path.Combine(basePath, subPath)
                If (File.Exists(vstoInstallerPath) = False) Then
                    Throw New Exception("The Visual Studio Tools for Office installer was not found.")
                End If
                Dim startInfo As ProcessStartInfo = New ProcessStartInfo(vstoInstallerPath)
                startInfo.Arguments = String.Format("/uninstall {0}  /silent", PreviousPath)
                Dim Process As Process = Process.Start(startInfo)
                Process.WaitForExit()
                Dim ErrorNumber = Process.ExitCode

                If ErrorNumber = 0 Then
                    Step6_Info.Text = "Uninstall done."
                Else
                    Throw New Exception("Uninstall failed.")
                End If
                Process.Close()

                '-------- When uninstall is Ok then delete the previous version.
                'Directory.Delete(PreviousPath.ToUpper().Replace("/CT.VSTO", "").Replace("FILE:///", ""), True)

            Else
                Step6_Info.Text = "Previous version doesn't exist."
            End If

            '-------- change the color to green as success and the progress bar
            Step6_UninstallCurrentOne.ForeColor = Color.Green
            ProgressBar1.Value = ProgressBar1.Value + If(Math.Ceiling((100 / TotalSteps)) > 100, 100, Math.Ceiling((100 / TotalSteps)))

        Catch ex As Exception
            Step6_UninstallCurrentOne.ForeColor = Color.Red
            UninstallCurrentVersion = ex.Message
        End Try
    End Function

    Private Function InstallLatestVersion() As String
        Try
            '-------- Clear the return error message text ----------
            InstallLatestVersion = String.Empty

            Dim basePath As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles)
            Dim subPath As String = "Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
            Dim vstoInstallerPath As String = Path.Combine(basePath, subPath)
            If (File.Exists(vstoInstallerPath) = False) Then
                Throw New Exception("The Visual Studio Tools for Office installer was not found.")
            End If
            Dim startInfo As ProcessStartInfo = New ProcessStartInfo(vstoInstallerPath)
            startInfo.Arguments = String.Format("/install {0} ", SetupFilePath)
            Dim Process As Process = Process.Start(startInfo)
            Process.WaitForExit()
            Dim ErrorNumber = Process.ExitCode
            Process.Close()

            If ErrorNumber = 0 Then
                Step7_Info.Text = "Install done."
            Else
                Step7_Info.Text = "Install failed."
                Throw New Exception("Install failed.")
            End If

            '-------- change the color to green as success and the progress bar
            Step7_InstallTheLatestOne.ForeColor = Color.Green
            ProgressBar1.Value = If(ProgressBar1.Value + Math.Ceiling((100 / TotalSteps)) > 100, 100, ProgressBar1.Value + Math.Ceiling((100 / TotalSteps)))


        Catch ex As Exception
            Step7_InstallTheLatestOne.ForeColor = Color.Red
            InstallLatestVersion = ex.Message
        End Try
    End Function

    Private Sub btnNextUpdate_Click(sender As Object, e As EventArgs) Handles btnNextUpdate.Click
        Dim strErrorMessage As String = String.Empty

        Try

            '------------------------------------------------------
            ' Step 1 current version
            '------------------------------------------------------
            strErrorMessage = GetCurrentVersion()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)


            '------------------------------------------------------
            ' Step 2 new available version
            '------------------------------------------------------
            strErrorMessage = GetLatestVersion()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)

            '------------------------------------------------------
            ' Validation
            '------------------------------------------------------
            If Step1_Info.Text = Step2_Info.Text Then
                btnNextUpdate.Enabled = False
                TxtMessage.ForeColor = Color.Green
                TxtMessage.Visible = True
                TxtMessage.Text = "The current version is the last version."
            Else
                btnNextUpdate.Text = "Download"
            End If

            '------------------------------------------------------
            ' Step 3 download latest version
            '------------------------------------------------------
            strErrorMessage = DownloadLatestVersion()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)

            '------------------------------------------------------
            ' Step 4 extract the files
            '------------------------------------------------------
            strErrorMessage = ExtractZipFile()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)

            '------------------------------------------------------
            ' Step 5 close Excel files
            '------------------------------------------------------
            strErrorMessage = ValidateClosedExcelApplication()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)

            '------------------------------------------------------
            ' Step 6 uninstall the current version with removing from registery
            '------------------------------------------------------
            strErrorMessage = UninstallCurrentVersion()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)

            '------------------------------------------------------
            ' Step 7 install the new version
            '------------------------------------------------------
            strErrorMessage = InstallLatestVersion()
            If strErrorMessage <> String.Empty Then Throw New Exception(strErrorMessage)
            btnNextUpdate.Enabled = False


        Catch ex As Exception
            btnNextUpdate.Enabled = False
            TxtMessage.ForeColor = Color.Red
            TxtMessage.Visible = True
            TxtMessage.Text = ex.Message
        End Try
    End Sub


End Class