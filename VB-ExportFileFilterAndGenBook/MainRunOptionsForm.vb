Imports System.Configuration
Imports System.IO

Public Class MainRunOptionsForm
    Private Sub CmdBEFSToBookDefaultInput_Click(sender As Object, e As EventArgs) Handles CmdBEFSToBookDefaultInput.Click
        DriveDefaultBlogExportFileSearchToBook()
    End Sub

    Private Sub CmdBEFSToBookPromptInput_Click(sender As Object, e As EventArgs) Handles CmdBEFSToBookPromptInput.Click
        DrivePromptInputBlogExportFileSearchToBook()
    End Sub

    Private Sub CmdFilterBEFDefaultInput_Click(sender As Object, e As EventArgs) Handles CmdFilterBEFDefaultInput.Click
        DriveDefaultFilterBlogExportFile()
    End Sub

    Private Sub CmdSBBGCLsDefaultInput_Click(sender As Object, e As EventArgs) Handles CmdSBBGCLsDefaultInput.Click
        DriveDefaultSplitBBAndGenCLs()
    End Sub

    Private Sub CmdSBBGCLsPromptInput_Click(sender As Object, e As EventArgs) Handles CmdSBBGCLsPromptInput.Click
        DrivePromptInputSplitBBAndGenCLs()
    End Sub

    Private Sub CmdExit_Click(sender As Object, e As EventArgs) Handles CmdExit.Click
        Me.Close()
    End Sub

    Private Sub CmdTestGCL_Click(sender As Object, e As EventArgs) Handles CmdTestGCL.Click
        TestDriveGenerateContentsLinks()
    End Sub

    Private Sub CmdTestSnGCL_Click(sender As Object, e As EventArgs) Handles CmdTestSnGCL.Click
        DriveTestSplitBBAndGenCLs()
    End Sub

    Private Sub CmdFilterBEFPromptInput_Click(sender As Object, e As EventArgs) Handles CmdFilterBEFPromptInput.Click
        DrivePromptFilterBlogExportFile()
    End Sub

    Private Sub CmdFilterByIndexDefaultInput_Click(sender As Object, e As EventArgs) Handles CmdFilterByIndexDefaultInput.Click
        DriveDefaultExportFileFilterByIndexList()
    End Sub

    Private Sub CmdFilterByIndexPromptInput_Click(sender As Object, e As EventArgs) Handles CmdFilterByIndexPromptInput.Click
        DrivePromptExportFileFilterByIndexList()
    End Sub

    Private Sub CmdShowAppConfig_Click(sender As Object, e As EventArgs) Handles CmdShowAppConfig.Click
        Dim ExeDir = Application.StartupPath
        Dim ExePath = System.Environment.ProcessPath
        Dim ExeFilenameWoutExtn = Path.GetFileNameWithoutExtension(ExePath)
        Dim AppConfigFilename = ExeFilenameWoutExtn & ".dll.config"
        'Dim msg = "Application.StartupPath: " & Application.StartupPath & vbCrLf &
        '        "System.Environment.ProcessPath: " & System.Environment.ProcessPath & vbCrLf &
        Dim msg = "App Config file of this program is in same directory as this program executable" &
            " which is " & Application.StartupPath & " ." & vbCrLf &
            "The App Config filename is: " & AppConfigFilename & " ." & vbCrLf & vbCrLf &
            "Note that above App Config file's appSettings can be modified to change behaviour " &
            "of this program related to that app setting. For example, DefaultInputDirectory " &
            "and DefaultXMLInputFileName can be changed and when the appropriate " &
            """Run With Default Input"" button is pressed in the main form of the program, the program" &
            " will look for the default file in the path specified by the changed app settings." &
            vbCrLf & vbCrLf &
            "After OK is pressed, Notepad.exe will be executed to open App Config file in a separate " &
            "process using Process.Start(""notepad.exe"", AppConfigFilename). However, changes made to " &
            "the file in this Notepad window usually cannot be saved as the file usually can be " &
            "modified only by administrators." & vbCrLf & vbCrLf &
            "To modify the App Config file, first ensure you have a backup working copy of the App " &
            "Config file. Then " &
            "run Notepad as administrator, open the App Config file in that Notepad window, make " &
            "changes and save. Close and run this program again for the new App Config file to take " &
            "effect." & vbCrLf & vbCrLf &
            "After OK is pressed, this process will return to the invoker of this function."
        If MsgBox(msg, MsgBoxStyle.OkCancel, MsgBoxTitle) = DialogResult.OK Then
            Process.Start("notepad.exe", AppConfigFilename)
        End If
    End Sub

    Private Sub CmdViewUserGuide_Click(sender As Object, e As EventArgs) Handles CmdViewUserGuide.Click
        'Dim DefaultBrowser As String = "chrome.exe"
        'Dim UserGuideURL As String =
        '    "https://ravisiyermisc.blogspot.com/2023/09/short-user-guide-to-creating-blogger.html"
        Dim DefaultInternetBrowserFilename = ConfigurationManager.AppSettings("DefaultInternetBrowserFilename")
        Dim UserGuideURL = ConfigurationManager.AppSettings("UserGuideURL")

        Dim msg = "User Guide URL: " & UserGuideURL & " ." & vbCrLf &
            "Default Internet Browser Filename: " & DefaultInternetBrowserFilename & vbCrLf & vbCrLf &
            "Note that above Default Internet Browser Filename can be changed in App Config file's " &
            "app settings key DefaultInternetBrowserFilename. " &
            "For example, it could be changed (from chrome) to msedge or iexplore ." & vbCrLf & vbCrLf &
            "After OK is pressed, above User Guide link will be opened in the abovementioned default " &
            "Internet browser filename program in a separate process using Process.Start() ." &
            vbCrLf & vbCrLf &
            "Then this process will return to the invoker of this function."
        If MsgBox(msg, MsgBoxStyle.OkCancel, MsgBoxTitle) = DialogResult.OK Then
            'Process.Start(DefaultBrowserExe, UserGuideURL)
            ' https://learn.microsoft.com/en-us/answers/questions/1168207/process-start-does-not-work-anymore
            Dim Pr As Process = New Process()
            Pr.StartInfo.UseShellExecute = True
            ' Pr.StartInfo.FileName = "chrome" ' works
            ' Pr.StartInfo.Arguments = "www.google.com" ' works
            Pr.StartInfo.FileName = DefaultInternetBrowserFilename
            Pr.StartInfo.Arguments = UserGuideURL
            'Handle start failure through try catch
            Try
                Pr.Start()
            Catch ex As Exception
                MsgBox("Failed to open user guide in browser. Details: " & ex.Message, , MsgBoxTitle)
            End Try
        End If

    End Sub

    Private Sub MainRunOptionsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DefaultCommandButtonsEnabled = ConfigurationManager.AppSettings("DefaultCommandButtonsEnabled")
        Dim TestButtonsVisible = ConfigurationManager.AppSettings("TestButtonsVisible")
        If DefaultCommandButtonsEnabled = "1" Then
            CmdBEFSToBookDefaultInput.Enabled = True
            CmdFilterBEFDefaultInput.Enabled = True
            CmdSBBGCLsDefaultInput.Enabled = True
            CmdFilterByIndexDefaultInput.Enabled = True
            CmdHowToEnableDefaultBtns.Enabled = False
        Else
            CmdBEFSToBookDefaultInput.Enabled = False
            CmdFilterBEFDefaultInput.Enabled = False
            CmdSBBGCLsDefaultInput.Enabled = False
            CmdFilterByIndexDefaultInput.Enabled = False
            CmdHowToEnableDefaultBtns.Enabled = True
        End If
        If TestButtonsVisible = "1" Then
            CmdTestGCL.Visible = True
            CmdTestSnGCL.Visible = True
            CmdTestGCL.Enabled = True
            CmdTestSnGCL.Enabled = True
        Else
            CmdTestGCL.Visible = False
            CmdTestSnGCL.Visible = False
            CmdTestGCL.Enabled = False
            CmdTestSnGCL.Enabled = False
        End If
    End Sub

    Private Sub CmdHowToEnableDefaultBtns_Click(sender As Object, e As EventArgs) Handles CmdHowToEnableDefaultBtns.Click
        Dim msg As String = "Procedure to enable ""Run with Default Input"" command buttons:" & vbCrLf & vbCrLf &
        "1. Ensure that App Config file app settings key DefaultCommandButtonsEnabled has value ""1"". " &
        "To see how to view and modify App Config file, click on ""View App Config file"" button " &
        "in main form." & vbCrLf & vbCrLf &
        "2. Ensure that App Config file app settings key ""DefaultInputDirectory"" has value as directory in " &
        "which user has permission to create files." & vbCrLf & vbCrLf &
        "3. Ensure that App Config file app settings keys: DefaultXMLInputFileName, DefaultHTMLInputFileName " &
        "and DefaultIndexListFileName have values that point to appropriate files in DefaultInputDirectory " &
        "and for which (files) user has read permission."
        MsgBox(msg, , MsgBoxTitle)
    End Sub
End Class
