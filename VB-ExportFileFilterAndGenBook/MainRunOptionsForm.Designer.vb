<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainRunOptionsForm
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
        Label1 = New Label()
        Panel1 = New Panel()
        CmdBEFSToBookPromptInput = New Button()
        CmdBEFSToBookDefaultInput = New Button()
        Panel2 = New Panel()
        CmdFilterBEFPromptInput = New Button()
        CmdFilterBEFDefaultInput = New Button()
        Label2 = New Label()
        Panel3 = New Panel()
        CmdSBBGCLsPromptInput = New Button()
        CmdSBBGCLsDefaultInput = New Button()
        Label3 = New Label()
        CmdExit = New Button()
        CmdTestGCL = New Button()
        CmdTestSnGCL = New Button()
        Panel4 = New Panel()
        CmdFilterByIndexPromptInput = New Button()
        CmdFilterByIndexDefaultInput = New Button()
        Label4 = New Label()
        CmdShowAppConfig = New Button()
        CmdViewUserGuide = New Button()
        CmdHowToEnableDefaultBtns = New Button()
        Panel1.SuspendLayout()
        Panel2.SuspendLayout()
        Panel3.SuspendLayout()
        Panel4.SuspendLayout()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(13, 16)
        Label1.Margin = New Padding(2, 0, 2, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(185, 80)
        Label1.TabIndex = 0
        Label1.Text = "Generate HTML Blogbook" & vbCrLf & "from XML entries input file" & vbCrLf & "optionally filtered by" & vbCrLf & "content matching string(s)"
        ' 
        ' Panel1
        ' 
        Panel1.BorderStyle = BorderStyle.FixedSingle
        Panel1.Controls.Add(CmdBEFSToBookPromptInput)
        Panel1.Controls.Add(CmdBEFSToBookDefaultInput)
        Panel1.Controls.Add(Label1)
        Panel1.Location = New Point(11, 9)
        Panel1.Margin = New Padding(2, 3, 2, 3)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(493, 115)
        Panel1.TabIndex = 1
        ' 
        ' CmdBEFSToBookPromptInput
        ' 
        CmdBEFSToBookPromptInput.Location = New Point(352, 25)
        CmdBEFSToBookPromptInput.Margin = New Padding(2, 3, 2, 3)
        CmdBEFSToBookPromptInput.Name = "CmdBEFSToBookPromptInput"
        CmdBEFSToBookPromptInput.Size = New Size(119, 57)
        CmdBEFSToBookPromptInput.TabIndex = 2
        CmdBEFSToBookPromptInput.Text = "Prompt For Input and Run"
        CmdBEFSToBookPromptInput.UseVisualStyleBackColor = True
        ' 
        ' CmdBEFSToBookDefaultInput
        ' 
        CmdBEFSToBookDefaultInput.Location = New Point(214, 27)
        CmdBEFSToBookDefaultInput.Margin = New Padding(2, 3, 2, 3)
        CmdBEFSToBookDefaultInput.Name = "CmdBEFSToBookDefaultInput"
        CmdBEFSToBookDefaultInput.Size = New Size(114, 55)
        CmdBEFSToBookDefaultInput.TabIndex = 1
        CmdBEFSToBookDefaultInput.Text = "Run with Default Input"
        CmdBEFSToBookDefaultInput.UseVisualStyleBackColor = True
        ' 
        ' Panel2
        ' 
        Panel2.BorderStyle = BorderStyle.FixedSingle
        Panel2.Controls.Add(CmdFilterBEFPromptInput)
        Panel2.Controls.Add(CmdFilterBEFDefaultInput)
        Panel2.Controls.Add(Label2)
        Panel2.Location = New Point(9, 220)
        Panel2.Margin = New Padding(2, 3, 2, 3)
        Panel2.Name = "Panel2"
        Panel2.Size = New Size(496, 93)
        Panel2.TabIndex = 3
        ' 
        ' CmdFilterBEFPromptInput
        ' 
        CmdFilterBEFPromptInput.Location = New Point(354, 21)
        CmdFilterBEFPromptInput.Margin = New Padding(2, 3, 2, 3)
        CmdFilterBEFPromptInput.Name = "CmdFilterBEFPromptInput"
        CmdFilterBEFPromptInput.Size = New Size(119, 57)
        CmdFilterBEFPromptInput.TabIndex = 2
        CmdFilterBEFPromptInput.Text = "Prompt For Input and Run"
        CmdFilterBEFPromptInput.UseVisualStyleBackColor = True
        ' 
        ' CmdFilterBEFDefaultInput
        ' 
        CmdFilterBEFDefaultInput.Location = New Point(214, 21)
        CmdFilterBEFDefaultInput.Margin = New Padding(2, 3, 2, 3)
        CmdFilterBEFDefaultInput.Name = "CmdFilterBEFDefaultInput"
        CmdFilterBEFDefaultInput.Size = New Size(114, 55)
        CmdFilterBEFDefaultInput.TabIndex = 0
        CmdFilterBEFDefaultInput.Text = "Run with Default Input"
        CmdFilterBEFDefaultInput.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(13, 16)
        Label2.Margin = New Padding(2, 0, 2, 0)
        Label2.Name = "Label2"
        Label2.Size = New Size(177, 60)
        Label2.TabIndex = 1
        Label2.Text = "Filter Blog XML" & vbCrLf & "Backup/Export File based" & vbCrLf & "on Date Range"
        ' 
        ' Panel3
        ' 
        Panel3.BorderStyle = BorderStyle.FixedSingle
        Panel3.Controls.Add(CmdSBBGCLsPromptInput)
        Panel3.Controls.Add(CmdSBBGCLsDefaultInput)
        Panel3.Controls.Add(Label3)
        Panel3.Location = New Point(11, 129)
        Panel3.Margin = New Padding(2, 3, 2, 3)
        Panel3.Name = "Panel3"
        Panel3.Size = New Size(493, 86)
        Panel3.TabIndex = 2
        ' 
        ' CmdSBBGCLsPromptInput
        ' 
        CmdSBBGCLsPromptInput.Location = New Point(352, 15)
        CmdSBBGCLsPromptInput.Margin = New Padding(2, 3, 2, 3)
        CmdSBBGCLsPromptInput.Name = "CmdSBBGCLsPromptInput"
        CmdSBBGCLsPromptInput.Size = New Size(119, 57)
        CmdSBBGCLsPromptInput.TabIndex = 1
        CmdSBBGCLsPromptInput.Text = "Prompt For Input and Run"
        CmdSBBGCLsPromptInput.UseVisualStyleBackColor = True
        ' 
        ' CmdSBBGCLsDefaultInput
        ' 
        CmdSBBGCLsDefaultInput.Location = New Point(214, 16)
        CmdSBBGCLsDefaultInput.Margin = New Padding(2, 3, 2, 3)
        CmdSBBGCLsDefaultInput.Name = "CmdSBBGCLsDefaultInput"
        CmdSBBGCLsDefaultInput.Size = New Size(114, 55)
        CmdSBBGCLsDefaultInput.TabIndex = 0
        CmdSBBGCLsDefaultInput.Text = "Run with Default Input"
        CmdSBBGCLsDefaultInput.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(13, 16)
        Label3.Margin = New Padding(2, 0, 2, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(183, 40)
        Label3.TabIndex = 0
        Label3.Text = "Generate Contents Links" & vbCrLf & "List and/or Split Blogbook"
        ' 
        ' CmdExit
        ' 
        CmdExit.Location = New Point(395, 433)
        CmdExit.Margin = New Padding(2, 3, 2, 3)
        CmdExit.Name = "CmdExit"
        CmdExit.Size = New Size(105, 49)
        CmdExit.TabIndex = 9
        CmdExit.Text = "Exit"
        CmdExit.UseVisualStyleBackColor = True
        ' 
        ' CmdTestGCL
        ' 
        CmdTestGCL.Location = New Point(151, 497)
        CmdTestGCL.Margin = New Padding(2, 3, 2, 3)
        CmdTestGCL.Name = "CmdTestGCL"
        CmdTestGCL.Size = New Size(97, 27)
        CmdTestGCL.TabIndex = 3
        CmdTestGCL.TabStop = False
        CmdTestGCL.Text = "Test GCL"
        CmdTestGCL.UseVisualStyleBackColor = False
        ' 
        ' CmdTestSnGCL
        ' 
        CmdTestSnGCL.Location = New Point(271, 497)
        CmdTestSnGCL.Margin = New Padding(2, 3, 2, 3)
        CmdTestSnGCL.Name = "CmdTestSnGCL"
        CmdTestSnGCL.Size = New Size(101, 27)
        CmdTestSnGCL.TabIndex = 5
        CmdTestSnGCL.TabStop = False
        CmdTestSnGCL.Text = "Test SnGCL"
        CmdTestSnGCL.UseVisualStyleBackColor = False
        ' 
        ' Panel4
        ' 
        Panel4.BorderStyle = BorderStyle.FixedSingle
        Panel4.Controls.Add(CmdFilterByIndexPromptInput)
        Panel4.Controls.Add(CmdFilterByIndexDefaultInput)
        Panel4.Controls.Add(Label4)
        Panel4.Location = New Point(9, 317)
        Panel4.Margin = New Padding(2, 3, 2, 3)
        Panel4.Name = "Panel4"
        Panel4.Size = New Size(496, 93)
        Panel4.TabIndex = 4
        ' 
        ' CmdFilterByIndexPromptInput
        ' 
        CmdFilterByIndexPromptInput.Location = New Point(354, 19)
        CmdFilterByIndexPromptInput.Margin = New Padding(2, 3, 2, 3)
        CmdFilterByIndexPromptInput.Name = "CmdFilterByIndexPromptInput"
        CmdFilterByIndexPromptInput.Size = New Size(117, 57)
        CmdFilterByIndexPromptInput.TabIndex = 1
        CmdFilterByIndexPromptInput.Text = "Prompt For Input and Run"
        CmdFilterByIndexPromptInput.UseVisualStyleBackColor = True
        ' 
        ' CmdFilterByIndexDefaultInput
        ' 
        CmdFilterByIndexDefaultInput.Location = New Point(214, 21)
        CmdFilterByIndexDefaultInput.Margin = New Padding(2, 3, 2, 3)
        CmdFilterByIndexDefaultInput.Name = "CmdFilterByIndexDefaultInput"
        CmdFilterByIndexDefaultInput.Size = New Size(114, 55)
        CmdFilterByIndexDefaultInput.TabIndex = 0
        CmdFilterByIndexDefaultInput.Text = "Run with Default Input"
        CmdFilterByIndexDefaultInput.UseVisualStyleBackColor = True
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(13, 16)
        Label4.Margin = New Padding(2, 0, 2, 0)
        Label4.Name = "Label4"
        Label4.Size = New Size(177, 60)
        Label4.TabIndex = 0
        Label4.Text = "Filter Blog XML" & vbCrLf & "Backup/Export File based" & vbCrLf & "on Index List"
        ' 
        ' CmdShowAppConfig
        ' 
        CmdShowAppConfig.Location = New Point(271, 433)
        CmdShowAppConfig.Margin = New Padding(2, 3, 2, 3)
        CmdShowAppConfig.Name = "CmdShowAppConfig"
        CmdShowAppConfig.Size = New Size(101, 59)
        CmdShowAppConfig.TabIndex = 10
        CmdShowAppConfig.Text = "View App Config file"
        CmdShowAppConfig.UseVisualStyleBackColor = True
        ' 
        ' CmdViewUserGuide
        ' 
        CmdViewUserGuide.Location = New Point(151, 433)
        CmdViewUserGuide.Margin = New Padding(2, 3, 2, 3)
        CmdViewUserGuide.Name = "CmdViewUserGuide"
        CmdViewUserGuide.Size = New Size(97, 59)
        CmdViewUserGuide.TabIndex = 11
        CmdViewUserGuide.Text = "View User Guide"
        CmdViewUserGuide.UseVisualStyleBackColor = True
        ' 
        ' CmdHowToEnableDefaultBtns
        ' 
        CmdHowToEnableDefaultBtns.Location = New Point(11, 433)
        CmdHowToEnableDefaultBtns.Margin = New Padding(2, 3, 2, 3)
        CmdHowToEnableDefaultBtns.Name = "CmdHowToEnableDefaultBtns"
        CmdHowToEnableDefaultBtns.Size = New Size(124, 59)
        CmdHowToEnableDefaultBtns.TabIndex = 12
        CmdHowToEnableDefaultBtns.Text = "How to Enable Default Buttons"
        CmdHowToEnableDefaultBtns.UseVisualStyleBackColor = True
        ' 
        ' MainRunOptionsForm
        ' 
        AutoScaleDimensions = New SizeF(8F, 20F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(519, 525)
        Controls.Add(CmdHowToEnableDefaultBtns)
        Controls.Add(CmdViewUserGuide)
        Controls.Add(CmdShowAppConfig)
        Controls.Add(Panel4)
        Controls.Add(CmdTestSnGCL)
        Controls.Add(CmdTestGCL)
        Controls.Add(CmdExit)
        Controls.Add(Panel3)
        Controls.Add(Panel2)
        Controls.Add(Panel1)
        Margin = New Padding(2, 3, 2, 3)
        Name = "MainRunOptionsForm"
        Text = "Blog Backup XML file to HTML Blogbook"
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        Panel2.ResumeLayout(False)
        Panel2.PerformLayout()
        Panel3.ResumeLayout(False)
        Panel3.PerformLayout()
        Panel4.ResumeLayout(False)
        Panel4.PerformLayout()
        ResumeLayout(False)
    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents CmdBEFSToBookDefaultInput As Button
    Friend WithEvents CmdBEFSToBookPromptInput As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents CmdFilterBEFPromptInput As Button
    Friend WithEvents CmdFilterBEFDefaultInput As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents CmdSBBGCLsPromptInput As Button
    Friend WithEvents CmdSBBGCLsDefaultInput As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents CmdExit As Button
    Friend WithEvents CmdTestGCL As Button
    Friend WithEvents CmdTestSnGCL As Button
    Friend WithEvents Panel4 As Panel
    Friend WithEvents CmdFilterByIndexPromptInput As Button
    Friend WithEvents CmdFilterByIndexDefaultInput As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents CmdShowAppConfig As Button
    Friend WithEvents CmdViewUserGuide As Button
    Friend WithEvents CmdHowToEnableDefaultBtns As Button
End Class
