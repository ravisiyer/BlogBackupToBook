<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BEFToBookParamForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(BEFToBookParamForm))
        Label1 = New Label()
        CheckBoxCombineUsingAND = New CheckBox()
        TextBoxSS1 = New TextBox()
        TextBoxSS3 = New TextBox()
        TextBoxSS2 = New TextBox()
        TextBoxSS4 = New TextBox()
        TextBoxSS5 = New TextBox()
        CheckBoxCreateBlogbook = New CheckBox()
        Label2 = New Label()
        CommandButtonRun = New Button()
        CommandButtonCancel = New Button()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 95)
        Label1.Name = "Label1"
        Label1.Size = New Size(147, 25)
        Label1.TabIndex = 0
        Label1.Text = "Search strings (5)"
        ' 
        ' CheckBoxCombineUsingAND
        ' 
        CheckBoxCombineUsingAND.AutoSize = True
        CheckBoxCombineUsingAND.Location = New Point(184, 95)
        CheckBoxCombineUsingAND.Name = "CheckBoxCombineUsingAND"
        CheckBoxCombineUsingAND.Size = New Size(493, 29)
        CheckBoxCombineUsingAND.TabIndex = 1
        CheckBoxCombineUsingAND.Text = "Combine Search Strings using AND instead of Default OR"
        CheckBoxCombineUsingAND.UseVisualStyleBackColor = True
        ' 
        ' TextBoxSS1
        ' 
        TextBoxSS1.Location = New Point(12, 140)
        TextBoxSS1.Name = "TextBoxSS1"
        TextBoxSS1.Size = New Size(1345, 31)
        TextBoxSS1.TabIndex = 2
        ' 
        ' TextBoxSS3
        ' 
        TextBoxSS3.Location = New Point(12, 237)
        TextBoxSS3.Name = "TextBoxSS3"
        TextBoxSS3.Size = New Size(1345, 31)
        TextBoxSS3.TabIndex = 4
        ' 
        ' TextBoxSS2
        ' 
        TextBoxSS2.Location = New Point(12, 186)
        TextBoxSS2.Name = "TextBoxSS2"
        TextBoxSS2.Size = New Size(1345, 31)
        TextBoxSS2.TabIndex = 3
        ' 
        ' TextBoxSS4
        ' 
        TextBoxSS4.Location = New Point(12, 283)
        TextBoxSS4.Name = "TextBoxSS4"
        TextBoxSS4.Size = New Size(1345, 31)
        TextBoxSS4.TabIndex = 5
        ' 
        ' TextBoxSS5
        ' 
        TextBoxSS5.Location = New Point(12, 331)
        TextBoxSS5.Name = "TextBoxSS5"
        TextBoxSS5.Size = New Size(1345, 31)
        TextBoxSS5.TabIndex = 6
        ' 
        ' CheckBoxCreateBlogbook
        ' 
        CheckBoxCreateBlogbook.AutoSize = True
        CheckBoxCreateBlogbook.Checked = True
        CheckBoxCreateBlogbook.CheckState = CheckState.Checked
        CheckBoxCreateBlogbook.Location = New Point(12, 402)
        CheckBoxCreateBlogbook.Name = "CheckBoxCreateBlogbook"
        CheckBoxCreateBlogbook.Size = New Size(732, 29)
        CheckBoxCreateBlogbook.TabIndex = 7
        CheckBoxCreateBlogbook.Text = "Create Blog book (and comments dictionary book)? If no, only log file(s) will be created."
        CheckBoxCreateBlogbook.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(12, 25)
        Label2.Name = "Label2"
        Label2.Size = New Size(1060, 50)
        Label2.TabIndex = 8
        Label2.Text = resources.GetString("Label2.Text")
        ' 
        ' CommandButtonRun
        ' 
        CommandButtonRun.Location = New Point(411, 453)
        CommandButtonRun.Name = "CommandButtonRun"
        CommandButtonRun.Size = New Size(112, 34)
        CommandButtonRun.TabIndex = 9
        CommandButtonRun.Text = "Run"
        CommandButtonRun.UseVisualStyleBackColor = True
        ' 
        ' CommandButtonCancel
        ' 
        CommandButtonCancel.Location = New Point(684, 453)
        CommandButtonCancel.Name = "CommandButtonCancel"
        CommandButtonCancel.Size = New Size(112, 34)
        CommandButtonCancel.TabIndex = 10
        CommandButtonCancel.Text = "Cancel"
        CommandButtonCancel.UseVisualStyleBackColor = True
        ' 
        ' BEFToBookParamForm
        ' 
        AutoScaleDimensions = New SizeF(10F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1369, 520)
        Controls.Add(CommandButtonCancel)
        Controls.Add(CommandButtonRun)
        Controls.Add(Label2)
        Controls.Add(CheckBoxCreateBlogbook)
        Controls.Add(TextBoxSS5)
        Controls.Add(TextBoxSS4)
        Controls.Add(TextBoxSS2)
        Controls.Add(TextBoxSS3)
        Controls.Add(TextBoxSS1)
        Controls.Add(CheckBoxCombineUsingAND)
        Controls.Add(Label1)
        Name = "BEFToBookParamForm"
        Text = "Parameters for Blog Export File Search To Book"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents CheckBoxCombineUsingAND As CheckBox
    Friend WithEvents TextBoxSS1 As TextBox
    Friend WithEvents TextBoxSS3 As TextBox
    Friend WithEvents TextBoxSS2 As TextBox
    Friend WithEvents TextBoxSS4 As TextBox
    Friend WithEvents TextBoxSS5 As TextBox
    Friend WithEvents CheckBoxCreateBlogbook As CheckBox
    Friend WithEvents Label2 As Label
    Friend WithEvents CommandButtonRun As Button
    Friend WithEvents CommandButtonCancel As Button
End Class
