<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ParamForm
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
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(ParamForm))
        Label1 = New Label()
        Label2 = New Label()
        TBSplitSize = New TextBox()
        Label3 = New Label()
        Label4 = New Label()
        TBEndOfPostMarker = New TextBox()
        GroupBox1 = New GroupBox()
        TBContentsLinksInsertMarker = New TextBox()
        Label5 = New Label()
        CBGenerateContentsLinks = New CheckBox()
        CmdRun = New Button()
        CmdCancel = New Button()
        LblSplitSizeError = New Label()
        GroupBox1.SuspendLayout()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 9)
        Label1.Name = "Label1"
        Label1.Size = New Size(676, 150)
        Label1.TabIndex = 0
        Label1.Text = resources.GetString("Label1.Text")
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(12, 173)
        Label2.Name = "Label2"
        Label2.Size = New Size(87, 25)
        Label2.TabIndex = 1
        Label2.Text = "Split Size:"
        ' 
        ' TBSplitSize
        ' 
        TBSplitSize.Location = New Point(105, 173)
        TBSplitSize.Name = "TBSplitSize"
        TBSplitSize.Size = New Size(150, 31)
        TBSplitSize.TabIndex = 2
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(286, 173)
        Label3.Name = "Label3"
        Label3.Size = New Size(292, 25)
        Label3.TabIndex = 3
        Label3.Text = "Use 0 or blank to not split input file"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(12, 222)
        Label4.Name = "Label4"
        Label4.Size = New Size(138, 25)
        Label4.TabIndex = 4
        Label4.Text = "End of Marker1:"
        ' 
        ' TBEndOfPostMarker
        ' 
        TBEndOfPostMarker.Location = New Point(12, 250)
        TBEndOfPostMarker.Name = "TBEndOfPostMarker"
        TBEndOfPostMarker.Size = New Size(676, 31)
        TBEndOfPostMarker.TabIndex = 5
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(TBContentsLinksInsertMarker)
        GroupBox1.Controls.Add(Label5)
        GroupBox1.Controls.Add(CBGenerateContentsLinks)
        GroupBox1.Location = New Point(12, 299)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(689, 150)
        GroupBox1.TabIndex = 6
        GroupBox1.TabStop = False
        GroupBox1.Text = "For .html files only"
        ' 
        ' TBContentsLinksInsertMarker
        ' 
        TBContentsLinksInsertMarker.Location = New Point(0, 119)
        TBContentsLinksInsertMarker.Name = "TBContentsLinksInsertMarker"
        TBContentsLinksInsertMarker.Size = New Size(676, 31)
        TBContentsLinksInsertMarker.TabIndex = 2
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(0, 83)
        Label5.Name = "Label5"
        Label5.Size = New Size(240, 25)
        Label5.TabIndex = 1
        Label5.Text = "Contents Links Insert Marker:"
        ' 
        ' CBGenerateContentsLinks
        ' 
        CBGenerateContentsLinks.AutoSize = True
        CBGenerateContentsLinks.Location = New Point(0, 40)
        CBGenerateContentsLinks.Name = "CBGenerateContentsLinks"
        CBGenerateContentsLinks.Size = New Size(307, 29)
        CBGenerateContentsLinks.TabIndex = 0
        CBGenerateContentsLinks.Text = "Generate Contents Links List (TOC)"
        CBGenerateContentsLinks.UseVisualStyleBackColor = True
        ' 
        ' CmdRun
        ' 
        CmdRun.Location = New Point(165, 487)
        CmdRun.Name = "CmdRun"
        CmdRun.Size = New Size(131, 34)
        CmdRun.TabIndex = 7
        CmdRun.Text = "Run Program"
        CmdRun.UseVisualStyleBackColor = True
        ' 
        ' CmdCancel
        ' 
        CmdCancel.Location = New Point(377, 487)
        CmdCancel.Name = "CmdCancel"
        CmdCancel.Size = New Size(112, 34)
        CmdCancel.TabIndex = 8
        CmdCancel.Text = "Cancel"
        CmdCancel.UseVisualStyleBackColor = True
        ' 
        ' LblSplitSizeError
        ' 
        LblSplitSizeError.AutoSize = True
        LblSplitSizeError.Location = New Point(286, 208)
        LblSplitSizeError.Name = "LblSplitSizeError"
        LblSplitSizeError.Size = New Size(0, 25)
        LblSplitSizeError.TabIndex = 9
        ' 
        ' ParamForm
        ' 
        AutoScaleDimensions = New SizeF(10.0F, 25.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(705, 543)
        Controls.Add(LblSplitSizeError)
        Controls.Add(CmdCancel)
        Controls.Add(CmdRun)
        Controls.Add(GroupBox1)
        Controls.Add(TBEndOfPostMarker)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(TBSplitSize)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Name = "ParamForm"
        Text = "Specify Program Parameters"
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub



    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TBSplitSize As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TBEndOfPostMarker As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TBContentsLinksInsertMarker As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents CBGenerateContentsLinks As CheckBox
    Friend WithEvents CmdRun As Button
    Friend WithEvents CmdCancel As Button
    Friend WithEvents LblSplitSizeError As Label
End Class
