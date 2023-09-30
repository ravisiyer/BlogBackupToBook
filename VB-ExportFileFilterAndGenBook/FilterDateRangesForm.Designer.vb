<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FilterDateRangesForm
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
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        StartPublishedDateTB = New TextBox()
        EndPublishedDateTB = New TextBox()
        Label4 = New Label()
        EndUpdatedDateTB = New TextBox()
        Label5 = New Label()
        StartUpdatedDateTB = New TextBox()
        Label6 = New Label()
        CmdRunFilter = New Button()
        CmdCancel = New Button()
        LblStartPublishedCDate = New Label()
        LblEndPublishedCDate = New Label()
        LblEndUpdatedCDate = New Label()
        LblStartUpdatedCDate = New Label()
        LblErrorMessage = New Label()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(19, 14)
        Label1.Margin = New Padding(2, 0, 2, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(416, 15)
        Label1.TabIndex = 0
        Label1.Text = "Please specify date ranges for filter program. All blank entries implies no filter."
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(19, 35)
        Label2.Margin = New Padding(2, 0, 2, 0)
        Label2.Name = "Label2"
        Label2.Size = New Size(423, 30)
        Label2.TabIndex = 1
        Label2.Text = "All dates may be specified as YYYY-MM-DD. E.g. 2022-01-01, 2022-12-31. Other" & vbCrLf & "formats like dd/mm/yyyy also seem to work but have not been tested much."
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(19, 76)
        Label3.Margin = New Padding(2, 0, 2, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(116, 15)
        Label3.TabIndex = 2
        Label3.Text = "Start Published Date:"
        ' 
        ' StartPublishedDateTB
        ' 
        StartPublishedDateTB.Location = New Point(146, 73)
        StartPublishedDateTB.Margin = New Padding(2, 2, 2, 2)
        StartPublishedDateTB.Name = "StartPublishedDateTB"
        StartPublishedDateTB.Size = New Size(106, 23)
        StartPublishedDateTB.TabIndex = 3
        ' 
        ' EndPublishedDateTB
        ' 
        EndPublishedDateTB.Location = New Point(146, 102)
        EndPublishedDateTB.Margin = New Padding(2, 2, 2, 2)
        EndPublishedDateTB.Name = "EndPublishedDateTB"
        EndPublishedDateTB.Size = New Size(106, 23)
        EndPublishedDateTB.TabIndex = 5
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(19, 106)
        Label4.Margin = New Padding(2, 0, 2, 0)
        Label4.Name = "Label4"
        Label4.Size = New Size(112, 15)
        Label4.TabIndex = 4
        Label4.Text = "End Published Date:"
        ' 
        ' EndUpdatedDateTB
        ' 
        EndUpdatedDateTB.Location = New Point(146, 166)
        EndUpdatedDateTB.Margin = New Padding(2, 2, 2, 2)
        EndUpdatedDateTB.Name = "EndUpdatedDateTB"
        EndUpdatedDateTB.Size = New Size(106, 23)
        EndUpdatedDateTB.TabIndex = 9
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(19, 169)
        Label5.Margin = New Padding(2, 0, 2, 0)
        Label5.Name = "Label5"
        Label5.Size = New Size(105, 15)
        Label5.TabIndex = 8
        Label5.Text = "End Updated Date:"
        ' 
        ' StartUpdatedDateTB
        ' 
        StartUpdatedDateTB.Location = New Point(146, 136)
        StartUpdatedDateTB.Margin = New Padding(2, 2, 2, 2)
        StartUpdatedDateTB.Name = "StartUpdatedDateTB"
        StartUpdatedDateTB.Size = New Size(106, 23)
        StartUpdatedDateTB.TabIndex = 7
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(19, 140)
        Label6.Margin = New Padding(2, 0, 2, 0)
        Label6.Name = "Label6"
        Label6.Size = New Size(109, 15)
        Label6.TabIndex = 6
        Label6.Text = "Start Updated Date:"
        ' 
        ' CmdRunFilter
        ' 
        CmdRunFilter.Location = New Point(54, 205)
        CmdRunFilter.Margin = New Padding(2, 2, 2, 2)
        CmdRunFilter.Name = "CmdRunFilter"
        CmdRunFilter.Size = New Size(78, 26)
        CmdRunFilter.TabIndex = 10
        CmdRunFilter.Text = "Run Filter"
        CmdRunFilter.UseVisualStyleBackColor = True
        ' 
        ' CmdCancel
        ' 
        CmdCancel.Location = New Point(164, 205)
        CmdCancel.Margin = New Padding(2, 2, 2, 2)
        CmdCancel.Name = "CmdCancel"
        CmdCancel.Size = New Size(78, 26)
        CmdCancel.TabIndex = 11
        CmdCancel.Text = "Cancel"
        CmdCancel.UseVisualStyleBackColor = True
        ' 
        ' LblStartPublishedCDate
        ' 
        LblStartPublishedCDate.AutoSize = True
        LblStartPublishedCDate.Location = New Point(283, 76)
        LblStartPublishedCDate.Margin = New Padding(2, 0, 2, 0)
        LblStartPublishedCDate.Name = "LblStartPublishedCDate"
        LblStartPublishedCDate.Size = New Size(0, 15)
        LblStartPublishedCDate.TabIndex = 12
        ' 
        ' LblEndPublishedCDate
        ' 
        LblEndPublishedCDate.AutoSize = True
        LblEndPublishedCDate.Location = New Point(283, 106)
        LblEndPublishedCDate.Margin = New Padding(2, 0, 2, 0)
        LblEndPublishedCDate.Name = "LblEndPublishedCDate"
        LblEndPublishedCDate.Size = New Size(0, 15)
        LblEndPublishedCDate.TabIndex = 13
        ' 
        ' LblEndUpdatedCDate
        ' 
        LblEndUpdatedCDate.AutoSize = True
        LblEndUpdatedCDate.Location = New Point(283, 169)
        LblEndUpdatedCDate.Margin = New Padding(2, 0, 2, 0)
        LblEndUpdatedCDate.Name = "LblEndUpdatedCDate"
        LblEndUpdatedCDate.Size = New Size(0, 15)
        LblEndUpdatedCDate.TabIndex = 15
        ' 
        ' LblStartUpdatedCDate
        ' 
        LblStartUpdatedCDate.AutoSize = True
        LblStartUpdatedCDate.Location = New Point(283, 140)
        LblStartUpdatedCDate.Margin = New Padding(2, 0, 2, 0)
        LblStartUpdatedCDate.Name = "LblStartUpdatedCDate"
        LblStartUpdatedCDate.Size = New Size(0, 15)
        LblStartUpdatedCDate.TabIndex = 14
        ' 
        ' LblErrorMessage
        ' 
        LblErrorMessage.AutoSize = True
        LblErrorMessage.Location = New Point(19, 241)
        LblErrorMessage.Margin = New Padding(2, 0, 2, 0)
        LblErrorMessage.Name = "LblErrorMessage"
        LblErrorMessage.Size = New Size(0, 15)
        LblErrorMessage.TabIndex = 16
        ' 
        ' FilterDateRangesForm
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(456, 258)
        Controls.Add(LblErrorMessage)
        Controls.Add(LblEndUpdatedCDate)
        Controls.Add(LblStartUpdatedCDate)
        Controls.Add(LblEndPublishedCDate)
        Controls.Add(LblStartPublishedCDate)
        Controls.Add(CmdCancel)
        Controls.Add(CmdRunFilter)
        Controls.Add(EndUpdatedDateTB)
        Controls.Add(Label5)
        Controls.Add(StartUpdatedDateTB)
        Controls.Add(Label6)
        Controls.Add(EndPublishedDateTB)
        Controls.Add(Label4)
        Controls.Add(StartPublishedDateTB)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Margin = New Padding(2, 2, 2, 2)
        Name = "FilterDateRangesForm"
        Text = "Specify Date  Ranges for Filter"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents StartPublishedDateTB As TextBox
    Friend WithEvents EndPublishedDateTB As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents EndUpdatedDateTB As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents StartUpdatedDateTB As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents CmdRunFilter As Button
    Friend WithEvents CmdCancel As Button
    Friend WithEvents LblStartPublishedCDate As Label
    Friend WithEvents LblEndPublishedCDate As Label
    Friend WithEvents LblEndUpdatedCDate As Label
    Friend WithEvents LblStartUpdatedCDate As Label
    Friend WithEvents LblErrorMessage As Label
End Class
