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
        Label1.Location = New Point(27, 24)
        Label1.Name = "Label1"
        Label1.Size = New Size(627, 25)
        Label1.TabIndex = 0
        Label1.Text = "Please specify date ranges for filter program. All blank entries implies no filter."
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(27, 59)
        Label2.Name = "Label2"
        Label2.Size = New Size(651, 50)
        Label2.TabIndex = 1
        Label2.Text = "All dates may be specified as YYYY-MM-DD. E.g. 2022-01-01, 2022-12-31. Other" & vbCrLf & "formats like dd/mm/yyyy also seem to work but have not been tested much."
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(27, 127)
        Label3.Name = "Label3"
        Label3.Size = New Size(176, 25)
        Label3.TabIndex = 2
        Label3.Text = "Start Published Date:"
        ' 
        ' StartPublishedDateTB
        ' 
        StartPublishedDateTB.Location = New Point(209, 121)
        StartPublishedDateTB.Name = "StartPublishedDateTB"
        StartPublishedDateTB.Size = New Size(150, 31)
        StartPublishedDateTB.TabIndex = 3
        ' 
        ' EndPublishedDateTB
        ' 
        EndPublishedDateTB.Location = New Point(209, 170)
        EndPublishedDateTB.Name = "EndPublishedDateTB"
        EndPublishedDateTB.Size = New Size(150, 31)
        EndPublishedDateTB.TabIndex = 5
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(27, 176)
        Label4.Name = "Label4"
        Label4.Size = New Size(170, 25)
        Label4.TabIndex = 4
        Label4.Text = "End Published Date:"
        ' 
        ' EndUpdatedDateTB
        ' 
        EndUpdatedDateTB.Location = New Point(209, 276)
        EndUpdatedDateTB.Name = "EndUpdatedDateTB"
        EndUpdatedDateTB.Size = New Size(150, 31)
        EndUpdatedDateTB.TabIndex = 9
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(27, 282)
        Label5.Name = "Label5"
        Label5.Size = New Size(162, 25)
        Label5.TabIndex = 8
        Label5.Text = "End Updated Date:"
        ' 
        ' StartUpdatedDateTB
        ' 
        StartUpdatedDateTB.Location = New Point(209, 227)
        StartUpdatedDateTB.Name = "StartUpdatedDateTB"
        StartUpdatedDateTB.Size = New Size(150, 31)
        StartUpdatedDateTB.TabIndex = 7
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(27, 233)
        Label6.Name = "Label6"
        Label6.Size = New Size(168, 25)
        Label6.TabIndex = 6
        Label6.Text = "Start Updated Date:"
        ' 
        ' CmdRunFilter
        ' 
        CmdRunFilter.Location = New Point(77, 342)
        CmdRunFilter.Name = "CmdRunFilter"
        CmdRunFilter.Size = New Size(112, 34)
        CmdRunFilter.TabIndex = 10
        CmdRunFilter.Text = "Run Filter"
        CmdRunFilter.UseVisualStyleBackColor = True
        ' 
        ' CmdCancel
        ' 
        CmdCancel.Location = New Point(234, 342)
        CmdCancel.Name = "CmdCancel"
        CmdCancel.Size = New Size(112, 34)
        CmdCancel.TabIndex = 11
        CmdCancel.Text = "Cancel"
        CmdCancel.UseVisualStyleBackColor = True
        ' 
        ' LblStartPublishedCDate
        ' 
        LblStartPublishedCDate.AutoSize = True
        LblStartPublishedCDate.Location = New Point(404, 127)
        LblStartPublishedCDate.Name = "LblStartPublishedCDate"
        LblStartPublishedCDate.Size = New Size(0, 25)
        LblStartPublishedCDate.TabIndex = 12
        ' 
        ' LblEndPublishedCDate
        ' 
        LblEndPublishedCDate.AutoSize = True
        LblEndPublishedCDate.Location = New Point(404, 176)
        LblEndPublishedCDate.Name = "LblEndPublishedCDate"
        LblEndPublishedCDate.Size = New Size(0, 25)
        LblEndPublishedCDate.TabIndex = 13
        ' 
        ' LblEndUpdatedCDate
        ' 
        LblEndUpdatedCDate.AutoSize = True
        LblEndUpdatedCDate.Location = New Point(404, 282)
        LblEndUpdatedCDate.Name = "LblEndUpdatedCDate"
        LblEndUpdatedCDate.Size = New Size(0, 25)
        LblEndUpdatedCDate.TabIndex = 15
        ' 
        ' LblStartUpdatedCDate
        ' 
        LblStartUpdatedCDate.AutoSize = True
        LblStartUpdatedCDate.Location = New Point(404, 233)
        LblStartUpdatedCDate.Name = "LblStartUpdatedCDate"
        LblStartUpdatedCDate.Size = New Size(0, 25)
        LblStartUpdatedCDate.TabIndex = 14
        ' 
        ' LblErrorMessage
        ' 
        LblErrorMessage.AutoSize = True
        LblErrorMessage.Location = New Point(27, 402)
        LblErrorMessage.Name = "LblErrorMessage"
        LblErrorMessage.Size = New Size(0, 25)
        LblErrorMessage.TabIndex = 16
        ' 
        ' FilterDateRangesForm
        ' 
        AutoScaleDimensions = New SizeF(10F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(693, 473)
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
