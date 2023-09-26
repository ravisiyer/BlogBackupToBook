Imports System.Windows

Public Class ParamForm
    Private Sub CmdRun_Click(sender As Object, e As EventArgs) Handles CmdRun.Click
        ValidateTBSplitSize()

        If Me.LblSplitSizeError.Text <> "" Then
            MsgBox("Form input data is invalid. Please correct and retry", , MsgBoxTitle)
            Exit Sub
        End If
        RunProgram = True
        Me.Hide()
    End Sub

    Private Sub CmdCancel_Click(sender As Object, e As EventArgs) Handles CmdCancel.Click
        RunProgram = False
        Me.Hide()
    End Sub

    Private Sub ValidateTBSplitSize()
        If Me.TBSplitSize.Text = "" Then
            Me.LblSplitSizeError.Text = ""
        ElseIf IsNumeric(Me.TBSplitSize.Text) Then
            Me.LblSplitSizeError.Text = ""
        Else
            Me.LblSplitSizeError.Text = "Split Size is not numeric"
        End If
    End Sub

    Private Sub TBSplitSize_LostFocus(sender As Object, e As EventArgs) Handles TBSplitSize.LostFocus
        ValidateTBSplitSize()
    End Sub
End Class