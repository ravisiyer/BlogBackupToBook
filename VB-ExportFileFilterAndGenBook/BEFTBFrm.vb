Public Class BEFToBookParamForm


    Private Sub CommandButtonRun_Click(sender As Object, e As EventArgs) Handles CommandButtonRun.Click
        RunBEFToBookProgram = True
        Me.Hide()
    End Sub

    Private Sub CommandButtonCancel_Click(sender As Object, e As EventArgs) Handles CommandButtonCancel.Click
        RunBEFToBookProgram = False
        Me.Hide()
    End Sub
End Class