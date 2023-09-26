Public Class FilterDateRangesForm

    Private Sub CmdRunFilter_Click(sender As Object, e As EventArgs) Handles CmdRunFilter.Click
        'Me.StartPublishedDateTB.SelStart = 0
        'Me.StartPublishedDateTB.SelLength = Len(Me.StartPublishedDateTB.Text)
        Me.StartPublishedDateTB.Select()

        If ValidateDates() Then
            RunFilterProgram = True
            Me.Hide()
        End If
    End Sub

    Private Function ValidateDates() As Boolean

        Dim d1 As Date

        If Me.StartPublishedDateTB.Text <> "" Then
            If IsDate(Me.StartPublishedDateTB.Text) Then
                d1 = CDate(Me.StartPublishedDateTB.Text)
            Else
                Me.LblErrorMessage.Text = "Invalid Start Published Date"
                ValidateDates = False
                Exit Function
            End If
        End If

        If Me.EndPublishedDateTB.Text <> "" Then
            If IsDate(Me.EndPublishedDateTB.Text) Then
                d1 = CDate(Me.EndPublishedDateTB.Text)
            Else
                Me.LblErrorMessage.Text = "Invalid End Published Date"
                ValidateDates = False
                Exit Function
            End If
        End If

        If Me.StartUpdatedDateTB.Text <> "" Then
            If IsDate(Me.StartUpdatedDateTB.Text) Then
                d1 = CDate(Me.StartUpdatedDateTB.Text)
            Else
                Me.LblErrorMessage.Text = "Invalid Start Updated Date"
                ValidateDates = False
                Exit Function
            End If
        End If

        If Me.EndUpdatedDateTB.Text <> "" Then
            If IsDate(Me.EndUpdatedDateTB.Text) Then
                d1 = CDate(Me.EndUpdatedDateTB.Text)
            Else
                Me.LblErrorMessage.Text = "Invalid End Updated Date"
                ValidateDates = False
                Exit Function
            End If
        End If

        ValidateDates = True
    End Function

    Private Sub CmdCancel_Click(sender As Object, e As EventArgs) Handles CmdCancel.Click
        RunFilterProgram = False
        Me.Hide()

    End Sub

    Private Sub StartPublishedDateTB_LostFocus(sender As Object, e As EventArgs) Handles StartPublishedDateTB.LostFocus
        Dim d1str As String
        Dim d1 As Date

        d1str = Me.StartPublishedDateTB.Text
        If d1str = "" Then
            Me.LblStartPublishedCDate.Text = ""
            Me.LblErrorMessage.Text = ""
            Exit Sub
        End If
        If IsDate(d1str) Then
            d1 = CDate(d1str)
            Me.LblStartPublishedCDate.Text = d1
            Me.LblErrorMessage.Text = ""
        Else
            Me.LblStartPublishedCDate.Text = "Invalid date"
            Me.LblErrorMessage.Text = "Typed in date string: " & d1str &
                ". IsDate() returned false indicating invalid date string"
        End If
    End Sub

    Private Sub EndPublishedDateTB_LostFocus(sender As Object, e As EventArgs) Handles EndPublishedDateTB.LostFocus
        Dim d1str As String
        Dim d1 As Date

        d1str = Me.EndPublishedDateTB.Text
        If d1str = "" Then
            Me.LblEndPublishedCDate.Text = ""
            Me.LblErrorMessage.Text = ""
            Exit Sub
        End If
        If IsDate(d1str) Then
            d1 = CDate(d1str)
            Me.LblEndPublishedCDate.Text = d1
            Me.LblErrorMessage.Text = ""
        Else
            Me.LblEndPublishedCDate.Text = "Invalid date"
            Me.LblErrorMessage.Text = "Typed in date string: " & d1str &
                ". IsDate() returned false indicating invalid date string"
        End If
    End Sub

    Private Sub StartUpdatedDateTB_LostFocus(sender As Object, e As EventArgs) Handles StartUpdatedDateTB.LostFocus
        Dim d1str As String
        Dim d1 As Date

        d1str = Me.StartUpdatedDateTB.Text
        If d1str = "" Then
            Me.LblStartUpdatedCDate.Text = ""
            Me.LblErrorMessage.Text = ""
            Exit Sub
        End If
        If IsDate(d1str) Then
            d1 = CDate(d1str)
            Me.LblStartUpdatedCDate.Text = d1
            Me.LblErrorMessage.Text = ""
        Else
            Me.LblStartUpdatedCDate.Text = "Invalid date"
            Me.LblErrorMessage.Text = "Typed in date string: " & d1str &
                ". IsDate() returned false indicating invalid date string"
        End If
    End Sub

    Private Sub EndUpdatedDateTB_LostFocus(sender As Object, e As EventArgs) Handles EndUpdatedDateTB.LostFocus
        Dim d1str As String
        Dim d1 As Date

        d1str = Me.EndUpdatedDateTB.Text
        If d1str = "" Then
            Me.LblEndUpdatedCDate.Text = ""
            Me.LblErrorMessage.Text = ""
            Exit Sub
        End If
        If IsDate(d1str) Then
            d1 = CDate(d1str)
            Me.LblEndUpdatedCDate.Text = d1
            Me.LblErrorMessage.Text = ""
        Else
            Me.LblEndUpdatedCDate.Text = "Invalid date"
            Me.LblErrorMessage.Text = "Typed in date string: " & d1str &
                ". IsDate() returned false indicating invalid date string"
        End If
    End Sub
End Class