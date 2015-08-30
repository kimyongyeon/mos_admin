Public Class Form_Opinion

    Private Sub Button_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Ok.Click
        Form_Result.Show()

        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form_Result.Show()

        Me.Hide()
    End Sub

    Private Sub Form_Opinion_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (e.CloseReason = CloseReason.UserClosing) OrElse (e.CloseReason = CloseReason.None) Then
            'If the user clicks the X button, do not close the form, but hide it.
            'To exit the application, the user must select "Exit" from the main or popup menu.
            e.Cancel = True
            'Hide()
        End If
    End Sub

    Private Sub Form_Opinion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class