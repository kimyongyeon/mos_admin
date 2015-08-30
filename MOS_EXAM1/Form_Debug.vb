Public Class Form_Debug

    Private Sub Button_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Close.Click
        Me.Close()
    End Sub

    Private Sub Button_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Clear.Click
        TextBox_Msg.Clear()
    End Sub

    Sub Log(ByVal msg As String)
        TextBox_Msg.Text = Format(Now(), "hh:mm;ss ") & msg & vbCrLf & TextBox_Msg.Text
    End Sub
End Class