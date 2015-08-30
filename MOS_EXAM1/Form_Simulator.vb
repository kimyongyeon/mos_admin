

Public Class Form_Simulator

    Private AB As New frmTaskBarSettings.TBAppBar 'AppBar Object

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' 2009.12.16 : SCS : 디자인 패턴을 적용하여 막아야 되는 핫키를 추가하는 로직 필요
        HookKeyboard()
    End Sub

    Private Sub Form_Simulator_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        UnhookKeyboard()
        AB.AlwaysOnTop = True 'Set Always On Top On
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        AB.AlwaysOnTop = False 'Set Always On Top Off
    End Sub

    Private Sub Form_Simulator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        End
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Form_Navigator.Show()
    End Sub
End Class
