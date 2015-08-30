Public Class Form_Result

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MessageBox.Show("시뮬레이션 프로그램에선 동작하지 않고 실제 시험에선 동작 합니다.", "시뮬 제한", MessageBoxButtons.OK)
    End Sub

    Private Sub Form_Result_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (e.CloseReason = CloseReason.UserClosing) OrElse (e.CloseReason = CloseReason.None) Then
            Form_Upload.Show()
        End If
    End Sub

    Private Sub Form_Result_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label_ExamTitle.Text = Class_Exam.examTitle
        Label_PassScore.Text = Class_Exam.iPassScore
        Label_TotalScore.Text = Class_Exam.iTotalScore

        If (Class_Exam.iTotalScore >= Class_Exam.iPassScore) Then
            Label_Result.Text = "합격"
        Else
            Label_Result.Text = "불합격"
            Label_Message.Text = "죄송합니다. 다음 시험에서 좋을 결과 있길 바랍니다."
        End If

    End Sub
End Class