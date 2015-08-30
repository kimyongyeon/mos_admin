Public Class Form_ExamInfo

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MessageBox.Show("시뮬레이션 프로그램에선 동작하지 않고 실제 시험에선 동작 합니다.", "시뮬 제한", MessageBoxButtons.OK)
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        MessageBox.Show("시뮬레이션 프로그램에선 동작하지 않고 실제 시험에선 동작 합니다.", "시뮬 제한", MessageBoxButtons.OK)
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        MessageBox.Show("시뮬레이션 프로그램에선 동작하지 않고 실제 시험에선 동작 합니다.", "시뮬 제한", MessageBoxButtons.OK)
    End Sub

    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Close()
    End Sub

    Private Sub Button_Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Start.Click
        Form_Navigator.Show()

        Me.Hide()
    End Sub

    Private Sub Form_ExamInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label_ExamTitle.Text = Class_Exam.examTitle
    End Sub
End Class