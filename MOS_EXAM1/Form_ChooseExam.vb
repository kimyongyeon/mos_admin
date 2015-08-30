Public Class Form_ChooseExam

    'Public Exam_name As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MessageBox.Show("시뮬레이션 프로그램에선 동작하지 않고 실제 시험에선 동작 합니다.", "시뮬 제한", MessageBoxButtons.OK)
    End Sub

    Private Sub Button_Continue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Continue.Click
        ' 선택 점검<=
        If Class_Exam.examTitle = Nothing Then
            MsgBox("시험 이름을 선택하세요")
            Exit Sub
        End If

        If Class_Exam.bookID <> "2010_M3PC1" Then
            MsgBox("파워포인트만 선택하실 수 있습니다.")
            Exit Sub
        End If

        If Class_Exam.examID = Nothing Then
            MsgBox("시험 회차를 선택하세요")
            Exit Sub
        End If

        If Class_Exam.examID <> "E1" And Class_Exam.examID <> "G1" Then
            MsgBox("해당 회차를 선택하실 수 없습니다.")
            Exit Sub
        End If

        ' 시험 정보 화면 출력
        Form_ExamInfo.Show()

        Me.Hide()
    End Sub


    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        End
    End Sub


    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

        ' 선택된 Item 얻기
        Dim selected_item As ListViewItem

        ' 현재 포커스가 있는 Item
        selected_item = Me.ListView1.FocusedItem

        ' 글로벌 변수 설정
        Class_Exam.examTitle = selected_item.SubItems(0).Text
        Select Case Class_Exam.examTitle
            Case "Microsoft Office Access 2003"
                Class_Exam.bookID = "2010_M3AC1"
                Class_Exam.examLgbn = "POWERPOINT2003"
                Class_Exam.examSgbn = "CORE"
            Case "Microsoft Office Excel 2003"
                Class_Exam.bookID = "2010_M3EC1"
                Class_Exam.examLgbn = "EXCEL2003"
                Class_Exam.examSgbn = "CORE"
            Case "Microsoft Office Excel 2003 Expert"
                Class_Exam.bookID = "2010_M3EE1"
                Class_Exam.examLgbn = "EXCEL2003"
                Class_Exam.examSgbn = "EXPERT"
            Case "Microsoft Office PowerPoint 2003"
                Class_Exam.bookID = "2010_M3PC1"
                Class_Exam.examLgbn = "POWERPOINT2003"
                Class_Exam.examSgbn = "CORE"
            Case "Microsoft Office Word 2003"
                Class_Exam.bookID = "2010_M3WC1"
                Class_Exam.examLgbn = "WORD2003"
                Class_Exam.examSgbn = "CORE"
            Case "Microsoft Office Word 2003 Expert"
                Class_Exam.bookID = "2010_M3WE1"
                Class_Exam.examLgbn = "WORD2003"
                Class_Exam.examSgbn = "EXPERT"
        End Select

        Button_Continue.Enabled = True
    End Sub


    Private Sub ComboBox_ExamIdTitle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_ExamIdTitle.SelectedIndexChanged

        ' 글로벌 변수 설정
        Dim examIdTitle As String = ComboBox_ExamIdTitle.Text
        Select Case examIdTitle
            Case "모의고사 1회"
                Class_Exam.examID = "E1"
            Case "모의고사 2회"
                Class_Exam.examID = "E2"
            Case "모의고사 3회"
                Class_Exam.examID = "E3"
            Case "모의고사 4회"
                Class_Exam.examID = "E4"
            Case "모의고사 5회"
                Class_Exam.examID = "E5"
            Case "최종고사 1회"
                Class_Exam.examID = "G1"
            Case "최종고사 2회"
                Class_Exam.examID = "G2"
            Case "최종고사 3회"
                Class_Exam.examID = "G3"
            Case "최종고사 4회"
                Class_Exam.examID = "G4"
            Case "최종고사 5회"
                Class_Exam.examID = "G5"
        End Select
    End Sub
End Class