Imports Excel = Microsoft.Office.Interop.Excel
Imports Ppt = Microsoft.Office.Interop.PowerPoint
Imports Wrd = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices


Public Class Form_Navigator
    Public Shared sWorkNewFullPathName As String
    Public Shared sWorkTplFullPathName As String
    Dim sWorkSrcFullPathName As String
    Dim sWorkAttFullPathName As String
    Dim sWorkQstFullPathName As String

    Private AB As New frmTaskBarSettings.TBAppBar 'AppBar Object

    Dim iCur_Qst_Seq As Integer
    Dim oQ

    'Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer
    Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Const SWP_NOSIZE As Int32 = &H1
    Const SWP_NOMOVE As Int32 = &H2
    Const HWND_TOPMOST As Short = -1

 

    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Int32, _
        ByVal Y As Int32, ByVal cx As Int32, ByVal cy As Int32, ByVal uFlags As Int32) As Boolean
    End Function

    '==========================================================================
    ' Navigation 화면 위치 및 크기 초기화
    '==========================================================================
    Private Sub navigation_init()

        ' 작업 표시줄 자동 숨김
        AB.AlwaysOnTop = False

        ' 화면 크기 및 위치
        Me.Width = Screen.PrimaryScreen.WorkingArea.Width
        Me.Left = 0
        Me.Top = Screen.PrimaryScreen.WorkingArea.Height - Me.Height

        ' 문제 지문 상자 크기 및 위치
        RichTextBox_Question.Left = 0
        RichTextBox_Question.Top = 0
        RichTextBox_Question.Width = Me.Width

        ' 종료 버튼 위치
        Button_End.Left = (GroupBox4.Width - Button_End.Width) / 2

        ' Always Top
        SetWindowPos(Me.Handle, HWND_TOPMOST, 0, 0, 0, 0, 3)

        ' 색
        Me.BackColor = ColorTranslator.FromOle(RGB(0, 73, 156))

    End Sub

    '==========================================================================
    ' Office 프로그램 화면 위치 및 크기 초기화
    '==========================================================================
    'Private Sub office_init()

    '    Static bCalibrated As Boolean = False               ' 최초 조정 여부
    '    Static dblWidth As Double, dblHeight As Double      ' 설정될 크기

    '    If (bCalibrated <> True) Then
    '        Dim cDesiredWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
    '        Dim cDesiredHeight As Integer = Me.Top
    '        'Dim cDesiredHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
    '        Dim lngSystemWidth, lngSystemHeight
    '        Dim dblWidthRatio As Double, dblHeightRatio As Double
    '        'Dim dblMaxWidth As Double, dblMaxHeight As Double

    '        oQ.oPpt.WindowState = Ppt.PpWindowState.ppWindowMaximized

    '        Const SM_CXSCREEN = 0, SM_CYSCREEN = 1
    '        lngSystemWidth = GetSystemMetrics(SM_CXSCREEN)
    '        lngSystemHeight = GetSystemMetrics(SM_CYSCREEN)

    '        cDesiredHeight = cDesiredHeight + (lngSystemHeight - Screen.PrimaryScreen.WorkingArea.Height)

    '        dblWidthRatio = (oQ.oPpt.Width - 5) / lngSystemWidth
    '        dblHeightRatio = oQ.oPpt.Height / lngSystemHeight
    '        dblWidth = cDesiredWidth * dblWidthRatio
    '        dblHeight = cDesiredHeight * dblHeightRatio

    '        bCalibrated = True
    '    End If

    '    oQ.oPpt.WindowState = Ppt.PpWindowState.ppWindowNormal
    '    oQ.oPpt.Width = dblWidth
    '    oQ.oPpt.Height = dblHeight

    '    oQ.oPpt.Left = 0   ' 왼쪽
    '    oQ.oPpt.Top = 0    ' 상단에

    '    oQ.oPpt.Visible = True

    'End Sub

    '==========================================================================
    ' 파일 초기화
    '==========================================================================
    Function file_init() As Boolean

        Dim sSubFolder As String = Class_Exam.bookID & "\" & Class_Exam.examID & "\" & Format(oQ.iCurrentQuestion, "00")
        Dim sSrcFullPathName = System.Windows.Forms.Application.StartupPath & "\data\" & sSubFolder & "\" & oQ.srcFile
        Dim sAttFullPathName = System.Windows.Forms.Application.StartupPath & "\data\" & sSubFolder & "\" & oQ.attFile
        sWorkSrcFullPathName = Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\" & oQ.srcFile
        sWorkAttFullPathName = Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\" & oQ.attFile
        sWorkNewFullPathName = Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\" & oQ.newFile
        sWorkQstFullPathName = System.Windows.Forms.Application.StartupPath & "\data\" & sSubFolder & "\" & oQ.qstFile
        sWorkTplFullPathName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Templates\" & oQ.tplFile

        '----------------------------------------------------------------------
        ' 작업 파일 준비
        '----------------------------------------------------------------------
        If (oQ.srcFile.Length > 0) Then
            ' 1. 기존 파일 삭제
            Try
                If (My.Computer.FileSystem.FileExists(sWorkSrcFullPathName) = False) Then
                    Exit Try
                End If
                My.Computer.FileSystem.DeleteFile(sWorkSrcFullPathName)
            Catch ex As Exception
                MsgBox("작업 파일을 초기화할 수 없습니다.file_init(),삭제오류")
                Return False
            End Try

            ' 2. 내문서로 복사
            Try
                If (My.Computer.FileSystem.FileExists(sSrcFullPathName) = False) Then
                    Exit Try
                End If
                My.Computer.FileSystem.CopyFile(sSrcFullPathName, sWorkSrcFullPathName, _
                                                FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
            Catch ex As Exception
                MsgBox("작업 파일을 초기화할 수 없습니다.file_init(),복사오류")
                Return False
            End Try
        End If

        '----------------------------------------------------------------------
        ' 추가 작업 파일 준비
        '----------------------------------------------------------------------
        If (oQ.attFile.Length > 0) Then
            ' 1. 기존 파일 삭제
            Try
                If (My.Computer.FileSystem.FileExists(sWorkAttFullPathName) = False) Then
                    Exit Try
                End If
                My.Computer.FileSystem.DeleteFile(sWorkAttFullPathName)
            Catch ex As Exception
                MsgBox("추가 작업 파일을 초기화할 수 없습니다.file_init(),삭제오류")
                Return False
            End Try

            ' 2. 내문서로 복사
            Try
                My.Computer.FileSystem.CopyFile(sAttFullPathName, sWorkAttFullPathName, _
                                                FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
            Catch ex As Exception
                MsgBox("추가 작업 파일을 초기화할 수 없습니다.file_init(),복사오류")
                Return False
            End Try
        End If

        '----------------------------------------------------------------------
        ' 생성될 결과 파일 준비(삭제)
        '----------------------------------------------------------------------
        If (oQ.newFile.Length > 0) Then
            ' 1. 기존 추가 파일 삭제
            Try
                If (My.Computer.FileSystem.FileExists(sWorkNewFullPathName) = False) Then
                    Exit Try
                End If
                My.Computer.FileSystem.DeleteFile(sWorkNewFullPathName)
            Catch ex As Exception
                MsgBox("생성될 결과 파일을 초기화할 수 없습니다.file_init(),삭제오류")
                Return False
            End Try
        End If

        '----------------------------------------------------------------------
        ' 생성될 템플릿 파일 준비(삭제)
        '----------------------------------------------------------------------
        If (oQ.tplFile.Length > 0) Then
            ' 1. 기존 추가 파일 삭제
            Try
                If (My.Computer.FileSystem.FileExists(sWorkTplFullPathName) = False) Then
                    Exit Try
                End If
                My.Computer.FileSystem.DeleteFile(sWorkTplFullPathName)
            Catch ex As Exception
                MsgBox("생성될 템플릿 파일을 초기화할 수 없습니다.file_init(),삭제오류")
                Return False
            End Try
        End If

        Return True

    End Function

    '==========================================================================
    ' 서버로 Next 전문 송신
    '==========================================================================
    Sub serverProcNext()
        Class_Server.cmdNext(Class_Exam.bookID, Class_Exam.examID, Class_Exam.userID, _
                    Class_Exam.examLgbn, Class_Exam.examSgbn, oQ.iCurrentQuestion, _
                    oQ.sWrongComment, oQ.iRealScore, second2minute(oQ.iElapsedTime))
    End Sub

    '==========================================================================
    ' 서버로 Skip 전문 송신
    '==========================================================================
    Sub serverProcSkip()
        Class_Server.cmdSkip(Class_Exam.bookID, Class_Exam.examID, Class_Exam.userID, _
                    Class_Exam.examLgbn, Class_Exam.examSgbn, oQ.iCurrentQuestion, second2minute(oQ.iElapsedTime))
    End Sub

    '==========================================================================
    ' 서버로 Start 전문 송신
    '==========================================================================
    Sub serverProcStart()
        Class_Server.cmdStart(Class_Exam.bookID, Class_Exam.examID, Class_Exam.userID, _
                    Class_Exam.examLgbn, Class_Exam.examSgbn)
    End Sub

    '==========================================================================
    ' 서버로 상태 요청 전문 송신
    '==========================================================================
    Sub serverRequest()
        Class_Server.requestStat(Class_Exam.bookID, Class_Exam.examID, Class_Exam.userID, _
                         Class_Exam.examLgbn, Class_Exam.examSgbn)
    End Sub

    '==========================================================================
    ' 문제 지문 파일 열기
    '==========================================================================
    Private Sub question_fill()
        RichTextBox_Question.LoadFile(sWorkQstFullPathName)
    End Sub

    '==========================================================================
    ' Office 파일 열기
    '==========================================================================
    Private Sub office_file_open()

        ' 문제 자료 초기화
        oQ.questionNo_init(oQ.iCurrentQuestion)

        ' 파일 초기화
        If (file_init() <> True) Then
            MsgBox("Error file_init()")
            End
        End If

        ' 문제 채우기
        question_fill()

        ' 문제 번호
        Label_Seq.Text = oQ.iTotalQuestion.ToString() & " 중 " & oQ.iCurrentQuestion.ToString()

        ' 오피스 파일 열기
        oQ.file_open(sWorkSrcFullPathName)

        ' 오피스 화면 최기화
        oQ.office_screen_init(Me)

    End Sub

    '==========================================================================
    ' 시험 종료 처리
    '==========================================================================
    Private Sub exam_quit()
        oQ.office_file_close()
        oQ.office_quit()

        ' 껍질이 남았음...살아있는 개체가 있는 것 같음...???

        Form_Opinion.Show()
        Me.Hide()
    End Sub

    '==========================================================================
    ' Debug Msg 
    '==========================================================================
    Private Sub debug(ByVal msg As String)
        If (Class_Exam.userID = "mrbr" Or Class_Exam.userID = "ciz2000") Then
            Form_Debug.Show()
            Form_Debug.Log(msg)
        End If
    End Sub

    '==========================================================================
    ' 종료 닫추
    '==========================================================================
    Private Sub Button_End_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_End.Click
        MsgBox("실제 시험에는 종료 기능이 없습니다.")

        Try
            oQ.office_file_close()
            oQ.office_quit()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

        Catch ex As Exception
        End Try

        End

    End Sub

    '==========================================================================
    ' 네비게이션 화면 이동시 원복
    '==========================================================================
    Private Sub Form_Navigator_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (e.CloseReason = CloseReason.UserClosing) OrElse (e.CloseReason = CloseReason.None) Then
            e.Cancel = True
        End If
    End Sub

    '==========================================================================
    ' 다음 단추
    '==========================================================================
    Private Sub Button_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Next.Click

        oQ.iRealScore = 0
        oQ.sWrongComment = ""

        ' skip된 문제 풀어야 됨. ???
        Select Case oQ.iCurrentQuestion
            Case 1 : Call oQ.question_examinate01()
            Case 2 : Call oQ.question_examinate02()
            Case 3 : Call oQ.question_examinate03()
            Case 4 : Call oQ.question_examinate04()
            Case 5 : Call oQ.question_examinate05()
            Case 6 : Call oQ.question_examinate06()
            Case 7 : Call oQ.question_examinate07()
            Case 8 : Call oQ.question_examinate08()
            Case 9 : Call oQ.question_examinate09()
            Case 10 : Call oQ.question_examinate10()
            Case 11 : Call oQ.question_examinate11()
            Case 12 : Call oQ.question_examinate12()
            Case 13 : Call oQ.question_examinate13()
            Case 14 : Call oQ.question_examinate14()
            Case 15 : Call oQ.question_examinate15()
            Case 16 : Call oQ.question_examinate16()
            Case 17 : Call oQ.question_examinate17()
            Case 18 : Call oQ.question_examinate18()
            Case 19 : Call oQ.question_examinate19()
            Case 20 : Call oQ.question_examinate20()
            Case Else : MsgBox("문제 번호 선택 오류") : End
        End Select

        ' 기존 문제 파일 닫기
        oQ.office_file_close()

        ' 글로벌 변수 설정
        oQ.iRealtotalScore = oQ.iRealtotalScore + oQ.iRealScore
        Class_Exam.iTotalScore = oQ.iRealtotalScore

        ' 점수 출력 디버깅용
        debug("총점: " & oQ.iRealtotalScore & " 점수: " & oQ.iRealScore)

        '----------------------------------------------------------------------
        ' Next 전문 전송
        '----------------------------------------------------------------------
        Call serverProcNext()

        ' 다음 문제 열기, 최종 문제인지 확인 필요
        If (oQ.iCurrentQuestion >= oQ.iTotalQuestion) Then
            exam_quit()
            Exit Sub
        End If

        oQ.iCurrentQuestion = oQ.iCurrentQuestion + 1
        Call office_file_open()

    End Sub

    '==========================================================================
    ' 건너뛰기 단추
    '==========================================================================
    Private Sub Button_Skip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Skip.Click

        ' 임시로
        MsgBox("실제 시험에선 건너뛰기 기능이 가능합다.")
        Exit Sub

        oQ.office_file_close()

        ' 다음 문제 열기, 최종 문제인지 확인 필요??? skip한 문제가 나타나야 한다.
        If (oQ.iCurrentQuestion >= oQ.iTotalQuestion) Then
            Exit Sub
        End If

        '----------------------------------------------------------------------
        ' Skip 전문 전송 
        '----------------------------------------------------------------------
        Call serverProcSkip()

        oQ.iCurrentQuestion = oQ.iCurrentQuestion + 1
        Call office_file_open()

    End Sub

    '==========================================================================
    ' 다시 단추
    '==========================================================================
    Private Sub Button_Retry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Retry.Click
        oQ.office_file_close()
        office_file_open()
    End Sub

    '==========================================================================
    ' 타이머 동작
    '==========================================================================
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim iCountdown As Integer

        ' 경과 시간 증가
        oQ.iElapsedTime = oQ.iElapsedTime + 1

        ' 남은 시간 감소
        iCountdown = 50 * 60 - oQ.iElapsedTime

        Label_Timer.Text = second2minute(iCountdown)

        ' 남은 시간이 없으면 종료 처리
        If (iCountdown <= 0) Then
            Timer1.Enabled = False
            oQ.iCurrentQuestion = oQ.iTotalQuestion
            serverProcNext()
            exam_quit()
        End If
    End Sub

    '==========================================================================
    ' 초를 "hh:mm:ss"로 바꾼다.
    '==========================================================================
    Function second2minute(ByVal iSecond As Integer) As String

        Dim temp As String
        temp = Format(Int(iSecond / 60 / 60), "00")
        temp = temp & ":" & Format(Int(iSecond / 60), "00")
        temp = temp & ":" & Format(iSecond Mod 60, "00")
        Return temp

    End Function

    '==========================================================================
    ' 화면 시작 : 프로그램 시작
    '==========================================================================
    Private Sub Form_Navigator_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '----------------------------------------------------------------------
        ' 파일 준비
        '----------------------------------------------------------------------
        Call file_ready(System.Windows.Forms.Application.StartupPath & "\data\2010_M3PC1.zip", _
                        System.Windows.Forms.Application.StartupPath & "\data")

        '----------------------------------------------------------------------
        ' Start 전문 전송
        '----------------------------------------------------------------------
        Call serverProcStart()

        '----------------------------------------------------------------------
        ' 화면 초기화
        '----------------------------------------------------------------------
        Call navigation_init()

        '----------------------------------------------------------------------
        ' 시험 시작
        '----------------------------------------------------------------------
        Call exam_init()

        'Call oQ.office_file_close()
        Call office_file_open()


    End Sub
    '==========================================================================
    ' 파일 준비 (압축풀기)
    '==========================================================================
    Private Sub file_ready(ByVal strSrcPath, ByVal strDestPath)
        Try
            'Dim ShellAppType As Type = Type.GetTypeFromProgID("Shell.Application")
            'Dim oShell As Object = Activator.CreateInstance(ShellAppType)
            'Dim SrcFlder As Object = oShell.NameSpace(strSrcPath.ToString)
            'Dim DestFlder As Object = oShell.NameSpace(strDestPath.ToString)
            'Dim items As Object = SrcFlder.Items()

            Dim sc As Shell32.ShellClass = New Shell32.ShellClass()
            Dim SrcFlder As Shell32.Folder = sc.NameSpace(strSrcPath)
            Dim DestFlder As Shell32.Folder = sc.NameSpace(strDestPath)
            Dim items As Shell32.FolderItems = SrcFlder.Items()
            DestFlder.CopyHere(items, 20)


        Catch ex As Exception
            MsgBox("데이터 파일 준비 오류입니다." & vbCrLf & _
                    ex.Message & vbCrLf & _
                   "src file=" & strSrcPath & vbCrLf & _
                   "dst file=" & strDestPath)

            If (My.Computer.FileSystem.FileExists(strSrcPath) = True) Then
                MsgBox("Zip File Ok")
            Else
                MsgBox("Zip File Error")
            End If

        End Try

    End Sub
    '==========================================================================
    ' 시험 관련 초기화
    '==========================================================================
    Private Sub exam_init()


        '----------------------------------------------------------------------
        ' 로그인 화면, 시험 선택 화면에서 설정했음, 디버깅용으로 강제 설정함
        '----------------------------------------------------------------------
        ''Class_Exam.userID = "ciz2000"

        'Class_Exam.userID = "mrbr"
        'Class_Exam.bookID = "2010_M3PC1"
        'Class_Exam.examID = "G1"
        'Class_Exam.examLgbn = "POWERPOINT2003"
        'Class_Exam.examSgbn = "CORE"

        Class_Exam.userID = "mrbr"
        Class_Exam.bookID = "2010_M3WE1"
        Class_Exam.examID = "E1"
        Class_Exam.examLgbn = "WORD2003"
        Class_Exam.examSgbn = "EXPERT"

        '----------------------------------------------------------------------
        ' 점수 기준 자체 점검
        '----------------------------------------------------------------------
        Select Case Class_Exam.bookID
            Case "2010_M3WE1"
                Select Case Class_Exam.examID
                    Case "E1"
                        oQ = New Class_Q_WORD1
                    Case Else
                        MsgBox("Class 생성 오류")
                        End
                End Select
            Case "2010_M3PC1"
                Select Case Class_Exam.examID
                    Case "E1"
                        oQ = New Class_Q_PPT1
                    Case "G1"
                        oQ = New Class_Q_PPT2
                    Case Else
                        MsgBox("Class 생성 오류")
                        End
                End Select
        End Select
        

        Dim i As Integer
        Dim iTotal As Integer
        For i = 1 To oQ.iTotalQuestion
            oQ.questionNo_init(i)
            iTotal = iTotal + oQ.iSubScore1 + oQ.iSubScore2
        Next
        If iTotal <> 1000 Then
            MsgBox("채점 기준 오류")
            End
        End If

        '----------------------------------------------------------------------
        ' 글로벌 변수 설정
        '----------------------------------------------------------------------
        Class_Exam.iPassScore = oQ.iPassScore

        '----------------------------------------------------------------------
        ' 서버에서 시험 진행 상태값 가져와 설정
        '----------------------------------------------------------------------
        Call serverRequest()
        If (Class_Server.r_curqstseq <> Nothing) Then
            'oQ.iCurrentQuestion = Class_Server.r_curqstseq + 1  ' 현재 진행할 문제 번호 : 주의
            oQ.iCurrentQuestion = 8  ' 현재 진행할 문제 번호 : 주의
            oQ.iRetakeCnt = Class_Server.r_retakecnt            ' 재시험 횟수
            oQ.setElapsedTime(Class_Server.r_elapsedtime)       ' 경과시간 
        Else
            oQ.iCurrentQuestion = 1
            oQ.iRetakeCnt = 0
            oQ.iElapsedTime = 0
        End If

        ' 임시 디버깅용
        'oQ.iCurrentQuestion = 17

        '----------------------------------------------------------------------
        ' 재시험 횟수 검사
        '----------------------------------------------------------------------
        If (oQ.iRetakeCnt >= 10) Then
            MsgBox("재시험을 5회 이상하셨습니다. 부정행위 방지를 위하여 더이상 시험 진행이 되지 않습니다.")
            End
        End If

        '----------------------------------------------------------------------
        ' 현재 진행중인 문제 번호 검사
        '----------------------------------------------------------------------
        If (oQ.iCurrentQuestion > oQ.iTotalQuestion) Then
            MsgBox("해당 회차 시험의 모든 문제를 다 풀었습니다.")
            End
        End If


        '----------------------------------------------------------------------
        ' 파일 닫기
        '----------------------------------------------------------------------
        oQ.office_file_close()

        '----------------------------------------------------------------------
        ' Start 전문 전송 
        '----------------------------------------------------------------------
        serverProcStart()

        '----------------------------------------------------------------------
        ' Timer 동작, Countdown Start
        '----------------------------------------------------------------------
        Timer1.Enabled = True

    End Sub

    Private Sub Form_Navigator_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        navigation_init()
    End Sub
End Class