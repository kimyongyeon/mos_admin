Imports System.Deployment.Application

Public Class Form_Login

    ' TODO: 제공된 사용자 이름과 암호를 사용하여 사용자 지정 인증을 수행하는 코드를 삽입합니다
    ' (자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=35339 참조).  
    ' 그러면 사용자 지정 보안 주체가 현재 스레드의 보안 주체에 다음과 같이 첨부될 수 있습니다. 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' 여기서 CustomPrincipal은 인증을 수행하는 데 사용되는 IPrincipal이 구현된 것입니다. 
    ' 나중에 My.User는 CustomPrincipal 개체에 캡슐화된 사용자 이름, 표시 이름 등의
    ' ID 정보를 반환합니다.

    Private Function convertUnicode(ByVal word As Char) As Boolean
        Dim ascii As Integer = Convert.ToInt32(word)
        If (ascii >= 41 And ascii <= 126) Then
            Return (True)
        Else
            Return (False)
        End If
    End Function

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        Class_Exam.userId = UsernameTextBox.Text

        ' 2010.01.17 : SCS 
        Call update_proc()

        ' 2010.01.20 : SCS
        Call login_proc()

        ' 2010.01.17 : SCS
        'MsgBox("시뮬레이션 프로그램은 2010년 1월 21일부터 사용가능합니다.")
    End Sub

    Private Sub Button_AdminLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AdminLogin.Click
        login_proc()
    End Sub

    Private Sub login_proc()
        If UsernameTextBox.Text.Length <= 0 Then
            MsgBox("사용자 이름을 입력하세요")
            Exit Sub
        End If

        If PasswordTextBox.Text.Length <= 0 Then
            MsgBox("암호를 입력하세요")
            Exit Sub
        End If

        If ComboBox_ExamType.Text.Length <= 0 Then
            MsgBox("시험 유형을 선택하세요")
            Exit Sub
        End If

        Dim correctWord As Boolean = False
        Dim character As Char() = PasswordTextBox.Text.ToCharArray()
        For i As Integer = 0 To character.Length - 1
            correctWord = convertUnicode(character(i))
            If (correctWord = False) Then
                Exit For
            End If
        Next

        If correctWord <> True Then
            MessageBox.Show("한글은 이용하실 수 없습니다. 숫자,영문자,특수문자만 가능합니다." & vbCrLf _
                            & "홈페이지에서 비밀번호를 바뀌세요.")
            Exit Sub
        End If

        Try
            Dim url As String = ("http://www.academysoft.kr/apps/exam/exam_login_check_proc.php?user_id=" & UsernameTextBox.Text & "&password=") & PasswordTextBox.Text
            Dim request As System.Net.HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(url), System.Net.HttpWebRequest)
            Dim response As System.Net.HttpWebResponse = DirectCast(request.GetResponse(), System.Net.HttpWebResponse)
            Dim reader As New System.IO.StreamReader(response.GetResponseStream(), System.Text.Encoding.GetEncoding(response.CharacterSet), True)
            Dim str As String = reader.ReadToEnd()
            If str <> "return:ok" Then
                MsgBox("사용자 이름 또는 비밀번호가 틀렸습니다")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("로그인 할 수 없습니다." & vbCrLf & ex.Message)
            Exit Sub
        End Try

        '----------------------------------------------------------------------
        ' 시험 선택 화면
        '----------------------------------------------------------------------
        Form_ChooseExam.Show()

        ' 2010.01.18 : SCS : imsi
        Me.Hide()

        '----------------------------------------------------------------------
        ' 키보드 핫키 lock
        '----------------------------------------------------------------------
        HookKeyboard()
    End Sub
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Sub update_proc()
        Dim info As UpdateCheckInfo = Nothing

        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment

            Try
                info = AD.CheckForDetailedUpdate()
            Catch dde As DeploymentDownloadException
                MessageBox.Show("The new version of the application cannot be downloaded at this time. " + ControlChars.Lf + ControlChars.Lf + "Please check your network connection, or try again later. Error: " + dde.Message)
                Return
            Catch ioe As InvalidOperationException
                MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message)
                Return
            End Try

            If (Not info.UpdateAvailable) Then
                'MessageBox.Show("There are no updates available at this time.")
                Return
            End If

            Dim doUpdate As Boolean = True

            If (Not info.IsUpdateRequired) Then
                Dim dr As DialogResult = MessageBox.Show("프로그램 업데이트가 가능합니다. 지금 업데이트 하시겠습니까?", "Update Available", MessageBoxButtons.YesNo)
                If (dr = System.Windows.Forms.DialogResult.No) Then
                    doUpdate = False
                End If
            Else
                ' Display a message that the app MUST reboot. Display the minimum required version.
                MessageBox.Show("This application has detected a mandatory update from your current " & _
                    "version to version " + info.MinimumRequiredVersion.ToString() & _
                    ". The application will now install the update and restart.", _
                    "Update Available", MessageBoxButtons.OK, _
                    MessageBoxIcon.Information)
            End If

            If (doUpdate) Then
                Try
                    AD.Update()
                    MessageBox.Show("프로그램이 업그레이드 되어서, 프로그램을 다시 시작합니다.")
                    'Application.Restart() : 2010.01.17 : SCS
                    System.Windows.Forms.Application.Restart()
                Catch dde As DeploymentDownloadException
                    MessageBox.Show("Cannot install the latest version of the application. " + ControlChars.Lf + ControlChars.Lf + "Please check your network connection, or try again later.")
                    Return
                End Try
            End If
        End If
    End Sub

    
    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        'Dim wbShow As New WebBrowser
        'wbShow.Navigate("http://www.academysoft.kr")
        'Shell("cmd ""iexplore""", AppWinStyle.NormalFocus)
        MsgBox("http://www.academysoft.kr 에서 등록하십시오.")
    End Sub



    Private Sub Form_Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
