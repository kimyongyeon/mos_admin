
Imports Ppt = Microsoft.Office.Interop.PowerPoint
Imports Wrd = Microsoft.Office.Interop.Word

Public Class Class_Q
    Public Const officeRate As Double = 0.0353  ' 도형의 크기 변경을 위한 비율               

    Public l_gbn As String                      ' 대구분
    Public s_gbn As String                      ' 소구분

    Public iTotalQuestion As Integer            ' 총 문제수
    Public iCurrentQuestion As Integer          ' 현재 진행중인 문제 번호
    Public aiSkipQuestionCount As Integer       ' 건너뛰기한 문제 개수
    Public aiSkipQuestionList(20) As Integer    ' 건너뛰기한 문제 번호 목록

    Public iPassScore As Integer                ' 합격 기준 점수
    Public iSubScore1 As Integer                ' 서브 1번 문항 기준 점수
    Public iSubScore2 As Integer                ' 서브 2번 문항 기준 점수
    Public iRealScore As Integer                ' 한 문제당 채점된 실제 점수
    Public iRealtotalScore As Integer           ' 채점된 실제 총점수
    Public sWrongComment As String              ' 틀린이유

    Public iRetakeCnt As Integer                ' 재시험 횟수
    Public iElapsedTime As Integer              ' 경과시간
    Public oStartTime As Date                   ' 시작시각
    '                                             파워포인트용
    Public oPpt As Ppt.ApplicationClass         ' 파워포인트 프로그램
    Public oPre As Ppt.Presentation             ' 파워포인트 파일
    Public oPres As Ppt.Presentations           ' 파워포인트 파일들
    Public oSlide As Ppt.Slide                  ' 파워포인트 슬라이드
    Public oSlides As Ppt.Slides                ' 파워포인트 슬라이드들
    Public oSlideShowWindow As Ppt.SlideShowWindow ' 파워포인트 슬아이드쇼창
    Public oSlideShowWindows As Ppt.SlideShowWindows ' 파워포인트 슬아이드쇼창들

    Public oWrd As Wrd.ApplicationClass
    Public oDoc As Wrd.Document
    Public oDocs As Wrd.Documents

    Public srcFile As String                    ' 문제에서 사용할 원본 파일(거의 필수)
    Public attFile As String                    ' 문제에서 사용할 추가 파일(선택 사항, 병합등에서 사용)
    Public fnlFile As String                    ' 문제에서 수정된 최종 결과 파일(나중에 사용 예정)
    Public newFile As String                    ' 문제에서 새롭게 생성될 결과 파일(선택 사항)
    Public tplFile As String                    ' 문제에서 새롭게 생성될 템플릿 파일(선택 사항)
    Public qstFile As String                    ' 문제 지문 파일

    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer

    Public Sub addWrongComment(ByVal sReason As String)
        sWrongComment = sWrongComment & sReason
    End Sub

    Public Sub setElapsedTime(ByVal sTime As String)
        Dim arrValue As String() = sTime.Split(New Char() {":"})

        Select Case arrValue.Length
            Case 0
                iElapsedTime = 50 * 60
            Case 2
                iElapsedTime = Integer.Parse(arrValue(0)) * 60 + Integer.Parse(arrValue(1))
            Case 3
                iElapsedTime = Integer.Parse(arrValue(0)) * 60 * 60 + Integer.Parse(arrValue(1)) * 60 + Integer.Parse(arrValue(2))
        End Select

    End Sub

    '==========================================================================
    ' Office 파일 열기
    '==========================================================================
    Public Sub file_open(ByVal sFileName As String)

        'Start Powerpoint and open the workbook.
        Select Case l_gbn
            Case "POWERPOINT"
                oPpt = CreateObject("PowerPoint.Application")
                oPpt.Visible = True
                oPres = oPpt.Presentations

                If (srcFile.Length > 0) Then
                    Try
                        oPre = oPres.Open(sFileName)
                    Catch ex As Exception
                        MsgBox("파일을 열 수 없습니다")
                        Exit Sub
                    End Try
                End If
            Case "WORD"
                oWrd = CreateObject("Word.Application")
                oWrd.Visible = True
                oDocs = oWrd.Documents
                If (srcFile.Length > 0) Then
                    Try
                        oDoc = oDocs.Open(sFileName)
                    Catch ex As Exception
                        MsgBox("파일을 열 수 없습니다")
                        Exit Sub
                    End Try
                End If

        End Select

        'office_init()

    End Sub

    '==========================================================================
    ' Office 파일 닫기
    '==========================================================================
    Public Sub office_file_close()

        Dim i As Integer

        Try
            Select Case l_gbn
                Case "POWERPOINT"
                    If oPpt Is Nothing Then oPpt = CreateObject("PowerPoint.Application")
                    For i = 1 To oPpt.Presentations.Count
                        oPre = oPpt.Presentations(i)
                        oPre.Close()
                        NAR(oPre)
                    Next
                    NAR(oPres)
                Case "WORD"
                    ' 왜 않되나?  ppt는 잘되는 데....쩝
                    If oWrd Is Nothing Then oWrd = CreateObject("Word.Application")
                    For i = 1 To oWrd.Documents.Count
                        oDoc = oWrd.Documents(i)
                        oDoc.Close()
                        NAR(oDoc)
                    Next
                    NAR(oDoc)
                    oWrd.Quit()
            End Select

        Catch e As Exception
            'Try
            '    Select Case l_gbn
            '        Case "POWERPOINT"
            '            For i = 1 To oPpt.Presentations.Count
            '                oPre = oPpt.Presentations(i)
            '                oPre.Close()
            '                NAR(oPre)
            '            Next
            '            NAR(oPres)
            '        Case "WORD"

            '    End Select
            'Catch ex As Exception

            'End Try

        End Try

    End Sub

    Private Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    '==========================================================================
    ' Office 종료하기
    '==========================================================================
    Public Sub office_quit()

        Try
            Select Case l_gbn
                Case "POWERPOINT"
                    oPpt.Quit()
                Case "WORD"
                    oWrd.Quit()
            End Select

        Catch ex As Exception
            MsgBox("Office 프로그램을 종료할 수 없습니다.office_quit() 오류")
        End Try

    End Sub

    '==========================================================================
    ' Office 프로그램 화면 위치 및 크기 초기화
    '==========================================================================
    'Private Sub office_screen_init(ByVal my As Object)
    Public Sub office_screen_init(ByVal my As Form)

        Static bCalibrated As Boolean = False               ' 최초 조정 여부
        Static dblWidth As Double, dblHeight As Double      ' 설정될 크기

        Dim oApplication As Object
        oApplication = oPpt
        Select Case l_gbn
            Case "POWERPOINT"
                oApplication = oPpt
            Case "WORD"
                oApplication = oWrd
        End Select

        If (bCalibrated <> True) Then
            Dim cDesiredWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
            'Dim cDesiredHeight As Integer = Me.Top
            Dim cDesiredHeight As Integer = my.Top
            'Dim cDesiredHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
            Dim lngSystemWidth, lngSystemHeight
            Dim dblWidthRatio As Double, dblHeightRatio As Double
            'Dim dblMaxWidth As Double, dblMaxHeight As Double

            'oPpt.WindowState = Ppt.PpWindowState.ppWindowMaximized
            Select Case l_gbn
                Case "POWERPOINT"
                    oApplication.WindowState = Ppt.PpWindowState.ppWindowMaximized
                Case "WORD"
                    oApplication.WindowState = Wrd.WdWindowState.wdWindowStateMaximize
            End Select


            Const SM_CXSCREEN = 0, SM_CYSCREEN = 1
            lngSystemWidth = GetSystemMetrics(SM_CXSCREEN)
            lngSystemHeight = GetSystemMetrics(SM_CYSCREEN)

            cDesiredHeight = cDesiredHeight + (lngSystemHeight - Screen.PrimaryScreen.WorkingArea.Height)

            dblWidthRatio = (oApplication.Width - 5) / lngSystemWidth
            dblHeightRatio = oApplication.Height / lngSystemHeight
            dblWidth = cDesiredWidth * dblWidthRatio
            dblHeight = cDesiredHeight * dblHeightRatio

            bCalibrated = True
        End If

        'oPpt.WindowState = Ppt.PpWindowState.ppWindowNormal
        'oPpt.Width = dblWidth
        'oPpt.Height = dblHeight

        'oPpt.Left = 0   ' 왼쪽
        'oPpt.Top = 0    ' 상단에

        'oPpt.Visible = True

        Select Case l_gbn
            Case "POWERPOINT"
                oApplication.WindowState = Ppt.PpWindowState.ppWindowNormal
            Case "WORD"
                oApplication.WindowState = Wrd.WdWindowState.wdWindowStateNormal
        End Select

        oApplication.Width = dblWidth
        oApplication.Height = dblHeight

        oApplication.Left = 0   ' 왼쪽
        oApplication.Top = 0    ' 상단에

        oApplication.Visible = True

    End Sub

End Class
