Imports Ppt = Microsoft.Office.Interop.PowerPoint

Public Class Class_Q_PPT2
    Inherits Class_Q
    Implements Interface_Q

    Sub New()
        l_gbn = "POWERPOINT"
        iTotalQuestion = 17
        iPassScore = 760
    End Sub


    Public Sub questionNo_init(ByVal iNo As Integer) Implements Interface_Q.questionNO_init

        iCurrentQuestion = iNo

        srcFile = ""
        attFile = ""
        qstFile = "question.rtf"
        tplFile = ""
        newFile = ""

        iSubScore1 = 0
        iSubScore2 = 0

        Select Case iNo
            Case 1
                srcFile = "연습1.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 2
                srcFile = "연습2.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 3
                srcFile = "연습3.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 4
                srcFile = "연습4.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 5
                srcFile = "연습5.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 6
                srcFile = "연습6.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 7
                attFile = "영업능력개발.doc"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 8
                srcFile = "연습8.ppt"
                iSubScore1 = 60

            Case 9
                srcFile = "연습9.ppt"
                attFile = "검토.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 10
                srcFile = "연습10.ppt"
                attFile = "전문지식분석.xls"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 11
                srcFile = "연습11.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 12
                srcFile = "연습12.ppt"
                tplFile = "MOS 서식 파일.pot"
                attFile = "협상.jpg"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 13
                srcFile = "연습13.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 14
                srcFile = "연습14.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 15
                srcFile = "연습15.ppt"
                iSubScore1 = 30
                iSubScore2 = 30

            Case 16
                srcFile = "연습16.ppt"
                iSubScore1 = 40

            Case 17
                newFile = "KETRi.mht"
                iSubScore1 = 30
                iSubScore2 = 30

        End Select

    End Sub

    Public Sub question_examinate20() Implements Interface_Q.question_examinate20
    End Sub

    Public Sub question_examinate19() Implements Interface_Q.question_examinate19
    End Sub

    Public Sub question_examinate18() Implements Interface_Q.question_examinate18
    End Sub

    Public Sub question_examinate03() Implements Interface_Q.question_examinate17
        ' ------------------------------------------ ---------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(2)
            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)

                ' placeholder인지 검사1
                If InStr(oShape.Name, "Rectangle") <= 0 Then  ' 2003에서 확인해야함
                    Continue For
                End If

                If (oShape.Type <> MsoShapeType.msoPlaceholder) Then
                    Exit Try
                End If

                '텍스트가 "그린"인지 확인
                If InStr(oShape.TextFrame.TextRange.Text, "아카데미소프트공학연구소") > 10 Then
                    bExamResult11 = True
                End If

            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(3)
            bExamResult21 = True
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("하이퍼링크를 만들지 않았거나 주소가 틀리거나 이름이 '더 조은 컴퓨터'가 아닙니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("애니메이션 효과가 적용되지 않았습니다.")
        End If
    End Sub
    Public Sub question_examinate16() Implements Interface_Q.question_examinate16
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(3)

            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                'If oShape.Type = MsoShapeType.msoChart And _
                If InStr(oShape.Name, "Rectangle") > 0 Then  ' 2003에서 확인해야함
                    'If oShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle Then

                    If (oShape.Type <> MsoShapeType.msoPlaceholder) Then
                        Exit Try
                    End If

                    If oShape.TextFrame.TextRange.Font.Emboss = True Then
                        bExamResult11 = True
                        Exit Try
                    End If

                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            'MsgBox("exam02-11 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore1
        Else
            'MsgBox("exam02-11 CustomDocumentProperties Error")
            addWrongComment("제목에 블록효과가 적용되지 않았습니다.")
        End If

    End Sub
    Public Sub question_examinate15() Implements Interface_Q.question_examinate15
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(5)
            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)

                ' AutoShape인지 검사1
                If InStr(oShape.Name, "AutoShape") <= 0 Then  ' 2003에서 확인해야함
                    Continue For
                End If

                ' AutoShape인지 검사2
                If (oShape.Type <> MsoShapeType.msoAutoShape) Then
                    Exit Try
                End If

                ' AutoShape중에서 되돌아가기 인지 검사
                If (oShape.AutoShapeType <> MsoAutoShapeType.msoShapeActionButtonReturn) Then
                    Exit Try
                End If

                ' ppMouseClick과 ppMouseOver 이벤트가 있는 지 확인
                If (oShape.ActionSettings.Count <> 2) Then
                    Exit Try
                End If

                Dim oAS As Ppt.ActionSetting
                oAS = oShape.ActionSettings(Ppt.PpMouseActivation.ppMouseOver)

                Dim kk
                kk = 1

                If (InStr(oAS.Hyperlink.SubAddress, "협상의 중요성") <= 0) Then
                    Exit Try
                End If


                bExamResult11 = True
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try

            Dim oSST As Ppt.SlideShowTransition
            oSST = oPpt.Presentations(1).Slides(3).SlideShowTransition

            ' 슬라이트쇼 자동 시간 검사
            If Int(oSST.AdvanceOnTime <> MsoTriState.msoTrue) Then
                Exit Try
            End If

            ' 슬라이트쇼 진행시간 검사
            If Int(oSST.AdvanceTime) <> 10 Then
                Exit Try
            End If

            bExamResult21 = True

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            'MsgBox("exam02-11 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore1
        Else
            'MsgBox("exam02-11 CustomDocumentProperties Error")
            addWrongComment("돌아가기 실행단추를 추가하지 않았거나 슬라이드의 링크가 잘못 되었습니다. ")
        End If

        If bExamResult21 Then
            'MsgBox("exam02-21 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore2
        Else
            'MsgBox("exam02-21 CustomDocumentProperties Error")
            addWrongComment("슬라이드 전환 시간이 10초로 설정되지 않았습니다.")
        End If
    End Sub
    Public Sub question_examinate14() Implements Interface_Q.question_examinate14
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            ' 2번째 슬라이드를 확인
            oSlide = oPpt.Presentations(1).Slides(2)
            Dim sText = oSlide.Shapes("Rectangle 2").TextFrame.TextRange.Text
            If (sText <> "의사소통") Then
                Exit Try
            End If

            oSlide = oPpt.Presentations(1).Slides(3)
            If oSlide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutChart Then
                bExamResult11 = True
            End If

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(3)
            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                'If oShape.Type = MsoShapeType.msoChart And _
                If InStr(oShape.Name, "Object") > 0 Then  ' 2003에서 확인해야함 

                    Dim bb
                    bb = oShape.PlaceholderFormat

                    If (bb.type = Ppt.PpPlaceholderType.ppPlaceholderChart) Then
                        bExamResult21 = True
                        Exit Try
                    End If
                    'Dim j As Integer
                    ''Dim ophs As Ppt.Placeholders
                    'Dim ophs
                    'ophs = oSlide.Shapes.Item(i)
                    'For j = 1 To ophs.Count
                    '    Dim oo
                    '    oo = ophs.itme(i)

                    '    bExamResult11 = True
                    '    'Exit Try
                    '    'End If
                    'Next
                End If
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("차트 슬라이드 레이아웃이 삽입되지 않았거나 삽입된 위치가 틀립니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("슬라이드에 차트가 추가되지 않았습니다")
        End If
    End Sub
    Public Sub question_examinate13() Implements Interface_Q.question_examinate13
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(2)

            If oSlide.Layout <> Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTwoColumnText Then
                'addWrongComment("제목 및 2단 텍스트 레이아웃이 적용되지 않았습니다")
                Exit Try
            End If

            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                If oShape.Name = "Rectangle 4" Then  ' 2003에서 확인해야함
                    If oShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle Then
                        bExamResult11 = True
                        Exit Try
                    End If
                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try

            oSlide = oPpt.Presentations(1).Slides(1)

            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            Dim ExamWidth As Double = 12.5  ' 문제 지문의 워드아트 너비
            Dim ExamHeight As Double = 5    ' 문제 지문의 워드아트 높이
            Dim oHeight As Integer          ' 파워포인트로부터 얻어온 높이
            Dim oWidth As Integer           ' 파워포인트로부터 얻어온 너비

            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "WordArt") > 0 Then

                    oHeight = Double.Parse(oShape.Height.ToString()) * officeRate * 10
                    oWidth = Double.Parse(oShape.Width.ToString()) * officeRate * 10
                    If oHeight = (ExamHeight * 10) And oWidth = (ExamWidth * 10) Then
                        bExamResult21 = True
                        Exit Try
                    End If
                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("레이아웃이 적용되지 않았거나 텍스트 고정위치가 중간이 아닙니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("워드아트 개체의 크기가 높이 5cm 또는 너비 12.5cm 가 아닙니다.")
        End If

    End Sub
    Public Sub question_examinate12() Implements Interface_Q.question_examinate12
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점     : 1번 문제 1번 지시 사항
        ' 문제     : 이미지개체를 2번 슬라이드 왼쪽에 위 구석에 삽입        '
        ' ---------------------------------------------------------------------

        Try
            oSlide = oPpt.Presentations(1).Slides(2)

            For i = 1 To oSlide.Shapes.Count
                Dim oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "Picture 4") > 0 Then  '이미지 불러오기 수정.
                    bExamResult11 = True
                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점     : 2번 문제 1번 지시 
        ' 문제     : 기본 폴더에 "명문"의 이름으로 디자인 서식파일로 슬라이드 저장
        ' 채점기준 : 
        ' ---------------------------------------------------------------------
        Try
            If (My.Computer.FileSystem.FileExists(Form_Navigator.sWorkTplFullPathName) = False) Then
                Exit Try
            Else
                bExamResult21 = True
            End If
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("협상.jpg의 이미지가 삽입되지 않았습니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("'MOS 서식 파일이 생성되지 않았거나 기본 폴더에 저장하지 않았습니다.")
        End If
    End Sub
    Public Sub question_examinate11() Implements Interface_Q.question_examinate11
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            ' 슬아이드 갯수 검사
            If oPpt.Presentations(1).Slides.Count <> 7 Then
                Exit Try
            End If


            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 3 To 4
                oPre = oPpt.Presentations(1)
                oSlide = oPpt.Presentations(1).Slides(i)

                ' 템플레이트 검사
                If oSlide.Design.Name <> "비즈니스" Then
                    Exit Try
                End If
            Next

            bExamResult11 = True
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try

            Dim osm As Ppt.Master
            osm = oPpt.Presentations(1).SlideMaster

            Dim i As Integer
            For i = 1 To osm.Shapes.Count
                Dim oShape As Ppt.Shape
                oShape = osm.Shapes(i)

                ' 마스터 제목 스타일을 찾을 때 까지
                If InStr(oShape.Name, "Rectangle") > 0 Then
                    If InStr(oShape.AlternativeText, "마스터 제목") > 0 Then
                        If oShape.TextFrame.TextRange.Font.Size = 24 And oShape.TextFrame.TextRange.Font.NameOther.ToString() = "돋움체" Then
                            bExamResult21 = True
                        End If
                    End If
                End If

            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("3,4번 슬라이드에 '비즈니스' 서식이 적용되지 않았습니다")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("1번 슬라이드의 슬라이드 마스터에 주 제목 서식을 돋움체, 24pt인지 확인하시오")
        End If

    End Sub
    Public Sub question_examinate10() Implements Interface_Q.question_examinate10
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(4)
            Dim sText = oSlide.Shapes("Rectangle 2").TextFrame.TextRange.Text
            If (sText <> "전문지식(Knowledge)") Then
                Exit Try
            End If

            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "Object") > 0 Then
                    ' 개체인지 확인
                    If (oShape.Type <> MsoShapeType.msoEmbeddedOLEObject) Then
                        Exit Try
                    End If

                    ' 아이콘인지 확인 : ???
                    If (oShape.Width > 100 Or oShape.Height > 100) Then
                        Exit Try
                    End If

                    ' 위치 확인: ???
                    bExamResult11 = True

                End If
            Next


        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(6)
            Dim oComment As Ppt.Comment
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Comments.Count
                oComment = oSlide.Comments.Item(i)
                If InStr(oComment.Text, "MOS:") And _
                    InStr(oComment.Text, "고객 신뢰(Belief)") > 0 Then  ' 2003에서 확인해야함
                    bExamResult21 = True
                End If
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("전문지식분석.xls 를 불러오지 않았거나 아이콘이 아닙니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("메모 첫머리에 'MOS:' 텍스트를 추가하지 않았거나 기존텍스트가 삭제되었습니다 ")
        End If
    End Sub
    Public Sub question_examinate09() Implements Interface_Q.question_examinate09
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' 임시사항
        bExamResult11 = True
        bExamResult21 = True

        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("검토.ppt 파일과 병합되지 않았습니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("'Gojh'가 변경한 내용이 적용되지 않았습니다.")
        End If
    End Sub
    Public Sub question_examinate08() Implements Interface_Q.question_examinate08
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' 문제 : 5번 슬라이드 즐거운여행 되세요 텍스트 1번 슬라이드로 복사
        ' 채점기준 : 1번 슬라이드의 '즐거운 여행 되세요' 텍스트 비교 후 채점
        ' ---------------------------------------------------------------------

        Try

            oSlide = oPpt.Presentations(1).Slides(2)
            For i = 1 To oSlide.Shapes.Count
                Dim oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "Rectangle") > 0 Then
                    If oShape.TextFrame.TextRange.Text = "시장 분석력" Then
                        bExamResult11 = True
                        Exit Try
                    End If
                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점     : 2번 문제 1번 지시 사항
        ' 문제     : 개요보기로 인쇄를 2매 하세요
        ' 채점기준 : 인쇄 옵션에서 개요보기와 복사횟수가 2개로 설정되어있는지 확인
        ' ---------------------------------------------------------------------
        ' 프린터 인쇄 문제는 SKIP
        'bExamResult21 = True
        'Try
        '    If (Microsoft.Office.Interop.PowerPoint.PpPrintOutputType.ppPrintOutputOutline.ToString() <> "ppPrintOutputOutline") Then
        '        Exit Try
        '    End If

        '    Dim printOption = oPpt.Presentations(1).PrintOptions
        '    If (printOption.NumberOfCopies.ToString() <> "2") Then
        '        Exit Try
        '    End If

        '    If (Microsoft.Office.Interop.PowerPoint.PpPrintOutputType.ppPrintOutputOutline.ToString() = "ppPrintOutputOutline") Then
        '        If (oPpt.Presentations(1).PrintOptions.NumberOfCopies.ToString() = "2") Then
        '            bExamResult21 = True
        '        End If
        '    End If

        'Catch e As Exception
        '    MsgBox(e.Message)
        'End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("텍스트상자가 복사되지 않았거나 원본서식이 유지되지 않았습니다..")
        End If

        'If bExamResult21 Then
        '    iRealScore = iRealScore + iSubScore2
        'Else
        '    addWrongComment("인쇄타입이 개요보기가 아니거나 2매로 인쇄하지 않았습니다.")
        'End If
    End Sub
    Public Sub question_examinate07() Implements Interface_Q.question_examinate07
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점      : 1번 문제 1번 지시 사항
        ' 문제      : 영업능력개발.doc 불러와서 프리젠테이션 작성
        ' 채점 기준 : 슬라이드의 1번 텍스트상자 "성공 전략" 비교로 점수 측정
        ' ---------------------------------------------------------------------
        Try

            If (oPpt.Presentations.Count < 1) Then
                Exit Try
            End If

            oSlide = oPpt.Presentations(1).Slides(1)
            Dim sText = oSlide.Shapes("Rectangle 2").TextFrame.TextRange.Text
            If (sText = "영업 능력 개발") Then
                bExamResult11 = True
            End If

        Catch e As Exception
            MsgBox(e.Message)

        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점     : 2번 문제 1번 지시 사항
        ' 문제     :프리젠테이션의 모든 슬라이드 배경색을 노랑색으로 변경
        ' 채점기준 : 슬라이드마스터의 배경색이 (255,255,0)인지 판별로 점수 측정
        ' ---------------------------------------------------------------------
        Try

            If (oPpt.Presentations.Count < 1) Then
                Exit Try
            End If

            oSlide = oPpt.Presentations(1).Slides(1)
            Dim bgColor = oSlide.Design.SlideMaster.Background.Fill
            If (bgColor.ForeColor.RGB = RGB(255, 255, 0)) Then
                bExamResult21 = True
            End If

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("마케팅개요.doc 파일이 아닙니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("모든 슬라이드에 노랑색을 적용하지 않았습니다.")
        End If
    End Sub
    Public Sub question_examinate06() Implements Interface_Q.question_examinate06
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try

            oSlide = oPpt.Presentations(1).Slides(1)

            Dim oShape As Ppt.Shape
            Dim bFound As Boolean = False
            Dim ExamWidth As Double = 1.5   ' 문제 지문의 도형 너비
            Dim ExamHeight As Double = 2.5  ' 문제 지문의 도형 높이
            Dim oHeight As Integer          ' 파워포인트로부터 얻어온 높이
            Dim oWidth As Integer           ' 파워포인트로부터 얻어온 너비

            Dim i As Integer
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)

                ' AutoShape인지 검사1
                If InStr(oShape.Name, "AutoShape") <= 0 Then  ' 2003에서 확인해야함
                    Continue For
                End If

                ' AutoShape인지 검사2
                If (oShape.Type <> MsoShapeType.msoAutoShape) Then
                    Exit Try
                End If

                ' AutoShape중에서 톱니 모양의 오른쪽 화살표인지 검사
                If (oShape.AutoShapeType <> MsoAutoShapeType.msoShapeNotchedRightArrow) Then
                    Exit Try
                End If

                ' 계산법 : oHeight,oWidth에 각각 officeRate값을 곱함               
                ' 위치 검사 : ???
                ' 위치 12.7cm, 세로 9.53
                For j = 1 To oSlide.Shapes.Count
                    oShape = oSlide.Shapes.Item(i)
                    oHeight = Double.Parse(oShape.Height.ToString()) * officeRate * 10
                    oWidth = Double.Parse(oShape.Width.ToString()) * officeRate * 10
                    If oHeight = (ExamHeight * 10) And oWidth = (ExamWidth * 10) Then
                        bExamResult11 = True
                        Exit Try
                    End If
                Next

            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try

            Dim oShape As Ppt.Shape
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)


                ' AutoShape인지 검사1
                If InStr(oShape.Name, "AutoShape") <= 0 Then  ' 2003에서 확인해야함
                    Continue For
                End If

                ' AutoShape인지 검사2
                If (oShape.Type <> MsoShapeType.msoAutoShape) Then
                    Exit Try
                End If

                ' AutoShape중에서 톱니 모양의 오른쪽 화살표인지 검사
                If (oShape.AutoShapeType <> MsoAutoShapeType.msoShapeNotchedRightArrow) Then
                    Exit Try
                End If

                For j = 1 To oSlide.Shapes.Count
                    oShape = oSlide.Shapes.Item(i)

                    '각도 검사
                    'If (oShape.Rotation < 270) Then
                    '    Exit Try
                    'End If
                    bExamResult21 = True

                Next

            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("화살표 모양이 톱니모양의 오른쪽 화살표 도형이 아니거나 너비와 높이가 문제와 맞지 않습니다")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("화살표 끝이 파워포인트 제목으로 향하지 않았습니다.")
        End If
    End Sub

    Public Sub question_examinate05() Implements Interface_Q.question_examinate05
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(1)
            Dim oComment As Ppt.Comment
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Comments.Count
                oComment = oSlide.Comments.Item(i)
                If InStr(oComment.Text, "MOS 파워포인트") Then
                    bExamResult11 = True
                End If
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(5)
            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)

                ' 잉크주석인지 검사2
                If InStr(oShape.Name, "Ink") > 0 Then

                    ' 잉크주석인지 검사2
                    If (oShape.Type <> MsoShapeType.msoInkComment) Then  '?? 형광색 싸인펜
                        Exit Try
                    End If

                    'Dim oCMT As Ppt.Comment
                    'oCMT = oSlide.Shapes.Item(i)


                    ' 잉크색               
                    'If (oShape.Line.ForeColor.RGB <> RGB(255, 255, 255)) Then
                    '    Exit Try
                    'End If


                    bExamResult21 = True

                End If
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("피라미드형 다이어그램을 추가하지 않았거나 배경색이 회색 -25%가 아닙니다")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("슬라이드에 회전 애니메이션이 적용되지 않았습니다")
        End If

    End Sub
    Public Sub question_examinate04() Implements Interface_Q.question_examinate04
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점     : 1번 문제 1번 지시 사항
        ' 문제     : 전체 프레젠테이션의 화면 전환 방법으로 가로 빗질과 요술봉 소리를 적용
        ' 채점기준 : 슬라이드전체에 가로빗질, 요술봉이 적용되어있는지 확인
        ' ---------------------------------------------------------------------

        Try
            Dim oSlides = oPpt.Presentations(1).Slides.Range.SlideShowTransition

            If oSlides.EntryEffect.ToString() <> "ppEffectCombHorizontal" Then
                Exit Try
            End If

            Dim soEffect = oPpt.Presentations(1).SlideMaster.SlideShowTransition.SoundEffect

            If soEffect.Name.ToString() <> "drumroll.wav" Then  ' 북소리로
                Exit Try
            End If

            If oSlides.EntryEffect.ToString() = "ppEffectCombHorizontal" Then
                If soEffect.Name.ToString() = "drumroll.wav" Then
                    bExamResult11 = True
                End If
            End If

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점     : 2번 문제 1번 지시 
        ' 문제     : 4번 슬라이드와 2번슬라이드 자리를 맞바꾸기
        '          : 여러 슬라이드 보기 상태로 두기
        ' 채점기준 : 2번째 슬라이드에서 (4번 슬라이드의 텍스트상자값을 읽어들여 비교)
        '          : 여러 슬라이드 보기 상태인지 체크
        ' ---------------------------------------------------------------------
        Try
            Dim oView = oPpt.Presentations(1).Application.SlideShowWindows.Application.ActiveWindow
            If (oView.ViewType.ToString() <> "ppViewSlideSorter") Then
                Exit Try
            End If

            oSlide = oPpt.Presentations(1).Slides(2)
            For i = 1 To oSlide.Shapes.Count
                Dim oShape = oSlide.Shapes.Item(i)
                If oShape.Name = "Rectangle 2" Then
                    If oShape.TextFrame.TextRange.Text = "협상 스타일" Then
                        bExamResult21 = True
                        Exit Try
                    End If
                End If
            Next


        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            'MsgBox("exam02-11 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore1
        Else
            'MsgBox("exam02-11 CustomDocumentProperties Error")
            addWrongComment("가로빗질효과와 북소리효과가 적용된게 아닙니다.")
        End If

        If bExamResult21 Then
            'MsgBox("exam02-21 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore2
        Else
            'MsgBox("exam02-21 CustomDocumentProperties Error")
            addWrongComment("3번 슬라이드가 이동되지 않았거나 여러 슬라이드 보기 상태가 아닙니다.")
        End If
    End Sub
    Public Sub question_examinate17() Implements Interface_Q.question_examinate03
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            If (oPpt.Presentations.Count < 1) Then
                Exit Try
            End If

            Dim oSlide = oPpt.Presentations(1).Slides(1)
            Dim oSlide1 = oPpt.Presentations(1).Slides(2)
            Dim Right1 = False
            Dim Right2 = False

            For i = 1 To oSlide.Shapes.Count
                Dim oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "Rectangle") > 0 Then
                    If oShape.TextFrame.TextRange.Text = "KETRi 서비스" Then
                        Right1 = True
                    End If
                End If
            Next

            For i = 1 To oSlide1.Shapes.Count
                Dim oShape = oSlide1.Shapes.Item(i)
                If InStr(oShape.Name, "Rectangle") > 0 Then
                    If oShape.TextFrame.TextRange.Text = "시장 요약" Then
                        Right2 = True
                    End If
                End If
            Next


            If Right1 <> False And Right2 <> False Then
                bExamResult11 = True
            End If


        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            If (My.Computer.FileSystem.FileExists(Form_Navigator.sWorkNewFullPathName) = False) Then
                Exit Try
            Else
                bExamResult21 = True
            End If
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("차트 슬라이드 레이아웃이 삽입되지 않았거나 삽입된 위치가 틀립니다.")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("슬라이드에 차트가 추가되지 않았습니다")
        End If

    End Sub
    Public Sub question_examinate02() Implements Interface_Q.question_examinate02
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(2)

            Dim oShape As Ppt.Shape

            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "WordArt") > 0 Then
                    If oShape.TextEffect.Text = "제안서 작성!" Then
                        bExamResult11 = True
                        Exit Try
                    End If
                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim i As Integer
            For i = 1 To oPpt.Presentations(1).Slides.Count

                ' 슬라이드쇼Trasition 객체 생성
                Dim oSST As Ppt.SlideShowTransition
                oSST = oPpt.Presentations(1).Slides(i).SlideShowTransition

                ' 슬라이트쇼 진행시간 검사
                If Int(oSST.AdvanceTime) <> 3 Then
                    Exit Try
                End If

            Next

            bExamResult21 = True
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            'MsgBox("exam02-11 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore1
        Else
            'MsgBox("exam02-11 CustomDocumentProperties Error")
            addWrongComment("레이아웃이 적용되지 않았거나 텍스트 고정위치가 중간이 아닙니다.")
        End If

        If bExamResult21 Then
            'MsgBox("exam02-21 CustomDocumentProperties Ok")
            iRealScore = iRealScore + iSubScore2
        Else
            'MsgBox("exam02-21 CustomDocumentProperties Error")
            addWrongComment("워드아트 개체의 크기가 높이 5cm 또는 너비 12cm 가 아닙니다.")
        End If

    End Sub

    Public Sub question_examinate01() Implements Interface_Q.question_examinate01
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(3)

            Dim oShape As Ppt.Shape
            Dim i As Integer
            Dim bFound As Boolean = False
            For i = 1 To oSlide.Shapes.Count
                oShape = oSlide.Shapes.Item(i)
                If InStr(oShape.Name, "Diagram") > 0 Then

                    ' 개체인지 확인
                    If (oShape.Type <> MsoShapeType.msoPlaceholder) Then
                        Exit Try
                    End If

                    Dim bb
                    bb = oShape.PlaceholderFormat
                    If (bb.type <> Ppt.PpPlaceholderType.ppPlaceholderObject) Then   '??? 확인필요 PPT1에서 잘 했는 지 확인 필요
                        Exit Try
                    End If

                    If (oShape.Fill.ForeColor.RGB <> RGB(150, 150, 150)) Then
                        Exit Try
                    End If

                    bExamResult11 = True

                End If
            Next


        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항 : "애니메이션 구성" 프로퍼티 찾아서 다시 하자. ????
        ' ---------------------------------------------------------------------
        Try
            oSlide = oPpt.Presentations(1).Slides(3)
            If oSlide.SlideShowTransition.EntryEffect <> Microsoft.Office.Interop.PowerPoint.PpEntryEffect.ppEffectSpiral Then
                'Exit Try
            End If

            Dim i As Integer
            For i = 1 To oSlide.Shapes.Count
                Dim oShape As Ppt.Shape
                oShape = oSlide.Shapes(i)


            Next

            bExamResult21 = True

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        iRealScore = 0
        If bExamResult11 Then
            iRealScore = iRealScore + iSubScore1
        Else
            addWrongComment("피라미드형 다이어그램을 추가하지 않았거나 배경색이 회색 -25%가 아닙니다")
        End If

        If bExamResult21 Then
            iRealScore = iRealScore + iSubScore2
        Else
            addWrongComment("슬라이드에 회전 애니메이션이 적용되지 않았습니다")
        End If

    End Sub
End Class
