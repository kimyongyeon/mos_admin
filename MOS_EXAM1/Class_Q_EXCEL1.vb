Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

Public Class Class_Q_EXCEL1


    Public oExcel As Excel.ApplicationClass
    Public oBook As Excel.WorkbookClass
    Public oBooks As Excel.Workbooks
    Public oSheet As Excel.Worksheet

    Public Sub question_examinate15()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항 : ??? 2003에서 해봐야 함.
        ' ---------------------------------------------------------------------
        Try
            'Dim oCommandBars = oExcel.CommandBars

            'Dim returnValue As VBE
            'returnValue = oExcel.VBE

            Dim oVBC = oBook.VBProject.VBComponents


            Dim i As Integer
            For i = 1 To oVBC.Count
                Dim CountOfLines = oVBC.Item(i).CodeModule.CountOfLines
                If (CountOfLines > 0) Then
                    Dim ProcStartLine = oVBC.Item(i).CodeModule.ProcStartLine("바닥글", vbext_ProcKind.vbext_pk_Proc)
                    Dim ProcCountLines = oVBC.Item(i).CodeModule.ProcCountLines("바닥글", vbext_ProcKind.vbext_pk_Proc)
                    If (ProcCountLines > 0) Then
                        Dim ProcCode = oVBC.Item(i).CodeModule.Lines(ProcStartLine, ProcCountLines)
                        ' 2009년 실적표는 없어야하고 바닥글은 있어야 한다.
                        If (InStr(ProcCode, "2009년 실적표") <= 0 And _
                            InStr(ProcCode, "바닥글") > 0) Then
                            bExamResult11 = True
                        End If
                    End If
                End If
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항 : 2003 버전에서 해보자 ???
        ' ---------------------------------------------------------------------
        Try




        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam15-11 CustomDocumentProperties Ok")
        Else
            MsgBox("exam15-11 CustomDocumentProperties Error")
        End If

        If bExamResult21 Then
            'MsgBox("exam15-21 CustomDocumentProperties Ok")
        Else
            'MsgBox("exam15-21 CustomDocumentProperties Error")
        End If
    End Sub

    Public Sub question_examinate14()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항 : ??? 2003에서 해봐야 함.
        ' ---------------------------------------------------------------------
        Try
            Dim oCommandBars = oExcel.CommandBars

            Dim i As Integer
            For i = 1 To oCommandBars.Count
                Dim oCommandBar = oCommandBars.Item(i)
                bExamResult21 = True
            Next

        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try




        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam14-11 CustomDocumentProperties Ok")
        Else
            MsgBox("exam14-11 CustomDocumentProperties Error")
        End If

        If bExamResult21 Then
            MsgBox("exam14-21 CustomDocumentProperties Ok")
        Else
            MsgBox("exam14-21 CustomDocumentProperties Error")
        End If
    End Sub
    Public Sub question_examinate13()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            'Dim oBDPs As Microsoft.Office.Core.DocumentProperties
            'oBDPs = oExcel.Workbooks(1).BuiltinDocumentProperties
            'Dim oBDPs As Microsoft.Office.Core.DocumentProperties = oExcel.Workbooks(1).BuiltinDocumentProperties

            Dim oBDPs = oExcel.Workbooks(1).BuiltinDocumentProperties
            Dim oCDPs = oExcel.Workbooks(1).CustomDocumentProperties

            Dim i As Integer
            For i = 1 To oCDPs.Count
                Dim oDP = oCDPs.Item(i).name
                Dim ovalue = oCDPs.Item(i).value
                Dim otype = oCDPs.item(i).type

                'If (oDP.ToString() = "언어" And otype = vbBoolean) Then ' 원래 작전인데???
                If (oDP.ToString() = "언어") Then
                    If (ovalue = True) Then
                        bExamResult11 = True
                    End If
                End If

                'bExamResult11 = True
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            'Dim oBDPs As Microsoft.Office.Core.DocumentProperties
            'oBDPs = oExcel.Workbooks(1).BuiltinDocumentProperties
            'Dim oBDPs As Microsoft.Office.Core.DocumentProperties = oExcel.Workbooks(1).BuiltinDocumentProperties

            Dim oBDPs = oExcel.Workbooks(1).BuiltinDocumentProperties
            Dim oCDPs = oExcel.Workbooks(1).CustomDocumentProperties
            Dim i As Integer
            For i = 1 To oCDPs.Count
                Dim oDP = oCDPs.Item(i).name
                Dim ovalue = oCDPs.Item(i).value
                Dim otype = oCDPs.item(i).type

                'If (oDP.ToString() = "언어" And otype = vbBoolean) Then ' 원래 작전인데???
                If (oDP.ToString() = "상태") Then
                    If (ovalue = "판매불가") Then
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
        If bExamResult11 Then
            MsgBox("exam13-11 CustomDocumentProperties Ok")
        Else
            MsgBox("exam13-11 CustomDocumentProperties Error")
        End If

        If bExamResult21 Then
            MsgBox("exam13-21 CustomDocumentProperties Ok")
        Else
            MsgBox("exam13-21 CustomDocumentProperties Error")
        End If
    End Sub

    Public Sub question_examinate12()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim oQueryTables = oSheet.QueryTables

            Dim i As Integer
            For i = 1 To oQueryTables.Count
                Dim oQueryTable As Excel.QueryTable = oQueryTables.Item(i)
                If oQueryTable.TextFileParseType <> Excel.XlTextParsingType.xlDelimited Then Exit Try
                If oQueryTable.TextFileTabDelimiter <> True Then Exit Try
                If oQueryTable.Destination.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1) <> "$B$2" Then
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
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        oSheet = oExcel.Workbooks(1).Worksheets(2)

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim oQueryTables = oSheet.QueryTables

            Dim i As Integer
            For i = 1 To oQueryTables.Count
                Dim oQueryTable As Excel.QueryTable = oQueryTables.Item(i)
                If oQueryTable.TextFileParseType <> Excel.XlTextParsingType.xlDelimited Then Exit Try
                If oQueryTable.TextFileCommaDelimiter <> True Then Exit Try
                If oQueryTable.Destination.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1) <> "$A$2" Then
                    Exit Try
                End If
                bExamResult21 = True
            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam12-11 탭으로 구분하여 외부 데이터 가져오기 Ok")
        Else
            MsgBox("exam12-11 탭으로 구분하여 외부 데이터 가져오기 Error")
        End If

        If bExamResult21 Then
            MsgBox("exam12-21 쉼표로 구분하여 외부 데이터 가져오기 Ok")
        Else
            MsgBox("exam12-21 쉼표로 구분하여 외부 데이터 가져오기 Error")
        End If
    End Sub

    Public Sub question_examinate11()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(5)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            If oSheet.Range("B2").Formula = "='1호점'!B2+'2호점'!B2+'3호점'!B2+'4호점'!B2" Then
                bExamResult11 = True
            End If
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            If oSheet.Range("C2").Formula = "='1호점'!C2+'2호점'!C2+'3호점'!C2+'4호점'!C2" And _
                oSheet.Range("C11").Formula = "='1호점'!C11+'2호점'!C11+'3호점'!C11+'4호점'!C11" Then
                bExamResult21 = True
            End If
        Catch e As Exception
            MsgBox(e.Message)
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam11-11 수식 검사 Ok")
        Else
            MsgBox("exam11-11 수식 검사 Error")
        End If

        If bExamResult11 Then
            MsgBox("exam11-21 수식 검사 Ok")
        Else
            MsgBox("exam11-21 수식 검사 Error")
        End If
    End Sub

    Public Sub question_examinate10()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        oBook = oExcel.Workbooks(1)

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim oXmlMaps As Excel.XmlMaps
            Dim oXmlmap As Excel.XmlMap
            Dim i As Integer

            oXmlMaps = oBook.XmlMaps
            For i = 1 To oXmlMaps.Count
                If oXmlMaps.Item(i).Name = "dataroot_맵" Then
                    oXmlmap = oXmlMaps.Item(i)
                    Dim oConnecton = oXmlmap.WorkbookConnection
                    Dim j As Integer
                    For j = 1 To oConnecton.Ranges.Count
                        Dim r = oConnecton.Ranges.Item(j)
                        'MsgBox(r)
                        If (r.Formula(1, 1) <> "사번") Then Exit For
                        If (r.Formula(1, 2) <> "이름") Then Exit For
                        If (r.Formula(1, 3) <> "전화번호") Then Exit For

                        If (r.Formula(2, 1) <> "1") Then Exit For
                        If (r.Formula(2, 2) <> "백동혁") Then Exit For
                        If (r.Formula(2, 3) <> "(02)912-2139") Then Exit For
                        bExamResult11 = True
                    Next
                    Exit For
                End If
            Next


        Catch e As Exception
            MsgBox(e.Message)
            'Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            ' 문제 시작 전에 파일을 먼지 지워야 함. ???
            If FileLen(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & _
                    "\사원연락처.xml") > 0 Then
                bExamResult21 = True
            End If


        Catch e As Exception
            'MsgBox(e.Message)
            'Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam10-11 XML 맵핑 및 데이터 가져오기 Ok")
        Else
            MsgBox("exam10-11 XML 맵핑 및 데이터 가져오기 Error")
        End If

        If bExamResult11 Then
            MsgBox("exam10-21 XML 데이터 내보내기 Ok")
        Else
            MsgBox("exam10-21 XML 데이터 내보내기 Error")
        End If

    End Sub
    Public Sub question_examinate09()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult12 As Boolean = False     ' 1번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult13 As Boolean = False     ' 1번 문제 3번 지시 사항 채점 결과

        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult22 As Boolean = False     ' 2번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult23 As Boolean = False     ' 2번 문제 3번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim oRange As Excel.Range
            oRange = oSheet.Range("G4").Precedents
            'MsgBox(oRange)
            '포기


        Catch e As Exception
            'MsgBox(e.Message)
            'Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항 : 문제를 ???
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam10-11 참조되는 셀 추적 Ok")
        Else
            MsgBox("exam10-11 참조되는 셀 추적 Error")
        End If

    End Sub

    Public Sub question_examinate08()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult12 As Boolean = False     ' 1번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult13 As Boolean = False     ' 1번 문제 3번 지시 사항 채점 결과

        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult22 As Boolean = False     ' 2번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult23 As Boolean = False     ' 2번 문제 3번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim oListObjects As Excel.ListObjects
            oListObjects = oSheet.ListObjects

            Dim i As Integer
            For i = 1 To oListObjects.Count
                Dim oListObject As Excel.ListObject

                oListObject = oListObjects.Item(i)
                If oListObject.SourceType <> Excel.XlListObjectSourceType.xlSrcRange Then Exit Try

                Dim sRange = oListObject.Range.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1)
                If sRange <> "$A$3:$I$11" Then Exit Try
            Next
            bExamResult11 = True

        Catch e As Exception
            'MsgBox(e.Message)
            'Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항 : 문제를 ???
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam7-11 목록 만들기 Ok")
        Else
            MsgBox("exam7-11 목록 만들기 Error")
        End If

    End Sub
    Public Sub question_examinate07()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult12 As Boolean = False     ' 1번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult13 As Boolean = False     ' 1번 문제 3번 지시 사항 채점 결과

        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult22 As Boolean = False     ' 2번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult23 As Boolean = False     ' 2번 문제 3번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim oListObjects As Excel.ListObjects
            oListObjects = oSheet.ListObjects

            Dim i As Integer
            For i = 1 To oListObjects.Count
                Dim oListObject As Excel.ListObject

                oListObject = oListObjects.Item(i)
                If oListObject.SourceType <> Excel.XlListObjectSourceType.xlSrcRange Then Exit Try

                Dim sRange = oListObject.Range.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1)
                If sRange <> "$A$3:$I$11" Then Exit Try

                bExamResult11 = True
            Next


        Catch e As Exception
            'MsgBox(e.Message)
            'Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항 : 문제를 ???
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam7-11 목록 만들기 Ok")
        Else
            MsgBox("exam7-11 목록 만들기 Error")
        End If


    End Sub
    Public Sub question_examinate06()
        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult12 As Boolean = False     ' 1번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult13 As Boolean = False     ' 1번 문제 3번 지시 사항 채점 결과

        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult22 As Boolean = False     ' 2번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult23 As Boolean = False     ' 2번 문제 3번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)  ' 주의
        End If


        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim v As Excel.Validation
            v = oSheet.Range("C5:C12").Validation

            If (v.Type <> Excel.XlDVType.xlValidateList) Then Exit Try
            If (v.Formula1.Replace(" ", "") <> "서울대,연세대,부산대,대구대,광주대,대전대") Then Exit Try
            If (v.InputMessage.Replace(" ", "") <> "해당하는 학교만 입력") Then Exit Try
            bExamResult11 = True

        Catch e As Exception
            'MsgBox(e.Message)
            'Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 2번 지시 사항 : 문제를 ???
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam6-11 데이터 유효성 검사 Ok")
        Else
            MsgBox("exam6-11 데이터 유효성 검사 Error")
        End If


    End Sub
    Public Sub question_examinate05()
        ' 나중에

        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult12 As Boolean = False     ' 1번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult13 As Boolean = False     ' 1번 문제 3번 지시 사항 채점 결과

        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과
        'Dim bExamResult22 As Boolean = False     ' 2번 문제 2번 지시 사항 채점 결과
        'Dim bExamResult23 As Boolean = False     ' 2번 문제 3번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(2)  ' 주의
        End If


        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            Dim o As Excel.Outline
            o = oSheet.Outline

            oSheet.Range("E:F").Group()

            oSheet.Columns.ClearOutline()

            bExamResult11 = True
        Catch e As Exception
            MsgBox(e.Message)
            Exit Sub
        End Try



        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam6-11 Y축(값)의 숫자 출력 형 Ok")
        Else
            MsgBox("exam6-11 Y축(값)의 숫자 출력 형 Error")
        End If


    End Sub

    Public Sub question_examinate04()

        ' ---------------------------------------------------------------------
        ' 채점 결과 초기화
        ' ---------------------------------------------------------------------
        Dim bExamResult11 As Boolean = False     ' 1번 문제 1번 지시 사항 채점 결과
        Dim bExamResult12 As Boolean = False     ' 1번 문제 2번 지시 사항 채점 결과
        Dim bExamResult13 As Boolean = False     ' 1번 문제 3번 지시 사항 채점 결과

        Dim bExamResult21 As Boolean = False     ' 2번 문제 1번 지시 사항 채점 결과
        Dim bExamResult22 As Boolean = False     ' 2번 문제 2번 지시 사항 채점 결과
        Dim bExamResult23 As Boolean = False     ' 2번 문제 3번 지시 사항 채점 결과

        ' ---------------------------------------------------------------------
        ' 1번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' ---------------------------------------------------------------------
        ' 차트 찾기 : 차트 이름은 엑셀 메뉴의 [보기]->[차트 창]에서 확인 가능
        ' ---------------------------------------------------------------------
        Dim oShape As Excel.Shape
        Dim i As Integer
        Dim bFound As Boolean = False
        For i = 1 To oSheet.Shapes.Count
            oShape = oSheet.Shapes.Item(i)
            If oShape.Type = MsoShapeType.msoChart And _
                oShape.Name = "Chart 3" Then
                bFound = True
                Exit For
            End If
        Next
        If bFound = False Then
            MsgBox("Error Chart 3 Not Found")
            Exit Sub
        End If

        ' ---------------------------------------------------------------------
        ' 차트 객체 초기화
        ' ---------------------------------------------------------------------
        Dim oChart As Excel.Chart = oShape.Chart  ' Warnning 없애는 방법은 ???

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            ' 차트의 Y축(값)의 숫자 출력 형식 구하기
            Dim sNumberFormat As String
            sNumberFormat = oChart.Axes(Excel.XlAxisType.xlValue).TickLabels.NumberFormatLocal
            sNumberFormat = sNumberFormat.Substring(1, sNumberFormat.IndexOf(";"))

            ' 소수점 이하 0이 1개 인지 검사
            If (sNumberFormat.IndexOf(".0") < 0) Then Exit Try ' "소수점 아래가 최소한 1개 0"
            If (sNumberFormat.IndexOf(".00") > 0) Then Exit Try '"소수점 아래가 최소한 2개 0"
            bExamResult11 = True
        Catch e As Exception
            MsgBox(e.Message)
            Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 2번 지시 사항
        ' ---------------------------------------------------------------------
        Dim oRangeUnion As Excel.Range      '차트 데이터 영역들을 합집합으로
        oRangeUnion = oSheet.Range("$L$1")  '임시, 꽁수, 정확하게 하려면???
        oRangeUnion.Clear()

        Dim oRange As Excel.Range

        For n = 1 To oChart.SeriesCollection.Count
            Dim oSeriesFormula = oChart.SeriesCollection(n).Formula   ' "=SERIES('11월'!$D$4,'11월'!$C$5:$C$10,'11월'!$D$5:$D$10,1)"

            Dim sRange As String
            sRange = oSeriesFormula.ToString
            sRange = sRange.Substring(sRange.IndexOf("("))
            sRange = sRange.Substring(1, sRange.LastIndexOf(",") - 1)

            oRange = oExcel.Range(sRange)                       ' "'11월'!$D$4,'11월'!$C$5:$C$10,'11월'!$D$5:$D$10"
            oRangeUnion = oExcel.Union(oRangeUnion, oRange)
        Next

        Dim oRangeIntersect As Excel.Range  ' 차트 데이터 영역과 추가했어야할 영역의 교집합
        oRangeIntersect = oExcel.Intersect(oRangeUnion, oExcel.Range("'11월'!$C$10:$E$10"))

        Dim sFormulaInterect As String      ' 교집합을 A1 같은 참조형식으로
        sFormulaInterect = oRangeIntersect.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1)

        If (sFormulaInterect = "$C$10:$E$10") Then
            bExamResult12 = True
        End If


        ' ---------------------------------------------------------------------
        ' 채점 : 1번 문제 3번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            ' 차트의 X축(항목)의 순서가 거꾸로 인지
            Dim bReverse As String
            bReverse = oChart.Axes(Excel.XlAxisType.xlCategory).ReversePlotOrder

            ' 소수점 이하 0이 1개 인지 검사
            If (bReverse = True) Then bExamResult13 = True
        Catch e As Exception
            MsgBox(e.Message)
            Exit Sub
        End Try

        ' ---------------------------------------------------------------------
        ' 2번 문제
        ' ---------------------------------------------------------------------
        ' 작업용 Worksheet 초기화
        ' ---------------------------------------------------------------------
        'If oSheet Is Nothing Then
        oSheet = oExcel.Workbooks(1).Worksheets(2)
        'End If

        ' ---------------------------------------------------------------------
        ' 차트 찾기
        ' ---------------------------------------------------------------------
        'Dim oShape As Excel.Shape
        'Dim i As Integer
        'Dim bFound As Boolean = False
        bFound = False
        For i = 1 To oSheet.Shapes.Count
            oShape = oSheet.Shapes.Item(i)
            If oShape.Type = MsoShapeType.msoChart And _
                oShape.Name = "Chart 1" Then
                bFound = True
                Exit For
            End If
        Next
        If bFound = False Then
            MsgBox("Error Chart 1 Not Found")
            Exit Sub
        End If

        ' ---------------------------------------------------------------------
        ' 차트 객체 초기화
        ' ---------------------------------------------------------------------
        'Dim oChart As Excel.Chart = oShape.Chart  ' Warnning 없애는 방법은 ???
        oChart = oShape.Chart  ' Warnning 없애는 방법은 ???

        ' ---------------------------------------------------------------------
        ' 채점 : 2번 문제 1번 지시 사항
        ' ---------------------------------------------------------------------
        Try
            For n = 1 To oChart.SeriesCollection.Count
                Dim oTrendLines As Excel.Trendlines
                oTrendLines = oChart.SeriesCollection(n).Trendlines

                For m = 1 To oTrendLines.Count

                    Dim oTrendLine = oTrendLines.Item(m)
                    If (oTrendLine.Type = Excel.XlTrendlineType.xlLinear) Then
                        bExamResult21 = True

                        If (oTrendLine.DisplayEquation = True) Then
                            bExamResult22 = True
                        End If

                        If (oTrendLine.DisplayRSquared = True) Then
                            bExamResult23 = True
                        End If
                    End If
                Next

            Next
        Catch e As Exception
            MsgBox(e.Message)
        End Try


        ' ---------------------------------------------------------------------
        ' 채점 종료
        ' ---------------------------------------------------------------------
        If bExamResult11 Then
            MsgBox("exam4-11 Y축(값)의 숫자 출력 형 Ok")
        Else
            MsgBox("exam4-11 Y축(값)의 숫자 출력 형 Error")
        End If

        If bExamResult12 Then
            MsgBox("exam4-12 차트에 데이터 추가 Ok")
        Else
            MsgBox("exam4-12 차트에 데이터 추가 Error")
        End If

        If bExamResult13 Then
            MsgBox("exam4-13 항목의 순서가 거꾸로 Ok")
        Else
            MsgBox("exam4-13 항목의 순서가 거꾸로 Error")
        End If

        If bExamResult21 Then
            MsgBox("exam4-21 추세선 추가 Ok")
        Else
            MsgBox("exam4-21 추세선 추가 Error")
        End If

        If bExamResult22 Then
            MsgBox("exam4-22 수식 표시 Ok")
        Else
            MsgBox("exam4-22 수식 표시 Error")
        End If

        If bExamResult23 Then
            MsgBox("exam4-23 R-제곱승 표시 Ok")
        Else
            MsgBox("exam4-23 R-제곱승 표시 Error")
        End If
    End Sub
    Public Sub question_examinate03()
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        Dim i As Integer
        Dim a As Excel.Shape
        Dim iMaxZorderPosition As Integer = -1
        Dim iGroupZorderPosition As Integer = 0
        Dim bFound As Boolean = False
        For i = 1 To oSheet.Shapes.Count
            a = oSheet.Shapes.Item(i)
            'If InStr(a.Name, "Group") = 1 Then
            '    MsgBox(a.msg)
            'End If
            If (iMaxZorderPosition < a.ZOrderPosition) Then
                iMaxZorderPosition = a.ZOrderPosition
            End If
            If a.Type = MsoShapeType.msoGroup Then
                'MsgBox(a.Name & ";" & a.Type)
                MsgBox(a.Name & ";" & a.ZOrderPosition)

                bFound = True
                iGroupZorderPosition = a.ZOrderPosition
                Dim j As Integer
                For j = 1 To a.GroupItems.Count
                    'MsgBox(a.GroupItems.Item(j).Name)
                    If (a.GroupItems.Item(j).Name <> "Picture 11" And _
                        a.GroupItems.Item(j).Name <> "Text Box 4") Then
                        bFound = False
                        iGroupZorderPosition = 0   ' 2개 문제 한꺼번에 나올 땐 불필요할 수 도 있음
                    End If
                Next

            End If
        Next

        If bFound Then
            MsgBox("exam3-1 Group Ok")
        Else
            MsgBox("exam3-1 Group Error")
        End If

        If iMaxZorderPosition = iGroupZorderPosition Then
            MsgBox("exam3-2 ZOrderPosition Ok")
        Else
            MsgBox("exam3-2 ZOrderPosition Error")
        End If


    End Sub


    Public Sub question_examinate02()
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        ' crop도해야겠지요?
        If oSheet.Shapes.Item(1).PictureFormat.Brightness = 0.5 And _
           oSheet.Shapes.Item(1).PictureFormat.Contrast = 0.5 And _
           oSheet.Shapes.Item(1).PictureFormat.ColorType = MsoPictureColorType.msoPictureAutomatic Then
            MsgBox("exam2-1 shape ok")
        Else
            MsgBox("exam2-1 shape error")
        End If
        '압축은 ???

        '크기 조정은???

        Dim a As Excel.Shape
        a = oSheet.Shapes.Item(1)

        Dim b As Excel.Shape
        b = oSheet.Shapes.Item(2)

    End Sub
    Public Sub question_examinate01()
        If oSheet Is Nothing Then
            oSheet = oExcel.Workbooks(1).Worksheets(1)
        End If

        If oSheet.Range("B3").NumberFormat = "@""광역시""" And _
           oSheet.Range("B7").NumberFormat = "@""광역시""" Then
            MsgBox("exam1-1 numberformat ok")
        Else
            MsgBox("exam1-1 numberformat error")
        End If

        If oSheet.Range("A3").NumberFormat = """CE-""#" And _
            oSheet.Range("A7").NumberFormat = """CE-""#" Then
            MsgBox("exam1-2 numberformat ok")
        Else
            MsgBox("exam1-2 numberformat error")
        End If
    End Sub
End Class
