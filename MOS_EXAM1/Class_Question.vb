Public Class Class_Question
    Public iQuestionNo As Integer
    Public srcFile As String
    Public finalFile As String
    Public questFile As String
    Public Sub New()

        iQuestionNo = 7
        Dim sSubFolder As String = Format(iQuestionNo, "00")
        srcFile = "data\" & sSubFolder & "\메모.ppt"
        finalFile = "data\" & sSubFolder & "\final.xls"
        questFile = "data\\" & sSubFolder & "\question.rtf"

        'iQuestionNo = 2
        'Dim sSubFolder As String = Format(iQuestionNo, "00")
        'srcFile = "data\" & sSubFolder & "\성공전략.ppt"
        'finalFile = "data\" & sSubFolder & "\final.xls"
        'questFile = "data\\" & sSubFolder & "\question.rtf"
    End Sub

    Public Sub New(ByVal iNo As Integer)

        iQuestionNo = iNo
        Dim sSubFolder As String = Format(iQuestionNo, "00")
        srcFile = "data\" & sSubFolder & "\메모.ppt"
        finalFile = "data\" & sSubFolder & "\final.xls"
        questFile = "data\\" & sSubFolder & "\question.rtf"

        'iQuestionNo = 2
        'Dim sSubFolder As String = Format(iQuestionNo, "00")
        'srcFile = "data\" & sSubFolder & "\성공전략.ppt"
        'finalFile = "data\" & sSubFolder & "\final.xls"
        'questFile = "data\\" & sSubFolder & "\question.rtf"
    End Sub



End Class
