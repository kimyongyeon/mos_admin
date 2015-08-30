Imports System.Net
Imports System.IO
Public Class Class_Server
    Public Shared r_examstart, r_examend, r_curqstseq, r_elapsedtime, r_retakecnt, r_skipno
    Public Shared Sub requestStat(ByVal bookid As String, ByVal examid As String, ByVal userid As String, _
                       ByVal examlgubun As String, ByVal examsgubun As String)
        Dim url As String = "http://www.academysoft.kr/apps/exam/exam_state_request_proc.php?" & _
            "p_bookid=" & bookid & _
            "&p_examid=" & examid & _
            "&p_userid=" & userid & _
            "&p_examlgubun=" & examlgubun & _
            "&p_examsgubun=" & examsgubun

        Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(url), HttpWebRequest)
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim reader As New StreamReader(response.GetResponseStream(), System.Text.Encoding.GetEncoding(response.CharacterSet), True)

        Dim str As String = reader.ReadToEnd()
        Dim arrValue As String() = str.Split(New Char() {","})

        r_examstart = arrValue(0)
        r_examend = arrValue(1)
        r_curqstseq = arrValue(2)
        r_elapsedtime = arrValue(3)
        r_retakecnt = arrValue(4)
        r_skipno = arrValue(5)
    End Sub

    Public Shared Sub cmdStart(ByVal bookid As String, ByVal examid As String, ByVal userid As String, _
                   ByVal examlgubun As String, ByVal examsgubun As String)
        Dim url As String = "http://www.academysoft.kr/apps/exam/exam_state_reg_proc.php?" & _
            "p_command=start" & _
            "&p_bookid=" & bookid & _
            "&p_examid=" & examid & _
            "&p_userid=" & userid & _
            "&p_examlgubun=" & examlgubun & _
            "&p_examsgubun=" & examsgubun


        Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(url), HttpWebRequest)
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim reader As New StreamReader(response.GetResponseStream(), System.Text.Encoding.GetEncoding(response.CharacterSet), True)

        Dim str As String = reader.ReadToEnd()
        Dim x
        x = str
        
    End Sub

    Public Shared Sub cmdSkip(ByVal bookid As String, ByVal examid As String, ByVal userid As String, _
                   ByVal examlgubun As String, ByVal examsgubun As String, ByVal qno As Integer, ByVal elapsedtime As String)
        Dim url As String = "http://www.academysoft.kr/apps/exam/exam_state_reg_proc.php?" & _
            "p_command=skip" & _
            "&p_bookid=" & bookid & _
            "&p_examid=" & examid & _
            "&p_userid=" & userid & _
            "&p_examlgubun=" & examlgubun & _
            "&p_examsgubun=" & examsgubun & _
            "&p_skipno=" & qno & _
            "&p_elapsedtime=" & elapsedtime

        Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(url), HttpWebRequest)
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim reader As New StreamReader(response.GetResponseStream(), System.Text.Encoding.GetEncoding(response.CharacterSet), True)

        Dim str As String = reader.ReadToEnd()
        Dim x
        x = str

    End Sub

    Public Shared Sub cmdNext(ByVal bookid As String, ByVal examid As String, ByVal userid As String, _
                   ByVal examlgubun As String, ByVal examsgubun As String, ByVal qno As Integer, _
                   ByVal wrongreason As String, ByVal score As Integer, ByVal elapsedtime As String)

        Dim url As String = "http://www.academysoft.kr/apps/exam/exam_state_reg_proc.php?" & _
            "p_command=next" & _
            "&p_bookid=" & bookid & _
            "&p_examid=" & examid & _
            "&p_userid=" & userid & _
            "&p_examlgubun=" & examlgubun & _
            "&p_examsgubun=" & examsgubun & _
            "&p_curqstseq=" & qno & _
            "&p_curjumsu=" & score & _
            "&p_elapsedtime=" & elapsedtime & _
            "&p_examwrong=" & System.Web.HttpUtility.UrlEncode(wrongreason, System.Text.Encoding.GetEncoding("EUC-KR"))

        Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(url), HttpWebRequest)
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim reader As New StreamReader(response.GetResponseStream(), System.Text.Encoding.GetEncoding(response.CharacterSet), True)

        Dim str As String = reader.ReadToEnd()
        Dim x
        x = str

    End Sub

End Class
