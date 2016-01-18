Imports System.IO
Imports System.Data.OleDb
Public Class ThreadFunc

    Dim thread As Threading.Thread
    Dim getUrl As String

    Public Sub ThreadHttpGet(_getUrl As String)
        getUrl = _getUrl
        thread = New System.Threading.Thread(AddressOf httpGet)  '創建執行緒實例
        thread.Start()  '開始執行緒
    End Sub

    Private Function httpGet() As Boolean
        Dim myHttpWebRequest As System.Net.HttpWebRequest
        Dim strTool As New StrTool()
        Dim success As Boolean
        Try
            Dim URL As String = getUrl
            Dim myUri As Uri = New Uri(getUrl)
            Dim myWebRequest As System.Net.WebRequest = System.Net.WebRequest.Create(URL)
            myHttpWebRequest = CType(myWebRequest, System.Net.HttpWebRequest)
            myHttpWebRequest.KeepAlive = True
            myHttpWebRequest.Timeout = 300000
            myHttpWebRequest.Method = "GET"
            myHttpWebRequest.GetResponse()
            myWebRequest.Abort()
            success = True
        Catch WebExcp As System.Net.WebException
            strTool.Debuger(Replace(WebExcp.Message.ToString(), "The remote server returned an error: (500) Internal Server Error.", "服務器出現故障無法連接"))
            httpsend_get = False
        Catch ex As Exception
            strTool.Debuger(ex.ToString)
            success = False
        Finally
            thread.Abort()
        End Try
        Return success
    End Function
End Class
