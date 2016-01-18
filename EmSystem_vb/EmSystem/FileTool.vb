Imports System.IO
Imports System.Data.OleDb
Public Class FileTool

    Public charset As String = "UTF-8"

    ''' <summary>
    ''' 直接下載檔案(檔案為頁面控制項)
    ''' </summary>
    ''' <param name="resp">HttpResponse</param>
    ''' <param name="control">要被下載的控制項</param>
    ''' <param name="fileName">檔案名稱</param>
    ''' <remarks></remarks>
    Public Sub DownloadPageAsFile(resp As HttpResponse, control As Web.UI.WebControls.WebControl, fileName As String)
        Try
            resp.ClearContent()
            resp.ContentType = GetContentType(fileName)
            resp.Write("<meta http-equiv=Content-Type content=text/html;charset=" & charset & ">")
            resp.AddHeader("Content-Disposition", "attachment;filename=" & fileName)
            resp.Charset = charset
            Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter()
            Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)
            control.RenderControl(htmlWrite)
            resp.Write(stringWrite.ToString())
            resp.End()
        Catch ex As Exception
            Dim emsys As New StrTool()
            emsys.Debuger(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 下載檔案(stream下載)，適用於小檔案
    ''' </summary>
    ''' <param name="resp">HttpResponse</param>
    ''' <param name="filePath">檔案實體路徑</param>
    ''' <remarks></remarks>
    Public Sub DownloadFile(resp As HttpResponse, filePath As String)
        Dim fi As FileInfo = New FileInfo(filePath)
        resp.Buffer = True
        resp.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(fi.Name))
        resp.Charset = charset
        resp.ContentType = GetContentType(fi.Name)
        resp.WriteFile(fi.FullName, False)
        resp.Flush()
        resp.End()
    End Sub

    ''' <summary>
    ''' 下載檔案(stream下載)，較適用於大檔案
    ''' </summary>
    ''' <param name="resp">HttpResponse</param>
    ''' <param name="fileSize">串流大小(能給0)</param>
    ''' <param name="filePath">檔案實體路徑</param>
    ''' <remarks></remarks>
    Public Sub DownloadFile(resp As HttpResponse, fileSize As Integer, filePath As String)
        If fileSize = 0 Then
            fileSize = 10000
        End If
        Dim stream As Stream = Nothing
        Dim buffer As Byte() = New Byte(fileSize) {}
        Dim length As Integer
        Dim dataToRead As Long
        Dim fileName As String = IO.Path.GetFileName(filePath)
        Try
            stream = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
            dataToRead = stream.Length
            resp.ContentType = GetContentType(fileName)
            resp.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(fileName))
            resp.Charset = charset
            While (dataToRead > 0)
                If resp.IsClientConnected Then
                    length = stream.Read(buffer, 0, length)
                    resp.OutputStream.Write(buffer, 0, length)
                    resp.Flush()
                    buffer = New Byte(fileSize) {}
                    dataToRead = dataToRead - length
                Else
                    dataToRead = -1
                End If
            End While

        Catch ex As Exception
            Dim emsys As New StrTool
            emsys.Debuger(ex.ToString())
        Finally
            If Not stream Is Nothing Then
                stream.Close()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 取得檔案的 ContentType
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    Public Shared Function GetContentType(fileName As String) As String
        Dim contentType As String = "application/octet-stream"
        Dim ext As String = System.IO.Path.GetExtension(fileName).ToLower()
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext)
        If Not registryKey Is Nothing And Not registryKey.GetValue("Content Type") Is Nothing Then
            contentType = registryKey.GetValue("Content Type").ToString()
        End If
        Return contentType
    End Function



    ''' <summary>
    ''' Excel版本 97至2003(含2003)為Under，反之為Upper
    ''' </summary>
    Public Enum ExcelVer
        Under2003 = 8
        Upper2003 = 12
    End Enum

    ''' <summary>
    ''' 是否第一列為表頭
    ''' </summary>
    Public Enum HRD
        NO = 0
        YES = 1
    End Enum
    ''' <summary>
    ''' 讀取excel內容，回傳DataTable。(需開啟IIS32位元應用程式)
    ''' </summary>
    ''' <param name="path">檔案實體路徑</param>
    Public Function XlsToDataTable(path As String, Optional ExcelVer As ExcelVer = ExcelVer.Under2003, Optional HRD As HRD = HRD.YES) As DataTable
        '連結字串中的HDR=YES， 代表略過第一欄資料
        Dim dt As New DataTable()
        Dim strCon As String = " Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Extended Properties='Excel " & ExcelVer & ".0;HDR=" & If(HRD.YES, "YES", "NO") & "'"
        Dim oledb_con As New OleDbConnection(strCon)
        Try
            oledb_con.Open()
            Dim excelShema As DataTable = oledb_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim firstSheetName As String = excelShema.Rows(0)("TABLE_NAME").ToString()
            Dim oledb_com As New OleDbCommand(String.Format(" SELECT * FROM [{0}] ", firstSheetName), oledb_con)
            Dim oledb_dr As OleDbDataReader = oledb_com.ExecuteReader()
            dt.Load(oledb_dr)
        Catch ex As Exception
            Dim emsys As New StrTool
            emsys.Debuger(ex.ToString())
        Finally
            oledb_con.Dispose()
        End Try
        Return dt
    End Function

End Class

