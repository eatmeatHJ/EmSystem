Imports System.IO
Imports System.Reflection

Public Class StrToolForm
    Public Sub Debuger(ByVal exStr As String)
        Dim strTool As New StrTool
        strTool.isForm = True
        strTool.Debuger(exStr)
    End Sub

    Public Function getDir(endDir As String) As String
        Dim strTool As New StrTool
        strTool.isForm = True
        Return strTool.getDir(endDir)
    End Function
End Class
Public Class StrTool
    Inherits System.Web.UI.Page

    'BY Jky EatMeat V1.00
    'for webpage

    Public isForm As Boolean = False
    
    ''' <summary>
    ''' 在根目錄下的debug資料夾的紀錄檔，寫入一筆紀錄
    ''' </summary>
    ''' <param name="exStr">訊息</param>
    ''' <remarks></remarks>
    Public Sub Debuger(ByVal exStr As String)
        Dim serverPath As String = getDir("/debug")
        Dim dir As DirectoryInfo = New DirectoryInfo(serverPath)
        If Not dir.Exists Then
            dir.Create()
        End If
        Dim fileName As String = "\" & CStr(Date.Now().Date).Replace("/", "-") & ".txt"
        Try
            Dim fileStream As FileStream = New FileStream(serverPath & fileName, FileMode.OpenOrCreate)
            fileStream.Close()
            Dim sw As StreamWriter
            sw = File.AppendText(serverPath & fileName)
            If isForm Then
                sw.WriteLine("[" & Date.Now() & "] " & New StackTrace(True).GetFrame(2).GetMethod().DeclaringType.FullName)
            Else
                sw.WriteLine("[" & Date.Now() & "] " & My.Request.Url.ToString())
            End If
            sw.WriteLine("[Debug] " & exStr)
            sw.Close()
        Catch ex As Exception
            Console.Write(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 公司產品解密文字
    ''' </summary>
    ''' <param name="inputStr">欲解碼文字</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DecodeStr(ByVal inputStr As String) As String
        Dim x As Integer, x1 As Integer
        Dim returnStr As String = ""
        Try
            If String.IsNullOrEmpty(inputStr) Then
                returnStr = ""
            Else
                Dim Login_Key As Long = 999
                x1 = CInt(Val("&h" & Mid$(inputStr, Len(inputStr) - 3, 4)))
                For i As Integer = (Len(inputStr) - 4) To 1 Step -4
                    x = CInt(Val("&h" & Mid$(inputStr, i - 3, 4)))
                    x = x1 Xor x
                    returnStr = ChrW(x) & returnStr
                    x1 = CInt(Val("&h" & Mid$(inputStr, i - 3, 4)))
                Next i
                Dim Code As Integer = CInt(Val(Right$(CStr(Login_Key), 5)) Mod 32767)
                x = Code Xor x1
                returnStr = ChrW(x) & returnStr
            End If
        Catch ex As Exception

        End Try
        Return returnStr
    End Function

    ''' <summary>
    ''' 公司產品加密文字
    ''' </summary>
    ''' <param name="inputStr">欲加密文字</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EncodeStr(ByVal inputStr As String) As String
        Dim returnStr As String = ""
        Dim temp As String = ""
        Dim x As Integer, x1 As Integer
        Dim Login_Key As Long = 999
        Dim Code As Integer = CInt(Val(Right$(CStr(Login_Key), 5)) Mod 32767)
        returnStr = ""
        x = Code
        For i As Integer = 1 To Len(inputStr)
            x1 = x
            x = x1 Xor AscW(Mid$(inputStr, i, 1))
            temp = Hex(x)
            temp = Left("0000", 4 - Len(temp)) & temp
            returnStr = returnStr & temp
        Next
        Return returnStr
    End Function

    ''' <summary>
    ''' 根目錄的txt檔內的數字+1 (記數用)
    ''' </summary>
    ''' <param name="files">txt檔名</param>
    Public Sub AddTxtFileNumber(files As String)
        Try
            Dim serverPath As String = getDir("/")
            Dim fileName As String = "\" & files & ".txt"
            Dim fileStream As FileStream = New FileStream(serverPath & fileName, FileMode.OpenOrCreate)
            fileStream.Close()
            Dim sr As String
            sr = File.ReadAllText(serverPath & fileName)
            If sr = "" Then
                sr = "0"
            End If
            File.WriteAllText(serverPath & fileName, CStr(CType(sr, Long) + 1))
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' 取得根目錄的txt檔內的數字 (記數用)
    ''' </summary>
    ''' <param name="files">txt檔名</param>
    Public Function GetTxtFileNumber(files As String) As String
        Try
            Dim serverPath As String = getDir("/")
            Dim fileName As String = "\" & files & ".txt"
            Dim fileStream As FileStream = New FileStream(serverPath & fileName, FileMode.OpenOrCreate)
            fileStream.Close()
            Dim sr As String
            sr = File.ReadAllText(serverPath & fileName)
            If sr = "" Then
                sr = "0"
            End If
            Return sr.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 取得endDir目錄的完整路徑，分web和form
    ''' </summary>
    ''' <param name="endDir"></param>
    ''' <returns></returns>
    Public Function getDir(endDir As String) As String
        If isForm Then
            Dim temp_dir As New DirectoryInfo(System.Environment.CurrentDirectory)
            While temp_dir.Parent.Exists
                now_dir = New DirectoryInfo(temp_dir.Parent.FullName & endDir)
                If now_dir.Exists Then
                    Return now_dir.FullName
                    Exit While
                Else
                    temp_dir = New DirectoryInfo(temp_dir.Parent.FullName)
                End If
            End While
            Return ""
        Else
            Try
                Return Server.MapPath(endDir)
            Catch ex As Exception
                isForm = True
                Return getDir(endDir)
            End Try
        End If
    End Function

    'API
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String,
         ByVal lpReturnedString As String, ByRef nSize As Integer, ByVal lpFileName As String) As Integer

    '取得ini檔案內容
    Public Function GetIniInfo(ByVal pSection As String, ByVal pKey As String, ByVal AppPath As String)
        Return GetIniString(pSection, pKey, AppPath)
        'Return Nothing
    End Function

    Public Function GetIniInfo(ByVal pSection As String, ByVal pKey As String)
        Dim AppPath As String = ""
        If isForm Then
            'Form 會走這
            Dim dir As New DirectoryInfo(System.Environment.CurrentDirectory)
            '一直往上一層找IWKernel
            While dir.Parent.Exists
                dir = dir.Parent
                Dim fi As New FileInfo(dir.FullName & "\IWKernel.ini")
                If fi.Exists Then
                    AppPath = dir.FullName & "\IWKernel.ini"
                    Exit While
                End If
            End While
            Return GetIniString(pSection, pKey, AppPath)
        Else
            'Web 會走這
            AppPath = New DirectoryInfo(Server.MapPath("/")).Parent.FullName & "\IWKernel.ini"
            Return GetIniString(pSection, pKey, AppPath)
        End If
    End Function

    Public Function GetIniString(ByVal lkSection As String, ByVal lkKey As String, ByVal lkIniPath As String) As String
        Dim lngRtn As Long
        Dim strRtn As String

        GetIniString = ""
        strRtn = New String(Chr(20), 255)

        Dim fi As New System.IO.FileInfo(lkIniPath)
        If Not fi.Exists Then
            Throw New Exception("設定檔不存在(" & lkIniPath & ")")
        End If

        lngRtn = GetPrivateProfileString(lkSection, lkKey, vbNullString, strRtn, Len(strRtn), lkIniPath)
        If lngRtn <> 0 Then
            'GetIniString = Left$(strRtn, InStr(strRtn, Chr(0)) - 1)
            GetIniString = strRtn.Substring(0, InStr(strRtn, Chr(0)) - 1)
        End If

    End Function

    '讀取ini
    Private Sub LoadIni(ByVal currentDir As String)
        Dim iniDir As String = "ini"
        Dim myDir As DirectoryInfo = New DirectoryInfo(currentDir + iniDir)
        Dim fileType As String = ""
        Dim fieldCount As Integer = 0
        Try
            If FileIO.FileSystem.DirectoryExists(currentDir + "/" + iniDir) Then
                For Each myFile As FileInfo In myDir.GetFiles
                    If myFile.Exists Then
                        '取得ini檔的各參數值
                        '[file]裡的type
                        fileType = GetIniInfo("file", "type", myFile.FullName)
                        '[field]裡的count
                        fieldCount = CType(GetIniInfo("field", "count", myFile.FullName), Int32)
                    Else
                        MsgBox("發生錯誤: " & myFile.Name & "檔案不存在!")
                    End If
                Next
            Else
                MsgBox("發生錯誤: " & myDir.Name & "目錄不存在!")
            End If
        Catch ex As Exception
            MsgBox("LoadIni error:" + ex.Message + Chr(10) + ex.ToString)
        End Try
    End Sub
End Class

