Imports System.Net.Mail
Imports System.Net

Public Class EmMail

    ''' <summary>
    ''' 發信者Email
    ''' </summary>
    Public fromAddr As String '系統Email 
    ''' <summary>
    ''' 發信者名稱
    ''' </summary>
    Public fromName As String '系統名
    ''' <summary>
    ''' 主旨
    ''' </summary>
    Public subject As String '主旨
    ''' <summary>
    ''' 信件內文
    ''' </summary>
    Public body As String '內文
    ''' <summary>
    ''' 收件者(多個收信者可以逗點隔開)
    ''' </summary>
    Public toMail As String '收件者
    ''' <summary>
    ''' 副本收信者(多個收信者可以逗點隔開)
    ''' </summary>
    Public ccMail As String = Nothing '副本
    ''' <summary>
    ''' 密件收信者(多個收信者可以逗點隔開)
    ''' </summary>
    Public bccMail As String '密件
    ''' <summary>
    ''' 重要性(優先權)
    ''' </summary>
    Public priority As MailPriority = MailPriority.Normal
    ''' <summary>
    ''' SMTP server 位置
    ''' </summary>
    Public smtpAddr As String
    ''' <summary>
    ''' SMTP 帳號
    ''' </summary>
    Public smtpAccount As String
    ''' <summary>
    ''' SMTP 密碼
    ''' </summary>
    Public smtpPwd As String

    Public Function send() As String
        Try
            Dim msg As MailMessage = New MailMessage()
            msg.From = New MailAddress(fromAddr, fromName)
            msg.To.Add(toMail)
            If Not String.IsNullOrEmpty(ccMail) Then
                msg.CC.Add(ccMail)
            End If
            If Not String.IsNullOrEmpty(bccMail) Then
                msg.Bcc.Add(bccMail)
            End If
            msg.Priority = priority
            msg.Subject = subject
            msg.Body = body
            msg.IsBodyHtml = True

            Dim smtp As SmtpClient = New SmtpClient(smtpAddr)
            smtp.Credentials = New NetworkCredential(smtpAccount, smtpPwd)
            smtp.Send(msg)
            Return True
        Catch ex As Exception
            Dim str As New StrTool
            str.Debuger(ex.ToString())
            Return False
        End Try
    End Function

End Class


