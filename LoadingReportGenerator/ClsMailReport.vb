Imports System.Net
Imports System.Net.Mail

Public Class ClsMailReport
    Private mailFrom As String = "reporting@topocean.com.hk"

    Sub MailReport(ByVal UID As String, ByVal UsrDtl As String(), ByVal RptFile As String, ByVal RptName As String, ByVal paraNames As String(), ByVal paraValues As String())
        Dim smtp As SmtpClient
        Dim mailBody As String = ""
        Dim mailTo As String = UsrDtl(1)
        Dim mailSubject As String = "Titan System Reporting Notice: " & RptName
        Dim mailMsg As New MailMessage
        Dim mailAttachment As Attachment
        Dim attachmentPath As String = My.Settings.ExportPath & RptFile
        Dim common As New common
        Dim i As Integer

        If common.NullVal(UsrDtl(1), "") <> "" Then
            ' Email Content
            mailBody &= "<span style=""font-family: verdana; font-size: 12px;"">" & _
                "Dear " & UsrDtl(0) & ",<br /><br />" & Chr(13) & _
                "Please be informed that your reporting request has been completed and attached to this email.<br /><br />" & Chr(13) & _
                "<table cellpadding=""2"" cellspacing=""0"" border=""1"">" & Chr(13)

            ' Report Detail
            mailBody &= "<tr bgcolor=""#FFD9B4""><td colspan=""2"" style=""font-family: verdana; font-size: 12px;""><b>Report Details</b></td></tr>" & Chr(13) & _
                "<tr><td style=""font-family: verdana; font-size: 12px;"">Report Name:</td>" & Chr(13) & _
                "<td style=""font-family: verdana; font-size: 12px;"">" & RptName & "</td></tr>" & Chr(13) & _
                "<tr><td style=""font-family: verdana; font-size: 12px;"">Request ID:</td>" & Chr(13) & _
                "<td style=""font-family: verdana; font-size: 12px;"">" & UID & "</td></tr>" & Chr(13) & _
                "<tr><td colspan=""2"" style=""font-family: verdana; font-size: 12px;"">&nbsp;</td></tr>" & Chr(13)

            ' Search Options
            mailBody &= "<tr bgcolor=""#FFD9B4""><td colspan=""2"" style=""font-family: verdana; font-size: 12px;""><b>Search Options</b></td></tr>" & Chr(13)


            For i = LBound(paraNames) To UBound(paraNames)
                mailBody &= "<tr><td style=""font-family: verdana; font-size: 12px; width: 120px;"">" & paraNames(i) & "&nbsp;&nbsp;</td>" & Chr(13) & _
                    "<td style=""font-family: verdana; font-size: 12px; width: 300px;"">" & paraValues(i) & "&nbsp;&nbsp;</td>" & Chr(13) & _
                    "</tr>" & Chr(13)
            Next

            mailBody &= "</table><br />" & Chr(13)

            ' No Data
            If Not My.Computer.FileSystem.FileExists(attachmentPath) Then
                mailBody &= "<font color=""red""><b>** NO DATA with above search options</b></font><br /><br />" & Chr(13)
            End If

            ' Mail Footer
            mailBody &= "For any inquires, please send email to <a href=""mailto:it@topocean.com.hk"">it@topocean.com.hk</a>.<br /><br />" & Chr(13) & _
                "Thanks,<br />" & Chr(13) & _
                "Titan Reporting Service.<br /><br />" & Chr(13) & _
                "<hr />" & Chr(13) & _
                "<font color=""red""><b>This is an automated system email, please DO NOT reply.</b></font>" & Chr(13) & _
                "</span>"

            mailMsg.From = New Mail.MailAddress(mailFrom, "Titan Reporting Service")
            mailMsg.To.Add(mailTo)
            mailMsg.Subject = mailSubject
            mailMsg.Body = mailBody
            mailMsg.IsBodyHtml = True

            ' Check File Existance and add file as attachment
            If My.Computer.FileSystem.FileExists(attachmentPath) Then
                mailAttachment = New Attachment(attachmentPath)
                mailMsg.Attachments.Add(mailAttachment)
            End If

            ' Send email
            smtp = New SmtpClient
            smtp.Host = My.Settings.SMTP
            smtp.Port = 25
            smtp.UseDefaultCredentials = True
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.Send(mailMsg)
        End If

        ' Destory Variables
        smtp = Nothing
        mailBody = Nothing
        mailTo = Nothing
        mailSubject = Nothing
        attachmentPath = Nothing

        If Not IsNothing(mailAttachment) Then
            mailAttachment.Dispose()
        End If

        mailMsg.Dispose()
        mailAttachment = Nothing
        mailMsg = Nothing
        i = Nothing
        common = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class
