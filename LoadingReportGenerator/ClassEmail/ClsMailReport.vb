Imports System.Net
Imports System.Net.Mail

Public Class ClsMailReport

    Private mailFrom As String = "reporting@topocean.com.hk"

    Sub MailReport(ByVal UID As String, ByVal UsrDtl As String(), ByVal RptFile As String, ByVal RptName As String, Optional ByVal paraNames As String() = Nothing, Optional ByVal paraValues As String() = Nothing, Optional ByVal RptNoData As Boolean = False, Optional ByVal isRetry As Boolean = False, Optional ByVal RptDate As String = "")

        Dim smtp As SmtpClient
        Dim mailBody As String = ""
        Dim mailTo As String = UsrDtl(1)
        Dim mailSubject As String = RptName
        Dim mailExtension As String = ""
        Dim mailMsg As New MailMessage
        Dim mailAttachment As Attachment
        Dim attachmentPath As String = My.Settings.ExportPath & RptFile
        Dim common As New common
        Dim iCount As Integer = 0
        Dim Sql As String
        Dim cn As String = ""
        Dim i As Integer = 0

        cn &= "Data Source=" & My.Settings.Server & ";"
        cn &= "Database=" & My.Settings.DB & ";"
        cn &= "User Id=" & My.Settings.Login & ";"
        cn &= "Password=" & My.Settings.Password & ";"

        Dim sqlConn As New Data.SqlClient.SqlConnection(cn)
        Dim cmd As New Data.SqlClient.SqlCommand

        sqlConn.Open()
        cmd = sqlConn.CreateCommand

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

            For iCount = LBound(paraNames) To UBound(paraNames)
                If Trim(paraNames(iCount)) <> "" And Trim(paraValues(iCount)) <> "" Then
                    mailBody &= "<tr><td style=""font-family: verdana; font-size: 12px; width: 120px;"">" & paraNames(iCount) & "&nbsp;&nbsp;</td>" & Chr(13) & _
                        "<td style=""font-family: verdana; font-size: 12px; width: 500px;"">" & paraValues(iCount) & "&nbsp;&nbsp;</td>" & Chr(13) & _
                        "</tr>" & Chr(13)
                End If
                If Trim(paraNames(iCount)) = "Branch" Or Trim(paraNames(iCount)) = "Voucher Set" Or Trim(paraNames(iCount)) = "Year" Or Trim(paraNames(iCount)) = "Week" Or Trim(paraNames(iCount)) = "Month" Then
                    mailExtension &= " " + Trim(paraNames(iCount)) + ":" + Trim(paraValues(iCount))
                End If
            Next

            mailBody &= "</table><br />" & Chr(13)

            ' Check if no data
            If isRetry Then
                If My.Computer.FileSystem.FileExists(attachmentPath) Then
                    System.Threading.Thread.CurrentThread.CurrentCulture = _
                        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

                    Dim objExcel As New Excel.Application
                    Dim objWB As Excel.Workbook = objExcel.Workbooks.Open(attachmentPath)
                    Dim objWS As Excel.Worksheet = objWB.ActiveSheet

                    RptNoData = True

                    With objWS
                        If objExcel.Application.WorksheetFunction.CountA(.Cells) <> 0 Then
                            RptNoData = False
                        Else
                            RptNoData = True
                        End If
                    End With


                    objWS.Application.Quit()
                    objWS = Nothing
                    objWB = Nothing
                    objExcel = Nothing

                    ' Wait 5 seconds
                    common.SaveLog("Waiting to close Excel. (UID: " & UID & ")")
                    frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Waiting to close Excel.")
                    frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
                    System.Threading.Thread.Sleep(5 * 1000)

                    If Not IsNothing(objExcel) Then
                        objExcel = Nothing
                    End If
                End If
            End If

            ' No Data
            'If Not My.Computer.FileSystem.FileExists(attachmentPath) Then
            If RptFile = "" Or RptNoData Then
                mailBody &= "<font color=""red""><b>** NO DATA with above search options</b></font><br /><br />" & Chr(13)
            End If
            'End If

            ' Mail Footer
            mailBody &= "For any inquires, please send email to <a href=""mailto:it@topocean.com.hk"">it@topocean.com.hk</a>.<br /><br />" & Chr(13) & _
                "Thanks,<br />" & Chr(13) & _
                "Titan Reporting Service.<br /><br />" & Chr(13) & _
                "<hr />" & Chr(13) & _
                "<font color=""red""><b>This is an automated system email, please DO NOT reply.</b></font>" & Chr(13) & _
                "</span>"

            mailMsg.From = New Mail.MailAddress(mailFrom, "Titan Reporting Service")
            'mailTo = "simon_fong@topocean.com.hk"
            mailMsg.To.Add(mailTo)
            mailSubject &= mailExtension

            If isRetry Then
                mailSubject &= " - Report Date: " & RptDate
            End If

            mailMsg.Subject = mailSubject

            mailMsg.Body = mailBody
            mailMsg.IsBodyHtml = True

            ' Check File Existance and add file as attachment
            If Not RptNoData Then
                If My.Computer.FileSystem.FileExists(attachmentPath) Then
                    mailAttachment = New Attachment(attachmentPath)
                    mailMsg.Attachments.Add(mailAttachment)
                End If
            End If

            ' Send email
            Try
                smtp = New SmtpClient
                smtp.Host = My.Settings.SMTP
                smtp.Port = 25
                smtp.UseDefaultCredentials = True
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network
                smtp.Send(mailMsg)

                Sql = "EXEC usp_PrintJob_SendMail '" & UID & "'"

                cmd.CommandText = Sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()
            Catch ex As Exception

                common.SaveLog("Sending email to requested user. (Error Message: " & ex.Message & ")")

                Sql = "EXEC usp_PrintJob_Fail '" & UID & "', N'" & common.setQuote(ex.Message) & "'"
                cmd.CommandText = Sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

            End Try
        End If

        ' Destory Variables
        smtp = Nothing
        mailBody = Nothing
        mailTo = Nothing
        mailSubject = Nothing
        attachmentPath = Nothing
        mailMsg.Dispose()
        mailAttachment = Nothing
        mailMsg = Nothing
        iCount = Nothing
        common = Nothing

        cmd.Dispose()
        sqlConn.Close()

        sqlConn = Nothing
        cmd = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

End Class
