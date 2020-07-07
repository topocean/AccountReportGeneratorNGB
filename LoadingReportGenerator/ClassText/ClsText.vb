Public Class ClsText
    Public rptFile As String

    Property rptFileName()
        Get
            rptFileName = Me.rptFile
        End Get
        Set(ByVal value)
            Me.rptFile = value
        End Set
    End Property

    Sub ExportText(ByVal UID As String, ByVal RptID As String, ByVal RptDataSet As DataSet, ByVal isEmail As Integer, ByVal UsrDtl As String(), ByVal inParaNames As String(), ByVal inParaValues As String())
        Dim rptFile As String = ""
        Dim rptName As String = ""
        Dim cn As String = ""
        Dim paraNames(), paraValues() As String
        Dim paraNameStr As String = ""
        Dim paraValueStr As String = ""
        Dim tmpParaVal As String = ""
        Dim tmpParaName As String = ""
        Dim startTime As Date
        Dim i As Integer
        Dim common As New common
        Dim hasError As Boolean = False
        Dim errMsg As String = ""
        Dim errArray() As String

        ' ------------------------------------------------------------
        ' Connection String
        ' ------------------------------------------------------------
        cn &= "Data Source=" & My.Settings.Server & ";"
        cn &= "Database=" & My.Settings.DB & ";"
        cn &= "User Id=" & My.Settings.Login & ";"
        cn &= "Password=" & My.Settings.Password & ";"

        Dim sqlConn As New MySql.Data.MySqlClient.MySqlConnection(cn)
        Dim cmd As New MySql.Data.MySqlClient.MySqlCommand
        Dim sql As String

        ' ------------------------------------------------------------
        ' Open Database Connection
        ' ------------------------------------------------------------
        sqlConn.Open()
        cmd = sqlConn.CreateCommand

        sql = "CALL usp_PrintJob_Export2File('" & UID & "');"
        cmd.CommandText = sql
        cmd.CommandTimeout = My.Settings.Timeout
        cmd.ExecuteNonQuery()

        common.SaveLog("Generating Report: " & RptID & " UID: " & UID)
        frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Generating Report: " & RptID & " (UID: " & UID & ")")
        frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1

        startTime = frmMain.ProcTime
        rptFile = ""
        paraNames = Split("", ",")
        paraValues = Split("", ",")

        Select Case RptID
            Case "RptInttra"
                ' ----------------------------------------------------------------------------
                ' AMS Submission (IES)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptInttra
                rptFile = rpt.RptInttra(UID, RptDataSet)
                rptName = "Inttra e-SI"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of AMS Submission (IES)
                ' ----------------------------------------------------------------------------

            Case Else
                hasError = True
                errMsg = "Requested report not found."
                rptFile = ""
                paraNames = Split("", ",")
                paraValues = Split("", ",")

        End Select

        errArray = Split(rptFile, ",")

        If rptFile = "" Then
            ' ------------------------------------------------------------
            ' Update Query Status
            ' ------------------------------------------------------------

            If hasError Then
                sql = "CALL usp_PrintJob_Fail('" & UID & "', '" & common.setQuote(errMsg) & "');"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog("Exporting Report " & RptID & " - " & errMsg & ". (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Export Report Failed, please review the log.")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            Else
                sql = "CALL usp_PrintJob_NoData('" & UID & "');"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog("Exporting Report " & RptID & " - No Data " & DateDiff(DateInterval.Minute, startTime, Now) & "min(s) (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - No Data " & DateDiff(DateInterval.Minute, startTime, Now) & " min(s)")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1

            End If
        Else
            If LCase(rptFile(0)) = "error" Then
                ' Update Query Status
                sql = "CALL usp_PrintJob_Fail('" & UID & "', '" & common.setQuote(errArray(1)) & "');"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog(errArray(1), "E")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Exporting Report '" & RptID & "' Failure, Reason: " & errArray(1))
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            Else
                ' ------------------------------------------------------------
                ' Update Query Status
                ' ------------------------------------------------------------
                sql = "CALL usp_PrintJob_Succ('" & UID & "', '" & rptFile & "');"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog("Exporting Report " & RptID & " Success in " & Replace(DateDiff(DateInterval.Minute, startTime, Now), 0, 1) & " min(s) (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Exporting Report " & RptID & " Success in " & Replace(DateDiff(DateInterval.Minute, startTime, Now), 0, 1) & " min(s)")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            End If
        End If

        ' ------------------------------------------------------------
        ' Close Database Connection
        ' ------------------------------------------------------------
        sqlConn.Close()

        ' ------------------------------------------------------------
        ' Send Email
        ' ------------------------------------------------------------
        If rptFile <> "" And isEmail = 1 Then
            common.SaveLog("Prepare sending email to requested user. (UID: " & UID & ")")
            frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Prepare sending email to requested user.")
            frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1


            Dim clsMail As New ClsMailReport_old
            clsMail.MailReport(UID, UsrDtl, rptFile, rptName, paraNames, paraValues)
            clsMail = Nothing

            If rptFile = ".txt" Then
                common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", ** No Data)")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (No Data)")
            Else
                common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", File Name: " & rptFile & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (File Name: " & rptFile & ")")
            End If

            frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
        End If

        ' ------------------------------------------------------------
        ' Destroy Variables
        ' ------------------------------------------------------------
        sqlConn.Dispose()
        cmd.Dispose()

        cn = Nothing
        rptFile = Nothing
        startTime = Nothing
        sql = Nothing
        paraNames = Nothing
        paraValues = Nothing
        i = Nothing
        tmpParaVal = Nothing
        tmpParaName = Nothing
        hasError = Nothing
        errMsg = Nothing
        errArray = Nothing

        ' ------------------------------------------------------------
        ' Release Memory
        ' ------------------------------------------------------------
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class
