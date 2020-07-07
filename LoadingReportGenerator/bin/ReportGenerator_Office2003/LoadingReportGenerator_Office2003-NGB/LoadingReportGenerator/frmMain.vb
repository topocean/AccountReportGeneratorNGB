Public Class frmMain
    Dim count As Integer = 0
    Dim startTime As Date
    Dim UID, RptID, sPID As String
    Dim starter, stopper As New Timer
    Dim inProcess As Boolean = False

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fvi As FileVersionInfo
        Dim common As New common

        Try
            ' --------------------------------------------------------------------
            ' Auto start service within 10 seconds
            ' --------------------------------------------------------------------

            starter.Interval = 10 * 1000
            starter.Start()
            AddHandler starter.Tick, AddressOf btnStart_Click

            ' ====================================================================


            ' --------------------------------------------------------------------
            ' Close Application After 8 Hours
            ' Default Starting service at 00:00:00
            ' Default Stopping service at 08:00:00
            ' --------------------------------------------------------------------
            stopper.Interval = My.Settings.Duration * 1000
            stopper.Start()

            AddHandler stopper.Tick, AddressOf CloseMe
            ' ====================================================================

        Catch ex As Exception
            common.SaveLog(ex.Message, "E")

            ' --------------------------------------------------------------------
            ' Save Log
            ' --------------------------------------------------------------------
            lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Application Error, please revise the error log.")
            lstDisplay.SelectedIndex = lstDisplay.Items.Count - 1
        End Try

        ' ------------------------------------------------------------
        ' Destroy Variables
        ' ------------------------------------------------------------
        fvi = Nothing
        common = Nothing

        ' ------------------------------------------------------------
        ' Release Memory
        ' ------------------------------------------------------------
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        ' ------------------------------------------------------------
        ' Stop Timer
        ' ------------------------------------------------------------
        Timer1.Stop()
        starter.Stop()
        btnStart.Enabled = True

        frmOptions.ShowDialog()
        frmOptions.Focus()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()

        ' ------------------------------------------------------------
        ' Release Memory
        ' ------------------------------------------------------------
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim common As New common

        Try
            Timer1.Interval = 1000
            Timer1.Start()
            btnStart.Enabled = False

            ' Disable Auto-Start timer
            starter.Stop()

            Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Service Started...")
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1

            ' ------------------------------------------------------------
            ' Save Log
            ' ------------------------------------------------------------
            common.SaveLog("Service Started...")

        Catch ex As Exception
            ' ------------------------------------------------------------
            ' Save Log
            ' ------------------------------------------------------------
            common.SaveLog(ex.Message, "E")

            Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Fail to start service, please review the error log.")
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim sql As String = ""
        Dim RptSQL As String = ""
        Dim cn As String = ""
        Dim i As Integer
        Dim common As New common
        Dim usrDtl() As String
        Dim paraNameStr, paraValueStr As String
        Dim paraNames(), paraValues() As String

        ' ------------------------------------------------------------
        ' Connection String
        ' ------------------------------------------------------------
        cn &= "Data Source=" & My.Settings.Server & ";"
        cn &= "User ID=" & My.Settings.Login & ";"
        cn &= "Password=" & My.Settings.Password & ";"
        cn &= "Network Library=DBMSSOCN;"
        cn &= "Initial Catalog=" & My.Settings.DB & ";"
        cn &= "Connection Timeout=" & My.Settings.Timeout & ";"

        Dim sqlConn As New Data.SqlClient.SqlConnection(cn)
        Dim cmd As New Data.SqlClient.SqlCommand
        Dim sda As New Data.SqlClient.SqlDataAdapter
        Dim ds As New Data.DataSet
        count += 1

        UID = ""
        RptID = ""
        sPID = ""
        paraNameStr = ""
        paraValueStr = ""

        If count = My.Settings.TimeInterval Then
            ' ------------------------------------------------------------
            ' Start Process
            ' ------------------------------------------------------------
            inProcess = True

            ' ------------------------------------------------------------
            ' Stop Timer
            ' ------------------------------------------------------------
            Timer1.Stop()

            sqlConn.Open()
            cmd = sqlConn.CreateCommand

            Try
                ' Get the first request query from last 1 hour
                sql = "EXEC usp_GetRptQuery '" & My.Settings.GenID & "'"

                If sql <> "" Then
                    cmd.CommandTimeout = My.Settings.Timeout
                    cmd.CommandText = sql

                    sda = New Data.SqlClient.SqlDataAdapter(cmd)
                    sda.Fill(ds)

                    ' Get Request UID, Report ID and stored procedure
                    With ds.Tables(0).Rows(0)
                        UID = UCase(common.NullVal(.Item("UID").ToString, ""))
                        RptID = common.NullVal(.Item("RptID").ToString, "")
                        sPID = common.NullVal(.Item("sPID").ToString, "")
                        usrDtl = Split(common.NullVal(.Item("Usr"), ""), "|")
                    End With

                    If UID <> "" Then
                        startTime = Now

                        ' Update Query Status
                        sql = "UPDATE PdfReport SET Status = 10, LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"
                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                        ' Execute the stored procedure for requested report
                        RptSQL = "EXEC " & sPID & " @UID='" & UID & "'"

                        For i = 0 To ds.Tables(1).Rows.Count - 1
                            RptSQL &= ", " & ds.Tables(1).Rows(i).Item("ParaName") & "="

                            If ds.Tables(1).Rows(i).Item("ParaType") = 1 Then
                                RptSQL &= "'" & ds.Tables(1).Rows(i).Item("ParaVal") & "'"
                            Else
                                RptSQL &= "" & ds.Tables(1).Rows(i).Item("ParaVal") & ""
                            End If

                            ' Pad Parameters Name
                            If paraNameStr = "" Then
                                paraNameStr &= ds.Tables(1).Rows(i).Item("ParaName")
                            Else
                                paraNameStr &= "," & ds.Tables(1).Rows(i).Item("ParaName")
                            End If

                            ' Pad Parameters Value
                            If paraValueStr = "" Then
                                paraValueStr &= ds.Tables(1).Rows(i).Item("ParaVal")
                            Else
                                paraValueStr &= "," & ds.Tables(1).Rows(i).Item("ParaVal")
                            End If
                        Next
                        ds.Clear()
                        sda.Dispose()

                        ' Put parameters into array
                        paraNames = Split(paraNameStr, ",")
                        paraValues = Split(paraValueStr, ",")

                        ' Update Display Screen
                        common.SaveLog("Generating Data: " & RptID & " by " & usrDtl(0) & " (UID: " & UID & ")")
                        Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Generating Data: " & RptID & " by " & usrDtl(0))
                        Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1

                        cmd.CommandText = RptSQL
                        cmd.CommandTimeout = My.Settings.Timeout
                        sda = New Data.SqlClient.SqlDataAdapter(cmd)
                        sda.Fill(ds)

                        ExportReport(UID, RptID, ds, usrDtl, paraNames, paraValues)

                        ds.Clear()
                    End If
                End If
            Catch ex As Exception
                If UID <> "" Then
                    ' Update Query Status
                    sql = "UPDATE PdfReport SET Status = 11, Reason = '" & Replace(ex.Message, "'", "''") & "', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"
                    cmd.CommandText = sql
                    cmd.ExecuteNonQuery()
                End If

                common.SaveLog(ex.Message, "E")
                Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Exporting Report '" & RptID & "' Failure, Reason: " & ex.Message)
                Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1
            End Try

            ' ------------------------------------------------------------
            ' Close Database Connection
            ' ------------------------------------------------------------
            sqlConn.Close()

            ' ------------------------------------------------------------
            ' Destroy Variables
            ' ------------------------------------------------------------
            sql = Nothing
            RptSQL = Nothing
            i = Nothing
            count = Nothing
            cn = Nothing
            UID = Nothing
            RptID = Nothing
            sPID = Nothing
            paraNames = Nothing
            paraValues = Nothing
            paraNameStr = Nothing
            paraValueStr = Nothing

            sqlConn.Dispose()
            cmd.Dispose()
            sda.Dispose()
            ds.Dispose()

            ' ------------------------------------------------------------
            ' Release Memory
            ' ------------------------------------------------------------
            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' ------------------------------------------------------------
            ' Reset Counter
            ' ------------------------------------------------------------
            count = 0

            ' ------------------------------------------------------------
            ' Start Timer
            ' ------------------------------------------------------------
            Timer1.Interval = 1000
            Timer1.Start()

            ' ------------------------------------------------------------
            ' End of Process
            ' ------------------------------------------------------------
            inProcess = False
        End If
    End Sub

    Sub ExportReport(ByVal UID As String, ByVal RptID As String, ByVal RptDataSet As DataSet, ByVal UsrDtl As String(), ByVal inParaNames As String(), ByVal inParaValues As String())
        Dim rptFile As String = ""
        Dim rptName As String = ""
        Dim cn As String = ""
        Dim paraNames(), paraValues() As String
        Dim paraNameStr As String = ""
        Dim paraValueStr As String = ""
        Dim tmpParaVal As String = ""
        Dim tmpParaName As String = ""
        Dim i As Integer
        Dim common As New common

        ' ------------------------------------------------------------
        ' Connection String
        ' ------------------------------------------------------------
        cn &= "Data Source=" & My.Settings.Server & ";"
        cn &= "User ID=" & My.Settings.Login & ";"
        cn &= "Password=" & My.Settings.Password & ";"
        cn &= "Network Library=DBMSSOCN;"
        cn &= "Initial Catalog=" & My.Settings.DB & ";"
        cn &= "Connection Timeout=" & My.Settings.Timeout & ";"

        Dim sqlConn As New Data.SqlClient.SqlConnection(cn)
        Dim cmd As Data.SqlClient.SqlCommand
        Dim sql As String

        common.SaveLog("Generating Report: " & RptID & " UID: " & UID)
        Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Generating Report: " & RptID & " (UID: " & UID & ")")
        Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1

        Select Case RptID
            ' ----------------------------------------------------------------------------
            ' Loading Report (Yearly / Monthly)
            ' ----------------------------------------------------------------------------
            Case "RptLoading_Monthly"
                Dim rpt As New RptLoading_Monthly()
                rptFile = rpt.RptLoading_Monthly(UID, RptDataSet) & ".xls"
                rptName = "Loading Report"

                ' Convert search option parameters for displaying on emails
                For i = LBound(inParaNames) To UBound(inParaNames)
                    If Trim(inParaNames(i)) <> "" Then
                        Select Case inParaNames(i)
                            Case "@WeekNo"
                                If inParaValues(i) = 0 Then
                                    tmpParaName = ""
                                    tmpParaVal = ""
                                Else
                                    tmpParaName = "Week"
                                    tmpParaVal = inParaValues(i)
                                End If

                            Case "@Month"
                                If inParaValues(i) = 0 Then
                                    tmpParaName = ""
                                    tmpParaVal = ""
                                Else
                                    tmpParaName = "Month"
                                    tmpParaVal = common.GetMonthName(inParaValues(i))
                                End If

                            Case "@SubBrhCd"
                                tmpParaName = "Branch"
                                tmpParaVal = common.GetSubBranchName(inParaValues(i))

                            Case "@BrhCd"
                                tmpParaName = ""
                                tmpParaVal = ""

                            Case "@Location"
                                tmpParaName = ""
                                tmpParaVal = ""

                            Case "@Traffic"
                                tmpParaName = "Traffic"
                                tmpParaVal = common.GetTrafficName(inParaValues(i))

                            Case "@POType"
                                tmpParaName = "PO Type"
                                tmpParaVal = common.GetPOType(inParaValues(i))

                            Case "@BkhMode"
                                tmpParaName = "Mode"
                                If Trim(inParaValues(i)) = "" Then
                                    tmpParaVal = "ALL"
                                Else
                                    tmpParaVal = inParaValues(i)
                                End If

                            Case "@OCF"
                                tmpParaName = "OCF"
                                If Trim(inParaValues(i)) = "" Then
                                    tmpParaVal = "ALL"
                                Else
                                    tmpParaVal = inParaValues(i)
                                End If

                            Case Else
                                tmpParaName = Replace(inParaNames(i), "@", "")
                                tmpParaVal = inParaValues(i)
                        End Select

                        ' Padding Parameters Name
                        If tmpParaName <> "" Then
                            If paraNameStr = "" Then
                                paraNameStr &= tmpParaName
                            Else
                                paraNameStr &= "," & tmpParaName
                            End If
                        End If

                        ' Padding Parameters Value
                        If tmpParaVal <> "" Then
                            If paraValueStr = "" Then
                                paraValueStr &= tmpParaVal
                            Else
                                paraValueStr &= "," & tmpParaVal
                            End If
                        End If
                    End If
                Next

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                ' ----------------------------------------------------------------------------
                ' End of Loading Report (Yearly / Monthly)
                ' ----------------------------------------------------------------------------

            Case Else
                rptFile = ""
                paraNames = Split("", ",")
                paraValues = Split("", ",")

        End Select

        ' ------------------------------------------------------------
        ' Open Database Connection
        ' ------------------------------------------------------------
        sqlConn.Open()
        cmd = sqlConn.CreateCommand

        If rptFile = "" Then
            ' ------------------------------------------------------------
            ' Update Query Status
            ' ------------------------------------------------------------
            sql = "UPDATE PdfReport SET Status = 20, LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"

            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            common.SaveLog("Exporting Report " & RptID & " - No Data " & DateDiff(DateInterval.Minute, startTime, Now) & "min(s) (UID: " & UID & ")")
            Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - No Data " & DateDiff(DateInterval.Minute, startTime, Now) & "min(s)")
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1
        Else
            ' ------------------------------------------------------------
            ' Update Query Status
            ' ------------------------------------------------------------
            sql = "UPDATE PdfReport SET Status = 6, Url = '" & rptFile & "', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"

            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            common.SaveLog("Exporting Report " & RptID & " Success in " & Replace(DateDiff(DateInterval.Minute, startTime, Now), 0, 1) & " min(s) (UID: " & UID & ")")
            Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Exporting Report " & RptID & " Success in " & Replace(DateDiff(DateInterval.Minute, startTime, Now), 0, 1) & " min(s)")
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1
        End If

        ' ------------------------------------------------------------
        ' Close Database Connection
        ' ------------------------------------------------------------
        sqlConn.Close()

        ' ------------------------------------------------------------
        ' Send Email
        ' ------------------------------------------------------------
        If rptFile <> "" Then
            common.SaveLog("Prepare sending email to requested user. (UID: " & UID & ")")
            Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Prepare sending email to requested user.")
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1

            Dim clsMail As New ClsMailReport
            clsMail.MailReport(UID, UsrDtl, rptFile, rptName, paraNames, paraValues)

            If rptFile = ".xls" Then
                common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", ** No Data)")
                Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (No Data)")
            Else
                common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", File Name: " & rptFile & ")")
                Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (File Name: " & rptFile & ")")
            End If
            
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1
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

        ' ------------------------------------------------------------
        ' Release Memory
        ' ------------------------------------------------------------
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.ExitToolStripMenuItem.PerformClick()
    End Sub

    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        ' ------------------------------------------------------------
        ' Show application onto Taskbar if double the icon on System
        ' Tray
        ' ------------------------------------------------------------
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        ' ------------------------------------------------------------
        ' Hide application to System Tray if minimized
        ' ------------------------------------------------------------
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            Me.Hide()
        End If
    End Sub

    Private Sub CloseMe(ByVal sender As Object, ByVal e As System.EventArgs)
        If inProcess Then
            ' ------------------------------------------------------------
            ' Reset timer if in report generating process
            ' and re-check within 30 seconds
            ' ------------------------------------------------------------
            stopper.Interval = 30 * 1000
            stopper.Start()

        Else
            ' ------------------------------------------------------------
            ' Close Application if not in process
            ' ------------------------------------------------------------
            Me.Close()
        End If
    End Sub
End Class
