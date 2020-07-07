Imports CGZipLibrary

Public Class frmMain

    Dim count As Integer = 0
    Dim startTime As Date
    Dim UID, RptID, sPID, sUID As String
    Dim starter, stopper As New Timer
    Dim inProcess As Boolean = False
    Dim PrevsUID As String = ""
    Dim PrevRptName As String = ""
    Dim PrevUsrDtl() As String
    Dim PrevParaNames() As String
    Dim PrevParaValues() As String
    Dim AutoPrintQuery As Integer = 0

    Property ProcTime()

        Get
            ProcTime = Me.startTime
        End Get

        Set(ByVal value)
            Me.startTime = value
        End Set

    End Property

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim fvi As FileVersionInfo
        Dim common As New common

        Me.lstDisplay.HorizontalScrollbar = True

        Try
            ' --------------------------------------------------------------------
            ' Auto start service within 10 seconds
            ' --------------------------------------------------------------------
            starter.Interval = 10 * 1000
            starter.Start()
            AddHandler starter.Tick, AddressOf btnStart_Click

            ' --------------------------------------------------------------------
            ' Close Application After 8 Hours
            ' Default Starting service at 00:00:00
            ' Default Stopping service at 08:00:00
            ' --------------------------------------------------------------------
            stopper.Interval = My.Settings.Duration * 1000
            stopper.Start()

            AddHandler stopper.Tick, AddressOf CloseMe

            ' --------------------------------------------------------------------
            ' Update Report Types
            ' --------------------------------------------------------------------
            common.UpdateRptType()

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
        Dim isEmail As Integer = 0
        Dim cn As String = ""
        Dim i As Integer
        Dim common As New common
        Dim usrDtl() As String
        Dim paraNameStr, paraValueStr As String
        Dim paraNames(), paraValues() As String
        Dim exportReport As String = ""
        Dim objZip As CGZipFiles
        Dim zipFile As String
        Dim RptName As String
        Dim AutoJob As Integer
        Dim AutoEmail As String

        ' ------------------------------------------------------------
        ' Connection String
        ' ------------------------------------------------------------
        cn &= "Data Source=" & My.Settings.Server & ";"
        cn &= "Database=" & My.Settings.DB & ";"
        cn &= "User Id=" & My.Settings.Login & ";"
        cn &= "Password=" & My.Settings.Password & ";"

        Dim sqlConn As New Data.SqlClient.SqlConnection(cn)
        Dim cmd As New Data.SqlClient.SqlCommand
        Dim sda As New Data.SqlClient.SqlDataAdapter
        Dim ds As New Data.DataSet

        count += 1

        UID = ""
        RptID = ""
        sPID = ""
        sUID = ""
        paraNameStr = ""
        paraValueStr = ""

        ' Reset Messages
        If Format(Now, "HH:mm") = "00:00" Then
            Me.lstDisplay.Items.Clear()
            Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Clear Messages...")
            Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1
        End If

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
                ' Auto generate the Account Interface report
                If AutoPrintQuery = 0 Then
                    common.SaveLog("Adding Account Interface Queries")
                    lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Adding Account Interface Queries.")

                    sql = "EXEC usp_GetRptAutoQuery " & My.Settings.GenID

                    cmd.CommandTimeout = My.Settings.Timeout
                    cmd.CommandText = sql

                    cmd.ExecuteNonQuery()

                    AutoPrintQuery += 1
                End If

                ' Get the first request query from last 1 hour
                sql = "EXEC usp_GetRptQuery " & My.Settings.GenID & ""

                If sql <> "" Then
                    cmd.CommandTimeout = My.Settings.Timeout
                    cmd.CommandText = sql

                    sda = New Data.SqlClient.SqlDataAdapter(cmd)
                    sda.Fill(ds)

                    ' Get Request UID, Report ID and stored procedure
                    With ds.Tables(0).Rows(0)
                        UID = UCase(common.NullVal(.Item("UID").ToString, ""))
                        RptID = common.NullVal(.Item("RptID").ToString, "")
                        RptName = common.NullVal(.Item("RptName").ToString, "")
                        sPID = common.NullVal(.Item("sPID").ToString, "")
                        sUID = common.NullVal(.Item("sUID").ToString, "")
                        AutoJob = common.NullVal(.Item("AutoJob"), 0)
                        AutoEmail = common.NullVal(.Item("AutoEmail").ToString, "")
                        If AutoJob = 1 Then
                            usrDtl = Split("All|" & AutoEmail, "|")
                        Else
                            usrDtl = Split(common.NullVal(.Item("Usr").ToString, ""), "|")
                        End If
                        exportReport = common.NullVal(.Item("RptType").ToString, "")
                    End With

                    If UID <> "" Then
                        If PrevsUID <> sUID Then
                            If PrevsUID <> "" Then
                                'Zip all the report files into one zip file
                                objZip = New CGZipFiles

                                zipFile = My.Settings.ExportPath & PrevsUID & ".zip"

                                objZip.ZipFileName = zipFile
                                objZip.RootDirectory = My.Settings.ExportPath & PrevsUID
                                'objZip.AddFile(My.Settings.ExportPath & PrevsUID & "\*.*")
                                objZip.AddFile("*.*")

                                If objZip.MakeZipFile <> 0 Then
                                    common.SaveLog("Error on zip file for " & PrevsUID & " Reason: " & objZip.GetLastMessage)
                                    lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Error on zip the report file.")

                                    ' Release Memory
                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                    Exit Try
                                End If

                                objZip = Nothing

                                common.SaveLog("Prepare sending email to requested user. (UID: " & PrevsUID & ")")
                                lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Prepare sending email to requested user.")

                                lstDisplay.SelectedIndex = lstDisplay.Items.Count - 1

                                Dim clsMail As New ClsMailReport
                                clsMail.MailReport(PrevsUID, PrevUsrDtl, PrevsUID & ".zip", PrevRptName, PrevParaNames, PrevParaValues)
                                clsMail = Nothing

                                'If rptFile = ".xls" Then
                                'common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", ** No Data)")
                                'frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (No Data)")
                                'Else
                                'common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", File Name: " & rptFile & ")")
                                'frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (File Name: " & rptFile & ")")
                                'End If

                                'frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
                            End If

                            PrevsUID = sUID
                            PrevRptName = RptName
                            PrevUsrDtl = usrDtl
                        End If

                        startTime = Now

                        ' Update Query Status
                        'sql = "CALL usp_PrintJob_Exec('" & UID & "');"
                        sql = "UPDATE PdfReport SET Status = 2, LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"
                        cmd.CommandText = sql
                        cmd.CommandTimeout = My.Settings.Timeout
                        cmd.ExecuteNonQuery()

                        ' Execute the stored procedure for requested report
                        RptSQL = "EXEC " & sPID & "'" & UID & "'"

                        For i = 0 To ds.Tables(1).Rows.Count - 1
                            'RptSQL &= ", " & ds.Tables(1).Rows(i).Item("ParaName") & ""

                            If ds.Tables(1).Rows(i).Item("ParaType") = 1 Then
                                RptSQL &= ", " & ds.Tables(1).Rows(i).Item("ParaName").ToString & "='" & ds.Tables(1).Rows(i).Item("ParaVal").ToString & "'"
                            Else
                                RptSQL &= ", " & ds.Tables(1).Rows(i).Item("ParaName").ToString & "=" & ds.Tables(1).Rows(i).Item("ParaVal").ToString & ""
                            End If

                            If AutoJob = 1 Then
                                ' Pad Parameters Name only if ParaNickName Exist
                                If paraNameStr = "" Then
                                    If ds.Tables(1).Rows(i).Item("ParaNickName").ToString <> "" Then
                                        paraNameStr &= ds.Tables(1).Rows(i).Item("ParaNickName").ToString
                                    End If
                                Else
                                    If ds.Tables(1).Rows(i).Item("ParaNickName").ToString <> "" Then
                                        paraNameStr &= "," & ds.Tables(1).Rows(i).Item("ParaNickName").ToString
                                    End If
                                End If

                                ' Pad Parameters Value only if ParaNickName Exist
                                If paraValueStr = "" Then
                                    If ds.Tables(1).Rows(i).Item("ParaNickName").ToString <> "" Then
                                        If ds.Tables(1).Rows(i).Item("ParaValMeaning").ToString <> "" Then
                                            paraValueStr &= ds.Tables(1).Rows(i).Item("ParaValMeaning").ToString
                                        Else
                                            paraValueStr &= ds.Tables(1).Rows(i).Item("ParaVal").ToString
                                        End If
                                    End If
                                Else
                                    If ds.Tables(1).Rows(i).Item("ParaNickName").ToString <> "" Then
                                        If ds.Tables(1).Rows(i).Item("ParaValMeaning").ToString <> "" Then
                                            paraValueStr &= "," & ds.Tables(1).Rows(i).Item("ParaValMeaning").ToString
                                        Else
                                            paraValueStr &= "," & ds.Tables(1).Rows(i).Item("ParaVal").ToString
                                        End If
                                    End If
                                End If
                            Else
                                ' Pad Parameters Value
                                If paraNameStr = "" Then
                                    paraNameStr &= ds.Tables(1).Rows(i).Item("ParaName").ToString
                                Else
                                    paraNameStr &= "," & ds.Tables(1).Rows(i).Item("ParaName").ToString
                                End If

                                ' Pad Parameters Value
                                If paraValueStr = "" Then
                                    If common.NullVal(ds.Tables(1).Rows(i).Item("ParaVal").ToString, "") = "" Then
                                        If ds.Tables(1).Rows(i).Item("ParaType") = "1" Then
                                            paraValueStr &= ds.Tables(1).Rows(i).Item("ParaVal").ToString & " "
                                        Else
                                            paraValueStr &= "0"
                                        End If
                                    Else
                                        paraValueStr &= ds.Tables(1).Rows(i).Item("ParaVal").ToString
                                    End If
                                Else
                                    If common.NullVal(ds.Tables(1).Rows(i).Item("ParaVal").ToString, "") = "" Then
                                        If ds.Tables(1).Rows(i).Item("ParaType") = "1" Then
                                            paraValueStr &= "," & ds.Tables(1).Rows(i).Item("ParaVal").ToString & " "
                                        Else
                                            paraValueStr &= ",0"
                                        End If
                                    Else
                                        paraValueStr &= "," & ds.Tables(1).Rows(i).Item("ParaVal").ToString
                                    End If
                                End If
                            End If
                        Next

                        RptSQL &= ""

                        ds.Clear()
                        sda.Dispose()

                        ' Put parameters into array
                        paraNames = Split(paraNameStr, ",")
                        paraValues = Split(paraValueStr, ",")

                        PrevParaNames = paraNames
                        PrevParaValues = paraValues

                        ' Update Display Screen
                        common.SaveLog("Generating Data: " & RptID & " by " & usrDtl(0) & " (UID: " & UID & ")")
                        Me.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Generating Data: " & RptID & " by " & usrDtl(0))
                        Me.lstDisplay.SelectedIndex = Me.lstDisplay.Items.Count - 1

                        cmd.CommandText = RptSQL
                        cmd.CommandTimeout = My.Settings.Timeout
                        sda = New SqlClient.SqlDataAdapter(cmd)
                        sda.Fill(ds)

                        Select Case UCase(exportReport)
                            Case "EXCEL"
                                If My.Settings.IsExcel = True Then
                                    Dim rptObj As New ClsExcel
                                    rptObj.ExportExcel(UID, sUID, RptID, RptName, ds, isEmail, usrDtl, paraNames, paraValues)
                                    rptObj = Nothing
                                Else
                                    ' Update Query Status
                                    'sql = "CALL usp_PrintJob_Fail('" & UID & "', 'Report Type: EXCEL, does not supported');"
                                    sql = "UPDATE PdfReport SET Status = 11, Reason = 'Report Type: EXCEL, does not supported', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"
                                    cmd.CommandText = sql
                                    cmd.CommandTimeout = My.Settings.Timeout
                                    cmd.ExecuteNonQuery()
                                End If

                            Case "PDF"
                                If My.Settings.IsPDF Then
                                    Dim rptObj As New ClsZip
                                    rptObj.ExportZip(UID, RptID, ds, isEmail, usrDtl, paraNames, paraValues)
                                    rptObj = Nothing
                                Else
                                    ' Update Query Status
                                    sql = "CALL usp_PrintJob_Fail('" & UID & "', 'Report Type: PDF, does not supported');"
                                    cmd.CommandText = sql
                                    cmd.CommandTimeout = My.Settings.Timeout
                                    cmd.ExecuteNonQuery()
                                End If

                            Case "TXT"
                                If My.Settings.IsTxt Then
                                    Dim rptObj As New ClsText
                                    rptObj.ExportText(UID, RptID, ds, isEmail, usrDtl, paraNames, paraValues)
                                    rptObj = Nothing
                                Else
                                    ' Update Query Status
                                    sql = "CALL usp_PrintJob_Fail('" & UID & "', 'Report Type: TEXT, does not supported');"
                                    cmd.CommandText = sql
                                    cmd.CommandTimeout = My.Settings.Timeout
                                    cmd.ExecuteNonQuery()
                                End If

                            Case "ZIP"
                                If My.Settings.IsZip Then
                                    Dim rptObj As New ClsZip
                                    rptObj.ExportZip(UID, RptID, ds, isEmail, usrDtl, paraNames, paraValues)
                                    rptObj = Nothing
                                Else
                                    ' Update Query Status
                                    sql = "CALL usp_PrintJob_Fail('" & UID & "', 'Report Type: ZIP, does not supported');"
                                    cmd.CommandText = sql
                                    cmd.CommandTimeout = My.Settings.Timeout
                                    cmd.ExecuteNonQuery()
                                End If
                        End Select

                        ds.Clear()
                    Else
                        If PrevsUID <> sUID Then
                            If PrevsUID <> "" Then
                                'Zip all the report files into one zip file
                                objZip = New CGZipFiles

                                zipFile = My.Settings.ExportPath & PrevsUID & ".zip"

                                objZip.ZipFileName = zipFile
                                objZip.RootDirectory = My.Settings.ExportPath & PrevsUID
                                'objZip.AddFile(My.Settings.ExportPath & PrevsUID & "\*.*")
                                objZip.AddFile("*.*")

                                If objZip.MakeZipFile <> 0 Then
                                    common.SaveLog("Error on zip file for " & PrevsUID & " Reason: " & objZip.GetLastMessage)
                                    lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Error on zip the report file.")

                                    ' Release Memory
                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                    Exit Try
                                End If

                                objZip = Nothing

                                common.SaveLog("Prepare sending email to requested user. (UID: " & PrevsUID & ")")
                                lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Prepare sending email to requested user.")

                                lstDisplay.SelectedIndex = lstDisplay.Items.Count - 1

                                Dim clsMail As New ClsMailReport
                                clsMail.MailReport(PrevsUID, PrevUsrDtl, PrevsUID & ".zip", PrevRptName, PrevParaNames, PrevParaValues)
                                clsMail = Nothing

                                'If rptFile = ".xls" Then
                                'common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", ** No Data)")
                                'frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (No Data)")
                                'Else
                                'common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", File Name: " & rptFile & ")")
                                'frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (File Name: " & rptFile & ")")
                                'End If

                                'frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
                            End If

                            PrevsUID = ""
                        End If
                    End If
                End If
            Catch ex As Exception
                If UID <> "" Then
                    ' Update Query Status
                    'sql = "CALL usp_PrintJob_Fail('" & UID & "', '" & common.setQuote(ex.Message) & "');"
                    sql = "UPDATE PdfReport SET Status = 11, Reason = '" & common.setQuote(ex.Message) & "', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"
                    cmd.CommandText = sql
                    cmd.CommandTimeout = My.Settings.Timeout
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Me.ExitToolStripMenuItem.PerformClick()

    End Sub

    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick

        ' Show application onto Taskbar if double the icon on System Tray
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
