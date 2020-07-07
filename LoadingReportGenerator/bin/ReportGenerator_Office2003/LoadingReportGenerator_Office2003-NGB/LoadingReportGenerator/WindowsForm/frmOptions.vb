Public Class frmOptions

    Private Sub frmOptions_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim clsCommon As New common

        Try
            ' General Settings
            txtGenID.Text = My.Settings.GenID
            txtExportPath.Text = My.Settings.ExportPath
            txtLogPath.Text = My.Settings.LogPath
            txtInterval.Text = My.Settings.TimeInterval
            txtDuration.Text = My.Settings.Duration
            txtSMTP.Text = My.Settings.SMTP

            ' ODBC Settings
            txtServer.Text = My.Settings.Server
            txtLogin.Text = My.Settings.Login
            txtPassword.Text = My.Settings.Password
            txtDB.Text = My.Settings.DB
            txtTimeout.Text = My.Settings.Timeout

            ' Report Type
            chkAll.Checked = My.Settings.IsAll
            chkExcel.Checked = My.Settings.IsExcel
            chkPDF.Checked = My.Settings.IsPDF
            chkTxt.Checked = My.Settings.IsTxt
            chkZip.Checked = My.Settings.IsZip
        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnGSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGSave.Click
        Dim clsCommon As New common

        Try
            My.Settings.GenID = txtGenID.Text
            My.Settings.ExportPath = txtExportPath.Text
            My.Settings.LogPath = txtLogPath.Text
            My.Settings.TimeInterval = txtInterval.Text
            My.Settings.Duration = txtDuration.Text
            My.Settings.SMTP = txtSMTP.Text

            My.Settings.Save()

            MsgBox("General setting saved, please restart application to take effect.", MsgBoxStyle.Exclamation)

            clsCommon.SaveLog("General setting saved, please restart application to take effect.")
        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnGCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGCancel.Click
        Dim clsCommon As New common

        Try
            txtGenID.Text = My.Settings.GenID
            txtExportPath.Text = Replace(My.Settings.ExportPath & "\", "\\", "\")
            txtLogPath.Text = Replace(My.Settings.LogPath & "\", "\\", "\")
            txtInterval.Text = My.Settings.TimeInterval
            txtDuration.Text = My.Settings.Duration
            txtSMTP.Text = My.Settings.SMTP
        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnOSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOSave.Click
        Dim clsCommon As New common

        Try
            My.Settings.Server = txtServer.Text
            My.Settings.Login = txtLogin.Text
            My.Settings.Password = txtPassword.Text
            My.Settings.DB = txtDB.Text
            My.Settings.Timeout = txtTimeout.Text

            My.Settings.Save()

            MsgBox("ODBC setting saved, please restart application to take effect.", MsgBoxStyle.Exclamation)

            clsCommon.SaveLog("ODBC setting saved, please restart application to take effect.")
        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnOCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOCancel.Click
        Dim clsCommon As New common

        Try
            txtServer.Text = My.Settings.Server
            txtLogin.Text = My.Settings.Login
            txtPassword.Text = My.Settings.Password
            txtDB.Text = My.Settings.DB
            txtTimeout.Text = My.Settings.Timeout
        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub chkAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged
        If chkAll.Checked Then
            If chkExcel.Enabled Then
                chkExcel.Checked = True
            End If

            If chkPDF.Enabled Then
                chkPDF.Checked = True
            End If

            If chkTxt.Enabled Then
                chkTxt.Checked = True
            End If

            If chkZip.Enabled Then
                chkZip.Checked = True
            End If
        Else
            chkExcel.Checked = False
            chkPDF.Checked = False
            chkTxt.Checked = False
            chkZip.Checked = False
        End If
    End Sub

    Private Sub btnRptSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRptSave.Click
        Dim clsCommon As New common
        Dim RptType As String = ""

        Try
            ' All
            If chkAll.Enabled Then
                If chkAll.Checked Then
                    My.Settings.IsAll = True
                Else
                    My.Settings.IsAll = False
                End If
            Else
                My.Settings.IsAll = False
            End If

            ' Excel
            If chkExcel.Enabled Then
                If chkExcel.Checked Then
                    My.Settings.IsExcel = True
                Else
                    My.Settings.IsExcel = False
                End If
            Else
                My.Settings.IsExcel = False
            End If

            ' PDF
            If chkPDF.Enabled Then
                If chkPDF.Checked Then
                    My.Settings.IsPDF = True
                Else
                    My.Settings.IsPDF = False
                End If
            Else
                My.Settings.IsPDF = False
            End If

            ' Text
            If chkTxt.Enabled Then
                If chkTxt.Checked Then
                    My.Settings.IsTxt = True
                Else
                    My.Settings.IsTxt = False
                End If
            Else
                My.Settings.IsTxt = False
            End If

            ' Zip
            If chkZip.Enabled Then
                If chkZip.Checked Then
                    My.Settings.IsZip = True
                Else
                    My.Settings.IsZip = False
                End If
            Else
                My.Settings.IsZip = False
            End If

            My.Settings.Save()

            MsgBox("Report types saved, please restart application to take effect.", MsgBoxStyle.Exclamation)

            clsCommon.SaveLog("Report types saved, please restart application to take effect.")

            clsCommon.UpdateRptType()

        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnRptCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRptCancel.Click
        Dim clsCommon As New common

        Try
            chkAll.Checked = My.Settings.IsAll
            chkExcel.Checked = My.Settings.IsExcel
            chkPDF.Checked = My.Settings.IsPDF
            chkTxt.Checked = My.Settings.IsTxt
            chkZip.Checked = My.Settings.IsZip
        Catch ex As Exception
            clsCommon.SaveLog(ex.Message, "E")
        End Try

        clsCommon = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class