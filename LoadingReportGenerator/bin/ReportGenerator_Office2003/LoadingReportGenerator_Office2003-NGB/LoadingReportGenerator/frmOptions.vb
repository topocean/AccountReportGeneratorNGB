Public Class frmOptions

    Private Sub frmOptions_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim clsCommon As New common

        Try
            txtGenID.Text = My.Settings.GenID
            txtExportPath.Text = My.Settings.ExportPath
            txtLogPath.Text = My.Settings.LogPath
            txtInterval.Text = My.Settings.TimeInterval
            txtDuration.Text = My.Settings.Duration
            txtSMTP.Text = My.Settings.SMTP

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
End Class