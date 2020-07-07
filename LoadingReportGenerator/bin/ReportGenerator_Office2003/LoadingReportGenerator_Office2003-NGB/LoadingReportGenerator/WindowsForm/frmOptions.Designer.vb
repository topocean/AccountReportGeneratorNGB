<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptions
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tabOptions = New System.Windows.Forms.TabControl
        Me.tabGnlSetting = New System.Windows.Forms.TabPage
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtDuration = New System.Windows.Forms.TextBox
        Me.btnGCancel = New System.Windows.Forms.Button
        Me.btnGSave = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtSMTP = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtInterval = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtLogPath = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtExportPath = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtGenID = New System.Windows.Forms.TextBox
        Me.tabODBC = New System.Windows.Forms.TabPage
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtDB = New System.Windows.Forms.TextBox
        Me.btnOCancel = New System.Windows.Forms.Button
        Me.btnOSave = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtTimeout = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtLogin = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtServer = New System.Windows.Forms.TextBox
        Me.tabRptType = New System.Windows.Forms.TabPage
        Me.btnRptCancel = New System.Windows.Forms.Button
        Me.btnRptSave = New System.Windows.Forms.Button
        Me.chkZip = New System.Windows.Forms.CheckBox
        Me.chkTxt = New System.Windows.Forms.CheckBox
        Me.chkPDF = New System.Windows.Forms.CheckBox
        Me.chkExcel = New System.Windows.Forms.CheckBox
        Me.chkAll = New System.Windows.Forms.CheckBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.tabOptions.SuspendLayout()
        Me.tabGnlSetting.SuspendLayout()
        Me.tabODBC.SuspendLayout()
        Me.tabRptType.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabOptions
        '
        Me.tabOptions.Controls.Add(Me.tabGnlSetting)
        Me.tabOptions.Controls.Add(Me.tabODBC)
        Me.tabOptions.Controls.Add(Me.tabRptType)
        Me.tabOptions.Location = New System.Drawing.Point(12, 12)
        Me.tabOptions.Name = "tabOptions"
        Me.tabOptions.SelectedIndex = 0
        Me.tabOptions.Size = New System.Drawing.Size(533, 249)
        Me.tabOptions.TabIndex = 0
        '
        'tabGnlSetting
        '
        Me.tabGnlSetting.Controls.Add(Me.Label13)
        Me.tabGnlSetting.Controls.Add(Me.Label14)
        Me.tabGnlSetting.Controls.Add(Me.txtDuration)
        Me.tabGnlSetting.Controls.Add(Me.btnGCancel)
        Me.tabGnlSetting.Controls.Add(Me.btnGSave)
        Me.tabGnlSetting.Controls.Add(Me.Label6)
        Me.tabGnlSetting.Controls.Add(Me.txtSMTP)
        Me.tabGnlSetting.Controls.Add(Me.Label5)
        Me.tabGnlSetting.Controls.Add(Me.Label4)
        Me.tabGnlSetting.Controls.Add(Me.txtInterval)
        Me.tabGnlSetting.Controls.Add(Me.Label3)
        Me.tabGnlSetting.Controls.Add(Me.txtLogPath)
        Me.tabGnlSetting.Controls.Add(Me.Label2)
        Me.tabGnlSetting.Controls.Add(Me.txtExportPath)
        Me.tabGnlSetting.Controls.Add(Me.Label1)
        Me.tabGnlSetting.Controls.Add(Me.txtGenID)
        Me.tabGnlSetting.Location = New System.Drawing.Point(4, 22)
        Me.tabGnlSetting.Name = "tabGnlSetting"
        Me.tabGnlSetting.Padding = New System.Windows.Forms.Padding(3)
        Me.tabGnlSetting.Size = New System.Drawing.Size(525, 223)
        Me.tabGnlSetting.TabIndex = 0
        Me.tabGnlSetting.Text = "General"
        Me.tabGnlSetting.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(251, 126)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(65, 13)
        Me.Label13.TabIndex = 15
        Me.Label13.Text = "(Seconds)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(18, 126)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(122, 13)
        Me.Label14.TabIndex = 14
        Me.Label14.Text = "Application Duration"
        '
        'txtDuration
        '
        Me.txtDuration.Location = New System.Drawing.Point(146, 123)
        Me.txtDuration.Name = "txtDuration"
        Me.txtDuration.Size = New System.Drawing.Size(99, 21)
        Me.txtDuration.TabIndex = 5
        '
        'btnGCancel
        '
        Me.btnGCancel.Location = New System.Drawing.Point(434, 185)
        Me.btnGCancel.Name = "btnGCancel"
        Me.btnGCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnGCancel.TabIndex = 8
        Me.btnGCancel.Text = "&Cancel"
        Me.btnGCancel.UseVisualStyleBackColor = True
        '
        'btnGSave
        '
        Me.btnGSave.Location = New System.Drawing.Point(353, 185)
        Me.btnGSave.Name = "btnGSave"
        Me.btnGSave.Size = New System.Drawing.Size(75, 23)
        Me.btnGSave.TabIndex = 7
        Me.btnGSave.Text = "&Save"
        Me.btnGSave.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(18, 153)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "SMTP Host"
        '
        'txtSMTP
        '
        Me.txtSMTP.Location = New System.Drawing.Point(146, 150)
        Me.txtSMTP.Name = "txtSMTP"
        Me.txtSMTP.Size = New System.Drawing.Size(347, 21)
        Me.txtSMTP.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(251, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "(Seconds)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 99)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(89, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Timer Interval"
        '
        'txtInterval
        '
        Me.txtInterval.Location = New System.Drawing.Point(146, 96)
        Me.txtInterval.Name = "txtInterval"
        Me.txtInterval.Size = New System.Drawing.Size(99, 21)
        Me.txtInterval.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Log Path"
        '
        'txtLogPath
        '
        Me.txtLogPath.Location = New System.Drawing.Point(146, 69)
        Me.txtLogPath.Name = "txtLogPath"
        Me.txtLogPath.Size = New System.Drawing.Size(347, 21)
        Me.txtLogPath.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Export Path"
        '
        'txtExportPath
        '
        Me.txtExportPath.Location = New System.Drawing.Point(146, 42)
        Me.txtExportPath.Name = "txtExportPath"
        Me.txtExportPath.Size = New System.Drawing.Size(347, 21)
        Me.txtExportPath.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Gen. ID"
        '
        'txtGenID
        '
        Me.txtGenID.Location = New System.Drawing.Point(146, 15)
        Me.txtGenID.Name = "txtGenID"
        Me.txtGenID.Size = New System.Drawing.Size(99, 21)
        Me.txtGenID.TabIndex = 1
        '
        'tabODBC
        '
        Me.tabODBC.Controls.Add(Me.Label12)
        Me.tabODBC.Controls.Add(Me.txtDB)
        Me.tabODBC.Controls.Add(Me.btnOCancel)
        Me.tabODBC.Controls.Add(Me.btnOSave)
        Me.tabODBC.Controls.Add(Me.Label11)
        Me.tabODBC.Controls.Add(Me.Label10)
        Me.tabODBC.Controls.Add(Me.txtTimeout)
        Me.tabODBC.Controls.Add(Me.Label9)
        Me.tabODBC.Controls.Add(Me.txtPassword)
        Me.tabODBC.Controls.Add(Me.Label8)
        Me.tabODBC.Controls.Add(Me.txtLogin)
        Me.tabODBC.Controls.Add(Me.Label7)
        Me.tabODBC.Controls.Add(Me.txtServer)
        Me.tabODBC.Location = New System.Drawing.Point(4, 22)
        Me.tabODBC.Name = "tabODBC"
        Me.tabODBC.Padding = New System.Windows.Forms.Padding(3)
        Me.tabODBC.Size = New System.Drawing.Size(525, 223)
        Me.tabODBC.TabIndex = 1
        Me.tabODBC.Text = "ODBC"
        Me.tabODBC.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(18, 99)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 16
        Me.Label12.Text = "Database"
        '
        'txtDB
        '
        Me.txtDB.Location = New System.Drawing.Point(162, 96)
        Me.txtDB.Name = "txtDB"
        Me.txtDB.Size = New System.Drawing.Size(347, 21)
        Me.txtDB.TabIndex = 4
        '
        'btnOCancel
        '
        Me.btnOCancel.Location = New System.Drawing.Point(434, 185)
        Me.btnOCancel.Name = "btnOCancel"
        Me.btnOCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnOCancel.TabIndex = 7
        Me.btnOCancel.Text = "&Cancel"
        Me.btnOCancel.UseVisualStyleBackColor = True
        '
        'btnOSave
        '
        Me.btnOSave.Location = New System.Drawing.Point(353, 185)
        Me.btnOSave.Name = "btnOSave"
        Me.btnOSave.Size = New System.Drawing.Size(75, 23)
        Me.btnOSave.TabIndex = 6
        Me.btnOSave.Text = "&Save"
        Me.btnOSave.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(281, 126)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(65, 13)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "(Seconds)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(18, 126)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(121, 13)
        Me.Label10.TabIndex = 11
        Me.Label10.Text = "Connection Timeout"
        '
        'txtTimeout
        '
        Me.txtTimeout.Location = New System.Drawing.Point(162, 123)
        Me.txtTimeout.Name = "txtTimeout"
        Me.txtTimeout.Size = New System.Drawing.Size(113, 21)
        Me.txtTimeout.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(18, 72)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(61, 13)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "Password"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(162, 69)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(347, 21)
        Me.txtPassword.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(18, 45)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(37, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Login"
        '
        'txtLogin
        '
        Me.txtLogin.Location = New System.Drawing.Point(162, 42)
        Me.txtLogin.Name = "txtLogin"
        Me.txtLogin.Size = New System.Drawing.Size(347, 21)
        Me.txtLogin.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(18, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 13)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Database Server"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(162, 15)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(347, 21)
        Me.txtServer.TabIndex = 1
        '
        'tabRptType
        '
        Me.tabRptType.Controls.Add(Me.btnRptCancel)
        Me.tabRptType.Controls.Add(Me.btnRptSave)
        Me.tabRptType.Controls.Add(Me.chkZip)
        Me.tabRptType.Controls.Add(Me.chkTxt)
        Me.tabRptType.Controls.Add(Me.chkPDF)
        Me.tabRptType.Controls.Add(Me.chkExcel)
        Me.tabRptType.Controls.Add(Me.chkAll)
        Me.tabRptType.Location = New System.Drawing.Point(4, 22)
        Me.tabRptType.Name = "tabRptType"
        Me.tabRptType.Size = New System.Drawing.Size(525, 223)
        Me.tabRptType.TabIndex = 2
        Me.tabRptType.Text = "Report Type"
        Me.tabRptType.UseVisualStyleBackColor = True
        '
        'btnRptCancel
        '
        Me.btnRptCancel.Location = New System.Drawing.Point(434, 185)
        Me.btnRptCancel.Name = "btnRptCancel"
        Me.btnRptCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnRptCancel.TabIndex = 6
        Me.btnRptCancel.Text = "&Cancel"
        Me.btnRptCancel.UseVisualStyleBackColor = True
        '
        'btnRptSave
        '
        Me.btnRptSave.Location = New System.Drawing.Point(353, 185)
        Me.btnRptSave.Name = "btnRptSave"
        Me.btnRptSave.Size = New System.Drawing.Size(75, 23)
        Me.btnRptSave.TabIndex = 5
        Me.btnRptSave.Text = "&Save"
        Me.btnRptSave.UseVisualStyleBackColor = True
        '
        'chkZip
        '
        Me.chkZip.AutoSize = True
        Me.chkZip.Location = New System.Drawing.Point(21, 111)
        Me.chkZip.Name = "chkZip"
        Me.chkZip.Size = New System.Drawing.Size(46, 17)
        Me.chkZip.TabIndex = 4
        Me.chkZip.Text = "ZIP"
        Me.chkZip.UseVisualStyleBackColor = True
        '
        'chkTxt
        '
        Me.chkTxt.AutoSize = True
        Me.chkTxt.Location = New System.Drawing.Point(21, 88)
        Me.chkTxt.Name = "chkTxt"
        Me.chkTxt.Size = New System.Drawing.Size(55, 17)
        Me.chkTxt.TabIndex = 3
        Me.chkTxt.Text = "TEXT"
        Me.chkTxt.UseVisualStyleBackColor = True
        '
        'chkPDF
        '
        Me.chkPDF.AutoSize = True
        Me.chkPDF.Enabled = False
        Me.chkPDF.Location = New System.Drawing.Point(21, 65)
        Me.chkPDF.Name = "chkPDF"
        Me.chkPDF.Size = New System.Drawing.Size(48, 17)
        Me.chkPDF.TabIndex = 2
        Me.chkPDF.Text = "PDF"
        Me.chkPDF.UseVisualStyleBackColor = True
        '
        'chkExcel
        '
        Me.chkExcel.AutoSize = True
        Me.chkExcel.Location = New System.Drawing.Point(21, 42)
        Me.chkExcel.Name = "chkExcel"
        Me.chkExcel.Size = New System.Drawing.Size(63, 17)
        Me.chkExcel.TabIndex = 1
        Me.chkExcel.Text = "EXCEL"
        Me.chkExcel.UseVisualStyleBackColor = True
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.Location = New System.Drawing.Point(21, 19)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(46, 17)
        Me.chkAll.TabIndex = 0
        Me.chkAll.Text = "ALL"
        Me.chkAll.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(470, 264)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 9
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(562, 299)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.tabOptions)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Options"
        Me.tabOptions.ResumeLayout(False)
        Me.tabGnlSetting.ResumeLayout(False)
        Me.tabGnlSetting.PerformLayout()
        Me.tabODBC.ResumeLayout(False)
        Me.tabODBC.PerformLayout()
        Me.tabRptType.ResumeLayout(False)
        Me.tabRptType.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabOptions As System.Windows.Forms.TabControl
    Friend WithEvents tabGnlSetting As System.Windows.Forms.TabPage
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtExportPath As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tabODBC As System.Windows.Forms.TabPage
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents txtGenID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSMTP As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtInterval As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtLogPath As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtTimeout As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtLogin As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents btnGCancel As System.Windows.Forms.Button
    Friend WithEvents btnGSave As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnOCancel As System.Windows.Forms.Button
    Friend WithEvents btnOSave As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtDB As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtDuration As System.Windows.Forms.TextBox
    Friend WithEvents tabRptType As System.Windows.Forms.TabPage
    Friend WithEvents chkZip As System.Windows.Forms.CheckBox
    Friend WithEvents chkTxt As System.Windows.Forms.CheckBox
    Friend WithEvents chkPDF As System.Windows.Forms.CheckBox
    Friend WithEvents chkExcel As System.Windows.Forms.CheckBox
    Friend WithEvents chkAll As System.Windows.Forms.CheckBox
    Friend WithEvents btnRptSave As System.Windows.Forms.Button
    Friend WithEvents btnRptCancel As System.Windows.Forms.Button
End Class
