Public Class RptAutoVoucher

    Public Function RptAutoVoucher(ByVal Uid As String, ByVal sUid As String, ByVal ds As DataSet) As String

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, iCount, MaxRow As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim setType As String = ""
        Dim common As New common
        Dim hasData As Boolean = False
        Dim rowLn As String = ""

        ' Start Excel Application
        objExcel = CreateObject("Excel.Application")
        objExcel.Visible = False

        Try
            ' Get a new workbook
            objWB = objExcel.Workbooks.Add
            objWS = objWB.ActiveSheet

            ' Set Worksheet Properties
            objWS.Application.Cells.Font.Name = "Verdana"
            objWS.Application.Cells.Font.Size = 9
            objWS.Application.Cells.VerticalAlignment = -4160

            ' Define the starting row and column number of the detail header
            iRow = 1
            iCol = 1

            ' Retrieve File Name
            If ds.Tables(0).Rows.Count > 0 Then
                fileName = ds.Tables(0).Rows(0).Item("fName").ToString
            End If

            ' Export Data to Excel
            If ds.Tables(1).Rows.Count > 0 Then
                hasData = True

                objExcel.Cells(iRow, iCol) = "Voucher No"
                objExcel.Cells(iRow, iCol + 1) = "HBL No"
                objExcel.Cells(iRow, iCol + 2) = "Agent Name"
                objExcel.Cells(iRow, iCol + 3) = "Currency"
                objExcel.Cells(iRow, iCol + 4) = "Amount"
                objExcel.Cells(iRow, iCol + 5) = "Last Update Date"
                objExcel.Cells(iRow, iCol + 6) = "Last Update User"
                objExcel.Cells(iRow, iCol + 7) = "Week No"

                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Borders(8).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Borders(10).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Borders(11).LineStyle = 1

                iRow = iRow + 1
                iCount = 0

                MaxRow = ds.Tables(1).Rows.Count - 1

                For i = 0 To MaxRow
                    'Set Content
                    objExcel.Cells(iRow, iCol) = ds.Tables(1).Rows(i).Item("VohVouNo").ToString
                    objExcel.Cells(iRow, iCol + 1) = "'" & ds.Tables(1).Rows(i).Item("BkhBLNo").ToString
                    objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(1).Rows(i).Item("AgentName").ToString
                    objExcel.Cells(iRow, iCol + 3) = ds.Tables(1).Rows(i).Item("BkvCurr").ToString
                    objExcel.Cells(iRow, iCol + 4) = ds.Tables(1).Rows(i).Item("VouAmt").ToString
                    If common.NullVal(ds.Tables(1).Rows(i).Item("VohLstUpd"), "") <> "" Then
                        objExcel.Cells(iRow, iCol + 5) = "'" & Format(Convert.ToDateTime(ds.Tables(1).Rows(i).Item("VohLstUpd")), "dd/MM/yyyy")
                    End If
                    objExcel.Cells(iRow, iCol + 6) = ds.Tables(1).Rows(i).Item("VohLstUsrName").ToString
                    objExcel.Cells(iRow, iCol + 7) = "'" & ds.Tables(1).Rows(i).Item("BkhWeek").ToString

                    iRow = iRow + 1
                Next
            End If

            iRow = iRow + 2

            objExcel.Range(objExcel.Cells(2, iCol + 4), objExcel.Cells(iRow, iCol + 4)).NumberFormatLocal = "#,##0.00_ "

            objExcel.Columns("A:A").ColumnWidth = 12
            objExcel.Columns("B:B").ColumnWidth = 13
            objExcel.Columns("C:C").ColumnWidth = 30
            objExcel.Columns("F:G").ColumnWidth = 15

            ' ----------------------------------------------------------------------
            ' Save File
            ' ----------------------------------------------------------------------
            Dim exportPath As String = My.Settings.ExportPath & sUid & "\"
            Dim exportFile As String = ""

            If fileName <> "" Then
                exportFile = exportPath & fileName & ".xls"
            Else
                exportFile = exportPath & Uid & ".xls"
            End If

            ' Create if export directory not found
            If Not My.Computer.FileSystem.DirectoryExists(exportPath) Then
                My.Computer.FileSystem.CreateDirectory(exportPath)
            End If

            ' Delete if file already exists
            If My.Computer.FileSystem.FileExists(exportFile) Then
                My.Computer.FileSystem.DeleteFile(exportFile)
            End If

            objWS.SaveAs(exportFile)
            objWS.Application.Quit()

            ' Destroy Variables
            objWS = Nothing
            objWB = Nothing
            objExcel = Nothing
            exportFile = Nothing
            exportPath = Nothing
            iCount = Nothing
            j = Nothing
            setType = Nothing
            i = Nothing
            iRow = Nothing
            iSRow = Nothing
            iCol = Nothing

            ' Release Memory
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception

            objExcel.ActiveWorkbook.SaveAs("C:\" & Uid & ".xls")
            objExcel.Quit()
            fileName = "Error," & ex.Message
        End Try

        ' Return File Path
        RptAutoVoucher = fileName

    End Function

End Class
