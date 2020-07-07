Public Class RptUnpaidList

    Function RptUnpaidList(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common

        If ds.Tables(1).Rows.Count > 0 Then
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

                ' ----------------------------------------------------------------------
                ' Define the starting row and column number of the detail header
                ' ----------------------------------------------------------------------

                iRow = 1
                iCol = 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Get File Name, Report Header
                ' ----------------------------------------------------------------------

                TblIndex = 0

                With ds.Tables(TblIndex).Rows(0)
                    fileName = .Item("RptFile").ToString
                    objExcel.Cells(1, 1) = .Item("BrhName").ToString
                    objExcel.Cells(2, 2) = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString

                    objExcel.Cells(5, 1) = "UNPAID LIST"

                    'Setting - bold
                    objExcel.Range("A1:G5").Font.Bold = True
                    objExcel.Range("A1:G5").HorizontalAlignment = -4108
                    objExcel.Range("A1:G1").Merge()
                    objExcel.Range("A2:G2").Merge()
                    objExcel.Range("A3:G3").Merge()
                    objExcel.Range("A5:G5").Merge()

                    iRow = 7
                End With

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Export Report Data
                ' ----------------------------------------------------------------------

                TblIndex += 1

                objExcel.Cells(iRow, iCol).value = "CUSTOMER"
                objExcel.Cells(iRow, iCol + 1).value = "CONSIGNEE"
                objExcel.Cells(iRow, iCol + 2).value = "TEL"
                objExcel.Cells(iRow, iCol + 3).value = "INVOICE NO."
                objExcel.Cells(iRow, iCol + 4).value = "HOUSE_BL"
                objExcel.Cells(iRow, iCol + 5).value = "ETA"
                objExcel.Cells(iRow, iCol + 6).value = "ORIGINAL AMOUNT"
                objExcel.Cells(iRow, iCol + 7).value = "OUTSTANDING AMOUNT"
                objExcel.Cells(iRow, iCol + 8).value = "GROUP"
                objExcel.Cells(iRow, iCol + 9).value = "PTERMS"
                objExcel.Cells(iRow, iCol + 10).value = "USER"
                objExcel.Cells(iRow, iCol + 11).value = "WK"
                objExcel.Cells(iRow, iCol + 12).value = "REMARK"

                ' setting border
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(8).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(10).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(11).LineStyle = 1

                iRow = iRow + 1
                startRow = iRow
                i = 0

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("Shipper").ToString
                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("Consignee").ToString
                    objExcel.Cells(iRow, iCol + 2).value = "'" & ds.Tables(TblIndex).Rows(i).Item("Tel").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("IvhInvNo").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("HouseBL").ToString
                    objExcel.Cells(iRow, iCol + 5).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("ETA")), "yyyy/MM/dd")
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("OriginalAmt").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("OutStanding").ToString
                    objExcel.Cells(iRow, iCol + 8).value = ds.Tables(TblIndex).Rows(i).Item("Grp").ToString
                    objExcel.Cells(iRow, iCol + 9).value = ds.Tables(TblIndex).Rows(i).Item("PTerms").ToString
                    objExcel.Cells(iRow, iCol + 10).value = ds.Tables(TblIndex).Rows(i).Item("Usr").ToString
                    objExcel.Cells(iRow, iCol + 11).value = ds.Tables(TblIndex).Rows(i).Item("Wk").ToString
                    objExcel.Cells(iRow, iCol + 12).value = ds.Tables(TblIndex).Rows(i).Item("IvhRemark").ToString
                    
                    iRow += 1
                Next

                iRow = iRow + 1

                objExcel.Cells(iRow, 5).value = "TOTAL AMOUNT"
                objExcel.Range(objExcel.Cells(iRow, 7), objExcel.Cells(iRow, 8)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"

                ' Setting - bold & underline
                objExcel.Range(objExcel.Cells(startRow, 5), objExcel.Cells(iRow - 2, 5)).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Range(objExcel.Cells(startRow, 7), objExcel.Cells(iRow, 7)).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Range(objExcel.Cells(startRow, 8), objExcel.Cells(iRow, 8)).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, 8)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 7), objExcel.Cells(iRow, 8)).Borders(9).LineStyle = -4119

                ' setting width
                objExcel.Columns("A:B").ColumnWidth = 40
                objExcel.Columns("C:F").ColumnWidth = 15
                objExcel.Columns("G:H").ColumnWidth = 24
                objExcel.Columns("I:J").ColumnWidth = 12
                objExcel.Columns("K:K").ColumnWidth = 16
                objExcel.Columns("L:L").ColumnWidth = 5
                objExcel.Columns("M:M").ColumnWidth = 40

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Save File
                ' ----------------------------------------------------------------------

                Dim exportPath As String = My.Settings.ExportPath
                Dim exportFile As String = ""

                If fileName <> "" Then
                    exportFile = exportPath & fileName & ".xls"
                    fileName &= ".xls"
                Else
                    exportFile = exportPath & UID & ".xls"
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

                ' **********************************************************************

                ' Destroy Variables
                objWS = Nothing
                objWB = Nothing
                objExcel = Nothing
                exportFile = Nothing
                exportPath = Nothing
                i = Nothing
                iRow = Nothing
                iSRow = Nothing
                iCol = Nothing
                startRow = Nothing
                TblIndex = Nothing

                ' Release Memory
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objExcel.ActiveWorkbook.SaveAs("C:\" & UID & ".xls")
                objExcel.Quit()
                fileName = "Error," & ex.Message
            End Try
        Else
            fileName = ""
        End If

        ' Return File Path
        RptUnpaidList = fileName
    End Function
End Class
