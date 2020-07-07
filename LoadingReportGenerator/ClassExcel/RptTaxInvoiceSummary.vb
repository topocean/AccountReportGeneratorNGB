Public Class RptTaxInvoiceSummary

    Function RptTaxInvoiceSummary(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim newRow As Integer
        Dim hkdCount As Double
        Dim rmbCount As Double
        Dim usdCount As Double

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

                fileName = ds.Tables(0).Rows(0).Item("RptFile")

                objExcel.Range("A" & iRow & ":I" & iRow).Merge()
                objExcel.Cells(iRow, iCol).value = "China Invoice Detail(中國發票明細表)"
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, 1)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, 1)).Font.Size = 14
                iRow = iRow + 1

                objExcel.Cells(iRow, 1).value = "中國發票號碼"
                objExcel.Cells(iRow, 2).value = "客戶抬頭"
                objExcel.Cells(iRow, 3).value = "帳戶號碼"
                objExcel.Cells(iRow, 4).value = "提單號"
                objExcel.Cells(iRow, 5).value = "LOT NO"
                objExcel.Cells(iRow, 6).value = "船名/班次/航班"
                objExcel.Cells(iRow, 7).value = "開航日期"
                objExcel.Cells(iRow, 8).value = "開單號"
                objExcel.Cells(iRow, 9).value = "開發票日期"
                objExcel.Cells(iRow, 10).value = "銀碼(RMB)"
                objExcel.Cells(iRow, 11).value = "銀碼(USD)"
                objExcel.Cells(iRow, 12).value = "銀碼(HKD)"
                objExcel.Cells(iRow, 13).value = "Equivalent RMB"
                objExcel.Cells(iRow, 14).value = "作廢日期"
                objExcel.Cells(iRow, 15).value = "中國發票號碼(隔月作廢)"
                objExcel.Cells(iRow, 16).value = "作廢金額"

                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Font.Bold = True

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Report Data
                ' ----------------------------------------------------------------------

                hkdCount = 0
                rmbCount = 0
                usdCount = 0

                objExcel.Columns("G:G").NumberFormatLocal = "dd/MM/YYYY"
                objExcel.Columns("I:I").NumberFormatLocal = "dd/MM/YYYY"
                objExcel.Columns("J:J").NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Columns("K:K").NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Columns("L:L").NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Columns("O:O").NumberFormatLocal = "#,###,##0.00_ "

                TblIndex = 1

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    iRow = iRow + 1
                    objExcel.Cells(iRow, 1).value = "'" & ds.Tables(TblIndex).Rows(i).Item("TxiInvNo").ToString
                    objExcel.Cells(iRow, 2).value = ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString
                    objExcel.Cells(iRow, 3).value = ds.Tables(TblIndex).Rows(i).Item("ShpAccCd").ToString
                    objExcel.Cells(iRow, 4).value = ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                    objExcel.Cells(iRow, 5).value = ds.Tables(TblIndex).Rows(i).Item("BkhLotNo").ToString

                    objExcel.Cells(iRow, 6).value = ds.Tables(TblIndex).Rows(i).Item("VslName").ToString & " / " & ds.Tables(TblIndex).Rows(i).Item("VoyName").ToString

                    If common.NullVal(ds.Tables(TblIndex).Rows(i).Item("BkhETD"), "") = "" Then
                        objExcel.Cells(iRow, 7).value = "'"
                    Else
                        objExcel.Cells(iRow, 7).value = "'" & ds.Tables(TblIndex).Rows(i).Item("BkhETD")
                    End If

                    objExcel.Cells(iRow, 8).value = ds.Tables(TblIndex).Rows(i).Item("BkhMBLNo").ToString

                    If common.NullVal(ds.Tables(TblIndex).Rows(i).Item("IssueDte"), "") = "" Then
                        objExcel.Cells(iRow, 9).value = "'"
                    Else
                        objExcel.Cells(iRow, 9).value = "'" & ds.Tables(TblIndex).Rows(i).Item("IssueDte")
                    End If

                    objExcel.Cells(iRow, 10).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("RMBAmt").ToString, 2)
                    objExcel.Cells(iRow, 11).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("USDAmt").ToString, 2)
                    objExcel.Cells(iRow, 12).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("HKDAmt").ToString, 2)

                    objExcel.Cells(iRow, 13).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ToRMB").ToString, 2)

                    If common.NullVal(ds.Tables(TblIndex).Rows(i).Item("LstVoidDte"), "") = "" Then
                        objExcel.Cells(iRow, 14).value = "'"
                    Else
                        objExcel.Cells(iRow, 14).value = "'" & ds.Tables(TblIndex).Rows(i).Item("LstVoidDte")
                    End If

                    objExcel.Cells(iRow, 15).value = ds.Tables(TblIndex).Rows(i).Item("LstVoidInv").ToString
                    objExcel.Cells(iRow, 16).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("LstVoidAmt").ToString, 2)

                    hkdCount = hkdCount + Convert.ToDouble(ds.Tables(TblIndex).Rows(i).Item("HKDAmt"))
                    rmbCount = rmbCount + Convert.ToDouble(ds.Tables(TblIndex).Rows(i).Item("RMBAmt"))
                    usdCount = usdCount + Convert.ToDouble(ds.Tables(TblIndex).Rows(i).Item("USDAmt"))
                Next

                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(iRow, iCol + 15)).Borders(7).LineStyle = 1
                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(iRow, iCol + 15)).Borders(8).LineStyle = 1
                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(iRow, iCol + 15)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(iRow, iCol + 15)).Borders(10).LineStyle = 1
                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(iRow, iCol + 15)).Borders(11).LineStyle = 1
                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(iRow, iCol + 15)).Borders(12).LineStyle = 1

                objExcel.Cells(iRow + 1, 10).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Cells(iRow + 1, 11).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Cells(iRow + 1, 12).NumberFormatLocal = "#,###,##0.00_ "

                objExcel.Cells(iRow + 1, 9).value = "Total Amount"
                objExcel.Cells(iRow + 1, 10) = FormatNumber(rmbCount, 2)
                objExcel.Cells(iRow + 1, 11) = FormatNumber(usdCount, 2)
                objExcel.Cells(iRow + 1, 12) = FormatNumber(hkdCount, 2)

                objExcel.Range(objExcel.Cells(2, 1), objExcel.Cells(2, iCol + 15)).Borders(9).LineStyle = -4119

                ' Voided Tax Invoice
                TblIndex += 1
                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    newRow = iRow + 5
                    iRow = iRow + 5

                    objExcel.Cells(iRow, 1).value = "本月作廢"
                    iRow = iRow + 1

                    objExcel.Cells(iRow, 1).value = "中國發票號碼"
                    objExcel.Cells(iRow, 2).value = "客戶抬頭"
                    objExcel.Cells(iRow, 3).value = "帳戶號碼"
                    objExcel.Cells(iRow, 4).value = "提單號"
                    objExcel.Cells(iRow, 5).value = "LOT NO"
                    objExcel.Cells(iRow, 6).value = "船名/班次/航班"
                    objExcel.Cells(iRow, 7).value = "開航日期"
                    objExcel.Cells(iRow, 8).value = "開單號"
                    objExcel.Cells(iRow, 9).value = "開發票日期"
                    objExcel.Cells(iRow, 10).value = "銀碼(RMB)"
                    objExcel.Cells(iRow, 11).value = "銀碼(USD)"
                    objExcel.Cells(iRow, 12).value = "銀碼(HKD)"
                    objExcel.Cells(iRow, 13).value = "作廢日期"
                    objExcel.Cells(iRow, 14).value = "作廢職員"

                    objExcel.Range(objExcel.Cells(newRow, 1), objExcel.Cells(iRow, iCol + 13)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 13)).Font.Bold = True

                    '-------Detail
                    hkdCount = 0
                    rmbCount = 0
                    usdCount = 0

                    objExcel.Columns("G:G").NumberFormatLocal = "DD/MM/YYYY"
                    objExcel.Columns("I:I").NumberFormatLocal = "DD/MM/YYYY"
                    objExcel.Columns("J:J").NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Columns("K:K").NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Columns("L:L").NumberFormatLocal = "#,###,##0.00_ "

                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        iRow = iRow + 1
                        objExcel.Cells(iRow, 1).value = ds.Tables(TblIndex).Rows(i).Item("TxiInvNo").ToString
                        objExcel.Cells(iRow, 2).value = ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString
                        objExcel.Cells(iRow, 3).value = ds.Tables(TblIndex).Rows(i).Item("ShpAccCd").ToString
                        objExcel.Cells(iRow, 4).value = ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                        objExcel.Cells(iRow, 5).value = ds.Tables(TblIndex).Rows(i).Item("BkhLotNo").ToString
                        objExcel.Cells(iRow, 6).value = ds.Tables(TblIndex).Rows(i).Item("VslName").ToString & "/" & ds.Tables(TblIndex).Rows(i).Item("VoyName").ToString

                        If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("BkhETD")) Then
                            objExcel.Cells(iRow, 7).value = "'"
                        Else
                            objExcel.Cells(iRow, 7).value = "'" & ds.Tables(TblIndex).Rows(i).Item("BkhETD")
                        End If
                        objExcel.Cells(iRow, 8).value = ds.Tables(TblIndex).Rows(i).Item("BkhMBLNo").ToString
                        If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("IssueDte")) Then
                            objExcel.Cells(iRow, 9).value = "'"
                        Else
                            objExcel.Cells(iRow, 9).value = "'" & ds.Tables(TblIndex).Rows(i).Item("IssueDte")
                        End If
                        objExcel.Cells(iRow, 10).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("RMBAmt").ToString, 2)
                        objExcel.Cells(iRow, 11).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("USDAmt").ToString, 2)
                        objExcel.Cells(iRow, 12).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("HKDAmt").ToString, 2)
                        If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("VoidDte")) Then
                            objExcel.Cells(iRow, 13).value = "'"
                        Else
                            objExcel.Cells(iRow, 13).value = "'" & ds.Tables(TblIndex).Rows(i).Item("VoidDte")
                        End If
                        objExcel.Cells(iRow, 14) = ds.Tables(TblIndex).Rows(i).Item("VoidUsr").ToString

                        hkdCount = hkdCount + Convert.ToDouble(ds.Tables(TblIndex).Rows(i).Item("HKDAmt"))
                        rmbCount = rmbCount + Convert.ToDouble(ds.Tables(TblIndex).Rows(i).Item("RMBAmt"))
                        usdCount = usdCount + Convert.ToDouble(ds.Tables(TblIndex).Rows(i).Item("USDAmt"))
                    Next

                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(iRow, iCol + 13)).Borders(7).LineStyle = 1
                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(iRow, iCol + 13)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(iRow, iCol + 13)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(iRow, iCol + 13)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(iRow, iCol + 13)).Borders(11).LineStyle = 1
                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(iRow, iCol + 13)).Borders(12).LineStyle = 1

                    objExcel.Range(objExcel.Cells(newRow + 1, 1), objExcel.Cells(2, iCol + 13)).Borders(9).LineStyle = -4119

                    objExcel.Cells(iRow + 1, 10).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Cells(iRow + 1, 11).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Cells(iRow + 1, 12).NumberFormatLocal = "#,###,##0.00_ "

                    objExcel.Cells(iRow + 1, 9).value = "Total Amount"
                    objExcel.Cells(iRow + 1, 10) = FormatNumber(rmbCount, 2)
                    objExcel.Cells(iRow + 1, 11) = FormatNumber(usdCount, 2)
                    objExcel.Cells(iRow + 1, 12) = FormatNumber(hkdCount, 2)

                End If

                objExcel.Columns("A:A").ColumnWidth = 14.5
                objExcel.Columns("B:B").ColumnWidth = 50
                objExcel.Columns("C:E").ColumnWidth = 17
                objExcel.Columns("F:F").ColumnWidth = 30
                objExcel.Columns("G:G").ColumnWidth = 15
                objExcel.Columns("H:H").ColumnWidth = 23.5
                objExcel.Columns("I:I").ColumnWidth = 15
                objExcel.Columns("J:L").ColumnWidth = 11.5
                objExcel.Columns("M:P").ColumnWidth = 22

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
        RptTaxInvoiceSummary = fileName
    End Function
End Class
