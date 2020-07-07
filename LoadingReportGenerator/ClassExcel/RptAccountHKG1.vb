Public Class RptAccountHKG1

    Public Function RptAccount1(ByVal uid As String, ByVal sUid As String, ByVal ds As DataSet) As String

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
        Dim PrevDocNo As String = ""
        Dim totalAmt As Decimal

        totalAmt = 0

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

            For j = 1 To 6
                PrevDocNo = ""

                If ds.Tables(j).Rows.Count > 0 Then
                    hasData = True

                    If ds.Tables(j).Rows(0).Item("SetType").ToString <> "" Then
                        setType = ds.Tables(j).Rows(0).Item("SetType").ToString

                        objExcel.Cells(iRow, iCol) = "Ledger Code"
                        objExcel.Cells(iRow, iCol + 1) = "Batch"
                        objExcel.Cells(iRow, iCol + 2) = "Period"
                        objExcel.Cells(iRow, iCol + 3) = "Voucher"
                        objExcel.Cells(iRow, iCol + 4) = "Voucher Date"
                        objExcel.Cells(iRow, iCol + 5) = "Voucher Description"
                        objExcel.Cells(iRow, iCol + 6) = "A/C Code"
                        objExcel.Cells(iRow, iCol + 7) = "Ana.1"
                        objExcel.Cells(iRow, iCol + 8) = "Ana.2"
                        objExcel.Cells(iRow, iCol + 9) = "Currency"
                        objExcel.Cells(iRow, iCol + 10) = "Orig.Amount"
                        objExcel.Cells(iRow, iCol + 11) = "Equv.Amount"
                        objExcel.Cells(iRow, iCol + 12) = "ExRate"
                        objExcel.Cells(iRow, iCol + 13) = "Docu.Type"
                        objExcel.Cells(iRow, iCol + 14) = "Docu.No."
                        objExcel.Cells(iRow, iCol + 15) = "Docu.Date"
                        objExcel.Cells(iRow, iCol + 16) = "Pay.Term"
                        objExcel.Cells(iRow, iCol + 17) = "Due Date"
                        objExcel.Cells(iRow, iCol + 18) = "Particulars 1"
                        objExcel.Cells(iRow, iCol + 19) = "Particulars 2"
                        objExcel.Cells(iRow, iCol + 20) = "Revise"
                        objExcel.Cells(iRow, iCol + 21) = "Week"

                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Interior.ColorIndex = 15
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Borders(8).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Borders(10).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Borders(11).LineStyle = 1

                        iRow = iRow + 1
                        iCount = 1
                        PrevDocNo = ""

                        MaxRow = ds.Tables(j).Rows.Count - 1

                        For i = 0 To MaxRow
                            'Set Content
                            If PrevDocNo <> ds.Tables(j).Rows(i).Item("DocuNo").ToString Then
                                If PrevDocNo <> "" Then
                                    iCount = iCount + 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 21)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                                End If

                                PrevDocNo = ds.Tables(j).Rows(i).Item("DocuNo").ToString
                            End If

                            objExcel.Cells(iRow, iCol) = ds.Tables(j).Rows(i).Item("Ledger").ToString
                            objExcel.Cells(iRow, iCol + 1) = ds.Tables(j).Rows(i).Item("Batch").ToString
                            objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(j).Rows(i).Item("Period").ToString
                            objExcel.Cells(iRow, iCol + 3) = "'" & ds.Tables(j).Rows(i).Item("Voucher").ToString & iCount
                            If common.NullVal(ds.Tables(j).Rows(i).Item("VoucherDte"), "") <> "" Then
                                objExcel.Cells(iRow, iCol + 4) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("VoucherDte")), "dd/MM/yyyy")
                            End If

                            objExcel.Cells(iRow, iCol + 5) = ds.Tables(j).Rows(i).Item("VouDesc").ToString
                            objExcel.Cells(iRow, iCol + 6) = "'" & ds.Tables(j).Rows(i).Item("AccCode").ToString
                            objExcel.Cells(iRow, iCol + 7) = ds.Tables(j).Rows(i).Item("Ana1").ToString
                            objExcel.Cells(iRow, iCol + 8) = ds.Tables(j).Rows(i).Item("Ana2").ToString
                            objExcel.Cells(iRow, iCol + 9) = ds.Tables(j).Rows(i).Item("Currency").ToString

                            objExcel.Cells(iRow, iCol + 10) = ds.Tables(j).Rows(i).Item("OrigAmt").ToString
                            objExcel.Cells(iRow, iCol + 11) = ds.Tables(j).Rows(i).Item("EquvAmt").ToString
                            objExcel.Cells(iRow, iCol + 12) = ds.Tables(j).Rows(i).Item("ExRate").ToString
                            objExcel.Cells(iRow, iCol + 13) = ds.Tables(j).Rows(i).Item("DocType").ToString
                            objExcel.Cells(iRow, iCol + 14) = ds.Tables(j).Rows(i).Item("DocuNo").ToString

                            If common.NullVal(ds.Tables(j).Rows(i).Item("DocDte"), "") <> "" Then
                                objExcel.Cells(iRow, iCol + 15) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DocDte")), "dd/MM/yyyy")
                            End If
                            objExcel.Cells(iRow, iCol + 16) = ds.Tables(j).Rows(i).Item("PayTerm").ToString
                            If common.NullVal(ds.Tables(j).Rows(i).Item("DueDte"), "") <> "" Then
                                objExcel.Cells(iRow, iCol + 17) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DueDte")), "dd/MM/yyyy")
                            End If
                            objExcel.Cells(iRow, iCol + 18) = ds.Tables(j).Rows(i).Item("Part1").ToString
                            objExcel.Cells(iRow, iCol + 19) = ds.Tables(j).Rows(i).Item("Part2").ToString

                            objExcel.Cells(iRow, iCol + 20) = ds.Tables(j).Rows(i).Item("Revise").ToString
                            objExcel.Cells(iRow, iCol + 21) = ds.Tables(j).Rows(i).Item("BkhWeek").ToString

                            rowLn = ds.Tables(j).Rows(i).Item("BkvLn").ToString

                            totalAmt = totalAmt + CDbl(ds.Tables(j).Rows(i).Item("EquvAmt").ToString)

                            If ds.Tables(j).Rows(i).Item("BkvLn") = "summary" Then
                                If CDbl(ds.Tables(j).Rows(i).Item("OrigAmt").ToString) = 0 And j = 2 Then 'If FL part and them Amt is 0, not show
                                    ' If Account Code 6210 at the end, no need to move back
                                    If ds.Tables(j).Rows(i).Item("AccCode").ToString <> "6210" Then
                                        iRow = iRow - 1
                                    End If
                                End If

                                If FormatNumber(totalAmt, "2") <> 0.0 And j = 2 Then
                                    iRow = iRow + 1

                                    objExcel.Cells(iRow, iCol) = ds.Tables(j).Rows(i).Item("Ledger").ToString
                                    objExcel.Cells(iRow, iCol + 1) = ds.Tables(j).Rows(i).Item("Batch").ToString
                                    objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(j).Rows(i).Item("Period").ToString
                                    objExcel.Cells(iRow, iCol + 3) = "'" & ds.Tables(j).Rows(i).Item("Voucher").ToString & iCount

                                    If common.NullVal(ds.Tables(j).Rows(i).Item("VoucherDte"), "") <> "" Then
                                        objExcel.Cells(iRow, iCol + 4) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("VoucherDte")), "dd/MM/yyyy")
                                    End If

                                    objExcel.Cells(iRow, iCol + 5) = ds.Tables(j).Rows(i).Item("VouDesc").ToString
                                    objExcel.Cells(iRow, iCol + 6) = "'7360"
                                    objExcel.Cells(iRow, iCol + 7) = ds.Tables(j).Rows(i).Item("Ana1").ToString
                                    objExcel.Cells(iRow, iCol + 8) = ds.Tables(j).Rows(i).Item("Ana2").ToString
                                    objExcel.Cells(iRow, iCol + 9) = "RMB"

                                    objExcel.Cells(iRow, iCol + 10) = 0
                                    objExcel.Cells(iRow, iCol + 11) = (-1 * FormatNumber(totalAmt, "2"))
                                    objExcel.Cells(iRow, iCol + 12) = 1.0 'ds.Tables(j).Rows(i).Item("ExRate").ToString
                                    objExcel.Cells(iRow, iCol + 13) = ds.Tables(j).Rows(i).Item("DocType").ToString
                                    objExcel.Cells(iRow, iCol + 14) = ds.Tables(j).Rows(i).Item("DocuNo").ToString

                                    If common.NullVal(ds.Tables(j).Rows(i).Item("DocDte"), "") <> "" Then
                                        objExcel.Cells(iRow, iCol + 15) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DocDte")), "dd/MM/yyyy")
                                    End If
                                    objExcel.Cells(iRow, iCol + 16) = ds.Tables(j).Rows(i).Item("PayTerm").ToString
                                    If common.NullVal(ds.Tables(j).Rows(i).Item("DueDte"), "") <> "" Then
                                        objExcel.Cells(iRow, iCol + 17) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DueDte")), "dd/MM/yyyy")
                                    End If
                                    objExcel.Cells(iRow, iCol + 18) = ds.Tables(j).Rows(i).Item("Part1").ToString
                                    objExcel.Cells(iRow, iCol + 19) = ds.Tables(j).Rows(i).Item("Part2").ToString

                                    objExcel.Cells(iRow, iCol + 20) = ds.Tables(j).Rows(i).Item("Revise").ToString
                                    objExcel.Cells(iRow, iCol + 21) = ds.Tables(j).Rows(i).Item("BkhWeek").ToString

                                    totalAmt = 0
                                End If

                                If j <> 2 Then
                                    i = i + 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 20)).Borders(9).LineStyle = 1
                                End If
                            End If

                            iRow = iRow + 1
                        Next
                    End If
                End If

                iRow = iRow + 2
            Next

            objExcel.Range(objExcel.Cells(2, iCol + 10), objExcel.Cells(iRow, iCol + 10)).NumberFormatLocal = "#,##0.00_ "
            objExcel.Range(objExcel.Cells(2, iCol + 11), objExcel.Cells(iRow, iCol + 11)).NumberFormatLocal = "#,##0.00_ "
            objExcel.Range(objExcel.Cells(2, iCol + 12), objExcel.Cells(iRow, iCol + 12)).NumberFormatLocal = "#,##0.000000_ "

            objExcel.Columns("F:F").ColumnWidth = 15
            objExcel.Columns("O:O").ColumnWidth = 11
            objExcel.Columns("S:S").ColumnWidth = 54
            objExcel.Columns("T:T").ColumnWidth = 11

            ' ----------------------------------------------------------------------
            ' Save File
            ' ----------------------------------------------------------------------
            Dim exportPath As String = My.Settings.ExportPath & sUid & "\"
            Dim exportFile As String = ""

            If fileName <> "" Then
                exportFile = exportPath & fileName & ".xls"
            Else
                exportFile = exportPath & uid & ".xls"
            End If

            'fileName &= ".xls"

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
            'MsgBox(ex.Message)
            objExcel.ActiveWorkbook.SaveAs("C:\" & uid & ".xls")
            objExcel.Quit()
            fileName = "Error," & ex.Message
        End Try

        ' Return File Path
        RptAccount1 = fileName

    End Function

End Class
