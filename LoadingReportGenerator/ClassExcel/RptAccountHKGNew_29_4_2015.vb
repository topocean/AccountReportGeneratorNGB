Public Class RptAccountHKGNew

    Public HasData As Boolean

    Property rptHasDataOccur()

        Get
            rptHasDataOccur = Me.HasData
        End Get
        Set(ByVal value)
            Me.HasData = value
        End Set

    End Property

    Public Function RptAccount(ByVal Uid As String, ByVal sUid As String, ByVal ds As DataSet) As String

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, iCount, MaxRow As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim setType As String = ""
        Dim common As New common
        Dim rowLn As String = ""
        Dim PrevDocNo As String = ""
        Dim totalAmt As Decimal
        Dim IsRevise As Integer = 0
        Dim ledger As String = "NGB"
        Dim FLX_ExDiffCurr As String = ""
        Dim CSG_ExDiffCurr As String = ""
        Dim FLX_ExDiffAccCode As String = ""
        Dim CSG_ExDiffAccCode As String = ""

        totalAmt = 0
        HasData = False

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
                IsRevise = ds.Tables(0).Rows(0).Item("IsRevise").ToString
                FLX_ExDiffCurr = ds.Tables(0).Rows(0).Item("FLX_ExDiffCurr").ToString
                FLX_ExDiffAccCode = ds.Tables(0).Rows(0).Item("FLX_ExDiffAccCode").ToString
                CSG_ExDiffCurr = ds.Tables(0).Rows(0).Item("CSG_ExDiffCurr").ToString
                CSG_ExDiffAccCode = ds.Tables(0).Rows(0).Item("CSG_ExDiffAccCode").ToString
            End If

            ' Export Data to Excel
            For j = 1 To 6
                PrevDocNo = ""

                If ds.Tables(j).Rows.Count > 0 Then

                    If ds.Tables(j).Rows(0).Item("SetType").ToString <> "" Then
                        setType = ds.Tables(j).Rows(0).Item("SetType").ToString

                        objExcel.Cells(iRow, iCol) = "Ledger Code"
                        objExcel.Cells(iRow, iCol + 1) = "Batch Number"
                        objExcel.Cells(iRow, iCol + 2) = "Account Period"
                        objExcel.Cells(iRow, iCol + 3) = "Voucher Number"
                        objExcel.Cells(iRow, iCol + 4) = "Voucher Date"
                        objExcel.Cells(iRow, iCol + 5) = "Voucher Description"
                        objExcel.Cells(iRow, iCol + 6) = "Account Code"
                        objExcel.Cells(iRow, iCol + 7) = "Analysis Code 1"
                        objExcel.Cells(iRow, iCol + 8) = "Analysis Code 2"
                        objExcel.Cells(iRow, iCol + 9) = "Analysis Code 5"
                        objExcel.Cells(iRow, iCol + 10) = "Currency Code"
                        objExcel.Cells(iRow, iCol + 11) = "Debit/Credit"
                        objExcel.Cells(iRow, iCol + 12) = "Original Amount"
                        objExcel.Cells(iRow, iCol + 13) = "Equivalent Amount"
                        objExcel.Cells(iRow, iCol + 14) = "Exchange Rate"
                        objExcel.Cells(iRow, iCol + 15) = "Document Type"
                        objExcel.Cells(iRow, iCol + 16) = "Document Number"
                        objExcel.Cells(iRow, iCol + 17) = "Document Date"
                        objExcel.Cells(iRow, iCol + 18) = "Payment Terms"
                        objExcel.Cells(iRow, iCol + 19) = "Document Due Date"
                        objExcel.Cells(iRow, iCol + 20) = "Particular 1"
                        objExcel.Cells(iRow, iCol + 21) = "Particular 2"
                        objExcel.Cells(iRow, iCol + 22) = "Open Item Number"
                        objExcel.Cells(iRow, iCol + 23) = "Applied Amount"
                        objExcel.Cells(iRow, iCol + 24) = "Alternate Doc. No."
                        objExcel.Cells(iRow, iCol + 25) = "Account Name(SC)"
                        objExcel.Cells(iRow, iCol + 26) = "Revise"
                        objExcel.Cells(iRow, iCol + 27) = "Week"

                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Interior.ColorIndex = 15
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(8).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(10).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(11).LineStyle = 1

                        iRow = iRow + 1
                        iCount = 0

                        MaxRow = ds.Tables(j).Rows.Count - 1

                        If MaxRow <> 0 Then
                            HasData = True
                        End If

                        For i = 0 To MaxRow
                            If PrevDocNo <> ds.Tables(j).Rows(i).Item("DocuNo").ToString Then

                                ' Exchange Rate Different
                                If FormatNumber(totalAmt, 2) <> 0.0 Then
                                    objExcel.Cells(iRow, iCol) = ds.Tables(j).Rows(i - 1).Item("Ledger").ToString
                                    objExcel.Cells(iRow, iCol + 1) = ds.Tables(j).Rows(i - 1).Item("Batch").ToString
                                    objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(j).Rows(i - 1).Item("Period").ToString
                                    objExcel.Cells(iRow, iCol + 3) = "'" & ds.Tables(j).Rows(i - 1).Item("Voucher").ToString & iCount

                                    If common.NullVal(ds.Tables(j).Rows(i - 1).Item("VoucherDte"), "") <> "" Then
                                        objExcel.Cells(iRow, iCol + 4) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i - 1).Item("VoucherDte")), "dd/MM/yyyy")
                                    End If

                                    objExcel.Cells(iRow, iCol + 5) = ds.Tables(j).Rows(i - 1).Item("VouDesc").ToString

                                    If Left(ds.Tables(j).Rows(i - 1).Item("Ledger").ToString, 1) = "C" Then
                                        objExcel.Cells(iRow, iCol + 6) = "'" & CSG_ExDiffAccCode
                                    Else
                                        objExcel.Cells(iRow, iCol + 6) = "'" & FLX_ExDiffAccCode
                                    End If

                                    objExcel.Cells(iRow, iCol + 7) = "'" & ds.Tables(j).Rows(i - 1).Item("Ana1").ToString
                                    objExcel.Cells(iRow, iCol + 8) = "'" & ds.Tables(j).Rows(i - 1).Item("Ana2").ToString
                                    objExcel.Cells(iRow, iCol + 9) = "'" & ds.Tables(j).Rows(i - 1).Item("Ana5").ToString

                                    If Left(ds.Tables(j).Rows(i - 1).Item("Ledger").ToString, 1) = "C" Then
                                        objExcel.Cells(iRow, iCol + 10) = "'" & CSG_ExDiffCurr
                                    Else
                                        objExcel.Cells(iRow, iCol + 10) = "'" & FLX_ExDiffCurr
                                    End If

                                    If (-1 * totalAmt) >= 0 Then
                                        objExcel.Cells(iRow, iCol + 11) = "D"
                                    Else
                                        objExcel.Cells(iRow, iCol + 11) = "C"
                                    End If

                                    objExcel.Cells(iRow, iCol + 12) = -1 * FormatNumber(totalAmt, 2)
                                    objExcel.Cells(iRow, iCol + 13) = -1 * FormatNumber(totalAmt, 2)
                                    objExcel.Cells(iRow, iCol + 14) = 1.0

                                    objExcel.Cells(iRow, iCol + 15) = ds.Tables(j).Rows(i - 1).Item("DocType").ToString
                                    objExcel.Cells(iRow, iCol + 16) = "'" & ds.Tables(j).Rows(i - 1).Item("DocuNo").ToString

                                    If common.NullVal(ds.Tables(j).Rows(i - 1).Item("DocDte"), "") <> "" Then
                                        objExcel.Cells(iRow, iCol + 17) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i - 1).Item("DocDte")), "dd/MM/yyyy")
                                    End If

                                    objExcel.Cells(iRow, iCol + 18) = ds.Tables(j).Rows(i - 1).Item("PayTerm").ToString

                                    If common.NullVal(ds.Tables(j).Rows(i - 1).Item("DueDte"), "") <> "" Then
                                        objExcel.Cells(iRow, iCol + 19) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i - 1).Item("DueDte")), "dd/MM/yyyy")
                                    End If

                                    objExcel.Cells(iRow, iCol + 20) = ds.Tables(j).Rows(i - 1).Item("Part1").ToString
                                    objExcel.Cells(iRow, iCol + 21) = "'" & ds.Tables(j).Rows(i - 1).Item("Part2").ToString
                                    objExcel.Cells(iRow, iCol + 22) = ""
                                    objExcel.Cells(iRow, iCol + 23) = ""
                                    objExcel.Cells(iRow, iCol + 24) = ds.Tables(j).Rows(i - 1).Item("AltDocNo").ToString
                                    objExcel.Cells(iRow, iCol + 25) = ds.Tables(j).Rows(i - 1).Item("BillTo").ToString
                                    objExcel.Cells(iRow, iCol + 26) = ds.Tables(j).Rows(i - 1).Item("Revise").ToString
                                    objExcel.Cells(iRow, iCol + 27) = ds.Tables(j).Rows(i - 1).Item("BkhWeek").ToString

                                    totalAmt = 0
                                    iRow = iRow + 1
                                End If

                                ' Add Report Header and 5 rows separator if Ledger Code changed
                                'If i + 1 <= MaxRow Then
                                If ds.Tables(j).Rows(i + 1).Item("Ledger").ToString <> ledger Then
                                    iRow = iRow + 5
                                    iCount = 0

                                    objExcel.Cells(iRow, iCol) = "Ledger Code"
                                    objExcel.Cells(iRow, iCol + 1) = "Batch Number"
                                    objExcel.Cells(iRow, iCol + 2) = "Account Period"
                                    objExcel.Cells(iRow, iCol + 3) = "Voucher Number"
                                    objExcel.Cells(iRow, iCol + 4) = "Voucher Date"
                                    objExcel.Cells(iRow, iCol + 5) = "Voucher Description"
                                    objExcel.Cells(iRow, iCol + 6) = "Account Code"
                                    objExcel.Cells(iRow, iCol + 7) = "Analysis Code 1"
                                    objExcel.Cells(iRow, iCol + 8) = "Analysis Code 2"
                                    objExcel.Cells(iRow, iCol + 9) = "Analysis Code 5"
                                    objExcel.Cells(iRow, iCol + 10) = "Currency Code"
                                    objExcel.Cells(iRow, iCol + 11) = "Debit/Credit"
                                    objExcel.Cells(iRow, iCol + 12) = "Original Amount"
                                    objExcel.Cells(iRow, iCol + 13) = "Equivalent Amount"
                                    objExcel.Cells(iRow, iCol + 14) = "Exchange Rate"
                                    objExcel.Cells(iRow, iCol + 15) = "Document Type"
                                    objExcel.Cells(iRow, iCol + 16) = "Document Number"
                                    objExcel.Cells(iRow, iCol + 17) = "Document Date"
                                    objExcel.Cells(iRow, iCol + 18) = "Payment Terms"
                                    objExcel.Cells(iRow, iCol + 19) = "Document Due Date"
                                    objExcel.Cells(iRow, iCol + 20) = "Particular 1"
                                    objExcel.Cells(iRow, iCol + 21) = "Particular 2"
                                    objExcel.Cells(iRow, iCol + 22) = "Open Item Number"
                                    objExcel.Cells(iRow, iCol + 23) = "Applied Amount"
                                    objExcel.Cells(iRow, iCol + 24) = "Alternate Doc. No."
                                    objExcel.Cells(iRow, iCol + 25) = "Account Name(SC)"
                                    objExcel.Cells(iRow, iCol + 26) = "Revise"
                                    objExcel.Cells(iRow, iCol + 27) = "Week"

                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Interior.ColorIndex = 15
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Font.Bold = True
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(8).LineStyle = 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(9).LineStyle = 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(10).LineStyle = 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(11).LineStyle = 1

                                    iRow = iRow + 1
                                End If
                                'End If

                                If PrevDocNo <> "" Then
                                    iCount = iCount + 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                                End If

                                PrevDocNo = ds.Tables(j).Rows(i).Item("DocuNo").ToString
                            End If

                    'Set Content
                    objExcel.Cells(iRow, iCol) = ds.Tables(j).Rows(i).Item("Ledger").ToString
                    objExcel.Cells(iRow, iCol + 1) = ds.Tables(j).Rows(i).Item("Batch").ToString
                    objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(j).Rows(i).Item("Period").ToString
                    objExcel.Cells(iRow, iCol + 3) = "'" & ds.Tables(j).Rows(i).Item("Voucher").ToString & iCount

                    If common.NullVal(ds.Tables(j).Rows(i).Item("VoucherDte"), "") <> "" Then
                        objExcel.Cells(iRow, iCol + 4) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("VoucherDte")), "dd/MM/yyyy")
                    End If

                    objExcel.Cells(iRow, iCol + 5) = ds.Tables(j).Rows(i).Item("VouDesc").ToString
                    objExcel.Cells(iRow, iCol + 6) = "'" & ds.Tables(j).Rows(i).Item("AccCode").ToString
                    objExcel.Cells(iRow, iCol + 7) = "'" & ds.Tables(j).Rows(i).Item("Ana1").ToString
                    objExcel.Cells(iRow, iCol + 8) = "'" & ds.Tables(j).Rows(i).Item("Ana2").ToString
                    objExcel.Cells(iRow, iCol + 9) = "'" & ds.Tables(j).Rows(i).Item("Ana5").ToString
                    objExcel.Cells(iRow, iCol + 10) = ds.Tables(j).Rows(i).Item("Currency").ToString

                    If ds.Tables(j).Rows(i).Item("OrigAmt") >= 0 Then
                        objExcel.Cells(iRow, iCol + 11) = "D"
                    Else
                        objExcel.Cells(iRow, iCol + 11) = "C"
                    End If

                    objExcel.Cells(iRow, iCol + 12) = ds.Tables(j).Rows(i).Item("OrigAmt").ToString
                    objExcel.Cells(iRow, iCol + 13) = FormatNumber(ds.Tables(j).Rows(i).Item("OrigAmt") * ds.Tables(j).Rows(i).Item("ExRate"), 2)

                    totalAmt = totalAmt + FormatNumber(ds.Tables(j).Rows(i).Item("OrigAmt") * ds.Tables(j).Rows(i).Item("ExRate"), 2)

                    objExcel.Cells(iRow, iCol + 14) = ds.Tables(j).Rows(i).Item("ExRate").ToString
                    objExcel.Cells(iRow, iCol + 15) = ds.Tables(j).Rows(i).Item("DocType").ToString
                    objExcel.Cells(iRow, iCol + 16) = "'" & ds.Tables(j).Rows(i).Item("DocuNo").ToString

                    If common.NullVal(ds.Tables(j).Rows(i).Item("DocDte"), "") <> "" Then
                        objExcel.Cells(iRow, iCol + 17) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DocDte")), "dd/MM/yyyy")
                    End If

                    objExcel.Cells(iRow, iCol + 18) = ds.Tables(j).Rows(i).Item("PayTerm").ToString

                    If common.NullVal(ds.Tables(j).Rows(i).Item("DueDte"), "") <> "" Then
                        objExcel.Cells(iRow, iCol + 19) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DueDte")), "dd/MM/yyyy")
                    End If

                    objExcel.Cells(iRow, iCol + 20) = ds.Tables(j).Rows(i).Item("Part1").ToString
                    objExcel.Cells(iRow, iCol + 21) = "'" & ds.Tables(j).Rows(i).Item("Part2").ToString
                    objExcel.Cells(iRow, iCol + 22) = ds.Tables(j).Rows(i).Item("OriDocuNo").ToString
                    If CDbl(ds.Tables(j).Rows(i).Item("AppliedAmt").ToString) = 0 Or IsRevise = 0 Then
                        objExcel.Cells(iRow, iCol + 23) = ""
                    Else
                        objExcel.Cells(iRow, iCol + 23) = ds.Tables(j).Rows(i).Item("AppliedAmt").ToString
                    End If
                    objExcel.Cells(iRow, iCol + 24) = ds.Tables(j).Rows(i).Item("AltDocNo").ToString
                    objExcel.Cells(iRow, iCol + 25) = ds.Tables(j).Rows(i).Item("BillTo").ToString
                    objExcel.Cells(iRow, iCol + 26) = ds.Tables(j).Rows(i).Item("Revise").ToString
                    objExcel.Cells(iRow, iCol + 27) = ds.Tables(j).Rows(i).Item("BkhWeek").ToString

                    rowLn = ds.Tables(j).Rows(i).Item("BkvLn").ToString
                    ledger = ds.Tables(j).Rows(i).Item("Ledger").ToString

                    iRow = iRow + 1
                        Next
                    End If
                End If

                iRow = iRow + 2
            Next

            objExcel.Range(objExcel.Cells(2, iCol + 12), objExcel.Cells(iRow, iCol + 12)).NumberFormatLocal = "#,##0.00_ "
            objExcel.Range(objExcel.Cells(2, iCol + 13), objExcel.Cells(iRow, iCol + 13)).NumberFormatLocal = "#,##0.00_ "
            objExcel.Range(objExcel.Cells(2, iCol + 23), objExcel.Cells(iRow, iCol + 23)).NumberFormatLocal = "#,##0.00_ "

            objExcel.Columns("A:A").ColumnWidth = 4.25
            objExcel.Columns("B:B").ColumnWidth = 3
            objExcel.Columns("C:C").ColumnWidth = 7
            objExcel.Columns("D:D").ColumnWidth = 6
            objExcel.Columns("E:E").ColumnWidth = 11
            objExcel.Columns("F:F").ColumnWidth = 13.25
            objExcel.Columns("G:G").ColumnWidth = 6.5
            objExcel.Columns("H:H").ColumnWidth = 10
            objExcel.Columns("I:I").ColumnWidth = 5
            objExcel.Columns("J:J").ColumnWidth = 7.5
            objExcel.Columns("K:K").ColumnWidth = 5
            objExcel.Columns("L:L").ColumnWidth = 3
            objExcel.Columns("M:N").ColumnWidth = 15
            objExcel.Columns("O:O").ColumnWidth = 9
            objExcel.Columns("P:P").ColumnWidth = 2
            objExcel.Columns("Q:Q").ColumnWidth = 14
            objExcel.Columns("R:R").ColumnWidth = 9.5
            objExcel.Columns("S:S").ColumnWidth = 13
            objExcel.Columns("T:T").ColumnWidth = 10
            objExcel.Columns("U:U").ColumnWidth = 32
            objExcel.Columns("V:V").ColumnWidth = 12
            objExcel.Columns("W:Y").ColumnWidth = 14
            objExcel.Columns("Z:Z").ColumnWidth = 50

            ' ----------------------------------------------------------------------
            ' Save File
            ' ----------------------------------------------------------------------
            Dim exportPath As String = My.Settings.ExportPath
            Dim exportFile As String = ""

            If fileName <> "" Then
                exportFile = exportPath & fileName & ".xls"
            Else
                exportFile = exportPath & Uid & ".xls"
            End If

            fileName &= ".xls"

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

            objExcel.ActiveWorkbook.SaveAs("C:\" & uid & ".xls")
            objExcel.Quit()
            fileName = "Error," & ex.Message
        End Try

        ' Return File Path
        RptAccount = fileName

    End Function

End Class
