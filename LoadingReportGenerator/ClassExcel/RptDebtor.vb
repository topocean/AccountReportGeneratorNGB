Public Class RptDebtor

    Public Function RptDebtor(ByVal uid As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, k, startRow As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim hasData As Boolean = False
        Dim tblIndex As Integer
        Dim title As String
        Dim Curr As String
        Dim RefId As Integer
        Dim totalCol(3) As Integer
        Dim CompanyName, CompanyAddr, CompanyTel, CompanyFax As String

        ' Start Excel Application
        objExcel = CreateObject("Excel.Application")
        objExcel.Visible = False

        Try
            ' Get a new workbook
            objWB = objExcel.Workbooks.Add
            objWS = objWB.ActiveSheet

            ' ----------------------------------------------------------------------
            ' Define the starting row and column number of the detail header
            ' ----------------------------------------------------------------------

            iRow = 1
            iCol = 1

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Retrieve File Name and Report Header
            ' ----------------------------------------------------------------------

            If ds.Tables(0).Rows.Count > 0 Then
                hasData = True

                fileName = ds.Tables(tblIndex).Rows(0).Item("RptFile").ToString
                CompanyName = ds.Tables(tblIndex).Rows(0).Item("BrhName").ToString
                CompanyAddr = ds.Tables(tblIndex).Rows(0).Item("BrhAddr").ToString
                CompanyTel = ds.Tables(tblIndex).Rows(0).Item("BrhTel").ToString
                CompanyFax = ds.Tables(tblIndex).Rows(0).Item("BrhFax").ToString
            Else
                hasData = False
            End If

            If hasData Then
                ' Default Values
                tblIndex = 1
                k = 1
                RefId = 0
                Curr = ""

                ' Export Data
                If ds.Tables(tblIndex).Rows.Count > 0 Then
                    hasData = True

                    For i = 0 To ds.Tables(tblIndex).Rows.Count - 1
                        ' Set Worksheet Properties
                        objWS.Application.Cells.Font.Name = "Verdana"
                        objWS.Application.Cells.Font.Size = 9
                        objWS.Application.Cells.VerticalAlignment = -4160

                        title = ds.Tables(tblIndex).Rows(i).Item("title").ToString

                        ' Add New WorkSheet if different Client or same Client but different currency
                        If CInt(ds.Tables(tblIndex).Rows(i).Item("AgtCd").ToString) <> RefId Or ds.Tables(tblIndex).Rows(i).Item("Currency").ToString <> Curr Then
                            RefId = CInt(ds.Tables(tblIndex).Rows(i).Item("AgtCd").ToString)
                            Curr = ds.Tables(tblIndex).Rows(i).Item("Currency").ToString

                            With objWS.PageSetup
                                .TopMargin = objExcel.InchesToPoints(0.5)
                                .BottomMargin = objExcel.InchesToPoints(0.5)
                                .Orientation = 1
                                .Zoom = False
                                .FitToPagesWide = 1
                                .FitToPagesTall = 1
                            End With

                            iRow = 1
                            iCol = 1

                            'Company Name, Address & Tel.
                            objExcel.Cells(iRow, iCol).value = CompanyName
                            objExcel.Cells(iRow + 1, iCol).value = CompanyAddr
                            objExcel.Cells(iRow + 2, iCol).value = "TEL: " & CompanyTel & "  FAX: " & CompanyFax
                            objExcel.Cells(iRow + 4, iCol).value = title & " Statement" 'Shipper Statment/Agent Statment/Subcontactor Statment
                            objExcel.Cells(iRow + 4, iCol + 5).value = "Print Date:"
                            objExcel.Cells(iRow + 4, iCol + 6).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("PrintDate")), "dd-MM-yy")

                            iRow = iRow + 6
                            objExcel.Cells(iRow, iCol).value = title & ":"  'SHIPPER/AGENT/SUBCONTACTOR
                            objExcel.Cells(iRow, iCol + 1).value = ds.Tables(tblIndex).Rows(i).Item("AgtName").ToString
                            objExcel.Cells(iRow, iCol + 5).value = "Acct Code:"
                            objExcel.Cells(iRow, iCol + 6).value = ds.Tables(tblIndex).Rows(i).Item("AgtAccCd").ToString
                            objExcel.Cells(iRow + 1, iCol).value = "Address:"
                            objExcel.Cells(iRow + 1, iCol + 1).value = ds.Tables(tblIndex).Rows(i).Item("AgtAddress").ToString

                            If title <> "Agent" Then       'DONT SHOW IN AGENT
                                objExcel.Cells(iRow + 1, iCol + 5).value = "Currency:"
                                objExcel.Cells(iRow + 1, iCol + 6).value = ds.Tables(tblIndex).Rows(i).Item("Currency").ToString
                            End If

                            objExcel.Cells(iRow + 2, iCol).value = "Attn:"
                            objExcel.Cells(iRow + 2, iCol + 1).value = ds.Tables(tblIndex).Rows(i).Item("AgtAttn").ToString
                            objExcel.Cells(iRow + 3, iCol).value = "Telephone No:"
                            objExcel.Cells(iRow + 3, iCol + 1).value = "'" & ds.Tables(tblIndex).Rows(i).Item("AgtTel").ToString
                            objExcel.Cells(iRow + 3, iCol + 2).value = "Fax No:"
                            objExcel.Cells(iRow + 3, iCol + 3).value = "'" & ds.Tables(tblIndex).Rows(i).Item("AgtFax").ToString
                            objExcel.Cells(iRow + 3, iCol + 5).value = "Period:"
                            objExcel.Cells(iRow + 3, iCol + 6).value = "From " & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("DateFrm")), "dd-MM-yy") & " To " & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("DateTo")), "dd-MM-yy")

                            'Setting - bold
                            objExcel.Range("A1:H3").Font.Bold = True
                            objExcel.Range("A5:A5").Font.Bold = True
                            objExcel.Range("F10:F10").Font.Bold = True
                            objExcel.Range("A1:J3").HorizontalAlignment = -4108
                            objExcel.Range("A1:J1").Merge()
                            objExcel.Range("A2:J2").Merge()
                            objExcel.Range("A3:J3").Merge()
                            objExcel.Range("B8:D8").Merge()
                            objExcel.Range("B8:D8").WrapText = True
                            objExcel.Range("A8:D8").VerticalAlignment = -4160

                            iRow = iRow + 6
                            If title = "Agent" Then
                                objExcel.Cells(iRow, iCol).value = "Invoice Date"
                                objExcel.Cells(iRow, iCol + 1).value = "OB Date"
                                objExcel.Cells(iRow, iCol + 2).value = "Invoice No"
                                objExcel.Cells(iRow, iCol + 3).value = "Customer"
                                objExcel.Cells(iRow, iCol + 4).value = "Total # of BL"
                                objExcel.Cells(iRow, iCol + 5).value = "BL"
                                objExcel.Cells(iRow, iCol + 6).value = "US$ Amount"
                                objExcel.Cells(iRow, iCol + 7).value = "US$ Outstanding"
                                objExcel.Cells(iRow, iCol + 8).value = "US$ Balance"
                                objExcel.Cells(iRow, iCol + 9).value = "Due Date"
                                objExcel.Cells(iRow, iCol + 10).value = "Container#"
                                objExcel.Cells(iRow, iCol + 11).value = "PO#"
                                totalCol(0) = 4
                                totalCol(1) = 8
                                totalCol(3) = 12
                            ElseIf title = "Shipper" Then
                                objExcel.Cells(iRow, iCol).value = "Invoice Date"
                                objExcel.Cells(iRow, iCol + 1).value = "OB Date"
                                objExcel.Cells(iRow, iCol + 2).value = "Invoice No"
                                objExcel.Cells(iRow, iCol + 3).value = "Customer"
                                objExcel.Cells(iRow, iCol + 4).value = "BL"
                                objExcel.Cells(iRow, iCol + 5).value = "Amount"
                                objExcel.Cells(iRow, iCol + 6).value = "Outstanding"
                                objExcel.Cells(iRow, iCol + 7).value = "Balance"
                                objExcel.Cells(iRow, iCol + 8).value = "Due Date"
                                objExcel.Cells(iRow, iCol + 9).value = "Container#"
                                objExcel.Cells(iRow, iCol + 10).value = "PO#"
                                totalCol(0) = 4
                                totalCol(1) = 7
                                totalCol(3) = 11
                            Else
                                objExcel.Cells(iRow, iCol).value = "Invoice Date"
                                objExcel.Cells(iRow, iCol + 1).value = "Invoice No"
                                objExcel.Cells(iRow, iCol + 2).value = "Customer"
                                objExcel.Cells(iRow, iCol + 3).value = "MBL"
                                objExcel.Cells(iRow, iCol + 4).value = "CNTR"
                                objExcel.Cells(iRow, iCol + 5).value = "Amount"
                                objExcel.Cells(iRow, iCol + 6).value = "Outstanding"
                                objExcel.Cells(iRow, iCol + 7).value = "Balance"
                                objExcel.Cells(iRow, iCol + 8).value = "Due Date"
                                objExcel.Cells(iRow, iCol + 9).value = "PO#"
                                totalCol(0) = 3
                                totalCol(1) = 7
                                totalCol(3) = 10
                            End If

                            iRow = iRow + 1
                            startRow = iRow
                        End If

                        If title = "Agent" Then
                            If IsDBNull(ds.Tables(tblIndex).Rows(i).Item("InvoiceDate")) = False Then
                                objExcel.Cells(iRow, iCol).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("InvoiceDate")), "dd/MM/yyyy")
                            End If

                            If IsDBNull(ds.Tables(tblIndex).Rows(i).Item("OBDate")) = False Then
                                objExcel.Cells(iRow, iCol + 1).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("OBDate")), "dd/MM/yyyy")
                            End If

                            objExcel.Cells(iRow, iCol + 2).value = ds.Tables(tblIndex).Rows(i).Item("InvoiceNo").ToString
                            objExcel.Cells(iRow, iCol + 3).value = ds.Tables(tblIndex).Rows(i).Item("ConName").ToString
                            objExcel.Cells(iRow, iCol + 4).value = ds.Tables(tblIndex).Rows(i).Item("TotalBL").ToString

                            objExcel.Cells(iRow, iCol + 5).value = ds.Tables(tblIndex).Rows(i).Item("BLNo").ToString
                            objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(tblIndex).Rows(i).Item("amount").ToString, 2)
                            objExcel.Cells(iRow, iCol + 7).value = FormatNumber(ds.Tables(tblIndex).Rows(i).Item("outstanding"), 2)

                            objExcel.Range(objExcel.Cells(iRow, iCol + 8), objExcel.Cells(iRow, iCol + 8)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C[-1]:RC[-1])"
                            objExcel.Cells(iRow, iCol + 9).value = ds.Tables(tblIndex).Rows(i).Item("DueDate")

                            objExcel.Cells(iRow, iCol + 10).value = ds.Tables(tblIndex).Rows(i).Item("CtnrNo").ToString
                            objExcel.Cells(iRow, iCol + 11).value = ds.Tables(tblIndex).Rows(i).Item("PONo").ToString
                        ElseIf title = "Shipper" Then
                            If IsDBNull(ds.Tables(tblIndex).Rows(i).Item("InvoiceDate")) = False Then
                                objExcel.Cells(iRow, iCol).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("InvoiceDate")), "dd/MM/yyyy")
                            End If

                            If IsDBNull(ds.Tables(tblIndex).Rows(i).Item("OBDate")) = False Then
                                objExcel.Cells(iRow, iCol + 1).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("OBDate")), "dd/MM/yyyy")
                            End If
                            objExcel.Cells(iRow, iCol + 2).value = ds.Tables(tblIndex).Rows(i).Item("InvoiceNo").ToString
                            objExcel.Cells(iRow, iCol + 3).value = ds.Tables(tblIndex).Rows(i).Item("ConName").ToString
                            objExcel.Cells(iRow, iCol + 4).value = ds.Tables(tblIndex).Rows(i).Item("BLNo").ToString
                            objExcel.Cells(iRow, iCol + 5).value = FormatNumber(ds.Tables(tblIndex).Rows(i).Item("amount").ToString, 2)
                            objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(tblIndex).Rows(i).Item("outstanding").ToString, 2)

                            objExcel.Range(objExcel.Cells(iRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C[-1]:RC[-1])"
                            objExcel.Cells(iRow, iCol + 8).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("DueDate")), "dd/MM/yyyy")

                            objExcel.Cells(iRow, iCol + 9).value = ds.Tables(tblIndex).Rows(i).Item("CtnrNo").ToString
                            objExcel.Cells(iRow, iCol + 10).value = ds.Tables(tblIndex).Rows(i).Item("PONo").ToString

                        Else
                            If IsDBNull(ds.Tables(tblIndex).Rows(i).Item("InvoiceDate")) = False Then
                                objExcel.Cells(iRow, iCol).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("InvoiceDate")), "dd/MM/yyyy")
                            End If
                            objExcel.Cells(iRow, iCol + 1).value = "'" & ds.Tables(tblIndex).Rows(i).Item("InvoiceNo").ToString
                            objExcel.Cells(iRow, iCol + 2).value = ds.Tables(tblIndex).Rows(i).Item("ConName").ToString

                            objExcel.Cells(iRow, iCol + 3).value = ds.Tables(tblIndex).Rows(i).Item("BLNo").ToString

                            objExcel.Cells(iRow, iCol + 4).value = ds.Tables(tblIndex).Rows(i).Item("CtnrNo").ToString
                            objExcel.Cells(iRow, iCol + 5).value = FormatNumber(ds.Tables(tblIndex).Rows(i).Item("amount").ToString, 2)
                            objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(tblIndex).Rows(i).Item("outstanding").ToString, 2)

                            objExcel.Range(objExcel.Cells(iRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C[-1]:RC[-1])"
                            objExcel.Cells(iRow, iCol + 8).value = "'" & Format(Convert.ToDateTime(ds.Tables(tblIndex).Rows(i).Item("DueDate")), "dd/MM/yyyy")

                            objExcel.Cells(iRow, iCol + 9).value = ds.Tables(tblIndex).Rows(i).Item("PONo").ToString
                        End If

                        iRow = iRow + 1

                        If i + 1 < ds.Tables(tblIndex).Rows.Count Then
                            If ds.Tables(tblIndex).Rows(i + 1).Item("AgtCd").ToString <> RefId Or ds.Tables(tblIndex).Rows(i + 1).Item("Currency").ToString <> Curr Then
                                'Add Footer before adding new worksheet
                                iRow = iRow + 2
                                objExcel.Cells(iRow, totalCol(1)).value = "GRAND TOTAL"
                                objExcel.Range(objExcel.Cells(iRow, totalCol(1)), objExcel.Cells(iRow, totalCol(1))).HorizontalAlignment = -4108
                                objExcel.Range(objExcel.Cells(iRow, totalCol(1) + 1), objExcel.Cells(iRow, totalCol(1) + 1)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C[-1]:R[-1]C[-1])"

                                'Setting - bold & line (Header & Detail)
                                objExcel.Range(objExcel.Cells(startRow - 1, iCol), objExcel.Cells(startRow - 1, totalCol(3))).Font.Bold = True
                                objExcel.Range(objExcel.Cells(startRow - 1, iCol), objExcel.Cells(startRow - 1, totalCol(3))).Borders(9).LineStyle = 1

                                If title = "Agent" Then
                                    objExcel.Range(objExcel.Cells(startRow, 5), objExcel.Cells(iRow - 2, 5)).HorizontalAlignment = -4108
                                End If

                                objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 8)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, totalCol(1) + 1), objExcel.Cells(iRow, totalCol(1) + 1)).Borders(9).LineStyle = -4119
                                objExcel.Range(objExcel.Cells(startRow, totalCol(1) - 1), objExcel.Cells(iRow, totalCol(1) + 1)).NumberFormatLocal = "#,##0.00_ "

                                If iRow >= 68 Then
                                    j = 3
                                    iRow = iRow + 2
                                Else
                                    j = 68 - iRow + 1
                                    iRow = 68
                                End If

                                objExcel.Range("B" & iRow & ":C" & iRow).Merge()
                                objExcel.Range("B" & iRow + 1 & ":C" & iRow + 1).Merge()
                                objExcel.Range("E" & iRow + 1 & ":F" & iRow + 1).Merge()
                                objExcel.Cells(iRow, iCol).value = "CURRENT"
                                objExcel.Cells(iRow, iCol + 1).value = "1-30 DAYS"
                                objExcel.Cells(iRow, iCol + 3).value = "OVER 30 DAYS"
                                objExcel.Cells(iRow, iCol + 4).value = "OVER 60 DAYS"
                                objExcel.Cells(iRow, iCol + 6).value = "OVER 90 DAYS"

                                iRow = iRow + 1

                                objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol)).FormulaR1C1 = "=+R[-" & j & "]C[" & totalCol(1) & "]"
                                'Setting-line
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(7).LineStyle = 1
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(8).LineStyle = 1
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(9).LineStyle = 1
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(10).LineStyle = 1
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(11).LineStyle = 1

                                If title <> "Subcontactor" Then
                                    iRow = iRow + 3
                                    objExcel.Cells(iRow, iCol).value = "Note : 1. E. & O.E."
                                    objExcel.Cells(iRow + 1, iCol).value = "          2. Transactions made after the statement date will be reflected in next month's statement."
                                    objExcel.Cells(iRow + 2, iCol).value = "          3. Please examine this statement and advise us of any discrepancies within 14 days of receipt."
                                    objExcel.Cells(iRow + 3, iCol).value = "          4. If we do not hear from you, we will take this statement as correct & binding."

                                    iRow = iRow + 5
                                    objExcel.Cells(iRow, iCol).value = "Our banker's details as follow:-"
                                    objExcel.Cells(iRow + 1, iCol).value = "Name of bank: The Hong Kong and Shanghai Banking Corporation Ltd. (Bonham Strand Branch)"
                                    objExcel.Cells(iRow + 2, iCol).value = "A/C Name : Topocean Consolidation Service Ltd."
                                    objExcel.Cells(iRow + 3, iCol).value = "Currency  : US$ "
                                    objExcel.Cells(iRow + 3, iCol + 2).value = "A/C No : 004-459-206751-274"

                                    If title = "Shipper" Then
                                        objExcel.Cells(iRow + 4, iCol).value = "Currency  : HK$ "
                                        objExcel.Cells(iRow + 4, iCol + 2).value = "A/C No : 004-459-206751-001"
                                    Else
                                        objExcel.Cells(iRow + 4, iCol).value = "Bank address : 35-45B Bonham Strand East, Hong Kong."
                                    End If
                                End If

                                objExcel.Columns("A:A").ColumnWidth = 14
                                objExcel.Columns("B:C").ColumnWidth = 13
                                objExcel.Columns("D:D").ColumnWidth = 40
                                objExcel.Columns("E:E").ColumnWidth = 14.5
                                objExcel.Columns("F:I").ColumnWidth = 13.5
                                objExcel.Columns("J:J").ColumnWidth = 11.5
                                objExcel.Columns("K:L").ColumnWidth = 15

                                objExcel.Application.Sheets.Add()
                                k += 1
                            End If
                        Else
                            'Add Footer before adding new worksheet
                            iRow = iRow + 2
                            objExcel.Cells(iRow, totalCol(1)).value = "GRAND TOTAL"
                            objExcel.Range(objExcel.Cells(iRow, totalCol(1)), objExcel.Cells(iRow, totalCol(1))).HorizontalAlignment = -4108
                            objExcel.Range(objExcel.Cells(iRow, totalCol(1) + 1), objExcel.Cells(iRow, totalCol(1) + 1)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C[-1]:R[-1]C[-1])"

                            'Setting - bold & line (Header & Detail)
                            objExcel.Range(objExcel.Cells(startRow - 1, iCol), objExcel.Cells(startRow - 1, totalCol(3))).Font.Bold = True
                            objExcel.Range(objExcel.Cells(startRow - 1, iCol), objExcel.Cells(startRow - 1, totalCol(3))).Borders(9).LineStyle = 1

                            If title = "Agent" Then
                                objExcel.Range(objExcel.Cells(startRow, 5), objExcel.Cells(iRow - 2, 5)).HorizontalAlignment = -4108
                            End If

                            objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 8)).Font.Bold = True
                            objExcel.Range(objExcel.Cells(iRow, totalCol(1) + 1), objExcel.Cells(iRow, totalCol(1) + 1)).Borders(9).LineStyle = -4119
                            objExcel.Range(objExcel.Cells(startRow, totalCol(1) - 1), objExcel.Cells(iRow, totalCol(1) + 1)).NumberFormatLocal = "#,##0.00_ "

                            If iRow >= 68 Then
                                j = 3
                                iRow = iRow + 2
                            Else
                                j = 68 - iRow + 1
                                iRow = 68
                            End If

                            objExcel.Range("B" & iRow & ":C" & iRow).Merge()
                            objExcel.Range("B" & iRow + 1 & ":C" & iRow + 1).Merge()
                            objExcel.Range("E" & iRow + 1 & ":F" & iRow + 1).Merge()
                            objExcel.Cells(iRow, iCol).value = "CURRENT"
                            objExcel.Cells(iRow, iCol + 1).value = "1-30 DAYS"
                            objExcel.Cells(iRow, iCol + 3).value = "OVER 30 DAYS"
                            objExcel.Cells(iRow, iCol + 4).value = "OVER 60 DAYS"
                            objExcel.Cells(iRow, iCol + 6).value = "OVER 90 DAYS"

                            iRow = iRow + 1

                            objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol)).FormulaR1C1 = "=+R[-" & j & "]C[" & totalCol(1) & "]"
                            'Setting-line
                            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(7).LineStyle = 1
                            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(8).LineStyle = 1
                            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(9).LineStyle = 1
                            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(10).LineStyle = 1
                            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(11).LineStyle = 1

                            If title <> "Subcontactor" Then
                                iRow = iRow + 3
                                objExcel.Cells(iRow, iCol).value = "Note : 1. E. & O.E."
                                objExcel.Cells(iRow + 1, iCol).value = "          2. Transactions made after the statement date will be reflected in next month's statement."
                                objExcel.Cells(iRow + 2, iCol).value = "          3. Please examine this statement and advise us of any discrepancies within 14 days of receipt."
                                objExcel.Cells(iRow + 3, iCol).value = "          4. If we do not hear from you, we will take this statement as correct & binding."

                                iRow = iRow + 5
                                objExcel.Cells(iRow, iCol).value = "Our banker's details as follow:-"
                                objExcel.Cells(iRow + 1, iCol).value = "Name of bank: The Hong Kong and Shanghai Banking Corporation Ltd. (Bonham Strand Branch)"
                                objExcel.Cells(iRow + 2, iCol).value = "A/C Name : Topocean Consolidation Service Ltd."
                                objExcel.Cells(iRow + 3, iCol).value = "Currency  : US$ "
                                objExcel.Cells(iRow + 3, iCol + 2).value = "A/C No : 004-459-206751-274"

                                If title = "Shipper" Then
                                    objExcel.Cells(iRow + 4, iCol).value = "Currency  : HK$ "
                                    objExcel.Cells(iRow + 4, iCol + 2).value = "A/C No : 004-459-206751-001"
                                Else
                                    objExcel.Cells(iRow + 4, iCol).value = "Bank address : 35-45B Bonham Strand East, Hong Kong."
                                End If
                            End If

                            objExcel.Columns("A:A").ColumnWidth = 14
                            objExcel.Columns("B:C").ColumnWidth = 13
                            objExcel.Columns("D:D").ColumnWidth = 40
                            objExcel.Columns("E:E").ColumnWidth = 14.5
                            objExcel.Columns("F:I").ColumnWidth = 13.5
                            objExcel.Columns("J:J").ColumnWidth = 11.5
                            objExcel.Columns("K:L").ColumnWidth = 15
                        End If
                    Next
                End If
            End If

            If Not hasData Then
                fileName = ""
            End If

            ' ----------------------------------------------------------------------
            ' Save File
            ' ----------------------------------------------------------------------

            Dim exportPath As String = My.Settings.ExportPath
            Dim exportFile As String = ""

            If fileName <> "" Then
                exportFile = exportPath & fileName & ".xls"
            Else
                exportFile = exportPath & uid & ".xls"
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

            ' Release Memory
            GC.Collect()
            GC.WaitForPendingFinalizers()

        Catch ex As Exception
            objExcel.ActiveWorkbook.SaveAs("C:\" & uid & ".xls")
            objExcel.Quit()
            fileName = "Error," & ex.Message
        End Try

        RptDebtor = fileName
    End Function
End Class
