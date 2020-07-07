Public Class RptDailyCashReceipt

    Function RptDailyCashReceipt(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
                System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String
        Dim POType As String = ""
        Dim common As New common
        Dim hasData As Boolean = False
        Dim ClientRefId As Integer = 0
        Dim isStart As Boolean = True
        Dim total As Double
        Dim hasOthers As Boolean = True

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
            ' Get File Name, Sub-Branch
            ' ----------------------------------------------------------------------

            With ds.Tables(0).Rows(0)
                fileName = common.NullVal(.Item("RptFile").ToString, UID)

                ' ----------------------------------------------------------------------
                ' Report Header (Company Name, Address, Tel, etc...)
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                objWS.Application.Cells(5, 1) = "Daily Cash Receipt (Period From: " & common.NullVal(.Item("DteFrm"), "") & " to " & common.NullVal(.Item("DteTo").ToString, "") & ")"

                ' ----------------------------------------------------------------------
                ' Setting Properties (Bold Header Details and Merge Cells)
                ' ----------------------------------------------------------------------

                objWS.Application.Range("A1:M6").Font.Bold = True
                objWS.Application.Range("A1:M6").HorizontalAlignment = -4108
                objWS.Application.Range("A1:M1").Merge()
                objWS.Application.Range("A2:M2").Merge()
                objWS.Application.Range("A3:M3").Merge()
                objWS.Application.Range("A5:M5").Merge()
                objWS.Application.Range("A6:M6").Merge()
            End With

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Set Column Header Line
            ' ----------------------------------------------------------------------

            iRow = 8

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Cash Payments
            ' ----------------------------------------------------------------------

            If ds.Tables(1).Rows.Count > 0 Then
                hasData = True

                objWS.Application.Cells(iRow, iCol) = "CASH"
                iRow += 1

                ' ----------------------------------------------------------------------
                ' Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol) = "BRANCH"
                objWS.Application.Cells(iRow, iCol + 1) = "WEEK"
                objWS.Application.Cells(iRow, iCol + 2) = "INVOICE NO."
                objWS.Application.Cells(iRow, iCol + 3) = "PAYOR"
                objWS.Application.Cells(iRow, iCol + 4) = "HOUSE BL"
                objWS.Application.Cells(iRow, iCol + 5) = "CURRENCY"
                objWS.Application.Cells(iRow, iCol + 6) = "USD"
                objWS.Application.Cells(iRow, iCol + 7) = "HKD"
                objWS.Application.Cells(iRow, iCol + 8) = "RMB"
                objWS.Application.Cells(iRow, iCol + 9) = "OUTSTANDING"
                objWS.Application.Cells(iRow, iCol + 10) = "REMARK"
                objWS.Application.Cells(iRow, iCol + 11) = "A/C GROUP"
                objWS.Application.Cells(iRow, iCol + 12) = "BANK A/C"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 13)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 13)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 13)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 13)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 13)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 13)).Borders(11).LineStyle = 1

                iRow += 1
                iSRow = iRow

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    With ds.Tables(1).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BrhName"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhWeek"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("IvhInvNo"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("ClientName"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("BkhBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("CurCd"), "")

                        ' USD
                        If common.NullVal(.Item("CurCd"), "") = "USD" Then
                            objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("LedgerAmt"), 0)
                        End If

                        ' HKD
                        If common.NullVal(.Item("CurCd"), "") = "HKD" Then
                            objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("LedgerAmt"), 0)
                        End If

                        ' RMB
                        If common.NullVal(.Item("CurCd"), "") = "RMB" Then
                            objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("LedgerAmt"), 0)
                        End If

                        objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("LedgerOutAmtInLocal"), 0)
                        objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("Remark"), "")
                        objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("AccGrp"), "")
                        objWS.Application.Cells(iRow, iCol + 12) = "'" & common.NullVal(.Item("BankCode"), "")
                    End With

                    ' **********************************************************************


                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1
                Next

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Calculate the total number of containers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 5) = "TOTAL: "
                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 5)).Font.Bold = True

                For i = 7 To 10
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Borders(9).LineStyle = -4119
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow, i)).NumberFormatLocal = "#,###,##0.00_ "
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Font.Bold = True
                Next

                iRow = iRow + 5

                ' **********************************************************************
            End If

            ' ----------------------------------------------------------------------
            ' End of Cash Payments
            ' ----------------------------------------------------------------------

            ' **********************************************************************


            ' ----------------------------------------------------------------------
            ' Cheque Payments
            ' ----------------------------------------------------------------------

            If ds.Tables(2).Rows.Count > 0 Then
                hasData = True

                objWS.Application.Cells(iRow, iCol) = "CHEQUE"

                iRow += 1
                objWS.Application.Range("K" & iRow & ":M" & iRow).Merge()
                objWS.Application.Cells(iRow, iCol + 10) = "ACTUAL RECEIVING"
                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 10), objWS.Application.Cells(iRow, iCol + 10)).Font.Bold = True

                iRow += 1

                ' ----------------------------------------------------------------------
                ' Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol) = "BRANCH"
                objWS.Application.Cells(iRow, iCol + 1) = "WEEK"
                objWS.Application.Cells(iRow, iCol + 2) = "INVOICE NO."
                objWS.Application.Cells(iRow, iCol + 3) = "PAYOR"
                objWS.Application.Cells(iRow, iCol + 4) = "HOUSE BL"
                objWS.Application.Cells(iRow, iCol + 5) = "CURRENCY"
                objWS.Application.Cells(iRow, iCol + 6) = "USD"
                objWS.Application.Cells(iRow, iCol + 7) = "HKD"
                objWS.Application.Cells(iRow, iCol + 8) = "RMB"
                objWS.Application.Cells(iRow, iCol + 9) = "OUTSTANDING"
                objWS.Application.Cells(iRow, iCol + 10) = "USD"
                objWS.Application.Cells(iRow, iCol + 11) = "HKD"
                objWS.Application.Cells(iRow, iCol + 12) = "RMB"
                objWS.Application.Cells(iRow, iCol + 13) = "CHEQUE NO."
                objWS.Application.Cells(iRow, iCol + 14) = "BANK"
                objWS.Application.Cells(iRow, iCol + 15) = "REMARK"
                objWS.Application.Cells(iRow, iCol + 16) = "A/C GROUP"
                objWS.Application.Cells(iRow, iCol + 17) = "BANK ACC"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 18)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 18)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 18)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 18)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 18)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 18)).Borders(11).LineStyle = 1

                iRow += 1
                iSRow = iRow

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(2).Rows.Count - 1
                    With ds.Tables(2).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BrhName"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhWeek"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("IvhInvNo"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("ClientName"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("BkhBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("CurCd"), "")

                        ' USD
                        If common.NullVal(.Item("CurCd"), "") = "USD" Then
                            objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("LedgerAmt"), 0)
                        End If

                        ' HKD
                        If common.NullVal(.Item("CurCd"), "") = "HKD" Then
                            objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("LedgerAmt"), 0)
                        End If

                        ' RMB
                        If common.NullVal(.Item("CurCd"), "") = "RMB" Then
                            objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("LedgerAmt"), 0)
                        End If

                        objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("LedgerOutAmtInLocal"), 0)

                        ' USD
                        If common.NullVal(.Item("IsChequeShow"), 0) = 0 Then
                            If common.NullVal(.Item("ChqCurCd"), "") = "USD" Then
                                objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("ChqAmt"), 0)
                            End If

                            ' HKD
                            If common.NullVal(.Item("ChqCurCd"), "") = "HKD" Then
                                objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("ChqAmt"), 0)
                            End If

                            ' RMB
                            If common.NullVal(.Item("ChqCurCd"), "") = "RMB" Then
                                objWS.Application.Cells(iRow, iCol + 12) = common.NullVal(.Item("ChqAmt"), 0)
                            End If

                            objWS.Application.Cells(iRow, iCol + 13) = "'" & common.NullVal(.Item("ChqNo"), "")
                        End If

                        objWS.Application.Cells(iRow, iCol + 14) = common.NullVal(.Item("BankName"), "")
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("Remark"), "")
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("AccGrp"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = "'" & common.NullVal(.Item("BankCode"), "")
                    End With

                    ' **********************************************************************


                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1
                Next

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Calculate the total number of containers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 5) = "TOTAL: "
                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 5)).Font.Bold = True

                For i = 7 To 13
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Borders(9).LineStyle = -4119
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow, i)).NumberFormatLocal = "#,###,##0.00_ "
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Font.Bold = True
                Next

                iRow = iRow + 2

                ' **********************************************************************
            End If

            ' ----------------------------------------------------------------------
            ' End of Cheque Payments
            ' ----------------------------------------------------------------------

            ' **********************************************************************


            ' ----------------------------------------------------------------------
            ' Setting Properties (Column Width)
            ' ----------------------------------------------------------------------

            objWS.Application.Columns("A:A").ColumnWidth = 9
            objWS.Application.Columns("B:B").ColumnWidth = 7
            objWS.Application.Columns("C:C").ColumnWidth = 14
            objWS.Application.Columns("D:D").ColumnWidth = 25
            objWS.Application.Columns("E:E").ColumnWidth = 16
            objWS.Application.Columns("F:F").ColumnWidth = 11
            objWS.Application.Columns("G:N").ColumnWidth = 15
            objWS.Application.Columns("L:R").ColumnWidth = 20


            ' **********************************************************************

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
                exportFile = exportPath & UID & ".xls"
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
            total = Nothing
            ClientRefId = Nothing
            iSRow = Nothing
            iCol = Nothing

            ' Release Memory
            GC.Collect()
            GC.WaitForPendingFinalizers()

        Catch ex As Exception
            objExcel.ActiveWorkbook.SaveAs("C:\" & UID & ".xls")
            objExcel.Quit()
            fileName = "Error," & ex.Message
        End Try

        RptDailyCashReceipt = fileName
    End Function
End Class
