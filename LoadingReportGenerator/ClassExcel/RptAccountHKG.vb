Public Class RptAccountHKG

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
        Dim hasData As Boolean = False
        Dim rowLn As String = ""
        Dim PrevDocNo As String = ""

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

                        objExcel.Cells(iRow, iCol) = "Branch"
                        objExcel.Cells(iRow, iCol + 1) = "Period"
                        objExcel.Cells(iRow, iCol + 2) = "Voucher"
                        objExcel.Cells(iRow, iCol + 3) = "Voucher Date"
                        objExcel.Cells(iRow, iCol + 4) = "Description"
                        objExcel.Cells(iRow, iCol + 5) = "A/C Code"
                        objExcel.Cells(iRow, iCol + 6) = "Ana.1"
                        objExcel.Cells(iRow, iCol + 7) = "Currency"
                        objExcel.Cells(iRow, iCol + 8) = "Orig.Amount"
                        objExcel.Cells(iRow, iCol + 9) = "ExRate"
                        objExcel.Cells(iRow, iCol + 10) = "Docu.No."
                        objExcel.Cells(iRow, iCol + 11) = "Date Sailed"
                        objExcel.Cells(iRow, iCol + 12) = "Account Name(SC)"
                        objExcel.Cells(iRow, iCol + 13) = "HBL/MBL"
                        objExcel.Cells(iRow, iCol + 14) = "Vessel/Voy"
                        objExcel.Cells(iRow, iCol + 15) = "China Invoice"
                        objExcel.Cells(iRow, iCol + 16) = "Week"

                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Interior.ColorIndex = 15
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Borders(8).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Borders(10).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 16)).Borders(11).LineStyle = 1

                        iRow = iRow + 1
                        iCount = 0

                        MaxRow = ds.Tables(j).Rows.Count - 1

                        For i = 0 To MaxRow
                            If PrevDocNo <> ds.Tables(j).Rows(i).Item("DocuNo").ToString Then
                                If PrevDocNo <> "" Then
                                    iCount = iCount + 1
                                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, 17)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                                End If

                                PrevDocNo = ds.Tables(j).Rows(i).Item("DocuNo").ToString
                            End If

                            'Set Content
                            objExcel.Cells(iRow, iCol) = ds.Tables(j).Rows(i).Item("Branch").ToString
                            objExcel.Cells(iRow, iCol + 1) = "'" & ds.Tables(j).Rows(i).Item("Period").ToString
                            objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(j).Rows(i).Item("Voucher").ToString & iCount
                            If common.NullVal(ds.Tables(j).Rows(i).Item("VoucherDte"), "") <> "" Then
                                objExcel.Cells(iRow, iCol + 3) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("VoucherDte")), "dd/MM/yyyy")
                            End If
                            objExcel.Cells(iRow, iCol + 4) = ds.Tables(j).Rows(i).Item("VouDesc").ToString
                            objExcel.Cells(iRow, iCol + 5) = ds.Tables(j).Rows(i).Item("AccCode").ToString
                            objExcel.Cells(iRow, iCol + 6) = ds.Tables(j).Rows(i).Item("Ana1").ToString
                            objExcel.Cells(iRow, iCol + 7) = ds.Tables(j).Rows(i).Item("Currency").ToString
                            objExcel.Cells(iRow, iCol + 8) = ds.Tables(j).Rows(i).Item("OrigAmt").ToString
                            objExcel.Cells(iRow, iCol + 9) = ds.Tables(j).Rows(i).Item("ExRate").ToString
                            objExcel.Cells(iRow, iCol + 10) = ds.Tables(j).Rows(i).Item("DocuNo").ToString
                            If common.NullVal(ds.Tables(j).Rows(i).Item("DocDte"), "") <> "" Then
                                objExcel.Cells(iRow, iCol + 11) = "'" & Format(Convert.ToDateTime(ds.Tables(j).Rows(i).Item("DocDte")), "dd/MM/yyyy")
                            End If
                            objExcel.Cells(iRow, iCol + 12) = ds.Tables(j).Rows(i).Item("BillTo").ToString
                            objExcel.Cells(iRow, iCol + 13) = ds.Tables(j).Rows(i).Item("Part1").ToString
                            objExcel.Cells(iRow, iCol + 14) = ds.Tables(j).Rows(i).Item("Vessel").ToString
                            objExcel.Cells(iRow, iCol + 15) = ds.Tables(j).Rows(i).Item("Part2").ToString
                            objExcel.Cells(iRow, iCol + 16) = ds.Tables(j).Rows(i).Item("BkhWeek").ToString

                            rowLn = ds.Tables(j).Rows(i).Item("BkvLn").ToString

                            iRow = iRow + 1
                        Next
                    End If
                End If

                iRow = iRow + 2
            Next

            objExcel.Range(objExcel.Cells(2, iCol + 8), objExcel.Cells(iRow, iCol + 8)).NumberFormatLocal = "#,##0.00_ "

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
                exportFile = exportPath & Uid & ".xls"
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

            objExcel.ActiveWorkbook.SaveAs("C:\" & uid & ".xls")
            objExcel.Quit()
            fileName = "Error," & ex.Message
        End Try

        ' Return File Path
        RptAccount = fileName

    End Function

End Class
