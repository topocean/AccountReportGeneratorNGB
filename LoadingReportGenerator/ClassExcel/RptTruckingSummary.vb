Public Class RptTruckingSummary

    Public Function RptTruckingSummary(ByVal uid As String, ByVal ds As DataSet) As String
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
        Dim total As Double

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
                fileName = common.NullVal(.Item("RptFile").ToString, uid)

                ' ----------------------------------------------------------------------
                ' Report Header (Company Name, Address, Tel, etc...)
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                objWS.Application.Cells(5, 1) = "Trucking Summary (Period From: " & common.NullVal(.Item("DteFrm"), "") & " to " & common.NullVal(.Item("DteTo").ToString, "") & ")"

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
            ' Trucking Summary (RMB)
            ' ----------------------------------------------------------------------

            If ds.Tables(1).Rows.Count > 0 Then
                hasData = True

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    If ds.Tables(1).Rows(i).Item("BkiCosSuppCd").ToString <> ClientRefId Then
                        ClientRefId = ds.Tables(1).Rows(i).Item("BkiCosSuppCd").ToString
                        
                        objWS.Application.Cells(iRow, iCol) = ds.Tables(1).Rows(i).Item("ClientName")
                        iRow += 1

                        objWS.Application.Cells(iRow, iCol) = "Date"
                        objWS.Application.Cells(iRow, iCol + 1) = "Invoice No."
                        objWS.Application.Cells(iRow, iCol + 2) = "Container Size"
                        objWS.Application.Cells(iRow, iCol + 3) = "Container No."
                        objWS.Application.Cells(iRow, iCol + 4) = "Lot No."
                        objWS.Application.Cells(iRow, iCol + 5) = "卡車費"
                        objWS.Application.Cells(iRow, iCol + 6) = "報關費"
                        objWS.Application.Cells(iRow, iCol + 7) = "轉關費"
                        objWS.Application.Cells(iRow, iCol + 8) = "壓夜費"
                        objWS.Application.Cells(iRow, iCol + 9) = "押架費"
                        objWS.Application.Cells(iRow, iCol + 10) = "異地提櫃費"
                        objWS.Application.Cells(iRow, iCol + 11) = "港口建設費"
                        objWS.Application.Cells(iRow, iCol + 12) = "查櫃費"
                        objWS.Application.Cells(iRow, iCol + 13) = "蒸燻費"
                        objWS.Application.Cells(iRow, iCol + 14) = "其他費用"
                        objWS.Application.Cells(iRow, iCol + 15) = "Total"
                        objWS.Application.Cells(iRow, iCol + 16) = "Currency"
                        objWS.Application.Cells(iRow, iCol + 17) = "Remarks"
                        objWS.Application.Cells(iRow, iCol + 18) = "Week"

                        ' **********************************************************************

                        ' ----------------------------------------------------------------------
                        ' Setting Properties of Detail Header
                        ' ----------------------------------------------------------------------

                        objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 19)).Font.Bold = True
                        objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 19)).Borders(9).LineStyle = 1
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 18)).HorizontalAlignment = -4108

                        iRow += 1
                        iSRow = iRow
                    End If

                    ' **********************************************************************


                    ' ----------------------------------------------------------------------
                    ' Export Report Data onto Excel Wooksheet
                    ' ----------------------------------------------------------------------

                    With ds.Tables(1).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BkiCosSuppInvDte"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhSONo"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("SizeName"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("CtnrNo"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("BkhLotNo"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("Cost1"), 0)
                        objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("Cost2"), 0)
                        objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("Cost3"), 0)
                        objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("Cost4"), 0)
                        objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("Cost5"), 0)
                        objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("Cost6"), 0)
                        objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("Cost7"), 0)
                        objWS.Application.Cells(iRow, iCol + 12) = common.NullVal(.Item("Cost8"), 0)
                        objWS.Application.Cells(iRow, iCol + 13) = common.NullVal(.Item("Cost9"), 0)
                        objWS.Application.Cells(iRow, iCol + 14) = common.NullVal(.Item("Cost10"), 0)
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("Total"), 0)
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("BkiCosCurr"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = common.NullVal(.Item("CtnrRemark"), "")
                        objWS.Application.Cells(iRow, iCol + 18) = common.NullVal(.Item("BkhWeek"), "")
                    End With

                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1

                    ' ----------------------------------------------------------------------
                    ' Calculate the total number of containers
                    ' ----------------------------------------------------------------------

                    If i + 1 < ds.Tables(1).Rows.Count Then
                        ' Show Total Amount of Current Supplier
                        If ds.Tables(1).Rows(i + 1).Item("BkiCosSuppCd").ToString <> ClientRefId Then
                            objWS.Application.Cells(iRow, iCol + 12) = "TOTAL: "

                            objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 15), objWS.Application.Cells(iRow - 1, iCol + 15)).Select()
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Activate()
                            objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Borders(9).LineStyle = -4119
                            objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 15)).NumberFormatLocal = "#,###,##0.00_ "
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol), objWS.Application.Cells(iRow, iCol + 15)).Font.Bold = True
                            total = total + objWS.Application.Cells(iRow, iCol + 15).value
                            iRow = iRow + 2
                        End If
                    Else
                        ' Show Total Amount of Last Supplier
                        objWS.Application.Cells(iRow, iCol + 12) = "TOTAL: "

                        objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 15), objWS.Application.Cells(iRow - 1, iCol + 15)).Select()
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Activate()
                        objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Borders(9).LineStyle = -4119
                        objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 15)).NumberFormatLocal = "#,###,##0.00_ "
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol), objWS.Application.Cells(iRow, iCol + 15)).Font.Bold = True
                        total = total + objWS.Application.Cells(iRow, iCol + 15).value
                        iRow = iRow + 2
                    End If
                Next
            End If

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Trucking Summary (HKD)
            ' ----------------------------------------------------------------------

            If ds.Tables(2).Rows.Count > 0 Then
                hasData = True

                For i = 0 To ds.Tables(2).Rows.Count - 1
                    If ds.Tables(2).Rows(i).Item("BkiCosSuppCd").ToString <> ClientRefId Then
                        ClientRefId = ds.Tables(2).Rows(i).Item("BkiCosSuppCd").ToString
                        
                        objWS.Application.Cells(iRow, iCol) = ds.Tables(2).Rows(i).Item("ClientName")
                        iRow += 1

                        objWS.Application.Cells(iRow, iCol) = "Date"
                        objWS.Application.Cells(iRow, iCol + 1) = "Invoice No."
                        objWS.Application.Cells(iRow, iCol + 2) = "Container Size"
                        objWS.Application.Cells(iRow, iCol + 3) = "Container No."
                        objWS.Application.Cells(iRow, iCol + 4) = "Lot No."
                        objWS.Application.Cells(iRow, iCol + 5) = "卡車費"
                        objWS.Application.Cells(iRow, iCol + 6) = "報關費"
                        objWS.Application.Cells(iRow, iCol + 7) = "轉關費"
                        objWS.Application.Cells(iRow, iCol + 8) = "壓夜費"
                        objWS.Application.Cells(iRow, iCol + 9) = "押架費"
                        objWS.Application.Cells(iRow, iCol + 10) = "異地提櫃費"
                        objWS.Application.Cells(iRow, iCol + 11) = "港口建設費"
                        objWS.Application.Cells(iRow, iCol + 12) = "查櫃費"
                        objWS.Application.Cells(iRow, iCol + 13) = "蒸燻費"
                        objWS.Application.Cells(iRow, iCol + 14) = "其他費用"
                        objWS.Application.Cells(iRow, iCol + 15) = "Total"
                        objWS.Application.Cells(iRow, iCol + 16) = "Currency"
                        objWS.Application.Cells(iRow, iCol + 17) = "Remarks"
                        objWS.Application.Cells(iRow, iCol + 18) = "Week"

                        ' **********************************************************************

                        ' ----------------------------------------------------------------------
                        ' Setting Properties of Detail Header
                        ' ----------------------------------------------------------------------

                        objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 19)).Font.Bold = True
                        objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 19)).Borders(9).LineStyle = 1
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 18)).HorizontalAlignment = -4108

                        iRow += 1
                        iSRow = iRow
                    End If

                    ' **********************************************************************


                    ' ----------------------------------------------------------------------
                    ' Export Report Data onto Excel Wooksheet
                    ' ----------------------------------------------------------------------

                    With ds.Tables(2).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BkiCosSuppInvDte"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhSONo"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("SizeName"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("CtnrNo"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("BkhLotNo"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("Cost1"), 0)
                        objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("Cost2"), 0)
                        objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("Cost3"), 0)
                        objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("Cost4"), 0)
                        objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("Cost5"), 0)
                        objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("Cost6"), 0)
                        objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("Cost7"), 0)
                        objWS.Application.Cells(iRow, iCol + 12) = common.NullVal(.Item("Cost8"), 0)
                        objWS.Application.Cells(iRow, iCol + 13) = common.NullVal(.Item("Cost9"), 0)
                        objWS.Application.Cells(iRow, iCol + 14) = common.NullVal(.Item("Cost10"), 0)
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("Total"), 0)
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("BkiCosCurr"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = common.NullVal(.Item("CtnrRemark"), "")
                        objWS.Application.Cells(iRow, iCol + 18) = common.NullVal(.Item("BkhWeek"), "")
                    End With

                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1


                    ' ----------------------------------------------------------------------
                    ' Calculate the total number of containers
                    ' ----------------------------------------------------------------------

                    If i + 1 < ds.Tables(1).Rows.Count Then
                        ' Show Total Amount of Current Supplier
                        If ds.Tables(1).Rows(i + 1).Item("BkiCosSuppCd").ToString <> ClientRefId Then
                            objWS.Application.Cells(iRow, iCol + 12) = "TOTAL: "

                            objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 15), objWS.Application.Cells(iRow - 1, iCol + 15)).Select()
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Activate()
                            objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Borders(9).LineStyle = -4119
                            objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 15)).NumberFormatLocal = "#,###,##0.00_ "
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol), objWS.Application.Cells(iRow, iCol + 15)).Font.Bold = True
                            total = total + objWS.Application.Cells(iRow, iCol + 15).value
                            iRow = iRow + 2
                        End If
                    Else
                        ' Show Total Amount of Last Supplier
                        objWS.Application.Cells(iRow, iCol + 12) = "TOTAL: "

                        objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 15), objWS.Application.Cells(iRow - 1, iCol + 15)).Select()
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Activate()
                        objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Borders(9).LineStyle = -4119
                        objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 15)).NumberFormatLocal = "#,###,##0.00_ "
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol), objWS.Application.Cells(iRow, iCol + 15)).Font.Bold = True
                        total = total + objWS.Application.Cells(iRow, iCol + 15).value
                        iRow = iRow + 2
                    End If
                Next

                ' **********************************************************************

            End If

            ' **********************************************************************


            ' ----------------------------------------------------------------------
            ' Setting Properties (Underline the total columns)
            ' ----------------------------------------------------------------------

            If hasData Then
                iRow += 2
                objWS.Application.Cells(iRow, iCol + 12).value = "GRAND TOTAL: "
                objWS.Application.Cells(iRow, iCol + 15).value = total
                objWS.Application.Range(objWS.Application.Cells(iSRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 15)).NumberFormatLocal = "#,###,##0.00_ "
                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 15), objWS.Application.Cells(iRow, iCol + 15)).Borders(9).LineStyle = -4119
                objWS.Application.Range(objWS.Application.Cells(iRow, iCol), objWS.Application.Cells(iRow, iCol + 15)).Font.Bold = True

                ' Page Setting
                objWS.Application.Columns("A:B").ColumnWidth = 12
                objWS.Application.Columns("C:E").ColumnWidth = 14
                objWS.Application.Columns("F:Q").ColumnWidth = 12
                objWS.Application.Columns("R:R").ColumnWidth = 50
                objWS.Application.Columns("S:S").ColumnWidth = 12
            End If

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
            total = Nothing
            ClientRefId = Nothing
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

        RptTruckingSummary = fileName
    End Function
End Class
