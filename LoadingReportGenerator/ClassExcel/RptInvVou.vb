Public Class RptInvVou

    Function RptInvVou(ByVal UID As String, ByVal sUid As String, ByVal ds As DataSet) As String

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim SubNo As String = ""
        Dim InvIndex As Integer = 0

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

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    With ds.Tables(TblIndex).Rows(0)
                        fileName = .Item("fname").ToString
                    End With

                End If

                TblIndex += 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    With ds.Tables(TblIndex).Rows(0)
                        objExcel.Cells(1, 1) = .Item("BrhName").ToString
                        objExcel.Cells(2, 2) = .Item("BrhAddr").ToString
                        objExcel.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString


                        If .Item("RptType") = 2 Then
                            objExcel.Cells(5, 1) = "INVOICE/VOUCHER SUMMARY " & " (YEAR " & .Item("RptYear1") & " WEEK " & .Item("RptWeek1") & " to YEAR " & .Item("RptYear2") & " WEEK " & .Item("RptWeek2") & ")"
                        Else
                            objExcel.Cells(5, 1) = "INVOICE/VOUCHER SUMMARY " & " (YEAR " & .Item("RptYear") & "; MONTH " & .Item("RptMonth") & ")"
                        End If
                    End With

                End If

                'Setting - bold
                objExcel.Range("A1:I6").Font.Bold = True
                objExcel.Range("A1:I6").HorizontalAlignment = -4108
                objExcel.Range("A1:I1").Merge()
                objExcel.Range("A2:I2").Merge()
                objExcel.Range("A3:I3").Merge()
                objExcel.Range("A5:I5").Merge()

                iRow = 6


                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Report Data
                ' ----------------------------------------------------------------------

                TblIndex += 1

                '-- Invoice Part
                objExcel.Cells(iRow, iCol).value = "HOUSEBL"
                objExcel.Cells(iRow, iCol + 1).value = "SHIPPER"
                objExcel.Cells(iRow, iCol + 2).value = "INVOICE NO"
                objExcel.Cells(iRow, iCol + 3).value = "CURRENCY"
                objExcel.Cells(iRow, iCol + 4).value = "AMOUNT-RMB"

                objExcel.Cells(iRow, iCol + 5).value = "AMOUNT-USD"
                objExcel.Cells(iRow, iCol + 6).value = "Account Code"
                objExcel.Cells(iRow, iCol + 7).value = "Last Update User"
                objExcel.Cells(iRow, iCol + 8).value = "Week"

                ' Setting Border
                'objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 8)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 8)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 8)).HorizontalAlignment = -4108

                iRow += 1
                iSRow = iRow

                For i = 0 To (ds.Tables(TblIndex).Rows.Count - 1)
                    objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("BLNo").ToString
                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString
                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("InvNo").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("InvCurr").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("InvAmt").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("InvUSDAmt").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("AccCode").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("LstUpdUsr").ToString
                    objExcel.Cells(iRow, iCol + 8).value = ds.Tables(TblIndex).Rows(i).Item("ShpWeek").ToString

                    iRow = iRow + 1
                Next

                If iRow > iSRow Then
                    objExcel.Range(objExcel.Cells(iRow, iCol + 4), objExcel.Cells(iRow, iCol + 4)).FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                    objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 5)).FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                End If

                objExcel.Range(objExcel.Cells(iSRow, iCol + 4), objExcel.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Borders(9).LineStyle = -4119
                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True

                objExcel.Application.Columns("A:A").ColumnWidth = 14
                objExcel.Application.Columns("B:B").ColumnWidth = 30
                objExcel.Application.Columns("C:I").ColumnWidth = 14

                '-- Voucher Part
                'objExcel.Application.Sheets.Add.Move(AFTER:=objExcel.Application.Sheets(objExcel.Application.Sheets.Count))
                objExcel.Application.Sheets.Add.Move(AFTER:=objExcel.Application.ActiveSheet)

                TblIndex += 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    With ds.Tables(TblIndex).Rows(0)
                        objExcel.Cells(1, 1) = .Item("BrhName").ToString
                        objExcel.Cells(2, 2) = .Item("BrhAddr").ToString
                        objExcel.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString

                        If .Item("RptType") = 2 Then
                            objExcel.Cells(5, 1) = "INVOICE/VOUCHER SUMMARY " & " (YEAR " & .Item("RptYear1") & " WEEK " & .Item("RptWeek1") & " to YEAR " & .Item("RptYear2") & " WEEK " & .Item("RptWeek2") & ")"
                        Else
                            objExcel.Cells(5, 1) = "INVOICE/VOUCHER SUMMARY " & " (YEAR " & .Item("RptYear") & "; MONTH " & .Item("RptMonth") & ")"
                        End If
                    End With
                End If

                ' Setting - bold
                objExcel.Application.Range("A1:J6").Font.Bold = True
                objExcel.Application.Range("A1:J6").HorizontalAlignment = -4108
                objExcel.Application.Range("A1:J1").Merge()
                objExcel.Application.Range("A2:J2").Merge()
                objExcel.Application.Range("A3:J3").Merge()
                objExcel.Application.Range("A5:J5").Merge()

                iRow = 6
                '-- Invoice Part
                objExcel.Cells(iRow, iCol).value = "HOUSEBL"
                objExcel.Cells(iRow, iCol + 1).value = "CARRIER"
                objExcel.Cells(iRow, iCol + 2).value = "VOUCHER NO"
                objExcel.Cells(iRow, iCol + 3).value = "CURRENCY"
                objExcel.Cells(iRow, iCol + 4).value = "AMOUNT-RMB"
                objExcel.Cells(iRow, iCol + 5).value = "AMOUNT-USD"
                objExcel.Cells(iRow, iCol + 6).value = "Account Code"
                objExcel.Cells(iRow, iCol + 7).value = "Last Update User"
                objExcel.Cells(iRow, iCol + 8).value = "Week"

                ' Setting Border
                'objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 7)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 8)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 8)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 8)).HorizontalAlignment = -4108

                TblIndex += 1
                iRow += 1
                iSRow = iRow

                For i = 0 To (ds.Tables(TblIndex).Rows.Count - 1)
                    objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("BLNo").ToString
                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("CarName").ToString
                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("VohVouNo").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("VohCurr").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("VouAmt").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("VouUSDAmt").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("AccCode").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("LstUpdUsr").ToString
                    objExcel.Cells(iRow, iCol + 8).value = ds.Tables(TblIndex).Rows(i).Item("ShpWeek").ToString

                    iRow = iRow + 1
                Next

                If iRow > iSRow Then
                    objExcel.Range(objExcel.Cells(iRow, iCol + 4), objExcel.Cells(iRow, iCol + 4)).FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                    objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 5)).FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                End If

                objExcel.Range(objExcel.Cells(iSRow, iCol + 4), objExcel.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,###,##0.00_ "
                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Borders(9).LineStyle = -4119
                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True

                objExcel.Application.Columns("A:A").ColumnWidth = 14
                objExcel.Application.Columns("B:B").ColumnWidth = 30
                objExcel.Application.Columns("C:I").ColumnWidth = 14

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Save File
                ' ----------------------------------------------------------------------
                Dim exportPath As String = My.Settings.ExportPath & sUid & "\"
                Dim exportFile As String = ""

                If fileName <> "" Then
                    exportFile = exportPath & fileName & ".xls"
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
                MsgBox(ex.Message)
                objExcel.ActiveWorkbook.SaveAs("C:\" & UID & ".xls")
                objExcel.Quit()
                fileName = "Error," & ex.Message
            End Try
        Else
            fileName = ""
        End If

        ' Return File Path
        RptInvVou = fileName
    End Function

End Class
