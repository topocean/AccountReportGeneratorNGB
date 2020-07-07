Public Class RptManifest

    Function RptManifest(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim tmpRefId, tmpRefId2, iCount As Integer
        Dim BkhWare, BkhDest, str1, str2, str3 As String

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
                    objExcel.Cells(3, 3) = "TEL: " & ds.Tables(0).Rows(0).Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString

                    objExcel.Cells(5, 1) = "MANIFEST CFS"

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

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    If CInt(ds.Tables(TblIndex).Rows(i).Item("VscRefId").ToString) <> tmpRefId Or CInt(ds.Tables(TblIndex).Rows(i).Item("BkhReceipt").ToString) <> BkhWare Or (ds.Tables(TblIndex).Rows(i).Item("BkhDest").ToString) <> BkhDest Then

                        BkhWare = CInt(ds.Tables(TblIndex).Rows(i).Item("BkhReceipt").ToString)
                        BkhDest = CInt(ds.Tables(TblIndex).Rows(i).Item("BkhDest").ToString)
                        tmpRefId = ds.Tables(TblIndex).Rows(i).Item("VscRefId")

                        iRow = iRow + 1
                        ' Vessel Header & Info
                        objExcel.Cells(iRow, iCol).value = "VESSEL/VOYAGE :"
                        objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("VslName").ToString & "/" & ds.Tables(TblIndex).Rows(i).Item("VslVoyName").ToString
                        objExcel.Cells(iRow + 1, iCol).value = "ETD"
                        objExcel.Cells(iRow + 1, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("BkhReceiptName") & " : " & Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("VslETD")), "MM/dd")
                        objExcel.Cells(iRow + 1, iCol + 2).value = "ETA"
                        objExcel.Cells(iRow + 1, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("BkhDestName").ToString & " : " & Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("VslETA")), "MM/dd")
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow + 1, iCol + 3)).Font.Bold = True

                        ' Container Header
                        iRow = iRow + 2
                        objExcel.Cells(iRow, iCol).value = "MASTER BL#"
                        objExcel.Cells(iRow, iCol + 1).value = "CONTAINER# / SEAL#"
                        objExcel.Cells(iRow, iCol + 2).value = "HOUSE BL#"
                        objExcel.Cells(iRow, iCol + 3).value = "TLX"
                        objExcel.Cells(iRow, iCol + 4).value = "SHIPPER"
                        objExcel.Cells(iRow, iCol + 5).value = "CONSIGEE"
                        objExcel.Cells(iRow, iCol + 6).value = "DEST"
                        objExcel.Cells(iRow, iCol + 7).value = "PKG"
                        objExcel.Cells(iRow, iCol + 8).value = ""
                        objExcel.Cells(iRow, iCol + 9).value = "KGS"
                        objExcel.Cells(iRow, iCol + 10).value = "CBM"
                        objExcel.Cells(iRow, iCol + 11).value = "AMS#"
                        objExcel.Cells(iRow, iCol + 12).value = "REMARK"

                        ' Setting Top Border
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Interior.ColorIndex = 15
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(8).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(10).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(11).LineStyle = 1
                    End If

                    ' Document No.
                    objExcel.Cells(iRow + 1, iCol).value = ""

                    iRow = iRow + 2
                    tmpRefId2 = CInt(ds.Tables(TblIndex).Rows(i).Item("CtnrRefId").ToString)
                    startRow = iRow
                    iCount = 0

                    Do While CInt(ds.Tables(TblIndex).Rows(i).Item("CtnrRefId").ToString) = tmpRefId2
                        iCount += 1
                        ' Container Info
                        str1 = "'" & ds.Tables(TblIndex).Rows(i).Item("SealNo").ToString
                        str2 = "'" & ds.Tables(TblIndex).Rows(i).Item("CtsName").ToString
                        str3 = ""

                        If iCount = 1 Then
                            objExcel.Cells(iRow, iCol).value = "'" & ds.Tables(TblIndex).Rows(i).Item("CtnrMBLNo").ToString
                            objExcel.Cells(iRow, iCol + 1).value = "'" & ds.Tables(TblIndex).Rows(i).Item("CtnrNo").ToString
                        ElseIf iCount = 2 Then
                            objExcel.Cells(iRow, iCol).value = str3
                            objExcel.Cells(iRow, iCol + 1).value = str1
                        ElseIf iCount = 3 Then
                            objExcel.Cells(iRow, iCol + 1).value = str2
                        End If

                        objExcel.Cells(iRow, iCol + 2).value = "'" & ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                        objExcel.Cells(iRow, iCol + 4).value = "'" & ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString
                        objExcel.Cells(iRow, iCol + 5).value = "'" & ds.Tables(TblIndex).Rows(i).Item("ConName").ToString
                        objExcel.Cells(iRow, iCol + 6).value = "'" & ds.Tables(TblIndex).Rows(i).Item("BkhDestName").ToString
                        objExcel.Cells(iRow, iCol + 7).value = Convert.ToSingle(ds.Tables(TblIndex).Rows(i).Item("Qty"))
                        objExcel.Cells(iRow, iCol + 8).value = "'" & ds.Tables(TblIndex).Rows(i).Item("UntName").ToString
                        objExcel.Cells(iRow, iCol + 9).value = Convert.ToDecimal(ds.Tables(TblIndex).Rows(i).Item("WGT"))
                        objExcel.Cells(iRow, iCol + 10).value = Convert.ToDecimal(ds.Tables(TblIndex).Rows(i).Item("CBM"))
                        objExcel.Cells(iRow, iCol + 11).value = "'" & ds.Tables(TblIndex).Rows(i).Item("AMS").ToString
                        objExcel.Cells(iRow, iCol + 12).value = "'" & ds.Tables(TblIndex).Rows(i).Item("Remark").ToString

                        iRow = iRow + 1
                        i = i + 1

                        If i >= ds.Tables(TblIndex).Rows.Count Then
                            Exit Do
                        End If
                    Loop

                    If iCount < 2 Then
                        objExcel.Cells(iRow, iCol).value = str3
                        objExcel.Cells(iRow, iCol + 1).value = str1
                        objExcel.Cells(iRow + 1, iCol + 1).value = str2
                        iRow = iRow + 2
                    ElseIf iCount < 3 Then
                        objExcel.Cells(iRow, iCol + 1).value = str2
                        iRow = iRow + 1
                    End If

                    For j = 8 To 11
                        If j <> 9 Then

                            objExcel.Range(objExcel.Cells(iRow, j), objExcel.Cells(iRow, j)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"

                        End If
                    Next

                    ' Setting Bottom Border
                    objExcel.Range(objExcel.Cells(startRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).NumberFormatLocal = "#,###,##0_ "
                    objExcel.Range(objExcel.Cells(startRow, iCol + 9), objExcel.Cells(iRow, iCol + 9)).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Range(objExcel.Cells(startRow, iCol + 10), objExcel.Cells(iRow, iCol + 10)).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(startRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(startRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(startRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(startRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(11).LineStyle = 1

                    iRow = iRow + 1
                Next

                objExcel.Columns("A:A").ColumnWidth = 14
                objExcel.Columns("B:B").ColumnWidth = 22
                objExcel.Columns("C:C").ColumnWidth = 14
                objExcel.Columns("D:D").ColumnWidth = 10
                objExcel.Columns("E:F").ColumnWidth = 35
                objExcel.Columns("G:K").ColumnWidth = 14
                objExcel.Columns("L:L").ColumnWidth = 17
                objExcel.Columns("M:M").ColumnWidth = 25

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
        RptManifest = fileName
    End Function
End Class
