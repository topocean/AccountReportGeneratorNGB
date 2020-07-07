Public Class RptBookingCFS

    Function RptBookingCFS(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, recCount, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim hasData As Boolean = False

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

                With ds.Tables(0).Rows(0)
                    fileName = .Item("RptFile").ToString
                    recCount = CInt(.Item("RecCount").ToString)
                    objWS.Application.Cells(1, 1) = .Item("BrhName").ToString
                    objWS.Application.Cells(2, 2) = .Item("BrhAddr").ToString
                    objWS.Application.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString
                End With

                objWS.Application.Cells(5, 1) = "BOOKING REPORT CFS"

                'Setting - bold
                objWS.Application.Range("A1:G5").Font.Bold = True
                objWS.Application.Range("A1:G5").HorizontalAlignment = -4108
                objWS.Application.Range("A1:G1").Merge()
                objWS.Application.Range("A2:G2").Merge()
                objWS.Application.Range("A3:G3").Merge()
                objWS.Application.Range("A5:G5").Merge()

                iRow = 8
                i = 1
                TblIndex = 0

                While i <= recCount
                    TblIndex += 1
                    iCol = 1

                    If ds.Tables(TblIndex).Rows.Count > 0 Then
                        hasData = True

                        ' Section Header
                        objWS.Application.Cells(iRow, iCol) = "VESSEL: " & ds.Tables(TblIndex).Rows(0).Item("VslName").ToString
                        objWS.Application.Cells(iRow + 1, iCol) = "ETD: " & Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(0).Item("VslETD")), "dd/MM/yyyy")
                        objWS.Application.Cells(iRow + 2, iCol) = "POL: " & ds.Tables(TblIndex).Rows(0).Item("PolName").ToString

                        iCol = 4
                        objWS.Application.Cells(iRow, iCol).value = "VOYAGE: " & ds.Tables(TblIndex).Rows(0).Item("VslVoy").ToString
                        objWS.Application.Cells(iRow + 1, iCol).value = "ETA: " & Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(0).Item("VslETA")), "dd/mm/yyyy")
                        objWS.Application.Cells(iRow + 2, iCol).value = "POD: " & ds.Tables(TblIndex).Rows(0).Item("PodName").ToString

                        ' Section Detail
                        TblIndex += 1
                        If recCount = 1 Then
                            If ds.Tables(TblIndex).Rows.Count = 0 Then
                                hasData = False
                            End If
                        End If

                        iCol = 1
                        iRow = iRow + 4

                        objWS.Application.Cells(iRow, iCol) = "S/O#"
                        objWS.Application.Cells(iRow, iCol + 1) = "CONNORM"
                        objWS.Application.Cells(iRow, iCol + 2) = "SHPR"
                        objWS.Application.Cells(iRow, iCol + 3) = "CNEE"
                        objWS.Application.Cells(iRow, iCol + 4) = "COMMODITY"

                        objWS.Application.Cells(iRow, iCol + 5) = "PKG"
                        objWS.Application.Cells(iRow, iCol + 6) = ""
                        objWS.Application.Cells(iRow, iCol + 7) = "KG"
                        objWS.Application.Cells(iRow, iCol + 8) = "'CBM"

                        objExcel.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 8)).Interior.ColorIndex = 15
                        objExcel.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 8)).Font.Bold = True
                        objExcel.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 8)).Borders(8).LineStyle = 1
                        objExcel.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 8)).Borders(9).LineStyle = 1
                        objExcel.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 8)).Borders(10).LineStyle = 1
                        objExcel.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 8)).Borders(11).LineStyle = 1

                        iRow = iRow + 1
                        startRow = iRow

                        For j = 0 To ds.Tables(TblIndex).Rows.Count - 1
                            objWS.Application.Cells(iRow, iCol) = ds.Tables(TblIndex).Rows(j).Item("BkhSONo").ToString
                            objWS.Application.Cells(iRow, iCol + 1) = ds.Tables(TblIndex).Rows(j).Item("ConNorm").ToString
                            objWS.Application.Cells(iRow, iCol + 2) = ds.Tables(TblIndex).Rows(j).Item("ShpSName").ToString
                            objWS.Application.Cells(iRow, iCol + 3) = ds.Tables(TblIndex).Rows(j).Item("ConSName").ToString
                            objWS.Application.Cells(iRow, iCol + 4) = ds.Tables(TblIndex).Rows(j).Item("BkhCarrComm").ToString

                            objWS.Application.Cells(iRow, iCol + 5) = ds.Tables(TblIndex).Rows(j).Item("BktPkg").ToString
                            objWS.Application.Cells(iRow, iCol + 6) = ds.Tables(TblIndex).Rows(j).Item("UntName").ToString
                            objWS.Application.Cells(iRow, iCol + 7) = ds.Tables(TblIndex).Rows(j).Item("BktWgt").ToString
                            objWS.Application.Cells(iRow, iCol + 8) = ds.Tables(TblIndex).Rows(j).Item("BktCBM").ToString

                            iRow = iRow + 1
                        Next

                        ' Return Section Totals
                        objWS.Application.Cells(iRow, iCol + 4) = "TOTAL"
                        objWS.Application.Cells(iRow, iCol + 5).Formula = "=SUBTOTAL(9,R[-" & (iRow - startRow) & "]C:R[-1]C)"
                        objWS.Application.Cells(iRow, iCol + 7).Formula = "=SUBTOTAL(9,R[-" & (iRow - startRow) & "]C:R[-1]C)"
                        objWS.Application.Cells(iRow, iCol + 8).Formula = "=SUBTOTAL(9,R[-" & (iRow - startRow) & "]C:R[-1]C)"

                        ' Set Propertise of Total Section
                        objWS.Application.Range(objWS.Application.Cells(startRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,###,##0_ "
                        objWS.Application.Range(objWS.Application.Cells(startRow, iCol + 7), objWS.Application.Cells(iRow, iCol + 8)).NumberFormatLocal = "#,###,##0.00_ "
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 8)).Font.Bold = True
                        objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 5), objWS.Application.Cells(iRow, iCol + 8)).Borders(9).LineStyle = -4119

                    End If

                    i = i + 1
                    iRow = iRow + 5

                End While

                objExcel.Columns("A:A").ColumnWidth = 10
                objExcel.Columns("B:B").ColumnWidth = 12
                objExcel.Columns("C:E").ColumnWidth = 30
                objExcel.Columns("F:I").ColumnWidth = 12

                ' **********************************************************************

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
        Else
            fileName = ""
        End If

        ' Return File Path
        RptBookingCFS = fileName
    End Function
End Class
