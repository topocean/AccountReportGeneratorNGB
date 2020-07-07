Public Class RptWeekComparison

    Function RptWeekComparison(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim str1, str2 As String
        Dim common As New common

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
                    objExcel.Cells(1, 1).value = .Item("BrhName").ToString
                    objExcel.Cells(2, 2).value = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3).value = "TEL: " & .Item("BrhTel").ToString & "  FAX: " & .Item("BrhFax").ToString
                    objExcel.Cells(5, 1).value = "WEEK COMPARISON BY CONSIGNEE"

                    If CInt(.Item("Figure1").ToString) = 0 Then
                        str1 = "Prospect"
                    Else
                        If CInt(.Item("Figure1").ToString) = 1 Then
                            str1 = "Booking"
                        Else
                            str1 = "Loading"
                        End If
                    End If

                    If CInt(.Item("Figure2").ToString) = 0 Then
                        str2 = "Prospect"
                    Else
                        If CInt(.Item("Figure2").ToString) = 1 Then
                            str2 = "Booking"
                        Else
                            str2 = "Loading"
                        End If
                    End If

                    objExcel.Cells(6, 1).value = .Item("YearFrm").ToString & .Item("WeekFrm").ToString & " " & str1 & " Vs " & .Item("YearTo").ToString & .Item("WeekTo").ToString & " " & str2

                    ' Setting - bold
                    objExcel.Range("A1:F5").Font.Bold = True
                    objExcel.Range("A1:F6").HorizontalAlignment = -4108
                    objExcel.Range("A1:F1").Merge()
                    objExcel.Range("A2:F2").Merge()
                    objExcel.Range("A3:F3").Merge()
                    objExcel.Range("A5:F5").Merge()
                    objExcel.Range("A6:F6").Merge()

                    iRow = 8
                    ' Upper part title
                    objExcel.Cells(iRow, iCol).value = "CONSIGNEE"
                    objExcel.Cells(iRow, iCol + 1).value = "SALES"
                    objExcel.Cells(iRow, iCol + 2).value = .Item("YearFrm").ToString & .Item("WeekFrm").ToString & " " & str1
                    objExcel.Cells(iRow, iCol + 3).value = .Item("YearTo").ToString & .Item("WeekTo").ToString & " " & str2
                    objExcel.Cells(iRow, iCol + 4).value = "VARIANCE"
                    objExcel.Cells(iRow, iCol + 5).value = "REMARKS"

                    ' setting border
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 5)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 5)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 5)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 5)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 5)).Borders(11).LineStyle = 1
                End With

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Export Data
                ' ----------------------------------------------------------------------

                TblIndex = 1
                iRow = iRow + 1
                startRow = iRow

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    If Not CDbl((ds.Tables(TblIndex).Rows(i).Item("SCompare").ToString) = 0 And CDbl(ds.Tables(TblIndex).Rows(i).Item("ECompare").ToString) = 0) Then
                        objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ConName").ToString
                        objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("Sales").ToString
                        objExcel.Cells(iRow, iCol + 2).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("SCompare").ToString)
                        objExcel.Cells(iRow, iCol + 3).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("ECompare").ToString)
                        objExcel.Cells(iRow, iCol + 4).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("ECompare").ToString) - CDbl(ds.Tables(TblIndex).Rows(i).Item("SCompare").ToString)
                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Remarks").ToString
                        iRow = iRow + 1
                    End If
                Next

                iRow = iRow + 1
                objExcel.Cells(iRow, 2).value = "TOTAL"

                For i = 3 To 5
                    objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"
                Next

                ' Setting - bold & underline
                objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 4)).Borders(9).LineStyle = -4119

                ' setting width
                objExcel.Columns("A:A").ColumnWidth = 35
                objExcel.Columns("B:B").ColumnWidth = 15
                objExcel.Columns("C:G").ColumnWidth = 17
                objExcel.Columns("H:H").ColumnWidth = 20

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
        RptWeekComparison = fileName
    End Function
End Class
