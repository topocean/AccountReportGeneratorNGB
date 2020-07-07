Public Class RptTopTen

    Function RptTopTen(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, recCount, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim ReportType As Integer
        Dim ReportBy As String
        Dim ReportRg As String
        Dim ReportBrhCd As Integer
        Dim ReportRefId As Integer
        Dim ReportTraffic As Integer
        Dim ReportNormName As String
        Dim ReportDteFrm As String
        Dim ReportDteTo As String
        Dim ReportYearFm As Double
        Dim ReportYearTo As Double
        Dim ReportWeekFm As Double
        Dim ReportWeekTo As Double
        Dim ReportPeriodFm As Double
        Dim ReportPeriodTo As Double
        Dim TraName As String

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
                    objExcel.Cells(5, 1).value = "LIFTING REPORT for the Period of " & .Item("PeriodFm").ToString & " to " & .Item("PeriodTo").ToString

                    ' Get Report Parameter
                    ReportBy = .Item("ReportBy").ToString
                    ReportRg = .Item("ReportRg").ToString

                    Select Case ReportBy
                        Case "0"
                            ReportBy = "Agent"
                        Case "1"
                            ReportBy = "Carrier"
                        Case "2"
                            ReportBy = "Consignee"
                        Case "3"
                            ReportBy = "Nomination"
                        Case "4"
                            ReportBy = "Sales"
                        Case "5"
                            ReportBy = "Shipper"
                        Case "6"
                            ReportBy = "Traffic"
                    End Select

                    Select Case ReportRg
                        Case "0"
                            ReportRg = "All"
                        Case "1"
                            ReportRg = "Top Ten"
                        Case "2"
                            ReportRg = "Specific"
                    End Select

                    ' Setting - bold
                    objExcel.Range("A1:I5").Font.Bold = True
                    objExcel.Range("A1:I5").HorizontalAlignment = -4108
                    objExcel.Range("A1:I1").Merge()
                    objExcel.Range("A2:I2").Merge()
                    objExcel.Range("A3:I3").Merge()
                    objExcel.Range("A5:I5").Merge()

                    iRow = 8
                    objExcel.Cells(iRow, 1).value = "Report Range : "
                    objExcel.Cells(iRow, 3).value = ReportRg
                    objExcel.Cells(iRow + 1, 1).value = "Report Type : "
                    If ReportRg = "Specific" Then
                        objExcel.Cells(iRow + 1, 3).value = "Detail"
                    Else
                        objExcel.Cells(iRow + 1, 3).value = "Summary"
                    End If
                    objExcel.Cells(iRow + 2, 1).value = "Report By :"
                    objExcel.Cells(iRow + 2, 3).value = ReportBy
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow + 2, 3)).Font.Bold = True

                    iRow = 12
                End With

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Export Report Data
                ' ----------------------------------------------------------------------

                TblIndex += 1

                If ReportRg = "All" Or ReportRg = "Top Ten" Then ' Summary
                    ' Upper part
                    objExcel.Cells(iRow, iCol).value = ReportBy
                    objExcel.Cells(iRow, iCol + 1).value = "CY"
                    objExcel.Cells(iRow, iCol + 2).value = "CFS"

                    ' setting border
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 2)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 2)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 2)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 2)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 2)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 2)).Borders(11).LineStyle = 1

                    iRow = iRow + 1
                    startRow = iRow

                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("sName").ToString
                        objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("CY").ToString
                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("CFS").ToString

                        iRow = iRow + 1
                    Next

                    iRow = iRow + 1

                    objExcel.Cells(iRow, 1).value = "TOTAL"

                    objExcel.Range(objExcel.Cells(iRow, 3), objExcel.Cells(iRow, 3)).Activate()
                    objExcel.Range(objExcel.Cells(iRow, 3), objExcel.Cells(iRow, 3)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"


                    objExcel.Range(objExcel.Cells(iRow, 2), objExcel.Cells(iRow, 2)).Activate()
                    objExcel.Range(objExcel.Cells(iRow, 2), objExcel.Cells(iRow, 2)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"

                    ' Setting - bold & underline
                    objExcel.Range(objExcel.Cells(startRow, 2), objExcel.Cells(iRow, 2)).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Range(objExcel.Cells(startRow, 3), objExcel.Cells(iRow, 3)).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, 3)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 2), objExcel.Cells(iRow, 3)).Borders(9).LineStyle = -4119

                    ' setting width
                    objExcel.Columns("A:A").ColumnWidth = 50
                    objExcel.Columns("B:C").ColumnWidth = 12

                Else
                    ' Upper part
                    objExcel.Cells(iRow, iCol).value = "Traffic"
                    objExcel.Cells(iRow, iCol + 1).value = "Sales"
                    objExcel.Cells(iRow, iCol + 2).value = "Nomination"
                    objExcel.Cells(iRow, iCol + 3).value = "Consignee"
                    objExcel.Cells(iRow, iCol + 4).value = "Shipper"
                    objExcel.Cells(iRow, iCol + 5).value = "POL"
                    objExcel.Cells(iRow, iCol + 6).value = "POD"
                    objExcel.Cells(iRow, iCol + 7).value = "HBL"
                    objExcel.Cells(iRow, iCol + 8).value = "CY"
                    objExcel.Cells(iRow, iCol + 9).value = "CFS"
                    objExcel.Cells(iRow, iCol + 10).value = "Carrier"
                    objExcel.Cells(iRow, iCol + 11).value = "Agent"

                    ' setting border
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Borders(11).LineStyle = 1

                    iRow = iRow + 1
                    startRow = iRow

                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("Traffic").ToString
                        objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("Sales").ToString
                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("Nomination").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("Consignee").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("Shipper").ToString
                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("POL").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("POD").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                        objExcel.Cells(iRow, iCol + 8).value = ds.Tables(TblIndex).Rows(i).Item("CY").ToString
                        objExcel.Cells(iRow, iCol + 9).value = ds.Tables(TblIndex).Rows(i).Item("CFS").ToString
                        objExcel.Cells(iRow, iCol + 10).value = ds.Tables(TblIndex).Rows(i).Item("Carrier").ToString
                        objExcel.Cells(iRow, iCol + 11).value = ds.Tables(TblIndex).Rows(i).Item("Agent").ToString

                        iRow = iRow + 1

                    Next
                    iRow = iRow + 1

                    'Range("A7:C7").Select
                    'Selection.Font.Bold = True

                    objExcel.Cells(iRow, 8).value = "TOTAL"

                    objExcel.Range(objExcel.Cells(iRow, 9), objExcel.Cells(iRow, 9)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"

                    objExcel.Range(objExcel.Cells(iRow, 10), objExcel.Cells(iRow, 10)).FormulaR1C1 = "=SUM(R[-" & (iRow - startRow) & "]C:R[-1]C)"

                    ' Setting - bold & underline
                    objExcel.Range(objExcel.Cells(startRow, 9), objExcel.Cells(iRow, 9)).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Range(objExcel.Cells(startRow, 10), objExcel.Cells(iRow, 10)).NumberFormatLocal = "#,###,##0.00_ "
                    objExcel.Range(objExcel.Cells(iRow, 8), objExcel.Cells(iRow, 10)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 9), objExcel.Cells(iRow, 10)).Borders(9).LineStyle = -4119

                    ' setting width
                    objExcel.Columns("A:B").ColumnWidth = 10
                    objExcel.Columns("C:C").ColumnWidth = 20
                    objExcel.Columns("D:E").ColumnWidth = 40
                    objExcel.Columns("F:G").ColumnWidth = 10
                    objExcel.Columns("H:H").ColumnWidth = 20
                    objExcel.Columns("I:J").ColumnWidth = 12
                    objExcel.Columns("K:L").ColumnWidth = 40
                End If

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

                ReportType = Nothing
                ReportBy = Nothing
                ReportRg = Nothing
                ReportBrhCd = Nothing
                ReportRefId = Nothing
                ReportTraffic = Nothing
                ReportNormName = Nothing
                ReportDteFrm = Nothing
                ReportDteTo = Nothing
                ReportYearFm = Nothing
                ReportYearTo = Nothing
                ReportWeekFm = Nothing
                ReportWeekTo = Nothing
                ReportPeriodFm = Nothing
                ReportPeriodTo = Nothing
                TraName = Nothing

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
        RptTopTen = fileName
    End Function
End Class
