Public Class RptBooking

    Function RptBooking(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, startRow As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim SubBrhCd, traffic As String
        Dim POType As String = ""
        Dim LocName As String = ""
        Dim OCF As String
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
                ' Get File Name, Sub-Branch, Traffic and Location
                ' ----------------------------------------------------------------------

                With ds.Tables(0).Rows(0)
                    fileName = .Item("RptFile").ToString

                    objExcel.Cells(1, 1) = .Item("BrhName").ToString
                    objExcel.Cells(2, 2) = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString

                    traffic = .Item("Traffic").ToString
                    SubBrhCd = CInt(.Item("SubBrhCd").ToString)

                    If traffic = 0 Then
                        If SubBrhCd = 0 Then
                            objExcel.Cells(5, 1) = "BOOKING REPORT - ALL, TRAFFIC - NON-USA (WEEK " & .Item("HdrWeek").ToString & ")"
                        Else
                            objExcel.Cells(5, 1) = "BOOKING REPORT - " & .Item("HdrBranch").ToString & ", TRAFFIC - NON-USA (WEEK " & .Item("HdrWeek").ToString & ")"
                        End If
                    Else
                        If SubBrhCd = 0 Then
                            objExcel.Cells(5, 1) = "BOOKING REPORT - ALL, TRAFFIC - USA (WEEK " & .Item("HdrWeek").ToString & ")"
                        Else
                            objExcel.Cells(5, 1) = "BOOKING REPORT - " & .Item("HdrBranch").ToString & ", TRAFFIC - USA (WEEK " & .Item("HdrWeek").ToString & ")"
                        End If
                    End If

                    objExcel.Application.Range("A1:I5").Font.Bold = True
                    objExcel.Application.Range("A1:I5").HorizontalAlignment = -4108
                    objExcel.Application.Range("A1:I1").Merge()
                    objExcel.Application.Range("A2:I2").Merge()
                    objExcel.Application.Range("A3:I3").Merge()
                    objExcel.Application.Range("A5:I5").Merge()

                    iRow = 7

                    objExcel.Cells(iRow, iCol) = "BRANCH"
                    objExcel.Cells(iRow, iCol + 1) = "PO"
                    objExcel.Cells(iRow, iCol + 2) = "NOMINATION"
                    objExcel.Cells(iRow, iCol + 3) = "TRAFFIC"
                    objExcel.Cells(iRow, iCol + 4) = "BKG_DATE"

                    objExcel.Cells(iRow, iCol + 5) = "AGENT"
                    objExcel.Cells(iRow, iCol + 6) = "SHIPPER"
                    objExcel.Cells(iRow, iCol + 7) = "CONSIGNEE"
                    objExcel.Cells(iRow, iCol + 8) = "'20"
                    objExcel.Cells(iRow, iCol + 9) = "'40"

                    objExcel.Cells(iRow, iCol + 10) = "'45"
                    objExcel.Cells(iRow, iCol + 11) = "'HQ"
                    objExcel.Cells(iRow, iCol + 12) = "FEUS"
                    objExcel.Cells(iRow, iCol + 13) = "TYPE"
                    objExcel.Cells(iRow, iCol + 14) = "CARRIER"

                    objExcel.Cells(iRow, iCol + 15) = "SO_NO"
                    objExcel.Cells(iRow, iCol + 16) = "VESSEL"
                    objExcel.Cells(iRow, iCol + 17) = "VOYAGE"
                    objExcel.Cells(iRow, iCol + 18) = "ETD"
                    objExcel.Cells(iRow, iCol + 19) = "ETA"

                    objExcel.Cells(iRow, iCol + 20) = "DEST"
                    objExcel.Cells(iRow, iCol + 21) = "DISC"
                    objExcel.Cells(iRow, iCol + 22) = "POL"
                    objExcel.Cells(iRow, iCol + 23) = "MBL_DEST"
                    objExcel.Cells(iRow, iCol + 24) = "HBL NO"

                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 24)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 24)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 24)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 24)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 24)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 24)).Borders(11).LineStyle = 1

                    iRow = iRow + 1
                    startRow = iRow


                    ' ----------------------------------------------------------------------
                    ' Export Data
                    ' ----------------------------------------------------------------------

                    For i = 0 To ds.Tables(1).Rows.Count - 1
                        With ds.Tables(1).Rows(i)
                            objExcel.Cells(iRow, iCol) = .Item("BkhSubBrh").ToString
                            objExcel.Cells(iRow, iCol + 1) = .Item("PO").ToString
                            objExcel.Cells(iRow, iCol + 2) = "'" & .Item("NomName").ToString
                            objExcel.Cells(iRow, iCol + 3) = .Item("BkhTraffic").ToString
                            objExcel.Cells(iRow, iCol + 4) = "'" & .Item("BkhCreDte").ToString

                            objExcel.Cells(iRow, iCol + 5) = .Item("AgtCode").ToString
                            objExcel.Cells(iRow, iCol + 6) = .Item("ShpName").ToString
                            objExcel.Cells(iRow, iCol + 7) = .Item("ConName").ToString
                            objExcel.Cells(iRow, iCol + 8) = .Item("Qty20").ToString
                            objExcel.Cells(iRow, iCol + 9) = .Item("Qty40").ToString

                            objExcel.Cells(iRow, iCol + 10) = .Item("Qty45").ToString
                            objExcel.Cells(iRow, iCol + 11) = .Item("QtyHQ").ToString
                            objExcel.Cells(iRow, iCol + 12) = .Item("Feus").ToString
                            objExcel.Cells(iRow, iCol + 13) = .Item("TypeRemark").ToString
                            objExcel.Cells(iRow, iCol + 14) = .Item("ClientSCAC").ToString

                            objExcel.Cells(iRow, iCol + 15) = "'" & .Item("BkhCarrNote").ToString
                            objExcel.Cells(iRow, iCol + 16) = "'" & .Item("VslName").ToString
                            objExcel.Cells(iRow, iCol + 17) = "'" & .Item("VslVoy").ToString
                            objExcel.Cells(iRow, iCol + 18) = "'" & .Item("BkhETD").ToString
                            objExcel.Cells(iRow, iCol + 19) = "'" & .Item("BkhETA").ToString

                            objExcel.Cells(iRow, iCol + 20) = .Item("BkhDest").ToString
                            objExcel.Cells(iRow, iCol + 21) = .Item("BkhDisc").ToString
                            objExcel.Cells(iRow, iCol + 22) = .Item("BkhLoad").ToString
                            objExcel.Cells(iRow, iCol + 23) = .Item("BkhDiscard").ToString
                            objExcel.Cells(iRow, iCol + 24) = .Item("BkhBLNo").ToString
                        End With

                        iRow = iRow + 1

                    Next

                    'iRow = iRow + 1
                    objExcel.Cells(iRow, iCol + 6) = "TOTAL"

                    For i = (iCol + 8) To (iCol + 12)
                        objExcel.Range(objExcel.Cells(startRow, i), objExcel.Cells(iRow - 2, i)).Select()
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                        'objExcel.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (iRow - startRow) & "]C:R[-1]C)"
                        objExcel.Cells(iRow, i).Formula = "=SUBTOTAL(9,R[-" & (iRow - startRow) & "]C:R[-1]C)"
                    Next

                    'Setting - bold & underline
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, 13)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 9), objExcel.Cells(iRow, 13)).Borders(9).LineStyle = -4119

                    objExcel.Columns("A:B").ColumnWidth = 16
                    objExcel.Columns("C:F").ColumnWidth = 14
                    objExcel.Columns("G:H").ColumnWidth = 30
                    objExcel.Columns("I:M").ColumnWidth = 10
                    objExcel.Columns("N:O").ColumnWidth = 12
                    objExcel.Columns("P:R").ColumnWidth = 20
                    objExcel.Columns("S:T").ColumnWidth = 14
                    objExcel.Columns("U:Y").ColumnWidth = 20

                End With

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
                traffic = Nothing
                OCF = Nothing
                POType = Nothing
                LocName = Nothing
                i = Nothing
                SubBrhCd = Nothing
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
        RptBooking = fileName
    End Function
End Class
