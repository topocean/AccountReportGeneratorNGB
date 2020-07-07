Public Class RptISF

    Function RptISF(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String
        Dim SubBrhCd As String
        Dim POType As String = ""
        Dim common As New common
        Dim hasData As Boolean = False

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
                SubBrhCd = common.NullVal(.Item("HdrBranch").ToString, "")

                ' ----------------------------------------------------------------------
                ' Report Header (Company Name, Address, Tel, etc...)
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                objWS.Application.Cells(5, 1) = "ISF BOOKING REPORT - " & SubBrhCd & " (YEAR: " & common.NullVal(.Item("HdrYear").ToString, "") & " , WEEK: " & common.NullVal(.Item("HdrWeek").ToString, "") & ")"

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
            ' Column Headers (CY Shipments)
            ' ----------------------------------------------------------------------

            If ds.Tables(1).Rows.Count > 0 Then
                hasData = True

                objWS.Application.Cells(iRow, iCol) = "FCL SHIPMENTS"
                iRow += 1

                objWS.Application.Cells(iRow, iCol) = "BRANCH"
                objWS.Application.Cells(iRow, iCol + 1) = "BKG DATE"
                objWS.Application.Cells(iRow, iCol + 2) = "AGENT"
                objWS.Application.Cells(iRow, iCol + 3) = "COLOAD"
                objWS.Application.Cells(iRow, iCol + 4) = "SHIPPER"
                objWS.Application.Cells(iRow, iCol + 5) = "CONSIGNEE"
                objWS.Application.Cells(iRow, iCol + 6) = "SI CUT OFF"
                objWS.Application.Cells(iRow, iCol + 7) = "NOMINATION"
                objWS.Application.Cells(iRow, iCol + 8) = "'20"
                objWS.Application.Cells(iRow, iCol + 9) = "'40"
                objWS.Application.Cells(iRow, iCol + 10) = "'45"
                objWS.Application.Cells(iRow, iCol + 11) = "'HQ"
                objWS.Application.Cells(iRow, iCol + 12) = "FEUS"
                objWS.Application.Cells(iRow, iCol + 13) = "TYPE"
                objWS.Application.Cells(iRow, iCol + 14) = "CARRIER"
                objWS.Application.Cells(iRow, iCol + 15) = "HBL#"
                objWS.Application.Cells(iRow, iCol + 16) = "ISF CUSTOMER"
                objWS.Application.Cells(iRow, iCol + 17) = "SERVICE OPTION"
                objWS.Application.Cells(iRow, iCol + 18) = "VESSEL"
                objWS.Application.Cells(iRow, iCol + 19) = "VOYAGE"
                objWS.Application.Cells(iRow, iCol + 20) = "ETD"
                objWS.Application.Cells(iRow, iCol + 21) = "ETA"
                objWS.Application.Cells(iRow, iCol + 22) = "DEST"
                objWS.Application.Cells(iRow, iCol + 23) = "POL"
                objWS.Application.Cells(iRow, iCol + 24) = "MBL DEST"
                objWS.Application.Cells(iRow, iCol + 25) = "PO#"
                objWS.Application.Cells(iRow, iCol + 26) = "NOTES"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 27)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 27)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 27)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 27)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 27)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 27)).Borders(11).LineStyle = 1

                ' **********************************************************************

                iRow += 1
                iSRow = iRow

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    With ds.Tables(1).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BkhSubBrh"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhCreDte"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("AgtCode"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("BkhCO"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("ShpName"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("ConName"), "")
                        objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("IsfSICutOff"), "")
                        objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("NomName"), "")
                        objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("Qty20"), 0)
                        objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("Qty40"), 0)
                        objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("Qty45"), 0)
                        objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("QtyHQ"), 0)
                        objWS.Application.Cells(iRow, iCol + 13) = common.NullVal(.Item("TypeRemark"), "")
                        objWS.Application.Cells(iRow, iCol + 14) = common.NullVal(.Item("ClientSCAC"), "")
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("BkhBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("IsTPTCust"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = common.NullVal(.Item("ServiceOption"), "")
                        objWS.Application.Cells(iRow, iCol + 18) = common.NullVal(.Item("VslName"), "")
                        objWS.Application.Cells(iRow, iCol + 19) = "'" & common.NullVal(.Item("VslVoy"), "")
                        objWS.Application.Cells(iRow, iCol + 20) = common.NullVal(.Item("BkhETD"), "")
                        objWS.Application.Cells(iRow, iCol + 21) = common.NullVal(.Item("BkhETA"), "")
                        objWS.Application.Cells(iRow, iCol + 22) = common.NullVal(.Item("BkhDest"), "")
                        objWS.Application.Cells(iRow, iCol + 23) = common.NullVal(.Item("BkhLoad"), "")
                        objWS.Application.Cells(iRow, iCol + 24) = common.NullVal(.Item("BkhDisc"), "")
                        objWS.Application.Cells(iRow, iCol + 25) = "'" & common.NullVal(.Item("IsfPO"), "")
                        objWS.Application.Cells(iRow, iCol + 26) = common.NullVal(.Item("Notes"), "")
                    End With

                    ' ----------------------------------------------------------------------
                    ' FEUS Calculation
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range(objWS.Application.Cells(iRow, 9), objWS.Application.Cells(iRow, 12)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, 13), objWS.Application.Cells(iRow, 13)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1],RC[-4]/2)"

                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1
                Next

                ' ----------------------------------------------------------------------
                ' Calculate the total number of containers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 7) = "TOTAL: "

                For i = 9 To 13
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                Next

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Underline the total columns)
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 8), objWS.Application.Cells(iRow, iCol + 12)).Borders(9).LineStyle = -4119

                ' **********************************************************************

            End If

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Column Headers (CFS Shipments)
            ' ----------------------------------------------------------------------

            If ds.Tables(2).Rows.Count > 0 Then
                iRow += 5

                hasData = True

                objWS.Application.Cells(iRow, iCol) = "LCL SHIPMENTS"
                iRow += 1

                objWS.Application.Cells(iRow, iCol) = "BRANCH"
                objWS.Application.Cells(iRow, iCol + 1) = "BKG DATE"
                objWS.Application.Cells(iRow, iCol + 2) = "AGENT"
                objWS.Application.Cells(iRow, iCol + 3) = "COLOAD"
                objWS.Application.Cells(iRow, iCol + 4) = "SHIPPER"
                objWS.Application.Cells(iRow, iCol + 5) = "CONSIGNEE"
                objWS.Application.Cells(iRow, iCol + 6) = "SI CUT OFF"
                objWS.Application.Cells(iRow, iCol + 7) = "NOMINATION"
                objWS.Application.Cells(iRow, iCol + 8) = "CBM"
                objWS.Application.Cells(iRow, iCol + 9) = "CARRIER"
                objWS.Application.Cells(iRow, iCol + 10) = "HBL#"
                objWS.Application.Cells(iRow, iCol + 11) = "ISF CUSTOMER"
                objWS.Application.Cells(iRow, iCol + 12) = "SERVICE OPTION"
                objWS.Application.Cells(iRow, iCol + 13) = "VESSEL"
                objWS.Application.Cells(iRow, iCol + 14) = "VOYAGE"
                objWS.Application.Cells(iRow, iCol + 15) = "ETD"
                objWS.Application.Cells(iRow, iCol + 16) = "ETA"
                objWS.Application.Cells(iRow, iCol + 17) = "DEST"
                objWS.Application.Cells(iRow, iCol + 18) = "POL"
                objWS.Application.Cells(iRow, iCol + 19) = "MBL DEST"
                objWS.Application.Cells(iRow, iCol + 20) = "PO#"
                objWS.Application.Cells(iRow, iCol + 21) = "NOTES"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 22)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 22)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 22)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 22)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 22)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 22)).Borders(11).LineStyle = 1

                ' **********************************************************************

                iRow += 1
                iSRow = iRow

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(2).Rows.Count - 1
                    With ds.Tables(2).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BkhSubBrh"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhCreDte"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("AgtCode"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("BkhCO"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("ShpName"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("ConName"), "")
                        objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("IsfSICutOff"), "")
                        objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("NomName"), "")
                        objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("BkhCBM"), 0)
                        objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("ClientSCAC"), "")
                        objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("BkhBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("IsTPTCust"), "")
                        objWS.Application.Cells(iRow, iCol + 12) = common.NullVal(.Item("ServiceOption"), "")
                        objWS.Application.Cells(iRow, iCol + 13) = common.NullVal(.Item("VslName"), "")
                        objWS.Application.Cells(iRow, iCol + 14) = "'" & common.NullVal(.Item("VslVoy"), "")
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("BkhETD"), "")
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("BkhETA"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = common.NullVal(.Item("BkhDest"), "")
                        objWS.Application.Cells(iRow, iCol + 18) = common.NullVal(.Item("BkhLoad"), "")
                        objWS.Application.Cells(iRow, iCol + 19) = common.NullVal(.Item("BkhDisc"), "")
                        objWS.Application.Cells(iRow, iCol + 20) = "'" & common.NullVal(.Item("IsfPO"), "")
                        objWS.Application.Cells(iRow, iCol + 21) = common.NullVal(.Item("Notes"), "")
                    End With

                    iRow += 1
                Next

                ' ----------------------------------------------------------------------
                ' Calculate the total number of CBM
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 7) = "TOTAL: "

                For i = 9 To 9
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                Next

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Underline the total columns)
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 8), objWS.Application.Cells(iRow, iCol + 8)).Borders(9).LineStyle = -4119

                ' **********************************************************************

            End If

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Setting Properties (Column Width)
            ' ----------------------------------------------------------------------

            objWS.Application.Range("A8:A8").ColumnWidth = 13
            objWS.Application.Range("B8:B8").ColumnWidth = 12
            objWS.Application.Range("C8:C8").ColumnWidth = 8
            objWS.Application.Range("D8:D8").ColumnWidth = 13
            objWS.Application.Range("E8:F8").ColumnWidth = 40
            objWS.Application.Range("G8:G8").ColumnWidth = 12
            objWS.Application.Range("H8:H8").ColumnWidth = 14
            objWS.Application.Range("I8:M8").ColumnWidth = 13
            objWS.Application.Range("N8:O8").ColumnWidth = 12
            objWS.Application.Range("P8:AA8").ColumnWidth = 20

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Save File
            ' ----------------------------------------------------------------------

            If Not hasData Then
                fileName = ""
            End If

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
            POType = Nothing
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

        ' Return File Path
        RptISF = fileName
    End Function
End Class
