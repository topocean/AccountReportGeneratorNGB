Public Class RptLoading

    Function RptLoading(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim SubBrhCd, traffic As String
        Dim POType As String = ""
        Dim LocName As String = ""
        Dim OCF, Mode As String
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
                    fileName = common.NullVal(.Item("RptFile").ToString, UID)
                    SubBrhCd = common.NullVal(.Item("HdrBranch").ToString, "")
                    traffic = common.NullVal(.Item("HdrTraffic").ToString, "")

                    OCF = common.NullVal(.Item("HdrOCF").ToString, "")
                    Mode = common.NullVal(.Item("HdrMode").ToString, "")
                    POType = common.NullVal(.Item("HdrPOType").ToString, "")

                    ' ----------------------------------------------------------------------
                    ' Report Header (Company Name, Address, Tel, etc...)
                    ' ----------------------------------------------------------------------

                    objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                    objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                    objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                    objWS.Application.Cells(5, 1) = "LOADING REPORT - " & SubBrhCd & " (YEAR: " & common.NullVal(.Item("HdrYear").ToString, "") & " , MONTH: " & common.NullVal(.Item("HdrMonth").ToString, "") & " , WEEK: " & common.NullVal(.Item("HdrWeek").ToString, "") & ")"
                    objWS.Application.Cells(6, 1) = "TRAFFIC: " & traffic & " , MODE: " & Mode & " , OCF: " & OCF & " , PO Type: " & POType

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
                ' Column Headers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol) = "MONTH"
                objWS.Application.Cells(iRow, iCol + 1) = "WEEK"
                objWS.Application.Cells(iRow, iCol + 2) = "BRANCH"
                objWS.Application.Cells(iRow, iCol + 3) = "TRAFFIC"
                objWS.Application.Cells(iRow, iCol + 4) = "BY"
                objWS.Application.Cells(iRow, iCol + 5) = "AGENT"
                objWS.Application.Cells(iRow, iCol + 6) = "SHIPPER"
                objWS.Application.Cells(iRow, iCol + 7) = "CONSIGNEE"
                objWS.Application.Cells(iRow, iCol + 8) = "NOMINATION"
                objWS.Application.Cells(iRow, iCol + 9) = "'20"
                objWS.Application.Cells(iRow, iCol + 10) = "'40"
                objWS.Application.Cells(iRow, iCol + 11) = "'45"
                objWS.Application.Cells(iRow, iCol + 12) = "HQ"
                objWS.Application.Cells(iRow, iCol + 13) = "FEUS"
                objWS.Application.Cells(iRow, iCol + 14) = "CONSOL"
                objWS.Application.Cells(iRow, iCol + 15) = "CBM"
                objWS.Application.Cells(iRow, iCol + 16) = "TYPE"
                objWS.Application.Cells(iRow, iCol + 17) = "CARRIER"
                objWS.Application.Cells(iRow, iCol + 18) = "SO NO"
                objWS.Application.Cells(iRow, iCol + 19) = "VESSEL"
                objWS.Application.Cells(iRow, iCol + 20) = "VOYAGE"
                objWS.Application.Cells(iRow, iCol + 21) = "ETD"
                objWS.Application.Cells(iRow, iCol + 22) = "ETA"
                objWS.Application.Cells(iRow, iCol + 23) = "POL"
                objWS.Application.Cells(iRow, iCol + 24) = "DEST"
                objWS.Application.Cells(iRow, iCol + 25) = "MBL BOOK TO"
                objWS.Application.Cells(iRow, iCol + 26) = "HOUSE BL#"
                objWS.Application.Cells(iRow, iCol + 27) = "MASTER BL#"
                objWS.Application.Cells(iRow, iCol + 28) = "CONTAINER #"
                objWS.Application.Cells(iRow, iCol + 29) = "STATUS"
                objWS.Application.Cells(iRow, iCol + 30) = "UPDATE"
                objWS.Application.Cells(iRow, iCol + 31) = "LOT #"
                objWS.Application.Cells(iRow, iCol + 32) = "LOCAL SALES"
                objWS.Application.Cells(iRow, iCol + 33) = "CHECK LOT NO"
                objWS.Application.Cells(iRow, iCol + 34) = "TYPE SIZE"
                objWS.Application.Cells(iRow, iCol + 35) = "CHECK DIFF"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 36)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 36)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 36)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 36)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 36)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, 36)).Borders(11).LineStyle = 1

                ' **********************************************************************

                iRow += 1
                iSRow = iRow

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    With ds.Tables(1).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BkhMonth"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhWeek"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("BkhSubBrhCd"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("BkhTraffic"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("BkhLstUsr"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = "'" & common.NullVal(.Item("AgtCd"), "")
                        objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("ShpName"), "")
                        objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("ConName"), "")
                        objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("NomName"), "")

                        If common.NullVal(.Item("Size20"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 9) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("Size20"), 0)
                        End If

                        If common.NullVal(.Item("Size40"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 10) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("Size40"), 0)
                        End If

                        If common.NullVal(.Item("Size45"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 11) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("Size45"), 0)
                        End If

                        If common.NullVal(.Item("SizeHQ"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 12) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 12) = common.NullVal(.Item("SizeHQ"), 0)
                        End If

                        objWS.Application.Cells(iRow, iCol + 14) = common.NullVal(.Item("SizeConsol"), 0)  ' Consol Box
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("LoadedCBM"), 0)
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("TypeRemark"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = common.NullVal(.Item("ClientSCAC"), "")
                        objWS.Application.Cells(iRow, iCol + 18) = "'" & common.NullVal(.Item("BkhCarrSO"), "")
                        objWS.Application.Cells(iRow, iCol + 19) = common.NullVal(.Item("VslName"), "")
                        objWS.Application.Cells(iRow, iCol + 20) = "'" & common.NullVal(.Item("VslVoy"), "")

                        ' ----------------------------------------------------------------------
                        ' ETD
                        ' ----------------------------------------------------------------------

                        'If Format(.Item("VslETD"), "yyyy/MM/dd") = "1900/01/01" Then
                        'objWS.Application.Cells(iRow, iCol + 21) = ""
                        'Else
                        objWS.Application.Cells(iRow, iCol + 21) = "'" & .Item("BkhETD")
                        'End If

                        ' **********************************************************************

                        ' ----------------------------------------------------------------------
                        ' ETA
                        ' ----------------------------------------------------------------------

                        'If common.NullVal(.Item("BkhETA"), "") = "" Then
                        'objWS.Application.Cells(iRow, iCol + 22) = ""
                        'Else
                        'If Format(.Item("BkhETA"), "yyyy/MM/dd") = "1900/01/01" Then
                        'objWS.Application.Cells(iRow, iCol + 22) = ""
                        'Else
                        objWS.Application.Cells(iRow, iCol + 22) = "'" & .Item("BkhETA")
                        'End If
                        'End If

                        ' **********************************************************************

                        objWS.Application.Cells(iRow, iCol + 23) = common.NullVal(.Item("BkhLoad"), "")
                        objWS.Application.Cells(iRow, iCol + 24) = common.NullVal(.Item("BkhDest"), "")
                        objWS.Application.Cells(iRow, iCol + 25) = common.NullVal(.Item("BkhDisc"), "")
                        objWS.Application.Cells(iRow, iCol + 26) = common.NullVal(.Item("BkhBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 27) = "'" & common.NullVal(.Item("BkhMBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 28) = common.NullVal(.Item("CtnrNo"), "")
                        objWS.Application.Cells(iRow, iCol + 29) = common.NullVal(.Item("BkhTLX"), "")

                        ' ----------------------------------------------------------------------
                        ' TLX Update date
                        ' ----------------------------------------------------------------------

                        objWS.Application.Cells(iRow, iCol + 30) = common.NullVal(.Item("BkhTLXDte"), "")

                        If InStr(UCase(CStr(common.NullVal(.Item("BkhTLXDte"), ""))), UCase(CStr(Format(Now, "dd-MMM-yyyy")))) > 0 Then
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 30), objWS.Application.Cells(iRow, iCol + 30)).Interior.Color = RGB(229, 229, 231)
                        End If

                        ' **********************************************************************

                        objWS.Application.Cells(iRow, iCol + 31) = "'" & common.NullVal(.Item("BkhLotNo"), "")
                        objWS.Application.Cells(iRow, iCol + 32) = common.NullVal(.Item("NomSales"), "")

                    End With

                    ' ----------------------------------------------------------------------
                    ' FEUS Calculation
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range(objWS.Application.Cells(iRow, 10), objWS.Application.Cells(iRow, 13)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, 14), objWS.Application.Cells(iRow, 14)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1],RC[-4]/2,RC[1]*-1)"

                    ' ----------------------------------------------------------------------
                    ' Check Lot No.
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range(objWS.Application.Cells(iRow, 34), objWS.Application.Cells(iRow, 34)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, 34), objWS.Application.Cells(iRow, 34)).Activate()
                    objWS.Application.ActiveCell.Formula = "=IF(AF" & iRow & " <> """", COUNTIF(AF:AF,AF" & iRow & "), 0)"

                    ' ----------------------------------------------------------------------
                    ' Type Size
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range(objWS.Application.Cells(iRow, 35), objWS.Application.Cells(iRow, 35)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, 35), objWS.Application.Cells(iRow, 35)).Activate()
                    objWS.Application.ActiveCell.FormulaLocal = "=LEFT(Q" & iRow & ", 2)"

                    ' ----------------------------------------------------------------------
                    ' Check Difference
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range(objWS.Application.Cells(iRow, 36), objWS.Application.Cells(iRow, 36)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, 36), objWS.Application.Cells(iRow, 36)).Activate()
                    objWS.Application.ActiveCell.Formula = "=+AI" & iRow & "-SUM(J" & iRow & ":M" & iRow & ")"

                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1
                Next

                ' ----------------------------------------------------------------------
                ' Calculate the total number of containers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 8) = "TOTAL: "

                For i = 10 To 16
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                Next

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Underline the total columns)
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 9), objWS.Application.Cells(iRow, iCol + 15)).Borders(9).LineStyle = -4119

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Column Width)
                ' ----------------------------------------------------------------------

                objWS.Application.Range("A8:B8").ColumnWidth = 8
                objWS.Application.Range("C8:C8").ColumnWidth = 14
                objWS.Application.Range("D8:D8").ColumnWidth = 10
                objWS.Application.Range("E8:E8").ColumnWidth = 15
                objWS.Application.Range("F8:F8").ColumnWidth = 10
                objWS.Application.Range("G8:H8").ColumnWidth = 40
                objWS.Application.Range("I8:I8").ColumnWidth = 15
                objWS.Application.Range("J8:O8").ColumnWidth = 8
                objWS.Application.Range("P8:R8").ColumnWidth = 10
                objWS.Application.Range("S8:T8").ColumnWidth = 15
                objWS.Application.Range("U8:U8").ColumnWidth = 10
                objWS.Application.Range("V8:W8").ColumnWidth = 15
                objWS.Application.Range("X8:AC8").ColumnWidth = 20
                objWS.Application.Range("AD8:AD8").ColumnWidth = 10
                objWS.Application.Range("AE8:AF8").ColumnWidth = 15
                objWS.Application.Range("AG8:AG8").ColumnWidth = 30
                objWS.Application.Range("AH8:AJ8").ColumnWidth = 15

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
        RptLoading = fileName
    End Function
End Class
