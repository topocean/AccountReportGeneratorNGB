Public Class RptAirLoading

    Function RptAirLoading(ByVal UID As String, ByVal ds As DataSet) As String
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
                    fileName = common.NullVal(.Item("RptFile").ToString, UID)
                    SubBrhCd = common.NullVal(.Item("HdrBranch").ToString, "")

                    ' ----------------------------------------------------------------------
                    ' Report Header (Company Name, Address, Tel, etc...)
                    ' ----------------------------------------------------------------------

                    objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                    objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                    objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                    objWS.Application.Cells(5, 1) = "LOADING REPORT - " & SubBrhCd & " (YEAR: " & common.NullVal(.Item("HdrYear").ToString, "") & " , WEEK: " & common.NullVal(.Item("HdrWeek").ToString, "") & ")"

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
                objWS.Application.Cells(iRow, iCol + 3) = "BY"
                objWS.Application.Cells(iRow, iCol + 4) = "AGENT"
                objWS.Application.Cells(iRow, iCol + 5) = "SHIPPER"
                objWS.Application.Cells(iRow, iCol + 6) = "CONSIGNEE"
                objWS.Application.Cells(iRow, iCol + 7) = "PIECE COUNT"
                objWS.Application.Cells(iRow, iCol + 8) = "GROSS WEIGHT"
                objWS.Application.Cells(iRow, iCol + 9) = "VOLUME WEIGHT"
                objWS.Application.Cells(iRow, iCol + 10) = "CBM"
                objWS.Application.Cells(iRow, iCol + 11) = "AIRLINE"
                objWS.Application.Cells(iRow, iCol + 12) = "FLIGHT NO"
                objWS.Application.Cells(iRow, iCol + 13) = "ETD"
                objWS.Application.Cells(iRow, iCol + 14) = "ETA"
                objWS.Application.Cells(iRow, iCol + 15) = "ORIGIN"
                objWS.Application.Cells(iRow, iCol + 16) = "DEST"
                objWS.Application.Cells(iRow, iCol + 17) = "HOUSE AIRWAY BILL #"
                objWS.Application.Cells(iRow, iCol + 18) = "MASTER AIRWAY BILL #"
                objWS.Application.Cells(iRow, iCol + 19) = "LOT #"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 19)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 19)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 19)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 19)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 19)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 19)).Borders(11).LineStyle = 1

                ' **********************************************************************

                iRow += 1
                iSRow = iRow

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    With ds.Tables(1).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("AbhMonth"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("AbhWeek"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("AbhSubBrhCd"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("AbhLstUsr"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = "'" & common.NullVal(.Item("AgtCd"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = common.NullVal(.Item("ShpName"), "")
                        objWS.Application.Cells(iRow, iCol + 6) = common.NullVal(.Item("ConName"), "")

                        If common.NullVal(.Item("AbhPCS"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 7) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 7) = common.NullVal(.Item("AbhPCS"), 0)
                        End If

                        If common.NullVal(.Item("AbhGW"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 8) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 8) = common.NullVal(.Item("AbhGW"), 0)
                        End If

                        If common.NullVal(.Item("AbhVW"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 9) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 9) = common.NullVal(.Item("AbhVW"), 0)
                        End If

                        If common.NullVal(.Item("AbhCBM"), 0) = 0 Then
                            objWS.Application.Cells(iRow, iCol + 10) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + 10) = common.NullVal(.Item("AbhCBM"), 0)
                        End If

                        objWS.Application.Cells(iRow, iCol + 11) = common.NullVal(.Item("ClientAirlineCode"), 0)  ' Consol Box
                        objWS.Application.Cells(iRow, iCol + 12) = common.NullVal(.Item("FgtNo"), 0)
                        objWS.Application.Cells(iRow, iCol + 13) = "'" & .Item("AbhETD")
                        objWS.Application.Cells(iRow, iCol + 14) = "'" & .Item("AbhETA")
                        objWS.Application.Cells(iRow, iCol + 15) = common.NullVal(.Item("AbhLoad"), "")
                        objWS.Application.Cells(iRow, iCol + 16) = common.NullVal(.Item("AbhDest"), "")
                        objWS.Application.Cells(iRow, iCol + 17) = common.NullVal(.Item("AbhHwabNo"), "")
                        objWS.Application.Cells(iRow, iCol + 18) = "'" & common.NullVal(.Item("AbhMawbNo"), "")
                        objWS.Application.Cells(iRow, iCol + 19) = common.NullVal(.Item("AbhLotNo"), "")

                    End With

                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1
                Next

                ' ----------------------------------------------------------------------
                ' Calculate the total number of containers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 6) = "TOTAL: "

                For i = 8 To 11
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                Next

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Underline the total columns)
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 7), objWS.Application.Cells(iRow, iCol + 10)).Borders(9).LineStyle = -4119

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Column Width)
                ' ----------------------------------------------------------------------

                objWS.Application.Range("A8:B8").ColumnWidth = 8
                objWS.Application.Range("C8:D8").ColumnWidth = 14
                objWS.Application.Range("E8:D8").ColumnWidth = 10
                objWS.Application.Range("E8:E8").ColumnWidth = 15
                objWS.Application.Range("F8:G8").ColumnWidth = 40
                objWS.Application.Range("H8:K8").ColumnWidth = 12.5
                objWS.Application.Range("L8:M8").ColumnWidth = 10
                objWS.Application.Range("N8:O8").ColumnWidth = 13
                objWS.Application.Range("P8:Q8").ColumnWidth = 20
                objWS.Application.Range("R8:T8").ColumnWidth = 15

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
        RptAirLoading = fileName
    End Function
End Class
