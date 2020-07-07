Public Class RptLoading_Monthly

    Function RptLoading_Monthly(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i As Integer
        Dim iRow, iSRow, iCol, addi As Integer
        Dim fileName As String
        Dim SubBrhCd, traffic, location As Integer
        Dim TrafficName As String = ""
        Dim LocName As String = ""
        Dim OCF As String
        Dim common As New common

        Dim brhCd As Integer

        addi = 0

        If ds.Tables(2).Rows.Count > 0 Then
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
                ' Get File Name
                ' ----------------------------------------------------------------------

                fileName = common.NullVal(ds.Tables(0).Rows(0).Item("fName").ToString, UID)

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Get Sub-Branch, Traffic and Location
                ' ----------------------------------------------------------------------

                With ds.Tables(1).Rows(0)
                    SubBrhCd = CInt(common.NullVal(.Item("SubBrhCd").ToString, 0))
                    traffic = CInt(common.NullVal(.Item("Traffic").ToString, 0))
                    location = CInt(common.NullVal(.Item("Loc").ToString, 0))

                    brhCd = CInt(common.NullVal(.Item("BrhCd").ToString, 0))

                    If common.NullVal(.Item("OCF").ToString, "") = "" Then
                        OCF = "ALL"
                    Else
                        OCF = common.NullVal(.Item("OCF").ToString, "")
                    End If
                End With

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Location
                ' ----------------------------------------------------------------------

                Select Case CInt(location)
                    Case 0
                        LocName = "ALL"
                    Case 1
                        LocName = "HONG KONG"
                    Case 2
                        LocName = "SHENZHEN"
                End Select

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Traffic
                ' ----------------------------------------------------------------------

                Select Case CInt(traffic)
                    Case 0
                        TrafficName = "ALL"
                    Case 1
                        TrafficName = "NON-USA"
                    Case 2
                        TrafficName = "USA"
                    Case 3
                        TrafficName = "EUR"
                    Case 4
                        TrafficName = "CHN"
                    Case 5
                        TrafficName = "SEA"
                    Case 6
                        TrafficName = "NEA"
                    Case 7
                        TrafficName = "AUS"
                    Case 8
                        TrafficName = "CAN"
                    Case 9
                        TrafficName = "SAM"
                    Case 10
                        TrafficName = "ISC"
                    Case 11
                        TrafficName = "AFR"
                End Select

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Export Report Header (Company Name, Address, Tel, etc...)
                ' ----------------------------------------------------------------------

                With ds.Tables(1).Rows(0)
                    objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                    objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                    objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                    objWS.Application.Cells(5, 1) = "LOADING REPORT - " & common.NullVal(.Item("SubBrh"), "") & " (YEAR: " & common.NullVal(.Item("HdrYear").ToString, "") & " , MONTH " & common.NullVal(.Item("HdrMonth").ToString, "") & ")"
                    objWS.Application.Cells(6, 1) = "TRAFFIC: " & TrafficName & " , LOCATION: " & LocName & " , OCF: " & OCF

                    ' ----------------------------------------------------------------------
                    ' Setting Properties (Bold Header Details and Merge Cells)
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range("A1:I6").Font.Bold = True
                    objWS.Application.Range("A1:I6").HorizontalAlignment = -4108
                    objWS.Application.Range("A1:I1").Merge()
                    objWS.Application.Range("A2:I2").Merge()
                    objWS.Application.Range("A3:I3").Merge()
                    objWS.Application.Range("A5:I5").Merge()
                    objWS.Application.Range("A6:I6").Merge()

                    ' **********************************************************************
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

                If TrafficName <> "USA" And brhCd <> 9 And brhCd <> 59 Then
                    addi = 1
                    objWS.Application.Cells(iRow, iCol + 5 + addi) = "AGENT NAME"
                End If

                objWS.Application.Cells(iRow, iCol + 6 + addi) = "SHIPPER"
                objWS.Application.Cells(iRow, iCol + 7 + addi) = "CONSIGNEE"
                objWS.Application.Cells(iRow, iCol + 8 + addi) = "NOMINATION"
                objWS.Application.Cells(iRow, iCol + 9 + addi) = "NOMINATION SALES"
                objWS.Application.Cells(iRow, iCol + 10 + addi) = "'20"
                objWS.Application.Cells(iRow, iCol + 11 + addi) = "'40"
                objWS.Application.Cells(iRow, iCol + 12 + addi) = "'45"
                objWS.Application.Cells(iRow, iCol + 13 + addi) = "HQ"
                objWS.Application.Cells(iRow, iCol + 14 + addi) = "FEUS"
                objWS.Application.Cells(iRow, iCol + 15 + addi) = "CONSOL"
                objWS.Application.Cells(iRow, iCol + 16 + addi) = "CBM"
                objWS.Application.Cells(iRow, iCol + 17 + addi) = "TYPE"
                objWS.Application.Cells(iRow, iCol + 18 + addi) = "CARRIER"
                objWS.Application.Cells(iRow, iCol + 19 + addi) = "CONTRACT #"
                objWS.Application.Cells(iRow, iCol + 20 + addi) = "SERVICE STRING"
                objWS.Application.Cells(iRow, iCol + 21 + addi) = "SO NO"
                objWS.Application.Cells(iRow, iCol + 22 + addi) = "VESSEL"
                objWS.Application.Cells(iRow, iCol + 23 + addi) = "VOYAGE"
                objWS.Application.Cells(iRow, iCol + 24 + addi) = "ETD"
                objWS.Application.Cells(iRow, iCol + 25 + addi) = "ETA"
                objWS.Application.Cells(iRow, iCol + 26 + addi) = "POL"
                objWS.Application.Cells(iRow, iCol + 27 + addi) = "DEST"
                objWS.Application.Cells(iRow, iCol + 28 + addi) = "MBL BOOK TO"
                objWS.Application.Cells(iRow, iCol + 29 + addi) = "HOUSE BL #"
                objWS.Application.Cells(iRow, iCol + 30 + addi) = "MASTER BL #"
                objWS.Application.Cells(iRow, iCol + 31 + addi) = "CONTAINER #"
                objWS.Application.Cells(iRow, iCol + 32 + addi) = "STATUS"
                objWS.Application.Cells(iRow, iCol + 33 + addi) = "SEND PRE-ALERT"
                objWS.Application.Cells(iRow, iCol + 34 + addi) = "UPDATE"
                objWS.Application.Cells(iRow, iCol + 35 + addi) = "LOT #"
                objWS.Application.Cells(iRow, iCol + 36 + addi) = "LOCAL SALES"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 36 + addi)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 36 + addi)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 36 + addi)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 36 + addi)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 36 + addi)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + 36 + addi)).Borders(11).LineStyle = 1

                ' **********************************************************************

                iRow += 1
                iSRow = iRow

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------

                For i = 0 To ds.Tables(2).Rows.Count - 1
                    With ds.Tables(2).Rows(i)
                        objWS.Application.Cells(iRow, iCol) = common.NullVal(.Item("BkhMonth"), "")
                        objWS.Application.Cells(iRow, iCol + 1) = common.NullVal(.Item("BkhWeek"), "")
                        objWS.Application.Cells(iRow, iCol + 2) = common.NullVal(.Item("SubBrh"), "")
                        objWS.Application.Cells(iRow, iCol + 3) = common.NullVal(.Item("BkhTraffic"), "")
                        objWS.Application.Cells(iRow, iCol + 4) = common.NullVal(.Item("BkhUsrName"), "")
                        objWS.Application.Cells(iRow, iCol + 5) = "'" & common.NullVal(.Item("BkhAgtName"), "")

                        If TrafficName <> "USA" And brhCd <> 9 And brhCd <> 59 Then
                            objWS.Application.Cells(iRow, iCol + 5 + addi) = "'" & common.NullVal(.Item("AgtName"), "")
                        End If

                        objWS.Application.Cells(iRow, iCol + 6 + addi) = common.NullVal(.Item("ShpName"), "")
                        objWS.Application.Cells(iRow, iCol + 7 + addi) = common.NullVal(.Item("ConName"), "")
                        objWS.Application.Cells(iRow, iCol + 8 + addi) = common.NullVal(.Item("NomSName"), "")
                        objWS.Application.Cells(iRow, iCol + 9 + addi) = common.NullVal(.Item("NomSales"), "")
                        objWS.Application.Cells(iRow, iCol + 10 + addi) = common.NullVal(.Item("Size1"), "")    ' 20
                        objWS.Application.Cells(iRow, iCol + 11 + addi) = common.NullVal(.Item("Size2"), "")   ' 40
                        objWS.Application.Cells(iRow, iCol + 12 + addi) = common.NullVal(.Item("Size3"), "")   ' 45
                        objWS.Application.Cells(iRow, iCol + 13 + addi) = common.NullVal(.Item("Size4"), "")   ' HQ
                        objWS.Application.Cells(iRow, iCol + 14 + addi) = common.NullVal(.Item("Consol"), "")  ' Consol Box
                        objWS.Application.Cells(iRow, iCol + 16 + addi) = common.NullVal(.Item("LoadCBM"), 0)
                        objWS.Application.Cells(iRow, iCol + 17 + addi) = common.NullVal(.Item("TypeRemark"), "")
                        objWS.Application.Cells(iRow, iCol + 18 + addi) = common.NullVal(.Item("CarSName"), "")
                        objWS.Application.Cells(iRow, iCol + 19 + addi) = common.NullVal(.Item("ContractNo"), "")
                        objWS.Application.Cells(iRow, iCol + 20 + addi) = common.NullVal(.Item("ServiceStr"), "")
                        objWS.Application.Cells(iRow, iCol + 21 + addi) = "'" & common.NullVal(.Item("SONo"), "")
                        objWS.Application.Cells(iRow, iCol + 22 + addi) = common.NullVal(.Item("VslName"), "")
                        objWS.Application.Cells(iRow, iCol + 23 + addi) = "'" & common.NullVal(.Item("VslVoyName"), "")

                        ' ----------------------------------------------------------------------
                        ' ETD
                        ' ----------------------------------------------------------------------
                        If common.NullVal(.Item("VslETD").ToString, "") = "" Then
                            objWS.Application.Cells(iRow, iCol + 24 + addi) = ""
                        Else
                            If Format(CDate(.Item("VslETD").ToString), "yyyy/MM/dd") = "1900/01/01" Then
                                objWS.Application.Cells(iRow, iCol + 24 + addi) = ""
                            Else
                                objWS.Application.Cells(iRow, iCol + 24 + addi) = Format(CDate(.Item("VslETD").ToString), "MMM dd, yyyy")
                            End If
                        End If

                        ' **********************************************************************

                        ' ----------------------------------------------------------------------
                        ' ETA
                        ' ----------------------------------------------------------------------
                        If common.NullVal(.Item("BkhETA").ToString, "") = "" Then
                            objWS.Application.Cells(iRow, iCol + 25 + addi) = ""
                        Else
                            If Format(CDate(.Item("BkhETA").ToString), "yyyy/MM/dd") = "1900/01/01" Then
                                objWS.Application.Cells(iRow, iCol + 25 + addi) = ""
                            Else
                                objWS.Application.Cells(iRow, iCol + 25 + addi) = Format(CDate(.Item("BkhETA").ToString), "MMM dd, yyyy")
                            End If
                        End If

                        ' **********************************************************************

                        objWS.Application.Cells(iRow, iCol + 26 + addi) = common.NullVal(.Item("BkhLoadName"), "")
                        objWS.Application.Cells(iRow, iCol + 27 + addi) = common.NullVal(.Item("BkhDestName"), "")
                        objWS.Application.Cells(iRow, iCol + 28 + addi) = common.NullVal(.Item("BkhDiscName"), "")
                        objWS.Application.Cells(iRow, iCol + 29 + addi) = common.NullVal(.Item("BLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 30 + addi) = "'" & common.NullVal(.Item("BkhMBLNo"), "")
                        objWS.Application.Cells(iRow, iCol + 31 + addi) = common.NullVal(.Item("CtnrNo"), "")
                        objWS.Application.Cells(iRow, iCol + 32 + addi) = common.NullVal(.Item("TLX"), "")

                        objWS.Application.Cells(iRow, iCol + 33 + addi) = common.NullVal(.Item("BkhPreAlert"), "")

                        ' ----------------------------------------------------------------------
                        ' TLX Update date
                        ' ----------------------------------------------------------------------

                        objWS.Application.Cells(iRow, iCol + 34 + addi) = common.NullVal(.Item("TLXDte"), "")

                        If common.NullVal(.Item("TLXDte"), "") = Format(Now, "MMM dd, yyyy") Then
                            objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 34 + addi), objWS.Application.Cells(iRow, iCol + 31 + addi)).Interior.Color = RGB(229, 229, 231)
                        End If

                        ' **********************************************************************

                        objWS.Application.Cells(iRow, iCol + 35 + addi) = "'" & common.NullVal(.Item("BkhLotNo"), "")
                        objWS.Application.Cells(iRow, iCol + 36 + addi) = common.NullVal(.Item("NomLSman"), "")
                    End With

                    ' ----------------------------------------------------------------------
                    ' FEUS Calculation
                    ' ----------------------------------------------------------------------

                    objWS.Application.Range(objWS.Application.Cells(iRow, 11 + addi), objWS.Application.Cells(iRow, 14 + addi)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, 15 + addi), objWS.Application.Cells(iRow, 15 + addi)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1],RC[-4]/2,RC[1]*-1)"

                    ' ----------------------------------------------------------------------
                    ' New Line
                    ' ----------------------------------------------------------------------

                    iRow += 1
                Next

                ' ----------------------------------------------------------------------
                ' Calculate the total number of containers
                ' ----------------------------------------------------------------------

                objWS.Application.Cells(iRow, iCol + 9 + addi) = "TOTAL: "

                For i = (11 + addi) To (17 + addi)
                    objWS.Application.Range(objWS.Application.Cells(iSRow, i), objWS.Application.Cells(iRow - 1, i)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, i), objWS.Application.Cells(iRow, i)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                Next

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties (Underline the total columns)
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, iCol + 10 + addi), objWS.Application.Cells(iRow, iCol + 16 + addi)).Borders(9).LineStyle = -4119

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
                objWS.Application.Range("I8:J8").ColumnWidth = 15
                objWS.Application.Range("K8:P8").ColumnWidth = 8
                objWS.Application.Range("Q8:T8").ColumnWidth = 10
                objWS.Application.Range("U8:W8").ColumnWidth = 15
                objWS.Application.Range("X8:X8").ColumnWidth = 10
                objWS.Application.Range("Y8:Z8").ColumnWidth = 15
                objWS.Application.Range("AA8:AF8").ColumnWidth = 20
                objWS.Application.Range("AG8:AG8").ColumnWidth = 10
                objWS.Application.Range("AH8:AI8").ColumnWidth = 15
                objWS.Application.Range("AJ8:AJ8").ColumnWidth = 30

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

                ' Create if export directory not found
                If Not My.Computer.FileSystem.DirectoryExists(exportPath) Then
                    My.Computer.FileSystem.CreateDirectory(exportPath)
                End If

                ' Delete if file already exists
                If My.Computer.FileSystem.FileExists(exportFile) Then
                    My.Computer.FileSystem.DeleteFile(exportFile)
                End If

                fileName &= ".xls"
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
                location = Nothing
                OCF = Nothing
                TrafficName = Nothing
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
                fileName = "error," & ex.Message
            End Try
        Else
            fileName = ""
        End If

        ' Return File Path
        RptLoading_Monthly = fileName
    End Function
End Class
