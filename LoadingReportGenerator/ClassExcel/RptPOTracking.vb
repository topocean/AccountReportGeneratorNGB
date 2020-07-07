Public Class RptPOTracking

    Function RptPOTracking(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, startRow As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim hasData As Boolean = False

        If ds.Tables(0).Rows.Count > 0 Then
            fileName = ds.Tables(0).Rows(0).Item("RptFile").ToString

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

                iRow = 11
                iCol = 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Export Report Data
                ' ----------------------------------------------------------------------

                If ds.Tables(1).Rows.Count > 0 Then
                    hasData = True

                    objExcel.Cells(iRow, iCol) = "SEQUENCE"
                    objExcel.Cells(iRow, iCol + 1) = "ORG"
                    objExcel.Cells(iRow, iCol + 2) = "DEST"
                    objExcel.Cells(iRow, iCol + 3) = "CNEE"
                    objExcel.Cells(iRow, iCol + 4) = "VENDOR"

                    objExcel.Cells(iRow, iCol + 5) = "PO#"
                    objExcel.Cells(iRow, iCol + 6) = "Alt. PO#"
                    objExcel.Cells(iRow, iCol + 7) = "SKU/Part#"
                    objExcel.Cells(iRow, iCol + 8) = "INVOICE#"
                    objExcel.Cells(iRow, iCol + 9) = "DESCRIPTION"

                    objExcel.Cells(iRow, iCol + 10) = "UNITS ORDER"
                    objExcel.Cells(iRow, iCol + 11) = "CTNS ORDER"
                    objExcel.Cells(iRow, iCol + 12) = "FOB POINT"
                    objExcel.Cells(iRow, iCol + 13) = "Schedule Ship Date"
                    objExcel.Cells(iRow, iCol + 14) = "First Ship Date"

                    objExcel.Cells(iRow, iCol + 15) = "Last Ship Date"
                    objExcel.Cells(iRow, iCol + 16) = "PO Ready Date"
                    objExcel.Cells(iRow, iCol + 17) = "Booking Date"
                    objExcel.Cells(iRow, iCol + 18) = "Ship From Factory Date"
                    objExcel.Cells(iRow, iCol + 19) = "Cargo RECEIVEd DATE"

                    objExcel.Cells(iRow, iCol + 20) = "LOADING PORT"
                    objExcel.Cells(iRow, iCol + 21) = "POL ETD"
                    objExcel.Cells(iRow, iCol + 22) = "ATD"
                    objExcel.Cells(iRow, iCol + 23) = "DISCHARGE PORT"
                    objExcel.Cells(iRow, iCol + 24) = "POD ETA"

                    objExcel.Cells(iRow, iCol + 25) = "IPI RAMP"
                    objExcel.Cells(iRow, iCol + 26) = "FINAL DEST"
                    objExcel.Cells(iRow, iCol + 27) = "UNITS SHIP"
                    objExcel.Cells(iRow, iCol + 28) = "CTNS SHIP"
                    objExcel.Cells(iRow, iCol + 29) = "WEIGHT KGS"

                    objExcel.Cells(iRow, iCol + 30) = "CBM"
                    objExcel.Cells(iRow, iCol + 31) = "MANIFEST SEQ"
                    objExcel.Cells(iRow, iCol + 32) = "VAN POSITION"
                    objExcel.Cells(iRow, iCol + 33) = "CARRIER"
                    objExcel.Cells(iRow, iCol + 34) = "VSL"

                    objExcel.Cells(iRow, iCol + 35) = "VOY"
                    objExcel.Cells(iRow, iCol + 36) = "SERVICE MODE"
                    objExcel.Cells(iRow, iCol + 37) = "CNTR TYPE"
                    objExcel.Cells(iRow, iCol + 38) = "CNTR NUMBER"
                    objExcel.Cells(iRow, iCol + 39) = "SEAL NUMBER"

                    objExcel.Cells(iRow, iCol + 40) = "MB/L"
                    objExcel.Cells(iRow, iCol + 41) = "FCR"
                    objExcel.Cells(iRow, iCol + 42) = "HB/L"
                    objExcel.Cells(iRow, iCol + 43) = "OCEAN FRT"
                    objExcel.Cells(iRow, iCol + 44) = "PSS / EBS"

                    objExcel.Cells(iRow, iCol + 45) = "ORC THC"
                    objExcel.Cells(iRow, iCol + 46) = "DOX FEE"
                    objExcel.Cells(iRow, iCol + 47) = "HULAGE"
                    objExcel.Cells(iRow, iCol + 48) = "EXPORT CUSTOM"
                    objExcel.Cells(iRow, iCol + 49) = "HANDLING CHARGE"

                    objExcel.Cells(iRow, iCol + 50) = "OTHERS"
                    objExcel.Cells(iRow, iCol + 51) = "REMARKS"
                    objExcel.Cells(iRow, iCol + 52) = "Model"
                    objExcel.Cells(iRow, iCol + 53) = "Export#"
                    objExcel.Cells(iRow, iCol + 54) = "Insurance Fee"

                    objExcel.Cells(iRow, iCol + 55) = "Unit Price"
                    objExcel.Cells(iRow, iCol + 56) = "LCL IPI"
                    objExcel.Cells(iRow, iCol + 57) = "Invoice Amt"
                    objExcel.Cells(iRow, iCol + 58) = "INVOICE DATE"
                    objExcel.Cells(iRow, iCol + 59) = "OrgDocRcvd"

                    objExcel.Cells(iRow, iCol + 60) = "BuyingAgent"
                    objExcel.Cells(iRow, iCol + 61) = "Sub-CNEE"
                    objExcel.Cells(iRow, iCol + 62) = "Customs Broker"
                    objExcel.Cells(iRow, iCol + 63) = "Tranload Facility"
                    objExcel.Cells(iRow, iCol + 64) = "Final ETA"

                    objExcel.Cells(iRow, iCol + 65) = "PO Rcvd Date"

                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 65)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 65)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 65)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 65)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 65)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 65)).Borders(11).LineStyle = 1

                    iRow = iRow + 1
                    startRow = iRow

                    'Set Content
                    For i = 0 To ds.Tables(1).Rows.Count - 1
                        With ds.Tables(1).Rows(i)
                            objExcel.Cells(iRow, iCol) = .Item("ConSeqNo").ToString
                            objExcel.Cells(iRow, iCol + 1) = .Item("orgName").ToString
                            objExcel.Cells(iRow, iCol + 2) = "'" & .Item("destName").ToString
                            objExcel.Cells(iRow, iCol + 3) = .Item("conName").ToString

                            If Len(.Item("ShpName").ToString) > 20 Then
                                objExcel.Cells(iRow, iCol + 4) = Mid(.Item("ShpName").ToString, 1, 20)
                            Else
                                objExcel.Cells(iRow, iCol + 4) = .Item("ShpName").ToString
                            End If

                            objExcel.Cells(iRow, iCol + 5) = .Item("BktPO").ToString
                            objExcel.Cells(iRow, iCol + 6) = ""
                            objExcel.Cells(iRow, iCol + 7) = .Item("BktSkuNo").ToString
                            objExcel.Cells(iRow, iCol + 8) = .Item("BktInvNo").ToString
                            objExcel.Cells(iRow, iCol + 9) = .Item("BkhCarrComm").ToString

                            If common.NullVal(.Item("BkhCtnr").ToString, "") = "" Then
                                objExcel.Cells(iRow, iCol + 10) = " "
                            Else
                                objExcel.Cells(iRow, iCol + 10) = CDbl(.Item("BkhCtnr").ToString)
                            End If

                            objExcel.Cells(iRow, iCol + 11) = CDbl(.Item("UntName").ToString)
                            objExcel.Cells(iRow, iCol + 12) = .Item("VslLoadName").ToString
                            objExcel.Cells(iRow, iCol + 13) = .Item("shpDte").ToString

                            If common.NullVal(.Item("FirstShpDte").ToString, "") = "" Then
                                objExcel.Cells(iRow, iCol + 14) = " "
                            Else
                                objExcel.Cells(iRow, iCol + 14) = "'" & .Item("FirstShpDte").ToString
                            End If

                            If common.NullVal(.Item("LastShpDte").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 15) = "'" & .Item("LastShpDte").ToString
                            End If

                            If common.NullVal(.Item("POReadyDte").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 16) = "'" & .Item("POReadyDte").ToString
                            End If

                            If common.NullVal(.Item("BookingDte").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 17) = "'" & .Item("BookingDte").ToString
                            End If

                            If common.NullVal(.Item("ShipFrmFactDte").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 18) = "'" & .Item("ShipFrmFactDte").ToString
                            End If

                            If common.NullVal(.Item("BkhCargoRec").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 19) = "'" & .Item("BkhCargoRec").ToString
                            End If

                            objExcel.Cells(iRow, iCol + 20) = .Item("LoadName").ToString

                            If common.NullVal(.Item("PolETD").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 21) = "'" & .Item("PolETD").ToString
                            End If

                            If common.NullVal(.Item("ATD").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 22) = "'" & .Item("ATD").ToString
                            End If

                            objExcel.Cells(iRow, iCol + 23) = "'" & .Item("DisName").ToString

                            If common.NullVal(.Item("PodETA").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 24) = "'" & .Item("PodETA").ToString
                            End If

                            objExcel.Cells(iRow, iCol + 25) = "'" & .Item("BktRamp").ToString
                            objExcel.Cells(iRow, iCol + 26) = "'" & .Item("FDestName").ToString

                            If common.NullVal(.Item("CtsName").ToString, 0) = 0 Then
                                objExcel.Cells(iRow, iCol + 27) = ""
                            Else
                                objExcel.Cells(iRow, iCol + 27) = CDbl(.Item("CtsName").ToString)
                            End If

                            objExcel.Cells(iRow, iCol + 28) = CDbl(.Item("CtnUntName").ToString)
                            objExcel.Cells(iRow, iCol + 29) = CDbl(.Item("BktWgt").ToString)

                            objExcel.Cells(iRow, iCol + 30) = CDbl(.Item("BktCBM").ToString)
                            objExcel.Cells(iRow, iCol + 31) = ""
                            objExcel.Cells(iRow, iCol + 32) = ""
                            objExcel.Cells(iRow, iCol + 33) = .Item("CarName").ToString
                            objExcel.Cells(iRow, iCol + 34) = .Item("VslName").ToString

                            objExcel.Cells(iRow, iCol + 35) = .Item("VslVoyName").ToString
                            objExcel.Cells(iRow, iCol + 36) = .Item("BktMode").ToString
                            objExcel.Cells(iRow, iCol + 37) = .Item("cntrType").ToString
                            objExcel.Cells(iRow, iCol + 38) = .Item("BktCtrNo").ToString
                            objExcel.Cells(iRow, iCol + 39) = .Item("BktSeal").ToString

                            objExcel.Cells(iRow, iCol + 40) = .Item("BkhMBLNo").ToString
                            objExcel.Cells(iRow, iCol + 41) = ""
                            objExcel.Cells(iRow, iCol + 42) = .Item("BkhBLNo").ToString

                            If common.NullVal(.Item("OceanFrt").ToString, "") <> "" Then
                                objExcel.Cells(iRow, iCol + 43) = CDbl(.Item("OceanFrt").ToString)
                            End If

                            objExcel.Cells(iRow, iCol + 44) = ""

                            objExcel.Cells(iRow, iCol + 45) = ""
                            objExcel.Cells(iRow, iCol + 46) = ""
                            objExcel.Cells(iRow, iCol + 47) = ""
                            objExcel.Cells(iRow, iCol + 48) = ""
                            objExcel.Cells(iRow, iCol + 49) = ""

                            objExcel.Cells(iRow, iCol + 50) = ""
                            objExcel.Cells(iRow, iCol + 51) = ""
                            objExcel.Cells(iRow, iCol + 52) = .Item("BkpModel").ToString
                            objExcel.Cells(iRow, iCol + 53) = .Item("BkpExport").ToString
                            objExcel.Cells(iRow, iCol + 54) = .Item("BkpInsFee").ToString

                            objExcel.Cells(iRow, iCol + 55) = .Item("BkpUnitPrice").ToString
                            objExcel.Cells(iRow, iCol + 56) = .Item("BkpLCLIPI").ToString
                            objExcel.Cells(iRow, iCol + 57) = .Item("BktIssueDte").ToString
                            objExcel.Cells(iRow, iCol + 58) = ""
                            objExcel.Cells(iRow, iCol + 59) = ""

                            objExcel.Cells(iRow, iCol + 60) = ""
                            objExcel.Cells(iRow, iCol + 61) = .Item("SubCNEE").ToString
                            objExcel.Cells(iRow, iCol + 62) = .Item("CustBroker").ToString
                            objExcel.Cells(iRow, iCol + 63) = .Item("TranloadFaci").ToString
                            objExcel.Cells(iRow, iCol + 64) = .Item("FinalETA").ToString

                            objExcel.Cells(iRow, iCol + 65) = .Item("PORcvdDate").ToString
                        End With

                        iRow = iRow + 1

                    Next

                    objExcel.Columns("A:A").ColumnWidth = 10
                    objExcel.Columns("B:C").ColumnWidth = 7
                    objExcel.Columns("D:BN").ColumnWidth = 20
                End If

                If Not hasData Then
                    fileName = ""
                End If

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
        End If

        RptPOTracking = fileName

    End Function
End Class
