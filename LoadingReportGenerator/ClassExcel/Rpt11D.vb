Public Class Rpt11D

    Function Rpt11D(ByVal UID As String, ByVal ds As DataSet) As String
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

                iRow = 1
                iCol = 1

                ' **********************************************************************

                If ds.Tables(1).Rows.Count > 0 Then
                    hasData = True

                    objExcel.Cells(iRow, iCol) = "SEQUENCE"
                    objExcel.Cells(iRow, iCol + 1) = "ORG"
                    objExcel.Cells(iRow, iCol + 2) = "DEST"
                    objExcel.Cells(iRow, iCol + 3) = "CNEE"
                    objExcel.Cells(iRow, iCol + 4) = "VENDOR"

                    objExcel.Cells(iRow, iCol + 5) = "PO#"
                    objExcel.Cells(iRow, iCol + 6) = "SKU#"
                    objExcel.Cells(iRow, iCol + 7) = "DESCRIPTION"
                    objExcel.Cells(iRow, iCol + 8) = "BOOKING DATE"
                    objExcel.Cells(iRow, iCol + 9) = "LOADING PORT"

                    objExcel.Cells(iRow, iCol + 10) = "DISCHARGE PORT"
                    objExcel.Cells(iRow, iCol + 11) = "IPI RAMP"
                    objExcel.Cells(iRow, iCol + 12) = "FINAL DEST"
                    objExcel.Cells(iRow, iCol + 13) = "UNITS SHIP"
                    objExcel.Cells(iRow, iCol + 14) = "CTNS SHIP"

                    objExcel.Cells(iRow, iCol + 15) = "CARRIER"
                    objExcel.Cells(iRow, iCol + 16) = "VSL"
                    objExcel.Cells(iRow, iCol + 17) = "VOY"
                    objExcel.Cells(iRow, iCol + 18) = "CNTR TYPE"
                    objExcel.Cells(iRow, iCol + 19) = "CNTR NUMBER"

                    objExcel.Cells(iRow, iCol + 20) = "MB/L"
                    objExcel.Cells(iRow, iCol + 21) = "HB/L"
                    objExcel.Cells(iRow, iCol + 22) = "POL ETD"
                    objExcel.Cells(iRow, iCol + 23) = "POL ATD"
                    objExcel.Cells(iRow, iCol + 24) = "POD ETA"

                    objExcel.Cells(iRow, iCol + 25) = "Ramp ETA"
                    objExcel.Cells(iRow, iCol + 26) = "REMARKS"
                    objExcel.Cells(iRow, iCol + 27) = "FINALETA"

                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Interior.ColorIndex = 15
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Font.Bold = True
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 27)).Borders(11).LineStyle = 1

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
                            objExcel.Cells(iRow, iCol + 6) = .Item("BktSkuNo").ToString
                            objExcel.Cells(iRow, iCol + 7) = .Item("BkhCarrComm").ToString
                            objExcel.Cells(iRow, iCol + 8) = "'" & .Item("BkhBkgDte").ToString
                            objExcel.Cells(iRow, iCol + 9) = .Item("LoadName").ToString

                            objExcel.Cells(iRow, iCol + 10) = .Item("DisName").ToString
                            objExcel.Cells(iRow, iCol + 11) = .Item("IPIRamp").ToString
                            objExcel.Cells(iRow, iCol + 12) = .Item("FDestName").ToString
                            objExcel.Cells(iRow, iCol + 13) = .Item("BkhCtnr").ToString
                            objExcel.Cells(iRow, iCol + 14) = .Item("CtnUntName").ToString

                            objExcel.Cells(iRow, iCol + 15) = .Item("CarName").ToString
                            objExcel.Cells(iRow, iCol + 16) = .Item("VslName").ToString
                            objExcel.Cells(iRow, iCol + 17) = .Item("VslVoyName").ToString
                            objExcel.Cells(iRow, iCol + 18) = .Item("cntrType").ToString
                            objExcel.Cells(iRow, iCol + 19) = .Item("BktCtrNo").ToString

                            objExcel.Cells(iRow, iCol + 20) = .Item("BkhMBLNo").ToString
                            objExcel.Cells(iRow, iCol + 21) = .Item("BkhBLNo").ToString
                            objExcel.Cells(iRow, iCol + 22) = "'" & .Item("PolETD").ToString

                            If common.NullVal(.Item("PolATD").ToString, "") = "" Then
                                objExcel.Cells(iRow, iCol + 23) = "'"
                            Else
                                objExcel.Cells(iRow, iCol + 23) = "'" & .Item("PolATD").ToString
                            End If
                            objExcel.Cells(iRow, iCol + 24) = "'" & .Item("PodETA").ToString

                            If common.NullVal(.Item("RampETA").ToString, "") = "" Then
                                objExcel.Cells(iRow, iCol + 25) = "'"
                            Else
                                objExcel.Cells(iRow, iCol + 25) = "'" & .Item("RampETA").ToString
                            End If
                            objExcel.Cells(iRow, iCol + 26) = ""
                            If common.NullVal(.Item("FINALETA").ToString, "") = "" Then
                                objExcel.Cells(iRow, iCol + 27) = "'"
                            Else
                                objExcel.Cells(iRow, iCol + 27) = "'" & .Item("FINALETA").ToString
                            End If
                        End With

                        iRow = iRow + 1

                    Next

                    objExcel.Columns("A:A").ColumnWidth = 12
                    objExcel.Columns("B:C").ColumnWidth = 7
                    objExcel.Columns("D:AB").ColumnWidth = 17

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
                End If

            Catch ex As Exception
                objExcel.ActiveWorkbook.SaveAs("C:\" & UID & ".xls")
                objExcel.Quit()
                fileName = "Error," & ex.Message
            End Try
        End If

        Rpt11D = fileName
    End Function
End Class
