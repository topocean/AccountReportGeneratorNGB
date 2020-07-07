Public Class RptBSShippingAdvice

    Function RptBSShippingAdvice(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim BkhMode, ModeStr As String
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
            ' Get File Name, Report Header
            ' ----------------------------------------------------------------------

            With ds.Tables(0).Rows(0)
                fileName = .Item("RptFile").ToString
                BkhMode = CInt(.Item("Mode").ToString)
            End With
            
            If BkhMode = 1 Then
                ModeStr = "CY Shipment:"
            Else
                If BkhMode = 2 Then
                    ModeStr = "CFS Shipment:"
                Else
                    ModeStr = "CY Shipment:"
                End If
            End If

            ' Report Header
            objExcel.Cells(1, 1) = "Date: " & Format(Now, "dd-MMM-yyyy")
            objExcel.Cells(3, 1) = ModeStr
            objExcel.Range("A1:A3").Font.Bold = True

            iRow = 4

            objExcel.Cells(iRow, iCol) = "P.O.NO."
            objExcel.Cells(iRow, iCol + 1) = "SKU NO."
            objExcel.Cells(iRow, iCol + 2) = "CTNS"
            objExcel.Cells(iRow, iCol + 3) = "PIECES"
            objExcel.Cells(iRow, iCol + 4) = "NO. OF UNIT"

            objExcel.Cells(iRow, iCol + 5) = "TOTAL CBM"
            objExcel.Cells(iRow, iCol + 6) = "NO. OF CNTR"
            objExcel.Cells(iRow, iCol + 7) = "CONTAINER NO."
            objExcel.Cells(iRow, iCol + 8) = "SEAL NO."
            objExcel.Cells(iRow, iCol + 9) = "HBLNO."

            objExcel.Cells(iRow, iCol + 10) = "RECEIVING DATE"
            objExcel.Cells(iRow, iCol + 11) = "V/NAME"
            objExcel.Cells(iRow, iCol + 12) = "REMARKS"

            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Interior.ColorIndex = 15
            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Font.Bold = True
            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(8).LineStyle = 1
            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(9).LineStyle = 1
            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(10).LineStyle = 1
            objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(11).LineStyle = 1
            objExcel.Columns("K:K").NumberFormatLocal = "yyyy/MM/dd"

            ' **********************************************************************

            ' ----------------------------------------------------------------------
            ' Export Report Data
            ' ----------------------------------------------------------------------

            TblIndex = 1

            If ds.Tables(TblIndex).Rows.Count > 0 Then
                hasData = True
            
                iRow = iRow + 1
                startRow = iRow

                For j = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol) = "'" & ds.Tables(TblIndex).Rows(j).Item("PONo").ToString
                    objExcel.Cells(iRow, iCol + 1) = "'" & ds.Tables(TblIndex).Rows(j).Item("SKUNo").ToString
                    objExcel.Cells(iRow, iCol + 2) = ds.Tables(TblIndex).Rows(j).Item("PKG").ToString
                    objExcel.Cells(iRow, iCol + 3) = ds.Tables(TblIndex).Rows(j).Item("WGT").ToString
                    objExcel.Cells(iRow, iCol + 4) = ds.Tables(TblIndex).Rows(j).Item("UntCd").ToString

                    objExcel.Cells(iRow, iCol + 5) = ds.Tables(TblIndex).Rows(j).Item("CBM").ToString
                    objExcel.Cells(iRow, iCol + 6) = ds.Tables(TblIndex).Rows(j).Item("BkhCtnr").ToString
                    objExcel.Cells(iRow, iCol + 7) = "'" & ds.Tables(TblIndex).Rows(j).Item("CtnrNo").ToString
                    objExcel.Cells(iRow, iCol + 8) = "'" & ds.Tables(TblIndex).Rows(j).Item("SealNo").ToString
                    objExcel.Cells(iRow, iCol + 9) = ds.Tables(TblIndex).Rows(j).Item("BkhBLNo").ToString

                    If IsDBNull(ds.Tables(TblIndex).Rows(j).Item("BkhCargoRec")) Then
                        objExcel.Cells(iRow, iCol + 10) = ""
                    Else
                        objExcel.Cells(iRow, iCol + 10) = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(j).Item("BkhCargoRec")), "yyyy/MM/dd")
                    End If
                    objExcel.Cells(iRow, iCol + 11) = ds.Tables(TblIndex).Rows(j).Item("BkhShpr").ToString
                    objExcel.Cells(iRow, iCol + 12) = "'" & ds.Tables(TblIndex).Rows(j).Item("Remark").ToString


                    iRow = iRow + 1
                Next
            End If

            ' CFS Shipments
            If BkhMode = 0 Then
                iRow += 5
                objExcel.Cells(iRow, 1) = "CFS Shipment:"
                objExcel.Cells(iRow, 1).Font.Bold = True

                iRow += 1

                objExcel.Cells(iRow, iCol) = "P.O.NO."
                objExcel.Cells(iRow, iCol + 1) = "SKU NO."
                objExcel.Cells(iRow, iCol + 2) = "CTNS"
                objExcel.Cells(iRow, iCol + 3) = "PIECES"
                objExcel.Cells(iRow, iCol + 4) = "NO. OF UNIT"

                objExcel.Cells(iRow, iCol + 5) = "TOTAL CBM"
                objExcel.Cells(iRow, iCol + 6) = "NO. OF CNTR"
                objExcel.Cells(iRow, iCol + 7) = "CONTAINER NO."
                objExcel.Cells(iRow, iCol + 8) = "SEAl NO."
                objExcel.Cells(iRow, iCol + 9) = "'HBLNO."

                objExcel.Cells(iRow, iCol + 10) = "RECEIVING DATE"
                objExcel.Cells(iRow, iCol + 11) = "V/NAME"
                objExcel.Cells(iRow, iCol + 12) = "REMARKS"

                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(8).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(10).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 12)).Borders(11).LineStyle = 1

                TblIndex += 1
                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    hasData = True

                    iRow = iRow + 1
                    startRow = iRow

                    For j = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol) = "'" & ds.Tables(TblIndex).Rows(j).Item("PONo").ToString
                        objExcel.Cells(iRow, iCol + 1) = "'" & ds.Tables(TblIndex).Rows(j).Item("SKUNo").ToString
                        objExcel.Cells(iRow, iCol + 2) = ds.Tables(TblIndex).Rows(j).Item("PKG").ToString
                        objExcel.Cells(iRow, iCol + 3) = ds.Tables(TblIndex).Rows(j).Item("WGT").ToString
                        objExcel.Cells(iRow, iCol + 4) = ds.Tables(TblIndex).Rows(j).Item("UntCd").ToString

                        objExcel.Cells(iRow, iCol + 5) = ds.Tables(TblIndex).Rows(j).Item("CBM").ToString
                        objExcel.Cells(iRow, iCol + 6) = ds.Tables(TblIndex).Rows(j).Item("BkhCtnr").ToString
                        objExcel.Cells(iRow, iCol + 7) = "'" & ds.Tables(TblIndex).Rows(j).Item("CtnrNo").ToString
                        objExcel.Cells(iRow, iCol + 8) = "'" & ds.Tables(TblIndex).Rows(j).Item("SealNo").ToString
                        objExcel.Cells(iRow, iCol + 9) = ds.Tables(TblIndex).Rows(j).Item("BkhBLNo").ToString

                        If IsDBNull(ds.Tables(TblIndex).Rows(j).Item("BkhCargoRec")) Then
                            objExcel.Cells(iRow, iCol + 10) = ""
                        Else
                            objExcel.Cells(iRow, iCol + 10) = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(j).Item("BkhCargoRec")), "yyyy/MM/dd")
                        End If
                        objExcel.Cells(iRow, iCol + 11) = ds.Tables(TblIndex).Rows(j).Item("BkhShpr").ToString
                        objExcel.Cells(iRow, iCol + 12) = "'" & ds.Tables(TblIndex).Rows(j).Item("Remark").ToString


                        iRow = iRow + 1
                    Next
                End If
            End If

            ' **********************************************************************

            iRow += 3

            ' --------------------------------------------------------------------
            ' Vessel
            ' --------------------------------------------------------------------

            TblIndex += 1

            If ds.Tables(TblIndex).Rows.Count > 0 Then
                objExcel.Cells(iRow, iCol).value = "VESSEL: " & ds.Tables(TblIndex).Rows(0).Item("VslName") & " V." & ds.Tables(TblIndex).Rows(0).Item("VslVoy")
                objExcel.Cells(iRow, iCol).Font.Bold = True
            End If

            ' **********************************************************************


            ' --------------------------------------------------------------------
            ' Ports & Dates
            ' --------------------------------------------------------------------

            TblIndex += 1

            If ds.Tables(TblIndex).Rows.Count > 0 Then
                iRow = iRow + 1
                'Load Port & Date
                For j = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol).value = "ETD " & ds.Tables(TblIndex).Rows(j).Item("BkhLoad") & ":"
                    objExcel.Cells(iRow, iCol + 1).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(j).Item("BkhETD")), "dd-MMM-yyyy")
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 1)).Font.Bold = True

                    iRow = iRow + 1
                Next

                iRow = iRow + 1
                'Disc Port & Date
                For j = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol).value = "ETA " & ds.Tables(TblIndex).Rows(j).Item("BkhDisc") & ":"
                    objExcel.Cells(iRow, iCol + 1).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(j).Item("BkhETA")), "dd-MMM-yyyy")
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 1)).Font.Bold = True

                    iRow = iRow + 1
                Next

                iRow += 1
                'Dest Port & date
                For j = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol).value = "ETA " & ds.Tables(TblIndex).Rows(j).Item("BkhDest") & ":"
                    objExcel.Cells(iRow, iCol + 1).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(j).Item("BkhDestETA")), "dd-MMM-yyyy")
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 1)).Font.Bold = True

                    iRow = iRow + 1
                Next
            End If

            ' **********************************************************************

            ' Set Propertise
            objExcel.Columns("A:B").ColumnWidth = 13
            objExcel.Columns("C:G").ColumnWidth = 15
            objExcel.Columns("H:K").ColumnWidth = 20
            objExcel.Columns("L:M").ColumnWidth = 30

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
                fileName &= ".xls"
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

        ' Return File Path
        RptBSShippingAdvice = fileName
    End Function
End Class
