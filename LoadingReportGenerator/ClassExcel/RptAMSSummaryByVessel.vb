Public Class RptAMSSummaryByVessel

    Function RptAMSSummaryByVessel(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim PortCount As Integer = 0
        Dim PortName As String = ""

        If ds.Tables(1).Rows.Count > 0 Then
            ' Start Excel Application
            objExcel = CreateObject("Excel.Application")
            objExcel.Visible = False

            Try
                ' Get a new workbook
                objWB = objExcel.Workbooks.Add
                objWS = objWB.ActiveSheet

                ' Set Worksheet Properties
                objWS.Application.Cells.Font.Name = "Trebuchet MS"
                objWS.Application.Cells.Font.Size = 12
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

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    With ds.Tables(TblIndex).Rows(0)
                        fileName = .Item("RptFile").ToString
                        PortCount = CInt(.Item("PortCount").ToString)
                    End With
                End If

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Export Report Data (Vessel Detail)
                ' ----------------------------------------------------------------------

                TblIndex = 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    startRow = 1
                    iRow = startRow
                    iCol = 1

                    objExcel.Cells(iRow, iCol) = "VESSEL"
                    objExcel.Cells(iRow, iCol + 1) = "VOY"
                    objExcel.Cells(iRow, iCol + 2) = "VSL CODE"
                    objExcel.Cells(iRow, iCol + 3) = "SVC"
                    objExcel.Cells(iRow, iCol + 4 + PortCount) = "CARRIER"
                    objExcel.Cells(iRow, iCol + 5 + PortCount) = "AMS THRU"

                    ' Set Propertise
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 5 + PortCount + (PortCount * 4))).Interior.Color = RGB(255, 255, 0)
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 5 + PortCount + (PortCount * 4))).Borders(8).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 5 + PortCount + (PortCount * 4))).Borders(9).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 5 + PortCount + (PortCount * 4))).Borders(10).LineStyle = 1
                    objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, iCol + 5 + PortCount + (PortCount * 4))).Borders(11).LineStyle = 1

                    iRow = iRow + 1

                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol)).Interior.Color = RGB(255, 255, 0)

                        objExcel.Cells(iRow, iCol) = "'" & ds.Tables(TblIndex).Rows(i).Item("VslName").ToString
                        objExcel.Cells(iRow, iCol + 1) = "'" & ds.Tables(TblIndex).Rows(i).Item("VslVoy").ToString
                        objExcel.Cells(iRow, iCol + 2) = "'" & ds.Tables(TblIndex).Rows(i).Item("VslCode").ToString
                        objExcel.Cells(iRow, iCol + 3) = "'" & ds.Tables(TblIndex).Rows(i).Item("VslSVC").ToString
                        objExcel.Cells(iRow, iCol + 4 + PortCount) = "'" & ds.Tables(TblIndex).Rows(i).Item("CarName").ToString
                        objExcel.Cells(iRow, iCol + 5 + PortCount) = "'" & ds.Tables(TblIndex).Rows(i).Item("AMSThru").ToString

                        ' Set Borders
                        objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, (iCol + 5 + PortCount + (PortCount * 4)))).Borders(8).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, (iCol + 5 + PortCount + (PortCount * 4)))).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, (iCol + 5 + PortCount + (PortCount * 4)))).Borders(10).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow, (iCol + 5 + PortCount + (PortCount * 4)))).Borders(11).LineStyle = 1

                        ' Set Column Width
                        objExcel.Range(objExcel.Cells(iRow, iCol + 4 + PortCount), objExcel.Cells(iRow, iCol + 5 + PortCount)).ColumnWidth = 12

                        

                        iRow = iRow + 1
                    Next

                    For j = 0 To PortCount - 1
                        TblIndex += 1

                        If ds.Tables(TblIndex).Rows.Count > 0 Then
                            'Title row (5, 10, 11, 12, 13)
                            iRow = startRow
                            objExcel.Cells(iRow, (iCol + 4 + j)) = "ETD " & ds.Tables(TblIndex).Rows(0).Item("LocSName").ToString
                            objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4))) = ds.Tables(TblIndex).Rows(0).Item("LocSName").ToString & " SI C-O"
                            objExcel.Range(objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4))), objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 1))).Merge()
                            objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 2)) = ""
                            objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 3)) = "ON WEB"

                            iRow += 1
                            TblIndex += 1

                            If ds.Tables(TblIndex).Rows.Count > 0 Then
                                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                                    If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("VslETD")) Then
                                        objExcel.Cells(iRow, (iCol + 4 + j)) = "NIL"
                                    Else
                                        objExcel.Cells(iRow, (iCol + 4 + j)) = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("VslETD")), "dd/MM/yyyy")
                                    End If

                                    If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("VslSICutOff")) Then
                                        objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4))) = "NIL"
                                        objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 1)) = "NIL"
                                    Else
                                        objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4))) = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("VslSICutOff")), "dd/MM/yyyy")
                                        objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 1)) = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("VslSICutOff")), "HH:mm:ss")
                                    End If

                                    objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 2)) = ""
                                    objExcel.Cells(iRow, (iCol + 6 + PortCount + (j * 4) + 3)) = ""

                                    objExcel.Range(objExcel.Cells(iRow, 5), objExcel.Cells(iRow, 5 + PortCount)).ColumnWidth = 12
                                    objExcel.Range(objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4)), objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4))).ColumnWidth = 12
                                    objExcel.Range(objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4) + 1), objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4) + 1)).ColumnWidth = 7.63
                                    objExcel.Range(objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4) + 2), objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4) + 2)).ColumnWidth = 20.88
                                    objExcel.Range(objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4) + 3), objExcel.Cells(iRow, iCol + 6 + PortCount + (j * 4) + 3)).ColumnWidth = 8.88

                                    ' Set Center
                                    objExcel.Range(objExcel.Cells(1, iCol + 1), objExcel.Cells(iRow, iCol + 3)).HorizontalAlignment = -4108
                                    objExcel.Range(objExcel.Cells(1, iCol + 4 + PortCount), objExcel.Cells(iRow, iCol + 5 + PortCount)).HorizontalAlignment = -4108

                                    iRow += 1
                                Next
                            End If
                        End If
                    Next
                End If

                objExcel.Columns("A:A").ColumnWidth = 29.25
                objExcel.Columns("B:B").ColumnWidth = 10.38
                objExcel.Columns("C:C").ColumnWidth = 13.13
                objExcel.Columns("D:D").ColumnWidth = 15.13

                ' ----------------------------------------------------------------------
                ' Save File
                ' ----------------------------------------------------------------------

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
        Else
        fileName = ""
        End If

        ' Return File Path
        RptAMSSummaryByVessel = fileName
    End Function
End Class
