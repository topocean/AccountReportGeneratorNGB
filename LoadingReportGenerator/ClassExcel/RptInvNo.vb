Public Class RptInvNo

    Function RptInvNo(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim SubNo As String = ""
        Dim InvIndex As Integer = 0
        Dim newRow As Integer

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
                    objExcel.Cells(1, 1) = .Item("BrhName").ToString
                    objExcel.Cells(2, 2) = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString
                End With

                objExcel.Cells(5, 1) = "Invoice Number Report"

                'Setting - bold
                objExcel.Range("A1:G5").Font.Bold = True
                objExcel.Range("A1:G5").HorizontalAlignment = -4108
                objExcel.Range("A1:G1").Merge()
                objExcel.Range("A2:G2").Merge()
                objExcel.Range("A3:G3").Merge()
                objExcel.Range("A5:G5").Merge()

                iRow = 7

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Report Data
                ' ----------------------------------------------------------------------

                TblIndex += 1

                '-- Invoice Part
                objExcel.Cells(iRow, iCol).value = "INVOICE NO."
                objExcel.Cells(iRow, iCol + 1).value = "H/BL NO."
                objExcel.Cells(iRow, iCol + 2).value = "SHIPPER"
                objExcel.Cells(iRow, iCol + 3).value = "ISSUED BY"
                objExcel.Cells(iRow, iCol + 4).value = "LOT NO."

                objExcel.Cells(iRow, iCol + 5).value = "ISSUE DATE"
                objExcel.Cells(iRow, iCol + 6).value = "WEEK"

                ' Setting Border
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(8).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(10).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(11).LineStyle = 1

                iRow = iRow + 1
                startRow = iRow

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("IvhInvNo").ToString
                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("Shipper").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("IssuedBy").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("BkhLotNo").ToString

                    If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("IssueDte")) Then
                        objExcel.Cells(iRow, iCol + 5).value = ""
                    Else
                        objExcel.Cells(iRow, iCol + 5).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("IssueDte")), "yyyy/MM/dd")
                    End If
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("BkhWeek").ToString

                    iRow = iRow + 1
                Next

                '-- Freight List Part
                iRow = iRow + 5
                TblIndex += 1

                If (ds.Tables.Count - 1) >= TblIndex Then
                    If ds.Tables(TblIndex).Rows.Count > 0 Then
                        objExcel.Cells(iRow, iCol).value = "F/L NO."
                        objExcel.Cells(iRow, iCol + 1).value = "M/BL NO."
                        objExcel.Cells(iRow, iCol + 2).value = "AGENT"
                        objExcel.Cells(iRow, iCol + 3).value = "ISSUED BY"
                        objExcel.Cells(iRow, iCol + 4).value = "LOT NO."

                        objExcel.Cells(iRow, iCol + 5).value = "ISSUE DATE"
                        objExcel.Cells(iRow, iCol + 6).value = "WEEK"

                        ' setting border
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Interior.ColorIndex = 15
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(8).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(10).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(11).LineStyle = 1

                        iRow = iRow + 1
                        newRow = iRow
                        For i = 0 To ds.Tables(TblIndex).Rows.Count - 1

                            objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ShhInvNo").ToString
                            objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                            objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("Agent").ToString
                            objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("IssuedBy").ToString
                            objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("BkhLotNo").ToString

                            If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("IssueDte")) Then
                                objExcel.Cells(iRow, iCol + 5).value = ""
                            Else
                                objExcel.Cells(iRow, iCol + 5).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("IssueDte")), "yyyy/MM/dd")
                            End If
                            objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("BkhWeek").ToString

                            iRow = iRow + 1
                        Next
                    End If
                End If

                iRow = iRow + 1

                objExcel.Columns("A:B").ColumnWidth = 15.5
                objExcel.Columns("C:C").ColumnWidth = 50
                objExcel.Columns("D:F").ColumnWidth = 15
                objExcel.Columns("G:G").ColumnWidth = 8

                ' **********************************************************************


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
        RptInvNo = fileName
    End Function

    Function RptInvNoByNumRange(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, j, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim common As New common
        Dim SubNo As String = ""
        Dim InvIndex As Integer = 0
        Dim startNo, endNo As Integer
        Dim RptType As String = ""
        Dim hasRecord As Boolean = False

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
                    startNo = CInt(.Item("InvNoFrm").ToString)
                    endNo = CInt(.Item("InvNoTo").ToString)
                    SubNo = .Item("InvHeader").ToString
                    RptType = .Item("RptType").ToString

                    objExcel.Cells(1, 1) = .Item("BrhName").ToString
                    objExcel.Cells(2, 2) = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3) = "TEL: " & .Item("BrhTel").ToString & "    " & "FAX: " & .Item("BrhFax").ToString
                End With

                objExcel.Cells(5, 1) = "Invoice Number Report"

                'Setting - bold
                objExcel.Range("A1:G5").Font.Bold = True
                objExcel.Range("A1:G5").HorizontalAlignment = -4108
                objExcel.Range("A1:G1").Merge()
                objExcel.Range("A2:G2").Merge()
                objExcel.Range("A3:G3").Merge()
                objExcel.Range("A5:G5").Merge()

                iRow = 7

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Report Data
                ' ----------------------------------------------------------------------

                If RptType = "1" Then
                    objExcel.Cells(iRow, iCol).value = "INVOICE NO."
                    objExcel.Cells(iRow, iCol + 1).value = "H/BL NO."
                Else
                    objExcel.Cells(iRow, iCol).value = "FREIGHT NO."
                    objExcel.Cells(iRow, iCol + 1).value = "M/BL NO."
                End If

                objExcel.Cells(iRow, iCol + 2).value = "SHIPPER"
                objExcel.Cells(iRow, iCol + 3).value = "ISSUED BY"
                objExcel.Cells(iRow, iCol + 4).value = "LOT NO."

                objExcel.Cells(iRow, iCol + 5).value = "ISSUE DATE"
                objExcel.Cells(iRow, iCol + 6).value = "WEEK"

                ' setting border
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Interior.ColorIndex = 15
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Font.Bold = True
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(8).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(9).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(10).LineStyle = 1
                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Borders(11).LineStyle = 1

                iRow = iRow + 1
                startRow = iRow
                i = 0
                TblIndex += 1

                For InvIndex = startNo To endNo
                    If i < ds.Tables(TblIndex).Rows.Count Then
                        If InvIndex = CInt(Right(ds.Tables(TblIndex).Rows(i).Item("IvhInvNo").ToString, 4)) Then
                            hasRecord = True
                        Else
                            hasRecord = False
                        End If
                    Else
                        hasRecord = False
                    End If

                    If hasRecord Then
                        objExcel.Cells(iRow, iCol).value = "'" & ds.Tables(TblIndex).Rows(i).Item("IvhInvNo").ToString
                        objExcel.Cells(iRow, iCol + 1).value = "'" & ds.Tables(TblIndex).Rows(i).Item("BkhBLNo").ToString
                        objExcel.Cells(iRow, iCol + 2).value = "'" & ds.Tables(TblIndex).Rows(i).Item("Shipper").ToString
                        objExcel.Cells(iRow, iCol + 3).value = "'" & ds.Tables(TblIndex).Rows(i).Item("IssuedBy").ToString
                        objExcel.Cells(iRow, iCol + 4).value = "'" & ds.Tables(TblIndex).Rows(i).Item("BkhLotNo").ToString

                        If IsDBNull(ds.Tables(TblIndex).Rows(i).Item("IssueDte")) Then
                            objExcel.Cells(iRow, iCol + 5).value = ""
                        Else
                            objExcel.Cells(iRow, iCol + 5).value = Format(Convert.ToDateTime(ds.Tables(TblIndex).Rows(i).Item("IssueDte")), "yyyy/MM/dd")
                        End If
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("BkhWeek").ToString

                        i += 1
                    Else
                        objExcel.Cells(iRow, iCol).value = "'" & SubNo & InvIndex.ToString.PadLeft(4, "0")
                        objExcel.Cells(iRow, iCol + 1).value = "NO DATA"

                        ' Highlight as Red
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 6)).Font.Color = RGB(255, 0, 0)
                    End If

                    iRow += 1
                Next

                iRow = iRow + 1

                objExcel.Columns("A:B").ColumnWidth = 15.5
                objExcel.Columns("C:C").ColumnWidth = 50
                objExcel.Columns("D:F").ColumnWidth = 15
                objExcel.Columns("G:G").ColumnWidth = 8

                ' **********************************************************************


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
        RptInvNoByNumRange = fileName
    End Function
End Class
