Public Class RptLiftingSummary

    Function RptLiftingSummary(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
        System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
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
                ' Get File Name, Report Header
                ' ----------------------------------------------------------------------

                TblIndex = 0

                With ds.Tables(TblIndex).Rows(0)
                    fileName = .Item("RptFile").ToString
                    objExcel.Cells(1, 1).value = .Item("BrhName").ToString
                    objExcel.Cells(2, 2).value = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3).value = "TEL: " & .Item("BrhTel").ToString & "  FAX: " & .Item("BrhFax").ToString

                    objExcel.Cells(5, 1).value = "LIFTING SUMMARY"
                    objExcel.Cells(6, 1).value = "DATE : WEEK " & .Item("WeekNo").ToString

                    'Setting - bold
                    objExcel.Range("A1:H7").Font.Bold = True
                    objExcel.Range("A1:H3").HorizontalAlignment = -4108
                    objExcel.Range("A1:H1").Merge()
                    objExcel.Range("A2:H2").Merge()
                    objExcel.Range("A3:H3").Merge()
                    objExcel.Range("A5:H5").Merge()
                    objExcel.Range("A6:H6").Merge()
                    objExcel.Range("A7:H8").HorizontalAlignment = -4108

                    iRow = 8
                    'Upper part title
                    objExcel.Cells(iRow, iCol + 2).value = .Item("CurrYear").ToString
                    objExcel.Cells(iRow, iCol + 5).value = .Item("PreYear").ToString
                    objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 4)).Merge()
                    objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 7)).Merge()

                    iRow = iRow + 1
                    objExcel.Cells(iRow, iCol + 2).value = "WK" & CInt(.Item("WeekNo").ToString) - 1
                    objExcel.Cells(iRow, iCol + 3).value = "WK" & .Item("WeekNo").ToString
                    objExcel.Cells(iRow, iCol + 4).value = "WK" & CInt(.Item("WeekNo").ToString) + 1

                    objExcel.Cells(iRow, iCol + 5).value = "WK" & CInt(.Item("WeekNo").ToString) - 1
                    objExcel.Cells(iRow, iCol + 6).value = "WK" & .Item("WeekNo").ToString
                    objExcel.Cells(iRow, iCol + 7).value = "WK" & CInt(.Item("WeekNo").ToString) + 1

                    iRow = iRow + 1
                    objExcel.Cells(iRow, iCol).value = "1."
                    objExcel.Cells(iRow, iCol + 1).value = "USA (FEU)"
                End With

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 2nd recordset USA(FEU)
                ' ----------------------------------------------------------------------
                TblIndex = 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    objExcel.Cells(iRow, iCol + 2).value = 0
                    objExcel.Cells(iRow, iCol + 3).value = 0
                    objExcel.Cells(iRow, iCol + 4).value = 0

                    objExcel.Cells(iRow, iCol + 5).value = 0
                    objExcel.Cells(iRow, iCol + 6).value = 0
                    objExcel.Cells(iRow, iCol + 7).value = 0
                Else
                    objExcel.Cells(iRow, iCol + 2).value = CInt(ds.Tables(TblIndex).Rows(0).Item("wk1").ToString)
                    objExcel.Cells(iRow, iCol + 3).value = CInt(ds.Tables(TblIndex).Rows(0).Item("wk2").ToString)
                    objExcel.Cells(iRow, iCol + 4).value = CInt(ds.Tables(TblIndex).Rows(0).Item("wk3").ToString)

                    objExcel.Cells(iRow, iCol + 5).value = CInt(ds.Tables(TblIndex).Rows(0).Item("Prewk1").ToString)
                    objExcel.Cells(iRow, iCol + 6).value = CInt(ds.Tables(TblIndex).Rows(0).Item("Prewk2").ToString)
                    objExcel.Cells(iRow, iCol + 7).value = CInt(ds.Tables(TblIndex).Rows(0).Item("Prewk3").ToString)
                End If

                iRow = iRow + 2
                objExcel.Cells(iRow, iCol).value = "2."
                objExcel.Cells(iRow, iCol + 1).value = "ALLOCATION"
                iRow = iRow + 1

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' 3rd recordset '2 - ALLOCATION
                ' ----------------------------------------------------------------------

                TblIndex = 2

                If ds.Tables(TblIndex).Rows.Count > 0 Then

                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1

                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("BrhSName").ToString
                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString
                        iRow = iRow + 1

                    Next
                End If

                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "3."
                objExcel.Cells(iRow, iCol + 1).value = "USA Document (SET)"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 4th recordset - USA Document(Set)
                ' ----------------------------------------------------------------------

                TblIndex = 3

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("UsrName")

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1

                    Next
                End If

                iRow = iRow + 2
                objExcel.Cells(iRow, iCol).value = "4."
                objExcel.Cells(iRow, iCol + 1).value = "Sales Volume (Feu)"
                objExcel.Cells(iRow, iCol + 5).value = "C&F"
                objExcel.Cells(iRow, iCol + 6).value = "FOB"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 5th recordset - Sale volume(Feu)
                ' ----------------------------------------------------------------------

                TblIndex = 4

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("SalName").ToString

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1

                    Next

                End If

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 6th recordset
                ' ----------------------------------------------------------------------

                TblIndex = 5

                objExcel.Cells(iRow, iCol + 1).value = "Consol Box"
                If ds.Tables(TblIndex).Rows.Count <= 0 Then
                    objExcel.Cells(iRow, iCol + 2).value = 0
                    objExcel.Cells(iRow, iCol + 3).value = 0
                    objExcel.Cells(iRow, iCol + 4).value = 0

                    objExcel.Cells(iRow, iCol + 5).value = 0
                    objExcel.Cells(iRow, iCol + 6).value = 0
                    objExcel.Cells(iRow, iCol + 7).value = 0
                    iRow = iRow + 1

                    objExcel.Cells(iRow, iCol + 2).value = 0
                    objExcel.Cells(iRow, iCol + 3).value = 0
                    objExcel.Cells(iRow, iCol + 4).value = 0

                    objExcel.Cells(iRow, iCol + 5).value = 0
                    objExcel.Cells(iRow, iCol + 6).value = 0
                    objExcel.Cells(iRow, iCol + 7).value = 0
                Else
                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(0).Item("wk1").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(0).Item("wk2").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(0).Item("wk3").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(0).Item("Prewk1").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(0).Item("Prewk2").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(0).Item("Prewk3").ToString
                    iRow = iRow + 1

                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(0).Item("wk1CBM").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(0).Item("wk2CBM").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(0).Item("wk3CBM").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(0).Item("Prewk1CBM").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(0).Item("Prewk2CBM").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(0).Item("Prewk3CBM").ToString
                End If

                iRow = iRow + 2
                objExcel.Cells(iRow, iCol).value = "5."
                objExcel.Cells(iRow, iCol + 1).value = "Top Customers (USA)"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 7th recordset 5-Top Customers (USA)
                ' ----------------------------------------------------------------------

                TblIndex = 6

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("ConName")

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1
                    Next
                End If
                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "6."
                objExcel.Cells(iRow, iCol + 1).value = "C & F Customers"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 8th recordset 6-C & F Customers
                ' ----------------------------------------------------------------------

                TblIndex = 7

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1
                    Next
                End If
                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "7."
                objExcel.Cells(iRow, iCol + 1).value = "Co-Loader Business"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 9th recordset 7: Co-Loader Business
                ' ----------------------------------------------------------------------

                TblIndex = 8

                If ds.Tables(TblIndex).Rows.Count > 0 Then

                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1

                    Next
                End If

                objExcel.Columns("A:A").ColumnWidth = 10
                objExcel.Columns("B:B").ColumnWidth = 25
                objExcel.Columns("C:AA").ColumnWidth = 12

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
        RptLiftingSummary = fileName
    End Function

    Function RptLiftingSummaryNUS(ByVal UID As String, ByVal ds As DataSet) As String
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
                    objExcel.Cells(1, 1).value = .Item("BrhName").ToString
                    objExcel.Cells(2, 2).value = .Item("BrhAddr").ToString
                    objExcel.Cells(3, 3).value = "TEL: " & .Item("BrhTel").ToString & "  FAX: " & .Item("BrhFax").ToString

                    objExcel.Cells(5, 1).value = "LIFTING SUMMARY NON-USA"
                    objExcel.Cells(6, 1).value = "DATE : WEEK " & .Item("WeekNo").ToString

                    'Setting - bold
                    objExcel.Range("A1:H7").Font.Bold = True
                    objExcel.Range("A1:H3").HorizontalAlignment = -4108
                    objExcel.Range("A1:H1").Merge()
                    objExcel.Range("A2:H2").Merge()
                    objExcel.Range("A3:H3").Merge()
                    objExcel.Range("A5:H5").Merge()
                    objExcel.Range("A6:H6").Merge()
                    objExcel.Range("A7:H8").HorizontalAlignment = -4108

                    iRow = 7
                    'Upper part title
                    objExcel.Cells(iRow, iCol + 2).value = .Item("CurrYear").ToString
                    objExcel.Cells(iRow, iCol + 5).value = .Item("PreYear").ToString
                    objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 4)).Merge()
                    objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 7)).Merge()

                    iRow = iRow + 1
                    objExcel.Cells(iRow, iCol + 2).value = "WK" & CInt(.Item("WeekNo").ToString) - 1
                    objExcel.Cells(iRow, iCol + 3).value = "WK" & .Item("WeekNo")
                    objExcel.Cells(iRow, iCol + 4).value = "WK" & CInt(.Item("WeekNo").ToString) + 1

                    objExcel.Cells(iRow, iCol + 5).value = "WK" & CInt(.Item("WeekNo").ToString) - 1
                    objExcel.Cells(iRow, iCol + 6).value = "WK" & .Item("WeekNo")
                    objExcel.Cells(iRow, iCol + 7).value = "WK" & CInt(.Item("WeekNo").ToString) + 1

                    iRow = iRow + 1
                    objExcel.Cells(iRow, iCol).value = "1."
                    objExcel.Cells(iRow, iCol + 1).value = "NON-USA (FEU)"
                End With

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 2nd recordset 1 - NON-USA (FEU)
                ' ----------------------------------------------------------------------

                TblIndex = 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then

                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(0).Item("wk1").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(0).Item("wk2").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(0).Item("wk3").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(0).Item("Prewk1").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(0).Item("Prewk2").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(0).Item("Prewk3").ToString

                End If

                iRow = iRow + 1
                TblIndex = 2

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("Traffic").ToString

                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString
                    iRow = iRow + 1

                Next

                iRow = iRow + 2
                objExcel.Cells(iRow, iCol).value = "2."
                objExcel.Cells(iRow, iCol + 1).value = "NON-USA Document (Set)"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 3rd recordset 2- NON-USA Document (Set)
                ' ----------------------------------------------------------------------

                TblIndex = 3

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1

                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("UsrName").ToString
                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString
                        iRow = iRow + 1

                    Next
                End If
                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "3."
                objExcel.Cells(iRow, iCol + 1).value = "Import (Feu)"
                iRow = iRow + 1

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' 4th recordset 3-Import (Feu)
                ' ----------------------------------------------------------------------

                TblIndex = 4

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("UsrName").ToString

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1
                    Next
                End If
                iRow = iRow + 2
                objExcel.Cells(iRow, iCol).value = "4."
                objExcel.Cells(iRow, iCol + 1).value = "TOP 10 customers (NON-USA)"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 5th recordset 4 TOP 10 customer(NON-USA)
                ' ----------------------------------------------------------------------

                TblIndex = 5

                If ds.Tables(TblIndex).Rows.Count > 0 Then

                    While ds.Tables(TblIndex).Rows(0).Item(0).ToString <> "END LOOP"
                        objExcel.Cells(iRow, iCol).value = "- " & ds.Tables(TblIndex).Rows(0).Item("Traffic").ToString

                        TblIndex += 1

                        If ds.Tables(TblIndex).Rows.Count > 0 Then
                            For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                                objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("ConName").ToString & " " & ds.Tables(TblIndex).Rows(i).Item("LocName").ToString

                                objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                                objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                                objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                                objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                                objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                                objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                                iRow = iRow + 1

                            Next
                        Else
                            objExcel.Cells(iRow, iCol + 1).value = "- "

                            objExcel.Cells(iRow, iCol + 2).value = 0
                            objExcel.Cells(iRow, iCol + 3).value = 0
                            objExcel.Cells(iRow, iCol + 4).value = 0

                            objExcel.Cells(iRow, iCol + 5).value = ""
                            objExcel.Cells(iRow, iCol + 6).value = ""
                            objExcel.Cells(iRow, iCol + 7).value = ""

                            iRow = iRow + 1
                        End If

                        TblIndex += 1
                    End While

                End If

                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "5."
                objExcel.Cells(iRow, iCol + 1).value = "Sales Volume (Feu)"
                objExcel.Cells(iRow, iCol + 5).value = "C&F"
                objExcel.Cells(iRow, iCol + 6).value = "FOB"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 6th recordset - Sales Volume(Feu)
                ' ----------------------------------------------------------------------

                TblIndex += 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("SalName").ToString

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1
                    Next
                End If
                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "6."
                objExcel.Cells(iRow, iCol + 1).value = "Non USA - C & F Customers"
                objExcel.Cells(iRow, iCol + 5).value = "C&F"
                objExcel.Cells(iRow, iCol + 6).value = "FOB"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 7th recordset 6- NON USA - C& F Customers
                ' ----------------------------------------------------------------------

                TblIndex += 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString & " " & ds.Tables(TblIndex).Rows(i).Item("LocName").ToString

                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                        iRow = iRow + 1
                    Next
                End If
                iRow = iRow + 1
                objExcel.Cells(iRow, iCol).value = "7."
                objExcel.Cells(iRow, iCol + 1).value = "Non USA - Co-Loader Business"
                iRow = iRow + 1

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' 8th recordset - Non USA - Co-Loader Business
                ' ----------------------------------------------------------------------

                TblIndex += 1

                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                    objExcel.Cells(iRow, iCol + 1).value = "- " & ds.Tables(TblIndex).Rows(i).Item("ShpName").ToString & " " & ds.Tables(TblIndex).Rows(i).Item("LocName").ToString

                    objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("wk1").ToString
                    objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("wk2").ToString
                    objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("wk3").ToString

                    objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("Prewk1").ToString
                    objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("Prewk2").ToString
                    objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("Prewk3").ToString

                    iRow = iRow + 1
                Next

                ' **********************************************************************


                ' ----------------------------------------------------------------------
                ' Set Properties
                ' ----------------------------------------------------------------------

                objExcel.Columns("A:A").ColumnWidth = 10
                objExcel.Columns("B:B").ColumnWidth = 25
                objExcel.Columns("C:AA").ColumnWidth = 12

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
        RptLiftingSummaryNUS = fileName
    End Function
End Class
