Public Class RptLoading_Carrier

    Function RptLoading_Carrier(ByVal UID As String, ByVal ds As DataSet) As String
        System.Threading.Thread.CurrentThread.CurrentCulture = _
                System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet
        Dim i As Integer
        Dim iRow, iSRow, iCol, cCol As Integer
        Dim fileName As String
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
                cCol = 1

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Get File Name
                ' ----------------------------------------------------------------------

                'fileName = common.NullVal(ds.Tables(0).Rows(0).Item("fName").ToString, UID)

                ' ----------------------------------------------------------------------
                ' Export Report Header (Company Name, Address, Tel, etc...)
                ' ----------------------------------------------------------------------

                With ds.Tables(0).Rows(0)
                    fileName = common.NullVal(.Item("RptFile").ToString, UID)
                    objWS.Application.Cells(1, 1) = common.NullVal(.Item("BrhName").ToString, "")
                    objWS.Application.Cells(2, 2) = common.NullVal(.Item("BrhAddr").ToString, "")
                    objWS.Application.Cells(3, 3) = "TEL: " & common.NullVal(.Item("BrhTel").ToString(), "")
                    objWS.Application.Cells(5, 1) = "LOADING REPORT - YEAR: " & common.NullVal(.Item("HdrYear").ToString, "")

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

                objWS.Application.Cells(iRow, iCol) = "WEEK"

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    objWS.Application.Cells(iRow, iCol + cCol) = ds.Tables(1).Rows(i).Item("SCAC")
                    cCol += 1
                Next

                objWS.Application.Cells(iRow, iCol + cCol) = "TOTAL"

                ' **********************************************************************

                ' ----------------------------------------------------------------------
                ' Setting Properties of Detail Header
                ' ----------------------------------------------------------------------

                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + cCol)).Interior.ColorIndex = 15
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + cCol)).Font.Bold = True
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + cCol)).Borders(8).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + cCol)).Borders(9).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + cCol)).Borders(10).LineStyle = 1
                objWS.Application.Range(objWS.Application.Cells(iRow, 1), objWS.Application.Cells(iRow, iCol + cCol)).Borders(11).LineStyle = 1

                ' **********************************************************************

                iRow += 1
                iSRow = iRow

                ' ----------------------------------------------------------------------
                ' Export Report Data onto Excel Wooksheet
                ' ----------------------------------------------------------------------
                Dim k, index As Integer

                index = 0
                cCol = 1

                For k = 1 To 53
                    For i = index To ds.Tables(2).Rows.Count - 1
                        If ds.Tables(2).Rows(i).Item("RptWeek") > k Then
                            Exit For
                        End If

                        objWS.Application.Cells(iRow, iCol) = "WK " & ds.Tables(2).Rows(i).Item("RptWeek")

                        If ds.Tables(2).Rows(i).Item("RptCount") = 0 Then
                            objWS.Application.Cells(iRow, iCol + cCol) = ""
                        Else
                            objWS.Application.Cells(iRow, iCol + cCol) = ds.Tables(2).Rows(i).Item("RptCount")
                        End If
                        cCol += 1
                    Next

                    ' Total
                    objWS.Application.Range(objWS.Application.Cells(iRow, 2), objWS.Application.Cells(iRow, iCol + cCol - 1)).Select()
                    objWS.Application.Range(objWS.Application.Cells(iRow, iCol + cCol), objWS.Application.Cells(iRow, iCol + cCol)).Activate()
                    objWS.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[-" & iCol + cCol - 2 & "]:RC[-1])"

                    ' Reset Counters
                    iRow += 1
                    index = i
                    cCol = 1
                Next

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
        Else
            fileName = ""
        End If

        ' Return File Path
        RptLoading_Carrier = fileName
    End Function
End Class
