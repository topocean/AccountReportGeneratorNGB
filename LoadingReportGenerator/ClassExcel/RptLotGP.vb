Public Class RptLotGP

    Function RptLotGP(ByVal Uid As String, ByVal sUid As String, ByVal ds As DataSet) As String

        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

        Dim objExcel As Excel.Application
        Dim objWB As Excel.Workbook
        Dim objWS As Excel.Worksheet

        Dim common As New common

        Dim i, startRow, TblIndex As Integer
        Dim iRow, iSRow, iCol As Integer
        Dim fileName As String = ""
        Dim Total1, Total2, Total3, Total4, Total5, Total6, Total7, Total8 As Double
        Dim PrintType, TypeRefId As Integer

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

            ' Define the starting row and column number of the detail header
            iRow = 1
            iCol = 1

            fileName = ds.Tables(0).Rows(0).Item("fname").ToString

            TblIndex = 1

            ' Retrieve File Name
            If ds.Tables(1).Rows.Count > 0 Then
                With ds.Tables(TblIndex).Rows(0)
                    PrintType = CInt(.Item("PrintType").ToString)
                    TypeRefId = CInt(.Item("TypeRefId").ToString)

                    Total1 = 0
                    Total2 = 0
                    Total3 = 0
                    Total4 = 0
                    Total5 = 0
                    Total6 = 0
                    Total7 = 0
                    Total8 = 0
                End With
            End If

            iRow = 1

            If PrintType = 11 Or PrintType = 12 Then 'LOT GROSS PROFIT DETAIL REPORT
                objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(0).Item("BrhHeader").ToString
                objExcel.Cells(iRow, iCol + 6).value = "Report Date:"
                objExcel.Cells(iRow, iCol + 7).value = Format(Now, "dd-MMM-yyyy")
                objExcel.Cells(iRow + 1, iCol).value = "LOT GROSS PROFIT"

                'bold font
                objExcel.Range(objExcel.Cells(iRow, iCol), objExcel.Cells(iRow + 1, iCol)).Font.Bold = True

                iRow = iRow + 1

                TblIndex += 1

                With ds.Tables(TblIndex).Rows(0)
                    objExcel.Cells(iRow + 2, iCol).value = "LOT NO: " & .Item("ShhLotNo").ToString
                    objExcel.Cells(iRow + 3, iCol).value = "VESSEL: " & .Item("VslName").ToString
                    objExcel.Cells(iRow + 4, iCol).value = "DEST: " & .Item("ShhDestName").ToString
                    objExcel.Cells(iRow + 5, iCol).value = "SHIPPER: " & .Item("ShhShprName").ToString

                    objExcel.Cells(iRow + 2, iCol + 2).value = "MB/L: " & .Item("ShhMBLNo").ToString
                    objExcel.Cells(iRow + 3, iCol + 2).value = "VOYAGE: " & .Item("VslVoyName").ToString
                    objExcel.Cells(iRow + 4, iCol + 2).value = "ETD: " & .Item("VslOnBoard").ToString
                    objExcel.Cells(iRow + 5, iCol + 2).value = "CONSIGNEE: " & .Item("ShhConName").ToString

                    objExcel.Cells(iRow + 8, iCol).value = "DOCUMENT NO."
                    objExcel.Cells(iRow + 8, iCol + 1).value = "CHINESE INV"
                    objExcel.Cells(iRow + 8, iCol + 2).value = "REFERENCE NO"
                    objExcel.Cells(iRow + 8, iCol + 3).value = "NAME"
                    objExcel.Cells(iRow + 8, iCol + 4).value = "CHARGE NAME"
                    objExcel.Cells(iRow + 8, iCol + 5).value = "CURR"
                    objExcel.Cells(iRow + 8, iCol + 6).value = "EX-RATE"
                    objExcel.Cells(iRow + 8, iCol + 7).value = "A/C EX-RATE"
                    objExcel.Cells(iRow + 8, iCol + 8).value = "INCOME"
                    objExcel.Cells(iRow + 8, iCol + 11).value = "COST"

                    objExcel.Cells(iRow + 9, iCol + 8).value = ""
                    objExcel.Cells(iRow + 9, iCol + 9).value = "RMB"
                    objExcel.Cells(iRow + 9, iCol + 10).value = "RMB" & Chr(10) & Chr(13) & "A/C EX-RATE"
                    objExcel.Cells(iRow + 9, iCol + 11).value = ""
                    objExcel.Cells(iRow + 9, iCol + 12).value = "RMB"
                    objExcel.Cells(iRow + 9, iCol + 13).value = "RMB" & Chr(10) & Chr(13) & "A/C EX-RATE"

                    'bold font
                    objExcel.Range(objExcel.Cells(iRow + 8, iCol), objExcel.Cells(iRow + 8, iCol + 13)).Font.Bold = True
                    objExcel.Range("I10:K10").HorizontalAlignment = -4108
                    objExcel.Range("L10:N10").HorizontalAlignment = -4108
                    objExcel.Range("I10:K10").Merge()
                    objExcel.Range("L10:N10").Merge()
                End With

                iRow = iRow + 10
                iSRow = iRow

                TblIndex += 1

                If ds.Tables(TblIndex).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                        objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ShhInvNo").ToString
                        objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("InvNo").ToString
                        objExcel.Cells(iRow, iCol + 2).value = ds.Tables(TblIndex).Rows(i).Item("ShdBLNo").ToString
                        objExcel.Cells(iRow, iCol + 3).value = ds.Tables(TblIndex).Rows(i).Item("AgtName").ToString
                        objExcel.Cells(iRow, iCol + 4).value = ds.Tables(TblIndex).Rows(i).Item("ChgName").ToString
                        objExcel.Cells(iRow, iCol + 5).value = ds.Tables(TblIndex).Rows(i).Item("ShdCurr").ToString
                        objExcel.Cells(iRow, iCol + 6).value = ds.Tables(TblIndex).Rows(i).Item("ShdExRate").ToString
                        objExcel.Cells(iRow, iCol + 7).value = ds.Tables(TblIndex).Rows(i).Item("AccExRate").ToString
                        objExcel.Cells(iRow, iCol + 8).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("orgIncome").ToString)
                        objExcel.Cells(iRow, iCol + 9).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("EquiIncome").ToString)
                        objExcel.Cells(iRow, iCol + 10).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("EquiIncomeAccRate").ToString)
                        objExcel.Cells(iRow, iCol + 11).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("orgOutlay").ToString)
                        objExcel.Cells(iRow, iCol + 12).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("EquiOutlay").ToString)
                        objExcel.Cells(iRow, iCol + 13).value = CDbl(ds.Tables(TblIndex).Rows(i).Item("EquiOutlayAccRate").ToString)

                        iRow = iRow + 1
                    Next

                    'For i = 10 To 12 Step 3
                    For i = 10 To 11
                        objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 1, i)).Select()
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                        objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Borders(9).LineStyle = -4119
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Borders(8).LineStyle = 1
                    Next

                    ' Account Rate
                    For i = 13 To 14
                        objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 1, i)).Select()
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                        objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Borders(9).LineStyle = -4119
                        objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Borders(8).LineStyle = 1
                    Next

                    iRow = iRow + 2

                    objExcel.Cells(iRow, 4).value = "GROSS PROFIT"
                    objExcel.Range(objExcel.Cells(iRow, 10), objExcel.Cells(iRow, 10)).Activate()
                    objExcel.ActiveCell.FormulaR1C1 = "=R[-2]C-R[-2]C[3]"

                    ' Account Rate
                    objExcel.Range(objExcel.Cells(iRow, 11), objExcel.Cells(iRow, 11)).Activate()
                    objExcel.ActiveCell.FormulaR1C1 = "=R[-2]C-R[-2]C[3]"

                    objExcel.Range(objExcel.Cells(iRow, 13), objExcel.Cells(iRow, 13)).Activate()
                    objExcel.ActiveCell.FormulaR1C1 = "=(RC[-3]/R[-2]C[-3])*10000%"

                    ' Account Rate
                    objExcel.Range(objExcel.Cells(iRow, 14), objExcel.Cells(iRow, 14)).Activate()
                    objExcel.ActiveCell.FormulaR1C1 = "=(RC[-3]/R[-2]C[-3])*10000%"

                    objExcel.Range(objExcel.Cells(iSRow, iCol + 9), objExcel.Cells(iRow, iCol + 14)).NumberFormatLocal = "#,##0.00_ "

                    'setting width
                    objExcel.Columns("A:A").ColumnWidth = 13
                    objExcel.Columns("B:B").ColumnWidth = 10
                    objExcel.Columns("C:C").ColumnWidth = 13
                    objExcel.Columns("D:D").ColumnWidth = 30
                    objExcel.Columns("E:E").ColumnWidth = 17
                    objExcel.Columns("F:F").ColumnWidth = 6
                    objExcel.Columns("G:N").ColumnWidth = 14
                End If
            Else
                If PrintType = 6 Or PrintType = 10 Or PrintType = 13 Then 'By Period or By Vessel or By MBL No
                    If TypeRefId = 0 Then
                        i = 0

                        TblIndex = 1

                        If ds.Tables(1).Rows.Count <= 0 Or ds.Tables(0).Rows.Count > 0 Then
                            While i < ds.Tables(TblIndex).Rows.Count
                                iSRow = iRow
                                objExcel.Cells(iRow, 1).value = "LOT GP SUMMARY REPORT BY " & ds.Tables(TblIndex).Rows.Item("title").ToString

                                If ds.Tables(TblIndex).Rows(0).Item("ReportBy") = "P" Then
                                    objExcel.Cells(iRow + 2, 1).value = "PERIOD (ETD): " & Format(ds.Tables(TblIndex).Rows(i).Item("DteFrm"), "dd/mm/yy") & "-" & Format(ds.Tables(TblIndex).Rows(i).Item("DteTo"), "dd/mm/yy")
                                Else
                                    If ds.Tables(TblIndex).Rows(i).Item("MonthNo") <> 0 Then
                                        objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " MONTH " & common.DigitToMonth(ds.Tables(TblIndex).Rows(i).Item("MonthNo"))
                                    Else
                                        objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " WEEK " & ds.Tables(TblIndex).Rows(i).Item("WeekNo")
                                    End If
                                End If

                                If i = 10 Then
                                    objExcel.Cells(iRow + 2, 1).value = "VESSEL/VOY: " & ds.Tables(TblIndex).Rows(i).Item("VslName") & "      CARRIER: " & ds.Tables(TblIndex).Rows(i).Item("CarrName") & "      DATE: " & Format(ds.Tables(TblIndex).Rows(i).Item("VslOnBoard"), "dd/mm/yy")
                                End If

                                iRow = iRow + 5

                                'Upper part title
                                If i = 13 Then
                                    objExcel.Cells(iRow, iCol).value = "MASTER BL"
                                    objExcel.Cells(iRow, iCol + 1).value = "HOUSE BL"
                                Else
                                    objExcel.Cells(iRow, iCol).value = "LOT"
                                    objExcel.Cells(iRow, iCol + 1).value = "MASTER BL"
                                End If
                                objExcel.Cells(iRow, iCol + 2).value = "INCOME"
                                objExcel.Cells(iRow, iCol + 3).value = "COST"
                                objExcel.Cells(iRow, iCol + 4).value = "GP"
                                objExcel.Cells(iRow, iCol + 5).value = "GP MARGIN(%)"
                                objExcel.Cells(iRow, iCol + 6).value = "INCOME" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 7).value = "COST" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 8).value = "GP" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 9).value = "GP MARGIN(%)" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 10).value = "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 11).value = "HOUSE BL"
                                objExcel.Cells(iRow, iCol + 12).value = "BOOKING ID"
                                objExcel.Cells(iRow, iCol + 13).value = "SHIPPER"
                                objExcel.Cells(iRow, iCol + 14).value = "CNEE"
                                objExcel.Cells(iRow, iCol + 15).value = "CTNR SIZE"
                                objExcel.Cells(iRow, iCol + 16).value = "FEUS"
                                objExcel.Cells(iRow, iCol + 17).value = "CBM"
                                objExcel.Cells(iRow, iCol + 18).value = "HANDLE BY"

                                'setting border
                                objExcel.Range(objExcel.Cells(iRow, iCol + 3), objExcel.Cells(iRow, iCol + 3)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, iCol + 9), objExcel.Cells(iRow, iCol + 9)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 18)).Borders(9).LineStyle = 1
                                objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).HorizontalAlignment = -4108
                                objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow, iCol + 5)).Merge()
                                objExcel.Range(objExcel.Cells(iSRow + 1, iCol), objExcel.Cells(iSRow + 1, iCol + 5)).Merge()
                                objExcel.Range(objExcel.Cells(iSRow + 3, iCol), objExcel.Cells(iSRow + 3, iCol + 5)).Merge()

                                iRow = iRow + 1
                                iSRow = iRow

                                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                                    objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ShhLotNo")
                                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("ShhMBLNo")

                                    objExcel.Cells(iRow, iCol + 2).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")
                                    Total1 = Total1 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")

                                    objExcel.Cells(iRow, iCol + 3).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")
                                    Total2 = Total2 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                                    objExcel.Cells(iRow, iCol + 4).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2"), "2")
                                    Total3 = Total3 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), 2) - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                                    If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                        objExcel.Cells(iRow, iCol + 5).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                                        Total4 = Total4 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                                    Else
                                        objExcel.Cells(iRow, iCol + 5).value = 0
                                    End If

                                    ' Account Rate
                                    objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")
                                    Total5 = Total5 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")

                                    objExcel.Cells(iRow, iCol + 7).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")
                                    Total6 = Total6 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                                    objExcel.Cells(iRow, iCol + 8).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2"), "2")
                                    Total7 = Total7 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), 2) - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                                    If ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") <> 0 Then
                                        objExcel.Cells(iRow, iCol + 9).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                                        Total8 = Total8 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                                    Else
                                        objExcel.Cells(iRow, iCol + 9).value = 0
                                    End If

                                    objExcel.Cells(iRow, iCol + 10).value = ds.Tables(TblIndex).Rows(i).Item("AccRate")
                                    objExcel.Cells(iRow, iCol + 11).value = ds.Tables(TblIndex).Rows(i).Item("ShhBLNo")
                                    objExcel.Cells(iRow, iCol + 12).value = ds.Tables(TblIndex).Rows(i).Item("ShhSoNo")
                                    objExcel.Cells(iRow, iCol + 13).value = ds.Tables(TblIndex).Rows(i).Item("ShhShipper")
                                    objExcel.Cells(iRow, iCol + 14).value = ds.Tables(TblIndex).Rows(i).Item("ShhConsignee")
                                    objExcel.Cells(iRow, iCol + 15).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrType")
                                    objExcel.Cells(iRow, iCol + 16).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrFEUS")
                                    objExcel.Cells(iRow, iCol + 17).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrCBM")
                                    objExcel.Cells(iRow, iCol + 18).value = ds.Tables(TblIndex).Rows(i).Item("HandleBy")

                                    iRow = iRow + 1
                                Next

                                iRow = iRow + 2

                                objExcel.Cells(iRow, iCol + 1).value = "GRAND TOTAL :"

                                For i = 3 To 5
                                    objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                                    objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                                    objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                                Next

                                ' Account Rate
                                For i = 7 To 9
                                    objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                                    objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                                    objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                                Next

                                objExcel.Range(objExcel.Cells(iRow, 6), objExcel.Cells(iRow, 6)).Activate()
                                objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                                ' Account Rate
                                objExcel.Range(objExcel.Cells(iRow, 10), objExcel.Cells(iRow, 10)).Activate()
                                objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,##0.00_ "
                                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Borders(8).LineStyle = 1

                                ' Account Rate
                                objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 9)).NumberFormatLocal = "#,##0.00_ "
                                objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 9)).Borders(8).LineStyle = 1

                                iRow = iRow + 3
                            End While
                        End If
                    Else
                        iSRow = iRow

                        TblIndex = 2
                        i = 0

                        objExcel.Cells(iRow, 1).value = "LOT GP SUMMARY REPORT BY " & ds.Tables(TblIndex).Rows(i).Item("title")

                        If PrintType = 6 Then
                            If ds.Tables(TblIndex).Rows(i).Item("ReportBy") = "P" Then
                                objExcel.Cells(iRow + 2, 1).value = "PERIOD (ETD): " & Format(ds.Tables(TblIndex).Rows(i).Item("DteFrm"), "dd/mm/yy") & "-" & Format(ds.Tables(TblIndex).Rows(i).Item("DteTo"), "dd/mm/yy")
                            Else
                                If ds.Tables(TblIndex).Rows(i).Item("MonthNo") <> 0 Then
                                    objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " MONTH " & common.DigitToMonth(ds.Tables(TblIndex).Rows(i).Item("MonthNo"))
                                Else
                                    objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " WEEK " & ds.Tables(TblIndex).Rows(i).Item("WeekNo")
                                End If
                            End If
                        Else
                            objExcel.Cells(iRow + 2, 1).value = "VESSEL/VOY: " & ds.Tables(TblIndex).Rows(i).Item("VslName") & "      CARRIER: " & ds.Tables(TblIndex).Rows(i).Item("CarrName") & "      DATE: " & Format(ds.Tables(TblIndex).Rows(i).Item("VslOnBoard"), "dd/mm/yy")
                        End If

                        iRow = iRow + 4

                        'Upper part title
                        objExcel.Cells(iRow, iCol).value = "LOT"
                        objExcel.Cells(iRow, iCol + 1).value = "MASTER BL"
                        objExcel.Cells(iRow, iCol + 2).value = "INCOME"
                        objExcel.Cells(iRow, iCol + 3).value = "COST"
                        objExcel.Cells(iRow, iCol + 4).value = "GP"
                        objExcel.Cells(iRow, iCol + 5).value = "GP MARGIN(%)"
                        objExcel.Cells(iRow, iCol + 6).value = "INCOME" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 7).value = "COST" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 8).value = "GP" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 9).value = "GP MARGIN(%)" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 10).value = "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 11).value = "HOUSE BL"
                        objExcel.Cells(iRow, iCol + 12).value = "BOOKING ID"
                        objExcel.Cells(iRow, iCol + 13).value = "SHIPPER"
                        objExcel.Cells(iRow, iCol + 14).value = "CNEE"
                        objExcel.Cells(iRow, iCol + 15).value = "CTNR SIZE"
                        objExcel.Cells(iRow, iCol + 16).value = "FEUS"
                        objExcel.Cells(iRow, iCol + 17).value = "CBM"
                        objExcel.Cells(iRow, iCol + 18).value = "HANDLE BY"

                        'setting border
                        objExcel.Range(objExcel.Cells(iRow, iCol + 3), objExcel.Cells(iRow, iCol + 3)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, iCol + 9), objExcel.Cells(iRow, iCol + 9)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 18)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow, iCol + 18)).Merge()
                        objExcel.Range(objExcel.Cells(iSRow + 1, iCol), objExcel.Cells(iSRow + 1, iCol + 5)).Merge()
                        objExcel.Range(objExcel.Cells(iSRow + 3, iCol), objExcel.Cells(iSRow + 3, iCol + 5)).Merge()

                        iRow = iRow + 1
                        iSRow = iRow

                        While i < ds.Tables(TblIndex).Rows.Count
                            objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ShhLotNo")
                            objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("ShhMBLNo")
                            objExcel.Cells(iRow, iCol + 2).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")
                            Total1 = Total1 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")

                            objExcel.Cells(iRow, iCol + 3).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")
                            Total2 = Total2 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                            objExcel.Cells(iRow, iCol + 4).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2"), "2")
                            Total3 = Total3 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                            If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                objExcel.Cells(iRow, iCol + 5).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                                Total4 = Total4 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                            Else
                                objExcel.Cells(iRow, iCol + 5).value = 0
                            End If

                            ' Account Rate
                            objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")
                            Total5 = Total5 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")

                            objExcel.Cells(iRow, iCol + 7).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")
                            Total6 = Total6 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                            objExcel.Cells(iRow, iCol + 8).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2"), "2")
                            Total7 = Total7 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                            If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                objExcel.Cells(iRow, iCol + 9).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                                Total8 = Total8 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                            Else
                                objExcel.Cells(iRow, iCol + 9).value = 0
                            End If

                            objExcel.Cells(iRow, iCol + 10).value = ds.Tables(TblIndex).Rows(i).Item("AccRate")
                            objExcel.Cells(iRow, iCol + 11).value = ds.Tables(TblIndex).Rows(i).Item("ShhBLNo")
                            objExcel.Cells(iRow, iCol + 12).value = ds.Tables(TblIndex).Rows(i).Item("ShhSoNo")
                            objExcel.Cells(iRow, iCol + 13).value = ds.Tables(TblIndex).Rows(i).Item("ShhShipper")
                            objExcel.Cells(iRow, iCol + 14).value = ds.Tables(TblIndex).Rows(i).Item("ShhConsignee")
                            objExcel.Cells(iRow, iCol + 15).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrType")
                            objExcel.Cells(iRow, iCol + 16).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrFEUS")
                            objExcel.Cells(iRow, iCol + 17).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrCBM")
                            objExcel.Cells(iRow, iCol + 18).value = ds.Tables(TblIndex).Rows(i).Item("HandleBy")

                            i = i + 1
                            iRow = iRow + 1
                        End While

                        iRow = iRow + 2

                        objExcel.Cells(iRow, iCol + 1).value = "GRAND TOTAL :"

                        For i = 3 To 5
                            objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                            objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                            objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        Next

                        ' Account Rate
                        For i = 7 To 9
                            objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                            objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                            objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        Next

                        objExcel.Range(objExcel.Cells(iRow, 6), objExcel.Cells(iRow, 6)).Activate()
                        objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                        ' Account Rate
                        objExcel.Range(objExcel.Cells(iRow, 10), objExcel.Cells(iRow, 10)).Activate()
                        objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                        objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,##0.00_ "
                        objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Borders(8).LineStyle = 1

                        ' Account Rate
                        objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 10)).NumberFormatLocal = "#,##0.00_ "
                        objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 10)).Borders(8).LineStyle = 1
                    End If
                Else ' Not type 6 or 10 or 11 or 12 or 13
                    If TypeRefId = 0 Then
                        i = 0

                        If ds.Tables(0).Rows.Count > 0 And ds.Tables(1).Rows.Count > 0 Then
                            TblIndex = 2

                            While i < ds.Tables(TblIndex).Rows.Count
                                iSRow = iRow
                                objExcel.Cells(iRow, 1).value = "LOT GP SUMMARY REPORT BY " & ds.Tables(TblIndex).Rows(i).Item("title")
                                If ds.Tables(TblIndex).Rows(i).Item("ReportBy") = "P" Then
                                    objExcel.Cells(iRow + 2, 1).value = "PERIOD (ETD): " & Format(ds.Tables(TblIndex).Rows(i).Item("DteFrm"), "dd/mm/yy") & "-" & Format(ds.Tables(TblIndex).Rows(i).Item("DteTo"), "dd/mm/yy")
                                Else
                                    If ds.Tables(TblIndex).Rows(i).Item("MonthNo") <> 0 Then
                                        objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " MONTH " & common.DigitToMonth(ds.Tables(TblIndex).Rows(i).Item("MonthNo"))
                                    Else
                                        objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " WEEK " & ds.Tables(TblIndex).Rows(i).Item("WeekNo")
                                    End If
                                End If
                                objExcel.Cells(iRow + 3, 1).value = ds.Tables(TblIndex).Rows(i).Item("title") & " : " & ds.Tables(TblIndex).Rows(i).Item("cName")

                                iRow = iRow + 5

                                ' Upper part title
                                objExcel.Cells(iRow, iCol).value = "LOT"
                                objExcel.Cells(iRow, iCol + 1).value = "MASTER BL"
                                objExcel.Cells(iRow, iCol + 2).value = "INCOME"
                                objExcel.Cells(iRow, iCol + 3).value = "COST"
                                objExcel.Cells(iRow, iCol + 4).value = "GP"
                                objExcel.Cells(iRow, iCol + 5).value = "GP MARGIN(%)"
                                objExcel.Cells(iRow, iCol + 6).value = "INCOME" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 7).value = "COST" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 8).value = "GP" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 9).value = "GP MARGIN(%)" & Chr(10) & Chr(13) & "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 10).value = "A/C EX-RATE"
                                objExcel.Cells(iRow, iCol + 11).value = "HOUSE BL"
                                objExcel.Cells(iRow, iCol + 12).value = "BOOKING ID"
                                objExcel.Cells(iRow, iCol + 13).value = "SHIPPER"
                                objExcel.Cells(iRow, iCol + 14).value = "CNEE"
                                objExcel.Cells(iRow, iCol + 15).value = "CTNR SIZE"
                                objExcel.Cells(iRow, iCol + 16).value = "FEUS"
                                objExcel.Cells(iRow, iCol + 17).value = "CBM"
                                objExcel.Cells(iRow, iCol + 18).value = "HANDLE BY"

                                ' setting border
                                objExcel.Range(objExcel.Cells(iRow, iCol + 3), objExcel.Cells(iRow, iCol + 3)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, iCol + 9), objExcel.Cells(iRow, iCol + 9)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 11)).Borders(9).LineStyle = 1
                                objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).Font.Bold = True
                                objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).HorizontalAlignment = -4108
                                objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow, iCol + 18)).Merge()
                                objExcel.Range(objExcel.Cells(iSRow + 1, iCol), objExcel.Cells(iSRow + 1, iCol + 5)).Merge()
                                objExcel.Range(objExcel.Cells(iSRow + 3, iCol), objExcel.Cells(iSRow + 3, iCol + 5)).Merge()

                                iRow = iRow + 1
                                iSRow = iRow

                                For i = 0 To ds.Tables(TblIndex).Rows.Count - 1
                                    objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ShhLotNo")
                                    objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("ShhMBLNo")
                                    objExcel.Cells(iRow, iCol + 2).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")
                                    Total1 = Total1 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")

                                    objExcel.Cells(iRow, iCol + 3).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")
                                    Total2 = Total2 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                                    objExcel.Cells(iRow, iCol + 4).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2"), "2")
                                    Total3 = Total3 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                                    If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                        objExcel.Cells(iRow, iCol + 5).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                                        Total4 = Total4 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                                    Else
                                        objExcel.Cells(iRow, iCol + 5).value = 0
                                    End If

                                    ' Account Rate
                                    objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")
                                    Total5 = Total5 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")

                                    objExcel.Cells(iRow, iCol + 7).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")
                                    Total6 = Total6 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                                    objExcel.Cells(iRow, iCol + 8).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2"), "2")
                                    Total7 = Total7 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                                    If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                        objExcel.Cells(iRow, iCol + 9).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                                        Total8 = Total8 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                                    Else
                                        objExcel.Cells(iRow, iCol + 9).value = 0
                                    End If

                                    objExcel.Cells(iRow, iCol + 10).value = ds.Tables(TblIndex).Rows(i).Item("AccRate")
                                    objExcel.Cells(iRow, iCol + 11).value = ds.Tables(TblIndex).Rows(i).Item("ShhBLNo")
                                    objExcel.Cells(iRow, iCol + 12).value = ds.Tables(TblIndex).Rows(i).Item("ShhSoNo")
                                    objExcel.Cells(iRow, iCol + 13).value = ds.Tables(TblIndex).Rows(i).Item("ShhShipper")
                                    objExcel.Cells(iRow, iCol + 14).value = ds.Tables(TblIndex).Rows(i).Item("ShhConsignee")
                                    objExcel.Cells(iRow, iCol + 15).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrType")
                                    objExcel.Cells(iRow, iCol + 16).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrFEUS")
                                    objExcel.Cells(iRow, iCol + 17).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrCBM")
                                    objExcel.Cells(iRow, iCol + 18).value = ds.Tables(TblIndex).Rows(i).Item("HandleBy")

                                    iRow = iRow + 1
                                    i = i + 1
                                Next

                                iRow = iRow + 2

                                objExcel.Cells(iRow, iCol + 1).value = "GRAND TOTAL :"

                                For i = 3 To 5
                                    objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                                    objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                                    objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                                Next

                                ' Account Rate
                                For i = 7 To 9
                                    objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                                    objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                                    objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                                Next

                                objExcel.Range(objExcel.Cells(iRow, 6), objExcel.Cells(iRow, 6)).Activate()
                                objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                                ' Account Rate
                                objExcel.Range(objExcel.Cells(iRow, 10), objExcel.Cells(iRow, 10)).Activate()
                                objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,##0.00_ "
                                objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Borders(8).LineStyle = 1

                                ' Account Rate
                                objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 9)).NumberFormatLocal = "#,##0.00_ "
                                objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 9)).Borders(8).LineStyle = 1

                                iRow = iRow + 3
                            End While
                        End If
                    Else
                        iSRow = iRow

                        TblIndex = 2
                        i = 0

                        objExcel.Cells(iRow, 1).value = "LOT GP SUMMARY REPORT BY " & ds.Tables(TblIndex).Rows(i).Item("title")

                        If ds.Tables(TblIndex).Rows(i).Item("ReportBy") = "P" Then
                            objExcel.Cells(iRow + 2, 1).value = "PERIOD (ETD): " & Format(ds.Tables(TblIndex).Rows(i).Item("DteFrm"), "dd/mm/yy") & "-" & Format(ds.Tables(TblIndex).Rows(i).Item("DteTo"), "dd/mm/yy")
                        Else
                            If ds.Tables(TblIndex).Rows(i).Item("MonthNo") <> 0 Then
                                objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " MONTH " & common.DigitToMonth(ds.Tables(TblIndex).Rows(i).Item("MonthNo"))
                            Else
                                objExcel.Cells(iRow + 2, 1).value = "FOR THE YEAR " & ds.Tables(TblIndex).Rows(i).Item("YearNo") & " WEEK " & ds.Tables(TblIndex).Rows(i).Item("WeekNo")
                            End If
                        End If
                        objExcel.Cells(iRow + 3, 1).value = ds.Tables(TblIndex).Rows(i).Item("title") & " : " & ds.Tables(TblIndex).Rows(i).Item("cName")

                        iRow = iRow + 5

                        ' Upper part title
                        objExcel.Cells(iRow, iCol).value = "LOT"
                        objExcel.Cells(iRow, iCol + 1).value = "MASTER BL"
                        objExcel.Cells(iRow, iCol + 2).value = "INCOME"
                        objExcel.Cells(iRow, iCol + 3).value = "COST"
                        objExcel.Cells(iRow, iCol + 4).value = "GP"
                        objExcel.Cells(iRow, iCol + 5).value = "GP MARGIN(%)"
                        objExcel.Cells(iRow, iCol + 6).value = "INCOME" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 7).value = "COST" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 8).value = "GP" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 9).value = "GP MARGIN(%)" & Chr(10) & Chr(13) & "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 10).value = "A/C EX-RATE"
                        objExcel.Cells(iRow, iCol + 11).value = "HOUSE BL"
                        objExcel.Cells(iRow, iCol + 12).value = "BOOKING ID"
                        objExcel.Cells(iRow, iCol + 13).value = "SHIPPER"
                        objExcel.Cells(iRow, iCol + 14).value = "CNEE"
                        objExcel.Cells(iRow, iCol + 15).value = "CTNR SIZE"
                        objExcel.Cells(iRow, iCol + 16).value = "FEUS"
                        objExcel.Cells(iRow, iCol + 17).value = "CBM"
                        objExcel.Cells(iRow, iCol + 18).value = "HANDLE BY"

                        ' setting border
                        objExcel.Range(objExcel.Cells(iRow, iCol + 3), objExcel.Cells(iRow, iCol + 3)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, iCol + 5), objExcel.Cells(iRow, iCol + 5)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, iCol + 7), objExcel.Cells(iRow, iCol + 7)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, iCol + 9), objExcel.Cells(iRow, iCol + 9)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iRow, 1), objExcel.Cells(iRow, iCol + 18)).Borders(9).LineStyle = 1
                        objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).Font.Bold = True
                        objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow + 1, iCol + 18)).HorizontalAlignment = -4108
                        objExcel.Range(objExcel.Cells(iSRow, iCol), objExcel.Cells(iSRow, iCol + 6)).Merge()
                        objExcel.Range(objExcel.Cells(iSRow + 1, iCol), objExcel.Cells(iSRow + 1, iCol + 5)).Merge()
                        objExcel.Range(objExcel.Cells(iSRow + 3, iCol), objExcel.Cells(iSRow + 3, iCol + 5)).Merge()

                        iRow = iRow + 1
                        iSRow = iRow

                        While i < ds.Tables(TblIndex).Rows.Count
                            objExcel.Cells(iRow, iCol).value = ds.Tables(TblIndex).Rows(i).Item("ShhLotNo")
                            objExcel.Cells(iRow, iCol + 1).value = ds.Tables(TblIndex).Rows(i).Item("ShhMBLNo")
                            objExcel.Cells(iRow, iCol + 2).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")
                            Total1 = Total1 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2")

                            objExcel.Cells(iRow, iCol + 3).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")
                            Total2 = Total2 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                            objExcel.Cells(iRow, iCol + 4).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2"), "2")
                            Total3 = Total3 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncome"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlay"), "2")

                            If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                objExcel.Cells(iRow, iCol + 5).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                                Total4 = Total4 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncome") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlay")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncome") * 100, 2)
                            Else
                                objExcel.Cells(iRow, iCol + 5).value = 0
                            End If

                            ' Account Rate
                            objExcel.Cells(iRow, iCol + 6).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")
                            Total5 = Total5 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2")

                            objExcel.Cells(iRow, iCol + 7).value = FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")
                            Total6 = Total6 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                            objExcel.Cells(iRow, iCol + 8).value = FormatNumber(FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2"), "2")
                            Total7 = Total7 + FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate"), "2") - FormatNumber(ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate"), "2")

                            If ds.Tables(TblIndex).Rows(i).Item("ShhIncome") <> 0 Then
                                objExcel.Cells(iRow, iCol + 9).value = FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                                Total8 = Total8 + FormatNumber((ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") - ds.Tables(TblIndex).Rows(i).Item("ShhOutlayAccRate")) / ds.Tables(TblIndex).Rows(i).Item("ShhIncomeAccRate") * 100, 2)
                            Else
                                objExcel.Cells(iRow, iCol + 9).value = 0
                            End If

                            objExcel.Cells(iRow, iCol + 10).value = ds.Tables(TblIndex).Rows(i).Item("AccRate")
                            objExcel.Cells(iRow, iCol + 11).value = ds.Tables(TblIndex).Rows(i).Item("ShhBLNo")
                            objExcel.Cells(iRow, iCol + 12).value = ds.Tables(TblIndex).Rows(i).Item("ShhSoNo")
                            objExcel.Cells(iRow, iCol + 13).value = ds.Tables(TblIndex).Rows(i).Item("ShhShipper")
                            objExcel.Cells(iRow, iCol + 14).value = ds.Tables(TblIndex).Rows(i).Item("ShhConsignee")
                            objExcel.Cells(iRow, iCol + 15).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrType")
                            objExcel.Cells(iRow, iCol + 16).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrFEUS")
                            objExcel.Cells(iRow, iCol + 17).value = ds.Tables(TblIndex).Rows(i).Item("ShhCtnrCBM")
                            objExcel.Cells(iRow, iCol + 18).value = ds.Tables(TblIndex).Rows(i).Item("HandleBy")

                            iRow = iRow + 1
                            i = i + 1
                        End While

                        iRow = iRow + 2
                        objExcel.Cells(iRow, iCol + 1).value = "GRAND TOTAL :"

                        For i = 3 To 5
                            objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                            objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                            objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        Next

                        ' Account Rate
                        For i = 7 To 9
                            objExcel.Range(objExcel.Cells(iSRow, i), objExcel.Cells(iRow - 2, i)).Select()
                            objExcel.Range(objExcel.Cells(iRow, i), objExcel.Cells(iRow, i)).Activate()
                            objExcel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & (iRow - iSRow) & "]C:R[-1]C)"
                        Next

                        objExcel.Range(objExcel.Cells(iRow, 6), objExcel.Cells(iRow, 6)).Activate()
                        objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                        ' Account Rate
                        objExcel.Range(objExcel.Cells(iRow, 10), objExcel.Cells(iRow, 10)).Activate()
                        objExcel.ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]*100"

                        objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).NumberFormatLocal = "#,##0.00_ "
                        objExcel.Range(objExcel.Cells(iRow, iCol + 2), objExcel.Cells(iRow, iCol + 5)).Borders(8).LineStyle = 1

                        ' Account Rate
                        objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 9)).NumberFormatLocal = "#,##0.00_ "
                        objExcel.Range(objExcel.Cells(iRow, iCol + 6), objExcel.Cells(iRow, iCol + 9)).Borders(8).LineStyle = 1
                    End If
                End If
            End If

            ' setting width
            objExcel.Columns("A:A").ColumnWidth = 20
            objExcel.Columns("B:B").ColumnWidth = 20
            objExcel.Columns("C:R").ColumnWidth = 14

            ' ----------------------------------------------------------------------
            ' Save File
            ' ----------------------------------------------------------------------
            Dim exportPath As String = My.Settings.ExportPath & sUid & "\"
            Dim exportFile As String = ""

            If fileName <> "" Then
                exportFile = exportPath & fileName & ".xls"
            Else
                exportFile = exportPath & Uid & ".xls"
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
        RptLotGP = fileName

    End Function

End Class
