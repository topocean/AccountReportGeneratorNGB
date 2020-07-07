Public Class ClsExcel

    Public rptFile As String

    Private RptNoData As Boolean = False

    Property rptFileName()

        Get
            rptFileName = Me.rptFile
        End Get
        Set(ByVal value)
            Me.rptFile = value
        End Set

    End Property

    Sub ExportExcel(ByVal UID As String, ByVal sUID As String, ByVal RptID As String, ByVal RptName As String, ByVal RptDataSet As DataSet, ByVal isEmail As Integer, ByVal UsrDtl As String(), ByVal inParaNames As String(), ByVal inParaValues As String())

        Dim rptFile As String = ""
        Dim cn As String = ""
        Dim paraNames(), paraValues() As String
        Dim paraNameStr As String = ""
        Dim paraValueStr As String = ""
        Dim tmpParaVal As String = ""
        Dim tmpParaName As String = ""
        Dim startTime As Date
        Dim i As Integer
        Dim common As New common
        Dim hasError As Boolean = False
        Dim errMsg As String = ""
        Dim errArray() As String

        ' ------------------------------------------------------------
        ' Connection String
        ' ------------------------------------------------------------
        cn &= "Data Source=" & My.Settings.Server & ";"
        cn &= "Database=" & My.Settings.DB & ";"
        cn &= "User Id=" & My.Settings.Login & ";"
        cn &= "Password=" & My.Settings.Password & ";"

        Dim sqlConn As New Data.SqlClient.SqlConnection(cn)
        Dim cmd As New Data.SqlClient.SqlCommand
        Dim sql As String

        ' ------------------------------------------------------------
        ' Open Database Connection
        ' ------------------------------------------------------------
        sqlConn.Open()
        cmd = sqlConn.CreateCommand

        'sql = "CALL usp_PrintJob_Export2File('" & UID & "');"
        sql = "UPDATE PdfReport SET Status = 10, LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"
        cmd.CommandText = sql
        cmd.CommandTimeout = My.Settings.Timeout
        cmd.ExecuteNonQuery()

        common.SaveLog("Generating Report: " & RptID & " UID: " & UID)
        frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Generating Report: " & RptID & " (UID: " & UID & ")")
        frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1

        startTime = frmMain.ProcTime
        rptFile = ""
        paraNames = Split("", ",")
        paraValues = Split("", ",")

        Select Case RptID
            ' ----------------------------------------------------------------------------
            ' Loading Report (Yearly / Monthly)
            ' ----------------------------------------------------------------------------
            Case "RptLoading"
                Dim rpt As New RptLoading
                rptFile = rpt.RptLoading(UID, RptDataSet)
                rptName = "Loading Report"

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "InWeek"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = ""
                                        tmpParaVal = ""
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "InMonth"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = ""
                                        tmpParaVal = ""
                                    Else
                                        tmpParaName = "Month"
                                        tmpParaVal = common.GetMonthName(inParaValues(i))
                                    End If

                                Case "InSubBrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetSubBranchName(inParaValues(i), My.Settings.GenID)

                                Case "InBrhCd"
                                    tmpParaName = ""
                                    tmpParaVal = ""

                                Case "InLocation"
                                    tmpParaName = ""
                                    tmpParaVal = ""

                                Case "InTraffic"
                                    tmpParaName = "Traffic"
                                    tmpParaVal = common.GetTrafficName(inParaValues(i))

                                Case "InPOType"
                                    tmpParaName = "PO Type"
                                    tmpParaVal = common.GetPOType(inParaValues(i))

                                Case "InMode"
                                    tmpParaName = "Mode"
                                    If Trim(inParaValues(i)) = "" Then
                                        tmpParaVal = "ALL"
                                    Else
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "InOCF"
                                    tmpParaName = "OCF"
                                    If Trim(inParaValues(i)) = "" Then
                                        tmpParaVal = "ALL"
                                    Else
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next

                    ' Put Parameter data into array
                    paraNames = Split(paraNameStr, ",")
                    paraValues = Split(paraValueStr, ",")
                End If

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Loading Report (Yearly / Monthly)
                ' ----------------------------------------------------------------------------

            Case "RptLoading_Carrier"
                ' ----------------------------------------------------------------------------
                ' Yearly Loading Report of Carriers
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptLoading_Carrier
                rptFile = rpt.RptLoading_Carrier(UID, RptDataSet)
                rptName = "Loading Report (Carrier)"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Yearly Loading Report of Carriers
                ' ----------------------------------------------------------------------------

            Case "RptLoading_Monthly"
                ' ----------------------------------------------------------------------------
                ' Yearly / Monthly Loading Report
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptLoading_Monthly
                rptFile = rpt.RptLoading_Monthly(UID, RptDataSet)
                rptName = "Loading Report"

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@Month"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Month"
                                        tmpParaVal = common.GetMonthName(inParaValues(i))
                                    End If

                                Case "@SubBrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetSubBranchName(inParaValues(i), My.Settings.GenID)

                                Case "@BrhCd"
                                    tmpParaName = " "
                                    tmpParaVal = " "

                                Case "@Location"
                                    tmpParaName = " "
                                    tmpParaVal = " "

                                Case "@Traffic"
                                    tmpParaName = "Traffic"
                                    tmpParaVal = common.GetTrafficName(inParaValues(i))

                                Case "@POType"
                                    tmpParaName = "PO Type"
                                    tmpParaVal = common.GetPOType(inParaValues(i))

                                Case "@BkhMode"
                                    tmpParaName = "Mode"
                                    If Trim(inParaValues(i)) = "" Then
                                        tmpParaVal = "ALL"
                                    Else
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@OCF"
                                    tmpParaName = "OCF"
                                    If Trim(inParaValues(i)) = "" Then
                                        tmpParaVal = "ALL"
                                    Else
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If

                        End If
                    Next
                End If

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Yearly / Monthly Loading Report
                ' ----------------------------------------------------------------------------

            Case "RptLoading_Monthly_THC"
                ' ----------------------------------------------------------------------------
                ' Yearly / Monthly Loading Report (THC)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptLoading_Monthly_THC
                rptFile = rpt.RptLoading_Monthly_THC(UID, RptDataSet)
                RptName = "Loading Report"

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@Month"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Month"
                                        tmpParaVal = common.GetMonthName(inParaValues(i))
                                    End If

                                Case "@SubBrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetSubBranchName(inParaValues(i), My.Settings.GenID)

                                Case "@BrhCd"
                                    tmpParaName = " "
                                    tmpParaVal = " "

                                Case "@Location"
                                    tmpParaName = " "
                                    tmpParaVal = " "

                                Case "@Traffic"
                                    tmpParaName = "Traffic"
                                    tmpParaVal = common.GetTrafficName(inParaValues(i))

                                Case "@POType"
                                    tmpParaName = "PO Type"
                                    tmpParaVal = common.GetPOType(inParaValues(i))

                                Case "@BkhMode"
                                    tmpParaName = "Mode"
                                    If Trim(inParaValues(i)) = "" Then
                                        tmpParaVal = "ALL"
                                    Else
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@OCF"
                                    tmpParaName = "OCF"
                                    If Trim(inParaValues(i)) = "" Then
                                        tmpParaVal = "ALL"
                                    Else
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Yearly / Monthly Loading Report
                ' ----------------------------------------------------------------------------

            Case "RptAMSSummaryVessel"
                ' ----------------------------------------------------------------------------
                ' AMS Summary (Vessel List)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptAMSSummaryByVessel
                rptFile = rpt.RptAMSSummaryByVessel(UID, RptDataSet)
                rptName = "AMS Summary (Vessel List)"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of AMS Summary (Vessel List)
                ' ----------------------------------------------------------------------------

            Case "RptBooking"
                ' ----------------------------------------------------------------------------
                ' Booking Report (CY)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptBooking
                rptFile = rpt.RptBooking(UID, RptDataSet)
                rptName = "Booking Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Booking Report (CY)
                ' ----------------------------------------------------------------------------

            Case "RptBookingCFS"
                ' ----------------------------------------------------------------------------
                ' Booking Report (CFS)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptBookingCFS
                rptFile = rpt.RptBookingCFS(UID, RptDataSet)
                rptName = "Booking Report CFS"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Booking Report (CFS)
                ' ----------------------------------------------------------------------------

            Case "RptBSBookingReport"
                ' ----------------------------------------------------------------------------
                ' Brookstone Booking Report
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptBSBookingReport
                rptFile = rpt.RptBSBookingReport(UID, RptDataSet)
                rptName = "Brookstone Booking Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Brookstone Booking Report
                ' ----------------------------------------------------------------------------

            Case "RptBSShippingAdvice"
                ' ----------------------------------------------------------------------------
                ' Brookstone Shipping Advice
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptBSShippingAdvice
                rptFile = rpt.RptBSShippingAdvice(UID, RptDataSet)
                rptName = "Brookstone Shipping Advice"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Brookstone Shipping Advice
                ' ----------------------------------------------------------------------------

            Case "Rpt11D"
                ' ----------------------------------------------------------------------------
                ' 11-D Report
                ' ----------------------------------------------------------------------------

                Dim rpt As New Rpt11D
                rptFile = rpt.Rpt11D(UID, RptDataSet)
                rptName = "11-D Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of 11-D Report
                ' ----------------------------------------------------------------------------

            Case "RptPOTracking"
                ' ----------------------------------------------------------------------------
                ' PO Tracking Report
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptPOTracking
                rptFile = rpt.RptPOTracking(UID, RptDataSet)
                rptName = "PO Tracking Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of PO Tracking Report
                ' ----------------------------------------------------------------------------

            Case "RptISF"
                ' ----------------------------------------------------------------------------
                ' ISF Report
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptISF
                rptFile = rpt.RptISF(UID, RptDataSet)
                rptName = "ISF Booking Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of ISF Report
                ' ----------------------------------------------------------------------------

            Case "RptManifest"
                ' ----------------------------------------------------------------------------
                ' Manifest CFS
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptManifest
                rptFile = rpt.RptManifest(UID, RptDataSet)
                rptName = "Manifest CFS"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Manifest CFS
                ' ----------------------------------------------------------------------------

            Case "RptTopTen"
                ' ----------------------------------------------------------------------------
                ' Top Ten Report
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptTopTen
                rptFile = rpt.RptTopTen(UID, RptDataSet)
                rptName = "Top Ten Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Top Ten Report
                ' ----------------------------------------------------------------------------

            Case "RptLiftingSummary", "RptLiftingSummaryNUS"
                ' ----------------------------------------------------------------------------
                ' Lifting Summary
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptLiftingSummary

                If RptID = "RptLiftingSummaryNUS" Then
                    rptFile = rpt.RptLiftingSummaryNUS(UID, RptDataSet)
                Else
                    rptFile = rpt.RptLiftingSummary(UID, RptDataSet)
                End If

                rptName = "Lifting Summary"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Top Lifting Summary
                ' ----------------------------------------------------------------------------

            Case "RptWeekComparison"
                ' ----------------------------------------------------------------------------
                ' Week Comparison
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptWeekComparison
                rptFile = rpt.RptWeekComparison(UID, RptDataSet)
                rptName = "Week Comparison"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Week Comparison
                ' ----------------------------------------------------------------------------

            Case "RptTruckingSummary"
                ' ----------------------------------------------------------------------------
                ' Trucking Summary
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptTruckingSummary
                rptFile = rpt.RptTruckingSummary(UID, RptDataSet)
                rptName = "Trucking Summary"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Trucking Summary
                ' ----------------------------------------------------------------------------

            Case "RptDailyCashReceipt"
                ' ----------------------------------------------------------------------------
                ' Daily Cash Receipt
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptDailyCashReceipt
                rptFile = rpt.RptDailyCashReceipt(UID, RptDataSet)
                rptName = "Daily Cash Receipt"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Daily Cash Receipt
                ' ----------------------------------------------------------------------------

            Case "RptAccount"
                ' ----------------------------------------------------------------------------
                ' Account Interface
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptAccount
                rptFile = rpt.RptAccount(UID, RptDataSet)
                rptName = "Account Interface"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Account Interface
                ' ----------------------------------------------------------------------------

            Case "RptAccountHKG"
                ' ----------------------------------------------------------------------------
                ' Account Interface (Flex Account)
                ' ----------------------------------------------------------------------------
                Dim rpt As New RptAccountHKG

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@YearNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Year"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@BrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetBranchName(inParaValues(i))

                                Case "@SetType"
                                    tmpParaName = "Voucher Set"
                                    If Trim(inParaValues(i)) = "1" Then
                                        tmpParaVal = "Invoice"
                                    Else
                                        If Trim(inParaValues(i)) = "2" Then
                                            tmpParaVal = "Freight List"
                                        Else
                                            tmpParaVal = "Voucher"
                                        End If
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                rptFile = rpt.RptAccount(UID, sUID, RptDataSet)
                rptName = "Account Interface (Flex Account)"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Account Interface 
                ' ----------------------------------------------------------------------------

            Case "RptAccountHKG1"
                ' ----------------------------------------------------------------------------
                ' Account Interface (Flex Account)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptAccountHKG1

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@YearNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Year"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@BrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetBranchName(inParaValues(i))

                                Case "@SetType"
                                    tmpParaName = "Voucher Set"
                                    If Trim(inParaValues(i)) = "1" Then
                                        tmpParaVal = "Invoice"
                                    Else
                                        If Trim(inParaValues(i)) = "2" Then
                                            tmpParaVal = "Freight List"
                                        Else
                                            tmpParaVal = "Voucher"
                                        End If
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                rptFile = rpt.RptAccount1(UID, sUID, RptDataSet)
                rptName = "New Account Interface (Flex Account)"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Account Interface
                ' ----------------------------------------------------------------------------

            Case "RptAccountHKGNew"
                ' ----------------------------------------------------------------------------
                ' Account Interface (Flex Account)
                ' ----------------------------------------------------------------------------
                Dim rpt As New RptAccountHKGNew

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@YearNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Year"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@BrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetBranchName(inParaValues(i))

                                Case "@SetType"
                                    tmpParaName = "Voucher Set"
                                    If Trim(inParaValues(i)) = "1" Then
                                        tmpParaVal = "Invoice"
                                    Else
                                        If Trim(inParaValues(i)) = "2" Then
                                            tmpParaVal = "Freight List"
                                        Else
                                            tmpParaVal = "Voucher"
                                        End If
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                rptFile = rpt.RptAccount(UID, sUID, RptDataSet)

                If RptName = "" Then
                    RptName = "New Account Interface (Flex Account)"
                End If

                RptNoData = Not (rpt.rptHasDataOccur)

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Account Interface 
                ' ----------------------------------------------------------------------------

            Case "RptAccountHKGNew1"
                ' ----------------------------------------------------------------------------
                ' Account Interface (Flex Account)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptAccountHKG1New

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@YearNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Year"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@BrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetBranchName(inParaValues(i))

                                Case "@SetType"
                                    tmpParaName = "Voucher Set"
                                    If Trim(inParaValues(i)) = "1" Then
                                        tmpParaVal = "Invoice"
                                    Else
                                        If Trim(inParaValues(i)) = "2" Then
                                            tmpParaVal = "Freight List"
                                        Else
                                            tmpParaVal = "Voucher"
                                        End If
                                    End If

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                rptFile = rpt.RptAccount1(UID, sUID, RptDataSet)

                If RptName = "" Then
                    RptName = "New Account Interface (Flex Account)"
                End If

                RptNoData = Not (rpt.rptHasDataOccur)

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Account Interface
                ' ----------------------------------------------------------------------------

            Case "RptInvVou"
                ' ----------------------------------------------------------------------------
                ' Invoice / Voucher Summary
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptInvVou

                ' Require Email
                isEmail = 1

                rptFile = rpt.RptInvVou(UID, sUID, RptDataSet)
                rptName = "Invoice/Voucher Summary"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Account Interface RptInvVou
                ' ----------------------------------------------------------------------------

            Case "RptDebtor"
                ' ----------------------------------------------------------------------------
                ' Debtor Monthly Statement
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptDebtor
                rptFile = rpt.RptDebtor(UID, RptDataSet)
                rptName = "Debtor Monthly Statement"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Debtor Monthly Statement
                ' ----------------------------------------------------------------------------

            Case "RptInvNo", "RptInvNoByNumRange"
                ' ----------------------------------------------------------------------------
                ' Invoice Number Report 
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptInvNo

                If RptID = "RptInvNo" Then
                    rptFile = rpt.RptInvNo(UID, RptDataSet)
                Else
                    rptFile = rpt.RptInvNoByNumRange(UID, RptDataSet)
                End If

                rptName = "Invoice Number Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Invoice Number Report
                ' ----------------------------------------------------------------------------

            Case "RptLotGP"
                ' ----------------------------------------------------------------------------
                ' Lot GP Report
                ' ----------------------------------------------------------------------------
                Dim rpt As New RptLotGP

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@MonthNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Month"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@YearNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Year"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@BrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetBranchName(inParaValues(i))

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                rptFile = rpt.RptLotGP(UID, sUID, RptDataSet)
                rptName = "Lot GP Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing
                ' ----------------------------------------------------------------------------
                ' End of Lot GP Report
                ' ----------------------------------------------------------------------------

            Case "RptAirLotGP"
                ' ----------------------------------------------------------------------------
                ' Air Lot GP Report
                ' ----------------------------------------------------------------------------
                Dim rpt As New RptAirtLotGP

                ' Require Email
                isEmail = 1

                ' Convert search option parameters for displaying on emails
                If isEmail = 1 Then
                    For i = LBound(inParaNames) To UBound(inParaNames)
                        If Trim(inParaNames(i)) <> "" Then
                            Select Case inParaNames(i)
                                Case "@WeekNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Week"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@MonthNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Month"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@YearNo"
                                    If inParaValues(i) = 0 Then
                                        tmpParaName = " "
                                        tmpParaVal = " "
                                    Else
                                        tmpParaName = "Year"
                                        tmpParaVal = inParaValues(i)
                                    End If

                                Case "@BrhCd"
                                    tmpParaName = "Branch"
                                    tmpParaVal = common.GetBranchName(inParaValues(i))

                                Case Else
                                    tmpParaName = Replace(inParaNames(i), "@", "")
                                    tmpParaVal = inParaValues(i)
                            End Select

                            ' Padding Parameters Name
                            If tmpParaName <> "" Then
                                If paraNameStr = "" Then
                                    paraNameStr &= tmpParaName
                                Else
                                    paraNameStr &= "," & tmpParaName
                                End If
                            End If

                            ' Padding Parameters Value
                            If tmpParaVal <> "" Then
                                If paraValueStr = "" Then
                                    paraValueStr &= tmpParaVal
                                Else
                                    paraValueStr &= "," & tmpParaVal
                                End If
                            End If
                        End If
                    Next
                End If

                rptFile = rpt.RptAirLotGP(UID, sUID, RptDataSet)
                rptName = "Air Lot GP Report"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing
                ' ----------------------------------------------------------------------------
                ' End of Air Lot GP Report
                ' ----------------------------------------------------------------------------

            Case "RptTaxInvoiceSummary"
                ' ----------------------------------------------------------------------------
                ' Tax Invoice Summary (中國發票明細表)
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptTaxInvoiceSummary
                rptFile = rpt.RptTaxInvoiceSummary(UID, RptDataSet)
                rptName = "Tax Invoice Summary (中國發票明細表)"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Tax Invoice Summary (中國發票明細表)
                ' ----------------------------------------------------------------------------

            Case "RptUnpaidList"
                ' ----------------------------------------------------------------------------
                ' Outstanding Paid List
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptUnpaidList
                rptFile = rpt.RptUnpaidList(UID, RptDataSet)
                rptName = "Unpaid List"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Outstanding Paid List
                ' ----------------------------------------------------------------------------

            Case "RptAutoVoucher"
                ' ----------------------------------------------------------------------------
                ' Agent Voucher with Auto Gen Number List
                ' ----------------------------------------------------------------------------

                Dim rpt As New RptAutoVoucher

                ' Require Email
                isEmail = 1

                rptFile = rpt.RptAutoVoucher(UID, sUID, RptDataSet)
                rptName = "Agent Voucher with Auto-Gen Number List"

                ' Put Parameter data into array
                paraNames = Split(paraNameStr, ",")
                paraValues = Split(paraValueStr, ",")

                rpt = Nothing

                ' ----------------------------------------------------------------------------
                ' End of Agent Voucher with Auto Gen Number List
                ' ----------------------------------------------------------------------------

            Case Else

                hasError = True
                errMsg = "Requested report not found."
                rptFile = ""
                paraNames = Split("", ",")
                paraValues = Split("", ",")

        End Select

        errArray = Split(rptFile, ",")

        If rptFile = "" Or rptFile = ".xls" Then
            ' ------------------------------------------------------------
            ' Update Query Status
            ' ------------------------------------------------------------

            If hasError Then
                'sql = "CALL usp_PrintJob_Fail('" & UID & "', '" & common.setQuote(errMsg) & "');"
                sql = "UPDATE PdfReport SET Status = 11, Reason = '" & common.setQuote(errMsg) & "', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog("Exporting Report " & RptID & " - " & errMsg & ". (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Export Report Failed, please review the log.")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            Else
                'sql = "CALL usp_PrintJob_NoData('" & UID & "');"
                sql = "UPDATE PdfReport SET Status = 20, LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog("Exporting Report " & RptID & " - No Data " & DateDiff(DateInterval.Minute, startTime, Now) & "min(s) (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - No Data " & DateDiff(DateInterval.Minute, startTime, Now) & " min(s)")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            End If
        Else
            If LCase(errArray(0)) = "error" Then
                ' Update Query Status
                'sql = "CALL usp_PrintJob_Fail('" & UID & "', '" & common.setQuote(errArray(1)) & "');"
                sql = "UPDATE PdfReport SET Status = 11, Reason = '" & common.setQuote(errArray(1)) & "', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog(errArray(1), "E")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Exporting Report '" & RptID & "' Failure, Reason: " & errArray(1))
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            Else
                ' ------------------------------------------------------------
                ' Update Query Status
                ' ------------------------------------------------------------
                'sql = "CALL usp_PrintJob_Succ('" & UID & "', '" & rptFile & "');"
                sql = "UPDATE PdfReport SET Status = 6, URL = '" & rptFile & "', LstUpdDte = GETDATE() WHERE UID = '" & UID & "'"

                cmd.CommandText = sql
                cmd.CommandTimeout = My.Settings.Timeout
                cmd.ExecuteNonQuery()

                common.SaveLog("Exporting Report " & RptID & " Success in " & Replace(DateDiff(DateInterval.Minute, startTime, Now), 0, 1) & " min(s) (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Exporting Report " & RptID & " Success in " & Replace(DateDiff(DateInterval.Minute, startTime, Now), 0, 1) & " min(s)")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            End If
        End If

        ' ------------------------------------------------------------
        ' Close Database Connection
        ' ------------------------------------------------------------
        sqlConn.Close()

        ' ------------------------------------------------------------
        ' Send Email - when no group id
        ' ------------------------------------------------------------
        If sUID = "" Then
            If rptFile <> "" And isEmail = 1 Then
                common.SaveLog("Prepare sending email to requested user. (UID: " & UID & ")")
                frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Prepare sending email to requested user.")
                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1

                Dim clsMail As New ClsMailReport_old

                clsMail.MailReport(UID, UsrDtl, rptFile, rptName, paraNames, paraValues, RptNoData)
                clsMail = Nothing

                If rptFile = ".xls" Then
                    common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", ** No Data)")
                    frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (No Data)")
                Else
                    common.SaveLog("Report file has sent to requested user. (UID: " & UID & ", File Name: " & rptFile & ")")
                    frmMain.lstDisplay.Items.Add(Format(Now, "yyyy.MM.dd HH:mm:ss") & " - Report file has sent to requested user. (File Name: " & rptFile & ")")
                End If

                frmMain.lstDisplay.SelectedIndex = frmMain.lstDisplay.Items.Count - 1
            End If
        End If

        ' ------------------------------------------------------------
        ' Destroy Variables
        ' ------------------------------------------------------------
        sqlConn.Dispose()
        cmd.Dispose()

        cn = Nothing
        rptFile = Nothing
        startTime = Nothing
        sql = Nothing
        paraNames = Nothing
        paraValues = Nothing
        i = Nothing
        tmpParaVal = Nothing
        tmpParaName = Nothing
        hasError = Nothing
        errMsg = Nothing
        errArray = Nothing

        ' ------------------------------------------------------------
        ' Release Memory
        ' ------------------------------------------------------------
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

End Class
