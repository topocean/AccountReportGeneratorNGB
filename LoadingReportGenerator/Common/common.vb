Imports System.IO

Public Class common

    Sub SaveLog(ByVal msg As String, Optional ByVal type As String = "I")

        Try
            Dim logPath As String = My.Settings.LogPath
            Dim MsgHdr As String = ""

            ' Check Log Path Existance
            If Not My.Computer.FileSystem.DirectoryExists(logPath) Then
                My.Computer.FileSystem.CreateDirectory(logPath)
            End If

            Dim fso As New StreamWriter(logPath & Format(Now, "yyyy.MM.dd") & ".txt", True)

            If type = "I" Then
                MsgHdr = "Message:"
            Else
                MsgHdr = "Error:"
            End If

            fso.WriteLine("======================================================")
            fso.WriteLine("Date: " & Format(Now, "dd/MM/yyyy HH:mm:ss"))
            fso.WriteLine()
            fso.WriteLine(MsgHdr)
            fso.WriteLine()
            fso.WriteLine(msg)
            fso.WriteLine("======================================================")
            fso.WriteLine()

            ' Close File Object
            fso.Close()

            MsgHdr = Nothing
            fso.Dispose()
            fso = Nothing
            logPath = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            frmMain.lstDisplay.Items.Add(Format(Now, "dd.MM.yyyy HH:mm:ss") & " - Save Log errors captured, please review the error log.")
        End Try

    End Sub

    Function NullVal(ByVal inValue As Object, ByVal replacement As String) As String

        Dim tmp As String = ""

        If IsNothing(inValue) Then
            tmp = replacement
        ElseIf IsDBNull(inValue) Then
            tmp = replacement
        ElseIf CStr(inValue) = "" Then
            tmp = replacement
        Else
            tmp = inValue
        End If

        ' Return Value
        NullVal = tmp

        ' Destroy Variables
        tmp = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Function

    Function GetMonthName(ByVal inMonth As Integer) As String

        Dim tmp As String = ""

        Select Case inMonth
            Case 1
                tmp = "JAN (WEEK 01 - 05)"
            Case 2
                tmp = "FEB (WEEK 06 - 09)"
            Case 3
                tmp = "MAR (WEEK 10 - 13)"
            Case 4
                tmp = "APR (WEEK 14 - 18)"
            Case 5
                tmp = "MAY (WEEK 19 - 22)"
            Case 6
                tmp = "JUN (WEEK 23 - 26)"
            Case 7
                tmp = "JUL (WEEK 27 - 31)"
            Case 8
                tmp = "AUG (WEEK 32 - 35)"
            Case 9
                tmp = "SEP (WEEK 36 - 39)"
            Case 10
                tmp = "OCT (WEEK 40 - 44)"
            Case 11
                tmp = "NOV (WEEK 45 - 48)"
            Case 12
                tmp = "DEC (WEEK 49 - 52)"
            Case 13
                tmp = "ALL (WEEK 01 - 52)"
        End Select

        ' Return Value
        GetMonthName = tmp

        ' Destroy Variables
        tmp = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Function

    Function GetTrafficName(ByVal traffic As Integer) As String
        Dim tmp As String = ""

        Select Case traffic
            Case 0
                tmp = "ALL"
            Case 1
                tmp = "NON-USA"
            Case 2
                tmp = "USA"
            Case 3
                tmp = "EUR"
            Case 4
                tmp = "CHN"
            Case 5
                tmp = "SEA"
            Case 6
                tmp = "NEA"
            Case 7
                tmp = "AUS"
            Case 8
                tmp = "CAN"
            Case 9
                tmp = "SAM"
            Case 10
                tmp = "ISC"
            Case 11
                tmp = "AFR"
        End Select

        ' Return Value
        GetTrafficName = tmp

        ' Destroy Variables
        tmp = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Function GetBranchName(ByVal BrhCd As Integer) As String

        Dim tmp As String = ""

        Select Case BrhCd
            Case 9
                tmp = "SHA"
            Case 32
                tmp = "SHA-NON-USA"
            Case 34
                tmp = "SHA-AIR"
            Case 38
                tmp = "SHA-IMPORT"
            Case 59
                tmp = "NINGBO"
            Case 61
                tmp = "NINGBO-NON-USA"
            Case 62
                tmp = "NINGBO-IMPORT"
            Case 63
                tmp = "NINGBO-AIR"
        End Select

        GetBranchName = tmp

        tmp = Nothing

        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Function

    Function GetSubBranchName(ByVal subBrh As Integer, ByVal GenID As Integer) As String

        Dim tmp As String = ""

        Select Case GenID
            Case 1, 2 ' HONG KONG / SHENZHEN
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 1
                        tmp = "HKG-EXPORT"
                    Case 4
                        tmp = "SHENZHEN"
                    Case 5
                        tmp = "HKG-IMPORT"
                    Case 12
                        tmp = "HKG-OFF"
                    Case 26
                        tmp = "GUANGZHOU"
                    Case 40
                        tmp = "ZHONGSHAN"
                    Case 46
                        tmp = "ZHONGSHAN-OFF"
                End Select

            Case 4 ' SHANGHAI
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 9
                        tmp = "SHANGHAI"
                    Case 16
                        tmp = "NINGBO"
                    Case 18
                        tmp = "LIANYUNGANG"
                    Case 19
                        tmp = "ZHANGJIAGANG"
                    Case 20
                        tmp = "NANJING"
                    Case 21
                        tmp = "NANTONG"
                    Case 22
                        tmp = "JIUJIANG"
                    Case 32
                        tmp = "SHA-NON-USA"
                    Case 33
                        tmp = "SHA-OFF"
                    Case 38
                        tmp = "SHA-IMPORT"
                    Case 46
                        tmp = "NINGBO-OFF"
                    Case 48
                        tmp = "NGB-IMPORT"
                End Select

            Case 7 ' MALAYSIA
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 28
                        tmp = "PENANG-USA"
                    Case 30
                        tmp = "MALAYSIA-IMP"
                    Case 31
                        tmp = "MALAYSIA-OFF"
                    Case 36
                        tmp = "SUBANG"
                    Case 37
                        tmp = "PENANG-NONUSA"
                    Case 49
                        tmp = "SABAH"
                    Case 50
                        tmp = "SABAH-OFF"
                End Select

            Case 8 ' XIAMEN
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 35
                        tmp = "XIAMEN"
                    Case 47
                        tmp = "XIAMEN-OFF"
                End Select

            Case 9 ' DALIAN
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 17
                        tmp = "DALIAN"
                    Case 47
                        tmp = "DALIAN-NON-USA"
                    Case 49
                        tmp = "DALIAN-OFF"
                    Case 50
                        tmp = "DALIAN-IMPORT"
                End Select

            Case 10 ' QINGDAO
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 39
                        tmp = "QINGDAO"
                    Case 44
                        tmp = "QINGDAO-OFF"
                End Select

            Case 13 ' FUZHOU
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 43
                        tmp = "FUZHOU"
                    Case 48
                        tmp = "FUZHOU-OFF"
                End Select

            Case 14 ' TIANJIN
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 44
                        tmp = "TIANJIN"
                    Case 49
                        tmp = "TIANJIN-OFF"
                End Select

            Case 15 ' JAKARTA
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 49
                        tmp = "JAKARTA"
                    Case 50
                        tmp = "JAKARTA-OFF"
                End Select

            Case 16 ' VIETNAM
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 45
                        tmp = "VIETNAM"
                    Case 47
                        tmp = "VIETNAM-OFF"
                End Select

            Case 17 ' VIETNAM
                Select Case subBrh
                    Case 0
                        tmp = "ALL"
                    Case 59
                        tmp = "NINGBO"
                    Case 61
                        tmp = "NGB-NON-USA"
                    Case 60
                        tmp = "NINGBO-OFF"
                    Case 62
                        tmp = "NGB-IMPORT"
                    Case 63
                        tmp = "NINGBO-AIR"
                    Case 64
                        tmp = "CMT-NINGBO"
                    Case 65
                        tmp = "LOCAL-NGB"

                End Select

        End Select

        ' Return Value
        GetSubBranchName = tmp

        ' Destroy Variables
        tmp = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Function GetPOType(ByVal POType As Integer) As String

        Dim tmp As String = ""

        Select Case POType
            Case 0
                tmp = "ALL"
            Case 51
                tmp = "11-D Report"
            Case 52
                tmp = "PO Tracking"
            Case 87
                tmp = "Others"
        End Select

        ' Return Value
        GetPOType = tmp

        ' Destroy Variables
        tmp = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Function

    Function setQuote(ByVal inVal As String) As String

        setQuote = Replace(inVal, "'", "''")

    End Function

    Function DigitToMonth(ByVal InMonth As Integer) As String

        DigitToMonth = "JAN"

        Select Case InMonth
            Case 1
                DigitToMonth = "JAN"
            Case 2
                DigitToMonth = "FEB"
            Case 3
                DigitToMonth = "MAR"
            Case 4
                DigitToMonth = "APR"
            Case 5
                DigitToMonth = "MAY"
            Case 6
                DigitToMonth = "JUN"
            Case 7
                DigitToMonth = "JUL"
            Case 8
                DigitToMonth = "AUG"
            Case 9
                DigitToMonth = "SEP"
            Case 10
                DigitToMonth = "OCT"
            Case 11
                DigitToMonth = "NOV"
            Case 12
                DigitToMonth = "DEC"
        End Select

    End Function

    Public Sub UpdateRptType()

        Dim RptType As String = ""

        ' Update Selected Report Types on Main Screen
        If My.Settings.IsExcel = True Then
            If RptType = "" Then
                RptType &= "Excel"
            Else
                RptType &= ", Excel"
            End If
        End If

        If My.Settings.IsPDF = True Then
            If RptType = "" Then
                RptType &= "PDF"
            Else
                RptType &= ", PDF"
            End If
        End If

        If My.Settings.IsTxt = True Then
            If RptType = "" Then
                RptType &= "Text"
            Else
                RptType &= ", Text"
            End If
        End If

        If My.Settings.IsZip = True Then
            If RptType = "" Then
                RptType &= "Zip"
            Else
                RptType &= ", Zip"
            End If
        End If

        frmMain.lblRptType.Text = "Selected Report Type(s): " & RptType

        ' Destroy Variables
        RptType = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

End Class
