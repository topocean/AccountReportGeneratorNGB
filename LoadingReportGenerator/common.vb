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

    Function GetSubBranchName(ByVal subBrh As Integer) As String
        Dim tmp As String = ""

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
            Case Else
                tmp = "N/A"
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
End Class
