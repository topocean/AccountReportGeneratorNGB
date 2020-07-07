Imports System.IO

Public Class RptInttra

    Public Function RptInttra(ByVal uid As String, ByVal ds As DataSet) As String
        Dim common As New common
        Dim filename As String = ""
        Dim tmpPath = "C:\TxtFiles"
        Dim i, j As Integer
        Dim txtStreamWriter As StreamWriter
        Dim txtFile As String = ""
        Dim exportFile As String = ""
        Dim valueStr As String()

        Try
            ' Check Existance of Temp. Directory
            If Not My.Computer.FileSystem.DirectoryExists(tmpPath) Then
                My.Computer.FileSystem.CreateDirectory(tmpPath)
            End If

            If ds.Tables(1).Rows.Count > 0 Then
                filename = ds.Tables(0).Rows(0).Item("RptFile").ToString
                txtFile = tmpPath & "\" & filename & "txt"

                txtStreamWriter = New StreamWriter(txtFile, True)

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    With ds.Tables(1).Rows(i)
                        valueStr = Split(Replace(.Item("Content"), Chr(10), ""), Chr(13))

                        For j = LBound(valueStr) To UBound(valueStr)
                            If Trim(valueStr(j)) <> "" Then
                                txtStreamWriter.WriteLine(Trim(Replace(valueStr(j), Chr(13), "")))
                            End If
                        Next
                    End With
                Next

                txtStreamWriter.Flush()
                txtStreamWriter.Close()

                ' Export File
                filename &= ".txt"
                exportFile = My.Settings.ExportPath & filename

                ' Delete file if already existed
                If My.Computer.FileSystem.FileExists(exportFile) Then
                    My.Computer.FileSystem.DeleteFile(exportFile)
                End If

                ' Move File
                My.Computer.FileSystem.MoveFile(txtFile, exportFile)
            End If
        Catch ex As Exception
            filename = "Error," & ex.Message
        End Try

        ' Destroy Variables
        i = Nothing
        j = Nothing
        txtStreamWriter = Nothing
        tmpPath = Nothing
        common = Nothing
        exportFile = Nothing
        txtFile = Nothing
        valueStr = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

        RptInttra = filename
    End Function
End Class
