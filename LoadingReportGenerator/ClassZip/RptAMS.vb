Imports System.Xml
Imports System.IO
Imports CGZipLibrary

Public Class RptAMS

    Public Function RptAMS(ByVal uid As String, ByVal ds As DataSet) As String
        Dim common As New common
        Dim filename As String = ""
        Dim xWriter As XmlTextWriter
        Dim tmpPath = "C:\AMS\" & uid
        Dim i, j As Integer
        Dim txtStreamWriter As StreamWriter
        Dim xmlFile As String = ""
        Dim txtFile As String = ""
        Dim zipFile As String = tmpPath & "\" & uid & ".zip"
        Dim exportFile As String = ""
        Dim ctnrRows As DataRow()
        Dim objZip As New CGZipFiles

        Try
            ' Check Existance of Temp. Directory
            If Not My.Computer.FileSystem.DirectoryExists(tmpPath) Then
                My.Computer.FileSystem.CreateDirectory(tmpPath)
            End If

            ' --------------------------------------------------------
            ' Export Shipment Warning Messages
            ' --------------------------------------------------------

            If ds.Tables(0).Rows.Count > 0 Then
                txtFile = tmpPath & "\" & "ErrorList.txt"
                txtStreamWriter = New StreamWriter(txtfile, True, System.Text.Encoding.ASCII)

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    With ds.Tables(0).Rows(i)
                        txtStreamWriter.WriteLine(.Item("ErrorTxt").ToString)
                    End With
                Next

                txtStreamWriter.Close()
                txtStreamWriter.Dispose()
            End If

            ' --------------------------------------------------------
            ' End of Export Shipment Warning Messages
            ' --------------------------------------------------------


            ' --------------------------------------------------------
            ' Export XML EDI Document
            ' --------------------------------------------------------

            If ds.Tables(1).Rows.Count > 0 Then
                For i = 0 To ds.Tables(1).Rows.Count - 1
                    With ds.Tables(1).Rows(i)
                        xmlFile = tmpPath & "\" & .Item("BkhBLNo").ToString & ".xml"
                        xWriter = New XmlTextWriter(xmlFile, System.Text.Encoding.UTF8)

                        ' XML Formatting
                        xWriter.WriteStartDocument()
                        xWriter.Formatting = Formatting.Indented
                        xWriter.Indentation = 4

                        ' --------------------------------------------------------

                        ' Root Element (Manifest)
                        xWriter.WriteStartElement("Manifest")

                        ' BillOfLading Segment
                        xWriter.WriteStartElement("BillOfLading")

                        xWriter.WriteElementString("SiteId", .Item("SiteID").ToString)
                        xWriter.WriteElementString("VesselName", .Item("VslName").ToString)
                        xWriter.WriteElementString("VesselFlag", .Item("VslFlag").ToString)
                        xWriter.WriteElementString("VoyageNumber", .Item("VslVoy").ToString)
                        xWriter.WriteElementString("ETD", .Item("BkhETD").ToString)
                        xWriter.WriteElementString("ETA", .Item("BkhETA").ToString)
                        xWriter.WriteElementString("HouseBillNumber", .Item("BkhBLNo").ToString)
                        xWriter.WriteElementString("MasterBillNumber", .Item("BkhMBLNo").ToString)
                        xWriter.WriteElementString("PortOfLoad", .Item("BkhLoad").ToString)
                        xWriter.WriteElementString("PortOfDischarge", .Item("BkhDisc").ToString)
                        xWriter.WriteElementString("LastForeignPort", .Item("VslLastPort").ToString)
                        xWriter.WriteElementString("TotalPieces", .Item("TotalPKG").ToString)
                        xWriter.WriteElementString("UnitOfMeasure", .Item("UnitOfMeasure").ToString)
                        xWriter.WriteElementString("TotalKilos", .Item("TotalWGT").ToString)
                        xWriter.WriteElementString("BillOfLadingType", .Item("BLType").ToString)
                        xWriter.WriteElementString("SCAC_Carrier", .Item("ClientSCAC").ToString)
                        xWriter.WriteElementString("SCAC_Secondary", "")
                        xWriter.WriteElementString("AmendmentFlag", .Item("AmdFlag").ToString)
                        xWriter.WriteElementString("TotalCBM", .Item("TotalCBM").ToString)
                        xWriter.WriteElementString("PlaceOfReceipt", .Item("BkhReceipt").ToString)
                        xWriter.WriteElementString("PlaceOfDelivery", .Item("BkhDest").ToString)
                        xWriter.WriteElementString("SenderUniqueReference", .Item("BkhBLNo").ToString)

                        ' --------------------------------------------------------

                        ' Shipper Segment inside BillOfLading
                        xWriter.WriteStartElement("ShipperPartyInfo")

                        xWriter.WriteElementString("Name", .Item("ShpName").ToString)
                        xWriter.WriteElementString("Address1", .Item("ShpAddr1").ToString)
                        xWriter.WriteElementString("Address2", .Item("ShpAddr2").ToString)
                        xWriter.WriteElementString("Address3", .Item("ShpAddr3").ToString)
                        xWriter.WriteElementString("CityName", .Item("ShpCity").ToString)
                        xWriter.WriteElementString("CountryCode", .Item("ShpCouCd").ToString)
                        xWriter.WriteElementString("StateOrProvinceCode", .Item("ShpStateCd").ToString)
                        xWriter.WriteElementString("PostCode", .Item("ShpPostCd").ToString)

                        xWriter.WriteEndElement()

                        ' --------------------------------------------------------

                        ' Consignee Segment inside BillOfLading
                        xWriter.WriteStartElement("ConsigneePartyInfo")

                        xWriter.WriteElementString("Name", .Item("ConName").ToString)
                        xWriter.WriteElementString("Address1", .Item("ConAddr1").ToString)
                        xWriter.WriteElementString("Address2", .Item("ConAddr2").ToString)
                        xWriter.WriteElementString("Address3", .Item("ConAddr3").ToString)
                        xWriter.WriteElementString("CityName", .Item("ConCity").ToString)
                        xWriter.WriteElementString("CountryCode", .Item("ConCouCd").ToString)
                        xWriter.WriteElementString("StateOrProvinceCode", .Item("ConStateCd").ToString)
                        xWriter.WriteElementString("PostCode", .Item("ConPostCd").ToString)

                        xWriter.WriteEndElement()

                        ' --------------------------------------------------------

                        ' Notify Segment inside BillOfLading
                        If common.NullVal(.Item("NotCd").ToString, 0) <> 0 Then
                            xWriter.WriteStartElement("NotifyPartyInfo")

                            xWriter.WriteElementString("Name", .Item("NotName").ToString)
                            xWriter.WriteElementString("Address1", .Item("NotAddr1").ToString)
                            xWriter.WriteElementString("Address2", .Item("NotAddr2").ToString)
                            xWriter.WriteElementString("Address3", .Item("NotAddr3").ToString)
                            xWriter.WriteElementString("CityName", .Item("NotCity").ToString)
                            xWriter.WriteElementString("CountryCode", .Item("NotCouCd").ToString)
                            xWriter.WriteElementString("StateOrProvinceCode", .Item("NotStateCd").ToString)
                            xWriter.WriteElementString("PostCode", .Item("NotPostCd").ToString)

                            xWriter.WriteEndElement()
                        End If

                        ' --------------------------------------------------------

                        ' Container Segment(s) of BillOfLading
                        ctnrRows = ds.Tables(2).Select("BkhRefId = " & common.NullVal(.Item("BkhRefId").ToString, 0))

                        For j = LBound(ctnrRows) To UBound(ctnrRows)
                            xWriter.WriteStartElement("Container")

                            With ctnrRows(j)
                                xWriter.WriteElementString("EquimentInitial", .Item("CtnrInit").ToString)
                                xWriter.WriteElementString("EquimentNum", .Item("CtnrNo").ToString)
                                xWriter.WriteElementString("EquimentTypeCode", .Item("CtnrSizCd").ToString)
                                xWriter.WriteElementString("TypeOfServiceCode", .Item("TypeOfSvc").ToString)
                                xWriter.WriteElementString("SealNum1", .Item("CtnrSeal").ToString)

                                ' Container Content
                                xWriter.WriteStartElement("ContainerContent")

                                xWriter.WriteElementString("Quantity", .Item("CtnrPKG").ToString)
                                xWriter.WriteElementString("UnitOfMeasure", .Item("UntCd").ToString)
                                xWriter.WriteElementString("FreeFormDescription", .Item("Description").ToString)
                                xWriter.WriteElementString("MarksAndNumbers", .Item("Marks").ToString)
                                xWriter.WriteElementString("Kilos", .Item("CtnrWGT").ToString)

                                xWriter.WriteEndElement()

                                ' -----------------------------------------------

                            End With

                            xWriter.WriteEndElement()
                        Next

                        ' --------------------------------------------------------

                        ' End of BillOfLading
                        xWriter.WriteEndElement()

                        ' End of Manifest
                        xWriter.WriteEndDocument()
                        xWriter.Flush()
                        xWriter.Close()
                    End With
                Next

                ' Add the EDI files into one Zip archivement
                objZip.ZipFileName = zipFile
                objZip.RootDirectory = tmpPath
                objZip.AddFile("*.*")

                If objZip.MakeZipFile <> 0 Then
                    filename = "Error," & objZip.GetLastMessage

                    ' Release Memory
                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                    Exit Try
                End If

                filename = uid & ".zip"

                ' Put the zip file onto Web Directory
                exportFile = My.Settings.ExportPath & filename

                ' Delete file if already existed
                If My.Computer.FileSystem.FileExists(exportFile) Then
                    My.Computer.FileSystem.DeleteFile(exportFile)
                End If

                ' Move File
                My.Computer.FileSystem.MoveFile(zipFile, exportFile)
            End If

            ' --------------------------------------------------------
            ' End of Export EDI Document
            ' --------------------------------------------------------

        Catch ex As Exception
            filename = "Error," & ex.Message
        End Try

        ' Destroy Variables
        i = Nothing
        j = Nothing
        objZip = Nothing
        txtFile = Nothing
        xmlFile = Nothing
        zipFile = Nothing
        exportFile = Nothing
        tmpPath = Nothing
        common = Nothing
        xWriter = Nothing
        txtStreamWriter = Nothing
        ctnrRows = Nothing

        ' Release Memory
        GC.Collect()
        GC.WaitForPendingFinalizers()

        ' Return & Exit Function
        RptAMS = filename
    End Function
End Class

