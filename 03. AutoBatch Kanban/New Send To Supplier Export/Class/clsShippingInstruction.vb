Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class clsShippingInstruction
    Shared Sub up_SendShippingInstruction(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,                              
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim ls_sql As String = ""
        Dim pAffiliate As String = ""
        Dim pForwarder As String = ""
        Dim pShippingInstruction As String = ""

        Dim pFileName1 As String = ""
        Dim pFileName2 As String = ""

        Dim ds As New DataSet

        Try
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data SI")

            '---------------------------------------excel Kanban ---------------------------------------'
            ls_sql = " Select distinct ShippingInstructionNo, AffiliateID, ForwarderID From ShippingInstruction_Master where isnull(TallyCls,'') = '1' "

            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pShippingInstruction = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))
                    pForwarder = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))

                    pFileName1 = ""
                    pFileName2 = ""

                    '88. Create PDF SI
                    log.WriteToProcessLog(Date.Now, pScreenName, "Start Create SI File, ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "]")
                    If CreateSIToPDF(GB, cfg, pAffiliate, pForwarder, pShippingInstruction, pFileName1, pResult, errMsg) = False Then
                        log.WriteToProcessLog(Date.Now, pScreenName, "End Create SI File, ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "] " & errMsg)
                        If errMsg = "Microsoft.VisualBasic.ErrObject" Then
                            End
                        End If                        
                        GoTo keluar
                    End If
                    log.WriteToProcessLog(Date.Now, pScreenName, "End Create SI File, ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "]")

                    '89. Create CSV
                    If CreateCSV(GB, log, pAffiliate, pForwarder, pShippingInstruction, pAtttacment, pResult, pScreenName, pFileName2, LogName, errMsg, ErrSummary) = False Then
                        GoTo keluar
                    End If

                    If pFileName1 <> "" And pFileName2 <> "" Then
                        If sendEmailtoForwarder(GB, pResult, pShippingInstruction, pAffiliate, pForwarder, errMsg, pFileName2, pFileName1) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Forwarder. ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "] ok.")
                        End If
                    End If

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send SI [" & pShippingInstruction & "], Forwarder [" & pForwarder & "], Affiliate [" & pAffiliate & "] ok", LogName)
                    LogName.Refresh()

                    Call UpdateTallyCls(pShippingInstruction, pAffiliate, pForwarder, errMsg)
keluar:
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "SI [" & pAffiliate & "-" & pForwarder & "-" & pShippingInstruction & "] " & ex.Message
            ErrSummary = "SI [" & pAffiliate & "-" & pForwarder & "-" & pShippingInstruction & "] " & ex.Message
        Finally
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    Shared Function CreateSIToPDF(ByVal GB As GlobalSetting.clsGlobal, ByVal cfg As GlobalSetting.clsConfig, ByVal pAffiliate As String, ByVal pForwarderID As String, ByVal pShippingNo As String, ByRef pFileName As String, ByVal pPathFile As String, ByRef errMsg As String) As Boolean
        Dim CrReport As New rptShippingInstruction()
        Dim CrExportOptions As ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

        Try
            Cursor.Current = Cursors.WaitCursor
            Dim dsPrint As New DataSet
            dsPrint = PrintSI(GB, pShippingNo, pAffiliate, pForwarderID)

            If IsNothing(dsPrint) Then
                Exit Try
            End If

            If dsPrint.Tables(0).Rows.Count = 0 Then Exit Try

            CrReport.SetDatabaseLogon(cfg.User, cfg.Password, cfg.Server, cfg.Database)
            CrReport.SetDataSource(dsPrint.Tables(0))

            Dim ls_FileName = pPathFile & "\ShippingInstruction-" & pShippingNo.Trim & "-" & pAffiliate.Trim & "-" & pForwarderID.Trim & ".pdf"
            pFileName = ls_FileName

            CrDiskFileDestinationOptions.DiskFileName = ls_FileName

            CrExportOptions = CrReport.ExportOptions

            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With

            Try
                CrReport.Export()
            Catch err As Exception
                errMsg = err.ToString()
                CreateSIToPDF = False
            End Try
            'PDF
            CreateSIToPDF = True
        Catch ex As Exception
            errMsg = Err.ToString()
            CreateSIToPDF = False
        Finally
            If Not CrReport Is Nothing Then
                clsGeneral.NAR(CrReport)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If
        End Try
    End Function

    Shared Function CreateCSV(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, _
                                   ByVal pAffiliate As String, ByVal pForwarderID As String, ByVal pShippingNo As String, _
                                   ByVal pAtttacment As String, ByVal pResult As String, _
                                   ByRef pScreenName As String, ByRef pFileName1 As String, _
                                   ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        CreateCSV = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""

        Dim dsDetail As New DataSet

        Dim i As Integer

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template CSV.xlsx") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send CSV to Forwarder STOPPED, File Excel isn't Found"
                errSummary = "Process Send CSV to Forwarder STOPPED, File Excel isn't Found"
                CreateCSV = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create CSV [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "]")

            NewFileCopy = pAtttacment & "\Template CSV.xlsx"
            NewFileCopyTO = pResult & "\Template CSV " & Format(Now, "HHmmss") & ".xlsx"

            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
            Else
                System.IO.File.Copy(NewFileCopy, pResult & "\Delivery.xlsx")
            End If

            Dim ls_file As String = NewFileCopyTO

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetail = PrintCSV(GB, pShippingNo, pAffiliate, pForwarderID)

            If dsDetail.Tables(0).Rows.Count > 0 Then
                log.WriteToProcessLog(Date.Now, pScreenName, "Input Header CSV [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "]")

                i = 0
                ExcelSheet.Range("A" & i + 1).Value = "Invoice No."
                ExcelSheet.Range("B" & i + 1).Value = "Consignee Code"
                ExcelSheet.Range("C" & i + 1).Value = "Buyer Code"
                ExcelSheet.Range("D" & i + 1).Value = "Shipment"
                ExcelSheet.Range("E" & i + 1).Value = "Shipping Line"
                ExcelSheet.Range("F" & i + 1).Value = "Vessel"
                ExcelSheet.Range("G" & i + 1).Value = "Voyage"
                ExcelSheet.Range("H" & i + 1).Value = "From Port"
                ExcelSheet.Range("I" & i + 1).Value = "VIA"
                ExcelSheet.Range("J" & i + 1).Value = "To Port"
                ExcelSheet.Range("K" & i + 1).Value = "ETD"
                ExcelSheet.Range("L" & i + 1).Value = "ETA"
                ExcelSheet.Range("M" & i + 1).Value = "Order No"
                ExcelSheet.Range("N" & i + 1).Value = "Original O/No"
                ExcelSheet.Range("O" & i + 1).Value = "Part No"
                ExcelSheet.Range("P" & i + 1).Value = "Part Group Name"
                ExcelSheet.Range("Q" & i + 1).Value = "Box No. From"
                ExcelSheet.Range("R" & i + 1).Value = "Box No. To"
                ExcelSheet.Range("S" & i + 1).Value = "Carton Count"
                ExcelSheet.Range("T" & i + 1).Value = "Quantity"
                ExcelSheet.Range("U" & i + 1).Value = "Total Quantity"
                ExcelSheet.Range("V" & i + 1).Value = "Net(Weight(KGM))"

                For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                    ExcelSheet.Range("A" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("InvoiceNo"))
                    ExcelSheet.Range("B" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Consignee"))
                    ExcelSheet.Range("C" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Buyer"))
                    ExcelSheet.Range("D" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Shipment"))
                    ExcelSheet.Range("E" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("ShippingLine"))
                    ExcelSheet.Range("F" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Vessel"))
                    ExcelSheet.Range("G" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Voyage"))
                    ExcelSheet.Range("H" & i + 2).Value = "JAKARTA" 'Trim(dsDetail.Tables(0).Rows(i)("FromPort"))
                    ExcelSheet.Range("I" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("VIA"))
                    ExcelSheet.Range("J" & i + 2).Value = Trim(Replace(dsDetail.Tables(0).Rows(i)("ToPort"), ",", ""))
                    ExcelSheet.Range("K" & i + 2).NumberFormat = "@"
                    ExcelSheet.Range("K" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("ETD"))
                    ExcelSheet.Range("L" & i + 2).NumberFormat = "@"
                    ExcelSheet.Range("L" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("ETA"))
                    ExcelSheet.Range("M" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("OrderNo"))
                    ExcelSheet.Range("N" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("OriginalNo"))
                    ExcelSheet.Range("O" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                    ExcelSheet.Range("P" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("PartGroupName"))
                    ExcelSheet.Range("Q" & i + 2).Value = Trim(Split(dsDetail.Tables(0).Rows(i)("BoxNo"), "-")(0))
                    ExcelSheet.Range("R" & i + 2).Value = Trim(Split(dsDetail.Tables(0).Rows(i)("BoxNo"), "-")(1))
                    ExcelSheet.Range("S" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("CartonCount"))
                    ExcelSheet.Range("T" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("QtyBox"))
                    ExcelSheet.Range("U" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Quantity"))
                    ExcelSheet.Range("V" & i + 2).Value = Trim(dsDetail.Tables(0).Rows(i)("Net"))
                Next

                'Save ke Local
                xlApp.DisplayAlerts = False
                pFileName1 = pResult & "\CSV " & Trim(pAffiliate) & "-" & Trim(pForwarderID) & "-" & Trim(pShippingNo) & ".csv"

                ExcelBook.SaveAs(pFileName1, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, CreateBackup:=False)

                xlApp.Workbooks.Close()
                xlApp.Quit()
            End If
        Catch ex As Exception
            CreateCSV = False

            log.WriteToErrorLog(pScreenName, "Process Create CSV [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Create CSV [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "] STOPPED, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Create Kanban [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "] STOPPED, because " & ex.Message, LogName)
            LogName.Refresh()
        Finally
            If Not xlApp Is Nothing Then
                xlApp.DisplayAlerts = False
                clsGeneral.NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                clsGeneral.NAR(ExcelBook)
                xlApp.Quit()
                clsGeneral.NAR(xlApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If

            If Not dsDetail Is Nothing Then
                dsDetail.Dispose()
            End If
        End Try
    End Function

    Shared Function PrintSI(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT DISTINCT " & vbCrLf & _
                  "   ShippingInstructionNo = SIM.ShippingInstructionNo,            " & vbCrLf & _
                  "   FWD = Rtrim(MF.ForwarderName) + ' ' + Rtrim(MF.Address) + ' ' + Rtrim(MF.City) + ' ' + Rtrim(MF.PostalCode),           " & vbCrLf & _
                  "   ATT = isnull(Rtrim(MF.Attn),''),            " & vbCrLf & _
                  "   FAx = isnull(Rtrim(MF.Fax),''),            " & vbCrLf & _
                  "   Tujuan = isnull(MA.DestinationPort,''), " & vbCrLf & _
                  "   Shipment = Case when TypeOfService = 'FCL' then 'SEA FREIGHT' WHEN TypeOfService = 'LCL' then 'SEA FREIGHT' ELSE 'AIR FREIGHT' END, " & vbCrLf & _
                  "   Vessel = ISNULL(SIM.Vessels,''), " & vbCrLf & _
                  "   ETD = Convert(Char(12), convert(Datetime, SIM.ETDPort),106),            " & vbCrLf & _
                  "   ETA = Convert(Char(12), convert(Datetime, SIM.ETAPort),106),            " & vbCrLf & _
                  "   tgltiba = Convert(Char(12), convert(Datetime, SIM.ETAPort),106),             "

        ls_SQL = ls_SQL + "   part = 'Automotive Component',            " & vbCrLf & _
                          "   jumlah = 0, " & vbCrLf & _
                          "   pallet = isnull(SIM.TotalPallet,0), " & vbCrLf & _
                          "   Box = isnull(SUM(SD.ShippingQty/MPM.QtyBox),0), " & vbCrLf & _
                          "   Qty = SUM(isnull(SD.ShippingQty,0)), " & vbCrLf & _
                          "   beratBersih = SUM(((netweight/MPM.QtyBox)* SD.ShippingQty)/1000),            " & vbCrLf & _
                          "   beratKotor = ISNULl(SIM.GrossWeight,0), " & vbCrLf & _
                          "   Buyer = Rtrim(BuyerName),            " & vbCrLf & _
                          "   BuyerAddress = Rtrim(BuyerAddress),  " & vbCrLf & _
                          "   Consignee = Rtrim(MA.ConsigneeName), ConsigneeAddress = Rtrim(MA.ConsigneeAddress), Attn =isnull(MA.Att,''),  " & vbCrLf & _
                          "   Freight = isnull(Freight,'') "

        ls_SQL = ls_SQL + " From ShippingInstruction_master SIM            " & vbCrLf & _
                          " LEFT JOIN ShippingInstruction_Detail SD            " & vbCrLf & _
                          " 	ON SIM.ShippingInstructionNo = SD.ShippingInstructionNo            " & vbCrLf & _
                          " 	AND SIM.AffiliateID = SD.AffiliateID          " & vbCrLf & _
                          " 	AND SIM.ForwarderID = SD.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SIM.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SIM.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SD.PartNo  "

        ls_SQL = ls_SQL + " 	and MPM.SupplierID = SD.SupplierID  " & vbCrLf & _
                          " 	and MPM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                          " WHERE SIM.ShippingInstructionNo = '" & Trim(ls_value1) & "' " & vbCrLf & _
                          " AND SIM.AffiliateID = '" & Trim(ls_value2) & "'" & vbCrLf & _
                          " AND SIM.ForwarderID = '" & Trim(ls_value3) & "' " & vbCrLf & _
                          " GROUP BY SIM.ShippingInstructionNo, " & vbCrLf & _
                          " 	Rtrim(MF.ForwarderName) ,Rtrim(MF.Address) , Rtrim(MF.City) , Rtrim(MF.PostalCode),  " & vbCrLf & _
                          " 	isnull(Rtrim(MF.Attn),''),   " & vbCrLf & _
                          " 	isnull(Rtrim(MF.Fax),''),           " & vbCrLf & _
                          " 	isnull(MA.DestinationPort,''), " & vbCrLf & _
                          " 	TypeOfService, SIM.VesselS, "

        ls_SQL = ls_SQL + " 	SIM.ETDPort, SIM.ETAPort, SIM.TotalPallet, SIM.GrossWeight, " & vbCrLf & _
                          " 	Rtrim(BuyerName),Rtrim(BuyerAddress),Rtrim(MA.ConsigneeName), Rtrim(MA.ConsigneeAddress), MA.Att, SIM.Freight " & vbCrLf & _
                          "  "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function PrintCSV(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT  DISTINCT " & vbCrLf & _
                  " 	InvoiceNo = SHM.ShippingInstructionNo,  " & vbCrLf & _
                  " 	Consignee = MA.ConsigneeCode, " & vbCrLf & _
                  " 	Buyer = MA.BuyerCode, " & vbCrLf & _
                  " 	Shipment = case when POM.ShipCls='A' then 'Air Freight' else 'Sea Freight' End, " & vbCrLf & _
                  " 	ShippingLine = ISNULL(SHM.ShippingLineS,''),   " & vbCrLf & _
                  " 	Vessel = isnull(SHM.NamaKapalS,''),   " & vbCrLf & _
                  " 	Voyage = isnull(SHM.VesselS,''), " & vbCrLf & _
                  " 	FromPort = MF.Port, " & vbCrLf & _
                  " 	VIA = ISNULL(SHM.Via,''), " & vbCrLf & _
                  " 	ToPort = MA.DestinationPort, "

        ls_SQL = ls_SQL + " 	ETD = Replace(CONVERT(CHAR(11), CONVERT(DATETIME, SHM.ETDPort), 104),'.',''),     " & vbCrLf & _
                          " 	ETA = Replace(CONVERT(CHAR(11), CONVERT(DATETIME, SHM.ETAPort), 104),'.',''), " & vbCrLf & _
                          " 	OrderNo = RD.OrderNo,   " & vbCrLf & _
                          " 	OriginalNo = RD.PONo, " & vbCrLf & _
                          " 	PartNo = SDM.PartNo,   " & vbCrLf & _
                          " 	PartGroupName = isnull(PartGroupName,''), " & vbCrLf & _
                          " 	SDM.BoxNo, " & vbCrLf & _
                          " 	CartonCount = SDM.BoxQty, " & vbCrLf & _
                          " 	QtyBox = MPM.QtyBox,   " & vbCrLf & _
                          " 	Quantity = MPM.QtyBox * SDM.BoxQty,   " & vbCrLf & _
                          " 	Net = MPM.NetWeight /1000 "

        ls_SQL = ls_SQL + " FROM ShippingInstruction_Master SHM  " & vbCrLf & _
                          " INNER JOIN ShippingInstruction_Detail SDM  " & vbCrLf & _
                          " 	ON ltrim(SDM.ShippingInstructionNo) = ltrim(SHM.ShippingInstructionNo)    " & vbCrLf & _
                          "   	AND ltrim(SDM.ForwarderID) = rtrim(SHM.ForwarderID)    " & vbCrLf & _
                          "   	AND rtrim(SDM.AffiliateID) = rtrim(SHM.AffiliateID) " & vbCrLf & _
                          " LEFT JOIN PO_Master_Export POM ON POM.PONo = SDM.OrderNo    " & vbCrLf & _
                          "      AND POM.AffiliateID = SDM.AffiliateID   " & vbCrLf & _
                          "      AND POM.SupplierID = SDM.SupplierID " & vbCrLf & _
                          " LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNO = SDM.SuratJalanno    " & vbCrLf & _
                          "    	AND RD.AffiliateID = SDM.AffiliateID     	 " & vbCrLf & _
                          " 	AND RD.SupplierID = SDM.SupplierID     	 "

        ls_SQL = ls_SQL + "    	AND RD.OrderNO = SDM.OrderNo    " & vbCrLf & _
                          "   	AND RD.PartNo = SDM.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SDM.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SHM.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Parts MP ON MP.PartNo = SDM.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SDM.PartNo  " & vbCrLf & _
                          " 	AND MPM.AffiliateID = SDM.AffiliateID AND MPM.SupplierID = SDM.SupplierID " & vbCrLf & _
                          " Where SHM.ShippingInstructionNo = '" & Trim(ls_value1) & "' " & vbCrLf & _
                          " AND SHM.AffiliateID = '" & Trim(ls_value2) & "' " & vbCrLf & _
                          " AND SHM.ForwarderID = '" & Trim(ls_value3) & "' "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Sub UpdateTallyCls(ByVal pShippingNo As String, ByVal pAffiliate As String, ByVal pForwarder As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & vbCrLf & _
                      " SET TallyCls='2'" & vbCrLf & _
                      " WHERE ShippingInstructionNo = '" & pShippingNo & "' " & vbCrLf & _
                      " AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                      " AND ForwarderID = '" & pForwarder & "'"

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send Shipping Instruction: ShippingInstructionNo [" & pShippingNo.Trim & "], AffiliateID [" & pAffiliate.Trim & "], ForwarderID [" & pForwarder.Trim & "] to Forwarder STOPPED, because " & ex.Message
        End Try
    End Sub

    Shared Function sendEmailtoForwarder(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                        ByVal pShippingNo As String, ByVal pAffiliate As String, ByVal pForwarder As String, ByRef errMsg As String, _
                                        Optional ByVal pDN1 As String = "", Optional ByVal pBarcodeFile As String = "") As Boolean
        Dim TempFilePath As String = Trim(pPathFile)

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoForwarder = True

            Dim dsEmail As New DataSet
            dsEmail = clsGeneral.getEmailAddressExport(GB, "", "PASI", "", pForwarder, "SupplierDeliveryCC", "SupplierDeliveryTo", "SupplierDeliveryTo", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "FWD" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "FWD" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If fromEmail = "" Then
                sendEmailtoForwarder = False
                errMsg = "Process Send SI [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoForwarder = False
                errMsg = "Process Send SI [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            'ls_Subject = "Send To Supplier PO Non Kanban : " & pAffiliate.Trim & "-" & pPONo & "-" & pKanbanNo & "-" & pSupplier.Trim
            ls_Subject = "SI-" & pAffiliate.Trim & "-" & pShippingNo.Trim & " Shipping Instruction [TRIAL]"

            ls_Body = clsNotification.GetNotification("22", "", pShippingNo)
            'ls_Body = Replace(ls_Body, "Kanban", "Non Kaban")
            ls_Attachment = Trim(pPathFile) & "\" & pDN1

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, pDN1, IIf(pBarcodeFile = "", "", pBarcodeFile), , , , , True) = False Then
                sendEmailtoForwarder = False
                Exit Function
            End If

            sendEmailtoForwarder = True

        Catch ex As Exception
            sendEmailtoForwarder = False
            errMsg = "Process Send SI [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function
End Class
