Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class clsInvoicePakcingList
    Shared Sub up_SendInvoice(ByVal cfg As GlobalSetting.clsConfig,
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
        Dim pTerm As String = ""
        Dim pService As String = ""

        Dim pFileName1 As String = ""

        Dim ds As New DataSet

        Try
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data Invoice")

            '---------------------------------------excel Kanban ---------------------------------------'
            ls_sql = " Select distinct ShippingInstructionNo, AffiliateID, ForwarderID, TermDelivery, TypeOfService From ShippingInstruction_Master where isnull(sendInvoice,'') = '1' "

            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pShippingInstruction = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))
                    pForwarder = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    pTerm = Trim(ds.Tables(0).Rows(xi)("TermDelivery"))
                    pService = Trim(ds.Tables(0).Rows(xi)("TypeOfService"))

                    pFileName1 = ""

                    If pTerm = "1" Or pTerm = "2" Then
                        pTerm = "FCA"
                    ElseIf pTerm = "3" Or pTerm = "4" Then
                        pTerm = "CIF"
                    ElseIf pTerm = "5" Then
                        pTerm = "DDU PASI"
                    ElseIf pTerm = "6" Then
                        pTerm = "DDU Affiliate"
                    ElseIf pTerm = "7" Then
                        pTerm = "EX-Work"
                    ElseIf pTerm = "8" Then
                        pTerm = "FOB"
                    End If

                    '88. Create PDF SI
                    log.WriteToProcessLog(Date.Now, pScreenName, "Start Create Invoice File, ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "]")
                    If CreateInvoiceToPDF(GB, cfg, pAffiliate, pForwarder, pShippingInstruction, pTerm, pService, pFileName1, pResult, errMsg) = False Then
                        log.WriteToProcessLog(Date.Now, pScreenName, "End Create Invoice File, ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "] " & errMsg)
                        If errMsg = "Microsoft.VisualBasic.ErrObject" Then
                            End
                        End If
                        GoTo keluar
                    End If
                    log.WriteToProcessLog(Date.Now, pScreenName, "End Create Invoice File, ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "]")

                    If pFileName1 <> "" Then
                        If sendEmailtoForwarder(GB, pResult, pShippingInstruction, pAffiliate, pForwarder, errMsg, "", pFileName1) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Forwarder. ShippingInstructionNo [" & pShippingInstruction & "], ForwarderID [" & pForwarder & "], Affiliate [" & pAffiliate & "] ok.")
                        End If
                    End If

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send Invoice [" & pShippingInstruction & "], Forwarder [" & pForwarder & "], Affiliate [" & pAffiliate & "] ok", LogName)
                    LogName.Refresh()

                    Call UpdateSendInvoice(pShippingInstruction, pAffiliate, pForwarder, errMsg)
keluar:
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "Invoice [" & pAffiliate & "-" & pForwarder & "-" & pShippingInstruction & "] " & ex.Message
            ErrSummary = "Invoice [" & pAffiliate & "-" & pForwarder & "-" & pShippingInstruction & "] " & ex.Message
        Finally
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    Shared Function CreateInvoiceToPDF(ByVal GB As GlobalSetting.clsGlobal, ByVal cfg As GlobalSetting.clsConfig, ByVal pAffiliate As String, ByVal pForwarderID As String, ByVal pShippingNo As String, ByVal pTerm As String, ByVal pService As String, ByRef pFileName As String, ByVal pPathFile As String, ByRef errMsg As String) As Boolean
        Dim CrReport As New Invoice()
        Dim CrExportOptions As ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

        Try
            Cursor.Current = Cursors.WaitCursor
            Dim dsPrint As New DataSet
            dsPrint = PrintSI(GB, pShippingNo, pAffiliate, pForwarderID, pTerm, pService)

            If IsNothing(dsPrint) Then
                Exit Try
            End If

            If dsPrint.Tables(0).Rows.Count = 0 Then Exit Try

            CrReport.SetDatabaseLogon(cfg.User, cfg.Password, cfg.Server, cfg.Database)
            CrReport.SetDataSource(dsPrint.Tables(0))

            Dim ls_FileName = pPathFile & "\Invoice-" & pShippingNo.Trim & "-" & pAffiliate.Trim & "-" & pForwarderID.Trim & ".pdf"
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
                CreateInvoiceToPDF = False
            End Try
            'PDF
            CreateInvoiceToPDF = True
        Catch ex As Exception
            errMsg = Err.ToString()
            CreateInvoiceToPDF = False
        Finally
            If Not CrReport Is Nothing Then
                clsGeneral.NAR(CrReport)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If
        End Try
    End Function

    Shared Function PrintSI(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String, ByVal pTerm As String, ByVal pService As String) As DataSet
        Dim ls_SQL As String = ""

        Dim tentukanBoat As String = pService.ToString.Trim
        Dim tentukanTerm As String = pTerm.ToString.Trim

        Dim PriceCls As String = 0

        If tentukanTerm = "FCA" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            PriceCls = "2"
        ElseIf tentukanTerm = "FCA" Then
            PriceCls = "1"
        End If

        If tentukanTerm = "CIF" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            PriceCls = "4"
        ElseIf tentukanTerm = "CIF" Then
            PriceCls = "3"
        End If

        If tentukanTerm = "DDU PASI" Then
            PriceCls = "5"
        ElseIf tentukanTerm = "DDU Affiliate" Then
            PriceCls = "6"
        ElseIf tentukanTerm = "EX-Work" Then
            PriceCls = "7"
        ElseIf tentukanTerm = "FOB" Then
            PriceCls = "8"
        End If

        ls_SQL = "  select distinct  " & vbCrLf & _
              "  buyer = Rtrim(MA.BuyerName) + CHAR(13)+CHAR(10) + Rtrim(MA.BuyerAddress),  " & vbCrLf & _
              "  Consignee = Rtrim(Coalesce(MA.ConsigneeName, MA.AffiliateName)) + CHAR(13)+CHAR(10) + Rtrim(coalesce(MA.ConsigneeAddress, Rtrim(MA.Address) + Rtrim(MA.City) )),  " & vbCrLf & _
              "  Attn = ISNULL(ma.Att,''),  " & vbCrLf & _
              "  Vessel = ISNULL(SHM.Vessels,''), " & vbCrLf & _
              "  Fromto = Isnull(MF.Port,''),  " & vbCrLf & _
              "  Toto = isnull(MA.DestinationPort,''),  " & vbCrLf & _
              "  About = Convert(Char(12), convert(Datetime, isnull(SHM.ETAPort,POM.ETAPort1)),106),  " & vbCrLf & _
              "  ONAbout = Convert(Char(12), convert(Datetime, isnull(SHM.ETDPort,POM.ETDPort1)),106),  " & vbCrLf & _
              "  Via = SHM.Via,  " & vbCrLf & _
              "  InvoiceNo = SHM.ShippingInstructionNo,   "

        ls_SQL = ls_SQL + "  OrderNo = (SELECT (STUFF((SELECT distinct ', ' + RTrim(ShippingInstruction_Detail.orderNo) FROM ShippingInstruction_Detail WHERE ShippingInstructionNo = '" & ls_value1 & "' AND AffiliateID = '" & ls_value2 & "' AND ForwarderID = '" & ls_value3 & "' FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          "  InvDate = Convert(Char(12), convert(Datetime, isnull(SHM.ShippingInstructionDate,'')),106),  " & vbCrLf & _
                          "  Place = 'JAKARTA',  " & vbCrLf & _
                          "  Privilege = '',  " & vbCrLf & _
                          "  AWB = '',  " & vbCrLf & _
                          "  ContainerNo = '', --TM.ContainerNo,  " & vbCrLf & _
                          "  Insurance = '',  " & vbCrLf & _
                          "  Remarks = '',  " & vbCrLf & _
                          "  paymentTerm = Isnull(MA.PaymentTerm,''),  " & vbCrLf & _
                          "  Marks = '',--Description = '',  " & vbCrLf & _
                          "  QtyBox = SHD.QtyBox,   " & vbCrLf

        ls_SQL = ls_SQL + "  Qty = RB.Box,  " & vbCrLf & _
                          "  Price = isnull(Price,0),   " & vbCrLf & _
                          "  Amount = 0,  " & vbCrLf & _
                          "  Net =  (isnull(NetWeight,0)/1000),  " & vbCrLf & _
                          "  Gross =(isnull(SHM.GrossWeight,0)/1000),  " & vbCrLf & _
                          "  DocNo = '',  " & vbCrLf & _
                          "  RevNo = '',  " & vbCrLf & _
                          "  partCust = isnull(PartGroupName,''),  " & vbCrLf & _
                          "  PartYazaki = SHD.PartNo,  " & vbCrLf & _
                          "  CaseNo = Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2),  " & vbCrLf & _
                          "  totalCarton = 0 , " & vbCrLf & _
                          "  Term = CASE WHEN RTRIM(MPC.Description) = 'FCA - BOAT' THEN 'FCA' " & vbCrLf & _
                          "  WHEN RTRIM(MPC.Description) = 'FCA - AIR' THEN 'FCA'	 " & vbCrLf & _
                          "  WHEN RTRIM(MPC.Description) = 'CIF - BOAT' THEN 'CIF' " & vbCrLf & _
                          "  WHEN RTRIM(MPC.Description) = 'CIF - AIR' THEN 'CIF' " & vbCrLf & _
                          "  ELSE RTRIM(MPC.Description) END,	" & vbCrLf & _
                          "  SHM.TotalPallet, SHM.Measurement, SHM.GrossWeight, SHM.Freight, CASE WHEN ISNULL(HSCodeCls,'0')  = '0'  THEN '' else HSCode END HSCode, SHD.OrderNo NewOrderNo, SHM.TypeOfService,  SHM.NamaKapalS  From  " & vbCrLf

        ls_SQL = ls_SQL + "  ShippingInstruction_Detail SHD   " & vbCrLf & _
                          "  LEFT JOIN ShippingInstruction_Master SHM ON SHM.ShippingInstructionNo = SHD.ShippingInstructionNo  " & vbCrLf & _
                          "  AND SHM.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  AND SHM.ForwarderID = SHD.ForwarderID  " & vbCrLf & _
                          "  LEFT JOIN Tally_Master TM ON TM.ShippingInstructionNo = SHM.ShippingInstructionNo and TM.AffiliateID = SHM.AffiliateID and TM.ForwarderID = SHM.ForwarderID " & vbCrLf & _
                          "  LEFT JOIN MS_Parts MP ON MP.PartNo = SHD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.Partno = SHD.PartNo and MPM.AffiliateID = SHD.AffiliateID and MPM.SupplierID = SHD.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SHD.ForwarderID  LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_master RM ON RM.SuratJalanNo = SHD.SuratJalanno  AND RM.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  AND RM.PONO = SHD.OrderNo  " & vbCrLf & _
                          "  AND SHD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNO = RM.SuratJalanno  " & vbCrLf

        ls_SQL = ls_SQL + "  AND RD.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                          "  AND RD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "  AND RD.PONo = RM.PONO  " & vbCrLf & _
                          "  AND RD.OrderNO = Rm.OrderNo  " & vbCrLf & _
                          "  AND RD.PartNo = SHD.PartNo " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = SHD.SuratJalanNo  " & vbCrLf & _
                          "  AND RB.SupplierID = SHD.SupplierID   " & vbCrLf & _
                          "  AND RB.AffiliateID = SHD.AffiliateID   " & vbCrLf & _
                          "  --AND RB.PONo = RD.PONo   " & vbCrLf & _
                          "  AND RB.OrderNo = SHD.OrderNo   " & vbCrLf & _
                          "  AND RB.PartNo = SHD.PartNo   " & vbCrLf

        ls_SQL = ls_SQL + "  AND RB.StatusDefect = '0'   " & vbCrLf & _
                          "  LEFT JOIN PO_Detail_Export POD ON POD.PONo = RD.PONO  " & vbCrLf & _
                          "  AND POD.OrderNo1 = RD.OrderNo  " & vbCrLf & _
                          "  AND POD.AffiliateID = RD.AffiliateID  AND POD.SupplierID = RD.SupplierID  " & vbCrLf & _
                          "  AND POD.PartNO = RD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN PO_Master_export POM ON POM.PONo = POD.PONO  " & vbCrLf & _
                          "  AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
                          "  AND POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN MS_PriceCls MPC ON MPC.PriceCls = SHM.TermDelivery " & vbCrLf & _
                          "  LEFT JOIN MS_Price MPR ON MPR.PartNO = SHD.PartNo  " & vbCrLf & _
                          "  AND MPR.AffiliateID = SHD.AffiliateID  " & vbCrLf

        ls_SQL = ls_SQL + "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112)  " & vbCrLf & _
                          "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)  " & vbCrLf & _
                          "  AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "'" & vbCrLf & _
                          "  WHERE SHM.ShippingInstructionNo = '" & Trim(ls_value1) & "'  " & vbCrLf & _
                          "  AND SHM.AffiliateID = '" & Trim(ls_value2) & "' " & vbCrLf & _
                          "  AND SHM.ForwarderID = '" & Trim(ls_value3) & "'  order by SHD.partno, Rtrim(RB.Label1) + '-' + Rtrim(Rb.Label2)  "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Sub UpdateSendInvoice(ByVal pShippingNo As String, ByVal pAffiliate As String, ByVal pForwarder As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & vbCrLf & _
                      " SET sendInvoice='2'" & vbCrLf & _
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
                errMsg = "Process Send Invoice [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoForwarder = False
                errMsg = "Process Send Invoice [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "INV-" & pAffiliate.Trim & "-" & pShippingNo.Trim & " TAX INVOICE [TRIAL]"

            ls_Body = clsNotification.GetNotification("22", "", pShippingNo)

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
