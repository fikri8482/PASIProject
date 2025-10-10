Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Public Class clsTallyData
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
        Dim pConsigneeCode As String = ""

        Dim pFileName2 As String = ""
        Dim pDate As String = ""
        Dim pTotalCarton As Integer = 0
        Dim pDestination As String = ""

        Dim ds As New DataSet

        Try
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data SI")

            '---------------------------------------excel Kanban ---------------------------------------'
            ls_sql = " Select distinct TotalCtn = SUM(SHD.BoxQty), Consignee = isnull(ConsigneeCode,''), SHM.AffiliateID, SHM.ForwarderID, SHM.ShippingInstructionNo, ETDPort = Convert(Char(12), convert(Datetime, isnull(SHM.ETDPort,'')),106), DestinationPort From ShippingInstruction_Master SHM " & vbCrLf & _
                      " LEFT JOIN ShippingInstruction_Detail SHD ON SHM.AffiliateID = SHD.AffiliateID " & vbCrLf & _
                      " AND SHM.ForwarderID = SHD.ForwarderID " & vbCrLf & _
                      " AND SHM.ShippingInstructionNo = SHD.ShippingInstructionNo " & vbCrLf & _
                      " LEFT JOIN ReceiveForwarder_Detail RD ON RD.AffiliateID = SHM.AffiliateID " & vbCrLf & _
                      " AND SHD.SupplierID = RD.SupplierID " & vbCrLf & _
                      " AND RD.PartNo = SHD.PartNo " & vbCrLf & _
                      " and RD.OrderNo = SHD.OrderNo  " & vbCrLf & _
                      " AND RD.SuratJalanNo = SHD.SuratJalanno " & vbCrLf & _
                      " AND SHD.SupplierID = RD.SupplierID " & vbCrLf & _
                      " LEFT JOIN PO_Master_Export POM ON POM.PONo = RD.PONO and POM.OrderNo1 = RD.OrderNo " & vbCrLf & _
                      " AND POM.AffiliateID = RD.AffiliateID and RD.SupplierID = POM.SupplierID " & vbCrLf & _
                      " LEFT JOIN MS_Affiliate MA ON POM.AffiliateID = MA.AffiliateID "

            ls_sql = ls_sql + "  where isnull(SHM.ExcelCls,'') = '1'  "
            ls_sql = ls_sql + " Group by ConsigneeCode, SHM.AffiliateID, SHM.ForwarderID, SHM.ShippingInstructionNo, SHM.ETDPort, DestinationPort  "

            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pShippingInstruction = Trim(ds.Tables(0).Rows(xi)("ShippingInstructionNo"))
                    pForwarder = Trim(ds.Tables(0).Rows(xi)("ForwarderID"))
                    pConsigneeCode = Trim(ds.Tables(0).Rows(xi)("Consignee"))
                    pDate = Trim(ds.Tables(0).Rows(xi)("ETDPort"))
                    pDestination = Trim(ds.Tables(0).Rows(xi)("DestinationPort"))
                    pTotalCarton = Trim(ds.Tables(0).Rows(xi)("TotalCtn"))

                    pFileName2 = ""

                    '89. Create CSV
                    If CreateTallyBlank(GB, log, pAffiliate, pForwarder, pShippingInstruction, pConsigneeCode, pDate, pDestination, pTotalCarton, pAtttacment, pResult, pScreenName, pFileName2, LogName, errMsg, ErrSummary) = False Then
                        GoTo keluar
                    End If

                    If pFileName2 <> "" Then
                        If sendEmailtoForwarder(GB, pResult, pShippingInstruction, pAffiliate, pForwarder, errMsg, pFileName2, "") = False Then
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

    Shared Function CreateTallyBlank(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, _
                                   ByVal pAffiliate As String, ByVal pForwarderID As String, ByVal pShippingNo As String, _
                                   ByVal pConsignee As String, ByVal pETD As String, ByVal pDestination As String, _
                                   ByVal pTotalCtn As String, ByVal pAtttacment As String, ByVal pResult As String, _
                                   ByRef pScreenName As String, ByRef pFileName1 As String, _
                                   ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        CreateTallyBlank = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""
        Dim fromEmail As String = ""
        Dim receiptCCEmail As String = ""
        Const ColorYellow As Single = 65535

        Dim dsDetail As New DataSet
        Dim dsEmail As New DataSet

        Dim i As Integer

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template Tally.xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send Tally to Forwarder STOPPED, File Excel isn't Found"
                errSummary = "Process Send Tally to Forwarder STOPPED, File Excel isn't Found"
                CreateTallyBlank = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Tally [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "]")

            NewFileCopy = pAtttacment & "\Template Tally.xlsm"
            NewFileCopyTO = pResult & "\Template Tally " & Format(Now, "HHmmss") & ".xlsm"

            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
            Else
                System.IO.File.Copy(NewFileCopy, pResult & "\Delivery.xlsm")
            End If

            Dim ls_file As String = NewFileCopyTO

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetail = PrintTally(GB, pShippingNo, pAffiliate, pForwarderID)

            If dsDetail.Tables(0).Rows.Count > 0 Then
                log.WriteToProcessLog(Date.Now, pScreenName, "Input Header Tally [" & pAffiliate & "-" & pForwarderID & "-" & pShippingNo & "]")

                dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "SupplierDeliverycc", "SupplierDeliverycc", "SupplierDeliveryTO", errMsg)

                For y = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If dsEmail.Tables(0).Rows(y)("flag") = "PASI" Then
                        fromEmail = dsEmail.Tables(0).Rows(y)("EmailFrom")
                        receiptCCEmail = dsEmail.Tables(0).Rows(y)("EmailCC")
                    End If
                Next

                ExcelSheet.Range("H1").Value = "TALLY"
                ExcelSheet.Range("H2").Value = fromEmail
                ExcelSheet.Range("H3").Value = Trim(pConsignee)
                ExcelSheet.Range("H4").Value = Trim(pForwarderID)
                ExcelSheet.Range("I8").Value = Trim(pShippingNo)

                ExcelSheet.Range("AA12").Value = Trim(pETD)
                ExcelSheet.Range("AA16").Value = Trim(pDestination)
                ExcelSheet.Range("I18").Value = Trim(pTotalCtn)

                ExcelSheet.Range("S1").Value = Trim(pAffiliate)
                ExcelSheet.Range("S1").Font.Color = Color.White

                ExcelSheet.Range("Y2").Value = ""


                For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                    'Header
                    ExcelSheet.Range("B" & i + 23 & ": C" & i + 23).Merge()
                    ExcelSheet.Range("B" & i + 23 & ": C" & i + 23).Value = i + 1
                    ExcelSheet.Range("B" & i + 23 & ": C" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).Merge()
                    ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).Interior.Color = ColorYellow

                    ExcelSheet.Range("I" & i + 23 & ": N" & i + 23).Merge()
                    ExcelSheet.Range("I" & i + 23 & ": N" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("OrderNo"))
                    ExcelSheet.Range("I" & i + 23 & ": N" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                    ExcelSheet.Range("O" & i + 23 & ": U" & i + 23).Merge()
                    ExcelSheet.Range("O" & i + 23 & ": U" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                    ExcelSheet.Range("O" & i + 23 & ": U" & i + 23).Merge()

                    ExcelSheet.Range("V" & i + 23 & ": AD" & i + 23).Merge()
                    ExcelSheet.Range("V" & i + 23 & ": AD" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                    ExcelSheet.Range("V" & i + 23 & ": AD" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                    ExcelSheet.Range("AE" & i + 23 & ": AJ" & i + 23).Merge()
                    ExcelSheet.Range("AE" & i + 23 & ": AJ" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("CaseNo1"))
                    ExcelSheet.Range("AE" & i + 23 & ": AJ" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                    ExcelSheet.Range("AK" & i + 23 & ": AP" & i + 23).Merge()
                    ExcelSheet.Range("AK" & i + 23 & ": AP" & i + 23).Value = Trim(dsDetail.Tables(0).Rows(i)("CaseNo2"))
                    ExcelSheet.Range("AK" & i + 23 & ": AP" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                    ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).Merge()
                    ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    ExcelSheet.Range("AQ" & i + 23 & ": AS" & i + 23).NumberFormat = "#,##0.00"

                    ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).Merge()
                    ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).NumberFormat = "#,##0.00"

                    ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).Merge()
                    ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).NumberFormat = "#,##0.00"

                    ExcelSheet.Range("D" & i + 23 & ": H" & i + 23).Merge()
                    ExcelSheet.Range("AT" & i + 23 & ": AV" & i + 23).Merge()

                    ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).Merge()
                    ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).Interior.Color = ColorYellow
                    ExcelSheet.Range("BC" & i + 23 & ": BE" & i + 23).Interior.Color = ColorYellow
                    ExcelSheet.Range("AW" & i + 23 & ": AY" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    ExcelSheet.Range("AE" & i + 23 & ": AY" & i + 23).Interior.Color = ColorYellow

                    ExcelSheet.Range("BC" & i + 23 & ": BE" & i + 23).Merge()
                    ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).Merge()
                    ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).FormulaR1C1 = "=RC[-9]*RC[-6]*RC[-3]"
                    ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    ExcelSheet.Range("AZ" & i + 23 & ": BB" & i + 23).NumberFormat = "#,##0.00"
                    clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & i + 23 & ": BE" & i + 23))
                Next

                ExcelSheet.Range("B" & i + 23).Value = "E"
                ExcelSheet.Range("B" & i + 23).Interior.Color = Color.Black
                ExcelSheet.Range("B" & i + 23).Font.Color = Color.White
                ExcelSheet.Range("B24").Font.Color = Color.Black
                ExcelSheet.Range("B24").Interior.Color = Color.White

                'Save ke Local
                xlApp.DisplayAlerts = False
                pFileName1 = pResult & "\Tally Data " & Trim(pAffiliate) & "-" & Trim(pForwarderID) & "-" & Trim(pShippingNo) & ".xlsm"

                ExcelBook.SaveAs(pFileName1)

                xlApp.Workbooks.Close()
                xlApp.Quit()
            End If
        Catch ex As Exception
            CreateTallyBlank = False

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

    Shared Function PrintTally(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value1 As String, ByVal ls_value2 As String, ByVal ls_value3 As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " select distinct " & vbCrLf & _
                  " InvoiceNo = SHM.ShippingInstructionNo, " & vbCrLf & _
                  " OrderNo = SHD.OrderNo, " & vbCrLf & _
                  " PartNo = SHD.PartNo, " & vbCrLf & _
                  " PartName = MP.PartGroupName, " & vbCrLf & _
                  " CaseNo1 = Label1,  " & vbCrLf & _
                  " CaseNo2 = Label2,  " & vbCrLf &
                  " Length = length, " & vbCrLf & _
                  " Width = Width, " & vbCrLf & _
                  " Height = Height, " & vbCrLf & _
                  " ForwarderID = SHM.ForwarderID " & vbCrLf & _
                  " From ShippingInstruction_master SHM  "

        ls_SQL = ls_SQL + " LEFT JOIN ShippingInstruction_Detail SHD " & vbCrLf & _
                          " ON SHM.ShippingInstructionNo = SHD.ShippingInstructionNo " & vbCrLf & _
                          " AND SHM.AffiliateID = SHD.AffiliateID " & vbCrLf & _
                          " AND SHM.ForwarderID = SHD.ForwarderID " & vbCrLf & _
                          " LEFT JOIN ReceiveForwarder_DetailBox RB " & vbCrLf & _
                          " ON RB.SuratJalanNo = SHD.SuratJalanNo  " & vbCrLf & _
                          "AND RB.AffiliateID = SHD.AffiliateID " & vbCrLf & _
                          "AND RB.SupplierID = SHD.SupplierID " & vbCrLf & _
                          "AND RB.OrderNo = SHD.OrderNo " & vbCrLf & _
                          "AND RB.PartNo = SHD.PartNo " & vbCrLf & _
                          "AND StatusDefect = '0' " & vbCrLf

        ls_SQL = ls_SQL + " LEFT JOIN MS_Parts MP ON MP.Partno = SHD.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SHD.PartNo " & vbCrLf & _
                          " AND MPM.AffiliateID = SHM.AffiliateID " & vbCrLf & _
                          " AND MPM.SupplierID = SHD.SupplierID " & vbCrLf & _
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
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE ShippingInstructionNo = '" & pShippingNo & "' " & vbCrLf & _
                      " AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                      " AND ForwarderID = '" & pForwarder & "'"

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send Tally Blank: ShippingInstructionNo [" & pShippingNo.Trim & "], AffiliateID [" & pAffiliate.Trim & "], ForwarderID [" & pForwarder.Trim & "] to Forwarder STOPPED, because " & ex.Message
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
                errMsg = "Process Send Tally [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoForwarder = False
                errMsg = "Process Send Tally [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "TA-" & pAffiliate.Trim & "-" & pShippingNo.Trim & " Tally Data [TRIAL]"

            ls_Body = clsNotification.GetNotification("21", "", pShippingNo)
            ls_Attachment = Trim(pPathFile) & "\" & pDN1

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, pDN1, IIf(pBarcodeFile = "", "", pBarcodeFile), , , , , True) = False Then
                sendEmailtoForwarder = False
                Exit Function
            End If

            sendEmailtoForwarder = True

        Catch ex As Exception
            sendEmailtoForwarder = False
            errMsg = "Process Send Tally [" & pShippingNo & "] from Affiliate [" & pAffiliate & "] to Forwarder [" & pForwarder & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function
End Class
