Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsReceivingAffiliate
    Shared Sub up_SendReceivingAffiliate(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")
        Dim xi As Integer

        Dim ds As New DataSet

        Dim ls_SQL As String = ""
        Dim pAffiliate As String = ""
        Dim pSuratJalanNo As String = ""
        Dim pDeliveryCls As String = ""
        Dim temp_Filename1 As String = "", temp_Filename2 As String = ""

        Try
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data GR Affiliate")

            ls_SQL = "SELECT distinct a.*, c.DeliveryByPASICls FROM ReceiveAffiliate_Master a " & vbCrLf & _
                     "left join ReceiveAffiliate_Detail b on a.AffiliateID = b.AffiliateID and a.SuratJalanNo = b.SuratJalanNo " & vbCrLf & _
                     "left join PO_Master c on b.PONo = c.PONo and b.AffiliateID = c.AffiliateID and b.SupplierID = c.SupplierID " & vbCrLf & _
                     "WHERE ExcelCls='1'"
            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pSuratJalanNo = Trim(ds.Tables(0).Rows(xi)("SuratJalanNo"))
                    pDeliveryCls = ds.Tables(0).Rows(0)("DeliveryByPASICls")

                    '01. Create data GR
                    If up_CreateGRAffiliate(GB, log, pSuratJalanNo, pAffiliate, pDeliveryCls, pAtttacment, pResult, pScreenName, temp_Filename1, LogName, errMsg, ErrSummary) = False Then
                        GoTo keluar
                    End If

                    ''02. Create data Invoice
                    'If up_CreateInvoiceAffiliate(GB, log, pSuratJalanNo, pAffiliate, pDeliveryCls, pAtttacment, pResult, pScreenName, temp_Filename2, LogName, errMsg, ErrSummary) = False Then
                    '    GoTo keluar
                    'End If

                    '03. Send Email to Affiliate
                    If sendEmailtoAffiliate(GB, pResult, temp_Filename1, temp_Filename2, pSuratJalanNo, pAffiliate, errMsg) = False Then
                        Exit Try
                    Else
                        log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to PASI. Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] ok.")
                    End If

                    Call UpdateExcelReceivingAffiliate(pAffiliate, pSuratJalanNo, errMsg)

                    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email. GR & Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] ok.")
                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send GR & Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] ok", LogName)
                    LogName.Refresh()
keluar:
                Next
            End If
        Catch ex As Exception
            errMsg = "GR & Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] " & ex.Message
            ErrSummary = "GR & Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] " & ex.Message
        Finally
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    Shared Function up_CreateGRAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, ByVal pSuratJalanNo As String, ByVal pAffiliate As String, ByVal pDeliveryCls As String, ByVal pAtttacment As String, ByVal pResult As String, ByRef pScreenName As String, ByRef pFileName As String, ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        up_CreateGRAffiliate = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""

        Dim dsDetail As New DataSet

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template GoodReceivingAffiliate.xlsx") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send GR Affiliate to PASI STOPPED, File Excel isn't Found"
                errSummary = "Process Send GR Affiliate to PASI STOPPED, File Excel isn't Found"
                up_CreateGRAffiliate = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "]")

            NewFileCopy = pAtttacment & "\Template GoodReceivingAffiliate.xlsx"
            NewFileCopyTO = pResult & "\Template GoodReceivingAffiliate " & Format(Now, "yyyyMMddHHmmss") & ".xlsx"

            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
            Else
                System.IO.File.Copy(NewFileCopy, pResult & "\GoodReceivingAffiliate.xlsx")
            End If

            Dim ls_file As String = NewFileCopyTO

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetail = bindDataReceivingAffiliate(GB, pAffiliate, pSuratJalanNo, pDeliveryCls)

            If dsDetail.Tables(0).Rows.Count > 0 Then
                log.WriteToProcessLog(Date.Now, pScreenName, "Input Header Excel Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "]")
                ExcelSheet.Range("M8").Value = dsDetail.Tables(0).Rows(0)("SuratJalanNo") & ""
                ExcelSheet.Range("M9").Value = Format(dsDetail.Tables(0).Rows(0)("DeliveryDate"), "dd-MMM-yyyy")
                ExcelSheet.Range("M10").Value = Format(dsDetail.Tables(0).Rows(0)("PlanDeliveryDate"), "dd-MMM-yyyy")
                ExcelSheet.Range("M11").Value = Format(dsDetail.Tables(0).Rows(0)("ReceiveDate"), "dd-MMM-yyyy")
                ExcelSheet.Range("M12").Value = Trim(dsDetail.Tables(0).Rows(0)("DeliveryLocationName") & "") & ""
                ExcelSheet.Range("M13").Value = Trim(dsDetail.Tables(0).Rows(0)("DeliveryAddress") & "") & ""
                ExcelSheet.Range("M13").WrapText = True

                ExcelSheet.Range("AS8").Value = Trim(dsDetail.Tables(0).Rows(0)("JenisArmada")) & ""
                ExcelSheet.Range("AS9").Value = Trim(dsDetail.Tables(0).Rows(0)("NoPol")) & ""
                ExcelSheet.Range("AS10").Value = Trim(dsDetail.Tables(0).Rows(0)("DriverName") & "")
                ExcelSheet.Range("AS11").Value = dsDetail.Tables(0).Rows(0)("TotalBox") & ""

                log.WriteToProcessLog(Date.Now, pScreenName, "Input Detail Excel Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "]")
                For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                    ExcelSheet.Range("B" & i + 19 & ": C" & i + 19).Merge()
                    ExcelSheet.Range("D" & i + 19 & ": I" & i + 19).Merge()
                    ExcelSheet.Range("J" & i + 19 & ": L" & i + 19).Merge()
                    ExcelSheet.Range("M" & i + 19 & ": P" & i + 19).Merge()
                    ExcelSheet.Range("Q" & i + 19 & ": U" & i + 19).Merge()
                    ExcelSheet.Range("V" & i + 19 & ": AD" & i + 19).Merge()
                    ExcelSheet.Range("AE" & i + 19 & ": AF" & i + 19).Merge()
                    ExcelSheet.Range("AG" & i + 19 & ": AH" & i + 19).Merge()

                    ExcelSheet.Range("AI" & i + 19 & ": AL" & i + 19).Merge()
                    ExcelSheet.Range("AM" & i + 19 & ": AP" & i + 19).Merge()
                    ExcelSheet.Range("AQ" & i + 19 & ": AT" & i + 19).Merge()
                    ExcelSheet.Range("AU" & i + 19 & ": AX" & i + 19).Merge()
                    ExcelSheet.Range("AY" & i + 19 & ": BB" & i + 19).Merge()

                    ExcelSheet.Range("B" & i + 19 & ": C" & i + 19).Value = i + 1
                    ExcelSheet.Range("B" & i + 19 & ": C" & i + 19).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ExcelSheet.Range("B" & i + 19 & ": C" & i + 19).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    ExcelSheet.Range("D" & i + 19 & ": I" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("PONo") & "")
                    ExcelSheet.Range("J" & i + 19 & ": L" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("POKanbanCls") & "")
                    ExcelSheet.Range("M" & i + 19 & ": P" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("KanbanNo") & "")

                    ExcelSheet.Range("Q" & i + 19 & ": U" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo") & "")
                    ExcelSheet.Range("V" & i + 19 & ": AD" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName") & "")
                    ExcelSheet.Range("AE" & i + 19 & ": AF" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("UnitDesc") & "")
                    ExcelSheet.Range("AG" & i + 19 & ": AH" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("QtyBox") & "")
                    ExcelSheet.Range("AG" & i + 19 & ": AH" & i + 19).NumberFormat = "#,##0"

                    ExcelSheet.Range("AI" & i + 19 & ": AL" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("SupplierDeliveryQty") & "")
                    ExcelSheet.Range("AI" & i + 19 & ": AL" & i + 19).NumberFormat = "#,##0"

                    ExcelSheet.Range("AM" & i + 19 & ": AP" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("RecQty") & "")
                    ExcelSheet.Range("AM" & i + 19 & ": AP" & i + 19).NumberFormat = "#,##0"

                    ExcelSheet.Range("AQ" & i + 19 & ": AT" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("DefectQty") & "")
                    ExcelSheet.Range("AQ" & i + 19 & ": AT" & i + 19).NumberFormat = "#,##0"

                    ExcelSheet.Range("AU" & i + 19 & ": AX" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("RemainingQty") & "")
                    ExcelSheet.Range("AU" & i + 19 & ": AX" & i + 19).NumberFormat = "#,##0"

                    ExcelSheet.Range("AY" & i + 19 & ": BB" & i + 19).Value = Trim(dsDetail.Tables(0).Rows(i)("ReceivingBox") & "")
                    ExcelSheet.Range("AY" & i + 19 & ": BB" & i + 19).NumberFormat = "#,##0"

                    clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & i + 19 & ": BB" & i + 19))
                Next

                ExcelSheet.EnableSelection = XlEnableSelection.xlNoRestrictions

                'ExcelSheet.Protect("tosis123", , , , , , , , , , , , , True)

                xlApp.DisplayAlerts = False

                Dim temp_Filename As String = "GoodReceivingAffiliate " & Trim(pAffiliate) & "-" & Trim(Replace(pSuratJalanNo, "/", "-")) & ".xlsx"
                pFileName = pResult & "\" & temp_Filename
                ExcelBook.SaveAs(pFileName)
                ExcelBook.Close()
                xlApp.Workbooks.Close()
                xlApp.Quit()

                My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

                log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] ok.")
            End If
        Catch ex As Exception
            up_CreateGRAffiliate = False

            log.WriteToErrorLog(pScreenName, "Process Create GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Create GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Create GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, LogName)
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

    Shared Function up_CreateInvoiceAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, ByVal pSuratJalanNo As String, ByVal pAffiliate As String, ByVal pDeliveryCls As String, ByVal pAtttacment As String, ByVal pResult As String, ByRef pScreenName As String, ByRef pFileName As String, ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        up_CreateInvoiceAffiliate = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""
        Dim fromEmail As String = ""
        Dim receiptCCEmail As String = ""

        Dim dsDetail As New DataSet
        Dim dsEmail As New DataSet
        Dim dsAffp As New DataSet

        Dim i As Integer

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template Invoice.xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send GR Affiliate to PASI STOPPED, File Excel isn't Found"
                errSummary = "Process Send GR Affiliate to PASI STOPPED, File Excel isn't Found"
                up_CreateInvoiceAffiliate = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "]")

            NewFileCopy = pAtttacment & "\Template Invoice.xlsm"
            NewFileCopyTO = pResult & "\Template Invoice" & Format(Now, "yyyyMMddHHmmss") & ".xlsm"

            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
            Else
                System.IO.File.Copy(NewFileCopy, pResult & "\Invoice.xlsm")
            End If

            Dim ls_file As String = NewFileCopyTO

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetail = bindDataReceivingAffiliate(GB, pAffiliate, pSuratJalanNo, pDeliveryCls)

            If dsDetail.Tables(0).Rows.Count > 0 Then
                log.WriteToProcessLog(Date.Now, pScreenName, "Input Header Excel Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "]")

                dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

                For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                        fromEmail = dsEmail.Tables(0).Rows(i)("EmailFrom")
                        receiptCCEmail = dsEmail.Tables(0).Rows(i)("EmailCC")
                    End If
                Next

                receiptCCEmail = Replace(receiptCCEmail, ",", ";")

                ExcelSheet.Range("H1").Value = "INV"
                ExcelSheet.Range("H2").Value = fromEmail
                ExcelSheet.Range("H3").Value = pAffiliate
                ExcelSheet.Range("H4").Value = dsDetail.Tables(0).Rows(0)("DeliveryLocationCode") & ""
                ExcelSheet.Range("H5").Value = ""

                ExcelSheet.Range("W11").Value = Format(Now, "dd-MMM-yy")

                dsAffp = clsGeneral.Affiliate(GB, Trim(pAffiliate))
                ExcelSheet.Range("I13").Value = dsAffp.Tables(0).Rows(0)("AffiliateName")
                ExcelSheet.Range("I14").Value = dsAffp.Tables(0).Rows(0)("Address")
                ExcelSheet.Range("I14").WrapText = True

                ExcelSheet.Range("AV37:BD37").UnMerge()

                Dim totalAmount As Double = 0

                log.WriteToProcessLog(Date.Now, pScreenName, "Input Detail Excel Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "]")
                For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                    ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Merge()
                    ExcelSheet.Range("D" & i + 36 & ": I" & i + 36).Merge()
                    ExcelSheet.Range("J" & i + 36 & ": P" & i + 36).Merge()
                    ExcelSheet.Range("Q" & i + 36 & ": T" & i + 36).Merge()
                    ExcelSheet.Range("U" & i + 36 & ": Y" & i + 36).Merge()
                    ExcelSheet.Range("Z" & i + 36 & ": AH" & i + 36).Merge()
                    ExcelSheet.Range("AI" & i + 36 & ": AJ" & i + 36).Merge()

                    ExcelSheet.Range("AK" & i + 36 & ": AL" & i + 36).Merge()
                    ExcelSheet.Range("AM" & i + 36 & ": AP" & i + 36).Merge()
                    ExcelSheet.Range("AQ" & i + 36 & ": AS" & i + 36).Merge()

                    ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).Merge()
                    ExcelSheet.Range("AV" & i + 36 & ": AY" & i + 36).Merge()
                    ExcelSheet.Range("AZ" & i + 36 & ": BD" & i + 36).Merge()

                    ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Value = i + 1
                    ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    ExcelSheet.Range("D" & i + 36 & ": I" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("SuratJalanNo") & "")
                    ExcelSheet.Range("J" & i + 36 & ": P" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PONo") & "")
                    ExcelSheet.Range("Q" & i + 36 & ": T" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("KanbanNo") & "")

                    ExcelSheet.Range("U" & i + 36 & ": Y" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo") & "")

                    ExcelSheet.Range("Z" & i + 36 & ": AH" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName") & "")
                    ExcelSheet.Range("AI" & i + 36 & ": AJ" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("UnitDesc") & "")
                    ExcelSheet.Range("AK" & i + 36 & ": AL" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("QtyBox") & "")
                    ExcelSheet.Range("AK" & i + 36 & ": AL" & i + 36).NumberFormat = "#,##0"

                    ExcelSheet.Range("AM" & i + 36 & ": AP" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("SupplierDeliveryQty") & "")
                    ExcelSheet.Range("AM" & i + 36 & ": AP" & i + 36).NumberFormat = "#,##0"

                    ExcelSheet.Range("AQ" & i + 36 & ": AS" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("RecQty") & "")
                    ExcelSheet.Range("AQ" & i + 36 & ": AS" & i + 36).NumberFormat = "#,##0"

                    ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("CurrDesc") & "")
                    ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                    ExcelSheet.Range("AV" & i + 36 & ": AY" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("Price") & "")
                    ExcelSheet.Range("AV" & i + 36 & ": AY" & i + 36).NumberFormat = "#,##0"

                    ExcelSheet.Range("AZ" & i + 36 & ": BD" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("Amount") & "")
                    ExcelSheet.Range("AZ" & i + 36 & ": BD" & i + 36).NumberFormat = "#,##0"
                    totalAmount = totalAmount + dsDetail.Tables(0).Rows(i)("Amount")

                    clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & i + 36 & ": BD" & i + 36))
                    ExcelSheet.Range("B" & i + 36 & ": BD" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    ExcelSheet.Range("B" & i + 36 & ": BD" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                Next

                ExcelSheet.Range("B38").Interior.Color = Color.White
                ExcelSheet.Range("B38").Font.Color = Color.Black
                ExcelSheet.Range("B" & i + 36).Value = "E"
                ExcelSheet.Range("B" & i + 36).Interior.Color = Color.Black
                ExcelSheet.Range("B" & i + 36).Font.Color = Color.White

                ExcelSheet.Range("AQ" & i + 36 & ": AS" & i + 36).Merge()
                ExcelSheet.Range("AQ" & i + 36 & ": AS" & i + 36).Value = "TOTAL"

                ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).Merge()
                ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(0)("CurrDesc") & "")
                ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ExcelSheet.Range("AT" & i + 36 & ": AU" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                ExcelSheet.Range("AV" & i + 36 & ": BD" & i + 36).Merge()
                ExcelSheet.Range("AV" & i + 36 & ": BD" & i + 36).Value = totalAmount
                ExcelSheet.Range("AV" & i + 36 & ": BD" & i + 36).NumberFormat = "#,##0"

                clsGeneral.DrawAllBorders(ExcelSheet.Range("AQ" & i + 36 & ": BD" & i + 36))
                ExcelSheet.Range("AQ" & i + 36 & ": BD" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ExcelSheet.Range("AQ" & i + 36 & ": BD" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                ExcelSheet.EnableSelection = XlEnableSelection.xlNoRestrictions

                'ExcelSheet.Protect("tosis123", , , , , , , , , , , , , True)

                xlApp.DisplayAlerts = False

                Dim temp_Filename As String = "Invoice " & Trim(pAffiliate) & "-" & Trim(Replace(pSuratJalanNo, "/", "-")) & ".xlsm"
                pFileName = pResult & "\" & temp_Filename
                ExcelBook.SaveAs(pFileName)
                ExcelBook.Close()
                xlApp.Workbooks.Close()
                xlApp.Quit()

                My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

                log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] ok.")
            End If
        Catch ex As Exception
            up_CreateInvoiceAffiliate = False

            log.WriteToErrorLog(pScreenName, "Process Create Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Create Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Create Invoice Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, LogName)
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
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
            If Not dsAffp Is Nothing Then
                dsAffp.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pFileName2 As String, ByVal pSuratJalanNo As String, ByVal pAffiliate As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoAffiliate = True

            dsEmail = clsGeneral.getEmailAddress(GB, pAffiliate, "PASI", "", "PASIReceivingCC", "PASIReceivingTO", "PASIReceivingTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            If receiptEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] Notification to Affiliate STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If fromEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send GR Affiliate [" & pAffiliate & "-" & pSuratJalanNo & "] Notification to Affiliate STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Good Receiving Affiliate: " & pSuratJalanNo
            ls_Body = clsNotification.GetNotification("62", , , , pSuratJalanNo.Trim)
            'ls_Attachment = Trim(pPathFile) & "\" & pFileName

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg, pFileName, pFileName2) = False Then
                sendEmailtoAffiliate = False
                Exit Function
            End If

            sendEmailtoAffiliate = True

        Catch ex As Exception
            sendEmailtoAffiliate = False
            errMsg = "Process Send GR [" & pAffiliate & "-" & pSuratJalanNo & "] Notification to Affiliate STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Sub UpdateExcelReceivingAffiliate(ByVal pAffCode As String, ByVal pSuratJalanNo As String, ByRef errMsg As String)

        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ReceiveAffiliate_Master " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE SuratJalanNo='" & pSuratJalanNo & "'  " & vbCrLf & _
                      " AND AffiliateID='" & pAffCode & "' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send SuratJalanNo [" & pSuratJalanNo & "], Affiliate [" & pAffCode & "] to Supplier STOPPED, because " & ex.Message
        End Try

    End Sub

    Shared Function bindDataReceivingAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pSuratJalan As String, ByVal pPONO As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = "  select distinct " & vbCrLf & _
                  " 	RAM.SuratJalanNo " & vbCrLf & _
                  " 	, RAM.ReceiveDate " & vbCrLf & _
                  " 	, KM.KanbanDate PlanDeliveryDate " & vbCrLf & _
                  " 	, DPM.DeliveryDate " & vbCrLf & _
                  " 	, ISNULL(RAM.JenisArmada,'') JenisArmada " & vbCrLf & _
                  " 	, ISNULL(RAM.NoPol,'') NoPol " & vbCrLf & _
                  " 	, ISNULL(RAM.DriverName,'') DriverName " & vbCrLf & _
                  " 	, ISNULL(RAM.TotalBox,0) TotalBox " & vbCrLf & _
                  "     , KM.DeliveryLocationCode " & vbCrLf & _
                  " 	, MDP.DeliveryLocationName " & vbCrLf & _
                  " 	, MDP.Address DeliveryAddress"

        ls_SQL = ls_SQL + " 	, RAD.PONo " & vbCrLf & _
                          " 	, RAD.PartNo " & vbCrLf & _
                          " 	, MP.PartName " & vbCrLf & _
                          " 	, CASE WHEN RAD.POKanbanCls = '1' then 'YES' else 'NO' end POKanbanCls " & vbCrLf & _
                          " 	, RAD.KanbanNo " & vbCrLf & _
                          " 	, MU.Description UnitDesc " & vbCrLf & _
                          " 	, QtyBox = ISNULL(DPD.POQtyBox,MPM.QtyBox) " & vbCrLf & _
                          " 	, DPD.DOQty SupplierDeliveryQty " & vbCrLf & _
                          " 	, RAD.RecQty " & vbCrLf & _
                          " 	, RAD.DefectQty " & vbCrLf & _
                          " 	, ISNULL((DPD.DOQty - RAD.RecQty),0) RemainingQty "

        ls_SQL = ls_SQL + " 	, ceiling(RAD.RecQty / ISNULL(DPD.POQtyBox,MPM.QtyBox)) ReceivingBox 	 " & vbCrLf & _
                          " 	, MPR.Price 	 " & vbCrLf & _
                          " 	, (MPR.Price * RAD.RecQty) Amount	 " & vbCrLf & _
                          " 	, MCC.Description CurrDesc 	 " & vbCrLf & _
                          "  from ReceiveAffiliate_Master RAM   " & vbCrLf & _
                          "  inner join ReceiveAffiliate_Detail RAD on RAM.SuratJalanNo = RAD.SuratJalanNo and RAM.AffiliateID = RAD.AffiliateID " & vbCrLf

        If pPONO = "1" Then
            ls_SQL = ls_SQL + "  left join DOPASI_Master DPM ON RAM.AffiliateID = DPM.AffiliateID and RAM.SuratJalanNo = DPM.SuratJalanNo " & vbCrLf & _
                              "  left join DOPASI_Detail DPD ON DPM.AffiliateID = DPD.AffiliateID and DPM.SuratJalanNo = DPD.SuratJalanNo AND RAD.PartNo = DPD.PartNo " & vbCrLf
        Else
            ls_SQL = ls_SQL + "  left join DOSupplier_Master DPM ON RAM.AffiliateID = DPM.AffiliateID and RAM.SuratJalanNo = DPM.SuratJalanNo " & vbCrLf & _
                              "  left join DOSupplier_Detail DPD ON DPM.AffiliateID = DPD.AffiliateID and DPM.SuratJalanNo = DPD.SuratJalanNo AND RAD.PartNo = DPD.PartNo " & vbCrLf
        End If
                          
        ls_SQL = ls_SQL + "  left join Kanban_Master KM on KM.KanbanNo = RAD.KanbanNo and KM.AffiliateID = RAD.AffiliateID and KM.SupplierID = RAD.SupplierID    " & vbCrLf & _
                          "  left join MS_DeliveryPlace MDP on KM.AffiliateID = MDP.AffiliateID and KM.DeliveryLocationCode = MDP.DeliveryLocationCode " & vbCrLf & _
                          "  left join MS_Parts MP on RAD.PartNo = MP.PartNo  " & vbCrLf & _
                          "  left join MS_UnitCls MU on MP.UnitCls = MU.UnitCls  " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = RAD.PartNo and MPM.AffiliateID = RAD.AffiliateID and MPM.SupplierID = RAD.SupplierID  " & vbCrLf & _
                          "  left join MS_Price MPR on MPR.PartNo = RAD.PartNo and MPR.AffiliateID = RAD.AffiliateID and (RAM.ReceiveDate between MPR.StartDate and MPR.EndDate)  " & vbCrLf & _
                          "  left join MS_CurrCls MCC ON MCC.CurrCls = MPR.CurrCls " & vbCrLf

        ls_SQL = ls_SQL + " where RAM.SuratJalanNo = '" & pSuratJalan & "' and RAM.AffiliateID = '" & pAffCode & "' "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

End Class
