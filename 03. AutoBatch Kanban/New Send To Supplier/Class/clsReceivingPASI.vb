Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsReceivingPASI
    Shared Sub up_SendReceivingPASI(ByVal cfg As GlobalSetting.clsConfig,
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
        Dim pSupplier As String = ""
        Dim temp_Filename1 As String = "", temp_Filename2 As String = ""

        Try
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data GR PASI")

            ls_SQL = "SELECT TOP 1 * FROM ReceivePASI_Master WHERE ExcelCls = '1' ORDER BY ReceiveDate"
            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pSuratJalanNo = Trim(ds.Tables(0).Rows(xi)("SuratJalanNo"))
                    pSupplier = Trim(ds.Tables(0).Rows(xi)("SupplierID"))

                    '01. Create data GR
                    If up_CreateGRSupplier(GB, log, pSuratJalanNo, pAffiliate, pSupplier, pAtttacment, pResult, pScreenName, temp_Filename1, LogName, errMsg, ErrSummary) = False Then
                        GoTo keluar
                    End If

                    '02. Create data Invoice
                    If up_CreateInvoiceAffiliate(GB, log, pSuratJalanNo, pAffiliate, pSupplier, pAtttacment, pResult, pScreenName, temp_Filename2, LogName, errMsg, ErrSummary) = False Then
                        GoTo keluar
                    End If

                    '03. Send Email to Supplier
                    If sendEmailtoSupplier(GB, pResult, temp_Filename1, temp_Filename2, pSuratJalanNo, pSupplier, errMsg) = False Then
                        Exit Try
                    Else
                        log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to PASI. Supplier [" & pSupplier & "-" & pSuratJalanNo & "] ok.")
                    End If

                    Call UpdateExcelReceivingPASI(pAffiliate, pSuratJalanNo, pSupplier, errMsg)

                    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email. GR & Invoice Supplier [" & pSupplier & "-" & pSuratJalanNo & "] ok.")
                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send GR & Invoice Supplier [" & pSupplier & "-" & pSuratJalanNo & "] ok", LogName)
                    LogName.Refresh()
keluar:
                Next
            End If
        Catch ex As Exception
            errMsg = "GR & Invoice PASI [" & pSupplier & "-" & pSuratJalanNo & "] " & ex.Message
            ErrSummary = "GR & Invoice PASI [" & pSupplier & "-" & pSuratJalanNo & "] " & ex.Message
        Finally
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    Shared Function up_CreateGRSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, ByVal pSuratJalanNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pAtttacment As String, ByVal pResult As String, ByRef pScreenName As String, ByRef pFileName As String, ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        up_CreateGRSupplier = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""

        Dim dsDetail As New DataSet

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template GoodReceivingPASI.xlsx") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send GR PASI to Supplier STOPPED, File Excel isn't Found"
                errSummary = "Process Send GR PASI to Supplier STOPPED, File Excel isn't Found"
                up_CreateGRSupplier = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel GR PASI [" & pSupplier & "-" & pSuratJalanNo & "]")

            NewFileCopy = pAtttacment & "\Template GoodReceivingPASI.xlsx"
            NewFileCopyTO = pResult & "\Template GoodReceivingPASI " & Format(Now, "yyyyMMddHHmmss") & ".xlsx"

            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
            Else
                System.IO.File.Copy(NewFileCopy, pResult & "\GoodReceivingPASI.xlsx")
            End If

            Dim ls_file As String = NewFileCopyTO

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetail = bindDataReceiving(GB, pAffiliate, pSuratJalanNo, pSupplier)

            If dsDetail.Tables(0).Rows.Count > 0 Then
                log.WriteToProcessLog(Date.Now, pScreenName, "Input Header Excel Supplier [" & pSupplier & "-" & pSuratJalanNo & "]")
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

                log.WriteToProcessLog(Date.Now, pScreenName, "Input Detail Excel Supplier [" & pSupplier & "-" & pSuratJalanNo & "]")
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

                Dim ls_SJ As String = ""
                If pSuratJalanNo.Contains("\") Then
                    ls_SJ = Replace(pSuratJalanNo, "\", "_")
                Else
                    ls_SJ = Replace(pSuratJalanNo, "/", "_")
                End If
                
                Dim temp_Filename As String = "GoodReceivingPASI " & Trim(pSupplier) & "-" & ls_SJ & ".xlsx"
                pFileName = pResult & "\" & temp_Filename
                ExcelBook.SaveAs(pFileName)
                ExcelBook.Close()
                xlApp.Workbooks.Close()
                xlApp.Quit()
                My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

                log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel GR Supplier [" & pSupplier & "-" & pSuratJalanNo & "] ok.")
            End If
        Catch ex As Exception
            up_CreateGRSupplier = False

            log.WriteToErrorLog(pScreenName, "Process Create GR Supplier [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Create GR Supplier [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Create GR Supplier [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, LogName)
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

    Shared Function up_CreateInvoiceAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, ByVal pSuratJalanNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pAtttacment As String, ByVal pResult As String, ByRef pScreenName As String, ByRef pFileName As String, ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
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
        Dim dsSupp As New DataSet

        Dim i As Integer

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template Invoice Supplier.xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send Invoice PASI to Supplier STOPPED, File Excel isn't Found"
                errSummary = "Process Send Invoice PASI to Supplier STOPPED, File Excel isn't Found"
                up_CreateInvoiceAffiliate = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel Invoice PASI [" & pSupplier & "-" & pSuratJalanNo & "]")

            NewFileCopy = pAtttacment & "\Template Invoice Supplier.xlsm"
            NewFileCopyTO = pResult & "\Template Invoice Supplier" & Format(Now, "yyyyMMddHHmmss") & ".xlsm"

            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
            Else
                System.IO.File.Copy(NewFileCopy, pResult & "\Invoice.xlsx")
            End If

            Dim ls_file As String = NewFileCopyTO

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetail = bindDataReceiving(GB, pAffiliate, pSuratJalanNo, pSupplier)

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
                ExcelSheet.Range("H5").Value = pSupplier

                ExcelSheet.Range("W11").Value = Format(Now, "dd-MMM-yy")

                dsAffp = clsGeneral.Affiliate(GB, Trim(pAffiliate))
                ExcelSheet.Range("I13").Value = dsAffp.Tables(0).Rows(0)("AffiliateName")
                ExcelSheet.Range("I14").Value = dsAffp.Tables(0).Rows(0)("Address")
                ExcelSheet.Range("I14").WrapText = True

                dsSupp = clsGeneral.Supplier(GB, Trim(pSupplier))
                ExcelSheet.Range("AM13").Value = dsSupp.Tables(0).Rows(0)("SupplierName")
                ExcelSheet.Range("AM14").Value = dsSupp.Tables(0).Rows(0)("Address")
                ExcelSheet.Range("AM14").WrapText = True

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
                    totalAmount = totalAmount + IIf(IsDBNull(dsDetail.Tables(0).Rows(i)("Amount")) = True, 0, dsDetail.Tables(0).Rows(i)("Amount"))

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

                'ExcelSheet.Protect("tosis123", , , , , , , , , , , , , True)

                xlApp.DisplayAlerts = False
                Dim ls_SJ As String = ""
                If pSuratJalanNo.Contains("\") Then
                    ls_SJ = Replace(pSuratJalanNo, "\", "_")
                Else
                    ls_SJ = Replace(pSuratJalanNo, "/", "_")
                End If
                Dim temp_Filename As String = "Invoice " & Trim(pSupplier) & "-" & Trim(ls_SJ) & ".xlsm"
                pFileName = pResult & "\" & temp_Filename
                ExcelBook.SaveAs(pFileName)
                ExcelBook.Close()
                xlApp.Workbooks.Close()
                xlApp.Quit()
                My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

                log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel Invoice Supplier [" & pSupplier & "-" & pSuratJalanNo & "] ok.")
            End If
        Catch ex As Exception
            up_CreateInvoiceAffiliate = False

            log.WriteToErrorLog(pScreenName, "Process Create Invoice Supplier [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Create Invoice Supplier [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Create Invoice Supplier [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message, LogName)
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
            If Not dsSupp Is Nothing Then
                dsSupp.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pFileName2 As String, ByVal pSuratJalanNo As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoSupplier = True

            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", pSupplier, "GoodReceiveCC", "GoodReceiveTO", "GoodReceiveTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
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
                sendEmailtoSupplier = False
                errMsg = "Process Send GR PASI [" & pSupplier & "-" & pSuratJalanNo & "] Notification to Affiliate STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If fromEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send GR PASI [" & pSupplier & "-" & pSuratJalanNo & "] Notification to Affiliate STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Good Receiving PASI and Invoice No.: " & pSuratJalanNo
            ls_Body = clsNotification.GetNotification("52", , , , pSuratJalanNo.Trim)
            ls_Attachment = pFileName

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg, pFileName, pFileName2) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send GR [" & pSupplier & "-" & pSuratJalanNo & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Sub UpdateExcelReceivingPASI(ByVal pAffCode As String, ByVal pSuratJalanNo As String, ByVal pSuppCode As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.ReceivePASI_Master " & vbCrLf & _
                          " SET ExcelCls='2'" & vbCrLf & _
                          " WHERE SuratJalanNo='" & pSuratJalanNo & "'  " & vbCrLf & _
                          " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                          " AND SupplierID='" & pSuppCode & "' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send Good Receiving [" & pSuratJalanNo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub

    Shared Function bindDataReceiving(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pSuratJalan As String, ByVal pSupplier As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = "   select distinct  " & vbCrLf & _
                  "  	RAM.SuratJalanNo  " & vbCrLf & _
                  "  	, RAM.ReceiveDate  " & vbCrLf & _
                  "  	, KM.KanbanDate PlanDeliveryDate  " & vbCrLf & _
                  "  	, DPM.DeliveryDate  " & vbCrLf & _
                  "  	, ISNULL(RAM.JenisArmada,'') JenisArmada  " & vbCrLf & _
                  "  	, ISNULL(RAM.NoPol,'') NoPol  " & vbCrLf & _
                  "  	, ISNULL(RAM.DriverName,'') DriverName  " & vbCrLf & _
                  "  	, ISNULL(RAM.TotalBox,0) TotalBox  " & vbCrLf & _
                  "      , KM.DeliveryLocationCode  " & vbCrLf & _
                  "  	, MDP.DeliveryLocationName  "

        ls_SQL = ls_SQL + "  	, MDP.Address DeliveryAddress 	, RAD.PONo  " & vbCrLf & _
                          "  	, RAD.PartNo  " & vbCrLf & _
                          "  	, MP.PartName  " & vbCrLf & _
                          "  	, CASE WHEN RAD.POKanbanCls = '1' then 'YES' else 'NO' end POKanbanCls  " & vbCrLf & _
                          "  	, RAD.KanbanNo  " & vbCrLf & _
                          "  	, MU.Description UnitDesc  " & vbCrLf & _
                          "  	, QtyBox = ISNULL(DPD.POQtyBox,MPM.QtyBox)  " & vbCrLf & _
                          "  	, DPD.DOQty SupplierDeliveryQty  " & vbCrLf & _
                          "  	, RAD.GoodRecQty RecQty  " & vbCrLf & _
                          "  	, RAD.DefectRecQty DefectQty  " & vbCrLf & _
                          "  	, ISNULL((DPD.DOQty - RAD.GoodRecQty),0) RemainingQty  	, ceiling(RAD.GoodRecQty / ISNULL(DPD.POQtyBox,MPM.QtyBox)) ReceivingBox 	  "

        ls_SQL = ls_SQL + "  	, ISNULL(RAD.Price,ISNULL(MPR.Price, ISNULL(MPRC.Price, 0))) Price 	  " & vbCrLf & _
                          "  	, (ISNULL(RAD.Price,ISNULL(MPR.Price, ISNULL(MPRC.Price, 0))) * RAD.GoodRecQty) Amount	  " & vbCrLf & _
                          "  	, MCC.Description CurrDesc 	  " & vbCrLf & _
                          "   from ReceivePASI_Master RAM    " & vbCrLf & _
                          "   inner join ReceivePASI_Detail RAD on RAM.SuratJalanNo = RAD.SuratJalanNo and RAM.AffiliateID = RAD.AffiliateID  and RAM.SupplierID = RAD.SupplierID " & vbCrLf & _
                          "   inner join DOSupplier_Master DPM ON RAM.AffiliateID = DPM.AffiliateID and RAM.SuratJalanNo = DPM.SuratJalanNo  " & vbCrLf & _
                          "   inner join DOSupplier_Detail DPD ON DPM.AffiliateID = DPD.AffiliateID and DPM.SuratJalanNo = DPD.SuratJalanNo AND RAD.PartNo = DPD.PartNo and RAD.SupplierID = DPD.SupplierID And RAD.KanbanNo = DPD.KanbanNo" & vbCrLf & _
                          "   left join Kanban_Master KM on KM.KanbanNo = RAD.KanbanNo and KM.AffiliateID = RAD.AffiliateID and KM.SupplierID = RAD.SupplierID     " & vbCrLf & _
                          "   left join MS_DeliveryPlace MDP on KM.AffiliateID = MDP.AffiliateID and KM.DeliveryLocationCode = MDP.DeliveryLocationCode  " & vbCrLf & _
                          "   left join MS_Parts MP on RAD.PartNo = MP.PartNo   " & vbCrLf & _
                          "   left join MS_UnitCls MU on MP.UnitCls = MU.UnitCls   "

        ls_SQL = ls_SQL + "   LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = RAD.PartNo and MPM.AffiliateID = RAD.AffiliateID and MPM.SupplierID = RAD.SupplierID   " & vbCrLf & _
                          "   left join MS_Price MPR on MPR.PartNo = RAD.PartNo and MPR.AffiliateID = RAD.SupplierID and MPR.DeliveryLocationID = RAD.AffiliateID and (RAM.ReceiveDate between MPR.StartDate and MPR.EndDate)   " & vbCrLf & _
                          "   left join MS_Price MPRC on MPRC.PartNo = RAD.PartNo and MPRC.AffiliateID = RAD.SupplierID and MPRC.DeliveryLocationID = '0000' and (RAM.ReceiveDate between MPR.StartDate and MPR.EndDate)   " & vbCrLf & _
                          "   left join MS_CurrCls MCC ON MCC.CurrCls = MPR.CurrCls  " & vbCrLf & _
                          "  where RAM.SuratJalanNo = '" & pSuratJalan & "' and RAM.AffiliateID = '" & pAffCode & "' and RAM.SupplierID = '" & pSupplier & "'  " & vbCrLf & _
                          "  "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

End Class
