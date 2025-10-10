Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports DevExpress.XtraPrinting

Public Class clsPONonKanban
    Shared Sub up_SendPONonKanban(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pBarcode As Boolean,
                              ByVal pInterval As String,                              
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim ls_sql As String = ""
        Dim Barcode As Boolean = pBarcode

        Dim ds As New DataSet

        Dim pAffiliate As String = ""
        Dim pSupplier As String = ""
        Dim pKanbanNo As String = ""
        Dim pDeliveryLocation As String = ""
        Dim pKanbanDate As Date
        Dim pPONo As String = ""

        Dim xi As Integer

        Dim pFileName1 As String = ""
        Dim pFileName2 As String = ""

        Try
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data PO Non Kanban")

            ls_sql = "select distinct PONO, a.affiliateID, a.SupplierID, a.KanbanDate, a.DeliveryLocationCode, a.KanbanNo " & vbCrLf & _
                     " from kanban_master a inner join kanban_Detail b  " & vbCrLf & _
                     " On a.kanbanno = b.kanbanno " & vbCrLf & _
                     " and a.affiliateID = b.AffiliateID " & vbCrLf & _
                     " and a.supplierID = b.supplierID " & vbCrLf & _
                     " LEFT JOIN ms_Etd_pasi MEP ON MEP.affiliateID = a.AffiliateID " & vbCrLf & _
                     "    AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(a.Kanbandate, " & vbCrLf & _
                     "          '')), 112) = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MEP.ETAAffiliate," & vbCrLf & _
                     "          '')), 112) " & vbCrLf & _
                     " WHERE ExcelCls='1' and kanbanstatus = '1' and convert(date,DATEADD(day," & pInterval & ",getdate())) >= ETDPASI" & vbCrLf

            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pKanbanNo = Trim(ds.Tables(0).Rows(xi)("KanbanNo"))
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pSupplier = Trim(ds.Tables(0).Rows(xi)("SupplierID"))
                    pDeliveryLocation = ds.Tables(0).Rows(xi)("DeliveryLocationCode") & ""
                    pKanbanDate = ds.Tables(0).Rows(xi)("KanbanDate")
                    pPONo = Trim(ds.Tables(0).Rows(xi)("PONO"))

                    pFileName1 = ""
                    pFileName2 = ""

                    '88. Create Kanban Barcode
                    If Barcode = True Then
                        log.WriteToProcessLog(Date.Now, pScreenName, "Start Create Barcode File, KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")
                        If CreateKanbanToPDF(GB, cfg, pKanbanDate, pKanbanNo, pAffiliate, pDeliveryLocation, pSupplier, pFileName1, pResult, errMsg) = False Then
                            log.WriteToProcessLog(Date.Now, pScreenName, "End Create Barcode File, KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & errMsg)
                            If errMsg = "Microsoft.VisualBasic.ErrObject" Then
                                End
                            End If
                            GoTo keluar
                        End If
                        log.WriteToProcessLog(Date.Now, pScreenName, "End Create Barcode File, KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")
                    End If

                    '89. Create DN
                    If CreateDelivery(GB, log, pAffiliate, pSupplier, pKanbanNo, pKanbanDate, pDeliveryLocation, pAtttacment, pResult, pScreenName, pFileName2, LogName, errMsg, ErrSummary) = False Then
                        GoTo keluar
                    End If

                    If pFileName1 <> "" And pFileName2 <> "" Then
                        If sendEmailtoSupplier(GB, pResult, pKanbanNo, pPONo, pAffiliate, pSupplier, errMsg, pFileName2, pFileName1) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        End If
                    End If

                    'If sendEmailtoAffiliate(GB, pResult, pKanbanNo, pAffiliate, pSupplier, pKanbanDate, pDeliveryLocation, errMsg) = False Then
                    'Else
                    '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC Affiliate. KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                    'End If

                    'If sendEmailtoPASI(GB, pResult, pKanbanNo, pAffiliate, pSupplier, pKanbanDate, pDeliveryLocation, errMsg) = False Then
                    'Else
                    '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC Affiliate. KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                    'End If

                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send PO Non Kanban [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
                    LogName.Refresh()

                    Call UpdateExcelNonKanban(pKanbanNo, pAffiliate, pSupplier, errMsg)
keluar:
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "Kanban [" & pKanbanNo & "-" & pAffiliate & "-" & pSupplier & "] " & ex.Message
            ErrSummary = "Kanban [" & pKanbanNo & "-" & pAffiliate & "-" & pSupplier & "] " & ex.Message
        Finally
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
        End Try
    End Sub

    Shared Function CreateKanbanToPDF(ByVal GB As GlobalSetting.clsGlobal, ByVal cfg As GlobalSetting.clsConfig, ByVal pKanbanDate As Date, ByVal pKanbanNo As String, ByVal pAffiliate As String, ByVal pDeliveryLocation As String, ByVal pSupplier As String, ByRef pFileName As String, ByVal pPathFile As String, ByRef errMsg As String) As Boolean
        Dim CrReport As New KanbanCard2

        Try
            Cursor.Current = Cursors.WaitCursor
            Dim dsPrint As New DataSet
            dsPrint = PrintBarcode(GB, pKanbanDate, pKanbanNo, pAffiliate, pSupplier, pDeliveryLocation)

            If IsNothing(dsPrint) Then
                Exit Try
            End If

            If dsPrint.Tables(0).Rows.Count = 0 Then Exit Try

            CrReport.DataSource = dsPrint.Tables(0)

            Dim ls_FileName = pPathFile & "\Barcode (NON KANBAN)-" & Replace(pKanbanNo, "'", "") & " " & pAffiliate.Trim & "-" & pSupplier.Trim & ".pdf"
            pFileName = ls_FileName

            Dim CrExportOptions As PdfExportOptions = CrReport.ExportOptions.Pdf

            CrExportOptions.ConvertImagesToJpeg = False
            CrExportOptions.ImageQuality = PdfJpegImageQuality.Medium

            Try
                CrReport.ExportToPdf(ls_FileName, CrExportOptions)

                CrReport = Nothing
                CrExportOptions = Nothing
            Catch err As Exception
                MessageBox.Show(err.ToString())
                CrReport = Nothing
                CrExportOptions = Nothing
            End Try

            'PDF
            CreateKanbanToPDF = True
        Catch ex As Exception
            errMsg = Err.ToString()
            CreateKanbanToPDF = False
        Finally
            If Not CrReport Is Nothing Then
                clsGeneral.NAR(CrReport)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If
        End Try
    End Function

    Shared Function CreateDelivery(ByVal GB As GlobalSetting.clsGlobal, ByVal log As GlobalSetting.clsLog, _
                                   ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pKanbanNo As String, ByVal pKanbanDate As Date, _
                                   ByVal pDeliveryLocation As String, ByVal pAtttacment As String, ByVal pResult As String, _
                                   ByRef pScreenName As String, ByRef pFileName1 As String, _
                                   ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        CreateDelivery = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""
        Dim fromEmail As String = ""
        Dim receiptCCEmail As String

        Dim dsDetailDelivery As New DataSet
        Dim dsEmail As New DataSet
        Dim dsSupp As New DataSet
        Dim dsAff As New DataSet
        Dim dsETAETD As New DataSet

        Dim y As Integer, k As Integer, j As Integer
        Dim ETDSupplier As String = ""

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template Delivery (PO non Kanban).xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send Kanban to Supplier STOPPED, File Excel isn't Found"
                errSummary = "Process Send Kanban to Supplier STOPPED, File Excel isn't Found"
                CreateDelivery = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel Kanban [" & pAffiliate & "-" & pSupplier & "-" & pKanbanDate & "]")

            NewFileCopy = pAtttacment & "\Template Delivery (PO non Kanban).xlsm"
            NewFileCopyTO = pResult & "\Template Delivery (PO non Kanban) " & Format(Now, "yyyyMMddHHmmss") & ".xlsm"

            If System.IO.File.Exists(NewFileCopyTO) = False Then
                If System.IO.File.Exists(NewFileCopy) = True Then
                    System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
                Else
                    System.IO.File.Copy(NewFileCopy, pResult & "\Delivery.xlsx")
                End If
            End If

            Dim ls_file As String = NewFileCopyTO

            'If Not (xlApp.Workbooks(ls_file) Is Nothing) Then
            '    'ExcelBook.Close()
            '    'xlApp.Workbooks.Close()
            '    'xlApp.Quit()
            'End If

            ExcelBook = xlApp.Workbooks.Open(ls_file)
            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

            dsDetailDelivery = bindDataDetailNonKanban(GB, pKanbanDate, pAffiliate, pSupplier, pDeliveryLocation, pKanbanNo)
            dsETAETD = bindHeaderETAETDNonKanban(GB, pAffiliate, pSupplier, pKanbanDate)

            If dsETAETD.Tables(0).Rows.Count > 0 Then
                ETDSupplier = dsETAETD.Tables(0).Rows(0)("ETDSupplier")
            Else
                ETDSupplier = ""
            End If

            If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                log.WriteToProcessLog(Date.Now, pScreenName, "Input Header Excel Affiliate [" & pAffiliate & "-" & pSupplier & "-" & pKanbanDate & "]")

                dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "KanbanCC", "KanbanTO", "KanbanTO", errMsg)

                For y = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If dsEmail.Tables(0).Rows(y)("flag") = "PASI" Then
                        fromEmail = dsEmail.Tables(0).Rows(y)("EmailFrom")
                        receiptCCEmail = dsEmail.Tables(0).Rows(y)("EmailCC")
                    End If
                Next

                ExcelSheet.Range("H2").Value = fromEmail
                'ExcelSheet.Range("Y2").Value = receiptCCEmail
                ExcelSheet.Range("H3").Value = pAffiliate
                ExcelSheet.Range("H4").Value = pDeliveryLocation
                ExcelSheet.Range("H5").Value = pSupplier

                dsSupp = clsGeneral.Supplier(GB, Trim(pSupplier))
                ExcelSheet.Range("I11").Value = dsSupp.Tables(0).Rows(0)("SupplierName")
                ExcelSheet.Range("I12").Value = dsSupp.Tables(0).Rows(0)("Address")
                ExcelSheet.Range("I12:X14").WrapText = True

                dsAff = clsGeneral.Affiliate(GB, Trim(pAffiliate))
                ExcelSheet.Range("I16").Value = dsAff.Tables(0).Rows(0)("AffiliateName")
                ExcelSheet.Range("I17").Value = dsAff.Tables(0).Rows(0)("Address")
                ExcelSheet.Range("I17:X19").WrapText = True

                k = 0
                Dim newKanbanNo As String = ""

                For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                    k = k

                    ExcelSheet.Range("B" & k + 39 & ": C" & k + 39).Merge()
                    ExcelSheet.Range("D" & k + 39 & ": H" & k + 39).Merge()
                    ExcelSheet.Range("i" & k + 39 & ": K" & k + 39).Merge()
                    ExcelSheet.Range("L" & k + 39 & ": O" & k + 39).Merge()
                    ExcelSheet.Range("P" & k + 39 & ": T" & k + 39).Merge()
                    ExcelSheet.Range("U" & k + 39 & ": AC" & k + 39).Merge()
                    ExcelSheet.Range("AD" & k + 39 & ": AE" & k + 39).Merge()
                    ExcelSheet.Range("AF" & k + 39 & ": AG" & k + 39).Merge()
                    ExcelSheet.Range("AH" & k + 39 & ": AJ" & k + 39).Merge()
                    ExcelSheet.Range("AK" & k + 39 & ": AN" & k + 39).Merge()
                    ExcelSheet.Range("AO" & k + 39 & ": AR" & k + 39).Merge()
                    ExcelSheet.Range("AS" & k + 39 & ": AV" & k + 39).Merge()
                    ExcelSheet.Range("AW" & k + 39 & ": AZ" & k + 39).Merge()
                    ExcelSheet.Range("BA" & k + 39 & ": BE" & k + 39).Merge()

                    ExcelSheet.Range("B" & k + 39 & ": C" & k + 39).Value = k + 1
                    ExcelSheet.Range("D" & k + 39 & ": H" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colpono")
                    ExcelSheet.Range("i" & k + 39 & ": K" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colpokanban")
                    newKanbanNo = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                    ExcelSheet.Range("L" & k + 39 & ": O" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                    ExcelSheet.Range("P" & k + 39 & ": T" & k + 39).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("colpartno"))
                    ExcelSheet.Range("U" & k + 39 & ": AC" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("coldescription")
                    ExcelSheet.Range("AD" & k + 39 & ": AE" & k + 39).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("coluom"))
                    ExcelSheet.Range("AF" & k + 39 & ": AG" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colqty")
                    ExcelSheet.Range("AH" & k + 39 & ": AJ" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colpallet")
                    ExcelSheet.Range("AK" & k + 39 & ": AN" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colcycle1")
                    ExcelSheet.Range("AO" & k + 39 & ": AR" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colcycle1")
                    ExcelSheet.Range("AS" & k + 39 & ": AV" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colbox1")
                    ExcelSheet.Range("AW" & k + 39 & ": AZ" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colNewPallet")
                    ExcelSheet.Range("AW" & k + 39 & ": AZ" & k + 39).NumberFormat = "#,##0.00"
                    ExcelSheet.Range("BA" & k + 39 & ": BE" & k + 39).Value = ETDSupplier
                    k = k + 1
                    k = k
                Next
                ExcelSheet.Range("B40").Interior.Color = Color.White
                ExcelSheet.Range("B40").Font.Color = Color.Black
                ExcelSheet.Range("B" & k + 39).Value = "E"
                ExcelSheet.Range("B" & k + 39).Interior.Color = Color.Black
                ExcelSheet.Range("B" & k + 39).Font.Color = Color.White

                clsGeneral.DrawAllBorders(ExcelSheet.Range("B39" & ": BE" & k + 38))
                ExcelSheet.Range("AO39" & ": AR" & k + 38).Interior.Color = Color.Yellow

                'Save ke Local
                xlApp.DisplayAlerts = False
                Dim pFile As String = newKanbanNo.Trim
                pFileName1 = pResult & "\Delivery(PO NON KANBAN) " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pFile & ".xlsm"

                ExcelBook.SaveAs(pFileName1)
                ExcelBook.Close()
                xlApp.Workbooks.Close()
                xlApp.Quit()
                My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

            End If
        Catch ex As Exception
            CreateDelivery = False

            log.WriteToErrorLog(pScreenName, "Process Create Kanban [" & pAffiliate & "-" & pSupplier & "-" & pKanbanDate & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
            log.WriteToProcessLog(Date.Now, pScreenName, "Process Create Kanban [" & pAffiliate & "-" & pSupplier & "-" & pKanbanDate & "] STOPPED, because " & ex.Message)

            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Create Kanban [" & pAffiliate & "-" & pSupplier & "-" & pKanbanDate & "] STOPPED, because " & ex.Message, LogName)
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

            If Not dsDetailDelivery Is Nothing Then
                dsDetailDelivery.Dispose()
            End If

            If Not dsETAETD Is Nothing Then
                dsETAETD.Dispose()
            End If
        End Try
    End Function

    Shared Function PrintBarcode(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value1 As Date, ByVal ls_value2 As String, ByVal ls_value3 As String, ByVal ls_value4 As String, ByVal ls_value5 As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = "   Select * from( SELECT Distinct ETACust = RTRIM(CONVERT(CHAR(5), CONVERT(DATETIME, ISNULL(KM.Kanbandate, " & vbCrLf & _
                  "                                                               '')), 103)) , " & vbCrLf & _
                  "             ETACustYear = '/' " & vbCrLf & _
                  "             + CONVERT(CHAR(4), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 120) , " & vbCrLf & _
                  "             ETACustTime = CONVERT(CHAR(5), KM.KanbanTime) , " & vbCrLf & _
                  "             ETAPasi = CONVERT(CHAR(5), CONVERT(DATETIME, ISNULL(MEP.ETDPASI, " & vbCrLf & _
                  "                                                               '')), 103) , " & vbCrLf & _
                  "             ETAPasiYear = '/' " & vbCrLf & _
                  "             + CONVERT(CHAR(4), CONVERT(DATETIME, ISNULL(MEP.ETDPasi, '')), 120) , " & vbCrLf & _
                  "             ETAPasiTime = '12:00' , " & vbCrLf & _
                  "             KanbanNo = RTRIM(KM.KanbanNo) , "

        ls_SQL = ls_SQL + "             SeqStart = RTRIM(CONVERT(NUMERIC, ISNULL(seqnoStart, 0))) , " & vbCrLf & _
                          "             SeqEnd = RTRIM(CONVERT(NUMERIC, ISNULL(seqnoEnd, 0))) , " & vbCrLf & _
                          "             PartNo1 = LEFT(Rtrim(KD.PartNo),2) , " & vbCrLf & _
                          "             PartNo2 = SUBSTRING(Rtrim(KD.PartNo),3,9) , " & vbCrLf & _
                          "             PartNo3 = SUBSTRING(Rtrim(KD.PartNo),12,10) , " & vbCrLf & _
                          "             PartName = RTRIM(MP.PartName) , " & vbCrLf & _
                          "             PartCMCode = RTRIM(ISNULL(MP.PartCarMaker, '')) , " & vbCrLf & _
                          "             PartCMName = RTRIM(ISNULL(MP.PartCarName, '')) , " & vbCrLf & _
                          "             Qty = REPLACE(RTRIM(ISNULL(KD.POQtyBox,ML.QtyBox)), '.00', '') , " & vbCrLf & _
                          "             BoxNo = RTRIM(ISNULL(KB.BoxNo, '')) , " & vbCrLf & _
                          "             Cust = RTRIM(KM.AffiliateID), " & vbCrLf & _
                          "             AFFCode = RTRIM(ISNULL(MA.AffiliateCode, '')) , " & vbCrLf & _
                          "             Location = RTRIM(ISNULL(ML.LocationID, '')) , "

        ls_SQL = ls_SQL + "             SupplierID = RTRIM(KM.SupplierID) + '#1' , " & vbCrLf & _
                          "             SupplierCode = RTRIM(ISNULL(MS.SupplierCode, '')) , " & vbCrLf & _
                          "             Barcode = RTRIM(KB.Barcode2) " & vbCrLf & _
                          "   FROM      dbo.Kanban_Master KM " & vbCrLf & _
                          "             LEFT JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID " & vbCrLf & _
                          "                                               AND kanbanqty <> 0 " & vbCrLf & _
                          "                                               AND KM.KanbanNo = KD.KanbanNo " & vbCrLf & _
                          "                                               AND KM.SupplierID = KD.SupplierID " & vbCrLf & _
                          "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = KM.SupplierID " & vbCrLf & _
                          "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo " & vbCrLf & _
                          "             INNER JOIN Kanban_Barcode KB ON KB.PONO = KD.PONO "

        ls_SQL = ls_SQL + "                                             AND KB.KanbanNo = KD.Kanbanno " & vbCrLf & _
                          "                                             AND KB.AffiliateID = KD.AffiliateID " & vbCrLf & _
                          "                                             AND KB.DeliveryLocationCode = KD.DeliveryLocationCode " & vbCrLf & _
                          "                                             AND KB.SupplierID = KD.SupplierID " & vbCrLf & _
                          "                                             AND KB.PartNo = KD.PartNo " & vbCrLf & _
                          "             LEFT JOIN MS_PartMapping ML ON KD.AffiliateID = ML.affiliateID AND ML.SupplierID = KD.SupplierID AND KD.PartNo = ML.PartNo " & vbCrLf & _
                          "             LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          "             LEFT JOIN ms_Etd_pasi MEP ON MEP.affiliateID = KM.AffiliateID " & vbCrLf & _
                          "                                          AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate, " & vbCrLf & _
                          "                                                               '')), 112) = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MEP.ETAAffiliate, "

        ls_SQL = ls_SQL + "                                                               '')), 112) " & vbCrLf & _
                          "   WHERE     KD.AffiliateID = '" & Trim(ls_value3) & "' " & vbCrLf & _
                          "             AND KD.SupplierID = '" & Trim(ls_value4) & "' " & vbCrLf & _
                          "             AND KD.kanbanno IN ('" & Trim(ls_value2) & "') " & vbCrLf & _
                          "   )xx   ORDER BY  kanbanno, Rtrim(PartNo1)+Rtrim(partno2), seqstart   "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Sub UpdateExcelNonKanban(ByVal pKanbanNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.Kanban_Master " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE KanbanStatus = '1'  " & vbCrLf & _
                      " AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                      " AND SupplierID = '" & pSupplier & "' " & vbCrLf & _
                      " AND KanbanNo = '" & pKanbanNo & "'"

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send PO Kanban: Affiliate [" & pAffiliate & "], Supplier [" & pSupplier & "], KanbanNo [" & pKanbanNo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub

    Shared Function bindDataDetailNonKanban(ByVal GB As GlobalSetting.clsGlobal, ByVal pDate As Date, ByVal pAffCode As String, ByVal pSupplierCode As String, ByVal pDeliveryLocation As String, ByVal pPONO As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = "  SELECT  colkanbanno, " & vbCrLf & _
                  "   	colpokanban = colpokanban, colno = ROW_NUMBER() OVER(ORDER BY cols DESC),   " & vbCrLf & _
                  "   	colpartno = colpartno,   " & vbCrLf & _
                  "   	coldescription=coldescription,   " & vbCrLf & _
                  "   	colpono=colpono,   " & vbCrLf & _
                  "   	coluom=coluom,   " & vbCrLf & _
                  "   	colqty=colqty,   " & vbCrLf & _
                  "   	colkanbanqty=colkanbanqty,   " & vbCrLf & _
                  "   	colcycle1 = colcycle1,   " & vbCrLf 

        ls_SQL = ls_SQL + "   	colbox1 = CEILING(colbox1), " & vbCrLf & _
                          "     colpallet1 = case when isnull(boxpallet,0) = 0 then 0 else CEILING(colbox1/boxpallet) END, colpallet = boxpallet " & vbCrLf & _
                          "     ,colNewPallet = round(colbox1 / boxpallet,2) " & vbCrLf & _
                          "   FROM (  " & vbCrLf & _
                          "  SELECT DISTINCT  colkanbanno = KM.Kanbanno, colpokanban = (case when isnull(KM.KanbanStatus,'') = '0' then 'YES' else 'NO' END), " & vbCrLf & _
                          "      cols = '1',  " & vbCrLf & _
                          "  	colno = '0',  " & vbCrLf & _
                          "  	colpartno = KD.partNo,  " & vbCrLf & _
                          "  	coldescription = MP.partname ,  " & vbCrLf & _
                          "  	colpono = KD.pono,   " & vbCrLf & _
                          "  	coluom = ISNULL(MUC.Description,''),  " & vbCrLf & _
                          "  	colmoq = ISNULL(KD.POMOQ,MPM.MOQ),  "

        ls_SQL = ls_SQL + "  	colqty = ISNULL(KD.POQtyBox,MPM.QtyBox),  " & vbCrLf & _
                          "  	colpoqty = COALESCE(PRD.POQty,PD.POQty),   " & vbCrLf & _
                          "  	colkanbanqty= ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI With(Nolock) "

        ls_SQL = ls_SQL + "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = 1  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  				+ ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI With(Nolock) " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  "

        ls_SQL = ls_SQL + "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = 2  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  				+ ISNULL((SELECT KanbanQty  				FROM dbo.Kanban_Master  KMI With(Nolock) " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  "

        ls_SQL = ls_SQL + "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = 3  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  				+ ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI With(Nolock) " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  "

        ls_SQL = ls_SQL + "  				WHERE KMI.KanbanCycle = 4  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  " & vbCrLf & _
                          "  	colcycle1 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI With(Nolock) " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  				WHERE KMI.KanbanCycle = 1  "

        ls_SQL = ls_SQL + "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  " & vbCrLf
                         
        ls_SQL = ls_SQL + "  	colbox1 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI With(Nolock) " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = 1  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  					 " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0) / CASE WHEN ISNULL(KD.POQtyBox,0) = 0 then KD.POMOQ else KD.POQtyBox end, " & vbCrLf

        ls_SQL = ls_SQL + "   boxpallet = MPM.BoxPallet, " & vbCrLf & _
                          "  	cols1 = '1', coluomcode = MP.UnitCls  " & vbCrLf & _
                          "  FROM dbo.Kanban_Master KM  With(Nolock) " & vbCrLf & _
                          "  	LEFT JOIN dbo.Kanban_Detail  KD ON KM.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                          "                                          AND KM.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                                          AND KM.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                                          AND KM.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf & _
                          "  	LEFT JOIN dbo.po_detailUpload PD ON KD.PartNo = PD.PartNo AND KD.PONo = PD.PONo  " & vbCrLf & _
                          "      LEFT JOIN PO_Master PM ON PM.PoNo = PD.PONo  " & vbCrLf & _
                          "                                  and PM.AffiliateID = PD.AffiliateID  " & vbCrLf & _
                          "                                  and PM.SupplierID = PD.SupplierID  "

        ls_SQL = ls_SQL + "      LEFT JOIN dbo.PORev_Master PRM ON PM.AffiliateID = PRM.AffiliateID  " & vbCrLf & _
                          "                                      AND PRM.PONo = PM.PONo  " & vbCrLf & _
                          "                                      AND PRM.SupplierID = PM.SupplierID  " & vbCrLf & _
                          "      LEFT JOIN dbo.PORev_Detail PRD ON PRD.PONo = PRM.PONo  " & vbCrLf & _
                          "                                      AND PRD.AffiliateID = PRM.AffiliateID  " & vbCrLf & _
                          "                                      AND PRD.SupplierID = PRM.SupplierID  " & vbCrLf & _
                          "                                      AND PRD.PartNo = PD.PartNo  " & vbCrLf & _
                          "                                      AND PRD.SeqNo = (SELECT MAX(seqNO) FROM PORev_Detail A With(Nolock) WHERE " & vbCrLf & _
                          "                                                          A.PONo = PD.PONo  " & vbCrLf & _
                          " 							                                AND A.AffiliateID = PD.AffiliateID   " & vbCrLf & _
                          " 							                                AND A.SupplierID = PD.SupplierID   "

        ls_SQL = ls_SQL + " 							                                AND A.PartNo = PD.PartNo)  " & vbCrLf & _
                          "  	 LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo  " & vbCrLf & _
                          "      LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo and MPM.AffiliateID = KD.AffiliateID and MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                          "      LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls  " & vbCrLf & _
                          "      LEFT JOIN dbo.MS_SupplierCapacity MSS ON MSS.SupplierID = PD.SupplierID AND MSS.PartNo = PD.PartNo	 " & vbCrLf & _
                          "  WHERE CONVERT(char(10), CONVERT(DATETIME,KM.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "   AND KM.AffiliateID = '" & pAffCode & "' And KM.DeliveryLocationcode = '" & pDeliveryLocation & "' AND KM.SupplierID = '" & Trim(pSupplierCode) & "' and KM.ExcelCls=1 and KD.KanbanNo = '" & Trim(pPONO) & "')xx " & _
                          "    where colkanbanqty <> 0 order by colpono asc"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return (ds)
    End Function

    Shared Function bindHeaderETAETDNonKanban(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pSupplierCode As String, ByVal pDate As Date) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT distinct AffiliateID, SupplierID, ETAPASI = CONVERT(CHAR(11),isnull(ETAPASI,''),106), ETDPASI = CONVERT(CHAR(11),isnull(ETDPASI,''),106), ETDSupplier = CONVERT(CHAR(11),isnull(ETDSUPPLIER,''),106)  " & vbCrLf & _
                 " FROM MS_ETD_PASI EP LEFT JOIN MS_ETD_Supplier_Pasi ES " & vbCrLf & _
                 " ON EP.ETDPASI = ES.ETAPASI WHERE AffiliateID = '" & Trim(pAffCode) & "' AND SupplierID = '" & Trim(pSupplierCode) & "' AND ETAAFFILIATE = '" & Format(pDate, "yyyy-MM-dd") & "'" & vbCrLf
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return (ds)
    End Function

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                        ByVal pKanbanNo As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String, _
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

            sendEmailtoSupplier = True

            Dim dsEmail As New DataSet
            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", pSupplier, "KanbanCC", "KanbanTO", "KanbanTO", errMsg)

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
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "SUPP" Then
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
                sendEmailtoSupplier = False
                errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Send To Supplier PO Non Kanban : " & pAffiliate.Trim & "-" & pPONo & "-" & pKanbanNo & "-" & pSupplier.Trim

            ls_Body = clsNotification.GetNotification("30", "", "", pKanbanNo)
            ls_Body = Replace(ls_Body, "Kanban", "Non Kaban")
            ls_Attachment = Trim(pPathFile) & "\" & pDN1

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, pSupplier, ls_Subject, ls_Body, errMsg, pDN1, IIf(pBarcodeFile = "", "", pBarcodeFile)) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Function sendEmailtoAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                              ByVal pKanbanNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, _
                                              ByVal pKanbanDate As Date, ByVal pDeliveryLocation As String, ByRef errMsg As String) As Boolean
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""


            sendEmailtoAffiliate = True

            Dim dsEmail As New DataSet
            dsEmail = clsGeneral.getEmailAddress(GB, pAffiliate, "PASI", "", "KanbanCC", "KanbanTO", "KanbanTO", errMsg)

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
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
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

            If fromEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If


            ls_URl = "http://" & clsNotification.pub_ServerName & "/Kanban/KanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pKanbanDate.Date) & "&t1=" & clsNotification.EncryptURL(pSupplier) & _
                                       "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/Kanban/KanbanList.aspx")

            ls_Subject = "Send To Supplier PO Non Kanban : " & pKanbanNo & "-" & pAffiliate.Trim & "-" & pSupplier.Trim

            ls_Body = clsNotification.GetNotification("30", ls_URl, "", pKanbanNo)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoAffiliate = False
                Exit Function
            End If

            sendEmailtoAffiliate = True

        Catch ex As Exception
            sendEmailtoAffiliate = False
            errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Function sendEmailtoPASI(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                        ByVal pKanbanNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, _
                                        ByVal pKanbanDate As Date, ByVal pDeliveryLocation As String, ByRef errMsg As String) As Boolean
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoPASI = True

            Dim dsEmail As New DataSet
            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "KanbanCC", "KanbanTO", "KanbanTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    Else
                        receiptEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("EmailTO")
                    End If
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    Else
                        receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("EmailCC")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            If fromEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to PASI [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to PASI [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerNamePASI & "/AffKanban/AffKanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pKanbanDate.Date) & "&t1=" & clsNotification.EncryptURL(pSupplier) & _
                                       "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/AffKanban/AffKanbanList.aspx")

            ls_Subject = "Send To Supplier PO Non Kanban : " & pKanbanNo & "-" & pAffiliate.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("30", ls_URl, "", pKanbanNo)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoPASI = False
                Exit Function
            End If

            sendEmailtoPASI = True

        Catch ex As Exception
            sendEmailtoPASI = False
            errMsg = "Process Send PO Non Kanban [" & pKanbanNo & "] from Affiliate [" & pAffiliate & "] to PASI [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

End Class
