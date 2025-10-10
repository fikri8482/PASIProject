Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports DevExpress.XtraPrinting

Public Class clsPOKanban
    Shared Sub up_SendPOKanban(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pBarcode As Boolean,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim ls_sql As String = ""
        Dim Barcode As Boolean = pBarcode

        Dim ds As New DataSet
        Dim dsHeader As New DataSet

        Dim pAffiliate As String = ""
        Dim pSupplier As String = ""
        Dim pKanbanNo As String = ""
        Dim xi As Integer, xy As Integer

        Dim pKanbanDate As Date
        Dim pDeliveryLocation As String = ""

        Dim pFileName1 As String = ""
        Dim pFileName2 As String = ""
        Dim pFileName3 As String = ""
        Dim pFileName4 As String = ""
        Dim pFileName5 As String = ""

        Dim pCycle As String = ""
        Dim pKanbanNo1 As String = ""
        Dim pKanbanNo2 As String = ""
        Dim pKanbanNo3 As String = ""
        Dim pKanbanNo4 As String = ""

        Try
            '01. Delete Kanban PASI
            log.WriteToProcessLog(Date.Now, pScreenName, "Start Delete data PO Kanban Supplier = PASI")
            Call DeleteKanbanSuppPASI()
            log.WriteToProcessLog(Date.Now, pScreenName, "End Delete data PO Kanban Supplier = PASI")

            '02. Get Data Kanban
            log.WriteToProcessLog(Date.Now, pScreenName, "Get data PO Kanban")

            ls_sql = " SELECT DISTINCT TOP 1 ISNULL(KanbanSeq_No,1) As Cycle,kanbandate2 = CONVERT(char(10), CONVERT(DATETIME,KanbanDate),120), KM.SupplierID SupplierID, KM.AffiliateID AffiliateID, KM.DeliveryLocationCode DeliveryLocationCode, DeliveryLocationName, SupplierName, KM.EntryDate FROM Kanban_Master KM " & vbCrLf & _
                     " LEFT JOIN MS_Supplier MS ON MS.SupplierID = KM.SupplierID " & vbCrLf & _
                     " LEFT JOIN MS_DeliveryPlace MD ON MD.DeliveryLOcationCode = KM.DeliveryLocationCode WHERE ExcelCls='1' and isnull(kanbanstatus,0) = 0 " & vbCrLf & _
                     " ORDER BY KM.EntryDate "
            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                    pSupplier = Trim(ds.Tables(0).Rows(xi)("SupplierID"))
                    pDeliveryLocation = ds.Tables(0).Rows(xi)("DeliveryLocationCode")
                    pKanbanDate = ds.Tables(0).Rows(xi)("KanbanDate2")
                    pCycle = ds.Tables(0).Rows(xi)("Cycle")

                    log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel PO Kanban, Affiliate [" & pAffiliate & "], Supplier [" & pSupplier & "], KanbanDate [" & pKanbanDate & "]")

                    dsHeader = bindDataHeaderKanban(GB, pKanbanDate, pAffiliate, pSupplier, pDeliveryLocation, pCycle)

                    If dsHeader.Tables(0).Rows.Count > 0 Then
                        pKanbanNo = ""
                        For xy = 0 To dsHeader.Tables(0).Rows.Count - 1
                            If pCycle = "1" Then
                                If Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "1" Then
                                    pKanbanNo1 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "2" Then
                                    pKanbanNo2 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "3" Then
                                    pKanbanNo3 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "4" Then
                                    pKanbanNo4 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                End If
                                If Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) <> "" Then
                                    If pKanbanNo = "" Then
                                        pKanbanNo = "'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    Else
                                        pKanbanNo = pKanbanNo & ",'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    End If
                                End If
                            ElseIf pCycle = "2" Then
                                If Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "5" Then
                                    pKanbanNo1 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "6" Then
                                    pKanbanNo2 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "7" Then
                                    pKanbanNo3 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "8" Then
                                    pKanbanNo4 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                End If
                                If Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) <> "" Then
                                    If pKanbanNo = "" Then
                                        pKanbanNo = "'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    Else
                                        pKanbanNo = pKanbanNo & ",'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    End If
                                End If
                            ElseIf pCycle = "3" Then
                                If Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "9" Then
                                    pKanbanNo1 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "10" Then
                                    pKanbanNo2 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "11" Then
                                    pKanbanNo3 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "12" Then
                                    pKanbanNo4 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                End If
                                If Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) <> "" Then
                                    If pKanbanNo = "" Then
                                        pKanbanNo = "'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    Else
                                        pKanbanNo = pKanbanNo & ",'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    End If
                                End If
                            ElseIf pCycle = "4" Then
                                If Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "13" Then
                                    pKanbanNo1 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "14" Then
                                    pKanbanNo2 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "15" Then
                                    pKanbanNo3 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "16" Then
                                    pKanbanNo4 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                End If
                                If Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) <> "" Then
                                    If pKanbanNo = "" Then
                                        pKanbanNo = "'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    Else
                                        pKanbanNo = pKanbanNo & ",'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    End If
                                End If
                            ElseIf pCycle = "5" Then
                                If Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "17" Then
                                    pKanbanNo1 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "18" Then
                                    pKanbanNo2 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "19" Then
                                    pKanbanNo3 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                ElseIf Trim(dsHeader.Tables(0).Rows(xy)("kanbancycle")) = "20" Then
                                    pKanbanNo4 = Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo"))
                                End If
                                If Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) <> "" Then
                                    If pKanbanNo = "" Then
                                        pKanbanNo = "'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    Else
                                        pKanbanNo = pKanbanNo & ",'" & Trim(dsHeader.Tables(0).Rows(xy)("KanbanNo")) & "'"
                                    End If
                                End If
                            End If
                        Next

                        '88. Create Kanban Barcode
                        If Barcode = True Then
                            log.WriteToProcessLog(Date.Now, pScreenName, "Start Create Barcode File, KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")
                            If CreateKanbanToPDF(GB, cfg, pKanbanDate, pKanbanNo, pAffiliate, pDeliveryLocation, pSupplier, pFileName1, pResult, errMsg, log) = False Then
                                log.WriteToProcessLog(Date.Now, pScreenName, "End Create Barcode File, KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & errMsg)
                                If errMsg = "Microsoft.VisualBasic.ErrObject" Then
                                    End
                                End If
                                GoTo keluar
                            End If
                            log.WriteToProcessLog(Date.Now, pScreenName, "End Create Barcode File, KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")
                        End If

                        '89. Create DN
                        If CreateDelivery(GB, log, pAffiliate, pSupplier, pKanbanDate, pCycle, pDeliveryLocation, pAtttacment, pResult, pScreenName, pFileName2, pFileName3, pFileName4, pFileName5, LogName, errMsg, ErrSummary) = False Then
                            GoTo keluar
                        End If

                        If pKanbanNo1 <> "" Then
                            If pFileName2 = "" Then
                                Exit Try
                            End If
                        End If

                        If pKanbanNo2 <> "" Then
                            If pFileName3 = "" Then
                                Exit Try
                            End If
                        End If

                        If pKanbanNo3 <> "" Then
                            If pFileName4 = "" Then
                                Exit Try
                            End If
                        End If

                        If pKanbanNo4 <> "" Then
                            If pFileName5 = "" Then
                                Exit Try
                            End If
                        End If

                        If pFileName1 <> "" Then
                            If sendEmailtoSupplier(GB, pResult, pAffiliate, pSupplier, errMsg, "", IIf(pKanbanNo1 = "", "", pKanbanNo1), IIf(pKanbanNo2 = "", "", pKanbanNo2), IIf(pKanbanNo3 = "", "", pKanbanNo3), IIf(pKanbanNo4 = "", "", pKanbanNo4), IIf(pFileName1 = "", "", pFileName1)) = False Then
                                Exit Try
                            Else
                                log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            End If
                        End If

                        'If sendEmailtoAffiliate(GB, pResult, pAffiliate, pSupplier, pKanbanDate, pDeliveryLocation, errMsg, IIf(pKanbanNo1 = "", "", pKanbanNo1), IIf(pKanbanNo2 = "", "", pKanbanNo2), IIf(pKanbanNo3 = "", "", pKanbanNo3), IIf(pKanbanNo4 = "", "", pKanbanNo4), IIf(pFileName1 = "", "", pFileName1)) = False Then
                        'Else
                        '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC Affiliate. KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        'End If

                        'If sendEmailtoPASI(GB, pResult, pAffiliate, pSupplier, pKanbanDate, pDeliveryLocation, errMsg, IIf(pKanbanNo1 = "", "", pKanbanNo1), IIf(pKanbanNo2 = "", "", pKanbanNo2), IIf(pKanbanNo3 = "", "", pKanbanNo3), IIf(pKanbanNo4 = "", "", pKanbanNo4), IIf(pFileName1 = "", "", pFileName1)) = False Then
                        'Else
                        '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC Affiliate. KanbanNo [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        'End If

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send PO Kanban [" & pKanbanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
                        LogName.Refresh()

                        log.WriteToProcessLog(Date.Now, pScreenName, "Start Update Cls PO Kanban [" & pKanbanNo & "]")
                        Call UpdateExcelKanban(pKanbanDate, pAffiliate, pSupplier, pDeliveryLocation, pKanbanNo, errMsg)

                        If errMsg.Substring(0, 6) <> "UPDATE" Then
                            log.WriteToProcessLog(Date.Now, pScreenName, "Error Update Cls PO Kanban [" & pKanbanNo & "] Message [" & errMsg & "]")
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Success Update Cls PO Kanban [" & pKanbanNo & "].ok Query [" & errMsg & "]")
                            errMsg = ""
                        End If

                        log.WriteToProcessLog(Date.Now, pScreenName, "End Update Cls PO Kanban [" & pKanbanNo & "]")
                    End If
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

    Shared Function CreateKanbanToPDF(ByVal GB As GlobalSetting.clsGlobal, ByVal cfg As GlobalSetting.clsConfig, ByVal pKanbanDate As Date, ByVal pKanbanNo As String, ByVal pAffiliate As String, ByVal pDeliveryLocation As String, ByVal pSupplier As String, ByRef pFileName As String, ByVal pPathFile As String, ByRef errMsg As String, ByVal log As GlobalSetting.clsLog) As Boolean
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

            Dim ls_FileName = pPathFile & "\Barcode-" & Replace(pKanbanNo, "'", "") & " " & pAffiliate.Trim & "-" & pSupplier.Trim & ".pdf"
            pFileName = ls_FileName

            Dim CrExportOptions As PdfExportOptions = CrReport.ExportOptions.Pdf

            CrExportOptions.ConvertImagesToJpeg = False
            CrExportOptions.ImageQuality = PdfJpegImageQuality.Medium

            Try
                CrReport.ExportToPdf(ls_FileName, CrExportOptions)

                CrReport = Nothing
                CrExportOptions = Nothing
            Catch err As Exception
                'MessageBox.Show(err.ToString())
                log.WriteToProcessLog(Date.Now, "SendPOKanban", "Error Create PDF " & err.ToString())
                CrReport = Nothing
                CrExportOptions = Nothing
            End Try

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
                                   ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pKanbanDate As Date, _
                                   ByVal pCycle As String, ByVal pDeliveryLocation As String, ByVal pAtttacment As String, ByVal pResult As String, _
                                   ByRef pScreenName As String, ByRef pFileName1 As String, ByRef pFileName2 As String, ByRef pFileName3 As String, _
                                   ByRef pFileName4 As String, ByVal LogName As RichTextBox, ByRef errMsg As String, ByRef errSummary As String) As Boolean
        CreateDelivery = True

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""
        Dim fromEmail As String = ""
        Dim receiptCCEmail As String = ""

        Dim dsDetailDelivery As New DataSet
        Dim dsEmail As New DataSet
        Dim dsSupp As New DataSet
        Dim dsAff As New DataSet
        Dim dsETAETD As New DataSet

        Dim i As Integer, y As Integer, k As Integer, j As Integer
        Dim jQty As Long
        Dim jQtyBox As Long
        Dim jQtyPallet As Double
        Dim ETDSupplier As String = ""

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template Delivery.xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send Kanban to Supplier STOPPED, File Excel isn't Found"
                errSummary = "Process Send Kanban to Supplier STOPPED, File Excel isn't Found"
                CreateDelivery = False
                Exit Function
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel Kanban [" & pAffiliate & "-" & pSupplier & "-" & pKanbanDate & "]")

            NewFileCopy = pAtttacment & "\Template Delivery.xlsm"

            For i = 0 To 3
                NewFileCopyTO = pResult & "\Template Delivery " & Format(Now, "yyyyMMddHHmmss") & ".xlsm"

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

                dsDetailDelivery = bindDataDetailKanbanDelivery(GB, pKanbanDate, pAffiliate, pSupplier, pDeliveryLocation, i + 1, pCycle)
                dsETAETD = bindHeaderETAETDKanban(GB, pAffiliate, pSupplier, pKanbanDate)

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
                            receiptCCEmail = dsEmail.Tables(0).Rows(y)("EmailCC") 'Ga dipake ini -__- 20220405
                        End If
                    Next

                    dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", pSupplier, "SupplierDeliveryCC", "KanbanTO", "KanbanTO", errMsg)
                    For y = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If dsEmail.Tables(0).Rows(y)("flag") = "SUPP" Then
                            receiptCCEmail = dsEmail.Tables(0).Rows(y)("EmailCC") 'yg dipake ini 20220405
                        End If
                    Next

                    ExcelSheet.Range("H2").Value = fromEmail
                    ExcelSheet.Range("H3").Value = pAffiliate
                    ExcelSheet.Range("H4").Value = pDeliveryLocation
                    ExcelSheet.Range("H5").Value = pSupplier
                    ExcelSheet.Range("Y2").Value = receiptCCEmail.Trim '20220405
                    ExcelSheet.Range("Y2:AT5").WrapText = True

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
                        'For i = 0 To 3
                        k = k

                        If i = 0 Then
                            newKanbanNo = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                            jQty = dsDetailDelivery.Tables(0).Rows(j)("colcycle1")
                            jQtyBox = dsDetailDelivery.Tables(0).Rows(j)("colbox1")
                            jQtyPallet = dsDetailDelivery.Tables(0).Rows(j)("colNewPallet")
                        End If

                        If i = 1 Then
                            newKanbanNo = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                            jQty = dsDetailDelivery.Tables(0).Rows(j)("colcycle2")
                            jQtyBox = dsDetailDelivery.Tables(0).Rows(j)("colbox2")
                            jQtyPallet = dsDetailDelivery.Tables(0).Rows(j)("colpallet2")
                        End If

                        If i = 2 Then
                            newKanbanNo = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                            jQty = dsDetailDelivery.Tables(0).Rows(j)("colcycle3")
                            jQtyBox = dsDetailDelivery.Tables(0).Rows(j)("colbox3")
                            jQtyPallet = dsDetailDelivery.Tables(0).Rows(j)("colpallet3")
                        End If

                        If i = 3 Then
                            newKanbanNo = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                            jQty = dsDetailDelivery.Tables(0).Rows(j)("colcycle4")
                            jQtyBox = dsDetailDelivery.Tables(0).Rows(j)("colbox4")
                            jQtyPallet = dsDetailDelivery.Tables(0).Rows(j)("colpallet4")
                        End If

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
                        ExcelSheet.Range("L" & k + 39 & ": O" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colkanbanno")
                        ExcelSheet.Range("P" & k + 39 & ": T" & k + 39).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("colpartno"))
                        ExcelSheet.Range("U" & k + 39 & ": AC" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("coldescription")
                        ExcelSheet.Range("AD" & k + 39 & ": AE" & k + 39).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("coluom"))
                        ExcelSheet.Range("AF" & k + 39 & ": AG" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colqty")
                        ExcelSheet.Range("AH" & k + 39 & ": AJ" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("colpallet")
                        ExcelSheet.Range("AK" & k + 39 & ": AN" & k + 39).Value = jQty
                        ExcelSheet.Range("AO" & k + 39 & ": AR" & k + 39).Value = jQty
                        ExcelSheet.Range("AS" & k + 39 & ": AV" & k + 39).Value = jQtyBox
                        ExcelSheet.Range("AW" & k + 39 & ": AZ" & k + 39).Value = jQtyPallet
                        ExcelSheet.Range("AW" & k + 39 & ": AZ" & k + 39).NumberFormat = "#,##0.00"
                        ExcelSheet.Range("BA" & k + 39 & ": BE" & k + 39).Value = ETDSupplier
                        k = k + 1
                        'Next
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
                    Dim pFile As String = ""
                    Dim pGeneralFileName = ""
                    If i = 0 Then
                        pFile = newKanbanNo.Trim
                        pFileName1 = pResult & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pFile & ".xlsm"
                        pGeneralFileName = pFileName1
                    End If
                    If i = 1 Then
                        pFile = newKanbanNo.Trim
                        pFileName2 = pResult & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pFile & ".xlsm"
                        pGeneralFileName = pFileName2
                    End If
                    If i = 2 Then
                        pFile = newKanbanNo.Trim
                        pFileName3 = pResult & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pFile & ".xlsm"
                        pGeneralFileName = pFileName3
                    End If
                    If i = 3 Then
                        pFile = newKanbanNo.Trim
                        pFileName4 = pResult & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pFile & ".xlsm"
                        pGeneralFileName = pFileName4
                    End If

                    ExcelBook.SaveAs(pGeneralFileName)
                    ExcelBook.Close()
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                    'My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

                End If
                xlApp.Workbooks.Close()
                xlApp.Quit()

                My.Computer.FileSystem.DeleteFile(NewFileCopyTO)
            Next
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

        ls_SQL = ls_SQL + "             SeqStart = RTRIM(CONVERT(NUMERIC, ISNULL(seqnoStart, 0))) , " & vbCrLf &
                          "             SeqEnd = RTRIM(CONVERT(NUMERIC, ISNULL(seqnoEnd, 0))) , " & vbCrLf &
                          "             PartNo1 = LEFT(Rtrim(KD.PartNo),2) , " & vbCrLf &
                          "             PartNo2 = SUBSTRING(Rtrim(KD.PartNo),3,9) , " & vbCrLf &
                          "             PartNo3 = SUBSTRING(Rtrim(KD.PartNo),12,10) , " & vbCrLf &
                          "             PartName = RTRIM(MP.PartName) , " & vbCrLf &
                          "             PartCMCode = RTRIM(ISNULL(MP.PartCarMaker, '')) , " & vbCrLf &
                          "             PartCMName = RTRIM(ISNULL(MP.PartGroupName, '')) , " & vbCrLf &
                          "             Qty = REPLACE(RTRIM(KD.POQtyBox), '.00', '') , " & vbCrLf &
                          "             BoxNo = RTRIM(ISNULL(KB.BoxNo, '')) , " & vbCrLf &
                          "             Cust = RTRIM(KM.AffiliateID), " & vbCrLf &
                          "             AFFCode = RTRIM(ISNULL(MA.AffiliateCode, '')) , " & vbCrLf &
                          "             Location = RTRIM(ISNULL(ML.LocationID, '')) , "

        ls_SQL = ls_SQL + "             SupplierID = RTRIM(KM.SupplierID) + '#1' , " & vbCrLf & _
                          "             SupplierCode = RTRIM(ISNULL(MS.SupplierCode, '')) , " & vbCrLf & _
                          "             Barcode = RTRIM(KB.barcode2) " & vbCrLf & _
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
                          "             AND CONVERT(CHAR(11), CONVERT(DATETIME, KanbanDate), 112) = '" & Format(ls_value1, "yyyyMMdd") & "' " & vbCrLf & _
                          "             AND KD.SupplierID = '" & Trim(ls_value4) & "' " & vbCrLf & _
                          "             AND KD.DeliveryLocationCode = '" & Trim(ls_value5) & "' " & vbCrLf & _
                          "             AND KD.kanbanno IN (" & Trim(ls_value2) & ") " & vbCrLf & _
                          "   )xx   ORDER BY  kanbanno, Rtrim(PartNo1)+Rtrim(partno2), BoxNo, cast(seqstart as numeric)   "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function bindDataHeaderKanban(ByVal GB As GlobalSetting.clsGlobal, ByVal pDate As Date, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pDeliveryLocation As String, ByVal pCycle As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = "Select kanbantime2 = convert(char(5),KanbanTime), * from kanban_Master where Isnull(excelcls,'') = '1' and KanbanDate = '" & Format(pDate, "yyyy-MM-dd") & "' " & vbCrLf & _
                 " AND AffiliateID = '" & pAffCode & "' and supplierID = '" & pSupplierID & "' and DeliveryLocationcode = '" & pDeliveryLocation & "' and ISNULL(KanbanSeq_No,1) = '" & pCycle & "' "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function bindHeaderETAETDKanban(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pSupplierCode As String, ByVal pDate As Date) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT AffiliateID, SupplierID, ETAPASI = CONVERT(CHAR(11),isnull(ETAPASI,''),106), ETDPASI = CONVERT(CHAR(11),isnull(ETDPASI,''),106), ETDSupplier = CONVERT(CHAR(11),isnull(ETDSUPPLIER,''),106)  " & vbCrLf & _
                 " FROM MS_ETD_PASI EP LEFT JOIN MS_ETD_Supplier_Pasi ES " & vbCrLf & _
                 " ON EP.ETDPASI = ES.ETAPASI WHERE AffiliateID = '" & Trim(pAffCode) & "' AND SupplierID = '" & Trim(pSupplierCode) & "' AND ETAAFFILIATE = '" & Format(pDate, "yyyy-MM-dd") & "'" & vbCrLf
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return (ds)
    End Function

    Shared Function bindDataDetailKanbanDelivery(ByVal GB As GlobalSetting.clsGlobal, ByVal pDate As Date, ByVal pAffCode As String, ByVal pSupplierCode As String, ByVal pDeliveryLocation As String, ByVal pCycle As Integer, ByVal pCycleSeq As Integer) As DataSet
        Dim ls_SQL As String = ""

        Dim C1 As String
        Dim C2 As String
        Dim C3 As String
        Dim C4 As String

        If pCycleSeq = "1" Then C1 = "1" : C2 = "2" : C3 = "3" : C4 = "4"
        If pCycleSeq = "2" Then C1 = "5" : C2 = "6" : C3 = "7" : C4 = "8"
        If pCycleSeq = "3" Then C1 = "9" : C2 = "10" : C3 = "11" : C4 = "12"
        If pCycleSeq = "4" Then C1 = "13" : C2 = "14" : C3 = "15" : C4 = "16"
        If pCycleSeq = "5" Then C1 = "17" : C2 = "18" : C3 = "19" : C4 = "20"

        ls_SQL = "  SELECT   colkanbanno, " & vbCrLf & _
                  "   	colpokanban = colpokanban, colno = ROW_NUMBER() OVER(ORDER BY cols DESC),   " & vbCrLf & _
                  "   	colpartno = colpartno,   " & vbCrLf & _
                  "   	coldescription=coldescription,   " & vbCrLf & _
                  "   	colpono=colpono,   " & vbCrLf & _
                  "   	coluom=coluom,   " & vbCrLf & _
                  "   	colqty=colqty,   " & vbCrLf & _
                  "   	colkanbanqty=colkanbanqty,   " & vbCrLf & _
                  "   	colcycle1 = colcycle1,   " & vbCrLf & _
                  "   	colcycle2 = colcycle2,   " & vbCrLf & _
                  "   	colcycle3 = colcycle3,    " & vbCrLf

        ls_SQL = ls_SQL + "   	colcycle4 = colcycle4,   " & vbCrLf & _
                          "   	colbox1 = CEILING(colbox1), " & vbCrLf & _
                          "     colbox2 = CEILING(colbox2), " & vbCrLf & _
                          "     colbox3 = CEILING(colbox3), " & vbCrLf & _
                          "     colbox4 = CEILING(colbox4), " & vbCrLf & _
                          "     colpallet1 = case when isnull(boxpallet,0) = 0 then 0 else CEILING(colbox1/boxpallet) END, " & vbCrLf & _
                          "     colpallet2 = case when isnull(boxpallet,0) = 0 then 0 else CEILING(colbox2/boxpallet) END, " & vbCrLf & _
                          "     colpallet3 = case when isnull(boxpallet,0) = 0 then 0 else CEILING(colbox3/boxpallet) END, " & vbCrLf & _
                          "     colpallet4 = case when isnull(boxpallet,0) = 0 then 0 else CEILING(colbox4/boxpallet) END, colpallet = boxpallet" & vbCrLf & _
                          "     ,colBarcode1 = convert(varchar(8000),colBarcode1), colbarcode2 = convert(varchar(8000),colBarcode2), colbarcode3 = convert(varchar(8000),colBarcode3), colbarcode4 =convert(varchar(8000),colBarcode4) " & vbCrLf & _
                          "     ,colNewPallet = round(colbox1 / boxpallet,2) " & vbCrLf & _
                          "   FROM (  " & vbCrLf & _
                          "  SELECT DISTINCT  colkanbanno = KM.KanbaNno, colpokanban = (case when isnull(KM.KanbanStatus,'') = '0' then 'YES' else 'NO' END), " & vbCrLf & _
                          "      cols = '1',  " & vbCrLf & _
                          "  	colno = '0',  " & vbCrLf & _
                          "  	colpartno = KD.partNo,  " & vbCrLf & _
                          "  	coldescription = MP.partname ,  " & vbCrLf & _
                          "  	colpono = KD.pono,   " & vbCrLf & _
                          "  	coluom = ISNULL(MUC.Description,''),  " & vbCrLf & _
                          "  	colmoq = KD.POMOQ,  " & vbCrLf

        ls_SQL = ls_SQL + "  	colqty = KD.POQtyBox,  " & vbCrLf & _
                          "  	colpoqty = COALESCE(PRD.POQty,PD.POQty),   " & vbCrLf & _
                          "  	colremainingpo= COALESCE(PRD.POQty,PD.POQty) - (SELECT SUM(ISNULL(KanbanQty,0)) FROM dbo.Kanban_Detail   " & vbCrLf & _
                          "  										WHERE PONo = PD.PoNo  										AND PartNo = PD.partNo),  " & vbCrLf & _
                          "  	colremainingsupplier=MSS.DailyDeliveryCapacity - (SELECT isnull(sum(KanbanQty),0) FROM dbo.Kanban_Detail A " & vbCrLf & _
                          " 						                                    LEFT JOIN dbo.Kanban_Master B ON A.KanbanNo = B.KanbanNo  " & vbCrLf & _
                          "                              							WHERE CONVERT(char(8), CONVERT(DATETIME,KanbanDate),112) = '20150502'  " & vbCrLf & _
                          "  														AND B.SupplierID = ''  AND A.PartNo = KD.PartNo) ,  " & vbCrLf & _
                          "  	coldeliveryqty = ISNULL(PD.DeliveryD2,0),  " & vbCrLf & _
                          "  	colkanbanqty= ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf

        ls_SQL = ls_SQL + "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C1 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  				+ ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf

        ls_SQL = ls_SQL + "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C2 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  				+ ISNULL((SELECT KanbanQty  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf

        ls_SQL = ls_SQL + "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C3 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  				+ ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf

        ls_SQL = ls_SQL + "  				WHERE KMI.KanbanCycle = " & C4 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  " & vbCrLf & _
                          "  	colcycle1 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  				WHERE KMI.KanbanCycle = " & C1 & "  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  " & vbCrLf & _
                          "  	colcycle2 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C2 & "  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.AffiliateID = KD.AffiliateID  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  " & vbCrLf & _
                          "  	colcycle3 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C3 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  	colcycle4 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C4 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0),  " & vbCrLf

        ls_SQL = ls_SQL + "  	colbox1 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C1 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0) / CASE WHEN ISNULL(KD.POQtyBox,0) = 0 then KD.POMOQ else KD.POQtyBox end, " & vbCrLf
        ls_SQL = ls_SQL + "  	colbox2 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C2 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0) / CASE WHEN ISNULL(KD.POQtyBox,0) = 0 then KD.POMOQ else KD.POQtyBox end, " & vbCrLf
        ls_SQL = ls_SQL + "  	colbox3 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C3 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0) / CASE WHEN ISNULL(KD.POQtyBox,0) = 0 then KD.POMOQ else KD.POQtyBox end, " & vbCrLf

        ls_SQL = ls_SQL + "  	colbox4 = ISNULL((SELECT KanbanQty  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C4 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0) / CASE WHEN ISNULL(KD.POQtyBox,0) = 0 then KD.POMOQ else KD.POQtyBox end, " & vbCrLf

        ls_SQL = ls_SQL + "   boxpallet = BoxPallet, " & vbCrLf & _
                          "  	cols1 = '1', coluomcode = MP.UnitCls  " & vbCrLf & _
                          "  	,kanbanno1= ISNULL((SELECT KMI.kanbanno  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C1 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbanno2= ISNULL((SELECT KMI.kanbanno  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C2 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbanno3= ISNULL((SELECT KMI.kanbanno  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C3 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbanno4= ISNULL((SELECT KMI.kanbanno  " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C4 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbantime1= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)   " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C1 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbantime2= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)   " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C2 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbantime3= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)   " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C3 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf

        ls_SQL = ls_SQL + "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf & _
                          "  	,kanbantime4= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)   " & vbCrLf & _
                          "  				FROM dbo.Kanban_Master  KMI  " & vbCrLf & _
                          "  					LEFT JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID  " & vbCrLf & _
                          "  						AND KMI.SupplierID = KDI.SupplierID  " & vbCrLf & _
                          "  						AND KMI.KanbanNo = KDI.KanbanNo AND KMI.DeliveryLocationCode = KDI.DeliveryLocationCode  " & vbCrLf & _
                          "  				WHERE KMI.KanbanCycle = " & C4 & "  " & vbCrLf & _
                          "  					AND KMI.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                      AND KDI.PartNo = KD.partNo  " & vbCrLf & _
                          "                      AND KDI.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND KMI.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND CONVERT(char(10), CONVERT(DATETIME,KMI.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'),0)  " & vbCrLf
        ls_SQL = ls_SQL + " , colbarcode1 = (select substring (( select  ' ;' + Rtrim(barcode) from  (  " & vbCrLf & _
                          " 										select distinct  barcode = Rtrim(KB.barcode), seqnostart from kanban_Barcode KB " & vbCrLf & _
                          " 										LEFT JOIN Kanban_Master KMR ON  " & vbCrLf & _
                          " 											KB.KanbanNo = KMR.KanbanNo and KB.AffiliateID = KMR.AffiliateID  " & vbCrLf & _
                          " 											AND KB.SupplierID = KMR.SupplierID " & vbCrLf & _
                          " 											AND KB.DeliveryLocationCode = KMR.DeliveryLocationCode  " & vbCrLf & _
                          " 										where KB.poNO = KD.PONo  " & vbCrLf & _
                          " 											and KB.affiliateID = KD.AffiliateID and KB.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf & _
                          " 											and KB.supplierID = KD.SupplierID  " & vbCrLf & _
                          " 											and KB.partno = KD.PartNo " & vbCrLf

        ls_SQL = ls_SQL + " 											and KanbanCycle = " & C1 & " and left(KB.kanbanno,8) = '" & Format(pDate, "yyyyMMdd") & "' )x" & vbCrLf & _
                          " 										order by convert(numeric,seqnostart)  " & vbCrLf & _
                          " 					FOR XML path(''), elements), 3, 5000)) " & vbCrLf & _
                          " 	, colbarcode2 = (select substring (( select  ' ;' + Rtrim(barcode) from  (  " & vbCrLf & _
                          " 										select distinct  barcode = Rtrim(KB.barcode), seqnostart from kanban_Barcode KB " & vbCrLf & _
                          " 										LEFT JOIN Kanban_Master KMR ON  " & vbCrLf & _
                          " 											KB.KanbanNo = KMR.KanbanNo and KB.AffiliateID = KMR.AffiliateID  " & vbCrLf & _
                          " 											AND KB.SupplierID = KMR.SupplierID " & vbCrLf & _
                          " 											AND KB.DeliveryLocationCode = KMR.DeliveryLocationCode  " & vbCrLf & _
                          " 										where KB.poNO = KD.PONo  " & vbCrLf & _
                          " 											and KB.affiliateID = KD.AffiliateID and KB.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf

        ls_SQL = ls_SQL + " 											and KB.supplierID = KD.SupplierID  " & vbCrLf & _
                          " 											and KB.partno = KD.PartNo " & vbCrLf & _
                          " 											and KanbanCycle = " & C2 & " and left(KB.kanbanno,8) = '" & Format(pDate, "yyyyMMdd") & "')x" & vbCrLf & _
                          " 										order by convert(numeric,seqnostart)  " & vbCrLf & _
                          " 					FOR XML path(''), elements), 3, 5000)) " & vbCrLf & _
                          " 	, colbarcode3 = (select substring (( select  ' ;' + Rtrim(barcode) from  (  " & vbCrLf & _
                          " 										select distinct  barcode = Rtrim(KB.barcode), seqnostart from kanban_Barcode KB " & vbCrLf & _
                          " 										LEFT JOIN Kanban_Master KMR ON  " & vbCrLf & _
                          " 											KB.KanbanNo = KMR.KanbanNo and KB.AffiliateID = KMR.AffiliateID  " & vbCrLf & _
                          " 											AND KB.SupplierID = KMR.SupplierID " & vbCrLf & _
                          " 											AND KB.DeliveryLocationCode = KMR.DeliveryLocationCode  " & vbCrLf

        ls_SQL = ls_SQL + " 										where KB.poNO = KD.PONo  " & vbCrLf & _
                          " 											and KB.affiliateID = KD.AffiliateID and KB.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf & _
                          " 											and KB.supplierID = KD.SupplierID  " & vbCrLf & _
                          " 											and KB.partno = KD.PartNo " & vbCrLf & _
                          " 											and KanbanCycle = " & C3 & " and left(KB.kanbanno,8) = '" & Format(pDate, "yyyyMMdd") & "')x " & vbCrLf & _
                          " 										order by convert(numeric,seqnostart)  " & vbCrLf & _
                          " 					FOR XML path(''), elements), 3, 5000))" & vbCrLf & _
                          " 	, colbarcode4 = (select substring (( select  ' ;' + Rtrim(barcode) from  (  " & vbCrLf & _
                          " 										select distinct  barcode = Rtrim(KB.barcode), seqnostart from kanban_Barcode KB " & vbCrLf & _
                          " 										LEFT JOIN Kanban_Master KMR ON  " & vbCrLf & _
                          " 											KB.KanbanNo = KMR.KanbanNo and KB.AffiliateID = KMR.AffiliateID  " & vbCrLf

        ls_SQL = ls_SQL + " 											AND KB.SupplierID = KMR.SupplierID " & vbCrLf & _
                          " 											AND KB.DeliveryLocationCode = KMR.DeliveryLocationCode  " & vbCrLf & _
                          " 										where KB.poNO = KD.PONo  " & vbCrLf & _
                          " 											and KB.affiliateID = KD.AffiliateID and KB.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf & _
                          " 											and KB.supplierID = KD.SupplierID  " & vbCrLf & _
                          " 											and KB.partno = KD.PartNo " & vbCrLf & _
                          " 											and KanbanCycle = " & C4 & " and left(KB.kanbanno,8) = '" & Format(pDate, "yyyyMMdd") & "')x" & vbCrLf & _
                          " 										order by convert(numeric,seqnostart)  " & vbCrLf & _
                          " 					FOR XML path(''), elements), 3, 5000)) " & vbCrLf
        ls_SQL = ls_SQL + "  FROM dbo.Kanban_Master KM   " & vbCrLf & _
                          "  	LEFT JOIN dbo.Kanban_Detail  KD ON KM.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                          "                                          AND KM.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "                                          AND KM.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "                                          AND KM.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf & _
                          "  	LEFT JOIN dbo.po_detailUpload PD ON KD.PartNo = PD.PartNo AND KD.PONo = PD.PONo And KD.AffiliateID = PD.AffiliateID And KD.SupplierID = PD.SupplierID " & vbCrLf & _
                          "      LEFT JOIN PO_Master PM ON PM.PoNo = PD.PONo  " & vbCrLf & _
                          "                                  and PM.AffiliateID = PD.AffiliateID  " & vbCrLf & _
                          "                                  and PM.SupplierID = PD.SupplierID  " & vbCrLf

        ls_SQL = ls_SQL + "      LEFT JOIN dbo.PORev_Master PRM ON PM.AffiliateID = PRM.AffiliateID  " & vbCrLf & _
                          "                                      AND PRM.PONo = PM.PONo  " & vbCrLf & _
                          "                                      AND PRM.SupplierID = PM.SupplierID  " & vbCrLf & _
                          "      LEFT JOIN dbo.PORev_Detail PRD ON PRD.PONo = PRM.PONo  " & vbCrLf & _
                          "                                      AND PRD.AffiliateID = PRM.AffiliateID  " & vbCrLf & _
                          "                                      AND PRD.SupplierID = PRM.SupplierID  " & vbCrLf & _
                          "                                      AND PRD.PartNo = PD.PartNo  " & vbCrLf & _
                          "                                      AND PRD.SeqNo = (SELECT MAX(seqNO) FROM PORev_Detail A WHERE " & vbCrLf & _
                          "                                                          A.PONo = PD.PONo  " & vbCrLf & _
                          " 							                                AND A.AffiliateID = PD.AffiliateID   " & vbCrLf & _
                          " 							                                AND A.SupplierID = PD.SupplierID   " & vbCrLf

        ls_SQL = ls_SQL + " 							                                AND A.PartNo = PD.PartNo)  " & vbCrLf & _
                          "  	LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo  " & vbCrLf & _
                          "     LEFT JOIN MS_PartMapping MPM ON MPM.Partno = KD.PartNo and MPM.AffiliateID = KD.AffiliateID and MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                          "      LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls  " & vbCrLf & _
                          "      LEFT JOIN dbo.MS_SupplierCapacity MSS ON MSS.SupplierID = PD.SupplierID AND MSS.PartNo = PD.PartNo	 " & vbCrLf & _
                          "  WHERE CONVERT(char(10), CONVERT(DATETIME,KM.KanbanDate),120) = '" & Format(pDate, "yyyy-MM-dd") & "'  " & vbCrLf & _
                          "   AND KM.AffiliateID = '" & pAffCode & "' And KM.DeliveryLocationcode = '" & pDeliveryLocation & "' AND KM.SupplierID = '" & Trim(pSupplierCode) & "' and KM.ExcelCls=1)xx " & _
                          "    where colkanbanqty <> 0 AND colcycle" & pCycle & " <> 0 order by colpono asc" & vbCrLf

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                              ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String, _
                                              Optional ByVal pKanban As String = "", Optional ByVal pDN1 As String = "", Optional ByVal pDN2 As String = "", _
                                              Optional ByVal pDN3 As String = "", Optional ByVal pDN4 As String = "", Optional ByVal pBarcodeFile As String = "") As Boolean
        Dim TempFilePath As String = Trim(pPathFile)
        Dim total_FileName As String = ""

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            Dim ls_DN1 As String = "", ls_DN2 As String = "", ls_DN3 As String = "", ls_DN4 As String = ""

            sendEmailtoSupplier = True

            If pDN1 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN1
                Else
                    total_FileName = total_FileName & "," & pDN1
                End If
                ls_DN1 = pPathFile & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pDN1 & ".xlsm"
            End If

            If pDN2 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN2
                Else
                    total_FileName = total_FileName & "," & pDN2
                End If
                ls_DN2 = pPathFile & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pDN2 & ".xlsm"
            End If

            If pDN3 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN3
                Else
                    total_FileName = total_FileName & "," & pDN3
                End If
                ls_DN3 = pPathFile & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pDN3 & ".xlsm"
            End If

            If pDN4 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN4
                Else
                    total_FileName = total_FileName & "," & pDN4
                End If
                ls_DN4 = pPathFile & "\Delivery " & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & pDN4 & ".xlsm"
            End If

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
                errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Send To Supplier Kanban No: " & total_FileName & "-" & pAffiliate.Trim & "-" & pSupplier.Trim

            ls_Body = clsNotification.GetNotification("30", "", "", total_FileName)
            ls_Attachment = Trim(pPathFile) & "\" & pKanban

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, pSupplier, ls_Subject, ls_Body, errMsg, ls_Attachment, IIf(pDN1 = "", "", ls_DN1), IIf(pDN2 = "", "", ls_DN2), IIf(pDN3 = "", "", ls_DN3), IIf(pDN4 = "", "", ls_DN4), IIf(pBarcodeFile = "", "", pBarcodeFile)) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Function sendEmailtoAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                              ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pKanbanDate As Date, ByVal pDeliveryLocation As String, ByRef errMsg As String, _
                                              Optional ByVal pKanban As String = "", Optional ByVal pDN1 As String = "", Optional ByVal pDN2 As String = "", _
                                              Optional ByVal pDN3 As String = "", Optional ByVal pDN4 As String = "", Optional ByVal pBarcodeFile As String = "") As Boolean
        Dim total_FileName As String = ""

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""


            sendEmailtoAffiliate = True

            If pKanban <> "" Then
                If total_FileName = "" Then
                    total_FileName = pKanban
                Else
                    total_FileName = total_FileName & "," & pKanban
                End If
            End If

            If pDN1 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN1
                Else
                    total_FileName = total_FileName & "," & pDN1
                End If
            End If

            If pDN2 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN2
                Else
                    total_FileName = total_FileName & "," & pDN2
                End If
            End If

            If pDN3 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN3
                Else
                    total_FileName = total_FileName & "," & pDN3
                End If
            End If

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
                errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If


            ls_URl = "http://" & clsNotification.pub_ServerName & "/Kanban/KanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pKanbanDate.Date) & "&t1=" & clsNotification.EncryptURL(pSupplier) & _
                                       "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/Kanban/KanbanList.aspx")

            ls_Subject = "Send To Supplier Kanban No: " & total_FileName & "-" & pAffiliate.Trim & "-" & pSupplier.Trim

            ls_Body = clsNotification.GetNotification("30", ls_URl, "", total_FileName)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoAffiliate = False
                Exit Function
            End If

            sendEmailtoAffiliate = True

        Catch ex As Exception
            sendEmailtoAffiliate = False
            errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Function sendEmailtoPASI(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, _
                                        ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pKanbanDate As Date, ByVal pDeliveryLocation As String, ByRef errMsg As String, _
                                        Optional ByVal pKanban As String = "", Optional ByVal pDN1 As String = "", Optional ByVal pDN2 As String = "", _
                                        Optional ByVal pDN3 As String = "", Optional ByVal pDN4 As String = "", Optional ByVal pBarcodeFile As String = "") As Boolean
        Dim total_FileName As String = ""

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoPASI = True

            If pKanban <> "" Then
                If total_FileName = "" Then
                    total_FileName = pKanban
                Else
                    total_FileName = total_FileName & "," & pKanban
                End If
            End If

            If pDN1 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN1
                Else
                    total_FileName = total_FileName & "," & pDN1
                End If
            End If

            If pDN2 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN2
                Else
                    total_FileName = total_FileName & "," & pDN2
                End If
            End If

            If pDN3 <> "" Then
                If total_FileName = "" Then
                    total_FileName = pDN3
                Else
                    total_FileName = total_FileName & "," & pDN3
                End If
            End If

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
                errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to PASI [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to PASI [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerNamePASI & "/AffKanban/AffKanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pKanbanDate.Date) & "&t1=" & clsNotification.EncryptURL(pSupplier) & _
                                       "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/AffKanban/AffKanbanList.aspx")

            ls_Subject = "Send To Supplier Kanban No: " & pAffiliate.Trim & "-" & total_FileName & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("30", ls_URl, "", total_FileName)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoPASI = False
                Exit Function
            End If

            sendEmailtoPASI = True

        Catch ex As Exception
            sendEmailtoPASI = False
            errMsg = "Process Send PO Kanban [" & total_FileName & "] from Affiliate [" & pAffiliate & "] to PASI [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Sub UpdateExcelKanban(ByVal KKanbandate As Date, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal KDeliveryLocation As String, ByVal KALLKANBANNO As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = "UPDATE dbo.Kanban_Master SET ExcelCls='2' WHERE CONVERT(char(10), CONVERT(DATETIME,KanbanDate),120)='" & Format(KKanbandate, "yyyy-MM-dd") & "' AND AffiliateID='" & pAffiliate & "' AND SupplierID='" & pSupplier & "' AND DeliveryLocationCode = '" & KDeliveryLocation & "' AND Kanbanno IN (" & KALLKANBANNO & ") "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
                errMsg = ls_SQL & "OK"
            End Using
        Catch ex As Exception
            errMsg = "Process Send PO Kanban: Affiliate [" & pAffiliate & "], Supplier [" & pSupplier & "], KanbanDate [" & KKanbandate & "] to Supplier STOPPED, because " & ex.Message & " and query " & ls_SQL
        End Try
    End Sub

    Shared Sub DeleteKanbanSuppPASI()

        Dim ls_SQL As String
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " Delete Kanban_master where supplierID = 'PASI' " & vbCrLf & _
                         " Delete Kanban_Detail where SupplierID = 'PASI' " & vbCrLf & _
                         " Delete UploadKanban where supplierID = 'PASI' " & vbCrLf & _
                         " Delete PO_Detail where supplierID = 'PASI' " & vbCrLf & _
                         " Delete PO_master where supplierID = 'PASI' " & vbCrLf & _
                         " Delete Affiliate_master where supplierID = 'PASI' " & vbCrLf & _
                         " Delete Affiliate_Detail where supplierID = 'PASI' " & vbCrLf & _
                         " Delete PO_MasterUpload where supplierID = 'PASI' " & vbCrLf & _
                         " Delete PO_DetailUpload where supplierID = 'PASI' " & vbCrLf
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception

        End Try
    End Sub
End Class
