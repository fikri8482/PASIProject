Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports System.Threading

Public Class clsDeliveryRemaining
    Shared Sub up_SendRemainingDelivery(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1




        Dim ls_sql As String = ""
        Dim pNamaFile As String = ""

        Dim NewFileCopy As String = ""
        Dim NewFileCopyTO As String = ""

        Dim pSuratJalanNo As String = ""
        Dim pSupplier As String = ""
        Dim pAffiliate As String = ""
        Dim pDeliveryLocation As String = ""

        Dim KKanbandate As String = ""

        Dim fromEmail As String = ""
        Dim receiptCCEmail As String

        Dim ds As New DataSet
        Dim dsDetailDelivery As New DataSet
        Dim dsEmail As New DataSet
        Dim dsSupp As New DataSet
        Dim dsAff As New DataSet
        Dim dsETAETD As New DataSet

        Dim k As Integer

        Try
            Dim fi As New FileInfo(Trim(pAtttacment) & "\Template Delivery.xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Template Delivery (PO Kanban), File Excel isn't Found"
                ErrSummary = "Template Delivery (PO Kanban), File Excel isn't Found"
                Exit Sub
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Get data Remaining Delivery")

            ls_sql = "select DISTINCT SupplierID, RPM.AffiliateID, SuratJalanNo, MD.DeliveryLocationCode " & vbCrLf & _
                     "from ReceivePASI_Master RPM LEFT JOIN MS_DeliveryPlace MD ON MD.AffiliateID = RPM.AffiliateID" & vbCrLf & _
                     "where ISNULL(RemainingCls,'1') = '1'"

            ds = GB.uf_GetDataSet(ls_sql)

            If ds.Tables(0).Rows.Count > 0 Then
                For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                    Try
                        Dim ETDSupplier As String = ""
                        pSuratJalanNo = Trim(ds.Tables(0).Rows(i_loop)("suratjalanno"))
                        pSupplier = Trim(ds.Tables(0).Rows(i_loop)("SupplierID"))
                        pAffiliate = Trim(ds.Tables(0).Rows(i_loop)("AffiliateID"))
                        pDeliveryLocation = Trim(ds.Tables(0).Rows(i_loop)("DeliveryLocationCode"))

                        log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel Remaining Delivery [" & pSuratJalanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")
                        'Thread.Sleep(1000)
                        NewFileCopy = pAtttacment & "\Template Delivery.xlsm"
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

                        dsDetailDelivery = bindDataDetailRemaining(GB, pSuratJalanNo, pAffiliate)

                        If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                            ExcelBook = xlApp.Workbooks.Open(ls_file)
                            ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                            log.WriteToProcessLog(Date.Now, pScreenName, "Input Header Excel Affiliate [" & pAffiliate & "-" & pSupplier & "-" & pSuratJalanNo & "]")

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
                            ExcelSheet.Range("AK37:AN38").Value = "REMAINING DELIVERY PLAN QTY"
                            ExcelSheet.Range("AO37:AR38").Value = "REMAINING DELIVERY QTY"

                            Dim newKanbanNo As String = ""

                            For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                                k = k

                                KKanbandate = dsDetailDelivery.Tables(0).Rows(j)("KanbanDate") 'Microsoft.VisualBasic.Left(Trim(dsDetailDelivery.Tables(0).Rows(j)("kanbanno")), 4) + "-" + Microsoft.VisualBasic.Mid(Trim(dsDetailDelivery.Tables(0).Rows(j)("kanbanno")), 5, 2) + "-" + Microsoft.VisualBasic.Mid(Trim(dsDetailDelivery.Tables(0).Rows(j)("kanbanno")), 7, 2)
                                dsETAETD = BindHeaderETAETDNonKanban(GB, pAffiliate, pSupplier, KKanbandate)

                                If dsETAETD.Tables(0).Rows.Count > 0 Then
                                    ETDSupplier = dsETAETD.Tables(0).Rows(0)("ETDSupplier")
                                Else
                                    ETDSupplier = ""
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
                                ExcelSheet.Range("D" & k + 39 & ": H" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("pono")
                                ExcelSheet.Range("i" & k + 39 & ": K" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("POKanbanCls")
                                ExcelSheet.Range("L" & k + 39 & ": O" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("kanbanno")
                                ExcelSheet.Range("P" & k + 39 & ": T" & k + 39).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partno"))
                                ExcelSheet.Range("U" & k + 39 & ": AC" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("partname")
                                ExcelSheet.Range("AD" & k + 39 & ": AE" & k + 39).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("uom"))
                                ExcelSheet.Range("AF" & k + 39 & ": AG" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("qtybox")
                                ExcelSheet.Range("AH" & k + 39 & ": AJ" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("boxpallet")
                                ExcelSheet.Range("AK" & k + 39 & ": AN" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("RemainingQty")
                                ExcelSheet.Range("AO" & k + 39 & ": AR" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("RemainingQty")
                                ExcelSheet.Range("AS" & k + 39 & ": AV" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("qtyboxdelivery")
                                ExcelSheet.Range("AW" & k + 39 & ": AZ" & k + 39).Value = dsDetailDelivery.Tables(0).Rows(j)("qtypallet")
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
                            'Dim pFile As String = newKanbanNo.Trim
                            Dim ls_SJ As String = ""
                            If pSuratJalanNo.Contains("\") Then
                                ls_SJ = Replace(pSuratJalanNo, "\", "_")
                            Else
                                ls_SJ = Replace(pSuratJalanNo, "/", "_")
                            End If
                            pNamaFile = "Delivery Remaining" & Trim(pAffiliate) & "-" & Trim(pSupplier) & "-" & ls_SJ & ".xlsm"

                            ExcelBook.SaveAs(pResult & "\" & pNamaFile)
                            ExcelBook.Close()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()
                            My.Computer.FileSystem.DeleteFile(NewFileCopyTO)

                            log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel Remaining Delivery [" & pSuratJalanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")

                            If sendEmailtoSupplier(GB, pResult, pNamaFile, pSuratJalanNo, pAffiliate, pSupplier, errMsg) = False Then
                                Exit Try
                            Else
                                log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. Remaining Delivery [" & pSuratJalanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            End If

                            Call UpdateStatusRemainingDelivery(pSuratJalanNo, pSupplier, pAffiliate, errMsg)

                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email. Remaining Delivery [" & pSuratJalanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send Delivery Remaining [" & pSuratJalanNo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
                            LogName.Refresh()
                        Else
                            Call UpdateStatusRemainingDelivery(pSuratJalanNo, pSupplier, pAffiliate, errMsg)
                        End If
                    Catch ex As Exception
                        xlApp.DisplayAlerts = False
                        ExcelBook.Close()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                        My.Computer.FileSystem.DeleteFile(NewFileCopyTO)
                        log.WriteToErrorLog(pScreenName, "Process Create Remaining Delivery [" & pAffiliate & "-" & pSuratJalanNo & "-" & pSupplier & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Create Remaining Delivery [" & pAffiliate & "-" & pSuratJalanNo & "-" & pSupplier & "] STOPPED, because " & ex.Message)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Send Remaining Delivery [" & pAffiliate & "-" & pSuratJalanNo & "-" & pSupplier & "] STOPPED, because " & ex.Message, LogName)
                        LogName.Refresh()
                    End Try
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "Affiliate [" & pAffiliate & "]" & ", Supplier [" & pSupplier & "]" & ", KanbanNo [" & pSuratJalanNo & "] " & ", " & ex.Message
            ErrSummary = "Affiliate [" & pAffiliate & "]" & ", Supplier [" & pSupplier & "]" & ", KanbanNo [" & pSuratJalanNo & "] " & ", " & ex.Message
        Finally
            If Not xlApp Is Nothing Then
                clsGeneral.NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                clsGeneral.NAR(ExcelBook)
                xlApp.Quit()
                clsGeneral.NAR(xlApp)
                GC.Collect()
            End If
        End Try

    End Sub

    Shared Function bindDataDetailRemaining(ByVal GB As GlobalSetting.clsGlobal, ByVal pSuratJalanNo As String, ByVal pAffiliateCode As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT  " & vbCrLf & _
                  " 	xyz.AffiliateID, " & vbCrLf & _
                  " 	xyz.SupplierID, " & vbCrLf & _
                  " 	POKanbanCls = CASE WHEN abc.POKanbanCls = '1' then 'YES' else 'NO' end, " & vbCrLf & _
                  " 	xyz.PONo, " & vbCrLf & _
                  " 	xyz.KanbanNo, " & vbCrLf & _
                  " 	xyz.PartNo, " & vbCrLf & _
                  " 	MP.PartName, " & vbCrLf & _
                  " 	MU.Description UOM, " & vbCrLf & _
                  " 	QtyBox = ISNULL(xyz.POQtyBox,MPM.QtyBox), " & vbCrLf & _
                  " 	MPM.BoxPallet, "

        ls_SQL = ls_SQL + " 	xyz.KanbanQty, " & vbCrLf & _
                          " 	xyz.KanbanQty - abc.TotalReceiving RemainingQty, " & vbCrLf & _
                          " 	Ceiling((xyz.KanbanQty - abc.TotalReceiving) / ISNULL(xyz.POQtyBox,MPM.QtyBox)) as qtyboxdelivery,	 " & vbCrLf & _
                          " 	ROUND(((xyz.KanbanQty - abc.TotalReceiving) / ISNULL(xyz.POQtyBox,MPM.QtyBox))  / boxpallet,2) as qtypallet, km.KanbanDate	 " & vbCrLf & _
                          " FROM " & vbCrLf & _
                          " ( " & vbCrLf & _
                          " 	--select PartNo and Kanban No must be send Remaining " & vbCrLf & _
                          " 	select DISTINCT RD.SupplierID, RD.AffiliateID, RD.PONo, RD.KanbanNo, RD.PartNo, KD.KanbanQty, KD.POMOQ, KD.POQtyBox " & vbCrLf & _
                          " 	from ReceivePASI_Detail RD " & vbCrLf & _
                          " 	left join Kanban_Detail KD ON RD.AffiliateID = KD.AffiliateID and RD.SupplierID	= KD.SupplierID " & vbCrLf & _
                          " 								 and RD.PartNo = KD.PartNo and RD.KanbanNo = KD.KanbanNo and RD.PONo = KD.PONo "

        ls_SQL = ls_SQL + " 	where RD.SuratJalanNo = '" & Trim(pSuratJalanNo) & "' and RD.AffiliateID = '" & Trim(pAffiliateCode) & "'" & vbCrLf & _
                          " )xyz " & vbCrLf & _
                          " LEFT JOIN " & vbCrLf & _
                          " ( " & vbCrLf & _
                          " 	select SupplierID, AffiliateID, PONo, KanbanNo, PartNo, POKanbanCls, SUM(GoodRecQty) TotalReceiving " & vbCrLf & _
                          " 	from ReceivePASI_Detail	 " & vbCrLf & _
                          " 	GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo, POKanbanCls  " & vbCrLf & _
                          " )abc ON abc.AffiliateID = xyz.AffiliateID and abc.SupplierID = xyz.SupplierID and abc.KanbanNo = xyz.KanbanNo and abc.PONo = xyz.PONo and xyz.PartNo = abc.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_Parts MP ON MP.PartNo = XYZ.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = XYZ.PartNo and MPM.AffiliateID =XYZ.AffiliateID and MPM.SupplierID = XYZ.SupplierID " & vbCrLf & _
                          " LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls "

        ls_SQL = ls_SQL + " LEFT JOIN Kanban_Master KM ON KM.AffiliateID = XYZ.AffiliateID and KM.SupplierID = XYZ.SupplierID and KM.KanbanNo = XYZ.KanbanNo " & vbCrLf & _
                          " WHERE (xyz.KanbanQty - abc.TotalReceiving) > 0 " & vbCrLf & _
                          "  "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return (ds)
    End Function

    Shared Function BindHeaderETAETDNonKanban(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pSupplierCode As String, ByVal pDate As Date) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT distinct AffiliateID, SupplierID, ETAPASI = CONVERT(CHAR(11),isnull(ETAPASI,''),106), ETDPASI = CONVERT(CHAR(11),isnull(ETDPASI,''),106), ETDSupplier = CONVERT(CHAR(11),isnull(ETDSUPPLIER,''),106)  " & vbCrLf & _
                 " FROM MS_ETD_PASI EP LEFT JOIN MS_ETD_Supplier_Pasi ES " & vbCrLf & _
                 " ON EP.ETDPASI = ES.ETAPASI WHERE AffiliateID = '" & Trim(pAffCode) & "' AND SupplierID = '" & Trim(pSupplierCode) & "' AND ETAAFFILIATE = '" & Format(pDate, "yyyy-MM-dd") & "'" & vbCrLf
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return (ds)
    End Function

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
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
                errMsg = "Process Send PO [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send PO [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Send To Supplier Remaining Delivery SuratJalanNo : " & Trim(pPONo) & "-" & pSupplier

            ls_Body = "Dear Sir/Madam,  " & vbCrLf & _
                      " " & vbCrLf & _
                      "This is notification for: " & vbCrLf & _
                      "Remaining Delivery SuratJalanNo : (" & Trim(pPONo) & ") " & vbCrLf & _
                      " " & vbCrLf & _
                      " " & vbCrLf & _
                      " " & vbCrLf & _
                      "Best Regard" & vbCrLf & _
                      "PO System" & vbCrLf

            ls_Attachment = Trim(pPathFile) & "\" & pFileName

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, pSupplier, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send PO [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Sub UpdateStatusRemainingDelivery(ByVal tSuratJalanNo As String, ByVal tSupp As String, ByVal tAff As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " Update dbo.ReceivePASI_Master set RemainingCls = '2'" & vbCrLf & _
                         " WHERE SuratJalanNo = '" & Trim(tSuratJalanNo) & "' and AffiliateID = '" & Trim(tAff) & "'" & vbCrLf & _
                         " AND SupplierID = '" & Trim(tSupp) & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send Remaining [" & tSuratJalanNo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub

End Class
