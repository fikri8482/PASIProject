Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsReceivingFWD
    Shared Sub up_ReceivingFWD(ByVal cfg As GlobalSetting.clsConfig,
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
        Dim xlApp = New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim i As Integer, k As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim fromEmail As String = ""
        Dim NewFileCopy As String
        Dim NewFileCopyTO As String

        Dim ls_SQL As String = ""
        Dim pAffCode As String = ""
        Dim pSupplier As String = ""
        Dim pPONo As String = ""
        Dim pOrderNo As String = ""
        Dim pFWD As String = ""

        Dim pSupplierName As String = ""
        Dim pSupplierAdd As String = ""
        Dim pFWDName As String = ""
        Dim pFWDAdd As String = ""
        Dim pAffName As String = ""
        Dim pAffAdd As String = ""
        Dim pAttn As String = ""
        Dim pTelp As String = ""
        Dim pPeriod As String = ""
        Dim pETDVendor As String = ""
        Dim pETDPort As String = ""
        Dim pETAPort As String = ""
        Dim pETAFactory As String = ""
        Dim pSuratJalanNo As String = ""

        Dim pConsigneeCode As String = ""
        Dim pConsigneeName As String = ""
        Dim pConsigneeAdd As String = ""
        Dim pCommercial As String = ""
        Dim dPeriod As Date

        Dim ds As New DataSet
        Dim dsDetailDelivery As New DataSet
        Dim dsEmail As New DataSet
        Dim dsSupp As New DataSet
        Dim dsAffp As New DataSet
        Dim dsSplit As New DataSet

        Dim pFileName1 As String = ""
        Dim pFileName2 As String = ""
        Dim pExcelCls As String = ""
        Dim booSplit As Boolean

        Try
            Dim fi As New FileInfo(pAtttacment & "\TEMPLATE DELIVERY NOTE FORWARDER_EXPORT.xlsm")
            If Not fi.Exists Then
                errMsg = "Process Send Receiving Forwarder Confirmation to Supplier STOPPED, File Excel isn't Found"
                ErrSummary = "Process Send Receiving Forwarder Confirmation STOPPED, File Excel isn't Found"
                Exit Try
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Get data GR")

            ls_SQL = " select distinct Consignee = isnull(MA.ConsigneeCode,''), attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), period = PME.Period, DOM.SuratJalanNo, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') + ' ' + isnull(MS.City,'') + ' ' + isnull(MS.Postalcode,'') SUPPAddress,  " & vbCrLf & _
                     " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  + ' ' + isnull(MF.City,'') + ' ' + isnull(MF.PostalCode,'') as FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                     " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'')  + ' ' + isnull(MA.City,'') + ' ' + isnull(MA.PostalCode,'') as AFFAddress, ISNULL(DOM.MovingList,0) MovingList, ConsigneName = isnull(MA.ConsigneeName,''), ConsigneeAdd = Rtrim(Isnull(MA.ConsigneeAddress,'')), isnull(DOM.ExcelCls,0) ExcelCls, ISNULL(DOM.SplitReffPONo, '') SplitReffPONo, ISNULL(DOM.CommercialCls,'1') CommercialCls from  " & vbCrLf & _
                     " DOSUpplier_Master_Export DOM LEFT JOIN PO_Master_Export PME  " & vbCrLf & _
                     " ON  DOM.PONo = PME.PONo and DOM.SupplierID = PME.SupplierID and DOM.AffiliateID = PME.AffiliateID AND DOM.OrderNo = PME.OrderNo1" & vbCrLf & _
                     " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                     " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                     " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                     " where isnull(DOM.ExcelCls,0) IN ('1', '3') and isnull(DOM.PONO,'') <> '' and DOM.SuratJalanno <> '' "

            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                    pFileName1 = ""
                    pFileName2 = ""
                    pExcelCls = Trim(ds.Tables(0).Rows(i_loop)("ExcelCls"))

                    If ds.Tables(0).Rows(i_loop)("CommercialCls") = "1" Then
                        pCommercial = "YES"
                    Else
                        pCommercial = "NO"
                    End If

                    If pExcelCls = "3" Then
                        booSplit = True

                        ls_SQL = " select distinct Consignee = isnull(MA.ConsigneeCode,''), attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), period = PME.Period, DOM.SuratJalanNo, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') + ' ' + isnull(MS.City,'') + ' ' + isnull(MS.Postalcode,'') SUPPAddress,  " & vbCrLf & _
                                 " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  + ' ' + isnull(MF.City,'') + ' ' + isnull(MF.PostalCode,'') as FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                                 " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'')  + ' ' + isnull(MA.City,'') + ' ' + isnull(MA.PostalCode,'') as AFFAddress, ISNULL(DOM.MovingList,0) MovingList, ConsigneName = isnull(MA.ConsigneeName,''), ConsigneeAdd = Rtrim(Isnull(MA.ConsigneeAddress,'')), isnull(DOM.ExcelCls,0) ExcelCls, ISNULL(DOM.SplitReffPONo, '') SplitReffPONo from  " & vbCrLf & _
                                 " DOSUpplier_Master_Export DOM LEFT JOIN PO_Master_Export PME  " & vbCrLf & _
                                 " ON  DOM.PONo = PME.PONo and DOM.SupplierID = PME.SupplierID and DOM.AffiliateID = PME.AffiliateID AND DOM.OrderNo = PME.OrderNo1" & vbCrLf & _
                                 " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                                 " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                                 " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                                 " WHERE DOM.PONo = '" & Trim(ds.Tables(0).Rows(i_loop)("PONo")) & "' " & vbCrLf & _
                                 " AND DOM.AffiliateID = '" & Trim(ds.Tables(0).Rows(i_loop)("AffiliateID")) & "' " & vbCrLf & _
                                 " AND DOM.SupplierID = '" & Trim(ds.Tables(0).Rows(i_loop)("SupplierID")) & "' " & vbCrLf & _
                                 " AND DOM.OrderNo = '" & Trim(ds.Tables(0).Rows(i_loop)("SplitReffPONo")) & "' "

                        dsSplit = GB.uf_GetDataSet(ls_SQL)

                        pSuratJalanNo = Trim(dsSplit.Tables(0).Rows(0)("SuratJalanNo"))
                        pSupplier = Trim(dsSplit.Tables(0).Rows(0)("supplierID"))
                        pSupplierName = Trim(dsSplit.Tables(0).Rows(0)("suppliername"))
                        pSupplierAdd = Trim(dsSplit.Tables(0).Rows(0)("SuppAddress"))
                        pFWD = Trim(dsSplit.Tables(0).Rows(0)("ForwarderID"))
                        pFWDName = Trim(dsSplit.Tables(0).Rows(0)("ForwarderName"))
                        pFWDAdd = Trim(dsSplit.Tables(0).Rows(0)("FWDAddress"))
                        pPONo = Trim(dsSplit.Tables(0).Rows(0)("PONo"))
                        pOrderNo = Trim(dsSplit.Tables(0).Rows(0)("OrderNo"))
                        pETDVendor = Format((dsSplit.Tables(0).Rows(0)("ETDVendor")), "yyyy-MM-dd")
                        pETDPort = Format((dsSplit.Tables(0).Rows(0)("ETDPort")), "yyyy-MM-dd")
                        pETAPort = Format((dsSplit.Tables(0).Rows(0)("ETAPort")), "yyyy-MM-dd")
                        pETAFactory = Format((dsSplit.Tables(0).Rows(0)("ETAFactory")), "yyyy-MM-dd")
                        pAffCode = Trim(dsSplit.Tables(0).Rows(0)("AFF"))
                        pAffName = Trim(dsSplit.Tables(0).Rows(0)("AFFName"))
                        pAffAdd = Trim(dsSplit.Tables(0).Rows(0)("AFFAddress"))
                        pAttn = Trim(dsSplit.Tables(0).Rows(0)("attn"))
                        pTelp = Trim(dsSplit.Tables(0).Rows(0)("telp"))
                        pConsigneeCode = Trim(dsSplit.Tables(0).Rows(0)("Consignee"))
                        pConsigneeName = Trim(dsSplit.Tables(0).Rows(0)("Consignename"))
                        pConsigneeAdd = Trim(dsSplit.Tables(0).Rows(0)("ConsigneeAdd"))
                        dPeriod = ds.Tables(0).Rows(0)("Period")
                    Else
Split:
                        booSplit = False

                        pSuratJalanNo = Trim(ds.Tables(0).Rows(i_loop)("SuratJalanNo"))
                        pSupplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                        pSupplierName = Trim(ds.Tables(0).Rows(i_loop)("suppliername"))
                        pSupplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))
                        pFWD = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))
                        pFWDName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                        pFWDAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))
                        pPONo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                        pOrderNo = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))
                        pETDVendor = Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")
                        pETDPort = Format((ds.Tables(0).Rows(i_loop)("ETDPort")), "yyyy-MM-dd")
                        pETAPort = Format((ds.Tables(0).Rows(i_loop)("ETAPort")), "yyyy-MM-dd")
                        pETAFactory = Format((ds.Tables(0).Rows(i_loop)("ETAFactory")), "yyyy-MM-dd")
                        pAffCode = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                        pAffName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                        pAffAdd = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                        pAttn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                        pTelp = Trim(ds.Tables(0).Rows(i_loop)("telp"))
                        pConsigneeCode = Trim(ds.Tables(0).Rows(i_loop)("Consignee"))
                        pConsigneeName = Trim(ds.Tables(0).Rows(i_loop)("Consignename"))
                        pConsigneeAdd = Trim(ds.Tables(0).Rows(i_loop)("ConsigneeAdd"))
                        dPeriod = ds.Tables(0).Rows(i_loop)("Period")
                    End If

                    log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel GR [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "]")

                    NewFileCopy = pAtttacment & "\TEMPLATE DELIVERY NOTE FORWARDER_EXPORT.xlsm"
                    NewFileCopyTO = pResult & "\TEMPLATE DELIVERY NOTE FORWARDER_EXPORT " & Format(Now, "HHmmss") & ".xlsm"

                    If System.IO.File.Exists(NewFileCopy) = True Then
                        System.IO.File.Copy(NewFileCopy, NewFileCopyTO)
                    Else
                        System.IO.File.Copy(NewFileCopy, pResult & "\Delivery.xlsm")
                    End If

                    Dim ls_file As String = NewFileCopyTO

                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    dsDetailDelivery = BindDataDeliveryInstruction(GB, pSuratJalanNo, pPONo, pOrderNo, pAffCode, pSupplier)

                    If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                        dsEmail = clsGeneral.getEmailAddressExport(GB, "", "PASI", "", "", "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)
                        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                            If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                                fromEmail = dsEmail.Tables(0).Rows(i)("EmailFrom")
                                receiptCCEmail = dsEmail.Tables(0).Rows(i)("EmailCC")
                            End If
                        Next

                        ExcelSheet.Range("H2").Value = fromEmail.Trim
                        'ExcelSheet.Range("Y2").Value = receiptCCEmail
                        ExcelSheet.Range("H3").Value = pConsigneeCode.Trim
                        ExcelSheet.Range("H4").Value = pFWD.Trim
                        ExcelSheet.Range("H5").Value = pSupplier.Trim

                        ExcelSheet.Range("I11:X11").Merge()
                        ExcelSheet.Range("I11:X11").Value = pSupplierName.Trim
                        ExcelSheet.Range("I12:X15").Merge()
                        ExcelSheet.Range("I12:X15").Value = pSupplierAdd.Trim
                        ExcelSheet.Range("I19:X19").Merge()
                        ExcelSheet.Range("I19:X19").Value = pFWDName.Trim
                        ExcelSheet.Range("I20:X22").Merge()
                        ExcelSheet.Range("I20:X22").Value = pFWDAdd.Trim
                        ExcelSheet.Range("I28:P28").Merge()
                        ExcelSheet.Range("I28:P28").Value = pSuratJalanNo.Trim
                        ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(pAttn) & "  TELP : " & Trim(pTelp)

                        ExcelSheet.Range("AE17:AE17").Value = pCommercial.Trim

                        If pPONo.Trim <> pOrderNo.Trim Then
                            ExcelSheet.Range("AE13:AE13").Value = pOrderNo
                            ExcelSheet.Range("AE15:AE15").Value = pPONo
                        Else
                            ExcelSheet.Range("AE13:AE13").Value = pPONo
                            ExcelSheet.Range("AE15:AE15").Value = pOrderNo
                        End If


                        ExcelSheet.Range("AE11:AE11").Value = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM-dd")

                        ExcelSheet.Range("AP11:AT11").Merge()
                        ExcelSheet.Range("AP11:AT11").Value = pETDVendor
                        ExcelSheet.Range("AP13:AT13").Merge()
                        ExcelSheet.Range("AP13:AT13").Value = pETDPort
                        ExcelSheet.Range("AP15:AT15").Merge()
                        ExcelSheet.Range("AP15:AT15").Value = pETAPort
                        ExcelSheet.Range("AP17:AT17").Merge()
                        ExcelSheet.Range("AP17:AT17").Value = pETAFactory

                        ExcelSheet.Range("AE19:AT19").Merge()
                        ExcelSheet.Range("AE19:AT19").Value = pConsigneeName.Trim

                        ExcelSheet.Range("AE20:AT20").Merge()
                        ExcelSheet.Range("AE20:AT20").Value = pConsigneeAdd.Trim
                        k = 0

                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            k = k
                            ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                            ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                            ExcelSheet.Range("i" & k + 34 & ": Q" & k + 34).Merge()
                            ExcelSheet.Range("R" & k + 34 & ": U" & k + 34).Merge()
                            ExcelSheet.Range("V" & k + 34 & ": Y" & k + 34).Merge()
                            ExcelSheet.Range("Z" & k + 34 & ": AA" & k + 34).Merge()
                            ExcelSheet.Range("AB" & k + 34 & ": AC" & k + 34).Merge()
                            ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).Merge()
                            ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).Merge()
                            ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).Merge()
                            ExcelSheet.Range("AP" & k + 34 & ": AS" & k + 34).Merge()

                            ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                            ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partno")).Trim
                            ExcelSheet.Range("i" & k + 34 & ": Q" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("partname"))
                            ExcelSheet.Range("R" & k + 34 & ": U" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo1") & "").Trim
                            ExcelSheet.Range("V" & k + 34 & ": Y" & k + 34).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("labelNo2") & "").Trim
                            ExcelSheet.Range("R" & k + 34).Interior.Color = Color.Yellow
                            ExcelSheet.Range("V" & k + 34).Interior.Color = Color.Yellow
                            ExcelSheet.Range("Z" & k + 34 & ": AA" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("uom")
                            ExcelSheet.Range("AB" & k + 34 & ": AC" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("qtybox")
                            ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("qty")
                            ExcelSheet.Range("AD" & k + 34 & ": AG" & k + 34).NumberFormat = "#,##0"

                            ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("DeliveryQty")
                            ExcelSheet.Range("AH" & k + 34 & ": AK" & k + 34).NumberFormat = "#,##0"
                            ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("BoxQty")
                            ExcelSheet.Range("AL" & k + 34 & ": AO" & k + 34).NumberFormat = "#,##0"
                            ExcelSheet.Range("AP" & k + 34 & ": AS" & k + 34).Value = 0
                            ExcelSheet.Range("AP" & k + 34).Font.Color = Color.Black

                            k = k + 1
                        Next
                        ExcelSheet.Range("B35").Interior.Color = Color.White
                        ExcelSheet.Range("B35").Font.Color = Color.Black
                        ExcelSheet.Range("B" & k + 34).Value = "E"
                        ExcelSheet.Range("B" & k + 34).Interior.Color = Color.Black
                        ExcelSheet.Range("B" & k + 34).Font.Color = Color.White

                        ExcelSheet.Range("AL34" & ": AS" & k + 33).Interior.Color = Color.Yellow
                        clsGeneral.DrawAllBorders(ExcelSheet.Range("B34" & ": AS" & k + 33))

                        xlApp.DisplayAlerts = False

                        'If pPONo.Trim = pOrderNo.Trim Then
                        '    ExcelBook.SaveAs(Trim(txtSaveAs.Text) & "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsm")
                        '    pFileName1 = "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & ".xlsm"
                        'Else
                        '    ExcelBook.SaveAs(Trim(txtSaveAs.Text) & "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNo1) & ")" & ".xlsm")
                        '    pFileName1 = "\GOOD RECEIVING-" & Trim(ls_Aff) & "-" & Trim(ls_Supplier) & "-" & Trim(ls_orderNo) & " Split (" & Trim(ls_orderNo1) & ")" & ".xlsm"
                        'End If

                        If pFileName1 = "" Then
                            If pPONo.Trim = pOrderNo.Trim Then
                                ExcelBook.SaveAs(pResult & "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & ".xlsm")
                                pFileName1 = "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & ".xlsm"
                            Else
                                ExcelBook.SaveAs(pResult & "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & " Split (" & Trim(pOrderNo) & ")" & ".xlsm")
                                pFileName1 = "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & " Split (" & Trim(pOrderNo) & ")" & ".xlsm"
                            End If
                        Else
                            If pPONo.Trim = pOrderNo.Trim Then
                                ExcelBook.SaveAs(pResult & "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & ".xlsm")
                                pFileName2 = "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & ".xlsm"
                            Else
                                ExcelBook.SaveAs(pResult & "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & " Split (" & Trim(pOrderNo) & ")" & ".xlsm")
                                pFileName2 = "\GOOD RECEIVING-" & Trim(pAffCode) & "-" & Trim(pSupplier) & "-" & Trim(pPONo) & " Split (" & Trim(pOrderNo) & ")" & ".xlsm"
                            End If
                        End If

                        xlApp.Workbooks.Close()
                        xlApp.Quit()

                        If booSplit Then GoTo Split

                        '---------------------------------------excel---------------------------------------'
                        If sendEmailtoForwarder(GB, pResult, pFileName1, pFileName2, pSuratJalanNo, pPONo, pOrderNo, pAffCode, pSupplier, pFWD, errMsg) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email Receiving to Forwarder. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok.")
                        End If

                        Call UpdateStatusDOExport(pAffAdd, pSupplier, pPONo, pSuratJalanNo, errMsg)
                    End If
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "Receiving PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] " & ex.Message
            ErrSummary = "Receiving PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] " & ex.Message
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
        End Try
    End Sub

    Public Shared Sub UpdateStatusDOExport(ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pSJ As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update dbo.DOSupplier_Master_Export set excelcls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND SuratJalanno = '" & pSJ & "'" & vbCrLf & _
                         " AND PONo = '" & pPoNo & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            errMsg = "Process Send PONo [" & pPoNo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub

    Public Shared Function BindDataDeliveryInstruction(ByVal GB As GlobalSetting.clsGlobal, ByVal pSJ As String, ByVal PONO As String, ByVal ls_orderNo As String, ByVal Aff As String, ByVal Supp As String) As DataSet
        Dim ls_sql As String

        ls_sql = "  select  distinct  " & vbCrLf & _
                  "  orderno = POD.PONo,  " & vbCrLf & _
                  "  Partno = POD.PartNo,  " & vbCrLf & _
                  "  Partname = MP.PartName,  " & vbCrLf & _
                  "  labelno1 = Rtrim(PL1.LabelNo),  " & vbCrLf & _
                  "  labelno2 = Rtrim(PL2.LabelNo),  " & vbCrLf & _
                  "  uom = MU.Description,  " & vbCrLf & _
                  "  qtybox = QtyBox,  " & vbCrLf & _
                  "  qty = convert(char,DOD.DOQty),  " & vbCrLf & _
                  "  remaining = convert(char,POD.Week1),  " & vbCrLf & _
                  "  DeliveryQty = convert(char,DOD.DOQty),  " & vbCrLf & _
                  "  boxqty = Ceiling(DOD.DOQty / QtyBox),  weight = NetWeight,  "

        ls_sql = ls_sql + "  barcode = convert(char(25),'') + Convert(char(20),POD.AFfiliateID) + convert(char(20), POD.Pono) + Convert(char(25), POD.PartNo) +  convert(char,DOD.DOQty)  " & vbCrLf & _
                          "  From PO_Detail_Export POD  " & vbCrLf & _
                          "  INNER JOIN PO_Master_Export POM  " & vbCrLf & _
                          "  ON POM.Pono = POD.PONo and POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
                          "  and POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  And POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                          "  LEFT JOIN DOSupplier_Master_Export DOM ON DOM.SupplierID = POM.SupplierID and DOM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "  AND DOM.PONo = POM.PONo and DOM.ORderNo = POM.OrderNo1  " & vbCrLf & _
                          "  INNER JOIN DOSupplier_Detail_Export DOD ON DOD.SuratJalanNo = DOM.SuratJalanNo and DOD.AffiliateID = DOM.AffiliateID  " & vbCrLf & _
                          " and DOD.SupplierID = DOM.SupplierID and DOD.PONo = DOM.POno and DOD.OrderNo = DOM.OrderNo and DOD.PartNo = POD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN (select OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, min(BoxNo) as labelno, SeqNo from DOSupplier_DetailBox_Export group by OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, SeqNo)PL1  "

        ls_sql = ls_sql + "  ON PL1.PONo = DOD.PONo  " & vbCrLf & _
                          "  and PL1.AffiliateID = DOD.AffiliateID  " & vbCrLf & _
                          "  AND PL1.SupplierID = DOD.SupplierID  " & vbCrLf & _
                          "  AND PL1.PartNO = DOD.PartNo  " & vbCrLf & _
                          "  AND PL1.SuratJalanno = DOD.SuratJalanno " & vbCrLf & _
                          "  AND PL1.OrderNo = DOD.OrderNo " & vbCrLf & _
                          "  AND PL1.SeqNo = DOD.SeqNo " & vbCrLf & _
                          "  LEFT JOIN (select OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, Max(BoxNo) as labelno, SeqNo from DOSupplier_DetailBox_Export group by OrderNo,SuratJalanno,POno, AffiliateID, SupplierID, PartNo, SeqNo)PL2  " & vbCrLf & _
                          "  ON PL2.PONo = DOD.PONo  " & vbCrLf & _
                          "  and PL2.AffiliateID = DOD.AffiliateID  " & vbCrLf & _
                          "  AND PL2.SupplierID = DOD.SupplierID  " & vbCrLf & _
                          "  AND PL2.PartNO = DOD.PartNo  " & vbCrLf & _
                          "  AND PL2.SuratJalanno = DOD.SuratJalanno " & vbCrLf & _
                          "  AND PL2.OrderNo = DOD.OrderNo " & vbCrLf & _
                          "  AND PL2.SeqNo = DOD.SeqNo " & vbCrLf

        ls_sql = ls_sql + "  INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo  " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM On MPM.PartNo = POD.PartNo  " & vbCrLf & _
                          "  AND MPM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  AND MPM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "  INNER JOIN ms_unitcls MU ON MU.UnitCls = MP.Unitcls  " & vbCrLf & _
                          " Where (Pl1.labelno is not null and pl2.labelno is not null) and DOM.SuratJalanno = '" & Trim(pSJ) & "'" & vbCrLf & _
                          " AND POD.OrderNo1 = '" & Trim(ls_orderNo) & "'" & vbCrLf & _
                          " AND POD.AffiliateID = '" & Trim(Aff) & "'" & vbCrLf & _
                          " AND POD.SupplierID = '" & Trim(Supp) & "'"

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function sendEmailtoForwarder(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pFileName2 As String, ByVal pSJ As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pFWD As String, ByRef errMsg As String) As Boolean
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_Attachment2 As String = ""
            Dim ls_URl As String = ""

            sendEmailtoForwarder = True

            Dim dsEmail As New DataSet
            dsEmail = clsGeneral.getEmailAddressExport(GB, "", "PASI", "", pFWD, "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

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
                errMsg = "Process Send PO Export [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoForwarder = False
                errMsg = "Process Send PO Export [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If pPONo.Trim <> pOrderNo.Trim Then
                ls_Subject = "DN-" & Trim(pSupplier) & "-" & Trim(pPONo) & " Split (" & pOrderNo & ")" & " Delivery Note From Supplier [TRIAL]"
            Else
                ls_Subject = "DN-" & Trim(pSupplier) & "-" & Trim(pPONo) & " Delivery Note From Supplier [TRIAL]"
            End If

            If Trim(pPONo) <> Trim(pOrderNo) Then
                ls_Body = clsNotification.GetNotification("20", "", pPONo.Trim & " Split (" & pOrderNo.Trim & ")", , pSj.Trim)
            Else
                ls_Body = clsNotification.GetNotification("20", "", pPONo.Trim, , pSj.Trim)
            End If

            ls_Attachment = Trim(pPathFile) & pFileName
            ls_Attachment2 = Trim(pPathFile) & pFileName2

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment, ls_Attachment2, , , , , True) = False Then
                sendEmailtoForwarder = False
                Exit Function
            End If

            sendEmailtoForwarder = True

        Catch ex As Exception
            sendEmailtoForwarder = False
            errMsg = "Process Send PO Export [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

End Class
