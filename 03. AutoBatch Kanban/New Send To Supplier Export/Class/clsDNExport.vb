Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsDNExport
    Shared Sub up_DNExport(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim ls_newLabelEx As Boolean = True 'untuk format label baru export

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp = New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim i As Integer, k As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim fromEmail As String = ""
        Dim NewFileCopy As String

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

        Dim ds As New DataSet
        Dim dsDetailDelivery As New DataSet
        Dim dsEmail As New DataSet

        Dim pFileName1 As String = ""
        Dim pFileName2 As String = ""

        Try
            ls_SQL = " select distinct attn = isnull(MF.Attn,''), telp = isnull(MF.MobilePhone,''), PME.Period, PME.PONo, PME.OrderNo1 as orderNo, PME.AffiliateID, PME.SupplierID,MS.SupplierName,isnull(MS.Address,'') SUPPAddress,  " & vbCrLf & _
                     " PME.ForwarderID, MF.ForwarderName, isnull(MF.Address,'')  FWDAddress, ETDVendor1 as ETDVendor, ETDPort1 as ETDPort, ETAPort1 as ETAPort, ETAFactory1 as ETAFactory, " & vbCrLf & _
                     " PME.AffiliateID as AFF, MA.AffiliateName as AFFName, isnull(MA.Address,'') AFFAddress from  " & vbCrLf & _
                     " PO_Master_Export PME  " & vbCrLf & _
                     " LEFT JOIN MS_Supplier MS ON MS.SupplierID = PME.SupplierID " & vbCrLf & _
                     " LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = PME.ForwarderID " & vbCrLf & _
                     " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PME.AffiliateID " & vbCrLf & _
                     " where isnull(FinalApprovalCls,0) =1 and isnull(PONO,'') <> ''"

            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For i_loop = 0 To ds.Tables(0).Rows.Count - 1
                    'Create Excel File
                    Dim fi As New FileInfo(pAtttacment & "\Template Customer Delivery Confirmation.xlsm")
                    If Not fi.Exists Then
                        errMsg = "Process Send Customer Delivery Confirmation to Supplier STOPPED, File Excel isn't Found"
                        ErrSummary = "Process Send Customer Delivery Confirmation to Supplier STOPPED, File Excel isn't Found"
                        Exit Try
                    End If

                    pPONo = Trim(ds.Tables(0).Rows(i_loop)("PONo"))
                    pOrderNo = Trim(ds.Tables(0).Rows(i_loop)("OrderNo"))

                    pSupplier = Trim(ds.Tables(0).Rows(i_loop)("supplierID"))
                    pAffCode = Trim(ds.Tables(0).Rows(i_loop)("AFF"))
                    pFWD = Trim(ds.Tables(0).Rows(i_loop)("ForwarderID"))

                    pSupplierName = Trim(ds.Tables(0).Rows(i_loop)("suppliername"))
                    pSupplierAdd = Trim(ds.Tables(0).Rows(i_loop)("SuppAddress"))

                    pFWDName = Trim(ds.Tables(0).Rows(i_loop)("ForwarderName"))
                    pFWDAdd = Trim(ds.Tables(0).Rows(i_loop)("FWDAddress"))

                    pETDVendor = Format((ds.Tables(0).Rows(i_loop)("ETDVendor")), "yyyy-MM-dd")

                    pAffName = Trim(ds.Tables(0).Rows(i_loop)("AFFName"))
                    pAffAdd = Trim(ds.Tables(0).Rows(i_loop)("AFFAddress"))
                    pPeriod = Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy-MM")

                    pAttn = Trim(ds.Tables(0).Rows(i_loop)("attn"))
                    pTelp = Trim(ds.Tables(0).Rows(i_loop)("telp"))

                    Call InsertPrintLabel(GB, pPONo, pOrderNo, pAffCode, pSupplier, Format((ds.Tables(0).Rows(i_loop)("Period")), "yyyy"), Format((ds.Tables(0).Rows(i_loop)("Period")), "MM"), errMsg)

                    dsDetailDelivery = BidDataDeliveryConfirm(GB, pPONo, pOrderNo, pAffCode, pSupplier)

                    If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                        dsEmail = clsGeneral.getEmailAddressPASI(GB, "", "PASI", "", "", "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)
                        For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                            If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                                fromEmail = dsEmail.Tables(0).Rows(i)("EmailFrom")
                                receiptCCEmail = dsEmail.Tables(0).Rows(i)("EmailCC")
                            End If
                        Next

                        NewFileCopy = pAtttacment & "\Template Customer Delivery Confirmation.xlsm"
                        ExcelBook = xlApp.Workbooks.Open(NewFileCopy)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        ExcelSheet.Range("H2").Value = fromEmail
                        ExcelSheet.Range("Y2").Value = receiptCCEmail
                        ExcelSheet.Range("H3").Value = pAffCode
                        ExcelSheet.Range("H4").Value = pFWD
                        ExcelSheet.Range("H5").Value = pSupplier

                        ExcelSheet.Range("I11:X11").Value = pSupplierName
                        ExcelSheet.Range("I12:X15").Value = pSupplierAdd

                        ExcelSheet.Range("I19:X19").Value = pFWDName
                        ExcelSheet.Range("I20:X22").Value = pFWDAdd
                        ExcelSheet.Range("I23:X23").Value = "ATTN : " & Trim(pAttn) & "   TELP : " & Trim(pTelp)

                        ExcelSheet.Range("AE19:AT19").Value = pAffName
                        ExcelSheet.Range("AE20:AT22").Value = pAffAdd

                        ExcelSheet.Range("AE11:AI11").Value = pPeriod
                        ExcelSheet.Range("AE13:AI13").Value = pPONo

                        If pPONo <> pOrderNo Then
                            ExcelSheet.Range("AE15:AI15").Value = pOrderNo
                        End If

                        ExcelSheet.Range("AE17:AI17").Value = pETDVendor

                        k = 0
                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Merge()
                            ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Merge()
                            ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Merge()
                            ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Merge()
                            ExcelSheet.Range("W" & k + 34 & ": Z" & k + 34).Merge()
                            ExcelSheet.Range("AA" & k + 34 & ": AB" & k + 34).Merge()
                            ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Merge()
                            ExcelSheet.Range("AE" & k + 34 & ": AH" & k + 34).Merge()
                            ExcelSheet.Range("AI" & k + 34 & ": AL" & k + 34).Merge()
                            ExcelSheet.Range("AM" & k + 34 & ": AP" & k + 34).Merge()
                            ExcelSheet.Range("AQ" & k + 34 & ": AT" & k + 34).Merge()

                            ExcelSheet.Range("B" & k + 34 & ": C" & k + 34).Value = k + 1
                            ExcelSheet.Range("D" & k + 34 & ": H" & k + 34).Value = pPONo
                            ExcelSheet.Range("I" & k + 34 & ": M" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("Partno")
                            ExcelSheet.Range("N" & k + 34 & ": V" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                            ExcelSheet.Range("W" & k + 34 & ": Z" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("labelno")
                            ExcelSheet.Range("AA" & k + 34 & ": AB" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                            ExcelSheet.Range("AC" & k + 34 & ": AD" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("MOQ")

                            ExcelSheet.Range("AE" & k + 34 & ": AH" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                            ExcelSheet.Range("AI" & k + 34 & ": AL" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")

                            ExcelSheet.Range("AM" & k + 34 & ": AP" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("SuppQty")
                            ExcelSheet.Range("AQ" & k + 34 & ": AT" & k + 34).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalPOQty")

                            k = k + 1
                        Next

                        ExcelSheet.Range("B35").Interior.Color = Color.White
                        ExcelSheet.Range("B35").Font.Color = Color.Black
                        ExcelSheet.Range("B" & k + 34).Value = "E"
                        ExcelSheet.Range("B" & k + 34).Interior.Color = Color.Black
                        ExcelSheet.Range("B" & k + 34).Font.Color = Color.White

                        k = k - 1
                        clsGeneral.DrawAllBorders(ExcelSheet.Range("B34" & ": AT" & k + 34))
                        ExcelSheet.Range("AM34" & ": AP" & k + 34).Interior.Color = ColorYellow

                        'Save ke Local
                        xlApp.DisplayAlerts = False

                        ExcelBook.SaveAs(pResult & "\Delivery Confirmation-" & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm")
                        'pFileName1 = "\Delivery Confirmation-" & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                        If pPONo <> pOrderNo Then
                            ExcelBook.SaveAs(pResult & "\Delivery Confirmation-" & Trim(pPONo) & " Split (" & Trim(pOrderNo) & ")-" & Trim(pSupplier) & ".xlsm")
                            pFileName1 = "\Delivery Confirmation-" & Trim(pPONo) & " Split (" & Trim(pOrderNo) & ")-" & Trim(pSupplier) & ".xlsm"
                        Else
                            ExcelBook.SaveAs(pResult & "\Delivery Confirmation-" & Trim(pOrderNo) & "-" & Trim(pSupplier) & ".xlsm")
                            pFileName1 = "\Delivery Confirmation-" & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                        End If

                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                    End If

                    '================================DELIVERY CONFIRMATION==================================

                    '=======================================LABEL===========================================

                    Dim fi2 As New FileInfo(pAtttacment & "\Print label2.xlsm")

                    If Not fi2.Exists Then
                        errMsg = "Process Send Customer Delivery Confirmation to Supplier STOPPED, File Excel isn't Found"
                        ErrSummary = "Process Send Customer Delivery Confirmation to Supplier STOPPED, File Excel isn't Found"
                        Exit Try
                    End If

                    dsDetailDelivery = BindDataLabelPrint(GB, pPONo, pOrderNo, pAffAdd, pSupplier)

                    If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                        NewFileCopy = pAtttacment & "\Print Label2.xlsm"
                        ExcelBook = xlApp.Workbooks.Open(NewFileCopy)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        ExcelSheet.Range("I4:W4").Value = pAffName
                        ExcelSheet.Range("I6:U6").Value = pFWDName
                        ExcelSheet.Range("I7:U9").Value = pFWDAdd
                        ExcelSheet.Range("AC6:AQ6").Value = pSupplierName
                        ExcelSheet.Range("AC7:AQ9").Value = pSupplierAdd

                        ExcelSheet.Range("I12:P12").Value = pPeriod
                        ExcelSheet.Range("I14:P14").Value = Format(pETDVendor, "yyyy-MM-dd")

                        k = 0

                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            k = k
                            ExcelSheet.Range("B" & k + 20 & ": C" & k + 20).Merge()
                            ExcelSheet.Range("D" & k + 20 & ": F" & k + 20).Merge()
                            ExcelSheet.Range("G" & k + 20 & ": I" & k + 20).Merge()
                            ExcelSheet.Range("J" & k + 20 & ": K" & k + 20).Merge()
                            ExcelSheet.Range("L" & k + 20 & ": P" & k + 20).Merge()
                            ExcelSheet.Range("Q" & k + 20 & ": X" & k + 20).Merge()
                            ExcelSheet.Range("Y" & k + 20 & ": AD" & k + 20).Merge()
                            ExcelSheet.Range("AE" & k + 20 & ": AI" & k + 20).Merge()
                            ExcelSheet.Range("AJ" & k + 20 & ": AK" & k + 20).Merge()
                            ExcelSheet.Range("AL" & k + 20 & ": AM" & k + 20).Merge()
                            ExcelSheet.Range("AN" & k + 20 & ": AQ" & k + 20).Merge()
                            ExcelSheet.Range("AR" & k + 20 & ": AT" & k + 20).Merge()

                            ExcelSheet.Range("J" & k + 20 & ": K" & k + 20).Value = k + 1
                            ExcelSheet.Range("D" & k + 20 & ": F" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("label1")
                            ExcelSheet.Range("G" & k + 20 & ": I" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("label2")
                            ExcelSheet.Range("L" & k + 20 & ": P" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("Partno")
                            ExcelSheet.Range("Q" & k + 20 & ": X" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("Partname")
                            ExcelSheet.Range("Y" & k + 20 & ": AD" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("labelNo")
                            ExcelSheet.Range("AE" & k + 20 & ": AI" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("OrderNo")
                            ExcelSheet.Range("AJ" & k + 20 & ": AK" & k + 20).Value = Trim(dsDetailDelivery.Tables(0).Rows(j)("Uom"))

                            ExcelSheet.Range("AL" & k + 20 & ": AM" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("QtyBox")
                            ExcelSheet.Range("AN" & k + 20 & ": AQ" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("Qty")
                            ExcelSheet.Range("AR" & k + 20 & ": AT" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("boxqty")
                            ExcelSheet.Range("AU" & k + 20 & ": AU" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("DestinationPort")
                            ExcelSheet.Range("AV" & k + 20 & ": AV" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("DestinationPoint")
                            ExcelSheet.Range("AW" & k + 20 & ": AW" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("custname")
                            ExcelSheet.Range("AX" & k + 20 & ": AX" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("custcode")
                            ExcelSheet.Range("AY" & k + 20 & ": AY" & k + 20).Value = dsDetailDelivery.Tables(0).Rows(j)("consigneecode")


                            k = k + 1
                        Next
                        ExcelSheet.Range("B21").Interior.Color = Color.White
                        ExcelSheet.Range("B21" & ": I" & k + 19).Interior.Color = ColorYellow
                        ExcelSheet.Range("AN20" & ": AT" & k + 20).Font.Color = Color.Black
                        ExcelSheet.Range("AU20" & ": AY" & k + 20).Font.Color = Color.White

                        ExcelSheet.Range("D20" & ": J" & k + 20).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        ExcelSheet.Range("AJ20" & ": AL" & k + 20).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                        ExcelSheet.Range("B21").Value = ""
                        ExcelSheet.Range("B21").Font.Color = Color.Black

                        ExcelSheet.Range("B" & k + 20).Value = "E"
                        ExcelSheet.Range("B" & k + 20).Interior.Color = Color.Black
                        ExcelSheet.Range("B" & k + 20).Font.Color = Color.White

                        clsGeneral.DrawAllBorders(ExcelSheet.Range("B20" & ": AT" & k + 19))
                        xlApp.DisplayAlerts = False

                        ExcelBook.SaveAs(pResult & "\Print Label-" & Trim(pOrderNo) & "-" & Trim(pSupplier) & ".xlsm")
                        pFileName2 = "\Print Label-" & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"

                        xlApp.Workbooks.Close()
                        xlApp.Quit()

                        If sendEmailtoSupplier(GB, pResult, pFileName1, pFileName2, pPONo, pAffCode, pSupplier, errMsg) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email DN to Supplier. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok.")
                        End If
                    End If
                    '=======================================LABEL===========================================

                    '==================================ORDER CONFIRMATION===================================
                    dsDetailDelivery = BindDataOrderConfirmation(GB, pPONo, pOrderNo, pAffCode, pSupplier)

                    If dsDetailDelivery.Tables(0).Rows.Count > 0 Then
                        NewFileCopy = pAtttacment & "\Template Customer Order Confirmation.xlsx"
                        ExcelBook = xlApp.Workbooks.Open(NewFileCopy)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        ExcelSheet.Range("A2:BA2").Value = IIf(dsDetailDelivery.Tables(0).Rows(0)("EmergencyCls") = "M", "PASI FINAL APPROVAL PO (MONTHLY)", "PASI FINAL APPROVAL PO (EMERGENCY)")
                        ExcelSheet.Range("G6:K6").Value = dsDetailDelivery.Tables(0).Rows(0)("period")                        
                        ExcelSheet.Range("G8:K8").Value = dsDetailDelivery.Tables(0).Rows(0)("Pono")                        
                        ExcelSheet.Range("G10:K10").Value = IIf(dsDetailDelivery.Tables(0).Rows(0)("OrderNo") <> dsDetailDelivery.Tables(0).Rows(0)("poNO"), dsDetailDelivery.Tables(0).Rows(0)("OrderNo"), "-")                        
                        ExcelSheet.Range("G12:K12").Value = IIf(dsDetailDelivery.Tables(0).Rows(0)("ShipBy") = "B", "BOAT", "AIR")

                        ExcelSheet.Range("R6:V6").Value = dsDetailDelivery.Tables(0).Rows(0)("ETDVENDOR")
                        ExcelSheet.Range("R8:V8").Value = dsDetailDelivery.Tables(0).Rows(0)("ETDPORT")
                        ExcelSheet.Range("R10:V10").Value = dsDetailDelivery.Tables(0).Rows(0)("ETAPORT")
                        ExcelSheet.Range("R12:V12").Value = dsDetailDelivery.Tables(0).Rows(0)("ETAFACTORY")

                        ExcelSheet.Range("G14:V14").Value = Trim(dsDetailDelivery.Tables(0).Rows(0)("AFFCode")) & "-" & dsDetailDelivery.Tables(0).Rows(0)("AFFName")
                        ExcelSheet.Range("G15:V17").Value = dsDetailDelivery.Tables(0).Rows(0)("AFFAdd")

                        ExcelSheet.Range("AE6:AT6").Value = dsDetailDelivery.Tables(0).Rows(0)("SuppName")
                        ExcelSheet.Range("AE7:AT10").Value = dsDetailDelivery.Tables(0).Rows(0)("SuppAdd")

                        ExcelSheet.Range("AE14:AT14").Value = dsDetailDelivery.Tables(0).Rows(0)("FWDName")
                        ExcelSheet.Range("AE15:AT17").Value = dsDetailDelivery.Tables(0).Rows(0)("FWDAdd")

                        k = 0

                        For j = 0 To dsDetailDelivery.Tables(0).Rows.Count - 1
                            k = k
                            ExcelSheet.Range("B" & k + 22 & ": C" & k + 22).Merge()
                            ExcelSheet.Range("D" & k + 22 & ": H" & k + 22).Merge()
                            ExcelSheet.Range("i" & k + 22 & ": P" & k + 22).Merge()

                            ExcelSheet.Range("Q" & k + 22 & ": R" & k + 22).Merge()
                            ExcelSheet.Range("S" & k + 22 & ": T" & k + 22).Merge()
                            ExcelSheet.Range("U" & k + 22 & ": X" & k + 22).Merge()

                            ExcelSheet.Range("Y" & k + 22 & ": AB" & k + 22).Merge()
                            ExcelSheet.Range("AC" & k + 22 & ": AF" & k + 22).Merge()
                            ExcelSheet.Range("AG" & k + 22 & ": AJ" & k + 22).Merge()
                            ExcelSheet.Range("AK" & k + 22 & ": AN" & k + 22).Merge()

                            ExcelSheet.Range("B" & k + 22 & ": C" & k + 22).Value = k + 1
                            ExcelSheet.Range("D" & k + 22 & ": H" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("Partno")
                            ExcelSheet.Range("i" & k + 22 & ": P" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("PartName")
                            ExcelSheet.Range("Q" & k + 22 & ": R" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("UOM")
                            ExcelSheet.Range("S" & k + 22 & ": T" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("QtyBox")
                            ExcelSheet.Range("U" & k + 22 & ": X" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("Qty")
                            ExcelSheet.Range("Y" & k + 22 & ": AB" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("SupplierQty")
                            ExcelSheet.Range("AC" & k + 22 & ": AF" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalBox")
                            ExcelSheet.Range("AG" & k + 22 & ": AJ" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("TotalNet")
                            ExcelSheet.Range("AK" & k + 22 & ": AN" & k + 22).Value = dsDetailDelivery.Tables(0).Rows(j)("Volume")

                            k = k + 1
                        Next

                        clsGeneral.DrawAllBorders(ExcelSheet.Range("B22" & ": AN" & k + 21))

                        xlApp.DisplayAlerts = False

                        ExcelBook.SaveAs(pResult & "\PASI FINAL APPROVAL-" & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsx")
                        pFileName1 = "\PASI FINAL APPROVAL-" & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsx"

                        xlApp.Workbooks.Close()
                        xlApp.Quit()

                        If sendEmailtoForwarder(GB, pResult, pFileName1, pPONo, pAffCode, pSupplier, pFWD, errMsg) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email DN to Forwarder. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok.")
                        End If
                    End If
                    '==================================ORDER CONFIRMATION===================================

                    Call UpdateStatusPOExport(GB, pAffCode, pSupplier, pPONo, pOrderNo, errMsg)
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "DN PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] " & ex.Message
            ErrSummary = "DN PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] " & ex.Message
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

    Public Shared Sub InsertPrintLabel(ByVal GB As GlobalSetting.clsGlobal, ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pTahun As Integer, ByVal pBulan As Integer, Optional ByRef pErrMsg As String = "")
        Dim ls_SQL As String = ""
        Dim admin As String = "administrator"
        Dim x As Integer = 0
        Dim ls_Startno As Integer
        Dim LabelNo As String = ""

        Dim dsData As New DataSet
        Dim dsSeqNo As New DataSet
        Dim dsAda As New DataSet

        Dim cfg As New GlobalSetting.clsConfig
        Dim ls_Period As String = ""
        Dim ls_newLabelEx As Boolean = True

        Dim ls_awalT As Integer = 2016
        Dim ls_selisihT As Integer = 0
        Dim ls_charT As Integer = 65
        Dim ls_codeT As String = ""
        Dim ls_awalB As Integer = 1
        Dim ls_selisihB As Integer = 0
        Dim ls_charB As Integer = 65
        Dim ls_codeB As String = ""
        'Dim ls_Period As String = ""

        Try
            dsAda = SelectInsertBarcodeExport(GB, pPono, pOrderNo, pAff, pSupp)
            If dsAda.Tables(0).Rows.Count = 0 Then
                dsData = InsertBarcodeExport(GB, pPono, pOrderNo, pAff, pSupp)
                If dsData.Tables(0).Rows.Count > 0 Then
                    ls_Period = Format(dsData.Tables(0).Rows(0)("Period"), "yyyy-MM-dd")
                    Using sqlConn As New SqlConnection(cfg.ConnectionString)
                        sqlConn.Open()
                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CreateKanban")
                            Dim sqlCommNew As SqlCommand = sqlConn.CreateCommand
                            sqlCommNew.Connection = sqlConn
                            sqlCommNew.Transaction = sqlTran

                            '============== Cari code =============
                            If ls_newLabelEx = True Then
                                'TAHUN
                                ls_selisihT = pTahun - ls_awalT
                                If ls_selisihT = 0 Then
                                    ls_codeT = Chr(ls_charT)
                                Else
                                    If (ls_charT + ls_selisihT) >= 73 Then 'I
                                        ls_selisihT = ls_selisihT + 1

                                        If ls_selisihT >= 79 Then 'O
                                            ls_selisihT = ls_selisihT + 1

                                            If ls_selisihT >= 83 Then 'S
                                                ls_selisihT = ls_selisihT + 1
                                            End If
                                        End If
                                    End If
                                    ls_codeT = Chr(ls_charT + ls_selisihT)
                                End If
                                'BULAN
                                ls_selisihB = pBulan - ls_awalB
                                If ls_selisihB = 0 Then
                                    ls_codeB = Chr(ls_charB)
                                Else
                                    If (ls_charB + ls_selisihB) >= 73 Then 'I
                                        ls_selisihB = ls_selisihB + 1

                                        If ls_selisihB >= 79 Then 'O
                                            ls_selisihB = ls_selisihB + 1

                                            If ls_selisihB >= 83 Then 'S
                                                ls_selisihB = ls_selisihB + 1
                                            End If
                                        End If
                                    End If
                                    ls_codeB = Chr(ls_charB + ls_selisihB)
                                End If
                            End If
                            '============== Cari code =============

                            For i = 0 To dsData.Tables(0).Rows.Count - 1
                                If ls_newLabelEx = False Then
                                    dsSeqNo = GetLABELNO(GB, pPono, pOrderNo, pAff, pSupp, dsData.Tables(0).Rows(i)("PartNo"))
                                Else
                                    dsSeqNo = GetLABELNONew(GB, ls_codeT + ls_codeB, ls_Period)
                                End If
                                If dsSeqNo.Tables(0).Rows.Count > 0 Then
                                    ls_Startno = dsSeqNo.Tables(0).Rows(0)("seqno")
                                Else
                                    ls_Startno = 0
                                End If
                                For x = 0 To dsData.Tables(0).Rows(i)("looping") - 1
                                    LabelNo = "00000" & ls_Startno + 1

                                    '------ NEW LABEL ------
                                    If ls_newLabelEx = False Then
                                        LabelNo = Trim(dsData.Tables(0).Rows(i)("LabelCode")) + Microsoft.VisualBasic.Right(LabelNo, 5)
                                    Else
                                        LabelNo = ls_codeT + ls_codeB + Microsoft.VisualBasic.Right(LabelNo, 5)
                                    End If
                                    '------ NEW LABEL ------

                                    ls_SQL = "INSERT INTO PrintLabelExport " & vbCrLf & _
                                             " VALUES ( " & vbCrLf & _
                                             " '" & pPono & "', " & vbCrLf & _
                                             " '" & pAff & "', " & vbCrLf & _
                                             " '" & pSupp & "', " & vbCrLf & _
                                             " '" & dsData.Tables(0).Rows(i)("PartNo") & "', " & vbCrLf & _
                                             " '" & LabelNo & "', " & vbCrLf & _
                                             " getdate(), " & vbCrLf & _
                                             " 'AdminDNX', " & vbCrLf & _
                                             " '" & pOrderNo & "','' )"
                                    sqlCommNew = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                    sqlCommNew.ExecuteNonQuery()
                                    ls_Startno = ls_Startno + 1
                                Next
                            Next

                            sqlCommNew.Dispose()
                            sqlTran.Commit()
                        End Using
                        sqlConn.Close()
                    End Using
                End If
            End If
        Catch ex As Exception
            pErrMsg = ex.Message
        End Try
    End Sub

    Public Shared Function SelectInsertBarcodeExport(ByVal GB As GlobalSetting.clsGlobal, ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String) As DataSet
        Dim ls_sql As String = ""

        ls_sql = " select * from PrintLabelExport" & vbCrLf & _
                 " where PONO = '" & pPono & "' and OrderNo = '" & pOrderNo & "' " & vbCrLf & _
                 " and AffiliateID = '" & pAff & "' and SupplierID = '" & pSupp & "' " & vbCrLf

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function InsertBarcodeExport(ByVal GB As GlobalSetting.clsGlobal, ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String) As DataSet
        Dim ls_sql As String = ""

        ls_sql = " Select distinct POM.Period, POD.PoNo,POD.OrderNo1, POD.AffiliateID, POD.SupplierID, POD.PartNo as PartNo, " & vbCrLf & _
                 " PUD.Week1 as qty, QtyBox, looping = Convert(numeric,PUD.Week1/QtyBox), MSP.LabelCode " & vbCrLf & _
                 " From PO_Detail_Export POD INNER JOIN PO_DetailUpload_Export PUD with(nolock) " & vbCrLf & _
                 " ON POD.PONo = PUD.PONo " & vbCrLf & _
                 " AND POD.AffiliateID = PUD.AffiliateID " & vbCrLf & _
                 " AND POD.SupplierID = PUD.SupplierID " & vbCrLf & _
                 " AND POD.partNO = PUD.PartNo " & vbCrLf & _
                 " INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo " & vbCrLf & _
                 " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo " & vbCrLf & _
                 " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                 " AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                 " LEFT JOIN MS_Supplier MSP ON MSP.SupplierID = POD.SupplierID " & vbCrLf & _
                 " LEFT JOIN PO_Master_Export POM ON POM.PONo = POD.PONo and POM.OrderNo1 = POD.OrderNo1 and POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                 " AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                 " where POD.PONO = '" & pPono & "' and POD.OrderNO1 = '" & pOrderNo & "' " & vbCrLf & _
                 " and POD.AffiliateID = '" & pAff & "' and POD.SupplierID = '" & pSupp & "' " & vbCrLf

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function GetLABELNO(ByVal GB As GlobalSetting.clsGlobal, ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pPartNo As String) As DataSet
        Dim ls_sql As String = ""

        ls_sql = " select seqno = Convert(numeric,replace(max(labelno), left(max(labelno),1),'')) from PrintLabelExport with(nolock) " & vbCrLf & _
                 " where SupplierID = '" & pSupp & "' " & vbCrLf & _
                 " and PartNo = '" & Trim(pPartNo) & "'" & vbCrLf & _
                 " GROUP BY SupplierID, PartNo "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function GetLABELNONew(ByVal GB As GlobalSetting.clsGlobal, ByVal pcode As String, ByVal pperiod As String) As DataSet
        Dim ls_sql As String = ""

        ls_sql = " select seqno = isnull(Convert(numeric,replace(max(labelno), left(max(labelno),2),'')),0) " & vbCrLf & _
                 " from PrintLabelExport PL with(nolock) " & vbCrLf & _
                 " LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                 " ON POM.PONo = PL.PONo and POM.OrderNo1 = PL.OrderNo and POM.AffiliateID = PL.AffiliateID and POM.SupplierID = PL.SupplierID " & vbCrLf & _
                 " where Left(LabelNo,2) = '" & pcode & "' " & vbCrLf & _
                 " and POM.Period = '" & pperiod & "' "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds

    End Function

    Public Shared Function BidDataDeliveryConfirm(ByVal GB As GlobalSetting.clsGlobal, ByVal PONO As String, ByVal OrderNo As String, ByVal affiliateID As String, ByVal SupplierID As String) As DataSet
        Dim ls_sql As String

        ls_sql = " select  " & vbCrLf & _
                  " Partno = POD.PartNo, " & vbCrLf & _
                  " PartName = MP.Partname, " & vbCrLf & _
                  " UOM = MUC.Description, " & vbCrLf & _
                  " MOQ = QtyBox, " & vbCrLf & _
                  " OrderQty = POD.Week1,  " & vbCrLf & _
                  " labelno = Rtrim(PL.Label1) + ' - ' + Rtrim(pl.Label2), " & vbCrLf & _
                  " SuppQty = PUD.week1, TotalPOQty = PUD.Week1 / QtyBox " & vbCrLf & _
                  " From PO_Detail_Export POD LEFT JOIN PO_DetailUpload_Export PUD " & vbCrLf & _
                  " ON POD.PONo = PUD.PONo " & vbCrLf & _
                  " And POD.AffiliateID = PUD.AffiliateID " & vbCrLf & _
                  " And POD.SupplierID = PUD.SupplierID "

        ls_sql = ls_sql + " And POD.PartNo = PUD.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo " & vbCrLf & _
                          " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          " AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                          " LEFT JOIN (select pono, orderNo, affiliateID, SupplierID, PartNo," & vbCrLf & _
                          "             Min(labelNo) as label1, Max(labelNo) as label2 from PrintLabelExport " & vbCrLf & _
                          "             Group by pono, orderNo, affiliateID, SupplierID, PartNo) PL ON PL.PONo = POD.PONo   " & vbCrLf & _
                          "         and PL.AffiliateID = POD.AffiliateID  AND PL.SupplierID = POD.SupplierID" & vbCrLf & _
                          "         AND PL.PartNO = POD.PartNo  AND PL.orderNo = POD.OrderNo1" & vbCrLf & _
                          " LEFT JOIN MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls" & vbCrLf & _
                          " where PUD.PONO = '" & Trim(PONO) & "' " & vbCrLf & _
                          " AND PUD.AffiliateID = '" & Trim(affiliateID) & "'" & vbCrLf & _
                          " AND PUD.SupplierID = '" & Trim(SupplierID) & "' " & vbCrLf & _
                          " AND PUD.OrderNO1 = '" & Trim(OrderNo) & "' " & vbCrLf

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function BindDataLabelPrint(ByVal GB As GlobalSetting.clsGlobal, ByVal pPono As String, ByVal pOrderNo As String, ByVal pAff As String, ByVal pSupp As String) As DataSet
        Dim ls_sql As String

        ls_sql = " select    " & vbCrLf & _
                  " Partno = POD.PartNo,   " & vbCrLf & _
                  " Partname = MP.PartName,   " & vbCrLf & _
                  " labelno = Rtrim(PL.Label1) + ' - ' + Rtrim(pl.Label2),  " & vbCrLf & _
                  " Label1 = Rtrim(label1), " & vbCrLf & _
                  " Label2 = Rtrim(label2), " & vbCrLf & _
                  " orderno = POD.PONo,   " & vbCrLf & _
                  " uom = MU.Description,   " & vbCrLf & _
                  " qtybox = Qtybox,   " & vbCrLf & _
                  " qty = convert(char,PUD.Week1),   " & vbCrLf & _
                  " boxqty = Ceiling(PUD.week1 / QtyBox),   "

        ls_sql = ls_sql + " DestinationPort = isnull(MA.DestinationPort,''), " & vbCrLf & _
                          " DestinationPoint = isnull(DeliveryPoint,''), " & vbCrLf & _
                          " CustName = POD.AffiliateID, " & vbCrLf & _
                          " CustCode = AffiliateCode, " & vbCrLf & _
                          " ConsigneeCode = isnull(MA.ConsigneeCode,'') " & vbCrLf & _
                          " From PO_Detail_Export POD INNER JOIN PO_DetailUpload_Export PUD  " & vbCrLf & _
                          " ON POD.PONo = PUD.PONo   " & vbCrLf & _
                          " AND POD.AffiliateID = PUD.AffiliateID   " & vbCrLf & _
                          " AND POD.SupplierID = PUD.SupplierID   " & vbCrLf & _
                          " AND POD.partNO = PUD.PartNo   " & vbCrLf & _
                          " INNER JOIN PO_Master_Export POM   " & vbCrLf & _
                          " ON POM.Pono = POD.PONo   "

        ls_sql = ls_sql + " and POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
                          " And POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                          " LEFT JOIN (select pono, orderNo, affiliateID, SupplierID, PartNo,  " & vbCrLf & _
                          " 			Min(labelNo) as label1, Max(labelNo) as label2 from PrintLabelExport " & vbCrLf & _
                          " 			Group by pono, orderNo, affiliateID, SupplierID, PartNo) PL ON PL.PONo = POD.PONo   " & vbCrLf & _
                          " and PL.AffiliateID = POD.AffiliateID  AND PL.SupplierID = POD.SupplierID   " & vbCrLf & _
                          " AND PL.PartNO = POD.PartNo   " & vbCrLf & _
                          " INNER JOIN MS_Parts MP ON MP.PartNo = POD.partNo   " & vbCrLf & _
                          " LEFT JOIN MS_PARTMApping MPM ON MPM.PartNo = PL.PartNo and MPM.AffiliateID = PL.AffiliateID and MPM.SupplierID = PL.SupplierID  " & vbCrLf & _
                          " INNER JOIN ms_unitcls MU ON MU.UnitCls = MP.Unitcls   " & vbCrLf & _
                          " LEFT JOIN ms_affiliate MA ON MA.AffiliateID = POD.AffiliateID "

        ls_sql = ls_sql + " LEFT JOIN ms_Forwarder MF ON MF.ForwarderID = POD.ForwarderID " & vbCrLf & _
                          " LEFT JOIN MS_Supplier MS ON MS.SupplierID = POD.SupplierID " & vbCrLf & _
                          " Where PL.POno = '" & Trim(pPono) & "'" & vbCrLf & _
                          " AND PL.OrderNo = '" & Trim(pOrderNo) & "' " & vbCrLf & _
                          " AND PL.AffiliateID = '" & Trim(pAff) & "' " & vbCrLf & _
                          " AND PL.SupplierID = '" & Trim(pSupp) & "'"

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function BindDataOrderConfirmation(ByVal GB As GlobalSetting.clsGlobal, ByVal PONO As String, ByVal OrderNo As String, ByVal affiliateID As String, ByVal SupplierID As String) As DataSet
        Dim ls_sql As String

        ls_sql = " SELECT DISTINCT period = CONVERT(char(7),POM.PERIOD), " & vbCrLf & _
                  " EmergencyCls, " & vbCrLf & _
                  " orderNO = POM.OrderNo1, " & vbCrLf & _
                  " PONo = POM.PONo, " & vbCrLf & _
                  " shipby = ISNULL(POM.ShipCls,''), " & vbCrLf & _
                  " AffCode = ISNULL(MA.ConsigneeCode,''), " & vbCrLf & _
                  " AffName = ISNULL(MA.AffiliateName,''), " & vbCrLf & _
                  " AffAdd = ISNULL(MA.ConsigneeAddress,''), " & vbCrLf & _
                  " SuppName = ISNULL(MS.SupplierName,''), " & vbCrLf & _
                  " SuppAdd = ISNULL(MS.ADDRESS,''), " & vbCrLf & _
                  " FwdName = ISNULL(MF.ForwarderName,''), " & vbCrLf & _
                  " FwdAdd = ISNULL(MF.Address,''), " & vbCrLf & _
                  " ETDVendor = convert(char(10),convert(datetime, ETDVendor1),120), " & vbCrLf & _
                  " ETDPort = convert(char(10),convert(datetime, ETDPort1),120), " & vbCrLf & _
                  " ETAPort = convert(char(10),convert(datetime, ETAPort1),120), " & vbCrLf & _
                  " ETAFactory = convert(char(10),convert(datetime, ETAFactory1),120), " & vbCrLf

        ls_sql = ls_sql + " PartNo = POD.PartNo, " & vbCrLf & _
                          " PartName = MP.PartName, " & vbCrLf & _
                          " UOM = ISNULL(MU.DESCRIPTION,''), " & vbCrLf & _
                          " QtyBox = Qtybox, " & vbCrLf & _
                          " Qty = POD.Week1, " & vbCrLf & _
                          " SupplierQty = PUD.Week1, " & vbCrLf & _
                          " TotalBox = POD.Week1 / QtyBox, " & vbCrLf & _
                          " TotalNet = POD.Week1 * (NetWeight/1000), " & vbCrLf & _
                          " Volume = ((Length * Height * Width) * (POD.Week1/QtyBox))/1000 " & vbCrLf & _
                          " FROM PO_MASTER_EXPORT POM LEFT JOIN PO_DETAIL_EXPORT POD " & vbCrLf & _
                          " ON POM.PONO = POD.PONO " & vbCrLf

        ls_sql = ls_sql + " AND POM.ORDERNO1 = POD.ORDERNO1 " & vbCrLf & _
                          " AND POM.AFFILIATEID = POD.AFFILIATEID " & vbCrLf & _
                          " AND POM.SUPPLIERID = POD.SUPPLIERID " & vbCrLf & _
                          " LEFT JOIN PO_DetailUpload_Export PUD " & vbCrLf & _
                          " ON POD.AffiliateID = PUD.AffiliateID " & vbCrLf & _
                          " And POD.SupplierID = PUD.SupplierID " & vbCrLf & _
                          " And POD.PartNo = PUD.PartNo " & vbCrLf & _
                          " AND POD.PONO = PUD.PONO	 " & vbCrLf & _
                          " AND POD.OrderNo1 = PUD.OrderNo1 " & vbCrLf & _
                          " LEFT JOIN MS_AFFILIATE MA ON MA.AFFILIATEID = POM.AFFILIATEID " & vbCrLf & _
                          " LEFT JOIN MS_SUPPLIER MS ON MS.SUPPLIERID = POM.SUPPLIERID " & vbCrLf

        ls_sql = ls_sql + " LEFT JOIN MS_PARTS MP ON MP.PARTNO = POD.PARTNO " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo " & vbCrLf & _
                          " AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                          " AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
                          " LEFT JOIN MS_FORWARDER MF ON MF.FORWARDERID = POM.FORWARDERID " & vbCrLf & _
                          " LEFT JOIN MS_UNITCLS MU ON MU.UNITCLS = MP.UNITCLS " & vbCrLf & _
                          " LEFT JOIN MS_SHIPCLS MSC ON MSC.SHIPCLS = POM.SHIPCLS " & vbCrLf & _
                          " where PUD.PONO = '" & Trim(PONO) & "' " & vbCrLf & _
                          " AND PUD.AffiliateID = '" & Trim(affiliateID) & "'" & vbCrLf & _
                          " AND PUD.SupplierID = '" & Trim(SupplierID) & "' " & vbCrLf & _
                          " AND PUD.OrderNO1 = '" & Trim(OrderNo) & "' " & vbCrLf


        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_sql)
        Return ds
    End Function

    Public Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pFileName2 As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_Attachment2 As String = ""
            Dim ls_URl As String = ""

            sendEmailtoSupplier = True

            Dim dsEmail As New DataSet
            dsEmail = clsGeneral.getEmailAddressExport(GB, "", "PASI", pSupplier, "", "SupplierDeliveryCC", "SupplierDeliveryTO", "SupplierDeliveryTO", errMsg)

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
                errMsg = "Process Send PO Export [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Send PO Export [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Send Delivery Confirmation : " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("18", "", pPONo.Trim & "-" & pSupplier.Trim)
            ls_Attachment = Trim(pPathFile) & pFileName
            ls_Attachment2 = Trim(pPathFile) & pFileName2

            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment, ls_Attachment2) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Send PO Export [" & pPONo & "] from Affiliate [" & pAffiliate & "] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Public Shared Function sendEmailtoForwarder(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pFWD As String, ByRef errMsg As String) As Boolean
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

            ls_Subject = "Send Final Approval Order No: " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("19", "", pPONo.Trim & "-" & pSupplier.Trim)
            ls_Attachment = Trim(pPathFile) & pFileName

            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
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

    Public Shared Sub UpdateStatusPOExport(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffiliateID As String, ByVal pSupp As String, ByVal pPoNo As String, ByVal pOrderNo1 As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " update dbo.PO_Master_Export set FinalApprovalCls = '2' " & vbCrLf & _
                         " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                         " AND SupplierID = '" & pSupp & "' " & vbCrLf & _
                         " AND PoNo = '" & pPoNo & "'" & vbCrLf & _
                         " AND OrderNo1 = '" & pOrderNo1 & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

        Catch ex As Exception
            errMsg = "Process Send PONo [" & pPoNo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub
End Class
