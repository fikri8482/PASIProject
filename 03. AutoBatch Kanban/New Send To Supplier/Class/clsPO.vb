Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsPO
    Shared Sub up_SendPODomestic(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim fromEmail As String = ""
        Dim NewFileCopy As String

        Dim ds As New DataSet
        Dim dsDetail As New DataSet
        Dim dsEta As New DataSet
        Dim dsEmail As New DataSet
        Dim dsSupp As New DataSet
        Dim dsAffp2 As New DataSet
        Dim dsAffp As New DataSet
        Dim dsPASI As New DataSet

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim ls_SQL As String = ""
        Dim pPeriod As Date
        Dim pPODate As Date

        Dim pAffiliate As String = ""
        Dim pSupplier As String = ""
        Dim pPONo As String = ""
        Dim pDel As String = ""
        Dim pCommercialCls As String = ""
        Dim pShipCls As String = ""
        Dim pAffiliateName As String = ""
        Dim pAffiliateAdd As String = ""
        Dim pConsigneeName As String = ""
        Dim pConsigneeAdd As String = ""
        Dim pDelivBy As String = ""
        Dim receiptCCEmail As String = ""

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template PO.xlsm") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send PO to Supplier STOPPED, File Excel isn't Found"
                ErrSummary = "Process Send PO to Supplier STOPPED, File Excel isn't Found"
                Exit Sub
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Get data PO")

            ls_SQL = "SELECT b.* FROM Affiliate_Master a " & vbCrLf & _
                     "inner join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID  = b.SupplierID" & vbCrLf & _
                     "WHERE a.ExcelCls='1'"
            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    Try
                        pAffiliate = Trim(ds.Tables(0).Rows(xi)("AffiliateID"))
                        pPONo = Trim(ds.Tables(0).Rows(xi)("PONo"))
                        pSupplier = Trim(ds.Tables(0).Rows(xi)("SupplierID"))
                        pPeriod = Format(ds.Tables(0).Rows(xi)("Period"), "MMM-yyyy")
                        pPODate = Format(ds.Tables(0).Rows(xi)("EntryDate"), "dd-MMM-yyyy")
                        pCommercialCls = IIf(ds.Tables(0).Rows(xi)("CommercialCls") = "1", "YES", "NO")
                        pShipCls = Trim(ds.Tables(0).Rows(xi)("ShipCls"))
                        pDelivBy = ds.Tables(0).Rows(xi)("DeliveryByPASICls")

                        If ds.Tables(0).Rows(xi)("DeliveryByPASICls") = "1" Then
                            pDel = "PASI"
                        Else
                            pDel = pAffiliate
                        End If

                        log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")

                        NewFileCopy = pAtttacment & "\Template PO.xlsm"

                        If System.IO.File.Exists(NewFileCopy) = True Then
                            System.IO.File.Delete(pResult & "\Template PO.xlsm")
                            System.IO.File.Copy(NewFileCopy, pResult & "\Template PO.xlsm")
                        Else
                            System.IO.File.Copy(NewFileCopy, pResult & "\PO.xlsm")
                        End If

                        Dim ls_file As String = pResult & "\Template PO.xlsm"

                        ExcelBook = xlApp.Workbooks.Open(ls_file)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        dsDetail = bindDataDetail(GB, pPeriod, pAffiliate, pPONo, pSupplier)
                        dsEta = bindDataETA(GB, pAffiliate, pSupplier, pPeriod)

                        If dsDetail.Tables(0).Rows.Count > 0 Then

                            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

                            For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                                If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                                    fromEmail = dsEmail.Tables(0).Rows(i)("EmailFrom")
                                    receiptCCEmail = dsEmail.Tables(0).Rows(i)("EmailCC")
                                End If
                            Next

                            receiptCCEmail = Replace(receiptCCEmail, ",", ";")

                            ExcelSheet.Range("H1").Value = "PO"
                            ExcelSheet.Range("H2").Value = fromEmail
                            ExcelSheet.Range("H3").Value = pAffiliate
                            ExcelSheet.Range("H5").Value = pSupplier

                            'ExcelSheet.Range("Y2").Value = receiptCCEmail

                            ExcelSheet.Range("I9").Value = pPONo
                            ExcelSheet.Range("T9").Value = Format(pPeriod, "MMM-yyyy")
                            ExcelSheet.Range("AE9").Value = Format(pPODate, "dd-MMM-yyyy")


                            dsSupp = clsGeneral.Supplier(GB, Trim(pSupplier))
                            ExcelSheet.Range("I11").Value = dsSupp.Tables(0).Rows(0)("SupplierName")
                            ExcelSheet.Range("I12").Value = dsSupp.Tables(0).Rows(0)("Address")
                            ExcelSheet.Range("I12:X14").WrapText = True

                            'Buyer
                            dsAffp2 = clsGeneral.Affiliate(GB, Trim(pAffiliate))
                            dsPASI = clsGeneral.PASI(GB, "PASI")

                            pAffiliateName = IIf(Trim(dsAffp2.Tables(0).Rows(0)("BuyerName")) = "", Trim(dsPASI.Tables(0).Rows(0)("AffiliateName")), Trim(dsAffp2.Tables(0).Rows(0)("BuyerName")))
                            pAffiliateAdd = IIf(Trim(dsAffp2.Tables(0).Rows(0)("BuyerName")) = "", Trim(dsPASI.Tables(0).Rows(0)("Address")), Trim(dsAffp2.Tables(0).Rows(0)("BuyerAddress")))
                            ExcelSheet.Range("I16").Value = pAffiliateName 'dsAffp2.Tables(0).Rows(0)("BuyerName")
                            ExcelSheet.Range("I17").Value = pAffiliateAdd 'dsAffp2.Tables(0).Rows(0)("BuyerAddress")
                            ExcelSheet.Range("I17:X19").WrapText = True

                            ExcelSheet.Range("AE12").Value = pCommercialCls
                            ExcelSheet.Range("AE14").Value = pShipCls

                            'Consignee
                            dsAffp = clsGeneral.Affiliate(GB, Trim(pAffiliate))
                            pConsigneeName = IIf(Trim(dsAffp.Tables(0).Rows(0)("ConsigneeName")) = "", Trim(dsAffp.Tables(0).Rows(0)("AffiliateName")), Trim(dsAffp.Tables(0).Rows(0)("ConsigneeName")))
                            pConsigneeAdd = IIf(Trim(dsAffp.Tables(0).Rows(0)("ConsigneeAddress")) = "", Trim(dsAffp.Tables(0).Rows(0)("ConsigneeAddress")), Trim(dsAffp.Tables(0).Rows(0)("ConsigneeAddress")))
                            ExcelSheet.Range("AE16").Value = pConsigneeName 'dsAffp.Tables(0).Rows(0)("ConsigneeName")
                            ExcelSheet.Range("AE17").Value = pConsigneeAdd 'dsAffp.Tables(0).Rows(0)("ConsigneeAddress")
                            ExcelSheet.Range("AE17:AT19").WrapText = True

                            If dsEta.Tables(0).Rows.Count > 0 Then
                                log.WriteToProcessLog(Date.Now, pScreenName, "Fill ETA detail Excel PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")

                                For i = 1 To dsEta.Tables(0).Rows.Count - 1
                                    ExcelSheet.Range("AX" & i + 33 & ": AY" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day1") '1
                                    ExcelSheet.Range("AZ" & i + 33 & ": BA" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day2") '2
                                    ExcelSheet.Range("BB" & i + 33 & ": BC" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day3") '3
                                    ExcelSheet.Range("BD" & i + 33 & ": BE" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day4") '4
                                    ExcelSheet.Range("BF" & i + 33 & ": BG" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day5") '5
                                    ExcelSheet.Range("BH" & i + 33 & ": BI" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day6") '6
                                    ExcelSheet.Range("BJ" & i + 33 & ": BK" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day7") '7
                                    ExcelSheet.Range("BL" & i + 33 & ": BM" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day8") '8
                                    ExcelSheet.Range("BN" & i + 33 & ": BO" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day9") '9
                                    ExcelSheet.Range("BP" & i + 33 & ": BQ" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day10") '10
                                    ExcelSheet.Range("BR" & i + 33 & ": BS" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day11") '11
                                    ExcelSheet.Range("BT" & i + 33 & ": BU" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day12") '12
                                    ExcelSheet.Range("BV" & i + 33 & ": BW" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day13") '13
                                    ExcelSheet.Range("BX" & i + 33 & ": BY" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day14") '14
                                    ExcelSheet.Range("BZ" & i + 33 & ": CA" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day15") '15
                                    ExcelSheet.Range("CB" & i + 33 & ": CC" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day16") '16
                                    ExcelSheet.Range("CD" & i + 33 & ": CE" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day17") '17
                                    ExcelSheet.Range("CF" & i + 33 & ": CG" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day18") '18
                                    ExcelSheet.Range("CH" & i + 33 & ": CI" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day19") '19
                                    ExcelSheet.Range("CJ" & i + 33 & ": CK" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day20") '20
                                    ExcelSheet.Range("CL" & i + 33 & ": CM" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day21") '21
                                    ExcelSheet.Range("CN" & i + 33 & ": CO" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day22") '22
                                    ExcelSheet.Range("CP" & i + 33 & ": CQ" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day23") '23
                                    ExcelSheet.Range("CR" & i + 33 & ": CS" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day24") '24
                                    ExcelSheet.Range("CT" & i + 33 & ": CU" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day25") '25
                                    ExcelSheet.Range("CV" & i + 33 & ": CW" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day26") '26
                                    ExcelSheet.Range("CX" & i + 33 & ": CY" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day27") '27
                                    ExcelSheet.Range("CZ" & i + 33 & ": DA" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day28") '28
                                    ExcelSheet.Range("DB" & i + 33 & ": DC" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day29") '29
                                    ExcelSheet.Range("DD" & i + 33 & ": DE" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day30") '30
                                    ExcelSheet.Range("DF" & i + 33 & ": DG" & i + 33).Value = dsEta.Tables(0).Rows(i)("Day31") '31
                                Next
                            End If

                            log.WriteToProcessLog(Date.Now, pScreenName, "Fill detail Excel PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")
                            For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                                'Header
                                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Merge()
                                ExcelSheet.Range("D" & i + 36 & ": H" & i + 36).Merge()
                                ExcelSheet.Range("I" & i + 36 & ": P" & i + 36).Merge()
                                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).Merge()
                                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).Merge()
                                ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Merge()
                                ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Merge()
                                ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Merge()
                                ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Merge()
                                ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Merge()
                                ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Merge()
                                ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Merge()
                                ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Merge()
                                ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Merge() '1
                                ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Merge() '2
                                ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Merge() '3
                                ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Merge() '4
                                ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Merge() '5
                                ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Merge() '6
                                ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Merge() '7
                                ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Merge() '8
                                ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Merge() '9
                                ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Merge() '10
                                ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Merge() '11
                                ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Merge() '12
                                ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Merge() '13
                                ExcelSheet.Range("BX" & i + 36 & ": BY" & i + 36).Merge() '14
                                ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Merge() '15
                                ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Merge() '16
                                ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Merge() '17
                                ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Merge() '18
                                ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Merge() '19
                                ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Merge() '20
                                ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Merge() '21
                                ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Merge() '22
                                ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Merge() '23
                                ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Merge() '24
                                ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Merge() '25
                                ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Merge() '26
                                ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Merge() '27
                                ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Merge() '28
                                ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Merge() '29
                                ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Merge() '30
                                ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Merge() '31

                                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("NoUrut"))
                                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                                ExcelSheet.Range("B" & i + 36 & ": C" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                                ExcelSheet.Range("D" & i + 36 & ": H" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                                ExcelSheet.Range("I" & i + 36 & ": P" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("KanbanCls"))
                                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                                ExcelSheet.Range("Q" & i + 36 & ": S" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("Description"))
                                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                                ExcelSheet.Range("T" & i + 36 & ": U" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                                ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("MOQ"))
                                ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).NumberFormat = "#,##0"
                                ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("QtyBox"))
                                ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).NumberFormat = "#,##0"
                                ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("Maker")) & ""
                                ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Value = If(IsDBNull(dsDetail.Tables(0).Rows(i)("POQty")), 0, dsDetail.Tables(0).Rows(i)("POQty"))
                                ExcelSheet.Range("AC" & i + 36 & ": DE" & i + 36).NumberFormat = "#,##0"

                                ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("ForecastN1"))
                                ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).NumberFormat = "#,##0"

                                ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("ForecastN2"))
                                ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).NumberFormat = "#,##0"

                                ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("ForecastN3"))
                                ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).NumberFormat = "#,##0"

                                ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Value = Trim(dsDetail.Tables(0).Rows(i)("BYWHAT"))

                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) = "BY AFFILIATE" Then
                                    ExcelSheet.Range("AG" & i + 36).Value = "YES"
                                    ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Value = "ORDER"
                                    ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Interior.Color = ColorYellow
                                Else
                                    ExcelSheet.Range("V" & i + 36 & ": W" & i + 36).Value = ""
                                    ExcelSheet.Range("X" & i + 36 & ": Y" & i + 36).Value = ""
                                    ExcelSheet.Range("Z" & i + 36 & ": AB" & i + 36).Value = ""
                                    ExcelSheet.Range("AI" & i + 36 & ": AK" & i + 36).Value = ""
                                    ExcelSheet.Range("AL" & i + 36 & ": AN" & i + 36).Value = ""
                                    ExcelSheet.Range("AO" & i + 36 & ": AQ" & i + 36).Value = ""
                                    ExcelSheet.Range("AR" & i + 36 & ": AW" & i + 36).Value = "SUPPLIER APPROVAL"
                                    ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Interior.Color = ColorYellow
                                    ExcelSheet.Range("AC" & i + 36 & ": AE" & i + 36).Locked = False
                                    ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).Interior.Color = ColorYellow
                                    ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).Locked = False
                                End If

                                ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD1") '1
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Locked = False

                                ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD2") '2
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("AZ" & i + 36 & ": BA" & i + 36).Locked = False

                                ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD3") '3
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BB" & i + 36 & ": BC" & i + 36).Locked = False

                                ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD4") '4
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BD" & i + 36 & ": BE" & i + 36).Locked = False

                                ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD5") '5
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BF" & i + 36 & ": BG" & i + 36).Locked = False

                                ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD6") '6
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BH" & i + 36 & ": BI" & i + 36).Locked = False

                                ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD7") '7
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BJ" & i + 36 & ": BK" & i + 36).Locked = False

                                ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD8") '8
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BL" & i + 36 & ": BM" & i + 36).Locked = False

                                ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD9") '9
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BN" & i + 36 & ": BO" & i + 36).Locked = False


                                ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD10") '10
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BP" & i + 36 & ": BQ" & i + 36).Locked = False

                                ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD11") '11
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BR" & i + 36 & ": BS" & i + 36).Locked = False

                                ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD12") '12
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BT" & i + 36 & ": BU" & i + 36).Locked = False

                                ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD13") '13
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BV" & i + 36 & ": BW" & i + 36).Locked = False

                                ExcelSheet.Range("BX" & i + 36 & ": BY" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD14") '14
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("AX" & i + 36 & ": AY" & i + 36).Locked = False

                                ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD15") '15
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("BZ" & i + 36 & ": CA" & i + 36).Locked = False

                                ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD16") '16
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CB" & i + 36 & ": CC" & i + 36).Locked = False

                                ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD17") '17
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CD" & i + 36 & ": CE" & i + 36).Locked = False

                                ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD18") '18
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CF" & i + 36 & ": CG" & i + 36).Locked = False

                                ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD19") '19
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CH" & i + 36 & ": CI" & i + 36).Locked = False

                                ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD20") '20
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CJ" & i + 36 & ": CK" & i + 36).Locked = False

                                ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD21") '21
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CL" & i + 36 & ": CM" & i + 36).Locked = False

                                ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD22") '22
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CN" & i + 36 & ": CO" & i + 36).Locked = False

                                ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD23") '23
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CP" & i + 36 & ": CQ" & i + 36).Locked = False

                                ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD24") '24
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CR" & i + 36 & ": CS" & i + 36).Locked = False

                                ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD25") '25
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CT" & i + 36 & ": CU" & i + 36).Locked = False

                                ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD26") '26
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CV" & i + 36 & ": CW" & i + 36).Locked = False

                                ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD27") '27
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CX" & i + 36 & ": CY" & i + 36).Locked = False

                                ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD28") '28
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("CZ" & i + 36 & ": DA" & i + 36).Locked = False

                                ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD29") '29
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("DB" & i + 36 & ": DC" & i + 36).Locked = False

                                ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD30") '30
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("DD" & i + 36 & ": DE" & i + 36).Locked = False

                                ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Value = dsDetail.Tables(0).Rows(i)("DeliveryD31") '31
                                If Trim(dsDetail.Tables(0).Rows(i)("BYWHAT")) <> "BY AFFILIATE" Then ExcelSheet.Range("DF" & i + 36 & ": DG" & i + 36).Locked = False

                                ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).NumberFormat = "#,##0"
                                ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                                ExcelSheet.Range("AX" & i + 36 & ": DG" & i + 36).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                                clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & i + 36 & ": AE" & i + 36))
                                clsGeneral.DrawAllBorders(ExcelSheet.Range("AI" & i + 36 & ": DG" & i + 36))
                                ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                ExcelSheet.Range("AF" & i + 36 & ": AH" & i + 36).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            Next

                            ExcelSheet.Range("B38").Interior.Color = Color.White
                            ExcelSheet.Range("B38").Font.Color = Color.Black
                            ExcelSheet.Range("B" & i + 36).Value = "E"
                            ExcelSheet.Range("B" & i + 36).Interior.Color = Color.Black
                            ExcelSheet.Range("B" & i + 36).Font.Color = Color.White

                            ExcelSheet.EnableSelection = XlEnableSelection.xlNoRestrictions

                            'ExcelSheet.Protect("tosis123", , , , , , , , , , , , , True)

                            xlApp.DisplayAlerts = False

                            Dim temp_Filename As String = "PO " & Trim(pPONo) & "-" & Trim(pAffiliate) & "-" & Trim(pSupplier) & ".xlsm"
                            ExcelBook.SaveAs(pResult & "\" & temp_Filename)
                            ExcelBook.Close()
                            xlApp.Workbooks.Close()
                            xlApp.Quit()

                            log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")

                            If sendEmailtoSupplier(GB, pResult, temp_Filename, pPONo, pAffiliate, pSupplier, errMsg) = False Then
                                Exit Try
                            Else
                                log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            End If

                            'If sendEmailtoPASI(GB, pPeriod, pPONo, pAffiliate, pSupplier, pAffiliateName, pDelivBy, errMsg) = False Then
                            'Else
                            '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC PASI. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            'End If

                            'If sendEmailtoAffiliate(GB, pPeriod, pPONo, pAffiliate, pSupplier, pAffiliateName, pDelivBy, errMsg) = False Then
                            'Else
                            '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC Affiliate. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            'End If

                            '0000 Perlu di aktifin
                            Call UpdateExcelPO(pAffiliate, pPONo, pSupplier, errMsg)

                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
                            LogName.Refresh()
                        End If
                    Catch ex As Exception
                        xlApp.DisplayAlerts = False
                        ExcelBook.Close()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                        log.WriteToErrorLog(pScreenName, "Process Create PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Create PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message, LogName)
                        LogName.Refresh()
                    End Try
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & ex.Message
            ErrSummary = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & ex.Message
        Finally
            If System.IO.File.Exists(NewFileCopy) = True Then
                System.IO.File.Delete(pResult & "\Template PO.xlsm")
            End If

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
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
            If Not dsDetail Is Nothing Then
                dsDetail.Dispose()
            End If
            If Not dsEta Is Nothing Then
                dsEta.Dispose()
            End If
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
            If Not dsSupp Is Nothing Then
                dsSupp.Dispose()
            End If
            If Not dsAffp2 Is Nothing Then
                dsAffp2.Dispose()
            End If
            If Not dsAffp Is Nothing Then
                dsAffp.Dispose()
            End If
            If Not dsPASI Is Nothing Then
                dsPASI.Dispose()
            End If
        End Try

    End Sub

    Shared Function bindDataDetail(ByVal GB As GlobalSetting.clsGlobal, ByVal pDate As Date, ByVal pAffCode As String, ByVal pPONo As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT Sort, NoUrut, PartNo, PartNos, PartName ,KanbanCls  ,ISNULL(Description, '')Description    " & vbCrLf & _
                  " ,MOQ, QtyBox, Maker ,BYWHAT    " & vbCrLf & _
                  " ,POQty, ForecastN1, ForecastN2, ForecastN3   " & vbCrLf & _
                  " ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5   " & vbCrLf & _
                  " ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10   " & vbCrLf & _
                  " ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15   " & vbCrLf & _
                  " ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20   " & vbCrLf & _
                  " ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25   " & vbCrLf & _
                  " ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31   " & vbCrLf & _
                  " FROM (    " & vbCrLf & _
                  "     SELECT row_number() over (order by AD.PONo) as Sort, CONVERT(CHAR,row_number() over (order by AD.PONo)) as NoUrut    " & vbCrLf

        ls_SQL = ls_SQL + "     ,AD.PartNo as PartNo ,AD.PartNo AS PartNos, PartName ,CASE WHEN PD.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls ,MU.Description    " & vbCrLf & _
                          "     ,MOQ = ISNULL(PD.POMOQ,MPM.MOQ), QtyBox = ISNULL(PD.POQtyBox,MPM.QtyBox), MPART.Maker   " & vbCrLf & _
                          "     ,'BY AFFILIATE' BYWHAT ,AD.POQty POqty " & vbCrLf & _
                          "     ,ISNULL(ForecastN1,0) ForecastN1 " & vbCrLf & _
                          "     ,ISNULL(ForecastN2,0) ForecastN2 " & vbCrLf & _
                          "     ,ISNULL(ForecastN3,0) ForecastN3 " & vbCrLf & _
                          "     ,AD.DeliveryD1, AD.DeliveryD2, AD.DeliveryD3, AD.DeliveryD4, AD.DeliveryD5   " & vbCrLf & _
                          "     ,AD.DeliveryD6, AD.DeliveryD7, AD.DeliveryD8, AD.DeliveryD9, AD.DeliveryD10   " & vbCrLf & _
                          "     ,AD.DeliveryD11, AD.DeliveryD12, AD.DeliveryD13, AD.DeliveryD14, AD.DeliveryD15   " & vbCrLf & _
                          "     ,AD.DeliveryD16, AD.DeliveryD17, AD.DeliveryD18, AD.DeliveryD19, AD.DeliveryD20   " & vbCrLf & _
                          "     ,AD.DeliveryD21, AD.DeliveryD22, AD.DeliveryD23, AD.DeliveryD24, AD.DeliveryD25   " & vbCrLf

        ls_SQL = ls_SQL + "   		,AD.DeliveryD26, AD.DeliveryD27, AD.DeliveryD28, AD.DeliveryD29, AD.DeliveryD30, AD.DeliveryD31   " & vbCrLf & _
                          "  		FROM dbo.Affiliate_Detail AD  " & vbCrLf & _
                          "  		LEFT JOIN dbo.PO_Detail PD ON AD.PONO = PD.PONo and AD.AffiliateID = PD.AffiliateID and AD.SupplierID = PD.SupplierID and AD.PartNo = PD.PartNo  " & vbCrLf & _
                          "  		LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo  " & vbCrLf & _
                          "  		LEFT JOIN dbo.MS_PartMapping MPM ON MPM.PartNo = AD.PartNo and AD.AffiliateID = MPM.AffiliateID and MPM.SupplierID = AD.SupplierID    " & vbCrLf & _
                          "  		LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls         " & vbCrLf & _
                          "  		WHERE AD.PONo='" & Trim(pPONo) & "' AND AD.AffiliateID='" & Trim(pAffCode) & "' AND AD.SupplierID='" & Trim(pSupplierID) & "' " & vbCrLf

        ls_SQL = ls_SQL + "   		GROUP BY AD.PONo,AD.PartNo,PartName,PD.KanbanCls,AD.POQty, ForecastN1, ForecastN2, ForecastN3, MU.Description, ISNULL(PD.POMOQ,MPM.MOQ), ISNULL(PD.POQtyBox,MPM.QtyBox), MPART.Maker,AD.PartNo,AD.AffiliateID    " & vbCrLf & _
                          "      	,AD.DeliveryD1,AD.DeliveryD2,AD.DeliveryD3,AD.DeliveryD4,AD.DeliveryD5   " & vbCrLf & _
                          "   		,AD.DeliveryD6,AD.DeliveryD7,AD.DeliveryD8,AD.DeliveryD9,AD.DeliveryD10   " & vbCrLf & _
                          "   		,AD.DeliveryD11,AD.DeliveryD12,AD.DeliveryD13,AD.DeliveryD14,AD.DeliveryD15   " & vbCrLf & _
                          "   		,AD.DeliveryD16,AD.DeliveryD17,AD.DeliveryD18,AD.DeliveryD19,AD.DeliveryD20   " & vbCrLf & _
                          "   		,AD.DeliveryD21,AD.DeliveryD22,AD.DeliveryD23,AD.DeliveryD24,AD.DeliveryD25   " & vbCrLf & _
                          "   		,AD.DeliveryD26,AD.DeliveryD27,AD.DeliveryD28,AD.DeliveryD29,AD.DeliveryD30,AD.DeliveryD31  )detail1   " & vbCrLf & _
                          "  UNION ALL  " & vbCrLf & _
                          "  SELECT Sort,NoUrut , PartNo, PartNos, PartName, KanbanCls, ISNULL(Description,'')Description   " & vbCrLf & _
                          "      ,MOQ, QtyBox ,Maker ,BYWHAT , POqty " & vbCrLf & _
                          "      ,ForecastN1 ,ForecastN2 ,ForecastN3  " & vbCrLf

        ls_SQL = ls_SQL + "      ,ISNULL(DeliveryD1,0)DeliveryD1,ISNULL(DeliveryD2,0)DeliveryD2,ISNULL(DeliveryD3,0)DeliveryD3,ISNULL(DeliveryD4,0)DeliveryD4,ISNULL(DeliveryD5,0)DeliveryD5  " & vbCrLf & _
                          "      ,ISNULL(DeliveryD6,0)DeliveryD6,ISNULL(DeliveryD7,0)DeliveryD7,ISNULL(DeliveryD8,0)DeliveryD8,ISNULL(DeliveryD9,0)DeliveryD9,ISNULL(DeliveryD10,0)DeliveryD10  " & vbCrLf & _
                          "      ,ISNULL(DeliveryD11,0)DeliveryD11,ISNULL(DeliveryD12,0)DeliveryD12,ISNULL(DeliveryD13,0)DeliveryD13,ISNULL(DeliveryD14,0)DeliveryD14,ISNULL(DeliveryD15,0)DeliveryD15  " & vbCrLf & _
                          "      ,ISNULL(DeliveryD16,0)DeliveryD16,ISNULL(DeliveryD17,0)DeliveryD17,ISNULL(DeliveryD18,0)DeliveryD18,ISNULL(DeliveryD19,0)DeliveryD19,ISNULL(DeliveryD20,0)DeliveryD20  " & vbCrLf & _
                          "      ,ISNULL(DeliveryD21,0)DeliveryD21,ISNULL(DeliveryD22,0)DeliveryD22,ISNULL(DeliveryD23,0)DeliveryD23,ISNULL(DeliveryD24,0)DeliveryD24,ISNULL(DeliveryD25,0)DeliveryD25  " & vbCrLf & _
                          "      ,ISNULL(DeliveryD26,0)DeliveryD26,ISNULL(DeliveryD27,0)DeliveryD27,ISNULL(DeliveryD28,0)DeliveryD28,ISNULL(DeliveryD29,0)DeliveryD29,ISNULL(DeliveryD30,0)DeliveryD30,ISNULL(DeliveryD31,0)DeliveryD31 " & vbCrLf & _
                          "       FROM (    " & vbCrLf & _
                          "   	    SELECT row_number() over (order by AD.PONo) as Sort,'' as NoUrut,'' PartNo,AD.PartNo PartNos,''PartName,'' KanbanCls,''Description, MOQ = ISNULL(PD.POMOQ,MPM.MOQ)     " & vbCrLf & _
                          "   	    ,QtyBox = ISNULL(PD.POQtyBox,MPM.QtyBox), ISNULL(MPART.Maker,'')Maker ,'BY PASI' BYWHAT ,AD.POqty POQty " & vbCrLf & _
                          "         ,ISNULL(ForecastN1,0) ForecastN1 " & vbCrLf & _
                          "    	    ,ISNULL(ForecastN2,0) ForecastN2 " & vbCrLf

        ls_SQL = ls_SQL + "    	    ,ISNULL(ForecastN3,0) ForecastN3 " & vbCrLf & _
                          "      	,AD.DeliveryD1 ,AD.DeliveryD2 ,AD.DeliveryD3 ,AD.DeliveryD4 ,AD.DeliveryD5 ,AD.DeliveryD6 ,AD.DeliveryD7 ,AD.DeliveryD8 ,AD.DeliveryD9 ,AD.DeliveryD10    " & vbCrLf & _
                          "         ,AD.DeliveryD11 ,AD.DeliveryD12 ,AD.DeliveryD13 ,AD.DeliveryD14 ,AD.DeliveryD15 ,AD.DeliveryD16 ,AD.DeliveryD17 ,AD.DeliveryD18 ,AD.DeliveryD19 ,AD.DeliveryD20   " & vbCrLf & _
                          "         ,AD.DeliveryD21 ,AD.DeliveryD22 ,AD.DeliveryD23 ,AD.DeliveryD24 ,AD.DeliveryD25 ,AD.DeliveryD26 ,AD.DeliveryD27 ,AD.DeliveryD28 ,AD.DeliveryD29 ,AD.DeliveryD30 ,AD.DeliveryD31  " & vbCrLf & _
                          "         FROM dbo.Affiliate_Detail AD  " & vbCrLf & _
                          "  		LEFT JOIN dbo.PO_Detail PD ON AD.PONO = PD.PONo and AD.AffiliateID = PD.AffiliateID and AD.SupplierID = PD.SupplierID and AD.PartNo = PD.PartNo  " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_Parts MPART ON AD.PartNo = MPART.PartNo   " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_PartMapping MPM ON MPM.PartNo = AD.PartNo and AD.AffiliateID = MPM.AffiliateID and MPM.SupplierID = AD.SupplierID     " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls   " & vbCrLf

        ls_SQL = ls_SQL + "  WHERE AD.PONo='" & Trim(pPONo) & "' AND AD.AffiliateID='" & Trim(pAffCode) & "' AND AD.SupplierID='" & Trim(pSupplierID) & "' " & vbCrLf & _
                          "     GROUP BY AD.PONo,AD.PartNo,PartName,PD.KanbanCls,AD.POQty, ForecastN1, ForecastN2, ForecastN3, MU.Description, ISNULL(PD.POMOQ,MPM.MOQ), ISNULL(PD.POQtyBox,MPM.QtyBox), MPART.Maker,AD.PartNo,AD.AffiliateID    " & vbCrLf & _
                          "     ,AD.DeliveryD1,AD.DeliveryD2,AD.DeliveryD3,AD.DeliveryD4,AD.DeliveryD5   " & vbCrLf & _
                          "   	,AD.DeliveryD6,AD.DeliveryD7,AD.DeliveryD8,AD.DeliveryD9,AD.DeliveryD10   " & vbCrLf & _
                          "   	,AD.DeliveryD11,AD.DeliveryD12,AD.DeliveryD13,AD.DeliveryD14,AD.DeliveryD15   " & vbCrLf & _
                          "   	,AD.DeliveryD16,AD.DeliveryD17,AD.DeliveryD18,AD.DeliveryD19,AD.DeliveryD20   " & vbCrLf & _
                          "   	,AD.DeliveryD21,AD.DeliveryD22,AD.DeliveryD23,AD.DeliveryD24,AD.DeliveryD25   " & vbCrLf & _
                          "   	,AD.DeliveryD26,AD.DeliveryD27,AD.DeliveryD28,AD.DeliveryD29,AD.DeliveryD30,AD.DeliveryD31)detail2   " & vbCrLf & _
                          "  ORDER BY sort, PartNos, NoUrut DESC  "


        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function bindDataETA(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pPeriod As Date) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " 	SELECT  " & vbCrLf & _
                  " 		'ETAAffiliate'as ETA  " & vbCrLf
        For i = 1 To 31
            ls_SQL = ls_SQL + " 		,Day" & i & " = ISNULL(MAX(CASE WHEN  DAY(ETAAffiliate) = " & i & " THEN DAY(ETAAffiliate) END),0) " & vbCrLf
        Next

        ls_SQL = ls_SQL + " FROM dbo.MS_ETD_PASI ETDP  " & vbCrLf & _
                          " LEFT JOIN dbo.MS_ETD_Supplier_PASI ETDSP ON ETDP.ETDPASI =  ETDSP.ETAPASI " & vbCrLf & _
                          " WHERE AffiliateID='" & pAffCode & "' AND SupplierID='" & pSupplierID & "' and YEAR(ETDP.ETAAffiliate) = '" & Year(pPeriod) & "' and MONTH(ETDP.ETAAffiliate) = " & Month(pPeriod) & " " & vbCrLf & _
                          " UNION ALL  " & vbCrLf & _
                          " SELECT  " & vbCrLf & _
                          " 	'ETAPASI'as ETA  " & vbCrLf
        For i = 1 To 31
            ls_SQL = ls_SQL + " 		,Day" & i & " = ISNULL(MAX(CASE WHEN  DAY(ETAAffiliate) = " & i & " THEN DAY(ETAPASI) END),0) " & vbCrLf
        Next
        ls_SQL = ls_SQL + " FROM dbo.MS_ETD_PASI ETDP  " & vbCrLf & _
                          " LEFT JOIN dbo.MS_ETD_Supplier_PASI ETDSP ON ETDP.ETDPASI =  ETDSP.ETAPASI " & vbCrLf & _
                          " WHERE AffiliateID='" & pAffCode & "' AND SupplierID='" & pSupplierID & "' and YEAR(ETDP.ETAAffiliate) = '" & Year(pPeriod) & "' and MONTH(ETDP.ETAAffiliate) = " & Month(pPeriod) & " " & vbCrLf & _
                          " UNION ALL  " & vbCrLf & _
                          " SELECT  " & vbCrLf & _
                          " 	'ETDSupplier'as ETA  " & vbCrLf
        For i = 1 To 31
            ls_SQL = ls_SQL + " 		,Day" & i & " = ISNULL(MAX(CASE WHEN  DAY(ETAAffiliate) = " & i & " THEN DAY(ETDSupplier) END),0) " & vbCrLf
        Next

        ls_SQL = ls_SQL + " FROM dbo.MS_ETD_PASI ETDP  " & vbCrLf & _
                          " LEFT JOIN dbo.MS_ETD_Supplier_PASI ETDSP ON ETDP.ETDPASI =  ETDSP.ETAPASI " & vbCrLf & _
                          " WHERE AffiliateID='" & pAffCode & "' AND SupplierID='" & pSupplierID & "' and YEAR(ETDP.ETAAffiliate) = '" & Year(pPeriod) & "' and MONTH(ETDP.ETAAffiliate) = " & Month(pPeriod) & " " & vbCrLf

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Sub UpdateExcelPO(ByVal pAffCode As String, ByVal pPONo As String, ByVal pSuppCode As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""        
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.Affiliate_Master " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
                      " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                      " AND SupplierID='" & pSuppCode & "' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send PONo [" & pPONo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub

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

            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", pSupplier, "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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

            ls_Subject = "Send To Supplier PO No: " & pPONo & "-" & pSupplier
            ls_Body = clsNotification.GetNotification("11", "", pPONo.Trim)
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

    Shared Function sendEmailtoPASI(ByVal GB As GlobalSetting.clsGlobal, ByVal pPeriod As Date, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pAffiliateName As String, ByVal pDelivBy As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailtoPASI = True

            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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
                errMsg = "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] Notification to PASI STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] Notification to PASI STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderDetail.aspx?id2=" & clsNotification.EncryptURL(pPONo.Trim) & _
                        "&t1=" & clsNotification.EncryptURL(pAffiliate) & "&t2=" & clsNotification.EncryptURL(pSupplier) & _
                        "&t3=" & clsNotification.EncryptURL(pPeriod) & _
                        "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderList.aspx")

            ls_Subject = "Send To Supplier PO No: " & pPONo & "-" & pSupplier
            ls_Body = clsNotification.GetNotification("11", ls_URl, pPONo.Trim)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoPASI = False
                Exit Function
            End If

            sendEmailtoPASI = True

        Catch ex As Exception
            sendEmailtoPASI = False
            errMsg = "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] Notification to PASI STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal pPeriod As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pAffiliateName As String, ByVal pDelivBy As String, ByRef errMsg As String) As Boolean
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

            dsEmail = clsGeneral.getEmailAddress(GB, pAffiliate, "PASI", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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

            If receiptEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] Notification to Affiliate STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If fromEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] Notification to Affiliate STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerName & "/PurchaseOrder/POEntry.aspx?id2=" & clsNotification.EncryptURL(pPONo.Trim) & "&t1=" & clsNotification.EncryptURL(pAffiliate.Trim) & _
                                   "&t2=" & clsNotification.EncryptURL("") & "&t3=" & clsNotification.EncryptURL(pPeriod) & "&t4=" & clsNotification.EncryptURL(pSupplier.Trim) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrder/POList.aspx")
            ls_Subject = "Send To Supplier PO No: " & pPONo & "-" & pSupplier
            ls_Body = clsNotification.GetNotification("11", ls_URl, pPONo.Trim)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoAffiliate = False
                Exit Function
            End If

            sendEmailtoAffiliate = True

        Catch ex As Exception
            sendEmailtoAffiliate = False
            errMsg = "Process Send PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] Notification to Affiliate STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function
End Class
