Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsPOExportEmergency
    Shared Sub up_SendPOExportEmergency(ByVal cfg As GlobalSetting.clsConfig,
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

        Dim i As Integer, xi As Integer
        Const ColorYellow As Single = 65535
        Dim receiptCCEmail As String = ""
        Dim fromEmail As String = ""

        Dim NewFileCopy As String

        Dim ls_SQL As String = ""

        Dim ds As New DataSet
        Dim dsHeader As New DataSet
        Dim dsDetail As New DataSet
        Dim dsEmail As New DataSet
        Dim dsSupp As New DataSet
        Dim dsAffp As New DataSet
        Dim dsDelivery As New DataSet

        Dim pPeriod As Date
        Dim pOrderNo1 As String

        Dim pAffCode As String = ""
        Dim pSupplier As String = ""
        Dim pPONo As String = ""
        Dim pDel As String = ""

        Try
            ls_SQL = "SELECT * FROM dbo.PO_Master_Export" & vbCrLf & _
                     "WHERE ExcelCls = '1' and EmergencyCls = 'E'"
            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    pAffCode = ds.Tables(0).Rows(xi)("AffiliateID").ToString.Trim
                    pPONo = ds.Tables(0).Rows(xi)("PONo").ToString.Trim
                    pSupplier = ds.Tables(0).Rows(xi)("SupplierID").ToString.Trim
                    pPeriod = ds.Tables(0).Rows(xi)("Period")
                    pDel = ds.Tables(0).Rows(xi)("ForwarderID").ToString.Trim
                    pOrderNo1 = ds.Tables(0).Rows(xi)("OrderNo1").ToString.Trim

                    log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "]")

                    dsEmail = clsGeneral.getEmailAddressPASI(GB, "", "PASI", "", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

                    For i = 0 To dsEmail.Tables(0).Rows.Count - 1
                        If dsEmail.Tables(0).Rows(i)("flag") = "PASI" Then
                            fromEmail = dsEmail.Tables(0).Rows(i)("EmailFrom")
                            receiptCCEmail = dsEmail.Tables(0).Rows(i)("EmailCC")
                        End If
                    Next

                    receiptCCEmail = Replace(receiptCCEmail, ",", ";")

                    'Create Excel File
                    Dim fi As New FileInfo(pAtttacment & "\Template PO Export (Emergency).xlsm") 'File dari Local
                    If Not fi.Exists Then
                        errMsg = "Process Send PO Export to Supplier STOPPED, File Excel isn't Found"
                        ErrSummary = "Process Send PO Export to Supplier STOPPED, File Excel isn't Found"
                        Exit Sub
                    End If

                    NewFileCopy = Trim(pAtttacment) & "\Template PO Export (Emergency).xlsm"
                    Dim ls_file As String = NewFileCopy
                    ExcelBook = xlApp.Workbooks.Open(ls_file)
                    ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                    ExcelSheet.Range("H1").Value = "POEE"
                    ExcelSheet.Range("H2").Value = fromEmail
                    ExcelSheet.Range("H3").Value = pAffCode
                    ExcelSheet.Range("H4").Value = pDel
                    ExcelSheet.Range("H5").Value = pSupplier

                    ExcelSheet.Range("S1").Value = pPONo

                    ExcelSheet.Range("Y2").Value = receiptCCEmail

                    'Order No
                    ExcelSheet.Range("I9").Value = pPONo

                    'Order No
                    If pPONo <> pOrderNo1 Then
                        ExcelSheet.Range("P9").Value = pOrderNo1
                    End If

                    'PO Date
                    ExcelSheet.Range("AE9").Value = ds.Tables(0).Rows(xi)("UploadDate")
                    ExcelSheet.Range("AE9").NumberFormat = "yyyy-MM-dd"

                    'Commercial Cls
                    ExcelSheet.Range("AE11").Value = IIf(ds.Tables(0).Rows(xi)("CommercialCls") = "1", "YES", "NO")

                    'To
                    ExcelSheet.Range("I11").Value = pSupplier
                    dsSupp = clsGeneral.Supplier(GB, Trim(pSupplier))
                    ExcelSheet.Range("I12").Value = dsSupp.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("I12:X14").WrapText = True

                    'Buyer
                    ExcelSheet.Range("I16").Value = pAffCode                    
                    dsAffp = clsGeneral.Affiliate(GB, Trim(pAffCode))
                    ExcelSheet.Range("I17").Value = dsAffp.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("I17:X19").WrapText = True

                    'Delivery To
                    ExcelSheet.Range("AE13").Value = pDel                    
                    dsDelivery = clsGeneral.Forwarder(GB, Trim(pDel))
                    ExcelSheet.Range("AE14").Value = dsDelivery.Tables(0).Rows(0)("Address")
                    ExcelSheet.Range("AE14:AT16").WrapText = True

                    ExcelSheet.Range("B38").Interior.Color = Color.White
                    ExcelSheet.Range("B38").Font.Color = Color.Black

                    dsDetail = bindDataDetailEmergency(GB, pAffCode, pPONo, pSupplier, pOrderNo1)

                    log.WriteToProcessLog(Date.Now, pScreenName, "Fill detail Excel PO Export [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "]")

                    If dsDetail.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dsDetail.Tables(0).Rows.Count - 1
                            'Header
                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).Merge() 'No
                            ExcelSheet.Range("D" & i + 37 & ": H" & i + 37).Merge() 'Part No
                            ExcelSheet.Range("I" & i + 37 & ": P" & i + 37).Merge() 'Part Name
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).Merge() 'UOM
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).Merge() 'MOQ
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Merge() ' Total Order
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Merge() 'ETD Supplier
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Merge() 'Total Firm Edit Supp

                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("NoUrut"))
                            ExcelSheet.Range("B" & i + 37 & ": C" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("D" & i + 37 & ": H" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("PartNo"))
                            ExcelSheet.Range("I" & i + 37 & ": P" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("PartName"))
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("UOM"))
                            ExcelSheet.Range("Q" & i + 37 & ": R" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("MOQ"))
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("S" & i + 37 & ": T" & i + 37).NumberFormat = "#,##0"

                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("U" & i + 37 & ": Y" & i + 37).NumberFormat = "#,##0"

                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).Value = Trim(ds.Tables(0).Rows(xi)("ETDVendor1"))
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            ExcelSheet.Range("Z" & i + 37 & ": AD" & i + 37).NumberFormat = "yyyy-MM-dd"

                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Value = Trim(dsDetail.Tables(0).Rows(i)("Week1"))
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).NumberFormat = "#,##0"
                            ExcelSheet.Range("AE" & i + 37 & ": AI" & i + 37).Locked = False

                            clsGeneral.DrawAllBorders(ExcelSheet.Range("B" & i + 37 & ": AI" & i + 37))
                            ExcelSheet.Range("AD" & i + 37 & ": AI" & i + 37).Interior.Color = ColorYellow
                            ExcelSheet.Range("B" & i + 37 & ": AD" & i + 37).Interior.Color = RGB(217, 217, 217)
                        Next
                    End If

                    ExcelSheet.Range("B" & i + 37).Value = "E"
                    ExcelSheet.Range("B" & i + 37).Interior.Color = Color.Black
                    ExcelSheet.Range("B" & i + 37).Font.Color = Color.White

                    ExcelSheet.EnableSelection = XlEnableSelection.xlNoRestrictions
                    'ExcelSheet.Protect("tosis123", , , , , , , , , , , , , True)
                    xlApp.DisplayAlerts = False

                    Dim temp_Filename As String = "PO Emergency " & Trim(pPONo) & "-" & Trim(pSupplier) & ".xlsm"
                    ExcelBook.SaveAs(Trim(pResult) & "\" & temp_Filename)
                    xlApp.Workbooks.Close()
                    xlApp.Quit()

                    log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel PO Export [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok.")

                    If sendEmailtoSupplier(GB, pResult, temp_Filename, pOrderNo1, pAffCode, pSupplier, errMsg) = False Then
                        Exit Try
                    Else
                        log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok.")
                    End If

                    Call UpdateExcelPOEmergency(pAffCode, pPONo, pOrderNo1, pSupplier, errMsg)

                    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email. PO Export [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok.")
                    clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send PO Export [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] ok", LogName)
                    LogName.Refresh()
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] " & ex.Message
            ErrSummary = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffCode & "] " & ex.Message
        Finally
            If Not xlApp Is Nothing Then
                clsGeneral.NAR(ExcelSheet)
                xlApp.Workbooks.Close()
                clsGeneral.NAR(ExcelBook)
                xlApp.Quit()
                clsGeneral.NAR(xlApp)
                GC.Collect()
            End If
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
            If Not dsDetail Is Nothing Then
                dsDetail.Dispose()
            End If
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
            If Not dsSupp Is Nothing Then
                dsSupp.Dispose()
            End If
            If Not dsAffp Is Nothing Then
                dsAffp.Dispose()
            End If
            If Not dsDelivery Is Nothing Then
                dsDelivery.Dispose()
            End If
        End Try
    End Sub

    Shared Function bindDataDetailEmergency(ByVal GB As GlobalSetting.clsGlobal, ByVal pAffCode As String, ByVal pPONo As String, ByVal pSupplierID As String, ByVal pOrderNo1 As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "  SELECT row_number() over (order by POD.PONo) as Sort,  " & vbCrLf & _
                  "  	CONVERT(CHAR,row_number() over (order by POD.PONo)) as NoUrut,  " & vbCrLf & _
                  "  	PONo = RTRIM(POD.PONo), PartNo = RTRIM(POD.PartNo), PartName = RTRIM(PartName),  " & vbCrLf & _
                  "  	UOM = MU.Description, MOQ = CONVERT(CHAR,MOQ), QtyBox = CONVERT(CHAR,QtyBox),  " & vbCrLf & _
                  "  	'ORDER' BYWHAT,  " & vbCrLf & _
                  "  	ISNULL(Week1,0)Week1, ETDVendor1 " & vbCrLf & _
                  "  FROM dbo.PO_Detail_Export POD  " & vbCrLf & _
                  "  INNER JOIN PO_Master_Export ME ON ME.PONo = POD.PONo AND ME.AffiliateID = POD.AffiliateID AND ME.SupplierID = POD.SupplierID  AND ME.OrderNo1 = POD.OrderNo1 " & vbCrLf & _
                  "  LEFT JOIN dbo.MS_Parts MPART ON POD.PartNo = MPART.PartNo  " & vbCrLf & _
                  "  LEFT JOIN MS_PartMapping MPM ON MPM.AffiliateID = POD.AffiliateID and MPM.SupplierID = POD.SupplierID and MPM.PartNo = POD.PartNo " & vbCrLf & _
                  "  LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MPART.UnitCls  " & vbCrLf

        ls_SQL = ls_SQL + " WHERE EmergencyCls = 'E' AND POD.PONo = '" & Trim(pPONo) & "' AND POD.AffiliateID = '" & Trim(pAffCode) & "' AND POD.SupplierID = '" & Trim(pSupplierID) & "' AND POD.OrderNo1 = '" & Trim(pOrderNo1) & "'" & vbCrLf

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Sub UpdateExcelPOEmergency(ByVal pAffCode As String, ByVal pPONo As String, ByVal pOrderNo As String, ByVal pSuppCode As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.PO_Master_Export " & vbCrLf & _
                      " SET ExcelCls='2'" & vbCrLf & _
                      " WHERE PONo='" & pPONo & "'  " & vbCrLf & _
                      " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                      " AND SupplierID='" & pSuppCode & "' " & vbCrLf & _
                      " AND OrderNo1='" & pOrderNo & "' " & vbCrLf & _
                      " AND EmergencyCls = 'E' "
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
            dsEmail = clsGeneral.getEmailAddressPASI(GB, "", "PASI", pSupplier, "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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

            ls_Subject = "Issued PO Export Emergency: " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("16", "", pPONo.Trim & "-" & pSupplier.Trim)
            ls_Attachment = Trim(pPathFile) & "\" & pFileName

            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
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

End Class
