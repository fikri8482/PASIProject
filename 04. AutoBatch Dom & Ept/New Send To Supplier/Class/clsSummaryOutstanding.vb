Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsSummaryOutstanding
    Shared Sub up_SendSummaryOutstanding(ByVal cfg As GlobalSetting.clsConfig,
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

        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim ls_SQL As String = ""
        Dim NewFileCopy As String = ""

        Dim temp_Filename As String = ""
        Dim pSupplier As String = ""
        Dim pPONo As String = ""
        Dim pDate As Date

        Dim ds As New DataSet
        Dim dsHeader As New DataSet

        Try
            Dim fi As New FileInfo(pAtttacment & "\Template Summary Outstanding.xlsx") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send PO to Supplier STOPPED, File Excel isn't Found"
                ErrSummary = "Process Send PO to Supplier STOPPED, File Excel isn't Found"
                Exit Sub
            End If

            ls_SQL = "SELECT * FROM SupplierSumOutstanding_Request WHERE SendExcel='1'"
            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                NewFileCopy = pAtttacment & "\Template Summary Outstanding.xlsx"

                If System.IO.File.Exists(NewFileCopy) = True Then
                    System.IO.File.Delete(pResult & "\Template Summary Outstanding")
                    System.IO.File.Copy(NewFileCopy, pResult & "\Template Summary Outstanding")
                Else
                    System.IO.File.Copy(NewFileCopy, pResult & "\Template Summary Outstanding")
                End If

                Dim ls_file As String = pResult & "\Template Summary Outstanding"

                ExcelBook = xlApp.Workbooks.Open(ls_file)
                ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                pPONo = Trim(ds.Tables(0).Rows(0)("PONo"))
                pSupplier = Trim(ds.Tables(0).Rows(0)("SupplierID"))

                dsHeader = bindDataOutstanding(GB, pSupplier, pPONo)

                If dsHeader.Tables(0).Rows.Count > 0 Then
                    'Header
                    Dim iRow As Long
                    With ExcelSheet
                        .Range("A2:A2").Value = "Period : " & Format(pDate, "MMM yyyy")
                        .Range("A2:A2").Font.Bold = True
                        .Range("A2:A2").Font.Size = 12

                        iRow = 4

                        Dim iCol As Integer = 0, iNextCol As Integer = 0

                        iRow = iRow + 2
                        'Detail Report                
                        Dim j As Long = 0
                        For j = 0 To dsHeader.Tables(0).Rows.Count - 1
                            .Cells(iRow, 1) = j + 1 'Trim(dsHeader.Tables(0).Rows(j)("ColNo")) & ""
                            .Cells(iRow, 2) = dsHeader.Tables(0).Rows(j)("Period") & ""
                            .Cells(iRow, 3) = Trim(dsHeader.Tables(0).Rows(j)("PONo")) & ""
                            .Cells(iRow, 4) = Trim(dsHeader.Tables(0).Rows(j)("AffiliateID")) & ""
                            .Cells(iRow, 5) = Trim(dsHeader.Tables(0).Rows(j)("SupplierID")) & ""
                            .Cells(iRow, 6) = Trim(dsHeader.Tables(0).Rows(j)("KanbanNo")) & ""
                            .Cells(iRow, 7) = Trim(dsHeader.Tables(0).Rows(j)("ETDSupp")) & ""

                            .Cells(iRow, 8) = Trim(dsHeader.Tables(0).Rows(j)("PartNo")) & ""
                            .Cells(iRow, 9) = Trim(dsHeader.Tables(0).Rows(j)("PartName")) & ""

                            .Cells(iRow, 10).Value = FormatNumber(IIf(Trim(dsHeader.Tables(0).Rows(j)("QtyPO")) = "", 0, Trim(dsHeader.Tables(0).Rows(j)("QtyPO"))), 0, TriState.True) & ""
                            .Cells(iRow, 11).Value = FormatNumber(IIf(Trim(dsHeader.Tables(0).Rows(j)("RemainingQtyPOPASI")) = "", 0, Trim(dsHeader.Tables(0).Rows(j)("RemainingQtyPOPASI"))), 0, TriState.True) & ""

                            .Cells(iRow, 12) = dsHeader.Tables(0).Rows(j)("SupplierDeliveryDate") & ""
                            .Cells(iRow, 13) = dsHeader.Tables(0).Rows(j)("SupplierSuratJalanNo") & ""
                            .Cells(iRow, 14) = Trim(dsHeader.Tables(0).Rows(j)("PASIReceiveDate")) & ""

                            .Cells(iRow, 15).Value = FormatNumber(IIf(Trim(dsHeader.Tables(0).Rows(j)("SupplierDeliveryQty")) = "", 0, Trim(dsHeader.Tables(0).Rows(j)("SupplierDeliveryQty"))), 0, TriState.True) & ""
                            .Cells(iRow, 16).Value = FormatNumber(IIf(Trim(dsHeader.Tables(0).Rows(j)("PASIReceivingQty")) = "", 0, Trim(dsHeader.Tables(0).Rows(j)("PASIReceivingQty"))), 0, TriState.True) & ""

                            .Cells(iRow, 17).Value = Trim(dsHeader.Tables(0).Rows(j)("InvoiceNoFromSupplier")) & ""
                            .Cells(iRow, 18).Value = Trim(dsHeader.Tables(0).Rows(j)("InvoiceDateFromSupplier")) & ""
                            .Cells(iRow, 19).Value = Trim(dsHeader.Tables(0).Rows(j)("InvoiceFromSupplierCurr")) & ""
                            .Cells(iRow, 20).Value = FormatNumber(IIf(Trim(dsHeader.Tables(0).Rows(j)("InvoiceFromSupplierAmount")) = "", 0, Trim(dsHeader.Tables(0).Rows(j)("InvoiceFromSupplierAmount"))), 0, TriState.True) & ""

                            With .Range(.Cells(iRow, 1), .Cells(iRow, 22))
                                .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                                .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                                .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                                .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                                .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
                                .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                            End With
                            iRow = iRow + 1
                            .Range(.Cells(iRow, 20), .Cells(iRow, 18)).NumberFormat = "#,###"

                            If .Range(.Cells(iRow, 22), .Cells(iRow, 22)).Value <> 0 Then
                                .Range(.Cells(iRow, 22), .Cells(iRow, 22)).NumberFormat = "#,###.00"
                            End If
                            .Range(.Cells(iRow, 2), .Cells(iRow, 33)).EntireColumn.AutoFit()
                        Next

                        xlApp.DisplayAlerts = False

                        temp_Filename = "Summary Outstanding " & Trim(pSupplier) & "-" & Format(Now, "ddMMyyyy hhmmss") & ".xlsm"
                        ExcelBook.SaveAs(pResult & "\" & temp_Filename)
                        ExcelBook.Close()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()

                    End With

                    If sendEmailSummaryOutstanding(GB, pResult, temp_Filename, pDate, pPONo, pSupplier, errMsg) = False Then
                        GoTo keluar
                    End If

                    Call UpdateExcelSummaryOutstanding(pDate, pPONo, pSupplier, errMsg)
keluar:
                    xlApp.Workbooks.Close()
                    xlApp.Quit()
                    Exit Sub
                Else
                    errMsg = "-"
                    ErrSummary = "-"
                    Exit Try
                End If
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "PONo [" & pPONo & "], Supplier [" & pSupplier & "] " & ex.Message
            ErrSummary = "PONo [" & pPONo & "], Supplier [" & pSupplier & "] " & ex.Message
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
            If Not dsHeader Is Nothing Then
                dsHeader.Dispose()
            End If
        End Try
    End Sub

    Shared Function bindDataOutstanding(ByVal GB As GlobalSetting.clsGlobal, ByVal pSupplierID As String, ByVal pPONo As String) As DataSet
        Dim ls_SQL As String = ""
        Dim pWhere As String = "" '= " and YEAR(POM.Period) = " & Year(pDate) & " and MONTH(POM.Period) = " & Month(pDate) & ""

        If pSupplierID <> "" Then
            pWhere = pWhere & " and POM.SupplierID = '" & pSupplierID & "'"
        End If

        If pPONo <> "" Then
            pWhere = pWhere & " and POM.PONo = '" & pPONo & "'"
        End If

        ls_SQL = "  SELECT DISTINCT * FROM  " & vbCrLf & _
                  "  (  " & vbCrLf & _
                  "  	SELECT   " & vbCrLf & _
                  "  		POM.Period  " & vbCrLf & _
                  "  		,POM.PONo  " & vbCrLf & _
                  "  		,POM.AffiliateID  " & vbCrLf & _
                  "  		,POM.SupplierID  " & vbCrLf & _
                  " 		,KD.KanbanNo " & vbCrLf & _
                  " 		,ETDSupp = ABC.ETDSupplier   		 " & vbCrLf & _
                  "  		,POD.PartNo  " & vbCrLf & _
                  "  		,MP.PartName  "

        ls_SQL = ls_SQL + "  		,QtyPO = ISNULL(POD.POQty,0)  " & vbCrLf & _
                          " 		,RemainingQtyPOPASI = ISNULL(KD.KanbanQty,0) -  " & vbCrLf & _
                          "  		                      ISNULL(  " & vbCrLf & _
                          "  		                        (select SUM(DOQty) from DOSupplier_Detail ABC  " & vbCrLf & _
                          "  		                         WHERE ABC.SupplierID = SDD.SupplierID and ABC.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                          "  		                         and ABC.KanbanNo = SDD.KanbanNo and ABC.PartNo = SDD.PartNo and ABC.PONo = SDD.PONo),0)  " & vbCrLf & _
                          "  		,SupplierDeliveryDate = SDM.DeliveryDate  " & vbCrLf & _
                          "  		,SupplierSuratJalanNo = SDM.SuratJalanNo " & vbCrLf & _
                          " 		,PASIReceiveDate = PRM.ReceiveDate  " & vbCrLf & _
                          "  		,SupplierDeliveryQty = SDD.DOQty  	 		  " & vbCrLf & _
                          "  		,PASIReceivingQty = PRD.GoodRecQty		  "

        ls_SQL = ls_SQL + "  		,InvoiceNoFromSupplier = ISM.InvoiceNo  " & vbCrLf & _
                          "  		,InvoiceDateFromSupplier = ISM.InvoiceDate  " & vbCrLf & _
                          "  		,InvoiceFromSupplierCurr = 'IDR' " & vbCrLf & _
                          "  		,InvoiceFromSupplierAmount = ISNULL(ISD.InvAmount,0)		  " & vbCrLf & _
                          "  	FROM PO_Master POM  " & vbCrLf & _
                          "  	LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  							AND POM.PoNo = POD.PONo  " & vbCrLf & _
                          "  							AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  								AND KD.PoNo = POD.PONo  " & vbCrLf & _
                          "  								AND KD.SupplierID = POD.SupplierID  "

        ls_SQL = ls_SQL + "  								AND KD.PartNo = POD.PartNo  " & vbCrLf & _
                          "  	LEFT JOIN Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID  " & vbCrLf & _
                          "  								AND KD.KanbanNo = KM.KanbanNo  " & vbCrLf & _
                          "  								AND KD.SupplierID = KM.SupplierID  " & vbCrLf & _
                          "  								AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                          "  	LEFT JOIN DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
                          "  									AND KD.KanbanNo = SDD.KanbanNo  " & vbCrLf & _
                          "  									AND KD.SupplierID = SDD.SupplierID  " & vbCrLf & _
                          "  									AND KD.PartNo = SDD.PartNo  " & vbCrLf & _
                          "  									AND KD.PoNo = SDD.PoNo  " & vbCrLf & _
                          "  	LEFT JOIN DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  "

        ls_SQL = ls_SQL + "  									AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
                          "  									AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf & _
                          "  	LEFT JOIN ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                          "  									AND SDD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
                          "  									AND SDD.SupplierID = PRD.SupplierID  " & vbCrLf & _
                          "  									AND SDD.PartNo = PRD.PartNo  " & vbCrLf & _
                          "  									AND SDD.PONo = PRD.PONo								  " & vbCrLf & _
                          "  									AND SDD.SuratJalanNo = PRD.SuratJalanNo  " & vbCrLf & _
                          "  	LEFT JOIN ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                          "  									AND PRM.SuratJalanNo = PRD.SuratJalanNo  " & vbCrLf & _
                          "  									AND PRM.SupplierID = PRD.SupplierID  "

        ls_SQL = ls_SQL + "  	LEFT JOIN InvoiceSupplier_Detail ISD ON ISD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                          "  										AND ISD.SupplierID = PRD.SupplierID  " & vbCrLf & _
                          "  										AND ISD.SuratJalanNo = PRD.SuratJalanNo  " & vbCrLf & _
                          "  										AND ISD.PONo = PRD.PONo  " & vbCrLf & _
                          "  										AND ISD.PartNo = PRD.PartNo  " & vbCrLf & _
                          "  										AND ISD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
                          "  	LEFT JOIN InvoiceSupplier_Master ISM ON ISM.InvoiceNo = ISD.InvoiceNo  " & vbCrLf & _
                          "    										AND ISM.AffiliateID = ISD.AffiliateID  " & vbCrLf & _
                          "    										AND ISM.SupplierID = ISD.SupplierID  " & vbCrLf & _
                          "    										AND ISM.suratJalanno = ISD.SuratJalanNo  	  " & vbCrLf & _
                          "  	LEFT JOIN (   "

        ls_SQL = ls_SQL + "   				SELECT * FROM MS_ETD_PASI a   " & vbCrLf & _
                          "   				INNER JOIN MS_ETD_Supplier_PASI b on a.ETDPASI =  b.ETAPASI   " & vbCrLf & _
                          "   				)ABC ON POM.SupplierID = ABC.SupplierID and POM.AffiliateID = ABC.AffiliateID AND KM.KanbanDate = ABC.ETAAffiliate   " & vbCrLf & _
                          "  	LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                          "  	LEFT JOIN MS_CurrCls MCS ON MCS.CurrCls = ISD.InvCurrCls  	 " & vbCrLf & _
                          "  	WHERE KD.KanbanQty > 0  AND (YEAR(POM.Period) = '" & Format(Now, "yyyy") & "' AND MONTH(POM.Period) = '" & Format(Now, "MM") & "')" & vbCrLf & _
                          "                       " & pWhere & "  " & vbCrLf & _
                          "  " & vbCrLf & _
                          "  )XYZ  "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function sendEmailSummaryOutstanding(ByVal GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pDate As Date, ByVal pPONo As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
        Dim ds As New DataSet
        Dim dsEmail As New DataSet

        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_Sql As String = ""
            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            sendEmailSummaryOutstanding = True

            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
            Next

            ls_Sql = "SELECT * FROM SupplierSumOutstanding_Request " & vbCrLf & _
                    " WHERE ReqDate='" & pDate & "' AND SupplierID='" & pSupplier & "' AND PONo='" & pPONo & "' AND SendExcel='1' "
            ds = GB.uf_GetDataSet(ls_Sql)
            If ds.Tables(0).Rows.Count > 0 Then
                receiptEmail = ds.Tables(0).Rows(0)("EmailFrom")
            End If

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If fromEmail = "" Then
                sendEmailSummaryOutstanding = False
                errMsg = "Process Send Summary Outstanding [" & pPONo & "], Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailSummaryOutstanding = False
                errMsg = "Process Send Summary Outstanding [" & pPONo & "], Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Send To Supplier Summary Outstanding " & pPONo
            ls_Body = clsNotification.GetNotification("90", "", pPONo)
            ls_Attachment = Trim(pPathFile) & "\" & pFileName

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailSummaryOutstanding = False
                Exit Function
            End If

            sendEmailSummaryOutstanding = True

        Catch ex As Exception
            sendEmailSummaryOutstanding = False
            errMsg = "Process Send Summary Outstanding [" & pPONo & "], Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
        End Try
    End Function

    Shared Sub UpdateExcelSummaryOutstanding(ByVal pDate As Date, ByVal pPONo As String, ByVal pSuppCode As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.SupplierSumOutstanding_Request " & vbCrLf & _
                      " SET SendExcel = '2'" & vbCrLf & _
                      " WHERE ReqDate = '" & pDate & "' " & vbCrLf & _
                      " AND PONo='" & pPONo & "' " & vbCrLf & _
                      " AND SupplierID = '" & pSuppCode & "' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send Summary Outstanding [" & pPONo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub
End Class
