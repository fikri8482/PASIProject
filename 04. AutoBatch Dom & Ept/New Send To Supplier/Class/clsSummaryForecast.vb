Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsSummaryForecast
    Shared Sub up_SendSummaryForecast(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")
        Dim xi As Integer, iRow As Integer

        Dim ExcelBook As Excel.Workbook
        Dim ExcelSheet As Excel.Worksheet
        Dim xlApp As New Excel.Application
        Dim sheetNumber As Integer = 1

        Dim receiptCCEmail As String = ""
        Dim receiptEmail As String = ""
        Dim fromEmail As String = ""

        Dim ls_SQL As String = ""
        Dim NewFileCopy As String = ""

        Dim pSupplier As String = ""
        Dim pPONo As String = ""
        Dim pDate As Date
        Dim pPartNo As String = ""
        Dim pForecastPeriod As Date

        Dim ds As New DataSet
        Dim dsDetail As New DataSet

        Try
            Dim fi As New FileInfo(pAtttacment & "\TemplateSummaryForecast.xlsx") 'File dari Local
            If Not fi.Exists Then
                errMsg = "Process Send Summary Forecast to Supplier STOPPED, File Excel isn't Found"
                ErrSummary = "Process Send Summary Forecast to Supplier STOPPED, File Excel isn't Found"
                Exit Sub
            End If

            log.WriteToProcessLog(Date.Now, pScreenName, "Get data Summary Forecast")

            ls_SQL = "SELECT * FROM SupplierSumForecast_Request WHERE SendExcel='1'"

            ds = GB.uf_GetDataSet(ls_SQL)

            If ds.Tables(0).Rows.Count > 0 Then
                For xi = 0 To ds.Tables(0).Rows.Count - 1
                    Try
                        pDate = ds.Tables(0).Rows(0)("ReqDate")
                        pPartNo = Trim(ds.Tables(0).Rows(0)("PartNo"))
                        pForecastPeriod = CDate(Trim(ds.Tables(0).Rows(0)("Period")))
                        pSupplier = Trim(ds.Tables(0).Rows(0)("SupplierID"))

                        log.WriteToProcessLog(Date.Now, pScreenName, "Create Excel Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "]")

                        NewFileCopy = pAtttacment & "\TemplateSummaryForecast.xlsx"

                        If System.IO.File.Exists(NewFileCopy) = True Then
                            System.IO.File.Delete(pResult & "\TemplateSummaryForecast.xlsx")
                            System.IO.File.Copy(NewFileCopy, pResult & "\TemplateSummaryForecast.xlsx")
                        Else
                            System.IO.File.Copy(NewFileCopy, pResult & "\SummaryForecast.xlsm")
                        End If

                        Dim ls_file As String = pResult & "\TemplateSummaryForecast.xlsx"

                        ExcelBook = xlApp.Workbooks.Open(ls_file)
                        ExcelSheet = CType(ExcelBook.Worksheets(sheetNumber), Excel.Worksheet)

                        dsDetail = bindDataForecast(GB, pDate, pSupplier, pForecastPeriod, pPartNo)

                        If dsDetail.Tables(0).Rows.Count > 0 Then
                            log.WriteToProcessLog(Date.Now, pScreenName, "Fill detail Excel Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "]")

                            With ExcelSheet
                                .Cells(5, 1).Value = "(" & Format(pForecastPeriod, "MMM yyyy") & " Production)"
                                .Cells(7, 1).Value = "Issue Date : " & Format(Now, "dd MMM yyyy")

                                .Cells(9, 8).Value = Format(DateAdd(DateInterval.Month, 1 - 1, pForecastPeriod), "MMM")
                                .Cells(9, 9).Value = Format(DateAdd(DateInterval.Month, 2 - 1, pForecastPeriod), "MMM")
                                .Cells(9, 10).Value = Format(DateAdd(DateInterval.Month, 3 - 1, pForecastPeriod), "MMM")
                                .Cells(9, 11).Value = Format(DateAdd(DateInterval.Month, 4 - 1, pForecastPeriod), "MMM")

                                iRow = 10
                                For j = 0 To dsDetail.Tables(0).Rows.Count - 1
                                    .Cells(iRow, 1) = j + 1 'Trim(dsDetail.Tables(0).Rows(j)("NoUrut")) & ""
                                    .Cells(iRow, 2) = Trim(dsDetail.Tables(0).Rows(j)("PartNo")) & ""
                                    .Cells(iRow, 3) = Trim(dsDetail.Tables(0).Rows(j)("AffiliateID")) & ""
                                    .Cells(iRow, 4) = Trim(dsDetail.Tables(0).Rows(j)("SupplierID")) & ""
                                    .Cells(iRow, 5) = Trim(dsDetail.Tables(0).Rows(j)("MOQ")) & ""
                                    .Cells(iRow, 6) = Trim(dsDetail.Tables(0).Rows(j)("Project")) & ""
                                    .Cells(iRow, 7) = Trim(dsDetail.Tables(0).Rows(j)("PONo")) & ""
                                    .Cells(iRow, 8) = Trim(dsDetail.Tables(0).Rows(j)("Bln1")) & ""
                                    .Cells(iRow, 9) = Trim(dsDetail.Tables(0).Rows(j)("Bln2")) & ""
                                    .Cells(iRow, 10) = Trim(dsDetail.Tables(0).Rows(j)("Bln3")) & ""
                                    .Cells(iRow, 11) = Trim(dsDetail.Tables(0).Rows(j)("Bln4")) & ""
                                    .Cells(iRow, 12) = ""

                                    With .Range(.Cells(iRow, 1), .Cells(iRow, 12))
                                        .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                                        .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                                        .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                                        .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                                        .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
                                        .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
                                    End With

                                    iRow = iRow + 1
                                    .Range(.Cells(iRow, 5), .Cells(iRow, 5)).NumberFormat = "#,###"
                                    .Range(.Cells(iRow, 8), .Cells(iRow, 11)).NumberFormat = "#,###"
                                    .Range(.Cells(iRow, 2), .Cells(iRow, 12)).EntireColumn.AutoFit()
                                Next
                                xlApp.DisplayAlerts = False

                                Dim temp_Filename As String = "Summary Forecast " & Trim(pSupplier) & " " & Format(Now, "ddMMyyyy hhmmss") & ".xlsx"
                                ExcelBook.SaveAs(pResult & "\" & temp_Filename)
                                ExcelBook.Close()
                                xlApp.Workbooks.Close()
                                xlApp.Quit()

                                log.WriteToProcessLog(Date.Now, pScreenName, "Finish Create Excel Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "] ok.")

                                If sendEmailSummaryForecast(GB, pResult, temp_Filename, pDate, pForecastPeriod, pSupplier, errMsg) = False Then
                                    Exit Try
                                Else
                                    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "] ok.")
                                End If

                                Call UpdateExcelSummaryForecast(pDate, pForecastPeriod, pPartNo, pSupplier, errMsg)

                                log.WriteToProcessLog(Date.Now, pScreenName, "Send Email. Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "] ok.")
                                clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Send Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "] ok", LogName)
                                LogName.Refresh()
                            End With
                        End If
                    Catch ex As Exception
                        xlApp.DisplayAlerts = False
                        ExcelBook.Close()
                        xlApp.Workbooks.Close()
                        xlApp.Quit()
                        log.WriteToErrorLog(pScreenName, "Process Send Summary Forecast [" & pSupplier & "-" & pDate & "] Notification to Supplier STOPPED, because " & ex.Message, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Send Summary Forecast [" & pSupplier & "-" & pDate & "] Notification to Supplier STOPPED, because " & ex.Message)

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Send Summary Forecast [" & pSupplier & "-" & pDate & "] Notification to Supplier STOPPED, because " & ex.Message, LogName)
                        LogName.Refresh()
                    End Try
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "] " & ex.Message
            ErrSummary = "Summary Forecast Supplier [" & pSupplier & "], Request Date [" & pDate & "] " & ex.Message
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
            If Not ds Is Nothing Then
                ds.Dispose()
            End If
            If Not dsDetail Is Nothing Then
                dsDetail.Dispose()
            End If
        End Try
    End Sub

    Shared Function bindDataForecast(ByVal GB As GlobalSetting.clsGlobal, ByVal pDate As Date, ByVal pSupplierID As String, ByVal pPeriod As Date, ByVal pPartNo As String) As DataSet
        Dim ls_SQL As String = ""
        Dim ls_End As String = ""
        Dim pWhere As String = ""

        If pPartNo.Trim <> "" Then
            pWhere = pWhere & " and b.PartNo = '" & pPartNo & "'"
        End If

        If pSupplierID.Trim <> "" Then
            pWhere = pWhere & " and b.SupplierID = '" & pSupplierID & "'"
        End If

        ls_SQL = " SELECT  " & vbCrLf & _
                                  "     DISTINCT " & vbCrLf & _
                                  " 	b.PartNo,  " & vbCrLf & _
                                  " 	b.AffiliateID, " & vbCrLf & _
                                  " 	b.SupplierID, " & vbCrLf & _
                                  " 	MOQ = ISNULL(b.POMOQ,d.MOQ), " & vbCrLf & _
                                  " 	c.Project, " & vbCrLf & _
                                  " 	b.PONo, " & vbCrLf & _
                                  " 	ISNULL(b.POQty,0) Bln1, " & vbCrLf & _
                                  " 	ISNULL(b.ForecastN1,0) Bln2, " & vbCrLf & _
                                  " 	ISNULL(b.ForecastN2,0) Bln3, " & vbCrLf & _
                                  " 	ISNULL(b.ForecastN3,0) Bln4 "

        ls_SQL = ls_SQL + " FROM PO_Master a " & vbCrLf & _
                          " INNER JOIN PO_Detail b on a.PONO = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping d on b.PartNo = b.PartNo and b.AffiliateID = d.AffiliateID and b.SupplierID = d.SupplierID " & vbCrLf & _
                          " WHERE YEAR(Period) = " & Year(pPeriod) & " and MONTH(Period) = " & Month(pPeriod) & " and FinalApproveDate IS NOT NULL " & pWhere & "" & vbCrLf & _
                          " ORDER BY b.PartNo "


        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function sendEmailSummaryForecast(GB As GlobalSetting.clsGlobal, ByVal pPathFile As String, ByVal pFileName As String, ByVal pDate As Date, ByVal pPeriod As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
        Dim dsEmail As New DataSet
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_sql As String = ""
            Dim ls_Subject As String = ""
            Dim ls_Body As String = ""
            Dim ls_Attachment As String = ""
            Dim ls_URl As String = ""

            Dim ds As New DataSet

            sendEmailSummaryForecast = True

            dsEmail = clsGeneral.getEmailAddress(GB, "", "PASI", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("FLAG") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("EmailFrom")
                End If
            Next

            ls_sql = "SELECT * FROM SupplierSumForecast_Request " & vbCrLf & _
                    " WHERE ReqDate='" & pDate & "' AND SupplierID='" & pSupplier & "' AND Period='" & pPeriod & "' AND SendExcel='1' "
            ds = GB.uf_GetDataSet(ls_sql)
            If ds.Tables(0).Rows.Count > 0 Then
                receiptEmail = ds.Tables(0).Rows(0)("EmailFrom")
            End If

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If fromEmail = "" Then
                sendEmailSummaryForecast = False
                errMsg = "Process Send Summary Forecast from PASI [PASI] to Supplier [" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailSummaryForecast = False
                errMsg = "Process Send Summary Forecast from PASI [PASI] to Supplier [" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Send To Supplier Summary Forecast Period: " & pPeriod & "-" & pSupplier
            ls_Body = clsNotification.GetNotification("90")
            ls_Body = Replace(ls_Body, "SUMMARY OUTSTANDING", "SUMMARY FORECAST")
            ls_Attachment = Trim(pPathFile) & "\" & pFileName

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailSummaryForecast = False
                Exit Function
            End If

            sendEmailSummaryForecast = True

        Catch ex As Exception
            sendEmailSummaryForecast = False
            errMsg = "Process Send Summary Forecast from PASI [PASI] to Supplier [" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Sub UpdateExcelSummaryForecast(ByVal pDate As Date, ByVal pPONo As String, ByVal pPartNo As String, ByVal pSuppCode As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " UPDATE dbo.SupplierSumForecast_Request " & vbCrLf & _
                      " SET SendExcel = '2'" & vbCrLf & _
                      " WHERE ReqDate = '" & pDate & "' " & vbCrLf & _
                      " AND Period='" & pPONo & "' " & vbCrLf & _
                      " AND PartNo = '" & pPartNo & "' " & vbCrLf & _
                      " AND SupplierID = '" & pSuppCode & "' "
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Send Summary Forecast [" & pPONo & "] to Supplier STOPPED, because " & ex.Message
        End Try
    End Sub
End Class
