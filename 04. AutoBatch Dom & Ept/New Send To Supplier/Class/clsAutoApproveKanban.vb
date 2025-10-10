Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsAutoApproveKanban
    Shared Sub AutoApproveKanban(ByVal cfg As GlobalSetting.clsConfig,
                                  ByVal log As GlobalSetting.clsLog,
                                  ByVal GB As GlobalSetting.clsGlobal,
                                  ByVal LogName As RichTextBox,
                                  ByVal pAtttacment As String,
                                  ByVal pResult As String,
                                  ByVal pScreenName As String,
                                  Optional ByRef errMsg As String = "",
                                  Optional ByRef ErrSummary As String = "")

        Dim dsGetKanbanApprove As New DataSet
        Dim dsGetIntervalKanbanApprove As New DataSet

        Dim iRow As Integer = 0

        Dim KAAff As String = ""
        Dim KASupplier As String = ""
        Dim KADeliveryLocation As String = ""
        Dim KAKanbanno As String = ""
        Dim KAKanbanDate As Date
        Dim KANeedApprove As Date
        Dim pDateNeedApproveKanban As Date
        Dim IntervalAkanbanApprove As TimeSpan

        Try
            log.WriteToProcessLog(Date.Now, "AutoApproveKanban", "Get data Kanban")

            dsGetKanbanApprove = getKanbanApprove(GB)
            If dsGetKanbanApprove.Tables(0).Rows.Count > 0 Then
                For iRow = 0 To dsGetKanbanApprove.Tables(0).Rows.Count - 1
                    KAAff = dsGetKanbanApprove.Tables(0).Rows(iRow)("AffiliateID")
                    KADeliveryLocation = If(IsDBNull(dsGetKanbanApprove.Tables(0).Rows(iRow)("DeliveryLocationCode")), "", dsGetKanbanApprove.Tables(0).Rows(iRow)("DeliveryLocationCode"))
                    KANeedApprove = If(IsDBNull(dsGetKanbanApprove.Tables(0).Rows(iRow)("AffiliateApproveDate")), "", dsGetKanbanApprove.Tables(0).Rows(iRow)("AffiliateApproveDate"))
                    KAKanbanDate = Format(dsGetKanbanApprove.Tables(0).Rows(iRow)("kanbanDate"), "yyyy-MM-dd")
                    KASupplier = dsGetKanbanApprove.Tables(0).Rows(iRow)("supplierID")
                    KAKanbanno = dsGetKanbanApprove.Tables(0).Rows(iRow)("KanbanNo")

                    pDateNeedApproveKanban = FormatDateTime(KANeedApprove, DateFormat.ShortTime)

                    dsGetIntervalKanbanApprove = clsGeneral.getIntervalApprove(GB)
                    IntervalAkanbanApprove = TimeSpan.FromHours(dsGetIntervalKanbanApprove.Tables(0).Rows(0)("KanbanApprovalHour"))
                    If Format(Now, "HH:mm:ss") >= Format(pDateNeedApproveKanban + IntervalAkanbanApprove, "HH:mm:ss") Then
                        log.WriteToProcessLog(Date.Now, "AutoApproveKanban", "Process Auto Approve Kanban [" & KAKanbanno & "], Supplier [" & KASupplier & "], Affiliate [" & KAAff & "]")
                        uf_ApproveKanban(KAKanbanDate, KAAff, KASupplier, KADeliveryLocation, KAKanbanno, errMsg)

                        If errMsg <> "" Then
                            log.WriteToErrorLog(pScreenName, "Process Auto Approve Kanban [" & KAAff & "-" & KAKanbanno & "-" & KASupplier & "] STOPPED, because " & errMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Auto Approve Kanban [" & KAAff & "-" & KAKanbanno & "-" & KASupplier & "] STOPPED, because " & errMsg)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Auto Approve Kanban [" & KAAff & "-" & KAKanbanno & "-" & KASupplier & "] STOPPED, because " & errMsg, LogName)
                            LogName.Refresh()
                            errMsg = ""
                            GoTo keluar
                        End If

                        'If sendEmailtoPASI(GB, KAKanbanno, KAAff, KASupplier, KADeliveryLocation, KAKanbanDate, errMsg) = False Then
                        'Else
                        '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC PASI. Kanban [" & KAKanbanno & "], Supplier [" & KASupplier & "], Affiliate [" & KAAff & "] ok.")
                        'End If

                        'If sendEmailtoAffiliate(GB, KAKanbanno, KAAff, KASupplier, KADeliveryLocation, KAKanbanDate, errMsg) = False Then
                        'Else
                        '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to CC Affiliate. Kanban [" & KAKanbanno & "], Supplier [" & KASupplier & "], Affiliate [" & KAAff & "] ok.")
                        'End If

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Auto Approve Kanban [" & KAKanbanno & "], Supplier [" & KASupplier & "], Affiliate [" & KAAff & "] ok", LogName)
                        LogName.Refresh()
                    End If
keluar:
                Next
            Else
                errMsg = "-"
                ErrSummary = "-"
                Exit Try
            End If
        Catch ex As Exception
            errMsg = "Kanban No [" & KAKanbanno & "], Supplier [" & KASupplier & "], Affiliate [" & KAAff & "] " & ex.Message
            ErrSummary = "Kanban No [" & KAKanbanno & "], Supplier [" & KASupplier & "], Affiliate [" & KAAff & "] " & ex.Message
        Finally
            If Not dsGetKanbanApprove Is Nothing Then
                dsGetKanbanApprove = Nothing
            End If
            If Not dsGetIntervalKanbanApprove Is Nothing Then
                dsGetIntervalKanbanApprove = Nothing
            End If
        End Try
    End Sub

    Shared Function getKanbanApprove(ByVal GB As GlobalSetting.clsGlobal) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT KanbanNo, KanbanDate, AffiliateID, SupplierID,DeliveryLocationCode,AffiliateApproveUser, AffiliateApproveDate, SupplierApproveuser, SupplierApproveDate, ExcelCls" & vbCrLf & _
                  " FROM Kanban_Master where Isnull(SupplierApproveDate,'') = '' and ExcelCls = 2 and isnull(AffiliateApproveUser,'') <> ''  " & vbCrLf & _
                  "  AND isnull(AffiliateApproveDate,'') <> '' and isnull(supplierApproveUser,'') = '' " & vbCrLf & _
                  " GROUP BY KanbanNo, KanbanDate, AffiliateID, SupplierID,DeliveryLocationCode,AffiliateApproveUser, AffiliateApproveDate, SupplierApproveuser, SupplierApproveDate, ExcelCls "

        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Sub uf_ApproveKanban(ByVal pkanbandate As Date, ByVal paff As String, ByVal psupp As String, ByVal pdel As String, ByVal pKanbaNo As String, ByRef errMsg As String)
        Dim ls_sql As String
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_sql = " Update Kanban_Master " & vbCrLf & _
                      " SET SupplierApproveDate = GetDate(), SupplierApproveUSer = 'AUTO APPROVED' " & vbCrLf & _
                      " WHERE KanbanDate='" & Format(pkanbandate, "yyyy-MM-dd") & "' AND AffiliateID='" & Trim(paff) & "' AND SupplierID='" & Trim(psupp) & "' AND DeliveryLocationCode='" & Trim(pdel) & "' and KanbanNo = '" & Trim(pKanbaNo) & "'"

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Auto Approve KanbanNo [" & pKanbaNo & "] STOPPED, because " & ex.Message
        End Try
        
    End Sub

    Shared Function sendEmailtoPASI(ByVal GB As GlobalSetting.clsGlobal, ByVal pKanbanNo As String, ByVal pAffiliate As String, _
                                    ByVal pSupplier As String, ByVal pDeliveryLocation As String, _
                                    ByVal pKanbanDate As Date, ByRef errMsg As String) As Boolean
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
                errMsg = "Process Auto Approve Kanban [" & pKanbanNo & "-" & pAffiliate & "-" & pSupplier & "] Notification to PASI STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Auto Approve Kanban [" & pKanbanNo & "-" & pAffiliate & "-" & pSupplier & "] Notification to PASI STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerNamePASI & "/AffKanban/AffKanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pKanbanDate.Date) & "&t1=" & clsNotification.EncryptURL(pSupplier) & _
                                       "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/AffKanban/AffKanbanList.aspx")


            ls_Subject = "AUTO APPROVE KANBAN NO : " & pKanbanNo & "-" & pAffiliate & "-" & pSupplier
            ls_Body = clsNotification.GetNotification("32", ls_URl, , pKanbanNo)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoPASI = False
                Exit Function
            End If

            sendEmailtoPASI = True

        Catch ex As Exception
            sendEmailtoPASI = False
            errMsg = "Process Send Kanban [" & pAffiliate & "-" & pKanbanNo & "-" & pSupplier & "] Notification to PASI STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Function sendEmailtoAffiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal pKanbanNo As String, ByVal pAffiliate As String, _
                                         ByVal pSupplier As String, ByVal pDeliveryLocation As String, _
                                         ByVal pKanbanDate As Date, ByRef errMsg As String) As Boolean
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

            If receiptEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Auto Approve Kanban [" & pKanbanNo & "-" & pAffiliate & "-" & pSupplier & "] Notification to Affiliate STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If fromEmail = "" Then
                sendEmailtoAffiliate = False
                errMsg = "Process Auto Approve Kanban [" & pKanbanNo & "-" & pAffiliate & "-" & pSupplier & "] Notification to Affiliate STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerName & "/Kanban/KanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pKanbanDate.Date) & "&t1=" & clsNotification.EncryptURL(pSupplier) & _
                           "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/Kanban/KanbanList.aspx")

            ls_Subject = "AUTO APPROVE KANBAN NO : " & pKanbanNo & "-" & pAffiliate & "-" & pSupplier
            ls_Body = clsNotification.GetNotification("32", ls_URl, , pKanbanNo)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoAffiliate = False
                Exit Function
            End If

            sendEmailtoAffiliate = True

        Catch ex As Exception
            sendEmailtoAffiliate = False
            errMsg = "Process Send Kanban [" & pAffiliate & "-" & pKanbanNo & "-" & pSupplier & "] Notification to Affiliate STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

End Class
