Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsPOAutoApprove
    Shared Sub up_AutoApprovePODomestic(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

        Dim ls_SQL As String = ""
        Dim pAffiliate As String = ""
        Dim pAffiliateName As String = ""
        Dim pSupplier As String = ""
        Dim pPONo As String = ""
        Dim pPeriod As String = ""
        Dim pDeliveryBy As String = ""
        Dim pShipCls As String = ""
        Dim pCommercialCls As String = ""
        Dim pRemarks As String = ""
        Dim pFinalApproval As String = ""

        Dim dsGetPOApprove As New DataSet
        Dim dsGetIntervalApprove As New DataSet

        Dim intervalApprove As TimeSpan
        Dim pDateNeedApprove As Date
        Dim pSendToSupplierDate As String

        Try

            log.WriteToProcessLog(Date.Now, pScreenName, "Get data PO")

            dsGetPOApprove = getPONoApprove(GB)

            If dsGetPOApprove.Tables(0).Rows.Count > 0 Then
                For iRow = 0 To dsGetPOApprove.Tables(0).Rows.Count - 1
                    pPeriod = If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("Period")), "", dsGetPOApprove.Tables(0).Rows(iRow)("Period"))
                    pPONo = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("PONo")), "", dsGetPOApprove.Tables(0).Rows(iRow)("PONo")))
                    pSupplier = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("SupplierID")), "", dsGetPOApprove.Tables(0).Rows(iRow)("SupplierID")))
                    pAffiliate = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("AffiliateID")), "", dsGetPOApprove.Tables(0).Rows(iRow)("AffiliateID")))
                    pAffiliateName = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("AffiliateName")), "", dsGetPOApprove.Tables(0).Rows(iRow)("AffiliateName")))
                    pDeliveryBy = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("DeliveryByPASICls")), "", dsGetPOApprove.Tables(0).Rows(iRow)("DeliveryByPASICls")))
                    pShipCls = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("ShipCls")), "", dsGetPOApprove.Tables(0).Rows(iRow)("ShipCls")))

                    pCommercialCls = If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("CommercialCls")), "", dsGetPOApprove.Tables(0).Rows(iRow)("CommercialCls"))
                    If pCommercialCls = "1" Then
                        pCommercialCls = "YES"
                    Else
                        pCommercialCls = "NO"
                    End If

                    pRemarks = ""

                    pSendToSupplierDate = If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("PASISendAffiliateDate")), "", dsGetPOApprove.Tables(0).Rows(iRow)("PASISendAffiliateDate"))
                    pDateNeedApprove = FormatDateTime(pSendToSupplierDate, DateFormat.GeneralDate)

                    'calculate interval date
                    dsGetIntervalApprove = clsGeneral.getIntervalApprove(GB)
                    intervalApprove = TimeSpan.FromDays(dsGetIntervalApprove.Tables(0).Rows(0)("POApprovalDate"))

                    If Format(Now, "yyyy-MM-dd") >= Format(pDateNeedApprove + intervalApprove, "yyyy-MM-dd") Then
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")

                        '0000 Perlu di aktifin
                        uf_Approve(pPONo, pAffiliate, pSupplier, errMsg)

                        If errMsg <> "" Then
                            log.WriteToErrorLog(pScreenName, "Process Auto Approve PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & errMsg, 9999, GlobalSetting.clsLog.ErrSeverity.ERR)
                            log.WriteToProcessLog(Date.Now, pScreenName, "Process Auto Approve PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & errMsg)

                            clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Auto Approve PO [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & errMsg, LogName)
                            LogName.Refresh()
                            errMsg = ""
                            GoTo keluar
                        End If

                        If sendEmailtoSupplier(GB, pPONo, pAffiliate, pSupplier, errMsg) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        End If

                        'If sendEmailtoPASI(GB, pPeriod, pPONo, pAffiliate, pAffiliateName, pSupplier, pDeliveryBy, pShipCls, pCommercialCls, pRemarks, pFinalApproval, errMsg) = False Then
                        'Else
                        '    log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to PASI. Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        'End If

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
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
            errMsg = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & ex.Message
            ErrSummary = "PONo [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] " & ex.Message
        Finally
            If Not dsGetPOApprove Is Nothing Then
                dsGetPOApprove.Dispose()
            End If
            If Not dsGetIntervalApprove Is Nothing Then
                dsGetIntervalApprove.Dispose()
            End If
        End Try

    End Sub

    Shared Function getPONoApprove(ByVal GB As GlobalSetting.clsGlobal) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT * FROM PO_Master POM " & vbCrLf & _                  
                  " LEFT JOIN dbo.MS_Affiliate MA ON POM.AffiliateID=MA.AffiliateID " & vbCrLf & _
                  " WHERE PASISendAffiliateDate is not null " & vbCrLf & _
                  " and (SupplierApproveDate is null and SupplierApprovePendingDate is null and SupplierUnApproveDate is null)" & vbCrLf & _
                  " and not exists ( " & vbCrLf & _
                  "         select * from PO_MasterUpload PODU where POM.PONo = PODU.PONo and POM.AffiliateID = PODU.AffiliateID AND POM.SupplierID=PODU.SupplierID" & vbCrLf & _
                  ") order by PASISendAffiliateDate"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Sub uf_Approve(ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " insert into PO_MasterUpload " & vbCrLf & _
                          " select PONo, AffiliateID, SupplierID, 'Auto Approved by System' Remarks, GETDATE(), 'AUTO APPROVED', GETDATE(), 'AUTO APPROVED' from Affiliate_Master " & vbCrLf & _
                          " where pono='" & Trim(pPONo) & "' and AffiliateID = '" & Trim(pAffiliate) & "' AND SupplierID='" & Trim(pSupplier) & "' "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " insert into PO_DetailUpload " & vbCrLf & _
                         " 	([PONo] ,[AffiliateID] ,[SupplierID], [PartNo] ,[POQty] ,[POQtyOld] " & vbCrLf & _
                         "            ,[DeliveryD1] ,[DeliveryD1Old] ,[DeliveryD2] ,[DeliveryD2Old] " & vbCrLf & _
                         "            ,[DeliveryD3] ,[DeliveryD3Old] ,[DeliveryD4] ,[DeliveryD4Old] " & vbCrLf & _
                         "            ,[DeliveryD5] ,[DeliveryD5Old] ,[DeliveryD6] ,[DeliveryD6Old] " & vbCrLf & _
                         "            ,[DeliveryD7] ,[DeliveryD7Old] ,[DeliveryD8] ,[DeliveryD8Old] " & vbCrLf & _
                         "            ,[DeliveryD9] ,[DeliveryD9Old] ,[DeliveryD10] ,[DeliveryD10Old] " & vbCrLf & _
                         "            ,[DeliveryD11] ,[DeliveryD11Old] ,[DeliveryD12] ,[DeliveryD12Old] " & vbCrLf & _
                         "            ,[DeliveryD13] ,[DeliveryD13Old] ,[DeliveryD14] ,[DeliveryD14Old] " & vbCrLf & _
                         "            ,[DeliveryD15] ,[DeliveryD15Old] ,[DeliveryD16] ,[DeliveryD16Old] " & vbCrLf & _
                         "            ,[DeliveryD17] ,[DeliveryD17Old] ,[DeliveryD18] ,[DeliveryD18Old] " & vbCrLf & _
                         "            ,[DeliveryD19] ,[DeliveryD19Old] ,[DeliveryD20] ,[DeliveryD20Old] " & vbCrLf

                ls_SQL = ls_SQL + "            ,[DeliveryD21] ,[DeliveryD21Old] ,[DeliveryD22] ,[DeliveryD22Old] " & vbCrLf & _
                                  "            ,[DeliveryD23] ,[DeliveryD23Old] ,[DeliveryD24] ,[DeliveryD24Old] " & vbCrLf & _
                                  "            ,[DeliveryD25] ,[DeliveryD25Old] ,[DeliveryD26] ,[DeliveryD26Old] " & vbCrLf & _
                                  "            ,[DeliveryD27] ,[DeliveryD27Old] ,[DeliveryD28] ,[DeliveryD28Old] " & vbCrLf & _
                                  "            ,[DeliveryD29] ,[DeliveryD29Old] ,[DeliveryD30] ,[DeliveryD30Old] " & vbCrLf & _
                                  "            ,[DeliveryD31] ,[DeliveryD31Old] ,[EntryDate] ,[EntryUser] " & vbCrLf & _
                                  "            ,[UpdateDate] ,[UpdateUser]) " & vbCrLf & _
                                  " select a.[PONo]      ,a.[AffiliateID]      ,a.[SupplierID],  a.[PartNo]      ,[POQty]       " & vbCrLf & _
                                  " 	  ,[POQtyOld]      ,[DeliveryD1]      ,[DeliveryD1Old] " & vbCrLf & _
                                  "       ,[DeliveryD2]      ,[DeliveryD2Old]      ,[DeliveryD3]      ,[DeliveryD3Old] " & vbCrLf & _
                                  "       ,[DeliveryD4]      ,[DeliveryD4Old]      ,[DeliveryD5]      ,[DeliveryD5Old] " & vbCrLf & _
                                  "       ,[DeliveryD6]      ,[DeliveryD6Old]	   ,[DeliveryD7]      ,[DeliveryD7Old]	   " & vbCrLf & _
                                  " 	  ,[DeliveryD8]      ,[DeliveryD8Old]      ,[DeliveryD9]      ,[DeliveryD9Old] " & vbCrLf & _
                                  "       ,[DeliveryD10]      ,[DeliveryD10Old]      ,[DeliveryD11]      ,[DeliveryD11Old] " & vbCrLf & _
                                  "       ,[DeliveryD12]      ,[DeliveryD12Old]      ,[DeliveryD13]      ,[DeliveryD13Old] " & vbCrLf & _
                                  "       ,[DeliveryD14]      ,[DeliveryD14Old]      ,[DeliveryD15]      ,[DeliveryD15Old] " & vbCrLf & _
                                  "       ,[DeliveryD16]      ,[DeliveryD16Old]      ,[DeliveryD17]      ,[DeliveryD17Old] " & vbCrLf

                ls_SQL = ls_SQL + "       ,[DeliveryD18]      ,[DeliveryD18Old]      ,[DeliveryD19]      ,[DeliveryD19Old] " & vbCrLf & _
                                  "       ,[DeliveryD20]      ,[DeliveryD20Old]      ,[DeliveryD21]      ,[DeliveryD21Old]       " & vbCrLf & _
                                  " 	  ,[DeliveryD22]      ,[DeliveryD22Old]      ,[DeliveryD23]      ,[DeliveryD23Old] " & vbCrLf & _
                                  "       ,[DeliveryD24]      ,[DeliveryD24Old]      ,[DeliveryD25]      ,[DeliveryD25Old] " & vbCrLf & _
                                  "       ,[DeliveryD26]      ,[DeliveryD26Old]      ,[DeliveryD27]      ,[DeliveryD27Old] " & vbCrLf & _
                                  "       ,[DeliveryD28]      ,[DeliveryD28Old]      ,[DeliveryD29]      ,[DeliveryD29Old]       " & vbCrLf & _
                                  " 	  ,[DeliveryD30]      ,[DeliveryD30Old]      ,[DeliveryD31]      ,[DeliveryD31Old] " & vbCrLf & _
                                  "       ,GETDATE()      ,'AUTO APPROVED'      ,GETDATE()      ,'AUTO APPROVED' " & vbCrLf & _
                                  " 	  from Affiliate_Detail a left join (select PONo, AffiliateID, SupplierID, PartNo, KanbanCls from PO_Detail) b on a.AffiliateID = b.AffiliateID " & vbCrLf & _
                                  "             and a.SupplierID = b.SupplierID and a.PONO = b.PONo and a.PartNo = b.PartNo" & vbCrLf & _
                                  " where a.pono='" & Trim(pPONo) & "' and a.AffiliateID = '" & Trim(pAffiliate) & "' AND a.SupplierID='" & Trim(pSupplier) & "' "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using

            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " Update PO_Master set SupplierApproveDate = getdate(), SupplierApproveUser = 'AUTO APPROVED'" & vbCrLf & _
                         " WHERE AffiliateID = '" & Trim(pAffiliate) & "' and PONo = '" & Trim(pPONo) & "' and SupplierID = '" & Trim(pSupplier) & "'" & vbCrLf

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Auto Approve PONo [" & pPONo & "] STOPPED, because " & ex.Message
        End Try
    End Sub

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
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
                errMsg = "Process Auto Approval [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Auto Approval [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Auto Approval PONo: " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("15", "", pPONo.Trim & "-" & pSupplier.Trim)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, pSupplier, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Auto Approval [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function

    Shared Function sendEmailtoPASI(ByVal GB As GlobalSetting.clsGlobal, ByVal pPeriod As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pAffiliateName As String, ByVal pSupplier As String, ByVal pDelivBy As String, ByVal pShipCls As String, ByVal pCommercialCls As String, ByVal pRemarks As String, ByVal pFinalApproveCls As String, ByRef errMsg As String) As Boolean
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

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If fromEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Auto Approval [" & pPONo & "-" & pSupplier & "] STOPPED, because Mailer's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoPASI = False
                errMsg = "Process Auto Approval [" & pPONo & "-" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_URl = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderAppDetail.aspx?id2=" & clsNotification.EncryptURL(pPONo.Trim) & _
                    "&t1=" & clsNotification.EncryptURL(pAffiliate.Trim) & "&t2=" & clsNotification.EncryptURL(pAffiliateName.Trim) & _
                    "&t3=" & clsNotification.EncryptURL(pPeriod) & "&t4=" & clsNotification.EncryptURL(pSupplier.Trim) & _
                    "&t5=" & clsNotification.EncryptURL(pRemarks.Trim) & "&t6=" & clsNotification.EncryptURL(pFinalApproveCls.Trim) & _
                    "&t7=" & clsNotification.EncryptURL(pDelivBy.Trim) & "&t8=" & clsNotification.EncryptURL(pShipCls.Trim) & _
                    "&t9=" & clsNotification.EncryptURL(pCommercialCls.Trim) & "&t10=" & clsNotification.EncryptURL("") & _
                    "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderAppList.aspx")
            ls_Subject = "Auto Approval PONo: " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("15", ls_URl, pPONo.Trim & "-" & pSupplier.Trim)

            If clsGeneral.sendEmail(GB, fromEmail, receiptEmail, receiptCCEmail, "", ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoPASI = False
                Exit Function
            End If

            sendEmailtoPASI = True

        Catch ex As Exception
            sendEmailtoPASI = False
            errMsg = "Process Auto Approval [" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        Finally
            If Not dsEmail Is Nothing Then
                dsEmail.Dispose()
            End If
        End Try
    End Function
End Class
