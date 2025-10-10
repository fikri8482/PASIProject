Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net

Public Class clsAutoApprovePOExport
    Shared Sub up_AutoApprovePOExport(ByVal cfg As GlobalSetting.clsConfig,
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
        Dim pOrderNo As String = ""
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

            dsGetPOApprove = getPONoApprovePOEXPORT(GB)

            If dsGetPOApprove.Tables(0).Rows.Count > 0 Then
                For iRow = 0 To dsGetPOApprove.Tables(0).Rows.Count - 1
                    pPONo = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("PONo")), "", dsGetPOApprove.Tables(0).Rows(iRow)("PONo")))
                    pSupplier = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("SupplierID")), "", dsGetPOApprove.Tables(0).Rows(iRow)("SupplierID")))
                    pAffiliate = Trim(If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("AffiliateID")), "", dsGetPOApprove.Tables(0).Rows(iRow)("AffiliateID")))
                    pOrderNo = If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("OrderNo1")), "", dsGetPOApprove.Tables(0).Rows(iRow)("OrderNo1"))
                    pSendToSupplierDate = If(IsDBNull(dsGetPOApprove.Tables(0).Rows(iRow)("PASISendToSupplierDate")), "", dsGetPOApprove.Tables(0).Rows(iRow)("PASISendToSupplierDate"))
                    pDateNeedApprove = FormatDateTime(pSendToSupplierDate, DateFormat.GeneralDate)

                    'calculate interval date
                    dsGetIntervalApprove = clsGeneral.getIntervalApproveExport(GB)
                    intervalApprove = TimeSpan.FromDays(dsGetIntervalApprove.Tables(0).Rows(0)("POApprovalDate"))

                    If Format(Now, "yyyy-MM-dd") >= Format(pDateNeedApprove + intervalApprove, "yyyy-MM-dd") Then
                        log.WriteToProcessLog(Date.Now, pScreenName, "Process Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "]")

                        uf_Approve(pPONo, pOrderNo, pAffiliate, pSupplier, errMsg)

                        If sendEmailtoSupplier(GB, pPONo, pAffiliate, pSupplier, errMsg) = False Then
                            Exit Try
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to Supplier. Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        End If

                        If sendEmailtoPASI(GB, "", pPONo, pAffiliate, pAffiliateName, pSupplier, pDeliveryBy, pShipCls, pCommercialCls, pRemarks, pFinalApproval, errMsg) = False Then
                        Else
                            log.WriteToProcessLog(Date.Now, pScreenName, "Send Email to PASI. Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok.")
                        End If

                        clsGeneral.up_displayLog(GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg, "Process Auto Approve PO [" & pPONo & "], Supplier [" & pSupplier & "], Affiliate [" & pAffiliate & "] ok", LogName)
                        LogName.Refresh()
                    End If
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

    Shared Function getPONoApprovePOEXPORT(ByVal GB As GlobalSetting.clsGlobal) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT * FROM PO_Master_export  PORM  " & vbCrLf & _
                  "  LEFT JOIN dbo.MS_Affiliate MA ON PORM.AffiliateID=MA.AffiliateID  " & vbCrLf & _
                  "  WHERE PASISendToSupplierDate is not null  " & vbCrLf & _
                  "  and (SupplierApproveDate is null and SupplierApprovePartialDate is null and SupplierUnApproveDate is null)  " & vbCrLf & _
                  "  and not exists (  " & vbCrLf & _
                  "         select * from PO_MasterUpload_Export PODU where PORM.PONo = PODU.PONo and PORM.OrderNo1 = PODU.OrderNo1 and PORM.AffiliateID = PODU.AffiliateID AND PORM.SupplierID=PODU.SupplierID " & vbCrLf & _
                  "  )  " & vbCrLf & _
                  "  order by PASISendToSupplierDate " & vbCrLf & _
                  "  "
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Sub uf_Approve(ByVal pPONo As String, ByVal pOrderNo1 As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String)
        Dim ls_SQL As String = ""
        Dim cfg As New GlobalSetting.clsConfig

        Try
            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " INSERT INTO PO_MasterUpload_Export " & vbCrLf & _
                       " SELECT POno, AffiliateID, SupplierID, ForwarderID, OrderNo1, ETDVendor1, '', GETDATE(), 'AUTO APP', GETDATE(), 'AUTO APP' " & vbCrLf & _
                       " FROM PO_Master_Export " & vbCrLf & _
                       " WHERE PONo='" & Trim(pPONo) & "' AND OrderNo1 = '" & Trim(pOrderNo1) & "' AND AffiliateID='" & Trim(pAffiliate) & "' AND SupplierID='" & Trim(pSupplier) & "' "

                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
                sqlConn.Close()
            End Using

            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " INSERT INTO PO_DetailUpload_Export " & vbCrLf & _
                      " SELECT POno, AffiliateID, SupplierID, ForwarderID, OrderNo1,PartNo, week1, Week1, totalPOQty,totalPOQty, GETDATE(), 'AUTO APP', GETDATE(), 'AUTO APP' " & vbCrLf & _
                      " FROM PO_Detail_Export " & vbCrLf & _
                      " WHERE PONo='" & Trim(pPONo) & "' AND OrderNo1 = '" & Trim(pOrderNo1) & "' AND AffiliateID='" & Trim(pAffiliate) & "' AND SupplierID='" & Trim(pSupplier) & "' "

                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
                sqlConn.Close()
            End Using

            Using sqlConn As New SqlConnection(cfg.ConnectionString)
                sqlConn.Open()
                ls_SQL = " Update PO_Master_Export set SupplierApproveDate = getdate(), SupplierApproveUser = 'AUTO APPROVED'" & vbCrLf & _
                         " WHERE PONo='" & Trim(pPONo) & "' AND OrderNo1 = '" & Trim(pOrderNo1) & "' AND AffiliateID='" & Trim(pAffiliate) & "' AND SupplierID='" & Trim(pSupplier) & "' "
                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            errMsg = "Process Auto Approve PONo [" & pPONo & "] STOPPED, because " & ex.Message
        End Try
    End Sub

    Shared Function sendEmailtoSupplier(ByVal GB As GlobalSetting.clsGlobal, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByRef errMsg As String) As Boolean
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
                errMsg = "Process Auto Approval [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            If receiptEmail = "" Then
                sendEmailtoSupplier = False
                errMsg = "Process Auto Approval [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because Recipient's e-mail address is not found"
                Exit Function
            End If

            ls_Subject = "Auto Approval (Export) PONo: " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("17", "", pPONo.Trim & "-" & pSupplier.Trim)

            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg, ls_Attachment) = False Then
                sendEmailtoSupplier = False
                Exit Function
            End If

            sendEmailtoSupplier = True

        Catch ex As Exception
            sendEmailtoSupplier = False
            errMsg = "Process Auto Approval [" & pAffiliate & "-" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function

    Shared Function sendEmailtoPASI(ByVal GB As GlobalSetting.clsGlobal, ByVal pPeriod As String, ByVal pPONo As String, ByVal pAffiliate As String, ByVal pAffiliateName As String, ByVal pSupplier As String, ByVal pDelivBy As String, ByVal pShipCls As String, ByVal pCommercialCls As String, ByVal pRemarks As String, ByVal pFinalApproveCls As String, ByRef errMsg As String) As Boolean
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
            dsEmail = clsGeneral.getEmailAddressPASI(GB, "", "PASI", "", "", "AffiliatePOCC", "AffiliatePOTO", "AffiliatePOTO", errMsg)

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

            'ls_URl = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderAppDetail.aspx?id2=" & clsNotification.EncryptURL(pPONo.Trim) & _
            '        "&t1=" & clsNotification.EncryptURL(pAffiliate.Trim) & "&t2=" & clsNotification.EncryptURL(pAffiliateName.Trim) & _
            '        "&t3=" & clsNotification.EncryptURL(pPeriod) & "&t4=" & clsNotification.EncryptURL(pSupplier.Trim) & _
            '        "&t5=" & clsNotification.EncryptURL(pRemarks.Trim) & "&t6=" & clsNotification.EncryptURL(pFinalApproveCls.Trim) & _
            '        "&t7=" & clsNotification.EncryptURL(pDelivBy.Trim) & "&t8=" & clsNotification.EncryptURL(pShipCls.Trim) & _
            '        "&t9=" & clsNotification.EncryptURL(pCommercialCls.Trim) & "&t10=" & clsNotification.EncryptURL("") & _
            '        "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderAppList.aspx")

            ls_Subject = "Auto Approval PONo: " & pPONo.Trim & "-" & pSupplier.Trim
            ls_Body = clsNotification.GetNotification("17", "", pPONo.Trim & "-" & pSupplier.Trim)

            If clsGeneral.sendEmailExport(GB, fromEmail, receiptEmail, receiptCCEmail, ls_Subject, ls_Body, errMsg) = False Then
                sendEmailtoPASI = False
                Exit Function
            End If

            sendEmailtoPASI = True

        Catch ex As Exception
            sendEmailtoPASI = False
            errMsg = "Process Auto Approval [" & pPONo & "-" & pSupplier & "] STOPPED, because " & ex.Message
            Exit Function
        End Try
    End Function
End Class
