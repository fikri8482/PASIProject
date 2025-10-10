Public Class clsPORevAutoApprove
    Shared Sub up_AutoApprovePORevDomestic(ByVal cfg As GlobalSetting.clsConfig,
                              ByVal log As GlobalSetting.clsLog,
                              ByVal GB As GlobalSetting.clsGlobal,
                              ByVal LogName As RichTextBox,
                              ByVal pAtttacment As String,
                              ByVal pResult As String,
                              ByVal pScreenName As String,
                              Optional ByRef errMsg As String = "",
                              Optional ByRef ErrSummary As String = "")

    End Sub

    'Private Sub AutoApproveRev()
    '    dsGetPOApproveRev = getPONoApproveRev()
    '    If dsGetPOApproveRev.Tables(0).Rows.Count > 0 Then
    '        For iRow = 0 To dsGetPOApproveRev.Tables(0).Rows.Count - 1
    '            pPONOAppR = dsGetPOApproveRev.Tables(0).Rows(iRow)("PONo")
    '            'pShipAppRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("ShipCls")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("ShipCls"))
    '            'pCommercialRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("CommercialCls")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("CommercialCls"))
    '            pAffiliateIDRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("AffiliateID")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("AffiliateID"))
    '            'pAffiliateNameRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("AffiliateName")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("AffiliateName"))
    '            'pPeriodRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("Period")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("Period"))
    '            'pSupplierIDRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierID")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierID"))
    '            'pRemarksRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("Remarks")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("Remarks"))
    '            'pFinalAppRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("FinalApproveCls")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("FinalApproveCls"))
    '            'pDelAppRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("DeliveryByPASICls")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("DeliveryByPASICls"))
    '            pPORevNOApp = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("PORevNO")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("PORevNO"))
    '            'pSupplierApproveDateRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierApproveDate")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierApproveDate"))
    '            'pSupplierApprovePendingDateRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierApprovePendingDate")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierApprovePendingDate"))
    '            'pSupplierUnApproveDateRev = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierUnApproveDate")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("SupplierUnApproveDate"))
    '            pSendToSupplierDate = If(IsDBNull(dsGetPOApproveRev.Tables(0).Rows(iRow)("PASISendAffiliateDate")), "", dsGetPOApproveRev.Tables(0).Rows(iRow)("PASISendAffiliateDate"))

    '            pDateNeedApproveRev = FormatDateTime(pSendToSupplierDate, DateFormat.GeneralDate)

    '            'If pSupplierUnApproveDateRev <> "" Then
    '            '    pDateNeedApproveRev = pSupplierUnApproveDateRev 'pDateNeedApprove = pSupplierApproveDateRev
    '            'End If
    '            'If pSupplierApprovePendingDateRev <> "" Then
    '            '    pDateNeedApproveRev = pSupplierApprovePendingDateRev
    '            'End If
    '            'If pSupplierApproveDateRev <> "" Then
    '            '    pDateNeedApproveRev = pSupplierApproveDateRev 'pDateNeedApprove = pSupplierUnApproveDateRev
    '            'End If

    '            dsGetIntervalApproveRev = getIntervalApprove()
    '            intervalApproveRev = TimeSpan.FromDays(dsGetIntervalApproveRev.Tables(0).Rows(0)("PORevisionApprovalDate"))
    '            If Format(Now, "yyyy-MM-dd") >= Format(pDateNeedApproveRev + intervalApproveRev, "yyyy-MM-dd") Then
    '                uf_ApproveRev(pPONOAppR, pPORevNOApp, pAffiliateIDRev, pSupplierIDRev)
    '                'sendEmailAppRev()
    '                sendEmailAppRevtoAff()
    '                sendEmailAppRevtoccPASI()
    '            End If
    '        Next
    '    End If
    'End Sub

    'Private Sub uf_ApproveRev(ByVal pPONOAppR As String, ByVal pPORevNOApp As String, ByVal pAffiliateIDRev As String, ByVal pSupplierIDRev As String)
    '    Dim ls_sql As String

    '    MdlConn.ReadConnection()

    '    Using sqlConn As New SqlConnection(MdlConn.uf_GetConString)
    '        sqlConn.Open()
    '        ls_sql = " insert into PORev_MasterUpload " & vbCrLf & _
    '              " select PORevNo, PONo, AffiliateID, SupplierID, SeqNo, 'Auto Approved by System' Remarks, GETDATE(), 'AUTO APPROVED', GETDATE(), 'AUTO APPROVED' from AffiliateRev_Master " & vbCrLf & _
    '              " WHERE PORevNo='" & Trim(pPORevNOApp) & "' AND PONo='" & Trim(pPONOAppR) & "' AND AffiliateID='" & Trim(pAffiliateIDRev) & "' AND SupplierID='" & Trim(pSupplierIDRev) & "' "

    '        Dim SqlComm As New SqlCommand(ls_sql, sqlConn)
    '        SqlComm.ExecuteNonQuery()
    '        SqlComm.Dispose()
    '        sqlConn.Close()
    '    End Using

    '    Using sqlConn As New SqlConnection(MdlConn.uf_GetConString)
    '        sqlConn.Open()
    '        ls_sql = " insert into PORev_DetailUpload " & vbCrLf & _
    '              " select [PORevNo], [PONo]      ,[AffiliateID]      ,[SupplierID]      ,[PartNo]      ,[SeqNo],  [DifferenceCls]      ,[KanbanCls]      ,[Maker]      ,[POQty]       " & vbCrLf & _
    '              " 	  ,[POQtyOld]      ,[CurrCls]      ,[Price]      ,[Amount]      ,[DeliveryD1]      ,[DeliveryD1Old] " & vbCrLf & _
    '              "       ,[DeliveryD2]      ,[DeliveryD2Old]      ,[DeliveryD3]      ,[DeliveryD3Old] " & vbCrLf & _
    '              "       ,[DeliveryD4]      ,[DeliveryD4Old]      ,[DeliveryD5]      ,[DeliveryD5Old] " & vbCrLf & _
    '              "       ,[DeliveryD6]      ,[DeliveryD6Old]	   ,[DeliveryD7]      ,[DeliveryD7Old]	   " & vbCrLf & _
    '              " 	  ,[DeliveryD8]      ,[DeliveryD8Old]      ,[DeliveryD9]      ,[DeliveryD9Old] " & vbCrLf & _
    '              "       ,[DeliveryD10]      ,[DeliveryD10Old]      ,[DeliveryD11]      ,[DeliveryD11Old] " & vbCrLf & _
    '              "       ,[DeliveryD12]      ,[DeliveryD12Old]      ,[DeliveryD13]      ,[DeliveryD13Old] " & vbCrLf & _
    '              "       ,[DeliveryD14]      ,[DeliveryD14Old]      ,[DeliveryD15]      ,[DeliveryD15Old] " & vbCrLf & _
    '              "       ,[DeliveryD16]      ,[DeliveryD16Old]      ,[DeliveryD17]      ,[DeliveryD17Old] "

    '        ls_sql = ls_sql + "       ,[DeliveryD18]      ,[DeliveryD18Old]      ,[DeliveryD19]      ,[DeliveryD19Old] " & vbCrLf & _
    '                          "       ,[DeliveryD20]      ,[DeliveryD20Old]      ,[DeliveryD21]      ,[DeliveryD21Old]       " & vbCrLf & _
    '                          " 	  ,[DeliveryD22]      ,[DeliveryD22Old]      ,[DeliveryD23]      ,[DeliveryD23Old] " & vbCrLf & _
    '                          "       ,[DeliveryD24]      ,[DeliveryD24Old]      ,[DeliveryD25]      ,[DeliveryD25Old] " & vbCrLf & _
    '                          "       ,[DeliveryD26]      ,[DeliveryD26Old]      ,[DeliveryD27]      ,[DeliveryD27Old] " & vbCrLf & _
    '                          "       ,[DeliveryD28]      ,[DeliveryD28Old]      ,[DeliveryD29]      ,[DeliveryD29Old]       " & vbCrLf & _
    '                          " 	  ,[DeliveryD30]      ,[DeliveryD30Old]      ,[DeliveryD31]      ,[DeliveryD31Old] " & vbCrLf & _
    '                          "       ,GETDATE()      ,'AUTO APPROVED'      ,GETDATE()      ,'AUTO APPROVED' " & vbCrLf & _
    '                          " 	  from AffiliateRev_Detail " & vbCrLf & _
    '                          " WHERE PORevNo='" & Trim(pPORevNOApp) & "' AND PONo='" & Trim(pPONOAppR) & "' AND AffiliateID='" & Trim(pAffiliateIDRev) & "' AND SupplierID='" & Trim(pSupplierIDRev) & "' " & vbCrLf

    '        Dim SqlComm As New SqlCommand(ls_sql, sqlConn)
    '        SqlComm.ExecuteNonQuery()
    '        SqlComm.Dispose()
    '        sqlConn.Close()
    '    End Using

    '    Using sqlConn As New SqlConnection(MdlConn.uf_GetConString)
    '        sqlConn.Open()
    '        ls_sql = " Update PORev_Master set SupplierApproveDate = getdate(), SupplierApproveUser = 'AUTO APPROVED'" & vbCrLf & _
    '                        " WHERE PORevNo='" & Trim(pPORevNOApp) & "' AND PONo='" & Trim(pPONOAppR) & "' AND AffiliateID='" & Trim(pAffiliateIDRev) & "' AND SupplierID='" & Trim(pSupplierIDRev) & "' " & vbCrLf
    '        Dim SqlComm As New SqlCommand(ls_sql, sqlConn)
    '        SqlComm.ExecuteNonQuery()
    '        SqlComm.Dispose()
    '        sqlConn.Close()
    '    End Using
    'End Sub

    'Private Function EmailToEmailCCAppRev(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
    '    Dim ls_SQL As String = ""

    '    MdlConn.ReadConnection()
    '    Using sqlConn As New SqlConnection(MdlConn.uf_GetConString)

    '        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
    '                 " select 'AFF' flag,affiliatepocc, affiliatepoto,FromEmail = '' from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
    '                 " union all " & vbCrLf & _
    '                 " --PASI TO -CC " & vbCrLf & _
    '                 " select 'PASI' flag,affiliatepocc,affiliatepoto='',FromEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)
    '        If ds.Tables(0).Rows.Count > 0 Then
    '            Return ds
    '        End If
    '    End Using
    'End Function

    'Private Sub sendEmailAppRev()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Body As String = ""

    '    'uf_GetNotification("26")

    '    'ls_Body = pLine1 & vbCr & pLine2 & "PO Revision No:" & Trim(pPONO) & vbCr & vbCr & pLine3 & vbCr & pLine4 & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
    '    '"PORevFinalApproval.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>&t5=<%#GetRemarks(Container)%>&t6=<%#GetFinalApproval(Container)%>&t7=<%#GetDeliveryBy(Container)%>&t8=<%#GetPORevNo(Container)%>&Session=~/PurchaseOrderRevision/PORevFinalApprovalList.aspx"
    '    Dim ls_URl As String = "http://" & clsNotification.pub_ServerName & "/PurchaseOrderRevision/PORevFinalApproval.aspx?id2=" & clsNotification.EncryptURL(pPONOAppR.Trim) & "&t1=" & clsNotification.EncryptURL(pShipAppRev.Trim) & _
    '                                  "&t2=" & clsNotification.EncryptURL(pCommercialRev) & "&t3=" & clsNotification.EncryptURL(pPeriodRev) & "&t4=" & clsNotification.EncryptURL(pSupplierIDRev) & "&t5=" & clsNotification.EncryptURL(pRemarksRev) & "&t6=" & clsNotification.EncryptURL(pFinalAppRev) & "&t7=" & clsNotification.EncryptURL(pDelAppRev) & "&t8=" & clsNotification.EncryptURL(pPORevNOApp) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevFinalApprovalList.aspx")

    '    ls_Body = clsNotification.GetNotification("26", ls_URl, pPONOAppR.Trim, "", "", pPORevNOApp.Trim)

    '    Dim dsEmail As New DataSet
    '    dsEmail = EmailToEmailCCAppRev(Trim(pAffiliateIDRev), "PASI", "")
    '    '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '    For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '        If receiptCCEmail = "" Then
    '            receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '        Else
    '            receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '        End If
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '            fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
    '        End If
    '        If receiptEmail = "" Then
    '            receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '        Else
    '            receiptEmail = receiptEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '        End If
    '    Next
    '    receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '    receiptEmail = Replace(receiptEmail, ",", ";")

    '    receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '    receiptEmail = Replace(receiptEmail, " ", "")
    '    fromEmail = Replace(fromEmail, " ", "")

    '    If receiptEmail = "" Then
    '        MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '        Exit Sub
    '    End If

    '    If fromEmail = "" Then
    '        MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
    '        Exit Sub
    '    End If

    '    'Make a copy of the file/Open it/Mail it/Delete it
    '    'If you want to change the file name then change only TempFileName

    '    Dim mailMessage As New Mail.MailMessage()
    '    mailMessage.From = New MailAddress(fromEmail)
    '    mailMessage.Subject = "[TRIAL] Approval PO Revision No: " & Trim(pPORevNOApp)

    '    If receiptEmail <> "" Then
    '        For Each recipient In receiptEmail.Split(";"c)
    '            If recipient <> "" Then
    '                Dim mailAddress As New MailAddress(recipient)
    '                mailMessage.To.Add(mailAddress)
    '            End If
    '        Next
    '    End If
    '    If receiptCCEmail <> "" Then
    '        For Each recipientCC In receiptCCEmail.Split(";"c)
    '            If recipientCC <> "" Then
    '                Dim mailAddress As New MailAddress(recipientCC)
    '                mailMessage.CC.Add(mailAddress)
    '            End If
    '        Next
    '    End If
    '    GetSettingEmail()
    '    mailMessage.Body = ls_Body
    '    'Dim filename As String = TempFilePath & TempFileName
    '    'mailMessage.Attachments.Add(New Attachment(filename))
    '    mailMessage.IsBodyHtml = False
    '    Dim smtp As New SmtpClient
    '    'smtp.Host = "smtp.atisicloud.com"
    '    'smtp.Host = "mail.fast.net.id"
    '    'smtp.EnableSsl = False
    '    'smtp.UseDefaultCredentials = True
    '    'smtp.Port = 25
    '    'smtp.Send(mailMessage)

    '    smtp.Host = smtpClient
    '    If smtp.UseDefaultCredentials = True Then
    '        smtp.EnableSsl = True
    '    Else
    '        smtp.EnableSsl = False
    '        Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '        smtp.Credentials = myCredential
    '    End If

    '    smtp.Port = portClient
    '    smtp.Send(mailMessage)

    'End Sub

    'Private Function getPONoApproveRev() As DataSet
    '    Dim ls_SQL As String = ""
    '    MdlConn.ReadConnection()
    '    ls_SQL = " SELECT * FROM PORev_Master  PORM " & vbCrLf & _
    '              "  LEFT JOIN dbo.PORev_MasterUpload POMU ON PORM.PONo=POMU.PONo AND PORM.AffiliateID=POMU.AffiliateID AND PORM.SupplierID=POMU.SupplierID and PORM.PORevNo = POMU.PORevNo " & vbCrLf & _
    '              "  LEFT JOIN dbo.MS_Affiliate MA ON PORM.AffiliateID=MA.AffiliateID " & vbCrLf & _
    '              "  --LEFT JOIN PO_Master POM ON POM.PONo=PORM.PONo AND POM.AffiliateID=PORM.AffiliateID AND POM.SupplierID=PORM.SupplierID  " & vbCrLf & _
    '              " WHERE PASISendAffiliateDate is not null " & vbCrLf & _
    '              " and (SupplierApproveDate is null and SupplierApprovePendingDate is null and SupplierUnApproveDate is null)" & vbCrLf & _
    '              " and not exists ( " & vbCrLf & _
    '              "         select PORevNo from PORev_MasterUpload PODU where PORM.PONo = PODU.PONo and PORM.AffiliateID = PODU.AffiliateID and PORM.PORevNo = PODU.PORevNo and PORM.SupplierID = PODU.SupplierID" & vbCrLf & _
    '              ") order by PASISendAffiliateDate"

    '    Dim ds As New DataSet
    '    ds = uf_GetDataSet(ls_SQL)
    '    Return ds
    'End Function

    'Private Function getIntervalApprove() As DataSet
    '    Dim ls_SQL As String = ""
    '    MdlConn.ReadConnection()
    '    ls_SQL = " SELECT POApprovalDate,PORevisionApprovalDate FROM MS_EmailSetting  "
    '    Dim ds As New DataSet
    '    ds = uf_GetDataSet(ls_SQL)
    '    Return ds
    'End Function

    'Private Sub sendEmailAppRevtoAff()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Body As String = ""

    '    Try
    '        Dim ls_URl As String = "http://" & clsNotification.pub_ServerName & "/PurchaseOrderRevision/PORevFinalApproval.aspx?id2=" & clsNotification.EncryptURL(pPONOAppR.Trim) & _
    '        "&t1=" & clsNotification.EncryptURL(pShipAppRev.Trim) & "&t2=" & clsNotification.EncryptURL(pCommercialRev) & _
    '        "&t3=" & clsNotification.EncryptURL(pPeriodRev) & "&t4=" & clsNotification.EncryptURL(pSupplierIDRev) & _
    '        "&t5=" & clsNotification.EncryptURL(pRemarksRev) & "&t6=" & clsNotification.EncryptURL(pFinalAppRev) & _
    '        "&t7=" & clsNotification.EncryptURL(pDelAppRev) & "&t8=" & clsNotification.EncryptURL(pPORevNOApp) & _
    '        "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevFinalApprovalList.aspx")

    '        ls_Body = clsNotification.GetNotification("26", ls_URl, pPONOAppR.Trim, "", "", pPORevNOApp.Trim)

    '        Dim dsEmail As New DataSet
    '        dsEmail = EmailToEmailCCAppRev(Trim(pAffiliateIDRev), "PASI", "")
    '        '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '        For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '            If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '                fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
    '            End If
    '            If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
    '                If receiptEmail = "" Then
    '                    receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '                Else
    '                    receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '                End If
    '            End If
    '            If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
    '                If receiptCCEmail = "" Then
    '                    receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '                Else
    '                    receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '                End If
    '            End If
    '        Next
    '        receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '        receiptEmail = Replace(receiptEmail, " ", "")
    '        fromEmail = Replace(fromEmail, " ", "")

    '        receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '        receiptEmail = Replace(receiptEmail, ",", ";")

    '        If receiptEmail = "" Then
    '            MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '            Exit Sub
    '        End If

    '        If fromEmail = "" Then
    '            MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
    '            Exit Sub
    '        End If

    '        'Make a copy of the file/Open it/Mail it/Delete it
    '        'If you want to change the file name then change only TempFileName

    '        Dim mailMessage As New Mail.MailMessage()
    '        mailMessage.From = New MailAddress(fromEmail)
    '        mailMessage.Subject = "[TRIAL] Approval PO Revision No: " & Trim(pPORevNOApp)

    '        If receiptEmail <> "" Then
    '            For Each recipient In receiptEmail.Split(";"c)
    '                If recipient <> "" Then
    '                    Dim mailAddress As New MailAddress(recipient)
    '                    mailMessage.To.Add(mailAddress)
    '                End If
    '            Next
    '        End If
    '        If receiptCCEmail <> "" Then
    '            For Each recipientCC In receiptCCEmail.Split(";"c)
    '                If recipientCC <> "" Then
    '                    Dim mailAddress As New MailAddress(recipientCC)
    '                    mailMessage.CC.Add(mailAddress)
    '                End If
    '            Next
    '        End If
    '        GetSettingEmail()
    '        mailMessage.Body = ls_Body
    '        mailMessage.IsBodyHtml = False
    '        Dim smtp As New SmtpClient
    '        smtp.Host = smtpClient
    '        If smtp.UseDefaultCredentials = True Then
    '            smtp.EnableSsl = True
    '        Else
    '            smtp.EnableSsl = False
    '            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '            smtp.Credentials = myCredential
    '        End If

    '        smtp.Port = portClient
    '        smtp.Send(mailMessage)
    '    Catch ex As Exception
    '        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process AUTO Approval PO STOPPED, because " & ex.Message & " " & vbCrLf & _
    '                        rtbProcess.Text
    '    End Try
    'End Sub

    'Private Sub sendEmailAppRevtoccPASI()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Body As String = ""

    '    Try
    '        '*******File di Server
    '        'uf_GetNotification("16")

    '        'ls_Body = pLine1 & vbCr & pLine2 & "PO No:" & Trim(pPONO) & vbCr & vbCr & pLine3 & vbCr & pLine4 & pLine5 & vbCr & pLine6 & vbCr & pLine7 & vbCr & pLine8
    '        '"AffiliateOrderRevAppDetail.aspx?id=<%#GetRowValue(Container)%>
    '        '&t1=<%#GetPeriod(Container)%>&t2=<%#GetPORevNo(Container)%>
    '        '&t3=<%#GetPONo(Container)%>&t4=<%#GetCommercial(Container)%>
    '        '&t5=<%#GetAffiliateID(Container)%>&t6=<%#GetAffiliateName(Container)%>
    '        '&t7=<%#GetSupplierID(Container)%>&t8=<%#GetSupplierName(Container)%>
    '        '&t9=<%#GetKanban(Container)%>&t10=<%#GetRemarks(Container)%>&Session=~/AffiliateRevision/AffiliateOrderRevAppList.aspx"

    '        Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateRevision/AffiliateOrderRevAppDetail.aspx?id2=" & clsNotification.EncryptURL(pPONOApp.Trim) & _
    '                                       "&t1=" & clsNotification.EncryptURL(pPeriodApp.Trim) & "&t2=" & clsNotification.EncryptURL(pPORevNOApp.Trim) & _
    '                                       "&t3=" & clsNotification.EncryptURL(pPONOAppR) & "&t4=" & clsNotification.EncryptURL("") & _
    '                                       "&t5=" & clsNotification.EncryptURL(pAffiliateIDRev.Trim) & "&t6=" & clsNotification.EncryptURL("") & _
    '                                       "&t7=" & clsNotification.EncryptURL("") & "&t8=" & clsNotification.EncryptURL("") & _
    '                                       "&t9=" & clsNotification.EncryptURL("2") & "&t10=" & clsNotification.EncryptURL(pRemarksApp) & _
    '                                       "&Session=" & clsNotification.EncryptURL("~/AffiliateRevision/AffiliateOrderRevAppList.aspx")

    '        ls_Body = clsNotification.GetNotification("26", ls_URl, pPONOApp.Trim)

    '        Dim dsEmail As New DataSet
    '        dsEmail = EmailToEmailCCApp(Trim(pAffiliateIDApp), "PASI", "")
    '        '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '        For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '            If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '                fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
    '                If receiptCCEmail = "" Then
    '                    receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '                Else
    '                    receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '                End If
    '            End If
    '            receiptEmail = receiptCCEmail
    '        Next

    '        receiptCCEmail = Replace(receiptCCEmail, " ", "")
    '        receiptEmail = Replace(receiptEmail, " ", "")
    '        fromEmail = Replace(fromEmail, " ", "")

    '        receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '        receiptEmail = Replace(receiptEmail, ",", ";")

    '        If receiptEmail = "" Then
    '            MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '            Exit Sub
    '        End If

    '        If fromEmail = "" Then
    '            MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
    '            Exit Sub
    '        End If

    '        Dim mailMessage As New Mail.MailMessage()
    '        mailMessage.From = New MailAddress(fromEmail)
    '        mailMessage.Subject = "[TRIAL] Approval PONo: " & Trim(pPONOApp)

    '        If receiptEmail <> "" Then
    '            For Each recipient In receiptEmail.Split(";"c)
    '                If recipient <> "" Then
    '                    Dim mailAddress As New MailAddress(recipient)
    '                    mailMessage.To.Add(mailAddress)
    '                End If
    '            Next
    '        End If
    '        If receiptCCEmail <> "" Then
    '            For Each recipientCC In receiptCCEmail.Split(";"c)
    '                If recipientCC <> "" Then
    '                    Dim mailAddress As New MailAddress(recipientCC)
    '                    mailMessage.CC.Add(mailAddress)
    '                End If
    '            Next
    '        End If
    '        GetSettingEmail()
    '        mailMessage.Body = ls_Body
    '        mailMessage.IsBodyHtml = False
    '        Dim smtp As New SmtpClient

    '        smtp.Host = smtpClient
    '        If smtp.UseDefaultCredentials = True Then
    '            smtp.EnableSsl = True
    '        Else
    '            smtp.EnableSsl = False
    '            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '            smtp.Credentials = myCredential
    '        End If

    '        smtp.Port = portClient
    '        smtp.Send(mailMessage)
    '    Catch ex As Exception
    '        rtbProcess.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " Process AUTO Approval PO STOPPED, because " & ex.Message & " " & vbCrLf & _
    '                        rtbProcess.Text
    '    End Try
    'End Sub
End Class
