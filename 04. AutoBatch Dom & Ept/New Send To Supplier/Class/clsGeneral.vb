Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports System.Windows.Forms

Public Class clsGeneral
    Shared Sub up_displayLog(ByVal pMsgType As GlobalSetting.clsGlobal.MsgTypeEnum,
                              ByVal vMsg As String,
                              ByVal txtLog As RichTextBox)
        Dim ls_msgtype As String = ""
        Dim lmsg As String = ""
        Dim i As Integer = 0
        Dim ls_duration As String = ""

        Try
            If pMsgType = GlobalSetting.clsGlobal.MsgTypeEnum.ErrorMsg Then
                ls_msgtype = "Err"
            ElseIf pMsgType = GlobalSetting.clsGlobal.MsgTypeEnum.InformationMsg Then
                ls_msgtype = "Info"
            End If

            If Len(txtLog.Text) > 50000 Then txtLog.Text = ""

            lmsg = Format(Date.Now, "dd/MM/yy hh:mm:ss") & "  [ " & ls_msgtype & " ] : " & vMsg.ToString & vbCrLf
            txtLog.SelectionStart = 0
            txtLog.Text = lmsg & txtLog.Text
            txtLog.Refresh()
        Catch ex As Exception

        End Try
    End Sub

    Shared Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Shared Sub NAR(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    Shared Function Supplier(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT ISNULL(SupplierName,'') SupplierName, ISNULL(Address,'')Address FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function Affiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT ISNULL(AffiliateName,'')AffiliateName, ISNULL(Address,'')Address, ISNULL(ConsigneeName,'')ConsigneeName, ISNULL(ConsigneeAddress,'')ConsigneeAddress, ISNULL(BuyerName,'')BuyerName, ISNULL(BuyerAddress,'')BuyerAddress FROM dbo.MS_Affiliate WHERE AffiliateID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function PASI(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT ISNULL(AffiliateName,'')AffiliateName, ISNULL(Address,'')Address, ISNULL(ConsigneeName,'')ConsigneeName, ISNULL(ConsigneeAddress,'')ConsigneeAddress, ISNULL(BuyerName,'')BuyerName, ISNULL(BuyerAddress,'')BuyerAddress FROM dbo.MS_Affiliate WHERE AffiliateID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function DeliveryLocation(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT * FROM dbo.MS_DeliveryPlace WHERE DeliveryLocationCode='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function getEmailAddress(ByVal GB As GlobalSetting.clsGlobal, ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String, _
                                    ByVal pEmailCC As String, ByVal pEmailTo As String, ByVal pEmailFrom As String, ByRef ErrMsg As String) As DataSet
        Dim ls_SQL As String = ""

        Try
            ls_SQL = " SELECT 'AFF' FLAG, " & pEmailCC & " AS EmailCC, " & pEmailTo & " AS EmailTO, " & pEmailFrom & " AS EmailFrom FROM MS_EmailAffiliate WHERE AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                 " UNION ALL " & vbCrLf & _
                 " SELECT 'PASI' FLAG, " & pEmailCC & " AS EmailCC, " & pEmailTo & " AS EmailTO, " & pEmailFrom & " AS EmailFrom FROM MS_EmailPASI WHERE AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                 " UNION ALL " & vbCrLf & _
                 " SELECT 'SUPP' FLAG, " & pEmailCC & " AS EmailCC, " & pEmailTo & " AS EmailTO, " & pEmailFrom & " as EmailFrom FROM MS_EmailSupplier WHERE SupplierID='" & Trim(pSupplierID) & "'"
            Dim ds As New DataSet
            ds = GB.uf_GetDataSet(ls_SQL)
            Return ds
        Catch ex As Exception
            Return Nothing
            ErrMsg = "Get Email Address Failed, " & ex.Message
        End Try
    End Function

    Shared Function GetSettingEmail(ByVal GB As GlobalSetting.clsGlobal, ByVal errMsg As String) As List(Of clsSendToSupplier)
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT * FROM dbo.Ms_EmailSetting"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then

            Dim SettingEmail As New List(Of clsSendToSupplier)

            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim SetEmail As New clsSendToSupplier
                SetEmail.smtpClient = Trim(ds.Tables(0).Rows(i)("SMTP"))
                SetEmail.portClient = Trim(ds.Tables(0).Rows(i)("PORTSMTP"))
                SetEmail.usernameSMTP = If(IsDBNull(ds.Tables(0).Rows(i)("usernameSMTP")), "", ds.Tables(0).Rows(i)("usernameSMTP"))
                SetEmail.PasswordSMTP = If(IsDBNull(ds.Tables(0).Rows(i)("passwordSMTP")), "", ds.Tables(0).Rows(i)("passwordSMTP"))
                SetEmail.DefaultCredentials = IIf(ds.Tables(0).Rows(i)("DefaultCredentials") = "1", True, False)
                SetEmail.SSL = IIf(ds.Tables(0).Rows(i)("SSL") = "1", True, False)

                'SetEmail.smtpClient = "smtp.atisicloud.com"
                'SetEmail.portClient = "25"
                'SetEmail.usernameSMTP = "admin-dms@pemi.co.id"
                'SetEmail.PasswordSMTP = "P4ssword!"
                SettingEmail.Add(SetEmail)
            Next

            Return SettingEmail
            Exit Function
        Else
            errMsg = "Process Send to Supplier STOPPED, because Email Setting Empty "
        End If
        Return Nothing
    End Function

    Shared Function getIntervalApprove(ByVal GB As GlobalSetting.clsGlobal) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT POApprovalDate, PORevisionApprovalDate, KanbanApprovalHour FROM MS_EmailSetting  "
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function getIntervalApproveExport(ByVal GB As GlobalSetting.clsGlobal) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " SELECT POApprovalDate,PORevisionApprovalDate FROM MS_EmailSetting_Export  "
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        Return ds
    End Function

    Shared Function sendEmail(ByVal GB As GlobalSetting.clsGlobal, ByVal pFromEmail As String, ByVal pToEmail As String, ByVal pCCEmail As String, _
                              ByVal pBCCEmail As String, ByVal pSubject As String, ByVal pBody As String, ByRef errMsg As String, _
                              Optional ByVal pFile1 As String = "", Optional ByVal pFile2 As String = "", Optional ByVal pFile3 As String = "", _
                              Optional ByVal pFile4 As String = "", Optional ByVal pFile5 As String = "", Optional ByVal pFile6 As String = "") As Boolean
        Try
            sendEmail = True

            'pFromEmail = "admin-dms@pemi.co.id"
            'pToEmail = "edi@tos.co.id"
            'pCCEmail = "edi@tos.co.id"

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(pFromEmail)

            If pToEmail <> "" Then
                For Each recipient In pToEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If

            If pCCEmail <> "" Then
                For Each recipientCC In pCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            'If pBCCEmail.Trim.ToUpper = "KMK" Then
            '    pBCCEmail = "Ariefrosyadi91@gmail.com;Ristian_a@yahoo.com;pasi.purchase@gmail.com"
            'Else
            '    pBCCEmail = "pasi.purchase@gmail.com"
            'End If

            'If pBCCEmail <> "" Then
            '    For Each recipientBCC In pBCCEmail.Split(";"c)
            '        If recipientBCC <> "" Then
            '            Dim mailAddress As New MailAddress(recipientBCC)
            '            mailMessage.Bcc.Add(mailAddress)
            '        End If
            '    Next
            'End If

            Dim setEmail As List(Of clsSendToSupplier)
            setEmail = GetSettingEmail(GB, errMsg)

            mailMessage.Subject = frmSendMail.SubjectEmail & pSubject
            mailMessage.Body = pBody

            If pFile1 <> "" Then
                Dim fi1 As New FileInfo(pFile1)
                If fi1.Exists Then
                    mailMessage.Attachments.Add(New Attachment(pFile1))
                End If
            End If

            If pFile2 <> "" Then
                Dim fi1 As New FileInfo(pFile2)
                If fi1.Exists Then
                    mailMessage.Attachments.Add(New Attachment(pFile2))
                End If
            End If

            If pFile3 <> "" Then
                Dim fi1 As New FileInfo(pFile3)
                If fi1.Exists Then
                    mailMessage.Attachments.Add(New Attachment(pFile3))
                End If
            End If

            If pFile4 <> "" Then
                Dim fi1 As New FileInfo(pFile4)
                If fi1.Exists Then
                    mailMessage.Attachments.Add(New Attachment(pFile4))
                End If
            End If

            If pFile5 <> "" Then
                Dim fi1 As New FileInfo(pFile5)
                If fi1.Exists Then
                    mailMessage.Attachments.Add(New Attachment(pFile5))
                End If
            End If

            If pFile6 <> "" Then
                Dim fi1 As New FileInfo(pFile6)
                If fi1.Exists Then
                    mailMessage.Attachments.Add(New Attachment(pFile6))
                End If
            End If

            mailMessage.IsBodyHtml = False

            Dim smtp As New SmtpClient
            smtp.Host = setEmail(0).smtpClient
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            smtp.UseDefaultCredentials = setEmail(0).DefaultCredentials
            smtp.EnableSsl = setEmail(0).SSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(setEmail(0).usernameSMTP), Trim(setEmail(0).PasswordSMTP))
            smtp.Credentials = myCredential

            'If smtp.UseDefaultCredentials = True Then
            '    smtp.EnableSsl = True
            'Else
            '    smtp.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(setEmail(0).usernameSMTP), Trim(setEmail(0).PasswordSMTP))
            '    smtp.Credentials = myCredential
            'End If

            smtp.Port = setEmail(0).portClient
            Threading.Thread.Sleep(5000)
            smtp.Send(mailMessage)
        Catch ex As Exception
            sendEmail = False
            errMsg = "Failed Send Email, " & ex.Message

            If InStr(ex.Message, "timed out") > 0 Then
                sendEmail = True
                errMsg = "Failed Send Email, " & ex.Message & " By Pass"
            End If
        End Try
    End Function

    Public Shared Function BatchProcessStatus(ByVal GB As GlobalSetting.clsGlobal) As String
        Dim sql As String = ""
        Dim ds As DataSet

        sql = "SELECT * FROM dbo.BatchProcessStatus "
        ds = GB.uf_GetDataSet(sql)

        If ds.Tables(0).Rows.Count > 0 Then
            BatchProcessStatus = ds.Tables(0).Rows(0)("BatchProcessStatus")
        Else
            BatchProcessStatus = ""
        End If
    End Function

    Public Shared Sub BatchProcessStatusUpdate(ByVal GB As GlobalSetting.clsGlobal)
        Dim sql As String = ""

        sql = "UPDATE dbo.BatchProcessStatus SET BatchProcessStatus = '3'"
        GB.uf_ExecuteSql(sql)
    End Sub
End Class
