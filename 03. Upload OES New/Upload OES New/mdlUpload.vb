Imports System.Data
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail

Module mdlUpload
    Public Declare Function uf_GetPrivateProfileString Lib _
        "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Integer, ByVal lpFileName As String) As _
        Integer

    Public Declare Function uf_WritePrivateProfileString Lib _
        "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpString As String, ByVal lpFileName As String) As _
        Long

    Public gs_ConString As String
    Public gs_UploadPath As String

    Dim strSMTPHost As String = ""
    Dim strSMTPPort As String = ""
    Dim strSMTPEmail As String = ""
    Dim strSMTPPassword As String = ""

    Dim strFromEmail As String = ""
    Dim strToEmail As String = ""
    Dim strCCEmail As String = ""
    Dim strBCCEmail As String = ""

    Public Sub up_GetSetting()
        Dim strConfigPath As String
        Dim lngResult As Long
        Dim strResult As String

        Dim strServer As String = ""
        Dim strDatabase As String = ""
        Dim strUser As String = ""
        Dim strPassword As String = ""

        Try
            'get setting
            strConfigPath = My.Application.Info.DirectoryPath & "\UPLOAD.ini"

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Setting", "Server", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strServer = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Setting", "Database", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strDatabase = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Setting", "User", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strUser = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Setting", "Password", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strPassword = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Setting", "Path", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then gs_UploadPath = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            gs_ConString = "Data Source=" & strServer & ";Initial Catalog=" & strDatabase & ";User ID=" & strUser & ";pwd=" & strPassword & ""

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Email", "SMTPHost", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strSMTPHost = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Email", "SMTPPort", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strSMTPPort = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Email", "SMTPEmail", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strSMTPEmail = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Email", "SMTPPassword", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strSMTPPassword = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Email", "EmailTo", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strToEmail = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            lngResult = 0
            strResult = Space(1500)
            lngResult = uf_GetPrivateProfileString("Email", "EmailCC", "", strResult, 1500, strConfigPath)
            If lngResult <> 0 Then strCCEmail = Microsoft.VisualBasic.Left(strResult, CInt(lngResult))

            strFromEmail = strSMTPEmail

            'check setting
            Using sqlCon As New SqlConnection(gs_ConString)
                sqlCon.Open()
            End Using

            If Dir(gs_UploadPath, FileAttribute.Directory) = "" Then
                MkDir(gs_UploadPath)
            End If

            If Dir(gs_UploadPath & "\Backup", FileAttribute.Directory) = "" Then
                MkDir(gs_UploadPath & "\Backup")
            End If

            If Dir(gs_UploadPath & "\Error", FileAttribute.Directory) = "" Then
                MkDir(gs_UploadPath & "\Error")
            End If

        Catch exSQL As SqlException
            MsgBox(exSQL.Message, MsgBoxStyle.Critical, "Error")
            End
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End
        End Try
    End Sub

    Public Function uf_GetDataSet(ByVal ls_query As String, Optional ByVal DbLock As SqlClient.SqlConnection = Nothing) As DataSet
        Dim lcon As New SqlConnection, lds As New DataSet

        Try
            If DbLock Is Nothing Then
                lcon.ConnectionString = gs_ConString

                lcon.Open()
                Dim lda As New SqlDataAdapter(ls_query, lcon)
                lda.Fill(lds)
                lcon.Close()
            Else
                Dim lda As New SqlDataAdapter(ls_query, DbLock)
                lda.Fill(lds)
                lcon.Close()
            End If

        Catch ex As SqlException
            ex.Message.ToString()
        End Try

        Return lds
    End Function

    Public Function uf_SendEmail(pMessage As String, strAffiliateCode As String) As String
        Try
            uf_SendEmail = ""

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(strFromEmail)

            If strToEmail <> "" Then
                For Each recipient In strToEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If

            If strCCEmail <> "" Then
                For Each recipientCC In strCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            If strBCCEmail <> "" Then
                For Each recipientBCC In strBCCEmail.Split(";"c)
                    If recipientBCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientBCC)
                        mailMessage.Bcc.Add(mailAddress)
                    End If
                Next
            End If

            mailMessage.Subject = "[Upload PO EDI] Error Notification - " & strAffiliateCode
            mailMessage.Body = pMessage

            mailMessage.IsBodyHtml = False

            Dim smtp As New SmtpClient
            smtp.Host = strSMTPHost
            If smtp.UseDefaultCredentials = True Then
                smtp.EnableSsl = True
            Else
                smtp.EnableSsl = False
                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(strSMTPEmail, strSMTPPassword)
                smtp.Credentials = myCredential
            End If

            smtp.Port = strSMTPPort
            smtp.Timeout = 600
            smtp.Send(mailMessage)

        Catch ex As Exception
            uf_SendEmail = ex.Message
        End Try
    End Function
End Module
