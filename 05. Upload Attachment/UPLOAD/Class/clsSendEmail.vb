Imports System.Net.Mail

Public Class clsSendEmail
    Shared Function SendEmail(ByVal Recipients As List(Of String), _
                        ByVal CC As List(Of String), _
                        ByVal FromAddress As String, _
                        ByVal Subject As String, _
                        ByVal Body As String, _
                        ByVal UserName As String, _
                        ByVal Password As String, _
                        ByVal EnableSSL As Boolean, _
                        ByVal DefaultCredentials As Boolean, _
                        ByVal Server As String, _
                        ByVal Port As Integer, ByVal pFile As String) As String

        Dim Email As New MailMessage()
        Dim retMsg As String
        'Dim pBCC As String = "adhi@tos.co.id"

        Try

            Email.From = New MailAddress(FromAddress) '("hadi@tos.co.id") '
            For Each Recipient As String In Recipients
                Email.To.Add(Recipient)
            Next
            For Each CC1 As String In CC
                If CC1 <> "" Then
                    Email.CC.Add(CC1)
                    'Email.To.Add(CC1)
                End If
            Next
            'Email.Bcc.Add(pBCC)

            Email.Subject = Subject
            Email.Body = Body

            Email.IsBodyHtml = False

            Dim SMTPServer As New SmtpClient


            Dim filename As String = pFile
            If pFile <> "" Then Email.Attachments.Add(New Attachment(filename))

            'SMTPServer.Host = Server '"tos-is.com" '
            SMTPServer.Host = Server
            '' 20220221 : Setting SSL and DefaultCredentials from Database Now
            SMTPServer.UseDefaultCredentials = DefaultCredentials
            SMTPServer.EnableSsl = EnableSSL

            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(UserName), Trim(Password))
            SMTPServer.Credentials = myCredential
            '' ''SMTPServer.Credentials = New System.Net.NetworkCredential(UserName, Password)
            '' ''SMTPServer.EnableSsl = EnableSSL
            'If SMTPServer.UseDefaultCredentials = True Then
            '    SMTPServer.EnableSsl = True
            'Else
            '    SMTPServer.EnableSsl = False
            '    Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(UserName), Trim(Password))
            '    SMTPServer.Credentials = myCredential
            'End If

            SMTPServer.Port = Port
            SMTPServer.Send(Email)
            Email.Dispose()
            Return ""

        Catch ex As SmtpException
            Email.Dispose()
            retMsg = "Sending Email Failed. Smtp Error."
        Catch ex As ArgumentOutOfRangeException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number."
        Catch Ex As InvalidOperationException
            Email.Dispose()
            retMsg = "Sending Email Failed. Check Port Number."
        End Try
        Return retMsg

    End Function
End Class
