Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Public Class clsTes
    Shared Sub tessend()
        Dim out As Outlook.MailItem
        Dim appOutlook As New Outlook.Application
        Try
            out = appOutlook.CreateItem(Outlook.OlItemType.olMailItem)

            MessageBox.Show(out.Session.Accounts.Item(1).DisplayName.ToString())

            out.SendUsingAccount = out.Session.Accounts.Item(1)

            Dim Recipents As Outlook.Recipients = out.Recipients

            Dim Recipent As Outlook.Recipient = Nothing
            Recipent = Recipents.Add("fikriismail82@gmail.com")
            Recipent.Type = Outlook.OlMailRecipientType.olTo


            Dim RecipentCC As Outlook.Recipient = Nothing
            RecipentCC = Recipents.Add("Ayulstari16@gmail.com")
            RecipentCC.Type = Outlook.OlMailRecipientType.olCC

            Dim RecipentBCC As Outlook.Recipient = Nothing
            RecipentBCC = Recipents.Add("lisdawati7562@gmail.com")
            RecipentBCC.Type = Outlook.OlMailRecipientType.olBCC


            out.Subject = "Testing e-mail"
            out.Body = "Hi Fikri," & vbCrLf & "This is a testing E-mail."
            out.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            out.Attachments.Add("D:\PASI EBWEB\01. TEMPLATE EXCEL\Result\Barcode-220221PM2 PEMI-PIOLAX.pdf")
            out.Attachments.Add("D:\PASI EBWEB\01. TEMPLATE EXCEL\Result\Delivery PEMI-PIOLAX-220221PM2.xlsm")
            out.Send()
            MessageBox.Show("Done!")
        Catch ex1 As System.Exception
            MessageBox.Show("failed!" & " _ " & ex1.Message & " _ " & ex1.ToString())
        Finally
            out = Nothing
            appOutlook = Nothing
        End Try
    End Sub
End Class
