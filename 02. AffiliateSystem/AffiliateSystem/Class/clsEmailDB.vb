Imports System.Data.SqlClient

Public Class clsEmailDB
    Public Shared Function GetSettingEmail(ByVal pConStr As String) As List(Of clsEmail)
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM dbo.Ms_EmailSetting "

            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            Dim EmailList As New List(Of clsEmail)

            For i = 0 To dt.Rows.Count - 1
                With dt.Rows(i)
                    Dim Email As New clsEmail
                    Email.smtpClient = .Item("SMTP") & ""
                    Email.portClient = .Item("PORTSMTP") & ""
                    Email.usernameSMTP = .Item("usernameSMTP") & ""
                    Email.PasswordSMTP = .Item("passwordSMTP") & ""
                    EmailList.Add(Email)
                End With
            Next
            Return EmailList
        End Using
    End Function

    Public Shared Function GetEmailList(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String, ByVal pFrom As String, ByVal pKolom As String, ByVal pConstr As String) As List(Of clsEmail)
        Dim que As String

        Using cn As New SqlConnection(pConstr)
            cn.Open()

            que = "--Affiliate TO-CC: " & vbCrLf & _
                    " select 'AFF' flag, affiliatepocc, affiliatepoto, fromEmail = " & IIf(pFrom = "0", pKolom, "'" & "'") & " from ms_emailAffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                    " union all " & vbCrLf & _
                    " --PASI TO -CC " & vbCrLf & _
                    " select 'PASI' flag, affiliatepocc, affiliatepoto, fromEmail = " & IIf(pFrom = "1", pKolom, "'" & "'") & " from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                    " union all " & vbCrLf & _
                    " --Supplier TO- CC " & vbCrLf & _
                    " select 'SUPP' flag, affiliatepocc, affiliatepoto, fromEmail='' from ms_emailSupplier where SupplierID='" & Trim(pSupplierID) & "'"

            Dim cmd As New SqlCommand(que, cn)
            Dim rd As SqlDataReader = cmd.ExecuteReader
            Dim EmailList As New List(Of clsEmail)
            Do While rd.Read
                Dim Email As New clsEmail
                Email.Flag = rd("flag") & ""
                Email.EmailFrom = rd("fromEmail") & ""
                Email.EmailTo = rd("affiliatepoto") & ""
                Email.EmailCC = rd("affiliatepocc") & ""
                EmailList.Add(Email)
            Loop
            rd.Close()
            Return EmailList
        End Using
    End Function
End Class
