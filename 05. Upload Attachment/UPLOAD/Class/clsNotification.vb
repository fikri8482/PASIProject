Imports System.Transactions
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text

Public Class clsNotification
    'Public Const pub_ServerNamePasi As String = "localhost:5832"
    'Public Const pub_ServerNameAffiliate As String = "localhost:48622"
    Public Const pub_ServerNamePasi As String = "10.62.212.17/PASISystem"
    Public Const pub_ServerNameAffiliate As String = "10.62.212.17/AffiliateSystem"

    'Public Const pub_ServerName As String = "VIOS"
    'Public Const pub_ServerName As String = "192.168.0.5"

    Public Shared Function GetNotification(ByVal pNotificationCode As String, Optional ByVal pUrl As String = "", _
                                     Optional ByVal pPONo As String = "", Optional ByVal pKanban As String = "", _
                                     Optional ByVal pSuratJalanNo As String = "", Optional ByVal pPORevision As String = "") As String

        Dim ls_SQL As String = ""
        Dim ls_BodyMessage As String = ""
        'Dim clsGlobal As New clsGlobal

        Dim ls_Line(8) As String
        Dim ls_Cls(8) As String

        Dim xyz As Integer

        MdlConn.ReadConnection()
        Using sqlConn As New SqlConnection(MdlConn.uf_GetConString)
            sqlConn.Open()

            ls_SQL = "select [Line1],[Line1Cls],[Line2],[Line2Cls],[Line3],[Line3Cls], " & vbCrLf & _
                     "[Line4],[Line4Cls],[Line5],[Line5Cls], " & vbCrLf & _
                     "[Line6],[Line6Cls],[Line7],[Line7Cls], " & vbCrLf & _
                     "[Line8],[Line8Cls] from ms_notification where notificationcode = '" & pNotificationCode & "'" & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If IsDBNull(ds.Tables(0).Rows(0)("Line1")) Then
                    ls_Line(0) = ""
                    ls_Cls(0) = ""
                Else
                    ls_Line(0) = ds.Tables(0).Rows(0)("Line1")
                    ls_Cls(0) = ds.Tables(0).Rows(0)("Line1Cls")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("Line2")) Then
                    ls_Line(1) = ""
                    ls_Cls(1) = ""
                Else
                    ls_Line(1) = ds.Tables(0).Rows(0)("Line2")
                    ls_Cls(1) = ds.Tables(0).Rows(0)("Line2Cls")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("Line3")) Then
                    ls_Line(2) = ""
                    ls_Cls(2) = ""
                Else
                    ls_Line(2) = ds.Tables(0).Rows(0)("Line3") & vbCrLf
                    ls_Cls(2) = ds.Tables(0).Rows(0)("Line3Cls")
                End If

                If IsDBNull(ds.Tables(0).Rows(0)("Line4")) Then
                    ls_Line(3) = ""
                    ls_Cls(3) = ""
                Else
                    ls_Line(3) = ds.Tables(0).Rows(0)("Line4") & vbCrLf
                    ls_Cls(3) = ds.Tables(0).Rows(0)("Line4Cls")
                End If

                If IsDBNull(ds.Tables(0).Rows(0)("Line5")) Then
                    ls_Line(4) = ""
                    ls_Cls(4) = ""
                Else
                    ls_Line(4) = ds.Tables(0).Rows(0)("Line5") & vbCrLf
                    ls_Cls(4) = ds.Tables(0).Rows(0)("Line5Cls")
                End If

                If IsDBNull(ds.Tables(0).Rows(0)("Line6")) Then
                    ls_Line(5) = ""
                    ls_Cls(5) = ""
                Else
                    ls_Line(5) = ds.Tables(0).Rows(0)("Line6") & vbCrLf
                    ls_Cls(5) = ds.Tables(0).Rows(0)("Line6Cls")
                End If

                If IsDBNull(ds.Tables(0).Rows(0)("Line7")) Then
                    ls_Line(6) = ""
                    ls_Cls(6) = ""
                Else
                    ls_Line(6) = ds.Tables(0).Rows(0)("Line7") & vbCrLf
                    ls_Cls(6) = ds.Tables(0).Rows(0)("Line7Cls")
                End If

                If IsDBNull(ds.Tables(0).Rows(0)("Line8")) Then
                    ls_Line(7) = ""
                    ls_Cls(7) = ""
                Else
                    ls_Line(7) = ds.Tables(0).Rows(0)("Line8") & vbCrLf
                    ls_Cls(7) = ds.Tables(0).Rows(0)("Line8Cls")
                End If
            End If

            For xyz = 0 To 7
                If ls_Cls(xyz) = "05" Then
                    ls_BodyMessage = ls_BodyMessage + ls_Line(xyz)
                End If
                If ls_Cls(xyz) <> "05" Then
                    If ls_Cls(xyz) = "01" Then 'Affiliate PO NO
                        ls_BodyMessage = ls_BodyMessage & vbCrLf
                        ls_BodyMessage = ls_BodyMessage & "PO No: " & pPONo
                    End If
                    If ls_Cls(xyz) = "02" Then 'Affiliate PO NO Revision
                        ls_BodyMessage = ls_BodyMessage & vbCrLf
                        ls_BodyMessage = ls_BodyMessage & "PO Rev No: " & pPORevision
                    End If
                    If ls_Cls(xyz) = "03" Then 'Affiliate Kanban No
                        ls_BodyMessage = ls_BodyMessage & vbCrLf
                        ls_BodyMessage = ls_BodyMessage & "Kanban No: " & pKanban
                    End If
                    If ls_Cls(xyz) = "04" Then 'Surat jalan
                        ls_BodyMessage = ls_BodyMessage & vbCrLf
                        ls_BodyMessage = ls_BodyMessage & "Surat Jalan No: " & pSuratJalanNo
                    End If
                    If ls_Cls(xyz) = "06" Then 'Affiliate PO Link
                        'String.Format("~/VB2.aspx?name={0}&technology={1}", name, technology)
                        ls_BodyMessage = ls_BodyMessage & vbCrLf
                        ls_BodyMessage = ls_BodyMessage & pUrl
                    End If
                End If
            Next

            Return ls_BodyMessage

        End Using
    End Function

    Public Shared Function DecryptURL(cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        cipherText = cipherText.Replace(" ", "+")
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function

    Public Shared Function EncryptURL(clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
End Class
