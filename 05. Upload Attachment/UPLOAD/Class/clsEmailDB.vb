Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration

Imports System.Net
Imports System.IO

Imports System.Windows.Forms
Imports System.Reflection


Public Class clsEmailDB
    Public Shared Function GetEmailSettingEx(ByVal pConStr As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailSetting_Export WITH(NOLOCK)"
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    Ems.SMTPServer = .Item("SMTP") & "" '"mail.fast.net.id" '
                    Ems.SenderName = .Item("UsernameSMTP") & "" '"hadi@tos.co.id" '
                    Ems.Port = .Item("PORTSMTP") & ""
                    Ems.Password = .Item("passwordSMTP") & "" '"0k3h0k3h" '
                    Ems.EnableSSL = .Item("SSL") & "" 'Val(.Item("EnableSSL") & "")
                    Ems.DefaultCredentials = IIf(.Item("DefaultCredentials") = "1", True, False)
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailSetting(ByVal pConStr As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailSetting WITH(NOLOCK)"
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    Ems.SMTPServer = .Item("SMTP") & "" '"mail.fast.net.id" '
                    Ems.SenderName = .Item("usernameSMTP") & "" '"hadi@tos.co.id" '
                    Ems.Port = .Item("PORTSMTP") & ""
                    Ems.Password = .Item("passwordSMTP") & "" '"0k3h0k3h" '
                    Ems.EnableSSL = .Item("SSL") & "" 'Val(.Item("EnableSSL") & "")
                    Ems.DefaultCredentials = IIf(.Item("DefaultCredentials") = "1", True, False)
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailPASI(ByVal pConStr As String, ByVal templateCode As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailPASI WITH(NOLOCK) WHERE AffiliateID = 'PASI' " '
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    If templateCode = "PO" Then
                        Ems.EmailPASITo = .Item("AffiliatePOTo") & ""
                        Ems.EmailPASICC = .Item("AffiliatePOCC") & ""
                    ElseIf templateCode = "POR" Then
                        Ems.EmailPASITo = .Item("AffiliatePORevisionTo") & ""
                        Ems.EmailPASICC = .Item("AffiliatePORevisionCC") & ""
                    ElseIf templateCode = "KB" Then
                        Ems.EmailPASITo = .Item("KanbanTo") & ""
                        Ems.EmailPASICC = .Item("KanbanCC") & ""
                    ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                        Ems.EmailPASITo = .Item("SupplierDeliveryTo") & ""
                        Ems.EmailPASICC = .Item("SupplierDeliveryCC") & ""
                    ElseIf templateCode = "INV" Then
                        Ems.EmailPASITo = .Item("InvoiceTo") & ""
                        Ems.EmailPASICC = .Item("InvoiceCC") & ""
                    End If
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailAffiliate(ByVal pConStr As String, ByVal pAffiliateID As String, ByVal templateCode As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailAffiliate WITH(NOLOCK) WHERE AffiliateID = '" & Trim(pAffiliateID) & "'"
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    If templateCode = "PO" Then
                        Ems.EmailAffiliateTo = .Item("AffiliatePOTo") & ""
                        Ems.EmailAffiliateCC = .Item("AffiliatePOCC") & ""
                    ElseIf templateCode = "POR" Then
                        Ems.EmailAffiliateTo = .Item("AffiliatePORevisionTo") & ""
                        Ems.EmailAffiliateCC = .Item("AffiliatePORevisionCC") & ""
                    ElseIf templateCode = "KB" Then
                        Ems.EmailAffiliateTo = .Item("KanbanTo") & ""
                        Ems.EmailAffiliateCC = .Item("KanbanCC") & ""
                    ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                        Ems.EmailAffiliateTo = .Item("SupplierDeliveryTo") & ""
                        Ems.EmailAffiliateCC = .Item("SupplierDeliveryCC") & ""
                    ElseIf templateCode = "INV" Then
                        Ems.EmailAffiliateTo = .Item("InvoiceTo") & ""
                        Ems.EmailAffiliateCC = .Item("InvoiceCC") & ""
                    End If
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailSupplier(ByVal pConStr As String, ByVal pSupplierID As String, ByVal templateCode As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailSupplier WITH(NOLOCK) WHERE SupplierID = '" & Trim(pSupplierID) & "'"
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    If templateCode = "PO" Then
                        Ems.EmailSupplierTo = .Item("AffiliatePOTo") & ""
                        Ems.EmailSupplierCC = .Item("AffiliatePOCC") & ""
                    ElseIf templateCode = "POR" Then
                        Ems.EmailSupplierTo = .Item("AffiliatePORevisionTo") & ""
                        Ems.EmailSupplierCC = .Item("AffiliatePORevisionCC") & ""
                    ElseIf templateCode = "KB" Then
                        Ems.EmailSupplierTo = .Item("KanbanTo") & ""
                        Ems.EmailSupplierCC = .Item("KanbanCC") & ""
                    ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                        Ems.EmailSupplierTo = .Item("SupplierDeliveryTo") & ""
                        Ems.EmailSupplierCC = .Item("SupplierDeliveryCC") & ""
                    ElseIf templateCode = "INV" Then
                        Ems.EmailSupplierTo = .Item("InvoiceTo") & ""
                        Ems.EmailSupplierCC = .Item("InvoiceCC") & ""
                    End If
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailFWD(ByVal pConStr As String, ByVal pFwd As String, ByVal templateCode As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailForwarder WITH(NOLOCK) WHERE ForwarderID = '" & Trim(pFwd) & "'"
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    Ems.EmailFWDTo = .Item("POExportTo") & ""
                    Ems.EmailFWDCC = .Item("POExportCC") & ""
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailSupplierEx(ByVal pConStr As String, ByVal pSupplierID As String, ByVal templateCode As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailSupplier_Export WITH(NOLOCK) WHERE SupplierID = '" & Trim(pSupplierID) & "'"
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    If templateCode = "POEM" Or templateCode = "POEE" Then
                        Ems.EmailSupplierTo = .Item("AffiliatePOTo") & ""
                        Ems.EmailSupplierCC = .Item("AffiliatePOCC") & ""
                    End If
                    If templateCode = "DO-EX" Then
                        Ems.EmailSupplierTo = .Item("SupplierDeliveryTo") & ""
                        Ems.EmailSupplierCC = .Item("SupplierDeliveryCC") & ""
                    End If
                    If templateCode = "REC-EX" Then
                        Ems.EmailSupplierTo = .Item("GoodReceiveTO") & ""
                        Ems.EmailSupplierCC = .Item("GoodReceiveCC") & ""
                    End If
                    If templateCode = "INV-EX" Or templateCode = "TALLY" Then
                        Ems.EmailSupplierTo = .Item("InvoiceTO") & ""
                        Ems.EmailSupplierCC = .Item("InvoiceCC") & ""
                    End If
                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailPASIEx(ByVal pConStr As String, ByVal templateCode As String) As clsEmail
        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = "SELECT * FROM MS_EmailPASI_Export WITH(NOLOCK) WHERE AffiliateID = 'PASI' " '
            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    If templateCode = "POEM" Or templateCode = "POEE" Then
                        Ems.EmailPASITo = .Item("AffiliatePOTo") & ""
                        Ems.EmailPASICC = .Item("AffiliatePOCC") & ""
                    ElseIf templateCode = "DO-EX" Then
                        Ems.EmailPASITo = .Item("SupplierDeliveryTO") & ""
                        Ems.EmailPASICC = .Item("SupplierDeliveryCC") & ""
                    ElseIf templateCode = "REC-EX" Then
                        Ems.EmailPASITo = .Item("GoodReceiveTO") & ""
                        Ems.EmailPASICC = .Item("GoodReceiveCC") & ""
                    ElseIf templateCode = "INV-EX" Then
                        Ems.EmailPASITo = .Item("InvoiceTO") & ""
                        Ems.EmailPASICC = .Item("InvoiceCC") & ""
                    ElseIf templateCode = "TALLY" Then
                        Ems.EmailPASITo = .Item("InvoiceTO") & ""
                        Ems.EmailPASICC = .Item("InvoiceCC") & ""
                    Else
                        Ems.EmailPASITo = .Item("KanbanTo") & ""
                        Ems.EmailPASICC = .Item("KanbanCC") & ""
                    End If

                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailSubjectBody(ByVal pConStr As String, ByVal templateCode As String, ByVal pURL As String, _
                                               ByVal pAllNo As String, Optional psts As Boolean = True) As clsEmail

        Dim ketLine1 As String = "", ketLine2 As String = "", ketLine3 As String = ""
        Dim ketLine4 As String = "", ketLine5 As String = "", ketLine6 As String = ""
        Dim ketLine7 As String = "", ketLine8 As String = ""

        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = ""
            Dim NotifCode As String = ""

            If templateCode = "PO" Then
                NotifCode = "12"
                pAllNo = "(PO NO: " & Trim(pAllNo) & ")"
            ElseIf templateCode = "POR" Then
                NotifCode = "22"
                pAllNo = "(PORev NO: " & Trim(pAllNo) & ")"
            ElseIf templateCode = "KB" Then
                NotifCode = "31"
                pAllNo = "(KANBAN NO: " & Trim(pAllNo) & ")"
            ElseIf templateCode = "DO" Or templateCode = "DO (PO non Kanban)" Then
                NotifCode = "40"
                If MsgstatusSave = "OVER" Then NotifCode = "41"
                pAllNo = "(SJ NO: " & Trim(pAllNo) & ")"
            ElseIf templateCode = "INV" Then
                NotifCode = "80"
                If psts = False Then NotifCode = "81"
                pAllNo = "(INVOICE NO: " & Trim(pAllNo) & ")"
            ElseIf templateCode = "REC-EX" Then
                NotifCode = "91"
                pAllNo = "(SURAT JALAN NO: " & Trim(pAllNo) & ")"
            End If

            q = " SELECT * FROM MS_NotificationCls A WITH(NOLOCK) " & vbCrLf & _
                " LEFT JOIN MS_Notification B WITH(NOLOCK) ON A.NotificationCode = B.NotificationCode " & vbCrLf & _
                " WHERE A.NotificationCode = '" & Trim(NotifCode) & "'"

            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    Ems.Subject = Trim(.Item("Description")) & " " & pAllNo

                    'CEK KETERANGAN NOTIFIKASI
                    '=========================
                    If Trim(.Item("Line1Cls")) = "05" Then
                        If Trim(.Item("Line1")) <> "" Then
                            ketLine1 = .Item("Line1") & ""
                        Else
                            ketLine1 = ""
                        End If
                    ElseIf Trim(.Item("Line1Cls")) = "06" Then
                        ketLine1 = pURL
                    ElseIf Trim(.Item("Line1Cls")) = "01" Or Trim(.Item("Line1Cls")) = "02" _
                        Or Trim(.Item("Line1Cls")) = "03" Or Trim(.Item("Line1Cls")) = "04" Then
                        ketLine1 = .Item("Line1") & " " & pAllNo
                    Else
                        ketLine1 = ""
                    End If

                    If Trim(.Item("Line2Cls")) = "05" Then
                        If Trim(.Item("Line2")) <> "" Then
                            ketLine2 = .Item("Line2") & ""
                        Else
                            ketLine2 = ""
                        End If
                    ElseIf Trim(.Item("Line2Cls")) = "06" Then
                        ketLine2 = pURL
                    ElseIf Trim(.Item("Line2Cls")) = "01" Or Trim(.Item("Line2Cls")) = "02" _
                        Or Trim(.Item("Line2Cls")) = "03" Or Trim(.Item("Line2Cls")) = "04" Then
                        ketLine2 = .Item("Line2") & " " & pAllNo
                    Else
                        ketLine2 = ""
                    End If

                    If Trim(.Item("Line3Cls")) = "05" Then
                        If Trim(.Item("Line3")) <> "" Then
                            ketLine3 = .Item("Line3") & ""
                        Else
                            ketLine3 = ""
                        End If
                    ElseIf Trim(.Item("Line3Cls")) = "06" Then
                        ketLine3 = pURL
                    ElseIf Trim(.Item("Line3Cls")) = "01" Or Trim(.Item("Line3Cls")) = "02" _
                        Or Trim(.Item("Line3Cls")) = "03" Or Trim(.Item("Line3Cls")) = "04" Then
                        ketLine3 = .Item("Line3") & " " & pAllNo
                    Else
                        ketLine3 = ""
                    End If

                    If Trim(.Item("Line4Cls")) = "05" Then
                        If Trim(.Item("Line4")) <> "" Then
                            ketLine4 = .Item("Line4") & ""
                        Else
                            ketLine4 = ""
                        End If
                    ElseIf Trim(.Item("Line4Cls")) = "06" Then
                        ketLine4 = pURL
                    ElseIf Trim(.Item("Line4Cls")) = "01" Or Trim(.Item("Line4Cls")) = "02" _
                        Or Trim(.Item("Line4Cls")) = "03" Or Trim(.Item("Line4Cls")) = "04" Then
                        ketLine4 = .Item("Line4") & " " & pAllNo
                    Else
                        ketLine4 = ""
                    End If

                    If Trim(.Item("Line5Cls")) = "05" Then
                        If Trim(.Item("Line5")) <> "" Then
                            ketLine5 = .Item("Line5") & ""
                        Else
                            ketLine5 = ""
                        End If
                    ElseIf Trim(.Item("Line5Cls")) = "06" Then
                        ketLine5 = pURL
                    ElseIf Trim(.Item("Line5Cls")) = "01" Or Trim(.Item("Line5Cls")) = "02" _
                        Or Trim(.Item("Line5Cls")) = "03" Or Trim(.Item("Line5Cls")) = "04" Then
                        ketLine5 = .Item("Line5") & " " & pAllNo
                    Else
                        ketLine5 = ""
                    End If

                    If Trim(.Item("Line6Cls")) = "05" Then
                        If Trim(.Item("Line6")) <> "" Then
                            ketLine6 = .Item("Line6") & ""
                        Else
                            ketLine6 = ""
                        End If
                    ElseIf Trim(.Item("Line6Cls")) = "06" Then
                        ketLine6 = pURL
                    ElseIf Trim(.Item("Line6Cls")) = "01" Or Trim(.Item("Line6Cls")) = "02" _
                        Or Trim(.Item("Line6Cls")) = "03" Or Trim(.Item("Line6Cls")) = "04" Then
                        ketLine6 = .Item("Line6") & " " & pAllNo
                    Else
                        ketLine6 = ""
                    End If

                    If Trim(.Item("Line7Cls")) = "05" Then
                        If Trim(.Item("Line7")) <> "" Then
                            ketLine7 = .Item("Line7") & ""
                        Else
                            ketLine7 = ""
                        End If
                    ElseIf Trim(.Item("Line7Cls")) = "06" Then
                        ketLine7 = pURL
                    ElseIf Trim(.Item("Line7Cls")) = "01" Or Trim(.Item("Line7Cls")) = "02" _
                        Or Trim(.Item("Line7Cls")) = "03" Or Trim(.Item("Line7Cls")) = "04" Then
                        ketLine7 = .Item("Line7") & " " & pAllNo
                    Else
                        ketLine7 = ""
                    End If

                    If Trim(.Item("Line8Cls")) = "05" Then
                        If Trim(.Item("Line8")) <> "" Then
                            ketLine8 = .Item("Line8") & ""
                        Else
                            ketLine8 = ""
                        End If
                    ElseIf Trim(.Item("Line8Cls")) = "06" Then
                        ketLine8 = pURL
                    ElseIf Trim(.Item("Line8Cls")) = "01" Or Trim(.Item("Line8Cls")) = "02" _
                        Or Trim(.Item("Line8Cls")) = "03" Or Trim(.Item("Line8Cls")) = "04" Then
                        ketLine8 = .Item("Line8") & " " & pAllNo
                    Else
                        ketLine8 = ""
                    End If
                    '=========================

                    'GABUNG NOTIFIKASI
                    '=========================
                    If ketLine1 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine1
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine1
                        End If
                    End If
                    If ketLine2 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine2
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine2 & vbCrLf
                        End If
                    End If
                    If ketLine3 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine3
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine3 & vbCrLf
                        End If
                    End If
                    If ketLine4 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine4
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine4
                        End If
                    End If
                    If ketLine5 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine5
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine5
                        End If
                    End If
                    If ketLine6 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine6
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine6
                        End If
                    End If
                    If ketLine7 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine7
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine7
                        End If
                    End If
                    If ketLine8 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine8
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine8
                        End If
                    End If
                    '=========================

                    Return Ems
                End With
            End If
        End Using
    End Function

    Public Shared Function GetEmailSubjectBodyEx(ByVal pConStr As String, ByVal templateCode As String, ByVal pURL As String, _
                                               ByVal pAllNo As String, ByVal sts As Boolean) As clsEmail

        Dim ketLine1 As String = "", ketLine2 As String = "", ketLine3 As String = ""
        Dim ketLine4 As String = "", ketLine5 As String = "", ketLine6 As String = ""
        Dim ketLine7 As String = "", ketLine8 As String = ""

        Using Cn As New SqlConnection(pConStr)
            Cn.Open()
            Dim q As String = ""
            Dim NotifCode As String = ""

            If templateCode = "REC-EX" Then
                If sts = False Then NotifCode = "91"
                If sts = True Then NotifCode = "92"
                If MsgstatusSave <> "" Then NotifCode = "95" 'already exists
                If MsgstatusSave = "DATA DN NO EXIST" Then NotifCode = "94" 'Surat Jalan Tidak ada di DN
                If MsgstatusSave = "PO" Then NotifCode = "108"
                If MsgstatusSave = "QTY" Then NotifCode = "109"
                pAllNo = "(SURAT JALAN NO: " & Trim(pAllNo) & ")"
            ElseIf templateCode = "DO-EX" Then
                If sts = False Then NotifCode = "93"
                If sts = True Then NotifCode = "93"
                If MsgstatusSave <> "" Then
                    If MsgstatusSave = "BOX ALREADY" Then
                        NotifCode = "99" 'already exists Box
                    ElseIf MsgstatusSave = "BOX NO NOT MATCH" Then
                        NotifCode = "110" 'Box Not Found
                    ElseIf MsgstatusSave = "QTY MOD QTYBOX, NOT MATCH" Then
                        NotifCode = "111" 'MOQ
                    Else
                        NotifCode = "96" 'already exists
                    End If
                End If
                pAllNo = "(SJ NO. : " & Trim(pAllNo) & ")"
            ElseIf templateCode = "POEM" Or templateCode = "POEE" Then
                NotifCode = "97"
                pAllNo = "(ORDER NO. : " & Trim(pAllNo) & ")"
            ElseIf templateCode = "INV-EX" Then
                NotifCode = "54"
                pAllNo = "(INVOICE NO. : " & Trim(pAllNo) & ")"
            End If

            q = " SELECT * FROM MS_NotificationCls A WITH(NOLOCK) " & vbCrLf & _
                " LEFT JOIN MS_Notification B WITH(NOLOCK) ON A.NotificationCode = B.NotificationCode " & vbCrLf & _
                " WHERE A.NotificationCode = '" & Trim(NotifCode) & "'"

            Dim cmd As New SqlCommand(q, Cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Return Nothing
            Else
                With dt.Rows(0)
                    Dim Ems As New clsEmail
                    Ems.Subject = Trim(.Item("Description")) & " " & pAllNo

                    'CEK KETERANGAN NOTIFIKASI
                    '=========================
                    If Trim(.Item("Line1Cls")) = "05" Then
                        If Trim(.Item("Line1")) <> "" Then
                            ketLine1 = .Item("Line1") & ""
                        Else
                            ketLine1 = ""
                        End If
                    ElseIf Trim(.Item("Line1Cls")) = "06" Then
                        ketLine1 = pURL
                    ElseIf Trim(.Item("Line1Cls")) = "01" Or Trim(.Item("Line1Cls")) = "02" _
                        Or Trim(.Item("Line1Cls")) = "03" Or Trim(.Item("Line1Cls")) = "04" Then
                        ketLine1 = .Item("Line1") & " " & pAllNo
                    Else
                        ketLine1 = ""
                    End If

                    If Trim(.Item("Line2Cls")) = "05" Then
                        If Trim(.Item("Line2")) <> "" Then
                            ketLine2 = .Item("Line2") & ""
                        Else
                            ketLine2 = ""
                        End If
                    ElseIf Trim(.Item("Line2Cls")) = "06" Then
                        ketLine2 = pURL
                    ElseIf Trim(.Item("Line2Cls")) = "01" Or Trim(.Item("Line2Cls")) = "02" _
                        Or Trim(.Item("Line2Cls")) = "03" Or Trim(.Item("Line2Cls")) = "04" Then
                        ketLine2 = .Item("Line2") & " " & pAllNo
                    Else
                        ketLine2 = ""
                    End If

                    If Trim(.Item("Line3Cls")) = "05" Then
                        If Trim(.Item("Line3")) <> "" Then
                            ketLine3 = .Item("Line3") & ""
                        Else
                            ketLine3 = ""
                        End If
                    ElseIf Trim(.Item("Line3Cls")) = "06" Then
                        ketLine3 = pURL
                    ElseIf Trim(.Item("Line3Cls")) = "01" Or Trim(.Item("Line3Cls")) = "02" _
                        Or Trim(.Item("Line3Cls")) = "03" Or Trim(.Item("Line3Cls")) = "04" Then
                        ketLine3 = .Item("Line3") & " " & pAllNo
                    Else
                        ketLine3 = ""
                    End If

                    If Trim(.Item("Line4Cls")) = "05" Then
                        If Trim(.Item("Line4")) <> "" Then
                            ketLine4 = .Item("Line4") & ""
                        Else
                            ketLine4 = ""
                        End If
                    ElseIf Trim(.Item("Line4Cls")) = "06" Then
                        ketLine4 = pURL
                    ElseIf Trim(.Item("Line4Cls")) = "01" Or Trim(.Item("Line4Cls")) = "02" _
                        Or Trim(.Item("Line4Cls")) = "03" Or Trim(.Item("Line4Cls")) = "04" Then
                        ketLine4 = .Item("Line4") & " " & pAllNo
                    Else
                        ketLine4 = ""
                    End If

                    If Trim(.Item("Line5Cls")) = "05" Then
                        If Trim(.Item("Line5")) <> "" Then
                            ketLine5 = .Item("Line5") & ""
                        Else
                            ketLine5 = ""
                        End If
                    ElseIf Trim(.Item("Line5Cls")) = "06" Then
                        ketLine5 = pURL
                    ElseIf Trim(.Item("Line5Cls")) = "01" Or Trim(.Item("Line5Cls")) = "02" _
                        Or Trim(.Item("Line5Cls")) = "03" Or Trim(.Item("Line5Cls")) = "04" Then
                        ketLine5 = .Item("Line5") & " " & pAllNo
                    Else
                        ketLine5 = ""
                    End If

                    If Trim(.Item("Line6Cls")) = "05" Then
                        If Trim(.Item("Line6")) <> "" Then
                            ketLine6 = .Item("Line6") & ""
                        Else
                            ketLine6 = ""
                        End If
                    ElseIf Trim(.Item("Line6Cls")) = "06" Then
                        ketLine6 = pURL
                    ElseIf Trim(.Item("Line6Cls")) = "01" Or Trim(.Item("Line6Cls")) = "02" _
                        Or Trim(.Item("Line6Cls")) = "03" Or Trim(.Item("Line6Cls")) = "04" Then
                        ketLine6 = .Item("Line6") & " " & pAllNo
                    Else
                        ketLine6 = ""
                    End If

                    If Trim(.Item("Line7Cls")) = "05" Then
                        If Trim(.Item("Line7")) <> "" Then
                            ketLine7 = .Item("Line7") & ""
                        Else
                            ketLine7 = ""
                        End If
                    ElseIf Trim(.Item("Line7Cls")) = "06" Then
                        ketLine7 = pURL
                    ElseIf Trim(.Item("Line7Cls")) = "01" Or Trim(.Item("Line7Cls")) = "02" _
                        Or Trim(.Item("Line7Cls")) = "03" Or Trim(.Item("Line7Cls")) = "04" Then
                        ketLine7 = .Item("Line7") & " " & pAllNo
                    Else
                        ketLine7 = ""
                    End If

                    If Trim(.Item("Line8Cls")) = "05" Then
                        If Trim(.Item("Line8")) <> "" Then
                            ketLine8 = .Item("Line8") & ""
                        Else
                            ketLine8 = ""
                        End If
                    ElseIf Trim(.Item("Line8Cls")) = "06" Then
                        ketLine8 = pURL
                    ElseIf Trim(.Item("Line8Cls")) = "01" Or Trim(.Item("Line8Cls")) = "02" _
                        Or Trim(.Item("Line8Cls")) = "03" Or Trim(.Item("Line8Cls")) = "04" Then
                        ketLine8 = .Item("Line8") & " " & pAllNo
                    Else
                        ketLine8 = ""
                    End If
                    '=========================

                    'GABUNG NOTIFIKASI
                    '=========================
                    If ketLine1 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine1
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine1
                        End If
                    End If
                    If ketLine2 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine2
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine2 & vbCrLf
                        End If
                    End If
                    If ketLine3 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine3
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine3 & vbCrLf
                        End If
                    End If
                    If ketLine4 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine4
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine4
                        End If
                    End If
                    If ketLine5 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine5
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine5
                        End If
                    End If
                    If ketLine6 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine6
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine6
                        End If
                    End If
                    If ketLine7 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine7
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine7
                        End If
                    End If
                    If ketLine8 <> "" Then
                        If Ems.Body = "" Then
                            Ems.Body = ketLine8
                        Else
                            Ems.Body = Ems.Body & vbCrLf & ketLine8
                        End If
                    End If
                    '=========================

                    Return Ems
                End With
            End If
        End Using
    End Function
End Class

