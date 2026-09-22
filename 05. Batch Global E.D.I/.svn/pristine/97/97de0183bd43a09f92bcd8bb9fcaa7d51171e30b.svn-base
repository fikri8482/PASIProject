Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

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

    Shared Function Supplier(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function Affiliate(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function DeliveryLocation(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT * FROM dbo.MS_DeliveryPlace WHERE DeliveryLocationCode='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function Forwarder(ByVal GB As GlobalSetting.clsGlobal, ByVal ls_value As String) As DataSet
        Dim ls_SQL As String = ""
        ls_SQL = "SELECT * FROM dbo.MS_Forwarder WHERE ForwarderID='" & ls_value & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function EmailToEmailCC(ByVal GB As GlobalSetting.clsGlobal, ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                " select 'AFF' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailAffiliate where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
                " union all " & vbCrLf & _
                " --PASI TO -CC " & vbCrLf & _
                " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                " union all " & vbCrLf & _
                " --Supplier TO- CC " & vbCrLf & _
                " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail='' from ms_emailSupplier where SupplierID='" & Trim(pSupplierID) & "'"
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
    End Function

    Shared Function EmailToEmailCCKanban(ByVal GB As GlobalSetting.clsGlobal, ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        ls_SQL = " select 'AFF' as FLAG , kanbanCC =  Rtrim(kanbanCC) + ';' + Rtrim(KanbanTo),kanbanTo = '' from MS_emailAffiliate where AffiliateID = '" & pAfffCode & "' " & vbCrLf & _
                " union ALL " & vbCrLf & _
                " select 'PASI' as FLAG , kanbanCC, kanbanTo = KanbanTo from MS_EmailPasi " & vbCrLf & _
                " UNION ALL " & vbCrLf & _
                " select 'SUPP' as FLAG , kanbanCC, kanbanTo = '' from MS_EmailSupplier where supplierID = '" & pSupplierID & "' "
        Dim ds As New DataSet
        ds = GB.uf_GetDataSet(ls_SQL)

        If ds.Tables(0).Rows.Count > 0 Then
            Return ds
        Else
            Return Nothing
        End If
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
                SettingEmail.Add(SetEmail)
            Next

            Return SettingEmail
        Else
            errMsg = "Process Send to Supplier STOPPED, because Email Setting Empty "
        End If
    End Function
End Class
