Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports System.Net
Imports System.Net.Mail

Public Class POExportEntry
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "B02"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean
    'Dim ls_TextFile 
    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim flag As Boolean = True
    Dim clsPO As New clsPO
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
            Session("M01Url") = Request.QueryString("Session")
            flag = False
        Else
            flag = True
        End If

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            'btnDelete.Enabled = ls_AllowDelete
            'btnSubmit.Enabled = ls_AllowUpdate
            If Session("M01Url") <> "" Then
                If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                    Session("MenuDesc") = "PO ENTRY"
                    pub_PONo = Request.QueryString("id")
                    pub_Ship = Request.QueryString("t1")
                    pub_Commercial = Request.QueryString("t2")
                    pub_Period = Request.QueryString("t3")
                    Session("SupplierID") = Request.QueryString("t4")

                    dtPeriodFrom.Value = pub_Period
                    txtPONo.Text = pub_PONo
                    txtShip.Text = pub_Ship
                    If pub_Commercial = "YES" Then
                        rdrCom1.Checked = True
                    Else
                        rdrCom2.Checked = True
                    End If

                    Session("Mode") = "Update"

                    bindData()
                    bindPOStatus()

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    txtPONo.ReadOnly = True
                    txtPONo.BackColor = Color.FromName("#CCCCCC")
                    dtPeriodFrom.ReadOnly = True
                    dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")
                    txtShip.ReadOnly = True
                    txtShip.BackColor = Color.FromName("#CCCCCC")
                    rdrCom1.ReadOnly = True
                    rdrCom2.ReadOnly = True

                    If clsPO.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID")) = "YES" Then
                        rdrEmergency2.Checked = True
                    Else
                        rdrEmergency3.Checked = True
                    End If


                    'btnCraete.Text = "UPDATE"
                    btnClear.Enabled = False

                    If txtDate1.Text.Trim <> "" And txtDate2.Text <> "" Then
                        btnSubmit.Enabled = False
                        btnDelete.Enabled = False
                    Else
                        btnSubmit.Enabled = True
                        btnDelete.Enabled = True
                    End If

                    If txtDate2.Text.Trim = "" Then
                        btnApprove.Text = "APPROVE"
                    Else
                        btnApprove.Text = "UNAPPROVE"
                    End If

                    If txtDate3.Text.Trim <> "" Then
                        btnApprove.Enabled = False
                    Else
                        btnApprove.Enabled = True
                    End If

                ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                    Session("MenuDesc") = "PO ENTRY"
                    pub_PONo = clsNotification.DecryptURL(Request.QueryString("id2"))
                    pub_Ship = clsNotification.DecryptURL(Request.QueryString("t1"))
                    pub_Commercial = clsNotification.DecryptURL(Request.QueryString("t2"))
                    pub_Period = clsNotification.DecryptURL(Request.QueryString("t3"))
                    Session("SupplierID") = clsNotification.DecryptURL(Request.QueryString("t4"))

                    dtPeriodFrom.Value = pub_Period
                    txtPONo.Text = pub_PONo
                    txtShip.Text = pub_Ship
                    If pub_Commercial = "YES" Then
                        rdrCom1.Checked = True
                    Else
                        rdrCom2.Checked = True
                    End If

                    Session("Mode") = "Update"

                    bindData()
                    bindPOStatus()

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    txtPONo.ReadOnly = True
                    txtPONo.BackColor = Color.FromName("#CCCCCC")
                    dtPeriodFrom.ReadOnly = True
                    dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")
                    txtShip.ReadOnly = True
                    txtShip.BackColor = Color.FromName("#CCCCCC")
                    rdrCom1.ReadOnly = True
                    rdrCom2.ReadOnly = True

                    If clsPO.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID")) = "YES" Then
                        rdrEmergency2.Checked = True
                    Else
                        rdrEmergency3.Checked = True
                    End If

                    btnClear.Enabled = False

                    If txtDate1.Text.Trim <> "" And txtDate2.Text <> "" Then
                        btnSubmit.Enabled = False
                        btnDelete.Enabled = False
                    Else
                        btnSubmit.Enabled = True
                        btnDelete.Enabled = True
                    End If

                    If txtDate2.Text.Trim = "" Then
                        btnApprove.Text = "APPROVE"
                    Else
                        btnApprove.Text = "UNAPPROVE"
                    End If

                    If txtDate3.Text.Trim <> "" Then
                        btnApprove.Enabled = False
                    Else
                        btnApprove.Enabled = True
                    End If
                Else
                    Session("MenuDesc") = "PO ENTRY"
                    Session("Mode") = "New"
                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    txtPONo.Focus()
                    dtPeriodFrom.Value = Now
                    rdrCom1.Checked = True
                    rdrEmergency2.Checked = True
                    bindData()
                    'btnApprove.Enabled = False
                    'btnSubmit.Enabled = True
                    'btnDelete.Enabled = True
                    'btnClear.Enabled = True
                End If
            Else
                Session("Mode") = "New"
                txtPONo.Focus()
                'btnApprove.Enabled = False
                'btnSubmit.Enabled = True
                'btnDelete.Enabled = True
                'btnClear.Enabled = True
                dtPeriodFrom.Value = Now
                rdrCom1.Checked = True
                rdrEmergency2.Checked = True
                bindData()
            End If

            bindData()
            ColorGrid()
            lblInfo.Text = ""

        ElseIf IsCallback Then
            If grid.VisibleRowCount = 0 Then Exit Sub
        End If

        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 4, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" _
             Or e.Column.FieldName = "UnitDesc" Or e.Column.FieldName = "MinOrderQty" Or e.Column.FieldName = "Maker" _
             Or e.Column.FieldName = "KanbanCls" Or e.Column.FieldName = "POQty" Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "CurrDesc" _
             Or e.Column.FieldName = "Price" Or e.Column.FieldName = "Amount" Or e.Column.FieldName = "QtyBox" _
             Or e.Column.FieldName = "ForecastN1" Or e.Column.FieldName = "ForecastN2" Or e.Column.FieldName = "ForecastN3") _
             And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/PurchaseOrder/POList.aspx")
        Else
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        grid.JSProperties("cpMessage") = ""
        Call bindData()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 4, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"

                    'If Session("B02IsSubmit") = "true" Then
                    '    grid.PageIndex = 0
                    '    Session.Remove("B02IsSubmit")
                    '    txtPONo.Text = Session("PONo")
                    '    grid.JSProperties("cpPONo") = Session("PONo")
                    '    Session.Remove("PONo")

                    '    grid.JSProperties("cpDate1") = Session("cpDate1")
                    '    Session.Remove("cpDate1")
                    '    grid.JSProperties("cpUser1") = Session("cpUser1")
                    '    Session.Remove("cpUser1")
                    'End If

                    'Call bindData()

                    'If grid.VisibleRowCount = 0 Then
                    '    Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    'End If
                    If Session("pub_Save") = True Then
                        If Session("B02IsSubmit") = "true" Then
                            grid.PageIndex = 0
                            Session.Remove("B02IsSubmit")
                            txtPONo.Text = Session("PONo")
                            grid.JSProperties("cpPONo") = Session("PONo")
                            Session.Remove("PONo")

                            grid.JSProperties("cpDate1") = Session("cpDate1")
                            Session.Remove("cpDate1")




                            grid.JSProperties("cpUser1") = Session("cpUser1")
                            Session.Remove("cpUser1")
                        End If

                        Call bindData()

                        If grid.VisibleRowCount = 0 Then
                            Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                            grid.JSProperties("cpMessage") = lblInfo.Text
                        End If
                    Else
                        Call bindData()

                        If grid.VisibleRowCount = 0 Then
                            Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                            grid.JSProperties("cpMessage") = lblInfo.Text
                        End If
                    End If

                    Session("pub_Save") = False
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "loadHeijunka"
                    Call bindHeijunka()
                Case "savedata"
                    Call saveData()
                Case "saveApprove"
                    Call uf_Approve()
                    Call bindPOStatus()
                Case "aftersave"
                    bindHeijunka()
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)

        If txtDate1.Text <> "" Then
            If pAction = "APPROVE" Then
                uf_Approve()
                'sendEmail()
                'sendEmailccPASI()
                'sendEmailtoAffiliate()
                btnApprove.Text = "UNAPPROVE"
                ButtonApprove.JSProperties("cpButton") = "UNAPPROVE"
            Else
                uf_UnApprove()
                btnApprove.Text = "APPROVE"
                ButtonApprove.JSProperties("cpButton") = "APPROVE"
            End If

            bindPOStatus("update")
        End If
        'sendEmailtoAffiliate()
        'sendEmailccPASI()
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 0
        Dim pIsUpdate As Boolean

        Dim ls_PONo As String = "", ls_Ship As String = "", ls_PartNo As String = "", ls_KanbanCls As String = ""
        Dim ls_Maker As String = "", ls_POQty As Double = 0, ls_CurrCls As String = "", ls_Price As Double = 0, ls_Amount As Double = 0
        Dim ls_Forecast1 As Double = 0, ls_Forecast2 As Double = 0, ls_Forecast3 As Double = 0
        Dim ls_D1 As Double = 0, ls_D2 As Double = 0, ls_D3 As Double = 0, ls_D4 As Double = 0, ls_D5 As Double = 0
        Dim ls_D6 As Double = 0, ls_D7 As Double = 0, ls_D8 As Double = 0, ls_D9 As Double = 0, ls_D10 As Double = 0
        Dim ls_D11 As Double = 0, ls_D12 As Double = 0, ls_D13 As Double = 0, ls_D14 As Double = 0, ls_D15 As Double = 0
        Dim ls_D16 As Double = 0, ls_D17 As Double = 0, ls_D18 As Double = 0, ls_D19 As Double = 0, ls_D20 As Double = 0
        Dim ls_D21 As Double = 0, ls_D22 As Double = 0, ls_D23 As Double = 0, ls_D24 As Double = 0, ls_D25 As Double = 0
        Dim ls_D26 As Double = 0, ls_D27 As Double = 0, ls_D28 As Double = 0, ls_D29 As Double = 0, ls_D30 As Double = 0
        Dim ls_D31 As Double = 0

        Dim ls_MOQ As Double = 0

        Dim ls_AffiliateID As String = Session("AffiliateID")
        Dim ls_SupplierID As String = ""
        Dim ls_TempSupplierID As String = ""
        Dim ls_PODeliveryBY As String = ""

        Dim ls_TotalCurr As String = "", ls_TotalAmount As Double = 0

        Dim a As Integer, xy As Integer = 0
        'Dim ls_tampungPO(10) As String
        'Dim wiplist As New List(Of clsPO)

        Dim sqlstring As String = ""
        Dim publi_PONO As String = ""

        Dim flgTempSupplier As String = ""
        Dim flgSupplier As Boolean = False
        'Dim ls_Seq As Integer = 0

        Session("pub_Save") = False

        'If Session("Mode") = "New" Then
        '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
        '        sqlConn.Open()
        '        ls_SQL = "select * from PO_Master where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "'"

        '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
        '        Dim ds As New DataSet
        '        sqlDA.Fill(ds)

        '        If ds.Tables(0).Rows.Count > 0 Then
        '            Call clsMsg.DisplayMessage(lblInfo, "5012", clsMessage.MsgType.ErrorMessage)
        '            Session("YA010IsSubmit") = lblInfo.Text
        '            grid.JSProperties("cpMessage") = lblInfo.Text
        '            Exit Sub
        '        End If

        '        sqlConn.Close()
        '    End Using
        'End If

        ' FOR VALIDATION
        a = e.UpdateValues.Count
        For iLoop = 0 To a - 1
            ls_Active = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())

            If flgTempSupplier <> e.UpdateValues(iLoop).NewValues("SupplierID").ToString() And flgTempSupplier <> "" Then
                flgSupplier = True
            End If

            flgTempSupplier = (e.UpdateValues(iLoop).NewValues("SupplierID").ToString())

            If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

            If ls_Active = "1" Then
                If e.UpdateValues(iLoop).NewValues("POQty") = 0 Then
                    Call clsMsg.DisplayMessage(lblInfo, "5001", clsMessage.MsgType.ErrorMessage)
                    Session("YA010IsSubmit") = lblInfo.Text
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Session("pub_Save") = False
                    Exit Sub
                End If
                If (e.UpdateValues(iLoop).NewValues("POQty") Mod e.UpdateValues(iLoop).NewValues("MinOrderQty")) <> 0 Then
                    Call clsMsg.DisplayMessage(lblInfo, "5005", clsMessage.MsgType.ErrorMessage)
                    Session("YA010IsSubmit") = lblInfo.Text
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Session("pub_Save") = False
                    Exit Sub
                End If
                'If IsNothing(e.UpdateValues(iLoop).NewValues("CurrCls")) Or IsDBNull(e.UpdateValues(iLoop).NewValues("CurrCls")) Or e.UpdateValues(iLoop).NewValues("CurrCls") = "" Then
                '    Call clsMsg.DisplayMessage(lblInfo, "5011", clsMessage.MsgType.ErrorMessage)
                '    Session("YA010IsSubmit") = lblInfo.Text
                '    grid.JSProperties("cpMessage") = lblInfo.Text
                '    Exit Sub
                'End If

                If txtHeijunka.Text <> "Heijunka" Then
                    Dim checkTotalDailyQty As Double = 0

                    For i = 1 To 31
                        checkTotalDailyQty = checkTotalDailyQty + e.UpdateValues(iLoop).NewValues("DeliveryD" & i)
                    Next

                    If checkTotalDailyQty <> e.UpdateValues(iLoop).NewValues("POQty") Then
                        Call clsMsg.DisplayMessage(lblInfo, "5006", clsMessage.MsgType.ErrorMessage)
                        Session("YA010IsSubmit") = lblInfo.Text
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("pub_Save") = False
                        Exit Sub
                    End If

                    For i = 1 To 31
                        If e.UpdateValues(iLoop).NewValues("DeliveryD" & i) <> 0 Then
                            If (e.UpdateValues(iLoop).NewValues("DeliveryD" & i) Mod e.UpdateValues(iLoop).NewValues("QtyBox")) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "5007", clsMessage.MsgType.ErrorMessage)
                                Session("YA010IsSubmit") = lblInfo.Text
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Session("pub_Save") = False
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        Next

        'Checking PONo already Exists or not
        If Session("mode") = "New" Then
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Dim ls_wherePONO As String = ""

                'If flgSupplier = False Then
                ls_wherePONO = "PONo ='" & txtPONo.Text & "'"
                'ElseIf flgSupplier = True Then
                '    ls_wherePONO = "SUBSTRING(PONo,1," & txtPONo.Text.Trim.Length & ") ='" & txtPONo.Text & "'"
                'End If

                sqlstring = "SELECT * FROM dbo.PO_Master WHERE " & ls_wherePONO & " AND AffiliateID = '" & Session("AffiliateID") & "' "

                Dim sqlDA As New SqlDataAdapter(sqlstring, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    ls_MsgID = "5012"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("YA010IsSubmit") = lblInfo.Text
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Session("pub_Save") = False
                    Exit Sub
                End If

                sqlConn.Close()

            End Using
        End If

        'FOR HEIJUNKA
        ''If txtHeijunka.Text = "Heijunka" Then
        ''    If e.UpdateValues.Count = 0 Then
        ''        Exit Sub
        ''    End If
        ''    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
        ''        sqlConn.Open()

        ''        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("item_code")

        ''            If grid.VisibleRowCount = 0 Then
        ''                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 4, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
        ''                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord, False)
        ''                Exit Sub
        ''            End If

        ''            Dim pub_Count As Boolean = False

        ''            a = e.UpdateValues.Count

        ''            ls_SQL = "  DELETE from dbo.tempPO_Detail" & vbCrLf & _
        ''                    "  where PONo = '" & txtPONo.Text & "'" & vbCrLf & _
        ''                    "  and AffiliateID = '" & Session("AffiliateID") & "' "

        ''            Dim SqlComm6 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
        ''            SqlComm6.ExecuteNonQuery()
        ''            SqlComm6.Dispose()
        ''            'sqlTran.Commit()

        ''            For iLoop = 0 To a - 1
        ''                ls_Active = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())

        ''                If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

        ''                ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
        ''                ls_POQty = Trim(e.UpdateValues(iLoop).NewValues("POQty").ToString())

        ''                Dim sqlComm As New SqlCommand()
        ''                ls_SQL = " INSERT INTO [dbo].[tempPO_Detail] " & vbCrLf & _
        ''                          "            ([AllowAcess] " & vbCrLf & _
        ''                          "            ,[PONo] " & vbCrLf & _
        ''                          "            ,[AffiliateID] " & vbCrLf & _
        ''                          "            ,[PartNo] " & vbCrLf & _
        ''                          "            ,[POQty]) " & vbCrLf & _
        ''                          "      VALUES " & vbCrLf & _
        ''                          "            ('1' " & vbCrLf & _
        ''                          "            ,'" & txtPONo.Text & "' " & vbCrLf & _
        ''                          "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
        ''                          "            ,'" & ls_PartNo & "' " & vbCrLf & _
        ''                          "            ,'" & ls_POQty & "' ) "
        ''                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
        ''                sqlComm.ExecuteNonQuery()
        ''                sqlComm.Dispose()
        ''            Next iLoop

        ''            sqlTran.Commit()
        ''            'Session("B02IsSubmit") = "true"
        ''            'Session("PONo") = publi_PONO
        ''        End Using

        ''        sqlConn.Close()
        ''    End Using

        ''    Exit Sub
        ''End If

        'for input remaining Qty
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("remainingItem")

                a = e.UpdateValues.Count

                For iLoop = 0 To a - 1

                    ls_SQL = " select a.SupplierID, isnull(b.MonthlyProductionCapacity,0)MonthlyProductionCapacity from [dbo].[MS_PartMapping] a " & vbCrLf & _
                              " left join [dbo].[MS_SupplierCapacity] b on a.SupplierID = b.SupplierID and a.PartNo = b.PartNo" & vbCrLf & _
                              " WHERE a.PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString()) & "' AND a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " ORDER BY a.SupplierID  " & vbCrLf

                    Dim sqlCommNew As SqlCommand = sqlConn.CreateCommand
                    sqlCommNew.Connection = sqlConn
                    sqlCommNew.Transaction = sqlTran

                    sqlCommNew.CommandText = ls_SQL
                    Dim da As New SqlDataAdapter(sqlCommNew)
                    Dim ds As New DataSet
                    da.Fill(ds)

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        Dim lk_SupplierID As String = ds.Tables(0).Rows(0)("SupplierID").ToString.Trim

                        sqlstring = "SELECT * FROM dbo.RemainingCapacity WHERE Period ='" & Format(dtPeriodFrom.Value, "yyyyMM") & "' AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString()) & "' AND SupplierID = '" & lk_SupplierID & "'"

                        Dim sqlComm As New SqlCommand(sqlstring, sqlConn, sqlTran)
                        sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                        Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                        If sqlRdr.Read Then
                            pIsUpdate = True
                        Else
                            pIsUpdate = False
                        End If
                        sqlRdr.Close()
                        If pIsUpdate = False Then
                            'INSERT DATA
                            ls_SQL = " INSERT INTO [dbo].[RemainingCapacity] " & vbCrLf & _
                                  "            ([Period] " & vbCrLf & _
                                  "            ,[PartNo] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[QtyRemaining]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & Format(dtPeriodFrom.Value, "yyyyMM") & "' " & vbCrLf & _
                                  "            ,'" & Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString()) & "' " & vbCrLf & _
                                  "            ,'" & lk_SupplierID & "'" & vbCrLf & _
                                  "            ,'" & ds.Tables(0).Rows(0)("MonthlyProductionCapacity").ToString.Trim & "' ) "
                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()
                        End If
                    Next
                Next

                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using

        'If flgSupplier Then
        '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
        '        sqlConn.Open()
        '        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("item_code")

        '            ls_SQL = "delete tempPODetail where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "'"

        '            Dim SqlComm6 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
        '            SqlComm6.ExecuteNonQuery()
        '            SqlComm6.Dispose()
        '            sqlTran.Commit()
        '        End Using

        '        sqlConn.Close()
        '    End Using
        'End If

        '''''NORMAL CASE NOT FOR HEIJUNKA'''''''''s
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("item_code")

                If grid.VisibleRowCount = 0 Then
                    'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowPager, False, False)
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord, False)
                    Exit Sub
                End If

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Exit Sub
                End If

                ''01. Check 1 part supply from 2 Supplier

                Dim pub_Count As Boolean = False

                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1
                    ls_Active = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())

                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                    ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
                    ls_SupplierID = Trim(e.UpdateValues(iLoop).NewValues("SupplierID").ToString())
                    ls_MOQ = Trim(e.UpdateValues(iLoop).NewValues("MinOrderQty").ToString())

                    If rdrEmergency2.Checked = True Then
                        ls_KanbanCls = 1
                    Else
                        ls_KanbanCls = 0
                    End If

                    'ls_Maker = Trim(e.UpdateValues(iLoop).NewValues("Maker").ToString())

                    If IsDBNull(e.UpdateValues(iLoop).NewValues("CurrCls")) Or IsNothing(e.UpdateValues(iLoop).NewValues("CurrCls")) Then
                        ls_CurrCls = "NULL"
                        ls_Price = 0
                        ls_Amount = 0
                    Else
                        ls_CurrCls = "'" & Trim(e.UpdateValues(iLoop).NewValues("CurrCls").ToString()) & "'"
                        ls_Price = Trim(e.UpdateValues(iLoop).NewValues("Price").ToString())
                        ls_Amount = Trim(e.UpdateValues(iLoop).NewValues("Amount").ToString())
                    End If

                    ls_POQty = Trim(e.UpdateValues(iLoop).NewValues("POQty").ToString())

                    ls_D1 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD1").ToString())
                    ls_D2 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD2").ToString())
                    ls_D3 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD3").ToString())
                    ls_D4 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD4").ToString())
                    ls_D5 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD5").ToString())
                    ls_D6 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD6").ToString())
                    ls_D7 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD7").ToString())
                    ls_D8 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD8").ToString())
                    ls_D9 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD9").ToString())
                    ls_D10 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD10").ToString())
                    ls_D11 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD11").ToString())
                    ls_D12 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD12").ToString())
                    ls_D13 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD13").ToString())
                    ls_D14 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD14").ToString())
                    ls_D15 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD15").ToString())
                    ls_D16 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD16").ToString())
                    ls_D17 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD17").ToString())
                    ls_D18 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD18").ToString())
                    ls_D19 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD19").ToString())
                    ls_D20 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD20").ToString())
                    ls_D21 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD21").ToString())
                    ls_D22 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD22").ToString())
                    ls_D23 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD23").ToString())
                    ls_D24 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD24").ToString())
                    ls_D25 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD25").ToString())
                    ls_D26 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD26").ToString())
                    ls_D27 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD27").ToString())
                    ls_D28 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD28").ToString())
                    ls_D29 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD29").ToString())
                    ls_D30 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD30").ToString())
                    ls_D31 = Trim(e.UpdateValues(iLoop).NewValues("DeliveryD31").ToString())
                    ls_PODeliveryBY = Trim(e.UpdateValues(iLoop).NewValues("PODeliveryBy").ToString())

                    ls_TotalAmount = ls_TotalAmount + ls_Amount
                    ls_TotalCurr = ls_CurrCls

                    'If iLoop = 0 And flgSupplier = True Then
                    '    ls_PONo = txtPONo.Text & "-" & ls_Seq + 1
                    'ElseIf flgSupplier = False Then
                    '    ls_PONo = txtPONo.Text
                    'End If

                    If Session("mode") = "New" Then
                        Dim ls_table As String = ""

                        'If flgSupplier = True Then
                        '    ls_table = "tempPODetail"
                        'Else
                        ls_table = "PO_Detail"
                        'End If

                        ls_SQL = " INSERT INTO [dbo].[" & ls_table & "] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[KanbanCls] " & vbCrLf & _
                                      "            ,[Maker] " & vbCrLf & _
                                      "            ,[POQty] " & vbCrLf & _
                                      "            ,[CurrCls] " & vbCrLf & _
                                      "            ,[Price] " & vbCrLf & _
                                      "            ,[Amount] "

                        ls_SQL = ls_SQL + "            ,[DeliveryD1] " & vbCrLf & _
                                          "            ,[DeliveryD2] " & vbCrLf & _
                                          "            ,[DeliveryD3] " & vbCrLf & _
                                          "            ,[DeliveryD4] " & vbCrLf & _
                                          "            ,[DeliveryD5] " & vbCrLf & _
                                          "            ,[DeliveryD6] " & vbCrLf & _
                                          "            ,[DeliveryD7] " & vbCrLf & _
                                          "            ,[DeliveryD8] "

                        ls_SQL = ls_SQL + "            ,[DeliveryD9] " & vbCrLf & _
                                          "            ,[DeliveryD10] " & vbCrLf & _
                                          "            ,[DeliveryD11] " & vbCrLf & _
                                          "            ,[DeliveryD12] " & vbCrLf & _
                                          "            ,[DeliveryD13] " & vbCrLf & _
                                          "            ,[DeliveryD14] " & vbCrLf & _
                                          "            ,[DeliveryD15] " & vbCrLf & _
                                          "            ,[DeliveryD16] " & vbCrLf & _
                                          "            ,[DeliveryD17] " & vbCrLf & _
                                          "            ,[DeliveryD18] " & vbCrLf & _
                                          "            ,[DeliveryD19] "

                        ls_SQL = ls_SQL + "            ,[DeliveryD20] " & vbCrLf & _
                                          "            ,[DeliveryD21] " & vbCrLf & _
                                          "            ,[DeliveryD22] " & vbCrLf & _
                                          "            ,[DeliveryD23] " & vbCrLf & _
                                          "            ,[DeliveryD24] " & vbCrLf & _
                                          "            ,[DeliveryD25] " & vbCrLf & _
                                          "            ,[DeliveryD26] " & vbCrLf & _
                                          "            ,[DeliveryD27] " & vbCrLf & _
                                          "            ,[DeliveryD28] " & vbCrLf & _
                                          "            ,[DeliveryD29] " & vbCrLf & _
                                          "            ,[DeliveryD30] "

                        ls_SQL = ls_SQL + "            ,[DeliveryD31] " & vbCrLf & _
                                          "            ,[EntryDate] " & vbCrLf & _
                                          "            ,[EntryUser]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                          "            ,'" & ls_AffiliateID & "' " & vbCrLf & _
                                          "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                                          "            ,'" & ls_PartNo & "' " & vbCrLf & _
                                          "            ,'" & ls_KanbanCls & "' "

                        ls_SQL = ls_SQL + "            , NULL " & vbCrLf & _
                                          "            ,'" & ls_POQty & "' " & vbCrLf & _
                                          "            ," & ls_CurrCls & " " & vbCrLf & _
                                          "            ,'" & ls_Price & "' " & vbCrLf & _
                                          "            ,'" & ls_Price * ls_POQty & "' " & vbCrLf & _
                                          "            ,'" & ls_D1 & "' " & vbCrLf & _
                                          "            ,'" & ls_D2 & "' " & vbCrLf & _
                                          "            ,'" & ls_D3 & "' "

                        ls_SQL = ls_SQL + "            ,'" & ls_D4 & "' " & vbCrLf & _
                                          "            ,'" & ls_D5 & "' " & vbCrLf & _
                                          "            ,'" & ls_D6 & "' " & vbCrLf & _
                                          "            ,'" & ls_D7 & "' " & vbCrLf & _
                                          "            ,'" & ls_D8 & "' " & vbCrLf & _
                                          "            ,'" & ls_D9 & "' " & vbCrLf & _
                                          "            ,'" & ls_D10 & "' " & vbCrLf & _
                                          "            ,'" & ls_D11 & "' " & vbCrLf & _
                                          "            ,'" & ls_D12 & "' " & vbCrLf & _
                                          "            ,'" & ls_D13 & "' " & vbCrLf & _
                                          "            ,'" & ls_D14 & "' "

                        ls_SQL = ls_SQL + "            ,'" & ls_D15 & "' " & vbCrLf & _
                                          "            ,'" & ls_D16 & "' " & vbCrLf & _
                                          "            ,'" & ls_D17 & "' " & vbCrLf & _
                                          "            ,'" & ls_D18 & "' " & vbCrLf & _
                                          "            ,'" & ls_D19 & "' " & vbCrLf & _
                                          "            ,'" & ls_D20 & "' " & vbCrLf & _
                                          "            ,'" & ls_D21 & "' " & vbCrLf & _
                                          "            ,'" & ls_D22 & "' " & vbCrLf & _
                                          "            ,'" & ls_D23 & "' " & vbCrLf & _
                                          "            ,'" & ls_D24 & "' " & vbCrLf & _
                                          "            ,'" & ls_D25 & "' "

                        ls_SQL = ls_SQL + "            ,'" & ls_D26 & "' " & vbCrLf & _
                                          "            ,'" & ls_D27 & "' " & vbCrLf & _
                                          "            ,'" & ls_D28 & "' " & vbCrLf & _
                                          "            ,'" & ls_D29 & "' " & vbCrLf & _
                                          "            ,'" & ls_D30 & "' " & vbCrLf & _
                                          "            ,'" & ls_D31 & "' " & vbCrLf & _
                                          "            , getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "' ) "
                        Dim sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()

                        Dim pub_Master As Boolean = False
                        ls_SQL = "select * from PO_Master where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "'"

                        Dim sqlComm1 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        Dim sqlRdr1 As SqlDataReader = sqlComm1.ExecuteReader()
                        Session("SupplierID") = ls_SupplierID

                        If sqlRdr1.Read Then
                            pub_Master = True
                        Else
                            pub_Master = False
                        End If
                        sqlRdr1.Close()

                        'New
                        If pub_Master = False Then
                            ls_SQL = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[Period] " & vbCrLf & _
                                      "            ,[CommercialCls] " & vbCrLf & _
                                      "            ,[DeliveryByPASICls] " & vbCrLf & _
                                      "            ,[ShipCls] " & vbCrLf & _
                                      "            ,[CurrCls] " & vbCrLf & _
                                      "            ,[Amount] " & vbCrLf & _
                                      "            ,[EntryDate] " & vbCrLf & _
                                      "            ,[EntryUser]) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                      "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                      "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                                      "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                      "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                      "            ,'" & ls_PODeliveryBY & "' " & vbCrLf & _
                                      "            ,'" & txtShip.Text & "' "

                            ls_SQL = ls_SQL + "            ," & ls_TotalCurr & " " & vbCrLf & _
                                              "            ,(select SUM(Amount) from PO_Detail where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "') " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "') "
                        Else
                            'Update
                            ls_SQL = " UPDATE [dbo].[PO_Master] SET [Amount] = (select SUM(Amount) from PO_Detail a where a.PONo = PO_Master.PONo and a.SupplierID = PO_Master.SupplierID and a.AffiliateID = PO_Master.AffiliateID), UpdateUser = '" & Session("UserID") & "', UpdateDate = getdate() WHERE AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "' and PONo='" & txtPONo.Text & "'"

                        End If

                        sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm1.ExecuteNonQuery()
                        sqlComm1.Dispose()
                        publi_PONO = txtPONo.Text

                    ElseIf Session("mode") = "Update" Then
                        sqlstring = "SELECT * FROM dbo.PO_Detail WHERE PONo ='" & txtPONo.Text & "' AND PartNo = '" & ls_PartNo & "' AND AffiliateID = '" & Session("AffiliateID") & "' AND SupplierID = '" & ls_SupplierID & "'"

                        Dim sqlComm As New SqlCommand(sqlstring, sqlConn, sqlTran)
                        sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                        Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                        If sqlRdr.Read Then
                            pIsUpdate = True
                        Else
                            pIsUpdate = False
                        End If
                        sqlRdr.Close()

                        If ls_Active = "1" Then
                            If pIsUpdate = False Then
                                'INSERT DATA
                                ls_SQL = " INSERT INTO [dbo].[PO_Detail] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[KanbanCls] " & vbCrLf & _
                                      "            ,[Maker] " & vbCrLf & _
                                      "            ,[POQty] " & vbCrLf & _
                                      "            ,[CurrCls] " & vbCrLf & _
                                      "            ,[Price] " & vbCrLf & _
                                      "            ,[Amount] "

                                ls_SQL = ls_SQL + "            ,[DeliveryD1] " & vbCrLf & _
                                                  "            ,[DeliveryD2] " & vbCrLf & _
                                                  "            ,[DeliveryD3] " & vbCrLf & _
                                                  "            ,[DeliveryD4] " & vbCrLf & _
                                                  "            ,[DeliveryD5] " & vbCrLf & _
                                                  "            ,[DeliveryD6] " & vbCrLf & _
                                                  "            ,[DeliveryD7] " & vbCrLf & _
                                                  "            ,[DeliveryD8] "

                                ls_SQL = ls_SQL + "            ,[DeliveryD9] " & vbCrLf & _
                                                  "            ,[DeliveryD10] " & vbCrLf & _
                                                  "            ,[DeliveryD11] " & vbCrLf & _
                                                  "            ,[DeliveryD12] " & vbCrLf & _
                                                  "            ,[DeliveryD13] " & vbCrLf & _
                                                  "            ,[DeliveryD14] " & vbCrLf & _
                                                  "            ,[DeliveryD15] " & vbCrLf & _
                                                  "            ,[DeliveryD16] " & vbCrLf & _
                                                  "            ,[DeliveryD17] " & vbCrLf & _
                                                  "            ,[DeliveryD18] " & vbCrLf & _
                                                  "            ,[DeliveryD19] "

                                ls_SQL = ls_SQL + "            ,[DeliveryD20] " & vbCrLf & _
                                                  "            ,[DeliveryD21] " & vbCrLf & _
                                                  "            ,[DeliveryD22] " & vbCrLf & _
                                                  "            ,[DeliveryD23] " & vbCrLf & _
                                                  "            ,[DeliveryD24] " & vbCrLf & _
                                                  "            ,[DeliveryD25] " & vbCrLf & _
                                                  "            ,[DeliveryD26] " & vbCrLf & _
                                                  "            ,[DeliveryD27] " & vbCrLf & _
                                                  "            ,[DeliveryD28] " & vbCrLf & _
                                                  "            ,[DeliveryD29] " & vbCrLf & _
                                                  "            ,[DeliveryD30] "

                                ls_SQL = ls_SQL + "            ,[DeliveryD31] " & vbCrLf & _
                                                  "            ,[EntryDate] " & vbCrLf & _
                                                  "            ,[EntryUser]) " & vbCrLf & _
                                                  "      VALUES " & vbCrLf & _
                                                  "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                                  "            ,'" & ls_AffiliateID & "' " & vbCrLf & _
                                                  "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                                                  "            ,'" & ls_PartNo & "' " & vbCrLf & _
                                                  "            ,'" & ls_KanbanCls & "' "

                                ls_SQL = ls_SQL + "            ,'" & ls_Maker & "' " & vbCrLf & _
                                                  "            ,'" & ls_POQty & "' " & vbCrLf & _
                                                  "            ," & ls_CurrCls & " " & vbCrLf & _
                                                  "            ,'" & ls_Price & "' " & vbCrLf & _
                                                  "            ,'" & ls_Price * ls_POQty & "' " & vbCrLf & _
                                                  "            ,'" & ls_D1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D2 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D3 & "' "

                                ls_SQL = ls_SQL + "            ,'" & ls_D4 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D5 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D6 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D7 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D8 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D9 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D10 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D11 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D12 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D13 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D14 & "' "

                                ls_SQL = ls_SQL + "            ,'" & ls_D15 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D16 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D17 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D18 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D19 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D20 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D21 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D22 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D23 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D24 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D25 & "' "

                                ls_SQL = ls_SQL + "            ,'" & ls_D26 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D27 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D28 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D29 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D30 & "' " & vbCrLf & _
                                                  "            ,'" & ls_D31 & "' " & vbCrLf & _
                                                  "            , getdate() " & vbCrLf & _
                                                  "            ,'" & Session("UserID") & "' ) "
                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            Else
                                ls_SQL = " UPDATE [dbo].[PO_Detail] " & vbCrLf & _
                                          "    SET [POQty] = '" & ls_POQty & "' " & vbCrLf & _
                                          "       ,[CurrCls] = " & ls_CurrCls & " " & vbCrLf & _
                                          "       ,[Price] = '" & ls_Price & "' " & vbCrLf & _
                                          "       ,[Amount] = '" & ls_Price * ls_POQty & "' " & vbCrLf & _
                                          "       ,[DeliveryD1] = '" & ls_D1 & "' " & vbCrLf & _
                                          "       ,[DeliveryD2] = '" & ls_D2 & "' " & vbCrLf & _
                                          "       ,[DeliveryD3] = '" & ls_D3 & "' "

                                ls_SQL = ls_SQL + "       ,[DeliveryD4] = '" & ls_D4 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD5] = '" & ls_D5 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD6] = '" & ls_D6 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD7] = '" & ls_D7 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD8] = '" & ls_D8 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD9] = '" & ls_D9 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD10] = '" & ls_D10 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD11] = '" & ls_D11 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD12] = '" & ls_D12 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD13] = '" & ls_D13 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD14] = '" & ls_D14 & "' "

                                ls_SQL = ls_SQL + "       ,[DeliveryD15] = '" & ls_D15 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD16] = '" & ls_D16 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD17] = '" & ls_D17 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD18] = '" & ls_D18 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD19] = '" & ls_D19 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD20] = '" & ls_D20 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD21] = '" & ls_D21 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD22] = '" & ls_D22 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD23] = '" & ls_D23 & "'" & vbCrLf & _
                                                  "       ,[DeliveryD24] = '" & ls_D24 & "' " & vbCrLf & _
                                                  "       ,[DeliveryD25] = '" & ls_D25 & "' "

                                ls_SQL = ls_SQL + "       ,[DeliveryD26] = '" & ls_D26 & "' " & vbCrLf & _
                                                    "       ,[DeliveryD27] = '" & ls_D27 & "' " & vbCrLf & _
                                                    "       ,[DeliveryD28] = '" & ls_D28 & "' " & vbCrLf & _
                                                    "       ,[DeliveryD29] = '" & ls_D29 & "' " & vbCrLf & _
                                                    "       ,[DeliveryD30] = '" & ls_D30 & "' " & vbCrLf & _
                                                    "       ,[DeliveryD31] = '" & ls_D31 & "' " & vbCrLf & _
                                                    "       ,[UpdateDate] = getdate() " & vbCrLf & _
                                                    "       ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                                    " 	 WHERE PONo ='" & txtPONo.Text & "' AND AffiliateID = '" & ls_AffiliateID & "' AND SupplierID = '" & ls_SupplierID & "' and PartNo = '" & ls_PartNo & "'"
                                ls_MsgID = "1002"
                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            End If

                        ElseIf ls_Active = "0" And pIsUpdate = True And Session("Mode") = "Update" Then
                            ls_SQL = "  DELETE from dbo.PO_Detail" & vbCrLf & _
                                     "  where PONo = '" & ls_PONo & "'" & vbCrLf & _
                                     "  and PartNo = '" & ls_PartNo & "' " & vbCrLf & _
                                     "  and AffiliateID = '" & ls_AffiliateID & "' " & vbCrLf & _
                                     "  and SupplierID = '" & ls_SupplierID & "'"
                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()
                        End If
                    End If
                Next iLoop

                'Insert to PO_Master
                'If flgSupplier = False Then
                'Dim pub_Master As Boolean = False
                'ls_SQL = "select * from PO_Master where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "'"

                'Dim sqlComm1 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                'sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                'Dim sqlRdr1 As SqlDataReader = sqlComm1.ExecuteReader()
                'Session("SupplierID") = ls_SupplierID

                'If sqlRdr1.Read Then
                '    pub_Master = True
                'Else
                '    pub_Master = False
                'End If
                'sqlRdr1.Close()

                ''New
                'If pub_Master = False Then
                '    ls_SQL = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                '              "            ([PONo] " & vbCrLf & _
                '              "            ,[AffiliateID] " & vbCrLf & _
                '              "            ,[SupplierID] " & vbCrLf & _
                '              "            ,[Period] " & vbCrLf & _
                '              "            ,[CommercialCls] " & vbCrLf & _
                '              "            ,[DeliveryByPASICls] " & vbCrLf & _
                '              "            ,[ShipCls] " & vbCrLf & _
                '              "            ,[CurrCls] " & vbCrLf & _
                '              "            ,[Amount] " & vbCrLf & _
                '              "            ,[EntryDate] " & vbCrLf & _
                '              "            ,[EntryUser]) " & vbCrLf & _
                '              "      VALUES " & vbCrLf & _
                '              "            ('" & txtPONo.Text & "' " & vbCrLf & _
                '              "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                '              "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                '              "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                '              "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                '              "            ,'" & ls_PODeliveryBY & "' " & vbCrLf & _
                '              "            ,'" & txtShip.Text & "' "

                '    ls_SQL = ls_SQL + "            ," & ls_TotalCurr & " " & vbCrLf & _
                '                      "            ,(select SUM(Amount) from PO_Detail where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "') " & vbCrLf & _
                '                      "            ,getdate() " & vbCrLf & _
                '                      "            ,'" & Session("UserID") & "') "
                'Else
                '    'Update
                '    ls_SQL = " UPDATE [dbo].[PO_Master] SET [Amount] = (select SUM(Amount) from PO_Detail a where a.PONo = PO_Master.PONo and a.SupplierID = PO_Master.SupplierID and a.AffiliateID = PO_Master.AffiliateID), UpdateUser = '" & Session("UserID") & "', UpdateDate = getdate() WHERE AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "' and PONo='" & txtPONo.Text & "'"

                'End If

                'sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                'sqlComm1.ExecuteNonQuery()
                'sqlComm1.Dispose()
                'publi_PONO = txtPONo.Text
                'ElseIf flgSupplier = True Then
                '    Dim ls_SeqNo As Integer = 1
                '    Dim k As Integer = 0

                '    ls_SQL = " SELECT [PONo],[AffiliateID],[SupplierID],[PartNo],[KanbanCls],[Maker],[POQty],[CurrCls],[Price],[Amount] " & vbCrLf & _
                '              "       ,[DeliveryD1],[DeliveryD2],[DeliveryD3],[DeliveryD4],[DeliveryD5],[DeliveryD6],[DeliveryD7],[DeliveryD8],[DeliveryD9],[DeliveryD10] " & vbCrLf & _
                '              "       ,[DeliveryD11],[DeliveryD12],[DeliveryD13],[DeliveryD14],[DeliveryD15],[DeliveryD16],[DeliveryD17],[DeliveryD18],[DeliveryD19],[DeliveryD20] " & vbCrLf & _
                '              "       ,[DeliveryD21],[DeliveryD22],[DeliveryD23],[DeliveryD24],[DeliveryD25],[DeliveryD26],[DeliveryD27],[DeliveryD28],[DeliveryD29],[DeliveryD30] " & vbCrLf & _
                '              "       ,[DeliveryD31] " & vbCrLf & _
                '              " FROM [dbo].[tempPODetail] " & vbCrLf & _
                '              " WHERE PONo = '" & txtPONo.Text & "' AND AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                '              " ORDER BY SupplierID  " & vbCrLf

                '    Dim sqlCommNew As SqlCommand = sqlConn.CreateCommand
                '    sqlCommNew.Connection = sqlConn
                '    sqlCommNew.Transaction = sqlTran

                '    sqlCommNew.CommandText = ls_SQL
                '    Dim da As New SqlDataAdapter(sqlCommNew)
                '    Dim ds As New DataSet
                '    da.Fill(ds)

                '    ls_TotalAmount = 0
                '    ls_SeqNo = 1
                '    Dim flagTempSupplier As String = ""

                '    ls_PONo = ds.Tables(0).Rows(0)("PONo").ToString.Trim & "-1"
                '    ls_SupplierID = ds.Tables(0).Rows(0)("SupplierID").ToString.Trim 'sqlRdr1("SupplierID").ToString

                '    ls_SQL = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                '                "            ([PONo] " & vbCrLf & _
                '                "            ,[AffiliateID] " & vbCrLf & _
                '                "            ,[SupplierID] " & vbCrLf & _
                '                "            ,[Period] " & vbCrLf & _
                '                "            ,[CommercialCls] " & vbCrLf & _
                '                "            ,[DeliveryByPASICls] " & vbCrLf & _
                '                "            ,[ShipCls] " & vbCrLf & _
                '                "            ,[CurrCls] " & vbCrLf & _
                '                "            ,[Amount] " & vbCrLf & _
                '                "            ,[EntryDate] " & vbCrLf & _
                '                "            ,[EntryUser]) " & vbCrLf & _
                '                "      VALUES " & vbCrLf & _
                '                "            ('" & ls_PONo & "' " & vbCrLf & _
                '                "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                '                "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                '                "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                '                "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                '                "            ,'" & ls_PODeliveryBY & "' " & vbCrLf & _
                '                "            ,'" & txtShip.Text & "' "

                '    ls_SQL = ls_SQL + "            ,NULL " & vbCrLf & _
                '                        "            ,0" & vbCrLf & _
                '                        "            ,getdate() " & vbCrLf & _
                '                        "            ,'" & Session("UserID") & "') "

                '    sqlCommNew.CommandText = ls_SQL
                '    Dim q As Integer = sqlCommNew.ExecuteNonQuery()


                '    For k = 0 To ds.Tables(0).Rows.Count - 1
                '        If flagTempSupplier <> ds.Tables(0).Rows(k)("SupplierID") And flagTempSupplier <> "" Then
                '            ls_SeqNo = ls_SeqNo + 1
                '            ls_PONo = ds.Tables(0).Rows(k)("PONo").ToString.Trim & "-" & ls_SeqNo
                '            ls_SupplierID = ds.Tables(0).Rows(k)("SupplierID") & ""
                '            ls_SQL = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                '                        "            ([PONo] " & vbCrLf & _
                '                        "            ,[AffiliateID] " & vbCrLf & _
                '                        "            ,[SupplierID] " & vbCrLf & _
                '                        "            ,[Period] " & vbCrLf & _
                '                        "            ,[CommercialCls] " & vbCrLf & _
                '                        "            ,[DeliveryByPASICls] " & vbCrLf & _
                '                        "            ,[ShipCls] " & vbCrLf & _
                '                        "            ,[CurrCls] " & vbCrLf & _
                '                        "            ,[Amount] " & vbCrLf & _
                '                        "            ,[EntryDate] " & vbCrLf & _
                '                        "            ,[EntryUser]) " & vbCrLf & _
                '                        "      VALUES " & vbCrLf & _
                '                        "            ('" & ls_PONo & "' " & vbCrLf & _
                '                        "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                '                        "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                '                        "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                '                        "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                '                        "            ,'" & ls_PODeliveryBY & "' " & vbCrLf & _
                '                        "            ,'" & txtShip.Text & "' "

                '            ls_SQL = ls_SQL + "            ,NULL " & vbCrLf & _
                '                                "            ,0 " & vbCrLf & _
                '                                "            ,getdate() " & vbCrLf & _
                '                                "            ,'" & Session("UserID") & "') "

                '            sqlCommNew = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                '            sqlCommNew.ExecuteNonQuery()
                '            sqlCommNew.Dispose()
                '        End If

                '        ls_PartNo = ds.Tables(0).Rows(k)("PartNo") & ""
                '        'ls_SupplierID = ds.Tables(0).Rows(k)("SupplierID") & ""
                '        flagTempSupplier = ls_SupplierID
                '        ls_KanbanCls = ds.Tables(0).Rows(k)("KanbanCls") & ""

                '        'ls_Maker = Trim(e.UpdateValues(iLoop).NewValues("Maker").ToString())

                '        'If IsDBNull(e.UpdateValues(iLoop).NewValues("CurrCls")) Or IsNothing(e.UpdateValues(iLoop).NewValues("CurrCls")) Then
                '        ls_CurrCls = "NULL"
                '        ls_Price = 0
                '        ls_Amount = 0
                '        'Else
                '        '    ls_CurrCls = "'" & Trim(e.UpdateValues(iLoop).NewValues("CurrCls").ToString()) & "'"
                '        '    ls_Price = Trim(e.UpdateValues(iLoop).NewValues("Price").ToString())
                '        '    ls_Amount = Trim(e.UpdateValues(iLoop).NewValues("Amount").ToString())
                '        'End If

                '        ls_D1 = ds.Tables(0).Rows(k)("DeliveryD1") & "" ' Trim(e.UpdateValues(iLoop).NewValues("DeliveryD1").ToString())
                '        ls_D2 = ds.Tables(0).Rows(k)("DeliveryD2") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD2").ToString())
                '        ls_D3 = ds.Tables(0).Rows(k)("DeliveryD3") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD3").ToString())
                '        ls_D4 = ds.Tables(0).Rows(k)("DeliveryD4") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD4").ToString())
                '        ls_D5 = ds.Tables(0).Rows(k)("DeliveryD5") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD5").ToString())
                '        ls_D6 = ds.Tables(0).Rows(k)("DeliveryD6") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD6").ToString())
                '        ls_D7 = ds.Tables(0).Rows(k)("DeliveryD7") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD7").ToString())
                '        ls_D8 = ds.Tables(0).Rows(k)("DeliveryD8") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD8").ToString())
                '        ls_D9 = ds.Tables(0).Rows(k)("DeliveryD9") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD9").ToString())
                '        ls_D10 = ds.Tables(0).Rows(k)("DeliveryD10") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD10").ToString())
                '        ls_D11 = ds.Tables(0).Rows(k)("DeliveryD11") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD11").ToString())
                '        ls_D12 = ds.Tables(0).Rows(k)("DeliveryD12") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD12").ToString())
                '        ls_D13 = ds.Tables(0).Rows(k)("DeliveryD13") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD13").ToString())
                '        ls_D14 = ds.Tables(0).Rows(k)("DeliveryD14") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD14").ToString())
                '        ls_D15 = ds.Tables(0).Rows(k)("DeliveryD15") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD15").ToString())
                '        ls_D16 = ds.Tables(0).Rows(k)("DeliveryD16") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD16").ToString())
                '        ls_D17 = ds.Tables(0).Rows(k)("DeliveryD17") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD17").ToString())
                '        ls_D18 = ds.Tables(0).Rows(k)("DeliveryD18") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD18").ToString())
                '        ls_D19 = ds.Tables(0).Rows(k)("DeliveryD19") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD19").ToString())
                '        ls_D20 = ds.Tables(0).Rows(k)("DeliveryD20") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD20").ToString())
                '        ls_D21 = ds.Tables(0).Rows(k)("DeliveryD21") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD21").ToString())
                '        ls_D22 = ds.Tables(0).Rows(k)("DeliveryD22") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD22").ToString())
                '        ls_D23 = ds.Tables(0).Rows(k)("DeliveryD23") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD23").ToString())
                '        ls_D24 = ds.Tables(0).Rows(k)("DeliveryD24") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD24").ToString())
                '        ls_D25 = ds.Tables(0).Rows(k)("DeliveryD25") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD25").ToString())
                '        ls_D26 = ds.Tables(0).Rows(k)("DeliveryD26") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD26").ToString())
                '        ls_D27 = ds.Tables(0).Rows(k)("DeliveryD27") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD27").ToString())
                '        ls_D28 = ds.Tables(0).Rows(k)("DeliveryD28") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD28").ToString())
                '        ls_D29 = ds.Tables(0).Rows(k)("DeliveryD29") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD29").ToString())
                '        ls_D30 = ds.Tables(0).Rows(k)("DeliveryD30") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD30").ToString())
                '        ls_D31 = ds.Tables(0).Rows(k)("DeliveryD31") & "" 'Trim(e.UpdateValues(iLoop).NewValues("DeliveryD31").ToString())

                '        ls_TotalAmount = ls_TotalAmount + ls_Amount
                '        ls_TotalCurr = ls_CurrCls


                '        ls_SQL = " INSERT INTO [dbo].[PO_Detail] " & vbCrLf & _
                '                    "            ([PONo] " & vbCrLf & _
                '                    "            ,[AffiliateID] " & vbCrLf & _
                '                    "            ,[SupplierID] " & vbCrLf & _
                '                    "            ,[PartNo] " & vbCrLf & _
                '                    "            ,[KanbanCls] " & vbCrLf & _
                '                    "            ,[Maker] " & vbCrLf & _
                '                    "            ,[POQty] " & vbCrLf & _
                '                    "            ,[CurrCls] " & vbCrLf & _
                '                    "            ,[Price] " & vbCrLf & _
                '                    "            ,[Amount] "

                '        ls_SQL = ls_SQL + "            ,[DeliveryD1] " & vbCrLf & _
                '                            "            ,[DeliveryD2] " & vbCrLf & _
                '                            "            ,[DeliveryD3] " & vbCrLf & _
                '                            "            ,[DeliveryD4] " & vbCrLf & _
                '                            "            ,[DeliveryD5] " & vbCrLf & _
                '                            "            ,[DeliveryD6] " & vbCrLf & _
                '                            "            ,[DeliveryD7] " & vbCrLf & _
                '                            "            ,[DeliveryD8] "

                '        ls_SQL = ls_SQL + "            ,[DeliveryD9] " & vbCrLf & _
                '                            "            ,[DeliveryD10] " & vbCrLf & _
                '                            "            ,[DeliveryD11] " & vbCrLf & _
                '                            "            ,[DeliveryD12] " & vbCrLf & _
                '                            "            ,[DeliveryD13] " & vbCrLf & _
                '                            "            ,[DeliveryD14] " & vbCrLf & _
                '                            "            ,[DeliveryD15] " & vbCrLf & _
                '                            "            ,[DeliveryD16] " & vbCrLf & _
                '                            "            ,[DeliveryD17] " & vbCrLf & _
                '                            "            ,[DeliveryD18] " & vbCrLf & _
                '                            "            ,[DeliveryD19] "

                '        ls_SQL = ls_SQL + "            ,[DeliveryD20] " & vbCrLf & _
                '                            "            ,[DeliveryD21] " & vbCrLf & _
                '                            "            ,[DeliveryD22] " & vbCrLf & _
                '                            "            ,[DeliveryD23] " & vbCrLf & _
                '                            "            ,[DeliveryD24] " & vbCrLf & _
                '                            "            ,[DeliveryD25] " & vbCrLf & _
                '                            "            ,[DeliveryD26] " & vbCrLf & _
                '                            "            ,[DeliveryD27] " & vbCrLf & _
                '                            "            ,[DeliveryD28] " & vbCrLf & _
                '                            "            ,[DeliveryD29] " & vbCrLf & _
                '                            "            ,[DeliveryD30] "

                '        ls_SQL = ls_SQL + "            ,[DeliveryD31] " & vbCrLf & _
                '                            "            ,[EntryDate] " & vbCrLf & _
                '                            "            ,[EntryUser]) " & vbCrLf & _
                '                            "      VALUES " & vbCrLf & _
                '                            "            ('" & ls_PONo & "' " & vbCrLf & _
                '                            "            ,'" & ls_AffiliateID & "' " & vbCrLf & _
                '                            "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                '                            "            ,'" & ls_PartNo & "' " & vbCrLf & _
                '                            "            ,'" & ls_KanbanCls & "' "

                '        ls_SQL = ls_SQL + "            ,'" & ls_Maker & "' " & vbCrLf & _
                '                            "            ,'" & ls_POQty & "' " & vbCrLf & _
                '                            "            ," & ls_CurrCls & " " & vbCrLf & _
                '                            "            ,'" & ls_Price & "' " & vbCrLf & _
                '                            "            ,'" & ls_Price * ls_POQty & "' " & vbCrLf & _
                '                            "            ,'" & ls_D1 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D2 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D3 & "' "

                '        ls_SQL = ls_SQL + "            ,'" & ls_D4 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D5 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D6 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D7 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D8 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D9 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D10 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D11 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D12 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D13 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D14 & "' "

                '        ls_SQL = ls_SQL + "            ,'" & ls_D15 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D16 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D17 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D18 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D19 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D20 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D21 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D22 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D23 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D24 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D25 & "' "

                '        ls_SQL = ls_SQL + "            ,'" & ls_D26 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D27 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D28 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D29 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D30 & "' " & vbCrLf & _
                '                            "            ,'" & ls_D31 & "' " & vbCrLf & _
                '                            "            , getdate() " & vbCrLf & _
                '                            "            ,'" & Session("UserID") & "' ) "

                '        sqlCommNew = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                '        sqlCommNew.ExecuteNonQuery()
                '        sqlCommNew.Dispose()
                '    Next

                '    publi_PONO = ls_PONo
                '    Session("SupplierID") = ls_SupplierID
                'End If

                sqlTran.Commit()
                Session("B02IsSubmit") = "true"
                Session("PONo") = publi_PONO
                Session("pub_Save") = True
            End Using

            sqlConn.Close()
        End Using

        If Session("Mode") = "New" Then
            ls_MsgID = "1001"
        Else
            ls_MsgID = "1002"
        End If

        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        Call bindPOStatus("new", publi_PONO)
        Session("Mode") = "Update"
        Session("YA010IsSubmit") = lblInfo.Text
        grid.JSProperties("cpMessage") = lblInfo.Text

    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        up_GridLoadWhenEventChange()
        txtPONo.Text = ""
        txtShip.Text = ""

        txtPONo.ReadOnly = False
        txtPONo.BackColor = Color.FromName("#FFFFFF")
        dtPeriodFrom.ReadOnly = False
        dtPeriodFrom.BackColor = Color.FromName("#FFFFFF")
        txtShip.ReadOnly = False
        txtShip.BackColor = Color.FromName("#FFFFFF")
        rdrCom1.ReadOnly = False
        rdrCom2.ReadOnly = False

        dtPeriodFrom.Value = Now
        rdrCom1.Checked = True
        rdrEmergency2.Checked = True

        'btnCraete.Text = "CREATE"

        txtDate1.Text = ""
        txtDate2.Text = ""
        txtDate3.Text = ""
        txtDate4.Text = ""
        txtDate5.Text = ""
        txtDate6.Text = ""
        txtDate7.Text = ""
        txtDate8.Text = ""

        txtUser1.Text = ""
        txtUser2.Text = ""
        txtUser3.Text = ""
        txtUser4.Text = ""
        txtUser5.Text = ""
        txtUser6.Text = ""
        txtUser7.Text = ""
        txtUser8.Text = ""

        lblInfo.Text = ""

        Session("Mode") = "New"
        Session.Remove("SupplierID")

        bindData()
    End Sub

    Private Sub ButtonDelete_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonDelete.Callback
        Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
        'If AlreadyUsed(pAffiliateID) = False Then
        Call deleteData(pAffiliateID)
        'End If
    End Sub

    Private Sub grid_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        With e.Column
            If .FieldName = "MinOrderQty" Then
                Dim ls_MinOrderQty As String = ""
                If IsNothing(e.GetFieldValue("MinOrderQty")) Then ls_MinOrderQty = 0 Else ls_MinOrderQty = IIf(e.GetFieldValue("MinOrderQty").ToString().Trim() = "", 0, e.GetFieldValue("MinOrderQty"))
                e.DisplayText = clsGlobal.FormatQty(ls_MinOrderQty)
            End If

            If .FieldName = "QtyBox" Then
                Dim ls_QtyBox As String = ""
                If IsNothing(e.GetFieldValue("QtyBox")) Then ls_QtyBox = 0 Else ls_QtyBox = IIf(e.GetFieldValue("QtyBox").ToString().Trim() = "", 0, e.GetFieldValue("QtyBox"))
                e.DisplayText = clsGlobal.FormatQty(ls_QtyBox)
            End If

            If .FieldName = "ForecastN1" Then
                Dim ls_ForecastN1 As String = ""
                If IsNothing(e.GetFieldValue("ForecastN1")) Then ls_ForecastN1 = 0 Else ls_ForecastN1 = IIf(e.GetFieldValue("ForecastN1").ToString().Trim() = "", 0, e.GetFieldValue("ForecastN1"))
                e.DisplayText = clsGlobal.FormatQty(ls_ForecastN1)
            End If

            If .FieldName = "ForecastN2" Then
                Dim ls_ForecastN2 As String = ""
                If IsNothing(e.GetFieldValue("ForecastN2")) Then ls_ForecastN2 = 0 Else ls_ForecastN2 = IIf(e.GetFieldValue("ForecastN2").ToString().Trim() = "", 0, e.GetFieldValue("ForecastN2"))
                e.DisplayText = clsGlobal.FormatQty(ls_ForecastN2)
            End If

            If .FieldName = "ForecastN3" Then
                Dim ls_ForecastN3 As String = ""
                If IsNothing(e.GetFieldValue("ForecastN3")) Then ls_ForecastN3 = 0 Else ls_ForecastN3 = IIf(e.GetFieldValue("ForecastN3").ToString().Trim() = "", 0, e.GetFieldValue("ForecastN3"))
                e.DisplayText = clsGlobal.FormatQty(ls_ForecastN3)
            End If

            For i = 1 To 31
                If .FieldName = "DeliveryD" & i Then
                    Dim ls_DeliveryD As String = ""
                    If IsNothing(e.GetFieldValue("DeliveryD" & i)) Then ls_DeliveryD = 0 Else ls_DeliveryD = IIf(e.GetFieldValue("DeliveryD" & i).ToString().Trim() = "", 0, e.GetFieldValue("DeliveryD" & i))
                    e.DisplayText = clsGlobal.FormatQty(ls_DeliveryD)
                End If
            Next
        End With
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        'If rdrKanban2.Checked = True Then
        '    pWhere = pWhere + " and KanbanCls = 'YES'"
        'End If

        'If rdrKanban3.Checked = True Then
        '    pWhere = pWhere + " and KanbanCls = 'NO'"
        'End If

        If Session("Mode") = "Update" Then
            pWhere = pWhere + " and SupplierID = '" & Session("SupplierID") & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT  " & vbCrLf & _
                  " 	AllowAccess,  " & vbCrLf & _
                  " 	row_number() over (order by AllowAccess desc) as NoUrut, " & vbCrLf & _
                  " 	PartNo, PartName, KanbanCls, UnitDesc, MinOrderQty, QtyBox, Maker, PONo, " & vbCrLf & _
                  " 	POQty, CurrDesc, Price, Amount, ForecastN1, ForecastN2, ForecastN3, " & vbCrLf & _
                  " 	ISNULL(DeliveryD1,0) DeliveryD1, ISNULL(DeliveryD2,0) DeliveryD2, ISNULL(DeliveryD3,0) DeliveryD3, ISNULL(DeliveryD4,0) DeliveryD4, ISNULL(DeliveryD5,0) DeliveryD5,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD6,0) DeliveryD6, ISNULL(DeliveryD7,0) DeliveryD7, ISNULL(DeliveryD8,0) DeliveryD8, ISNULL(DeliveryD9,0) DeliveryD9, ISNULL(DeliveryD10,0) DeliveryD10, " & vbCrLf & _
                  " 	ISNULL(DeliveryD11,0) DeliveryD11, ISNULL(DeliveryD12,0) DeliveryD12, ISNULL(DeliveryD13,0) DeliveryD13, ISNULL(DeliveryD14,0) DeliveryD14, ISNULL(DeliveryD15,0) DeliveryD15, " & vbCrLf & _
                  " 	ISNULL(DeliveryD16,0) DeliveryD16, ISNULL(DeliveryD17,0) DeliveryD17, ISNULL(DeliveryD18,0) DeliveryD18, ISNULL(DeliveryD19,0) DeliveryD19, ISNULL(DeliveryD20,0) DeliveryD20, " & vbCrLf & _
                  " 	ISNULL(DeliveryD21,0) DeliveryD21, ISNULL(DeliveryD22,0) DeliveryD22, ISNULL(DeliveryD23,0) DeliveryD23, ISNULL(DeliveryD24,0) DeliveryD24, ISNULL(DeliveryD25,0) DeliveryD25, " & vbCrLf & _
                  " 	ISNULL(DeliveryD26,0) DeliveryD26, ISNULL(DeliveryD27,0) DeliveryD27, ISNULL(DeliveryD28,0) DeliveryD28, ISNULL(DeliveryD29,0) DeliveryD29, ISNULL(DeliveryD30,0) DeliveryD30, " & vbCrLf

            ls_SQL = ls_SQL + " 	ISNULL(DeliveryD31,0) DeliveryD31, countPartNo, SupplierID, CurrCls, UnitCls, PODeliveryBy " & vbCrLf & _
                              " FROM " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " select distinct " & vbCrLf & _
                              " 	'0'AllowAccess, a.PartNo, a.PartName, " & vbCrLf & _
                              " 	case a.KanbanCls when '1' then 'YES' else 'NO' end KanbanCls,   " & vbCrLf & _
                              " 	c.Description UnitDesc,  " & vbCrLf & _
                              " 	a.MOQ MinOrderQty,  " & vbCrLf & _
                              " 	a.QtyBox,  " & vbCrLf & _
                              " 	a.Maker,  " & vbCrLf & _
                              " 	'' PONo,  " & vbCrLf

            ls_SQL = ls_SQL + " 	0 POQty,  " & vbCrLf & _
                              " 	e.Description CurrDesc, isnull(d.Price,0)Price,  " & vbCrLf & _
                              " 	0 Amount,  " & vbCrLf & _
                              " 	ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and b.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	NULL DeliveryD1, NULL DeliveryD2, NULL DeliveryD3, NULL DeliveryD4, NULL DeliveryD5,  " & vbCrLf & _
                              " 	NULL DeliveryD6, NULL DeliveryD7, NULL DeliveryD8, NULL DeliveryD9, NULL DeliveryD10,  " & vbCrLf & _
                              " 	NULL DeliveryD11, NULL DeliveryD12, NULL DeliveryD13, NULL DeliveryD14, NULL DeliveryD15,  " & vbCrLf & _
                              " 	NULL DeliveryD16, NULL DeliveryD17, NULL DeliveryD18, NULL DeliveryD19, NULL DeliveryD20,  " & vbCrLf & _
                              " 	NULL DeliveryD21, NULL DeliveryD22, NULL DeliveryD23, NULL DeliveryD24, NULL DeliveryD25,  " & vbCrLf & _
                              " 	NULL DeliveryD26, NULL DeliveryD27, NULL DeliveryD28, NULL DeliveryD29, NULL DeliveryD30,  " & vbCrLf & _
                              " 	NULL DeliveryD31, countPartNo, case countPartNo when '1' then b.SupplierID else '' end SupplierID, e.CurrCls, c.UnitCls, f.PODeliveryBy " & vbCrLf

            ls_SQL = ls_SQL + " from MS_Parts a  inner join MS_PartMapping b on b.PartNo = a.PartNo  " & vbCrLf & _
                              " inner join MS_PartSetting g on g.AffiliateID = b.AffiliateID and g.PartNo = b.PartNo " & vbCrLf & _
                              " left join MS_UnitCls c on a.UnitCls = c.UnitCls  " & vbCrLf & _
                              " left join MS_Affiliate f on b.AffiliateID = f.AffiliateID " & vbCrLf & _
                              " left join MS_Price d on a.PartNo = d.PartNo and d.AffiliateID = b.AffiliateID and ('" & Format(Now, "yyyy-MM-dd") & "' between StartDate and EndDate) " & vbCrLf & _
                              " left join MS_CurrCls e on d.CurrCls = e.CurrCls  " & vbCrLf & _
                              " left join (select COUNT(PartNo) countPartNo, f.AffiliateID, f.PartNo from MS_PartMapping f group by f.AffiliateID, f.PartNo) z on z.AffiliateID = b.AffiliateID and z.PartNo = a.PartNo" & vbCrLf & _
                              " where b.AffiliateID = '" & Session("AffiliateID") & "' and ShowCls = 1 " & vbCrLf & _
                              " and a.PartNo not in " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select PartNo  " & vbCrLf & _
                              " 		from PO_Master a  " & vbCrLf & _
                              " 		inner join PO_Detail b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo " & vbCrLf & _
                              " 		where a.PONo = '" & txtPONo.Text & "' " & vbCrLf & _
                              " 	) " & vbCrLf

            ls_SQL = ls_SQL + " union all " & vbCrLf & _
                              " select  " & vbCrLf & _
                              " 	'1'AllowAccess, b.PartNo, c.PartName,  " & vbCrLf & _
                              " 	case b.KanbanCls when '1' then 'YES' else 'NO' end KanbanCls,  " & vbCrLf & _
                              " 	d.Description UnitDesc, " & vbCrLf & _
                              " 	c.MOQ, c.QtyBox, c.Maker, b.PONo, " & vbCrLf & _
                              " 	b.POQty, " & vbCrLf & _
                              " 	e.Description CurrDesc, " & vbCrLf & _
                              " 	b.Price, b.Amount, " & vbCrLf & _
                              " 	ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and a.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and a.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and a.AffiliateID = MF.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	b.DeliveryD1 DeliveryD1, b.DeliveryD2 DeliveryD2, b.DeliveryD3 DeliveryD3, b.DeliveryD4 DeliveryD4, b.DeliveryD5 DeliveryD5,  " & vbCrLf

            ls_SQL = ls_SQL + " 	b.DeliveryD6 DeliveryD6, b.DeliveryD7 DeliveryD7, b.DeliveryD8 DeliveryD8, b.DeliveryD9 DeliveryD9, b.DeliveryD10 DeliveryD10, " & vbCrLf & _
                              " 	b.DeliveryD11 DeliveryD11, b.DeliveryD12 DeliveryD12, b.DeliveryD13 DeliveryD13, b.DeliveryD14 DeliveryD14, b.DeliveryD15 DeliveryD15, " & vbCrLf & _
                              " 	b.DeliveryD16 DeliveryD16, b.DeliveryD17 DeliveryD17, b.DeliveryD18 DeliveryD18, b.DeliveryD19 DeliveryD19, b.DeliveryD20 DeliveryD20, " & vbCrLf & _
                              " 	b.DeliveryD21 DeliveryD21, b.DeliveryD22 DeliveryD22, b.DeliveryD23 DeliveryD23, b.DeliveryD24 DeliveryD24, b.DeliveryD25 DeliveryD25, " & vbCrLf & _
                              " 	b.DeliveryD26 DeliveryD26, b.DeliveryD27 DeliveryD27, b.DeliveryD28 DeliveryD28, b.DeliveryD29 DeliveryD29, b.DeliveryD30 DeliveryD30, " & vbCrLf & _
                              " 	b.DeliveryD31 DeliveryD31, '1' countPartNo, a.SupplierID, e.CurrCls, d.UnitCls, f.PODeliveryBy " & vbCrLf & _
                              " from PO_Master a " & vbCrLf & _
                              " inner join PO_Detail b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo " & vbCrLf & _
                              " inner join MS_PartSetting g on g.AffiliateID = b.AffiliateID and g.PartNo = b.PartNo " & vbCrLf & _
                              " left join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " left join MS_Affiliate f on a.AffiliateID = f.AffiliateID " & vbCrLf & _
                              " left join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " left join MS_CurrCls e on e.CurrCls = b.CurrCls " & vbCrLf

            ls_SQL = ls_SQL + " where a.PONo = '" & txtPONo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and ShowCls = 1" & vbCrLf & _
                              " )X " & vbCrLf & _
                              " WHERE 'A' = 'A' " & pWhere & "" & vbCrLf & _
                              " ORDER BY AllowAccess DESC, PartNo "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, False)
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using

        clsGlobal.HideColumTanggal(dtPeriodFrom, grid)

    End Sub

    Private Sub bindPOStatus(Optional ByVal pUpdate As String = "", Optional ByVal pPONO As String = "")
        Dim ls_SQL As String = ""
        Dim ls_PONo As String = ""

        If pPONO <> "" Then
            ls_PONo = pPONO
        Else
            ls_PONo = txtPONo.Text
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	EntryDate, " & vbCrLf & _
                  " 	ISNULL(EntryUser,'')EntryUser, " & vbCrLf & _
                  " 	AffiliateApproveDate, " & vbCrLf & _
                  " 	ISNULL(AffiliateApproveUser,'')AffiliateApproveUser, " & vbCrLf & _
                  " 	PASISendAffiliateDate, " & vbCrLf & _
                  " 	ISNULL(PASISendAffiliateUser,'')PASISendAffiliateUser, " & vbCrLf & _
                  " 	SupplierApproveDate, " & vbCrLf & _
                  " 	ISNULL(SupplierApproveUser,'')SupplierApproveUser, " & vbCrLf & _
                  " 	SupplierApprovePendingDate, " & vbCrLf & _
                  " 	ISNULL(SupplierApprovePendingUser,'')SupplierApprovePendingUser, "

            ls_SQL = ls_SQL + " 	SupplierUnApproveDate, " & vbCrLf & _
                              " 	ISNULL(SupplierUnApproveUser,'')SupplierUnApproveUser, " & vbCrLf & _
                              " 	PASIApproveDate, " & vbCrLf & _
                              " 	ISNULL(PASIApproveUser,'')PASIApproveUser, " & vbCrLf & _
                              " 	FinalApproveDate, " & vbCrLf & _
                              " 	ISNULL(FinalApproveUser,'')FinalApproveUser  " & vbCrLf & _
                              " from PO_Master where PONo = '" & ls_PONo & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "'"


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                If IsDBNull(ds.Tables(0).Rows(0)("EntryDate")) Then
                    txtDate1.Text = ""
                Else
                    txtDate1.Text = Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")) Then
                    txtDate2.Text = ""
                Else
                    txtDate2.Text = Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")) Then
                    txtDate3.Text = ""
                Else
                    txtDate3.Text = Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")) Then
                    txtDate4.Text = ""
                Else
                    txtDate4.Text = Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")) Then
                    txtDate5.Text = ""
                Else
                    txtDate5.Text = Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")) Then
                    txtDate6.Text = ""
                Else
                    txtDate6.Text = Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")) Then
                    txtDate7.Text = ""
                Else
                    txtDate7.Text = Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd HH:mm:ss")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")) Then
                    txtDate8.Text = ""
                Else
                    txtDate8.Text = Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd HH:mm:ss")
                End If

                txtUser1.Text = ds.Tables(0).Rows(0)("EntryUser")
                txtUser2.Text = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                txtUser3.Text = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                txtUser4.Text = ds.Tables(0).Rows(0)("SupplierApproveUser")
                txtUser5.Text = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                txtUser6.Text = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                txtUser7.Text = ds.Tables(0).Rows(0)("PASIApproveUser")
                txtUser8.Text = ds.Tables(0).Rows(0)("FinalApproveUser")

                If pUpdate = "update" Then
                    ButtonApprove.JSProperties("cpDate1") = txtDate1.Text
                    ButtonApprove.JSProperties("cpDate2") = txtDate2.Text
                    ButtonApprove.JSProperties("cpDate3") = txtDate3.Text
                    ButtonApprove.JSProperties("cpDate4") = txtDate4.Text
                    ButtonApprove.JSProperties("cpDate5") = txtDate5.Text
                    ButtonApprove.JSProperties("cpDate6") = txtDate6.Text
                    ButtonApprove.JSProperties("cpDate7") = txtDate7.Text
                    ButtonApprove.JSProperties("cpDate8") = txtDate8.Text

                    ButtonApprove.JSProperties("cpUser1") = txtUser1.Text
                    ButtonApprove.JSProperties("cpUser2") = txtUser2.Text
                    ButtonApprove.JSProperties("cpUser3") = txtUser3.Text
                    ButtonApprove.JSProperties("cpUser4") = txtUser4.Text
                    ButtonApprove.JSProperties("cpUser5") = txtUser5.Text
                    ButtonApprove.JSProperties("cpUser6") = txtUser6.Text
                    ButtonApprove.JSProperties("cpUser7") = txtUser7.Text
                    ButtonApprove.JSProperties("cpUser8") = txtUser8.Text

                    If txtDate2.Text.Trim = "" Then
                        Call clsMsg.DisplayMessage(lblInfo, "1011", clsMessage.MsgType.InformationMessage)
                    Else
                        Call clsMsg.DisplayMessage(lblInfo, "1007", clsMessage.MsgType.InformationMessage)
                    End If

                    ButtonApprove.JSProperties("cpMessage") = lblInfo.Text
                ElseIf pUpdate = "new" Then
                    Session("cpDate1") = txtDate1.Text
                    Session("cpUser1") = txtUser1.Text
                End If
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' AllowAccess, '' NoUrut, " & vbCrLf & _
                  " '' PartNo, '' PartName, '' KanbanCls,'' UnitDesc, '' MinOrderQty, '' QtyBox,'' Maker, " & vbCrLf & _
                  " '' PONo, 0 POQty, '' CurrDesc, '' Price, 0 Amount,   " & vbCrLf & _
                  " 0 ForecastN1, 0 ForecastN2, 0 ForecastN3, " & vbCrLf & _
                  " 0 DeliveryD1, 0 DeliveryD2, 0 DeliveryD3, 0 DeliveryD4, 0 DeliveryD5, " & vbCrLf & _
                  " 0 DeliveryD6, 0 DeliveryD7, 0 DeliveryD8, 0 DeliveryD9, 0 DeliveryD10, " & vbCrLf & _
                  " 0 DeliveryD11, 0 DeliveryD12, 0 DeliveryD13, 0 DeliveryD14, 0 DeliveryD15, " & vbCrLf & _
                  " 0 DeliveryD16, 0 DeliveryD17, 0 DeliveryD18, 0 DeliveryD19, 0 DeliveryD20, " & vbCrLf & _
                  " 0 DeliveryD21, 0 DeliveryD22, 0 DeliveryD23, 0 DeliveryD24, 0 DeliveryD25, " & vbCrLf & _
                  " 0 DeliveryD26, 0 DeliveryD27, 0 DeliveryD28, 0 DeliveryD29, 0 DeliveryD30, " & vbCrLf & _
                  " 0 DeliveryD31, '' countPartNo, '' SupplierID, '' CurrCls, '' UnitCls, ''PODeliveryBy "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()

            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub ColorGrid()
        grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.White
        'grid.VisibleColumns(9).CellStyle.BackColor = Drawing.Color.White

        grid.VisibleColumns(16).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(17).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(18).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(19).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(20).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(21).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(22).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(23).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(24).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(25).CellStyle.BackColor = Drawing.Color.White

        grid.VisibleColumns(26).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(27).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(28).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(29).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(30).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(31).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(32).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(33).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(34).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(35).CellStyle.BackColor = Drawing.Color.White

        grid.VisibleColumns(36).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(37).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(38).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(39).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(40).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(41).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(42).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(43).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(44).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(45).CellStyle.BackColor = Drawing.Color.White

        grid.VisibleColumns(46).CellStyle.BackColor = Drawing.Color.White
        grid.VisibleColumns(47).CellStyle.BackColor = Drawing.Color.White
    End Sub

    Private Sub bindHeijunka()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        'If rdrKanban2.Checked = True Then
        '    pWhere = pWhere + " and KanbanCls = 'YES'"
        'End If

        'If rdrKanban3.Checked = True Then
        '    pWhere = pWhere + " and KanbanCls = 'NO'"
        'End If

        If Session("Mode") = "Update" Then
            pWhere = pWhere + " and SupplierID = '" & Session("SupplierID") & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "declare @JlhHari numeric " & vbCrLf & _
                  "set @JlhHari =  (select COUNT(CalendarDate) from MS_Calendar where HolidayCls = 0  and year(CalendarDate ) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(CalendarDate) = '" & Month(dtPeriodFrom.Value) & "')  " & vbCrLf & _
                  " " & vbCrLf & _
                  " SELECT  " & vbCrLf & _
                  " 	AllowAccess,  " & vbCrLf & _
                  " 	row_number() over (order by AllowAccess desc) as NoUrut, " & vbCrLf & _
                  " 	PartNo, PartName, KanbanCls, UnitDesc, MinOrderQty, QtyBox, Maker, PONo, " & vbCrLf & _
                  " 	POQty, CurrDesc, Price, Amount, ForecastN1, ForecastN2, ForecastN3, " & vbCrLf & _
                  " 	ISNULL(DeliveryD1,0) DeliveryD1, ISNULL(DeliveryD2,0) DeliveryD2, ISNULL(DeliveryD3,0) DeliveryD3, ISNULL(DeliveryD4,0) DeliveryD4, ISNULL(DeliveryD5,0) DeliveryD5,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD6,0) DeliveryD6, ISNULL(DeliveryD7,0) DeliveryD7, ISNULL(DeliveryD8,0) DeliveryD8, ISNULL(DeliveryD9,0) DeliveryD9, ISNULL(DeliveryD10,0) DeliveryD10, " & vbCrLf & _
                  " 	ISNULL(DeliveryD11,0) DeliveryD11, ISNULL(DeliveryD12,0) DeliveryD12, ISNULL(DeliveryD13,0) DeliveryD13, ISNULL(DeliveryD14,0) DeliveryD14, ISNULL(DeliveryD15,0) DeliveryD15, " & vbCrLf & _
                  " 	ISNULL(DeliveryD16,0) DeliveryD16, ISNULL(DeliveryD17,0) DeliveryD17, ISNULL(DeliveryD18,0) DeliveryD18, ISNULL(DeliveryD19,0) DeliveryD19, ISNULL(DeliveryD20,0) DeliveryD20, " & vbCrLf & _
                  " 	ISNULL(DeliveryD21,0) DeliveryD21, ISNULL(DeliveryD22,0) DeliveryD22, ISNULL(DeliveryD23,0) DeliveryD23, ISNULL(DeliveryD24,0) DeliveryD24, ISNULL(DeliveryD25,0) DeliveryD25, " & vbCrLf & _
                  " 	ISNULL(DeliveryD26,0) DeliveryD26, ISNULL(DeliveryD27,0) DeliveryD27, ISNULL(DeliveryD28,0) DeliveryD28, ISNULL(DeliveryD29,0) DeliveryD29, ISNULL(DeliveryD30,0) DeliveryD30, "

            ls_SQL = ls_SQL + " 	ISNULL(DeliveryD31,0) DeliveryD31, countPartNo, SupplierID, CurrCls, UnitCls " & vbCrLf & _
                              " FROM " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " select distinct " & vbCrLf & _
                              " 	'0'AllowAccess, a.PartNo, a.PartName, " & vbCrLf & _
                              " 	case a.KanbanCls when '1' then 'YES' else 'NO' end KanbanCls,   " & vbCrLf & _
                              " 	c.Description UnitDesc,  " & vbCrLf & _
                              " 	a.MOQ MinOrderQty,  " & vbCrLf & _
                              " 	a.QtyBox,  " & vbCrLf & _
                              " 	'' Maker,  " & vbCrLf & _
                              " 	'' PONo,  "

            ls_SQL = ls_SQL + " 	0 POQty,  " & vbCrLf & _
                              " 	e.Description CurrDesc, isnull(d.Price,0)Price,  " & vbCrLf & _
                              " 	0 Amount,  " & vbCrLf & _
                              " 	ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              "     ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	NULL DeliveryD1, NULL DeliveryD2, NULL DeliveryD3, NULL DeliveryD4, NULL DeliveryD5,  " & vbCrLf & _
                              " 	NULL DeliveryD6, NULL DeliveryD7, NULL DeliveryD8, NULL DeliveryD9, NULL DeliveryD10,  " & vbCrLf & _
                              " 	NULL DeliveryD11, NULL DeliveryD12, NULL DeliveryD13, NULL DeliveryD14, NULL DeliveryD15,  " & vbCrLf & _
                              " 	NULL DeliveryD16, NULL DeliveryD17, NULL DeliveryD18, NULL DeliveryD19, NULL DeliveryD20,  " & vbCrLf & _
                              " 	NULL DeliveryD21, NULL DeliveryD22, NULL DeliveryD23, NULL DeliveryD24, NULL DeliveryD25,  " & vbCrLf & _
                              " 	NULL DeliveryD26, NULL DeliveryD27, NULL DeliveryD28, NULL DeliveryD29, NULL DeliveryD30,  " & vbCrLf & _
                              " 	NULL DeliveryD31, countPartNo, case countPartNo when '1' then b.SupplierID else '' end SupplierID, e.CurrCls, c.UnitCls "

            ls_SQL = ls_SQL + " from MS_Parts a  inner join MS_PartMapping b on b.PartNo = a.PartNo  " & vbCrLf & _
                              " left join MS_UnitCls c on a.UnitCls = c.UnitCls  " & vbCrLf & _
                              " left join MS_Price d on a.PartNo = d.PartNo and d.AffiliateID = b.AffiliateID and ('" & Format(Now, "yyyy-MM-dd") & "' between StartDate and EndDate) " & vbCrLf & _
                              " left join MS_CurrCls e on d.CurrCls = e.CurrCls  " & vbCrLf & _
                              " left join (select COUNT(PartNo) countPartNo, f.AffiliateID, f.PartNo from MS_PartMapping f group by f.AffiliateID, f.PartNo) z on z.AffiliateID = b.AffiliateID and z.PartNo = a.PartNo" & vbCrLf & _
                              " where b.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                              " and a.PartNo not in " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select PartNo  " & vbCrLf & _
                              " 		from tempPO_Detail a  " & vbCrLf & _
                              " 		where a.PONo = '" & txtPONo.Text & "' " & vbCrLf & _
                              " 	) " & vbCrLf

            ls_SQL = ls_SQL + " union all " & vbCrLf & _
                              " select distinct " & vbCrLf & _
                              "  	'1'AllowAccess, a.PartNo, c.PartName,   " & vbCrLf & _
                              "  	case c.KanbanCls when '1' then 'YES' else 'NO' end KanbanCls,   " & vbCrLf & _
                              "  	d.Description UnitDesc,  " & vbCrLf & _
                              "  	c.MOQ, c.QtyBox, '' Maker, '' PONo,  " & vbCrLf & _
                              "  	a.POQty,  " & vbCrLf & _
                              "  	e.Description CurrDesc,  " & vbCrLf & _
                              "  	isnull(b.Price,0) Price, a.POQty * isnull(b.Price,0) Amount,  " & vbCrLf & _
                              "  	ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'2015-04-09')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'2015-04-09'))),0),  " & vbCrLf & _
                              "     ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'2015-04-09')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'2015-04-09'))),0),  " & vbCrLf & _
                              "     ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'2015-04-09')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'2015-04-09'))),0),  " & vbCrLf

            ls_SQL = ls_SQL + "  	tgl1 DeliveryD1, tgl2 DeliveryD2, tgl3 DeliveryD3, tgl4 DeliveryD4, tgl5 DeliveryD5, " & vbCrLf & _
                              "  	tgl6 DeliveryD6, tgl7 DeliveryD7, tgl8 DeliveryD8, tgl9 DeliveryD9, tgl10 DeliveryD10,  " & vbCrLf & _
                              "  	tgl11 DeliveryD11, tgl12 DeliveryD12, tgl13 DeliveryD13, tgl14 DeliveryD14, tgl15 DeliveryD15,  " & vbCrLf & _
                              "  	tgl16 DeliveryD16, tgl17 DeliveryD17, tgl18 DeliveryD18, tgl19 DeliveryD19, tgl20 DeliveryD20,  " & vbCrLf & _
                              "  	tgl21 DeliveryD21, tgl22 DeliveryD22, tgl23 DeliveryD23, tgl24 DeliveryD24, tgl25 DeliveryD25,  " & vbCrLf & _
                              "  	tgl26 DeliveryD26, tgl27 DeliveryD27, tgl28 DeliveryD28, tgl29 DeliveryD29, tgl30 DeliveryD30,  " & vbCrLf & _
                              "  	tgl31 DeliveryD31, countPartNo, case countPartNo when '1' then g.SupplierID else '' end SupplierID, e.CurrCls, d.UnitCls  " & vbCrLf & _
                              "  from  " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select PONo, a.PartNo, POQty, QtyBox, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=1 then QtyKirim else 0 end) tgl1,  " & vbCrLf

            ls_SQL = ls_SQL + " 		sum(case when day(calendardate)=2 then QtyKirim else 0 end) tgl2,  " & vbCrLf & _
                              " 		sum(case when day(calendardate)=3 then QtyKirim else 0 end) tgl3,  " & vbCrLf & _
                              " 		sum(case when day(calendardate)=4 then QtyKirim else 0 end) tgl4,  " & vbCrLf & _
                              " 		sum(case when day(calendardate)=5 then QtyKirim else 0 end) tgl5, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=6 then QtyKirim else 0 end) tgl6, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=7 then QtyKirim else 0 end) tgl7, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=8 then QtyKirim else 0 end) tgl8, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=9 then QtyKirim else 0 end) tgl9, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=10 then QtyKirim else 0 end) tgl10, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=11 then QtyKirim else 0 end) tgl11, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=12 then QtyKirim else 0 end) tgl12, " & vbCrLf

            ls_SQL = ls_SQL + " 		sum(case when day(calendardate)=13 then QtyKirim else 0 end) tgl13, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=14 then QtyKirim else 0 end) tgl14, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=15 then QtyKirim else 0 end) tgl15, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=16 then QtyKirim else 0 end) tgl16, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=17 then QtyKirim else 0 end) tgl17, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=18 then QtyKirim else 0 end) tgl18, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=19 then QtyKirim else 0 end) tgl19, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=20 then QtyKirim else 0 end) tgl20, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=21 then QtyKirim else 0 end) tgl21, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=22 then QtyKirim else 0 end) tgl22, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=23 then QtyKirim else 0 end) tgl23, " & vbCrLf

            ls_SQL = ls_SQL + " 		sum(case when day(calendardate)=24 then QtyKirim else 0 end) tgl24, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=25 then QtyKirim else 0 end) tgl25, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=26 then QtyKirim else 0 end) tgl26, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=27 then QtyKirim else 0 end) tgl27, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=28 then QtyKirim else 0 end) tgl28, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=29 then QtyKirim else 0 end) tgl29, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=30 then QtyKirim else 0 end) tgl30, " & vbCrLf & _
                              " 		sum(case when day(calendardate)=31 then QtyKirim else 0 end) tgl31 " & vbCrLf & _
                              " 	from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select " & vbCrLf

            ls_SQL = ls_SQL + " 			AllowAcess, PONo, a.PartNo, POQty, QtyBox, CalendarDate, 		 " & vbCrLf & _
                              " 			(case when jlh%@JlhHari=0 or jlh%@JlhHari>=num then ceiling(jlh/@JlhHari) else (case when ceiling(jlh/@JlhHari)>1 then ceiling(jlh/@JlhHari)-1 else 0 end) end)*QtyBox QtyKirim " & vbCrLf & _
                              " 		from " & vbCrLf & _
                              " 		( " & vbCrLf & _
                              " 			select AllowAcess, PONo, a.PartNo , POQty, QtyBox, (POQty / QtyBox) jlh from tempPO_Detail a " & vbCrLf & _
                              " 			left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " 			where PONo = '" & txtPONo.Text & "' " & vbCrLf & _
                              " 		) a " & vbCrLf & _
                              " 		cross join " & vbCrLf & _
                              " 		( " & vbCrLf & _
                              " 			select CalendarDate,ROW_NUMBER() over (order by calendardate) as num "

            ls_SQL = ls_SQL + " 			from MS_Calendar where year(calendardate) = '" & Year(dtPeriodFrom.Value) & "' and MONTH(CalendarDate) = '" & Month(dtPeriodFrom.Value) & "' and HolidayCls = 0 " & vbCrLf & _
                              " 		) b " & vbCrLf & _
                              " 	) a " & vbCrLf & _
                              " 		group by PONo, a.PartNo, POQty, QtyBox " & vbCrLf & _
                              "  ) a " & vbCrLf & _
                              "  left join MS_Price b on a.PartNo = b.PartNo and b.AffiliateID = '" & Session("AffiliateID") & "' and ('" & Format(Now, "yyyy-MM-dd") & "' between StartDate and EndDate) " & vbCrLf & _
                              "  left join MS_Parts c on a.PartNo = c.PartNo  " & vbCrLf & _
                              "  left join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              "  left join MS_CurrCls e on e.CurrCls = b.CurrCls " & vbCrLf & _
                              "  left join MS_PartMapping g on g.PartNo = a.PartNo and g.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                              "  left join (select COUNT(PartNo) countPartNo, f.AffiliateID, f.PartNo from MS_PartMapping f group by f.AffiliateID, f.PartNo) z on z.AffiliateID = '" & Session("AffiliateID") & "' and z.PartNo = a.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " )X " & vbCrLf & _
                              " WHERE 'A' = 'A' " & pWhere & "" & vbCrLf & _
                              " ORDER BY AllowAccess DESC, PartNo "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, False)
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub saveData()
        Dim i As Integer
        Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        Dim GrandTotal As Double = 0
        Dim GrandCurr As String = ""
        Dim ls_PONo As String = ""
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""

        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "AllowAccess").ToString = "1" Then
                    ls_Check = True
                    Exit For
                End If
            Next i

            If ls_Check = False Then
                lblInfo.Text = "[ Please give a checkmark to save data ! ] "
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()

            SqlTran = SqlCon.BeginTransaction

            Try
                '2. Check MODE UPDATE atau NEW
                If Session("Mode") = "Update" Then
                    '2.1 delete data 
                    Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                    SQLCom.Connection = SqlCon
                    SQLCom.Transaction = SqlTran

                    ls_Sql = "delete PO_Detail where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "'"

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()


                    '2.2 Insert New Detail Data
                    For i = 0 To grid.VisibleRowCount - 1
                        If grid.GetRowValues(i, "AllowAccess").ToString = "1" Then
                            ls_Sql = " INSERT INTO [dbo].[PO_Detail] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[KanbanCls] " & vbCrLf & _
                                      "            ,[Maker] " & vbCrLf & _
                                      "            ,[POQty] " & vbCrLf & _
                                      "            ,[CurrCls] " & vbCrLf & _
                                      "            ,[Price] " & vbCrLf & _
                                      "            ,[Amount] "

                            ls_Sql = ls_Sql + "            ,[ForecastN1] " & vbCrLf & _
                                              "            ,[ForecastN2] " & vbCrLf & _
                                              "            ,[ForecastN3] " & vbCrLf & _
                                              "            ,[DeliveryD1] " & vbCrLf & _
                                              "            ,[DeliveryD2] " & vbCrLf & _
                                              "            ,[DeliveryD3] " & vbCrLf & _
                                              "            ,[DeliveryD4] " & vbCrLf & _
                                              "            ,[DeliveryD5] " & vbCrLf & _
                                              "            ,[DeliveryD6] " & vbCrLf & _
                                              "            ,[DeliveryD7] " & vbCrLf & _
                                              "            ,[DeliveryD8] "

                            ls_Sql = ls_Sql + "            ,[DeliveryD9] " & vbCrLf & _
                                              "            ,[DeliveryD10] " & vbCrLf & _
                                              "            ,[DeliveryD11] " & vbCrLf & _
                                              "            ,[DeliveryD12] " & vbCrLf & _
                                              "            ,[DeliveryD13] " & vbCrLf & _
                                              "            ,[DeliveryD14] " & vbCrLf & _
                                              "            ,[DeliveryD15] " & vbCrLf & _
                                              "            ,[DeliveryD16] " & vbCrLf & _
                                              "            ,[DeliveryD17] " & vbCrLf & _
                                              "            ,[DeliveryD18] " & vbCrLf & _
                                              "            ,[DeliveryD19] "

                            ls_Sql = ls_Sql + "            ,[DeliveryD20] " & vbCrLf & _
                                              "            ,[DeliveryD21] " & vbCrLf & _
                                              "            ,[DeliveryD22] " & vbCrLf & _
                                              "            ,[DeliveryD23] " & vbCrLf & _
                                              "            ,[DeliveryD24] " & vbCrLf & _
                                              "            ,[DeliveryD25] " & vbCrLf & _
                                              "            ,[DeliveryD26] " & vbCrLf & _
                                              "            ,[DeliveryD27] " & vbCrLf & _
                                              "            ,[DeliveryD28] " & vbCrLf & _
                                              "            ,[DeliveryD29] " & vbCrLf & _
                                              "            ,[DeliveryD30] "

                            ls_Sql = ls_Sql + "            ,[DeliveryD31] " & vbCrLf & _
                                              "            ,[EntryDate] " & vbCrLf & _
                                              "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                              "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                              "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "KanbanCls").ToString = "YES", "1", "0") & "' "

                            ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "Maker").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "POQty").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "CurrCls").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Price").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Amount").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "ForecastN1").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "ForecastN1").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "ForecastN1").ToString & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD1").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD1").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD2").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD2").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD3").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD3").ToString) & "' "

                            ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD4").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD4").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD5").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD5").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD6").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD6").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD7").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD7").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD8").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD8").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD9").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD9").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD10").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD10").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD11").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD11").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD12").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD12").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD13").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD13").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD14").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD14").ToString) & "' "

                            ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD15").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD15").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD16").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD16").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD17").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD17").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD18").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD18").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD19").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD19").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD20").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD20").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD21").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD21").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD22").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD22").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD23").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD23").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD24").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD24").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD25").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD25").ToString) & "' "

                            ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD26").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD26").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD27").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD27").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD28").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD28").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD29").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD29").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD30").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD30").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD31").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD31").ToString) & "' " & vbCrLf & _
                                              "            , getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' ) "
                            GrandTotal = GrandTotal + grid.GetRowValues(i, "Amount").ToString
                            GrandCurr = grid.GetRowValues(i, "CurrCls").ToString
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()
                            ls_MsgID = "1002"
                        End If
                    Next i


                    '2.3 Insert data to Master
                    Dim pub_Master As Boolean = False
                    ls_Sql = "select * from PO_Master where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "'"

                    SQLCom.CommandText = ls_Sql
                    Dim da As New SqlDataAdapter(SQLCom)
                    Dim ds As New DataSet

                    da.Fill(ds)

                    If ds.Tables(0).Rows.Count > 0 Then
                        pub_Master = True
                    Else
                        pub_Master = False
                    End If

                    da = Nothing
                    ds = Nothing

                    'New
                    If pub_Master = False Then
                        ls_Sql = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                                  "            ([PONo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[Period] " & vbCrLf & _
                                  "            ,[CommercialCls] " & vbCrLf & _
                                  "            ,[ShipCls] " & vbCrLf & _
                                  "            ,[CurrCls] " & vbCrLf & _
                                  "            ,[Amount] " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                  "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                  "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                  "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                  "            ,'" & txtShip.Text & "' "

                        ls_Sql = ls_Sql + "            ,'" & GrandCurr & "' " & vbCrLf & _
                                          "            ,'" & GrandTotal & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "' "

                    Else
                        'Update
                        ls_Sql = " UPDATE [dbo].[PO_Master] SET [Amount] = '" & GrandTotal & "' WHERE AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "' and PONo='" & txtPONo.Text & "'"

                    End If

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()
                Else
                    '2.1 delete data 
                    Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                    SQLCom.Connection = SqlCon
                    SQLCom.Transaction = SqlTran

                    ls_Sql = "delete PO_Detail where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "'"

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()


                    '2.2 Insert New Detail Data
                    For i = 0 To grid.VisibleRowCount - 1
                        If grid.GetRowValues(i, "AllowAccess").ToString = "1" Then
                            ls_Sql = " INSERT INTO [dbo].[PO_Detail] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[KanbanCls] " & vbCrLf & _
                                      "            ,[Maker] " & vbCrLf & _
                                      "            ,[POQty] " & vbCrLf & _
                                      "            ,[CurrCls] " & vbCrLf & _
                                      "            ,[Price] " & vbCrLf & _
                                      "            ,[Amount] "

                            ls_Sql = ls_Sql + "            ,[ForecastN1] " & vbCrLf & _
                                              "            ,[ForecastN2] " & vbCrLf & _
                                              "            ,[ForecastN3] " & vbCrLf & _
                                              "            ,[DeliveryD1] " & vbCrLf & _
                                              "            ,[DeliveryD2] " & vbCrLf & _
                                              "            ,[DeliveryD3] " & vbCrLf & _
                                              "            ,[DeliveryD4] " & vbCrLf & _
                                              "            ,[DeliveryD5] " & vbCrLf & _
                                              "            ,[DeliveryD6] " & vbCrLf & _
                                              "            ,[DeliveryD7] " & vbCrLf & _
                                              "            ,[DeliveryD8] "

                            ls_Sql = ls_Sql + "            ,[DeliveryD9] " & vbCrLf & _
                                              "            ,[DeliveryD10] " & vbCrLf & _
                                              "            ,[DeliveryD11] " & vbCrLf & _
                                              "            ,[DeliveryD12] " & vbCrLf & _
                                              "            ,[DeliveryD13] " & vbCrLf & _
                                              "            ,[DeliveryD14] " & vbCrLf & _
                                              "            ,[DeliveryD15] " & vbCrLf & _
                                              "            ,[DeliveryD16] " & vbCrLf & _
                                              "            ,[DeliveryD17] " & vbCrLf & _
                                              "            ,[DeliveryD18] " & vbCrLf & _
                                              "            ,[DeliveryD19] "

                            ls_Sql = ls_Sql + "            ,[DeliveryD20] " & vbCrLf & _
                                              "            ,[DeliveryD21] " & vbCrLf & _
                                              "            ,[DeliveryD22] " & vbCrLf & _
                                              "            ,[DeliveryD23] " & vbCrLf & _
                                              "            ,[DeliveryD24] " & vbCrLf & _
                                              "            ,[DeliveryD25] " & vbCrLf & _
                                              "            ,[DeliveryD26] " & vbCrLf & _
                                              "            ,[DeliveryD27] " & vbCrLf & _
                                              "            ,[DeliveryD28] " & vbCrLf & _
                                              "            ,[DeliveryD29] " & vbCrLf & _
                                              "            ,[DeliveryD30] "

                            ls_Sql = ls_Sql + "            ,[DeliveryD31] " & vbCrLf & _
                                              "            ,[EntryDate] " & vbCrLf & _
                                              "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                              "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                              "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "KanbanCls").ToString = "YES", "1", "0") & "' "

                            ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "Maker").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "POQty").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "CurrCls").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Price").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Amount").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "ForecastN1").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "ForecastN1").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "ForecastN1").ToString & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD1").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD1").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD2").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD2").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD3").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD3").ToString) & "' "

                            ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD4").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD4").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD5").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD5").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD6").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD6").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD7").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD7").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD8").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD8").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD9").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD9").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD10").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD10").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD11").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD11").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD12").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD12").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD13").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD13").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD14").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD14").ToString) & "' "

                            ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD15").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD15").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD16").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD16").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD17").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD17").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD18").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD18").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD19").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD19").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD20").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD20").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD21").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD21").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD22").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD22").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD23").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD23").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD24").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD24").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD25").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD25").ToString) & "' "

                            ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD26").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD26").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD27").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD27").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD28").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD28").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD29").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD29").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD30").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD30").ToString) & "' " & vbCrLf & _
                                              "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD31").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD31").ToString) & "' " & vbCrLf & _
                                              "            , getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' ) "
                            GrandTotal = GrandTotal + grid.GetRowValues(i, "Amount").ToString
                            GrandCurr = grid.GetRowValues(i, "CurrCls").ToString
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()
                            ls_MsgID = "1002"
                        End If
                    Next i


                    '2.3 Insert data to Master
                    Dim pub_Master As Boolean = False
                    ls_Sql = "select * from PO_Master where PONo = '" & txtPONo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "'"

                    SQLCom.CommandText = ls_Sql
                    Dim da As New SqlDataAdapter(SQLCom)
                    Dim ds As New DataSet

                    da.Fill(ds)

                    If ds.Tables(0).Rows.Count > 0 Then
                        pub_Master = True
                    Else
                        pub_Master = False
                    End If

                    da = Nothing
                    ds = Nothing

                    'New
                    If pub_Master = False Then
                        ls_Sql = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                                  "            ([PONo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[Period] " & vbCrLf & _
                                  "            ,[CommercialCls] " & vbCrLf & _
                                  "            ,[ShipCls] " & vbCrLf & _
                                  "            ,[CurrCls] " & vbCrLf & _
                                  "            ,[Amount] " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & txtPONo.Text & "' " & vbCrLf & _
                                  "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                  "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                  "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                  "            ,'" & txtShip.Text & "' "

                        ls_Sql = ls_Sql + "            ,'" & GrandCurr & "' " & vbCrLf & _
                                          "            ,'" & GrandTotal & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "' "

                    Else
                        'Update
                        ls_Sql = " UPDATE [dbo].[PO_Master] SET [Amount] = '" & GrandTotal & "' WHERE AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & Session("SupplierID") & "' and PONo='" & txtPONo.Text & "'"

                    End If

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()
                End If

                SqlTran.Commit()
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
            Catch ex As Exception
                Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                SqlTran.Rollback()
                SqlCon.Close()
                Exit Sub
            End Try

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Sub

    Private Sub deleteData(ByVal pPONo As String)
        Dim sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("DeleteRevisionMasterData")

                sql = " delete PO_Detail " & vbCrLf & _
                    " where PONo='" & txtPONo.Text.Trim & "' and AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
                    " and SupplierID ='" & Session("SupplierID") & "' "

                Dim SqlComm As New SqlCommand(sql, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()

                sql = " delete PO_Master " & vbCrLf & _
                    " where PONo='" & txtPONo.Text.Trim & "' and AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
                    " and SupplierID ='" & Session("SupplierID") & "' "

                Dim SqlComm1 As New SqlCommand(sql, sqlConn, sqlTran)
                SqlComm1.ExecuteNonQuery()
                SqlComm1.Dispose()

                sqlTran.Commit()

                Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
                ButtonDelete.JSProperties("cpMessage") = lblInfo.Text
                Session("delete") = "delete"
            End Using

            sqlConn.Close()

        End Using
    End Sub

    Private Sub uf_Approve()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PO_Master set AffiliateApproveDate = getdate(), AffiliateApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & txtPONo.Text & "' and SupplierID = '" & Session("SupplierID") & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub uf_UnApprove()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PO_Master set AffiliateApproveDate = NULL, AffiliateApproveUser = NULL" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & txtPONo.Text & "' and SupplierID = '" & Session("SupplierID") & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    'Private Sub sendEmail()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Body As String

    '    Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderDetail.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(dtPeriodFrom.Value) & _
    '                           "&t2=" & clsNotification.EncryptURL(Session("AffiliateID")) & "&t3=" & clsNotification.EncryptURL(Session("SupplierID")) & "&t4=" & clsNotification.EncryptURL(txtPONo.Text) & "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderList.aspx")
    '    ls_Body = clsNotification.GetNotification("10", ls_URl, txtPONo.Text.Trim, , , , txtPONo.Text.Trim & "-" & Session("SupplierID"))

    '    Dim dsEmail As New DataSet
    '    dsEmail = EmailToEmailCC(Session("AffiliateID"), "PASI", "")
    '    '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '    For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '        If receiptCCEmail = "" Then
    '            receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '        Else
    '            receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '        End If
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
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

    '    'If receiptCCEmail <> "" Then
    '    '    receiptCCEmail = Left(receiptCCEmail, receiptCCEmail.Length - 1)
    '    'End If


    '    If receiptEmail = "" Then
    '        MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '        Exit Sub
    '    End If

    '    'Make a copy of the file/Open it/Mail it/Delete it
    '    'If you want to change the file name then change only TempFileName


    '    'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
    '    Dim mailMessage As New Mail.MailMessage()
    '    mailMessage.From = New MailAddress(fromEmail)
    '    mailMessage.Subject = "Issued PONo: " & txtPONo.Text

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
    '        smtp.EnableSsl = False
    '    Else
    '        smtp.EnableSsl = True
    '        Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(trim(usernameSMTP), trim(PasswordSMTP))
    '        smtp.Credentials = myCredential
    '    End If

    '    smtp.Port = portClient
    '    smtp.Send(mailMessage)

    'End Sub

    Private Sub sendEmailtoAffiliate()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""
            Dim ls_Body As String = ""

            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNameAffiliate & "/PurchaseOrder/POEntry.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(txtShip.Text.Trim) & _
                                    "&t2=" & clsNotification.EncryptURL(IIf(rdrCom1.Checked = True, "YES", "NO")) & "&t3=" & clsNotification.EncryptURL(dtPeriodFrom.Value) & "&t4=" & clsNotification.EncryptURL(Session("SupplierID")) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrder/POList.aspx")

            ls_Body = clsNotification.GetNotification("10", ls_URl, txtPONo.Text.Trim, , , , txtPONo.Text.Trim & "-" & Session("SupplierID"))

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCC(Session("AffiliateID"), "PASI", "")
            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("FromEmail")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next
            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            End If

            If fromEmail = "" Then
                MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Issued PONo: " & Trim(txtPONo.Text)

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If

            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail()

            mailMessage.Body = ls_Body
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            If smtp.UseDefaultCredentials = True Then
                smtp.EnableSsl = True
            Else
                smtp.EnableSsl = False
                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
                smtp.Credentials = myCredential
            End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try


    End Sub

    Private Sub sendEmailccPASI()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""
            Dim ls_Body As String = ""

            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderDetail.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(dtPeriodFrom.Value) & _
                              "&t2=" & clsNotification.EncryptURL(Session("AffiliateID")) & "&t3=" & clsNotification.EncryptURL(Session("SupplierID")) & "&t4=" & clsNotification.EncryptURL(txtPONo.Text) & "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderList.aspx")
            ls_Body = clsNotification.GetNotification("10", ls_URl, txtPONo.Text.Trim, , , , txtPONo.Text.Trim & "-" & Session("SupplierID"))

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCC(Session("AffiliateID"), "PASI", "")
            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    'fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                'receiptEmail = receiptCCEmail
            Next
            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            End If

            If fromEmail = "" Then
                MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "Issued PONo: " & Trim(txtPONo.Text)

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If

            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If

            GetSettingEmail()

            mailMessage.Body = ls_Body
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            If smtp.UseDefaultCredentials = True Then
                smtp.EnableSsl = True
            Else
                smtp.EnableSsl = False
                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
                smtp.Credentials = myCredential
            End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub GetSettingEmail()
        Dim ls_SQL As String = ""
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.Ms_EmailSetting"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                smtpClient = Trim(ds.Tables(0).Rows(0)("SMTP"))
                portClient = Trim(ds.Tables(0).Rows(0)("PORTSMTP"))
                usernameSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("usernameSMTP")), "", ds.Tables(0).Rows(0)("usernameSMTP"))
                PasswordSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("passwordSMTP")), "", ds.Tables(0).Rows(0)("passwordSMTP"))
            End If
        End Using
    End Sub

    Private Function EmailToEmailCC(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                     " select 'AFF' flag,affiliatepocc, affiliatepoto='',FromEmail = affiliatepoto from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,affiliatepocc,affiliatepoto,FromEmail = ''  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function
#End Region

End Class