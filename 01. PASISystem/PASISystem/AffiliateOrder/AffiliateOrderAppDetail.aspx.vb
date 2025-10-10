Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports System.Net
Imports System.Net.Mail

Public Class AffiliateOrderAppDetail
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "B04"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_AffiliateID As String, pub_AffiliateName As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String, pub_SupplierName As String, pub_Remarks As String
    Dim pub_FinalApproval As String, pub_DeliveyBy As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim flag As Boolean = True
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
            Session("M01Url") = Request.QueryString("Session")
            flag = False
        Else
            flag = True
        End If

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            If Session("M01Url") <> "" Then
                If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                    Session("MenuDesc") = "AFFILIATE ORDER APPROVAL DETAIL"
                    up_Fillcombo()
                    pub_PONo = Request.QueryString("id")
                    pub_AffiliateID = Request.QueryString("t1")
                    pub_AffiliateName = Request.QueryString("t2")
                    pub_Period = Request.QueryString("t3")
                    pub_SupplierID = Request.QueryString("t4")
                    pub_Remarks = Request.QueryString("t5")
                    pub_FinalApproval = Request.QueryString("t6")
                    pub_DeliveyBy = Request.QueryString("t7")
                    pub_Ship = Request.QueryString("t8")
                    pub_Commercial = Request.QueryString("t9")
                    pub_SupplierName = Request.QueryString("t10")

                    dtPeriod.Value = pub_Period
                    cboAffiliateCode.Text = pub_AffiliateID
                    txtAffiliateName.Text = pub_AffiliateName
                    cboPONo.Text = pub_PONo
                    txtDeliveryBy.Text = pub_Ship
                    txtShipBy.Text = pub_DeliveyBy
                    txtRemarks.Text = pub_Remarks
                    txtDeliveryBy.Text = IIf(pub_DeliveyBy = "1", "VIA PASI", "DIRECT AFFILIATE")
                    txtCommercial.Text = pub_Commercial
                    txtSupplierCode.Text = pub_SupplierID
                    txtSupplierName.Text = pub_SupplierName

                    Session("Mode") = "Update"

                    'rblPOKanban.Value = 2
                    'rdrDiff1.Checked = True

                    bindData(pub_Period, pub_PONo, pub_AffiliateID, pub_SupplierID, pub_Commercial, pub_DeliveyBy, pub_Ship)
                    bindPOStatus("", pub_PONo, pub_AffiliateID)

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    If (txtPASIAppDate.Text <> "" And txtPASIAppDate.Text <> "-") Or (txtAffFinalAppDate.Text <> "" And txtAffFinalAppDate.Text <> "-") Then
                        btnApprove.Enabled = False
                    End If
                    'cboPONo.ReadOnly = True
                    'cboPONo.BackColor = Color.FromName("#CCCCCC")
                    'dtPeriod.ReadOnly = True
                    'dtPeriod.BackColor = Color.FromName("#CCCCCC")

                    'If pub_FinalApproval <> "1" Then
                    '    btnApprove.Enabled = False
                    'Else
                    '    btnApprove.Enabled = True
                    'End If
                ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                    Session("MenuDesc") = "AFFILIATE ORDER APPROVAL DETAIL"
                    up_Fillcombo()
                    pub_PONo = clsNotification.DecryptURL(Request.QueryString("id2"))
                    pub_AffiliateID = clsNotification.DecryptURL(Request.QueryString("t1"))
                    pub_AffiliateName = clsNotification.DecryptURL(Request.QueryString("t2"))
                    pub_Period = clsNotification.DecryptURL(Request.QueryString("t3"))
                    pub_SupplierID = clsNotification.DecryptURL(Request.QueryString("t4"))
                    pub_Remarks = clsNotification.DecryptURL(Request.QueryString("t5"))
                    pub_FinalApproval = clsNotification.DecryptURL(Request.QueryString("t6"))
                    pub_DeliveyBy = clsNotification.DecryptURL(Request.QueryString("t7"))
                    pub_Ship = clsNotification.DecryptURL(Request.QueryString("t8"))
                    pub_Commercial = clsNotification.DecryptURL(Request.QueryString("t9"))
                    pub_SupplierName = clsNotification.DecryptURL(Request.QueryString("t10"))

                    dtPeriod.Value = pub_Period
                    cboAffiliateCode.Text = pub_AffiliateID
                    txtAffiliateName.Text = pub_AffiliateName
                    cboPONo.Text = pub_PONo
                    txtDeliveryBy.Text = pub_Ship
                    txtShipBy.Text = pub_DeliveyBy
                    txtRemarks.Text = pub_Remarks
                    txtDeliveryBy.Text = pub_DeliveyBy
                    txtCommercial.Text = pub_Commercial
                    txtSupplierCode.Text = pub_SupplierID
                    txtSupplierName.Text = pub_SupplierName

                    Session("Mode") = "Update"

                    'rblPOKanban.Value = 2
                    'rdrDiff1.Checked = True

                    bindData(pub_Period, pub_PONo, pub_AffiliateID, pub_SupplierID, pub_Commercial, pub_DeliveyBy, pub_Ship)
                    bindPOStatus("", pub_PONo, pub_AffiliateID)

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    If (txtPASIAppDate.Text <> "" And txtPASIAppDate.Text <> "-") Or (txtAffFinalAppDate.Text <> "" And txtAffFinalAppDate.Text <> "-") Then
                        btnApprove.Enabled = False
                    End If
                    'cboPONo.ReadOnly = True
                    'cboPONo.BackColor = Color.FromName("#CCCCCC")
                    'dtPeriod.ReadOnly = True
                    'dtPeriod.BackColor = Color.FromName("#CCCCCC")

                    'If pub_FinalApproval <> "1" Then
                    '    btnApprove.Enabled = False
                    'Else
                    '    btnApprove.Enabled = True
                    'End If
                Else
                    Session("MenuDesc") = "AFFILIATE ORDER APPROVAL DETAIL"
                    Session("Mode") = "New"
                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    cboPONo.Focus()
                    dtPeriod.Value = Now
                    up_Fillcombo()
                    rblPOKanban.Value = 2
                End If
            Else
                Session("Mode") = "New"
                cboPONo.Focus()
                dtPeriod.Value = Now
                up_Fillcombo()
                rblPOKanban.Value = 2
            End If

            lblInfo.Text = ""

        ElseIf IsCallback Then
            If grid.VisibleRowCount = 0 Then Exit Sub
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/AffiliateOrder/AffiliateOrderAppList.aspx")
        Else
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        grid.JSProperties("cpMessage") = ""
        bindData(pub_Period, pub_PONo, pub_PONo, pub_SupplierID, pub_Commercial, pub_DeliveyBy, pub_Ship)
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Dim pDate As Date = Split(e.Parameters, "|")(1)
                    Dim pAffCode As String = Split(e.Parameters, "|")(2)
                    Dim pPONo As String = Split(e.Parameters, "|")(3)
                    Dim pSuppCode As String = Split(e.Parameters, "|")(4)
                    Dim pComm As String = Split(e.Parameters, "|")(5)
                    Dim pDelBy As String = Split(e.Parameters, "|")(6)
                    Dim pKanban As String = Split(e.Parameters, "|")(7)
                    Dim pShipBy As String = Split(e.Parameters, "|")(8)
                    Call bindData(pDate, pPONo, pAffCode, pSuppCode, pComm, pDelBy, pShipBy)
                    Call bindPOStatus("", pPONo, pAffCode)

                    grid.JSProperties("cpSearch") = "search"

                    'Dim TempASPxGridViewCellMerger As ASPxGridViewCellMerger = New ASPxGridViewCellMerger(grid, "NoUrut,PartNo,PartName,KanbanCls,UnitDesc,MOQ,QtyBox,Maker")
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    grid.JSProperties("cpSearch") = ""
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
        Try
            Dim ls_MsgID As String = ""
            If getApp(Trim(cboPONo.Text)) = True Then
                ls_MsgID = "6030"
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                Session("YA010IsSubmit") = lblInfo.Text
                Exit Sub
            End If
            uf_Approve()
            bindPOStatus("update", pub_PONo, pub_AffiliateID)
            'sendEmail()
            sendEmailtoAffiliate()
            sendEmailccPASI()
            If (txtPASIAppDate.Text <> "" And txtPASIAppDate.Text <> "-") Or (txtAffFinalAppDate.Text <> "" And txtAffFinalAppDate.Text <> "-") Then
                btnApprove.Enabled = False
            End If
            ls_MsgID = "6030"
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            Session("YA010IsSubmit") = lblInfo.Text
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
        
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        grid.CollapseAll()

        cboPONo.Text = ""
        txtShipBy.Text = ""
        txtRemarks.Text = ""
        txtDeliveryBy.Text = ""
        txtCommercial.Text = ""

        cboPONo.ReadOnly = False
        cboPONo.BackColor = Color.FromName("#FFFFFF")
        dtPeriod.ReadOnly = False
        dtPeriod.BackColor = Color.FromName("#FFFFFF")

        dtPeriod.Value = Now

        rblPOKanban.Value = 2

        up_Fillcombo()

        'txtEntryDate.Text = ""
        'txtAffAppDate.Text = ""
        'txtSendDate.Text = ""
        'txtSuppAppDate.Text = ""
        'txtSuppPendDate.Text = ""
        'txtSuppUnpDate.Text = ""
        'txtPASIAppDate.Text = ""
        'txtAffFinalAppDate.Text = ""

        'txtEntryUser.Text = ""
        'txtAffAppUser.Text = ""
        'txtSendUser.Text = ""
        'txtSuppAppUser.Text = ""
        'txtSuppPendUser.Text = ""
        'txtSuppUnpUser.Text = ""
        'txtPASIAppUser.Text = ""
        'txtAffFinalAppUser.Text = ""

        lblInfo.Text = ""

        Session("Mode") = "New"
        Session.Remove("SupplierID")

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                If e.GetValue("AffiliateName") = "BY AFFILIATE" Then
                    e.Cell.BackColor = Color.AliceBlue
                End If

                If e.GetValue("AffiliateName") = "BY PASI" Then
                    If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" Then
                        e.Cell.Text = ""
                    End If
                    If e.DataColumn.FieldName = "ForecastN1" Or e.DataColumn.FieldName = "ForecastN2" Or e.DataColumn.FieldName = "ForecastN3" Then
                        e.Cell.Text = ""
                    End If
                    If CDbl(e.GetValue("POQty")) <> CDbl(e.GetValue("POQtyOld")) Then
                        If e.DataColumn.FieldName = "POQty" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD1")) <> CDbl(e.GetValue("DeliveryD1Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD1" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD2")) <> CDbl(e.GetValue("DeliveryD2Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD2" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD3")) <> CDbl(e.GetValue("DeliveryD3Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD3" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD4")) <> CDbl(e.GetValue("DeliveryD4Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD4" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD5")) <> CDbl(e.GetValue("DeliveryD5Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD5" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD6")) <> CDbl(e.GetValue("DeliveryD6Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD6" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD7")) <> CDbl(e.GetValue("DeliveryD7Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD7" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD8")) <> CDbl(e.GetValue("DeliveryD8Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD8" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD9")) <> CDbl(e.GetValue("DeliveryD9Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD9" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD10")) <> CDbl(e.GetValue("DeliveryD10Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD10" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD11")) <> CDbl(e.GetValue("DeliveryD11Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD11" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD12")) <> CDbl(e.GetValue("DeliveryD12Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD12" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD13")) <> CDbl(e.GetValue("DeliveryD13Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD13" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD14")) <> CDbl(e.GetValue("DeliveryD14Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD14" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD15")) <> CDbl(e.GetValue("DeliveryD15Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD15" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD16")) <> CDbl(e.GetValue("DeliveryD16Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD16" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD17")) <> CDbl(e.GetValue("DeliveryD17Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD17" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD18")) <> CDbl(e.GetValue("DeliveryD18Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD18" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD19")) <> CDbl(e.GetValue("DeliveryD19Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD19" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD20")) <> CDbl(e.GetValue("DeliveryD20Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD20" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD21")) <> CDbl(e.GetValue("DeliveryD21Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD21" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD22")) <> CDbl(e.GetValue("DeliveryD22Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD22" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD23")) <> CDbl(e.GetValue("DeliveryD23Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD23" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD24")) <> CDbl(e.GetValue("DeliveryD24Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD24" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD25")) <> CDbl(e.GetValue("DeliveryD25Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD25" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD26")) <> CDbl(e.GetValue("DeliveryD26Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD26" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD27")) <> CDbl(e.GetValue("DeliveryD27Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD27" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD28")) <> CDbl(e.GetValue("DeliveryD28Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD28" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD29")) <> CDbl(e.GetValue("DeliveryD29Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD29" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD30")) <> CDbl(e.GetValue("DeliveryD30Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD30" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD31")) <> CDbl(e.GetValue("DeliveryD31Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD31" Then
                            e.Cell.BackColor = Color.Yellow
                        End If
                    End If
                End If

                If e.GetValue("AffiliateName") = "BY SUPPLIER" Then
                    If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" Then
                        e.Cell.Text = ""
                    End If
                    If e.DataColumn.FieldName = "ForecastN1" Or e.DataColumn.FieldName = "ForecastN2" Or e.DataColumn.FieldName = "ForecastN3" Then
                        e.Cell.Text = ""
                    End If
                    If CDbl(e.GetValue("POQty")) <> CDbl(e.GetValue("POQtyOld")) Then
                        If e.DataColumn.FieldName = "POQty" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD1")) <> CDbl(e.GetValue("DeliveryD1Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD1" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD2")) <> CDbl(e.GetValue("DeliveryD2Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD2" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD3")) <> CDbl(e.GetValue("DeliveryD3Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD3" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD4")) <> CDbl(e.GetValue("DeliveryD4Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD4" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD5")) <> CDbl(e.GetValue("DeliveryD5Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD5" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD6")) <> CDbl(e.GetValue("DeliveryD6Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD6" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD7")) <> CDbl(e.GetValue("DeliveryD7Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD7" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD8")) <> CDbl(e.GetValue("DeliveryD8Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD8" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD9")) <> CDbl(e.GetValue("DeliveryD9Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD9" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD10")) <> CDbl(e.GetValue("DeliveryD10Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD10" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD11")) <> CDbl(e.GetValue("DeliveryD11Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD11" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD12")) <> CDbl(e.GetValue("DeliveryD12Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD12" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD13")) <> CDbl(e.GetValue("DeliveryD13Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD13" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD14")) <> CDbl(e.GetValue("DeliveryD14Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD14" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD15")) <> CDbl(e.GetValue("DeliveryD15Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD15" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD16")) <> CDbl(e.GetValue("DeliveryD16Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD16" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD17")) <> CDbl(e.GetValue("DeliveryD17Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD17" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD18")) <> CDbl(e.GetValue("DeliveryD18Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD18" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD19")) <> CDbl(e.GetValue("DeliveryD19Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD19" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD20")) <> CDbl(e.GetValue("DeliveryD20Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD20" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD21")) <> CDbl(e.GetValue("DeliveryD21Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD21" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD22")) <> CDbl(e.GetValue("DeliveryD22Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD22" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD23")) <> CDbl(e.GetValue("DeliveryD23Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD23" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD24")) <> CDbl(e.GetValue("DeliveryD24Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD24" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD25")) <> CDbl(e.GetValue("DeliveryD25Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD25" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD26")) <> CDbl(e.GetValue("DeliveryD26Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD26" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD27")) <> CDbl(e.GetValue("DeliveryD27Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD27" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD28")) <> CDbl(e.GetValue("DeliveryD28Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD28" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD29")) <> CDbl(e.GetValue("DeliveryD29Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD29" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD30")) <> CDbl(e.GetValue("DeliveryD30Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD30" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                    If CDbl(e.GetValue("DeliveryD31")) <> CDbl(e.GetValue("DeliveryD31Old")) Then
                        If e.DataColumn.FieldName = "DeliveryD31" Then
                            e.Cell.BackColor = Color.GreenYellow
                        End If
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub ButtonPartNo_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonPartNo.Callback
        Call up_GridLoadWhenEventChange()
        Call bindHeader(Split(e.Parameter, "|")(0))
    End Sub

    Private Sub cboPONo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPONo.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_value As String = Split(e.Parameter, "|")(0)
        Dim ls_sql As String = ""

        ls_sql = "SELECT '" & clsGlobal.gs_All & "' PONo UNION ALL SELECT RTRIM(PONo)PONo FROM dbo.PO_Master WHERE YEAR(Period) = YEAR('" & dtPeriod.Value & "') AND MONTH(Period) = MONTH('" & dtPeriod.Value & "')  AND AffiliateID='" & ls_value & "' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPONo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PONo")
                .Columns(0).Width = 50

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub cbPONo_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbPONo.Callback

        Dim ls_sql As String = ""

        Dim pAction As String = Split(e.Parameter, "|")(0)
        Dim pDate As Date = Split(e.Parameter, "|")(1)
        Dim pPONo As String = Split(e.Parameter, "|")(2)
        Dim pAffCode As String = Split(e.Parameter, "|")(3)


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_sql = "   SELECT DISTINCT Period,POD.AffiliateID,AffiliateName,POM.PONo  " & vbCrLf & _
                  "   ,CASE WHEN CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls  " & vbCrLf & _
                  "   ,POD.SupplierID,SupplierName,ShipCls   " & vbCrLf & _
                  "   ,PODeliveryBy   " & vbCrLf & _
                  "   ,POD.KanbanCls   " & vbCrLf & _
                  "   FROM dbo.PO_Master POM    " & vbCrLf & _
                  "   LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID  AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                  "   LEFT JOIN dbo.MS_Affiliate MA ON POD.AffiliateID = MA.AffiliateID  " & vbCrLf & _
                  "   LEFT JOIN dbo.MS_Parts MP ON POD.PartNo = MP.PartNo   " & vbCrLf & _
                  "   LEFT JOIN dbo.MS_Supplier MS ON POD.SupplierID = MS.SupplierID   " & vbCrLf & _
                  "  WHERE YEAR(Period) = YEAR('" & pDate & "') AND MONTH(Period) = MONTH('" & pDate & "')  "

            ls_sql = ls_sql + "  AND POM.PONo = '" & pPONo & "'    " & vbCrLf & _
                              "  AND POM.AffiliateID='" & pAffCode & "' "


            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With ds.Tables(0)
                If ds.Tables(0).Rows.Count > 0 Then
                    cbPONo.JSProperties("cpCommercialCls") = .Rows(0).Item("CommercialCls")
                    cbPONo.JSProperties("cpSupplierID") = .Rows(0).Item("SupplierID")
                    cbPONo.JSProperties("cpSupplierName") = .Rows(0).Item("SupplierName")
                    cbPONo.JSProperties("cpShipCls") = .Rows(0).Item("ShipCls")
                    cbPONo.JSProperties("cpPODeliveryBy") = IIf(.Rows(0).Item("PODeliveryBy")="0","DIRECT TO AFFILIATE","VIA PASI")
                    cbPONo.JSProperties("cpKanbanCls") = .Rows(0).Item("KanbanCls")
                Else
                    cbPONo.JSProperties("cpCommercialCls") = ""
                    cbPONo.JSProperties("cpSupplierID") = ""
                    cbPONo.JSProperties("cpSupplierName") = ""
                    cbPONo.JSProperties("cpShipCls") = ""
                    cbPONo.JSProperties("cpPODeliveryBy") = ""
                    cbPONo.JSProperties("cpKanbanCls") = 2
                End If
            End With

            sqlConn.Close()
        End Using
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub up_Fillcombo()
        Dim ls_SQL As String = ""
        'Combo Affiliate
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 50
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 120

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliateName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
        'Combo PONo
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PONo UNION ALL SELECT RTRIM(PONo)PONo FROM dbo.PO_Master" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPONo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PONo")
                .Columns(0).Width = 50

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindData(ByVal pDate As Date, ByVal pPONo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pComm As String, ByVal pDeliv As String, ByVal pShip As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If pPONo <> clsGlobal.gs_All Then
            pWhere = pWhere + " a.PONo = '" & pPONo & "' " & vbCrLf
        End If

        If pAff <> clsGlobal.gs_All Then
            pWhere = pWhere + " AND a.AffiliateID='" & pAff & "' " & vbCrLf
        End If

        pWhere = pWhere + " AND a.SupplierID = '" & pSupp & "'"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select tbl2.NoUrut, tbl1.AffiliateName, tbl2.PartNo, tbl1.PartNo1, tbl2.PartName, tbl2.KanbanCls, tbl2.UnitDesc, tbl2.MOQ, tbl2.QtyBox, tbl2.Maker, " & vbCrLf & _
                      " 	ISNULL(POQty,0)POQty, ISNULL(POQtyOld,0)POQtyOld, ForecastN1, ForecastN2, ForecastN3, " & vbCrLf & _
                      " 	ISNULL(DeliveryD1,0)DeliveryD1, ISNULL(DeliveryD2,0)DeliveryD2, ISNULL(DeliveryD3,0)DeliveryD3, ISNULL(DeliveryD4,0)DeliveryD4, ISNULL(DeliveryD5,0)DeliveryD5,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD6,0)DeliveryD6, ISNULL(DeliveryD7,0)DeliveryD7, ISNULL(DeliveryD8,0)DeliveryD8, ISNULL(DeliveryD9,0)DeliveryD9, ISNULL(DeliveryD10,0)DeliveryD10,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD11,0)DeliveryD11, ISNULL(DeliveryD12,0)DeliveryD12, ISNULL(DeliveryD13,0)DeliveryD13, ISNULL(DeliveryD14,0)DeliveryD14, ISNULL(DeliveryD15,0)DeliveryD15,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD16,0)DeliveryD16, ISNULL(DeliveryD17,0)DeliveryD17, ISNULL(DeliveryD18,0)DeliveryD18, ISNULL(DeliveryD19,0)DeliveryD19, ISNULL(DeliveryD20,0)DeliveryD20,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD21,0)DeliveryD21, ISNULL(DeliveryD22,0)DeliveryD22, ISNULL(DeliveryD23,0)DeliveryD23, ISNULL(DeliveryD24,0)DeliveryD24, ISNULL(DeliveryD25,0)DeliveryD25,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD26,0)DeliveryD26, ISNULL(DeliveryD27,0)DeliveryD27, ISNULL(DeliveryD28,0)DeliveryD28, ISNULL(DeliveryD29,0)DeliveryD29, ISNULL(DeliveryD30,0)DeliveryD30,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD31,0)DeliveryD31, " & vbCrLf & _
                      " 	ISNULL(DeliveryD1Old,0)DeliveryD1Old, ISNULL(DeliveryD2Old,0)DeliveryD2Old, ISNULL(DeliveryD3Old,0)DeliveryD3Old, ISNULL(DeliveryD4Old,0)DeliveryD4Old, ISNULL(DeliveryD5Old,0)DeliveryD5Old,  " & vbCrLf & _
                      " 	ISNULL(DeliveryD6Old,0)DeliveryD6Old, ISNULL(DeliveryD7Old,0)DeliveryD7Old, ISNULL(DeliveryD8Old,0)DeliveryD8Old, ISNULL(DeliveryD9Old,0)DeliveryD9Old, ISNULL(DeliveryD10Old,0)DeliveryD10Old,  "

            ls_SQL = ls_SQL + " 	ISNULL(DeliveryD11Old,0)DeliveryD11Old, ISNULL(DeliveryD12Old,0)DeliveryD12Old, ISNULL(DeliveryD13Old,0)DeliveryD13Old, ISNULL(DeliveryD14Old,0)DeliveryD14Old, ISNULL(DeliveryD15Old,0)DeliveryD15Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD16Old,0)DeliveryD16Old, ISNULL(DeliveryD17Old,0)DeliveryD17Old, ISNULL(DeliveryD18Old,0)DeliveryD18Old, ISNULL(DeliveryD19Old,0)DeliveryD19Old, ISNULL(DeliveryD20Old,0)DeliveryD20Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD21Old,0)DeliveryD21Old, ISNULL(DeliveryD22Old,0)DeliveryD22Old, ISNULL(DeliveryD23Old,0)DeliveryD23Old, ISNULL(DeliveryD24Old,0)DeliveryD24Old, ISNULL(DeliveryD25Old,0)DeliveryD25Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD26Old,0)DeliveryD26Old, ISNULL(DeliveryD27Old,0)DeliveryD27Old, ISNULL(DeliveryD28Old,0)DeliveryD28Old, ISNULL(DeliveryD29Old,0)DeliveryD29Old, ISNULL(DeliveryD30Old,0)DeliveryD30Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD31Old,0)DeliveryD31Old " & vbCrLf & _
                              " from  " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select * from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select '1' NoUrutDesc, 'BY AFFILIATE' AffiliateName " & vbCrLf & _
                              " 		union all " & vbCrLf

            ls_SQL = ls_SQL + " 		select '2' NoUrutDesc, 'BY PASI' AffiliateName " & vbCrLf & _
                              " 		union all " & vbCrLf & _
                              " 		select '3' NoUrutDesc, 'BY SUPPLIER' AffiliateName " & vbCrLf & _
                              " 	)tbla " & vbCrLf & _
                              " 	cross join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select  " & vbCrLf & _                              
                              " 			b.PartNo, b.PartNo PartNo1 " & vbCrLf & _
                              " 		from PO_Master a " & vbCrLf & _
                              " 		inner join po_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 		inner join MS_UnitCls d on d.UnitCls = c.UnitCls " & vbCrLf & _
                              " 		where " & pWhere & " " & vbCrLf & _
                              " 	)tb1b " & vbCrLf & _
                              " )tbl1 " & vbCrLf & _
                              " left join " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		convert(char,row_number() over (order by b.PartNo asc))as NoUrut, 'BY AFFILIATE' AffiliateName, '1' NoUrutDesc,  " & vbCrLf & _
                              " 		b.PartNo, b.PartNo PartNo1, c.PartName, case when b.KanbanCls = '1' then 'YES' else 'NO' end KanbanCls, d.Description UnitDesc, " & vbCrLf & _
                              " 		e.MOQ, e.QtyBox, c.Maker, b.POQty, 0 POQtyOld, " & vbCrLf

            ls_SQL = ls_SQL + " 		ISNULL(ForecastN1,0) ForecastN1 ,  " & vbCrLf & _
                              " 		ISNULL(ForecastN2,0) ForecastN2 ,  " & vbCrLf & _
                              " 		ISNULL(ForecastN3,0) ForecastN3,  " & vbCrLf & _
                              "  		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5, " & vbCrLf & _
                              "  		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,  " & vbCrLf & _
                              "  		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,  " & vbCrLf & _
                              "  		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,  " & vbCrLf & _
                              "  		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,  " & vbCrLf & _
                              "  		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,  " & vbCrLf & _
                              "  		b.DeliveryD31, " & vbCrLf & _
                              "  		0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old, " & vbCrLf

            ls_SQL = ls_SQL + "  		0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old,  " & vbCrLf & _
                              "  		0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old,  " & vbCrLf & _
                              "  		0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old,  " & vbCrLf & _
                              "  		0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old,  " & vbCrLf & _
                              "  		0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old,  " & vbCrLf & _
                              "  		0 DeliveryD31Old " & vbCrLf & _
                              " 		from PO_Master a " & vbCrLf & _
                              " 	inner join po_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	Left join MS_PartMapping e on e.PartNo = b.PartNo and e.AffiliateID = b.AffiliateID and e.SupplierID = b.SupplierID " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf

            ls_SQL = ls_SQL + " 	where " & pWhere & " " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		'' NoUrut,'BY PASI' AffiliateName, '2' NoUrutDesc,  " & vbCrLf & _
                              " 		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, '' Maker, b.POQty, b.POQtyOld, " & vbCrLf & _
                              " 		ISNULL(ForecastN1,0) ForecastN1 ,  " & vbCrLf & _
                              " 		ISNULL(ForecastN2,0) ForecastN2 ,  " & vbCrLf & _
                              " 		ISNULL(ForecastN3,0) ForecastN3,  " & vbCrLf & _
                              "  		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5, " & vbCrLf & _
                              "  		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,  " & vbCrLf & _
                              "  		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,  " & vbCrLf

            ls_SQL = ls_SQL + "  		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,  " & vbCrLf & _
                              "  		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,  " & vbCrLf & _
                              "  		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,  " & vbCrLf & _
                              "  		b.DeliveryD31, " & vbCrLf & _
                              "  		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old, " & vbCrLf & _
                              "  		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,  " & vbCrLf & _
                              "  		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,  " & vbCrLf & _
                              "  		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,  " & vbCrLf & _
                              "  		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,  " & vbCrLf & _
                              "  		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,  " & vbCrLf & _
                              "  		b.DeliveryD31Old " & vbCrLf

            ls_SQL = ls_SQL + " 		from Affiliate_Master a " & vbCrLf & _
                              " 	inner join Affiliate_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "     left join po_detail f on f.PONo = b.PONo and f.SupplierID = b.SupplierID and f.AffiliateID = b.AffiliateID and f.PartNo = b.PartNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	Left join MS_PartMapping e on e.PartNo = b.PartNo and e.AffiliateID = b.AffiliateID and e.SupplierID = b.SupplierID " & vbCrLf & _
                              " 	where " & pWhere & "  " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		'' NoUrut, 'BY SUPPLIER' AffiliateName, '3' NoUrutDesc,  " & vbCrLf & _
                              " 		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, '' Maker, b.POQty, b.POQtyOld, " & vbCrLf

            ls_SQL = ls_SQL + " 		ISNULL(ForecastN1,0) ForecastN1 ,  " & vbCrLf & _
                              " 		ISNULL(ForecastN2,0) ForecastN2 ,  " & vbCrLf & _
                              " 		ISNULL(ForecastN3,0) ForecastN3,  " & vbCrLf & _
                              "  		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5, " & vbCrLf & _
                              "  		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,  " & vbCrLf & _
                              "  		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,  " & vbCrLf & _
                              "  		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,  " & vbCrLf & _
                              "  		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,  " & vbCrLf & _
                              "  		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,  " & vbCrLf & _
                              "  		b.DeliveryD31, " & vbCrLf & _
                              "  		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old, " & vbCrLf & _
                              "  		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,  " & vbCrLf

            ls_SQL = ls_SQL + "  		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,  " & vbCrLf & _
                              "  		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,  " & vbCrLf & _
                              "  		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,  " & vbCrLf & _
                              "  		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,  " & vbCrLf & _
                              "  		b.DeliveryD31Old " & vbCrLf & _
                              " 		from PO_MasterUpload a " & vbCrLf & _
                              " 	inner join PO_DetailUpload b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "     left join po_detail f on f.PONo = b.PONo and f.SupplierID = b.SupplierID and f.AffiliateID = b.AffiliateID and f.PartNo = b.PartNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	Left join MS_PartMapping e on e.PartNo = b.PartNo and e.AffiliateID = b.AffiliateID and e.SupplierID = b.SupplierID " & vbCrLf & _
                              " 	where " & pWhere & " " & vbCrLf

            ls_SQL = ls_SQL + " )tbl2 on tbl2.AffiliateName = tbl1.AffiliateName and tbl1.PartNo = tbl2.PartNo1 and tbl1.NoUrutDesc = tbl2.NoUrutDesc " & vbCrLf & _
                              " ORDER BY PartNo1, tbl1.NoUrutDesc, NoUrut" & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub bindPOStatus(Optional ByVal pUpdate As String = "", Optional ByVal pPONO As String = "", Optional ByVal pAffiliate As String = "")
        Dim ls_SQL As String = ""
        Dim ls_PONo As String = ""
        Dim ls_Affiliate As String = ""

        If pPONO <> "" Then
            ls_PONo = pPONO
        Else
            ls_PONo = Trim(cboPONo.Text)
        End If

        If pAffiliate <> "" Then
            ls_Affiliate = pAffiliate
        Else
            ls_Affiliate = Trim(cboAffiliateCode.Text)
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  DISTINCT CASE WHEN CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls,POM.SupplierID,SupplierName" & vbCrLf & _
                  "     ,CASE WHEN DeliveryByPASICls='0' THEN 'DIRECT TO AFFILIATE' ELSE 'VIA PASI' END DeliveryByPASICls,ISNULL(ShipCls,'')ShipCls,KanbanCls,ISNULL(Remarks,'')Remarks," & vbCrLf & _
                  " 	POM.EntryDate, " & vbCrLf & _
                  " 	ISNULL(POM.EntryUser,'')EntryUser, " & vbCrLf & _
                  " 	POM.AffiliateApproveDate, " & vbCrLf & _
                  " 	ISNULL(POM.AffiliateApproveUser,'')AffiliateApproveUser, " & vbCrLf & _
                  " 	POM.PASISendAffiliateDate, " & vbCrLf & _
                  " 	ISNULL(POM.PASISendAffiliateUser,'')PASISendAffiliateUser, " & vbCrLf & _
                  " 	POM.SupplierApproveDate, " & vbCrLf & _
                  " 	ISNULL(POM.SupplierApproveUser,'')SupplierApproveUser, " & vbCrLf & _
                  " 	POM.SupplierApprovePendingDate, " & vbCrLf & _
                  " 	ISNULL(POM.SupplierApprovePendingUser,'')SupplierApprovePendingUser, "

            ls_SQL = ls_SQL + " 	POM.SupplierUnApproveDate, " & vbCrLf & _
                              " 	ISNULL(POM.SupplierUnApproveUser,'')SupplierUnApproveUser, " & vbCrLf & _
                              " 	POM.PASIApproveDate, " & vbCrLf & _
                              " 	ISNULL(POM.PASIApproveUser,'')PASIApproveUser, " & vbCrLf & _
                              " 	POM.FinalApproveDate, " & vbCrLf & _
                              " 	ISNULL(POM.FinalApproveUser,'')FinalApproveUser  " & vbCrLf & _
                              " from PO_Master POM " & vbCrLf & _
                              " LEFT JOIN dbo.PO_Detail POD ON POM.PONo = POD.PONo AND POM.AffiliateID = POD.AffiliateID AND POD.SupplierID=POM.SupplierID " & vbCrLf & _
                              " LEFT JOIN dbo.PO_MasterUpload POMU ON POM.AffiliateID = POMU.AffiliateID AND POM.PONo = POMU.PONo AND POM.SupplierID = POMU.SupplierID " & vbCrLf & _
                              " LEFT JOIN dbo.MS_Supplier MS ON POD.SupplierID = MS.SupplierID" & vbCrLf & _
                              " where POM.PONo = '" & ls_PONo & "' and POM.AffiliateID = '" & ls_Affiliate & "' and POM.SupplierID = '" & txtSupplierCode.Text & "'"


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                If IsDBNull(ds.Tables(0).Rows(0)("EntryDate")) Then
                    txtEntryDate.Text = "-"
                    txtEntryUser.Text = "-"
                Else
                    txtEntryDate.Text = Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd HH:mm:ss")
                    txtEntryUser.Text = ds.Tables(0).Rows(0)("EntryUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")) Then
                    txtAffAppDate.Text = "-"
                    txtAffAppUser.Text = "-"
                Else
                    txtAffAppDate.Text = Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtAffAppUser.Text = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")) Then
                    txtSendDate.Text = "-"
                    txtSendUser.Text = "-"
                Else
                    txtSendDate.Text = Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                    txtSendUser.Text = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")) Then
                    txtSuppAppDate.Text = "-"
                    txtSuppAppUser.Text = "-"
                Else
                    txtSuppAppDate.Text = Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtSuppAppUser.Text = ds.Tables(0).Rows(0)("SupplierApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")) Then
                    txtSuppPendDate.Text = "-"
                    txtSuppPendUser.Text = "-"
                Else
                    txtSuppPendDate.Text = Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd HH:mm:ss")
                    txtSuppPendUser.Text = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")) Then
                    txtSuppUnpDate.Text = "-"
                    txtSuppUnpUser.Text = "-"
                Else
                    txtSuppUnpDate.Text = Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtSuppUnpUser.Text = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")) Then
                    txtPASIAppDate.Text = "-"
                    txtPASIAppUser.Text = "-"
                Else
                    txtPASIAppDate.Text = Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtPASIAppUser.Text = ds.Tables(0).Rows(0)("PASIApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")) Then
                    txtAffFinalAppDate.Text = ""
                    txtAffFinalAppUser.Text = ""
                Else
                    txtAffFinalAppDate.Text = Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtAffFinalAppUser.Text = ds.Tables(0).Rows(0)("FinalApproveUser")
                End If

                If pUpdate = "update" Then
                    txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                    txtSupplierCode.Text = ds.Tables(0).Rows(0)("SupplierID")
                    txtSupplierName.Text = ds.Tables(0).Rows(0)("SupplierName")
                    txtShipBy.Text = ds.Tables(0).Rows(0)("ShipCls")
                    txtDeliveryBy.Text = ds.Tables(0).Rows(0)("DeliveryByPASICls")
                    rblPOKanban.Value = ds.Tables(0).Rows(0)("KanbanCls")
                    txtRemarks.Value = ds.Tables(0).Rows(0)("Remarks")
                    txtEntryDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("EntryDate")), "", Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtEntryUser.Text = ds.Tables(0).Rows(0)("EntryUser")
                    txtAffAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")), "", Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtAffAppUser.Text = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                    txtSendDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")), "", Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtSendUser.Text = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                    txtSuppAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtSuppAppUser.Text = ds.Tables(0).Rows(0)("SupplierApproveUser")
                    txtSuppPendDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtSuppPendUser.Text = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                    txtSuppUnpDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtSuppUnpUser.Text = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                    txtPASIAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")), "", Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtPASIAppUser.Text = ds.Tables(0).Rows(0)("PASIApproveUser")
                    txtAffFinalAppDate.Text = If(IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")), "", Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    txtAffFinalAppUser.Text = ds.Tables(0).Rows(0)("FinalApproveUser")

                    ButtonApprove.JSProperties("cpKanban") = ds.Tables(0).Rows(0)("KanbanCls")
                    ButtonApprove.JSProperties("cpEntryDate") = If(IsDBNull(ds.Tables(0).Rows(0)("EntryDate")), "", Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpEntryUser") = ds.Tables(0).Rows(0)("EntryUser")
                    ButtonApprove.JSProperties("cpAffAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")), "", Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpAffAppUser") = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                    ButtonApprove.JSProperties("cpSendDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")), "", Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpSendUser") = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                    ButtonApprove.JSProperties("cpSuppAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpSuppAppUser") = ds.Tables(0).Rows(0)("SupplierApproveUser")
                    ButtonApprove.JSProperties("cpSuppAppPendingDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpSuppAppPendingUser") = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                    ButtonApprove.JSProperties("cpSuppUnApproveDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpSuppUnApproveUser") = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                    ButtonApprove.JSProperties("cpPASIAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")), "", Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpPASIAppUser") = ds.Tables(0).Rows(0)("PASIApproveUser")
                    ButtonApprove.JSProperties("cpFinalAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")), "", Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    ButtonApprove.JSProperties("cpFinalAppUser") = ds.Tables(0).Rows(0)("FinalApproveUser")

                    Call clsMsg.DisplayMessage(lblInfo, "1009", clsMessage.MsgType.InformationMessage)
                    ButtonApprove.JSProperties("cpMessage") = lblInfo.Text
                Else
                    'dtPeriod.Value = ds.Tables(0).Rows(0)("Period")
                    'cboAffiliateCode.Text = ds.Tables(0).Rows(0)("AffiliateID")
                    'txtAffiliateName.Text = ds.Tables(0).Rows(0)("AffiliateName")
                    'cboPONo.Text = ds.Tables(0).Rows(0)("PONo")
                    txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                    txtSupplierCode.Text = ds.Tables(0).Rows(0)("SupplierID")
                    txtSupplierName.Text = ds.Tables(0).Rows(0)("SupplierName")
                    txtShipBy.Text = ds.Tables(0).Rows(0)("ShipCls")
                    txtDeliveryBy.Text = ds.Tables(0).Rows(0)("DeliveryByPASICls")
                    rblPOKanban.Value = ds.Tables(0).Rows(0)("KanbanCls")
                    txtRemarks.Value = ds.Tables(0).Rows(0)("Remarks")
                    grid.JSProperties("cpKanban") = ds.Tables(0).Rows(0)("KanbanCls")
                    grid.JSProperties("cpEntryDate") = If(IsDBNull(ds.Tables(0).Rows(0)("EntryDate")), "", Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpEntryUser") = ds.Tables(0).Rows(0)("EntryUser")
                    grid.JSProperties("cpAffAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")), "", Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpAffAppUser") = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                    grid.JSProperties("cpSendDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")), "", Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpSendUser") = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                    grid.JSProperties("cpSuppAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpSuppAppUser") = ds.Tables(0).Rows(0)("SupplierApproveUser")
                    grid.JSProperties("cpSuppAppPendingDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")), "", Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpSuppAppPendingUser") = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                    grid.JSProperties("cpSuppUnApproveDate") = If(IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")), "", Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpSuppUnApproveUser") = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                    grid.JSProperties("cpPASIAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")), "", Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpPASIAppUser") = ds.Tables(0).Rows(0)("PASIApproveUser")
                    grid.JSProperties("cpFinalAppDate") = If(IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")), "", Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd hh:mm:ss"))
                    grid.JSProperties("cpFinalAppUser") = ds.Tables(0).Rows(0)("FinalApproveUser")
                End If
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub bindHeader(ByVal pPONO As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	case when DeliveryByPASICls = '1' then 'VIA PASI' else 'VIA SUPPLIER' end DeliveryByPASICls, " & vbCrLf & _
                  " 	case when CommercialCls = '1' then 'YES' else 'NO' end CommercialCls, " & vbCrLf & _
                  " 	a.SupplierID, ShipCls, isnull(Remarks,'')Remarks " & vbCrLf & _
                  " from PO_Master a " & vbCrLf & _
                  " left join PO_MasterUpload b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                  " where a.PONo = '" & pPONO & "' and a.AffiliateID = '" & Session("AffiliateID") & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtDeliveryBy.Text = ds.Tables(0).Rows(0)("DeliveryByPASICls")
                txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                txtShipBy.Text = ds.Tables(0).Rows(0)("ShipCls")
                txtRemarks.Text = ds.Tables(0).Rows(0)("Remarks")
                Session("SupplierID") = ds.Tables(0).Rows(0)("SupplierID")

                cbPONo.JSProperties("cpDelivery") = txtDeliveryBy.Text
                cbPONo.JSProperties("cpCommercial") = txtCommercial.Text
                cbPONo.JSProperties("cpShip") = txtShipBy.Text
                cbPONo.JSProperties("cpRemarks") = txtRemarks.Text
            Else
                cbPONo.JSProperties("cpDelivery") = ""
                cbPONo.JSProperties("cpCommercial") = ""
                cbPONo.JSProperties("cpShip") = ""
                cbPONo.JSProperties("cpRemarks") = ""
                Session("SupplierID") = ""
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' AffiliateName, '' PartNo, '' PartNo1, '' PartName, '' KanbanCls, '' UnitDesc, '' MOQ, '' QtyBox, '' Maker, " & vbCrLf & _
                  " 0 POQty, 0 POQtyOld, '' CurrDesc, '' Price, '' Amount, 0 ForecastN1, 0 ForecastN2, 0 ForecastN3,   " & vbCrLf & _
                  " 0 DeliveryD1, 0 DeliveryD2, 0 DeliveryD3, 0 DeliveryD4, 0 DeliveryD5, " & vbCrLf & _
                  " 0 DeliveryD6, 0 DeliveryD7, 0 DeliveryD8, 0 DeliveryD9, 0 DeliveryD10, " & vbCrLf & _
                  " 0 DeliveryD11, 0 DeliveryD12, 0 DeliveryD13, 0 DeliveryD14, 0 DeliveryD15, " & vbCrLf & _
                  " 0 DeliveryD16, 0 DeliveryD17, 0 DeliveryD18, 0 DeliveryD19, 0 DeliveryD20, " & vbCrLf & _
                  " 0 DeliveryD21, 0 DeliveryD22, 0 DeliveryD23, 0 DeliveryD24, 0 DeliveryD25, " & vbCrLf & _
                  " 0 DeliveryD26, 0 DeliveryD27, 0 DeliveryD28, 0 DeliveryD29, 0 DeliveryD30, " & vbCrLf & _
                  " 0 DeliveryD31, " & vbCrLf & _
                  " 0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old, " & vbCrLf & _
                  " 0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old, " & vbCrLf & _
                  " 0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old, " & vbCrLf & _
                  " 0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old, " & vbCrLf & _
                  " 0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old, " & vbCrLf & _
                  " 0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old, " & vbCrLf & _
                  " 0 DeliveryD31Old "

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

    Private Function getApp(ByVal ls_value As String) As Boolean
        Dim ls_SQL As String = ""
        Dim doneApp As Boolean = False
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " SELECT * FROM PO_Master " & vbCrLf & _
                  " WHERE PONO='" & ls_value & "' AND SupplierID = '" & pub_SupplierID & "' AND AffiliateID = '" & pub_AffiliateID & "' AND  " & vbCrLf & _
                  " (ISNULL(PASIApproveDate,'') <> '' OR  " & vbCrLf & _
                  " ISNULL(FinalApproveDate,'') <> '') "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                doneApp = True
            End If
        End Using
        Return doneApp
    End Function

    Private Sub uf_Approve()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PO_Master set PASIApproveDate = getdate(), PASIApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' and PONo = '" & Trim(cboPONo.Text) & "' and SupplierID = '" & Trim(txtSupplierCode.Text) & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
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
                SmtpClient = Trim(ds.Tables(0).Rows(0)("SMTP"))
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
                     "  select 'AFF' flag,affiliatepocc, affiliatepoto,FromEmail = '' from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,affiliatepocc,affiliatepoto='',FromEmail = affiliatepoto from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Function EmailToEmailCCNotif(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                     "  select 'AFF' flag,affiliatepocc, affiliatepoto,FromEmail = '' from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,affiliatepocc,affiliatepoto='',FromEmail = affiliatepoto from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Sub sendEmailtoAffiliate()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""
            Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
            Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
            Dim ls_Body As String = ""
            Dim pApproval As String
            If txtPASIAppDate.Text <> "" Then
                pApproval = "1"
            Else
                pApproval = "0"
            End If

            '"POFinalApproval.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>
            '&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%
            '>&t5=<%#GetRemarks(Container)%>&t6=<%#GetFinalApproval(Container)%>&t7=<%#GetDeliveryBy(Container)%>&Session=~/PurchaseOrder/POFinalApprovalList.aspx"
            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNameAffiliate & "/PurchaseOrder/POFinalApproval.aspx?id2=" & clsNotification.EncryptURL(cboPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(cboAffiliateCode.Text.Trim) & _
                                           "&t2=" & clsNotification.EncryptURL(txtAffiliateName.Text.Trim) & "&t3=" & clsNotification.EncryptURL(dtPeriod.Value) & "&t4=" & clsNotification.EncryptURL(txtSupplierCode.Text.Trim) & _
                                           "&t5=" & clsNotification.EncryptURL(txtRemarks.Text.Trim) & "&t6=" & clsNotification.EncryptURL(pApproval) & "&t7=" & clsNotification.EncryptURL(txtDeliveryBy.Text.Trim) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrder/POFinalApprovalList.aspx")


            ls_Body = clsNotification.GetNotification("13", ls_URl, cboPONo.Text.Trim)

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCNotif(Trim(cboAffiliateCode.Text), "PASI", "")
            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
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
            mailMessage.Subject = "PASI Approval PONo: " & Trim(cboPONo.Text)

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
            Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
            Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
            Dim ls_Body As String = ""
            Dim pApproval As String
            If txtPASIAppDate.Text <> "" Then
                pApproval = "1"
            Else
                pApproval = "0"
            End If

            '"AffiliateOrderAppDetail.aspx?id=<%#GetRowValue(Container)%>
            '&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetAffiliateName(Container)%>
            '&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>
            '&t5=<%#GetRemarks(Container)%>&t6=<%#GetFinalApproval(Container)%>
            '&t7=<%#GetDeliveryBy(Container)%>&t8=<%#GetShipCls(Container)%>
            '&t9=<%#GetCommercialCls(Container)%>&t10=<%#GetSupplierName(Container)%>&Session=~/AffiliateOrder/AffiliateOrderAppList.aspx"
            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderAppDetail.aspx?id2=" & clsNotification.EncryptURL(cboPONo.Text.Trim) & _
                "&t1=" & clsNotification.EncryptURL(cboAffiliateCode.Text.Trim) & "&t2=" & clsNotification.EncryptURL(txtAffiliateName.Text.Trim) & _
                "&t3=" & clsNotification.EncryptURL(dtPeriod.Value) & "&t4=" & clsNotification.EncryptURL(txtSupplierCode.Text.Trim) & _
                "&t5=" & clsNotification.EncryptURL(txtRemarks.Text.Trim) & "&t6=" & clsNotification.EncryptURL(pApproval) & _
                "&t7=" & clsNotification.EncryptURL(txtDeliveryBy.Text.Trim) & "&t8=" & clsNotification.EncryptURL(txtShipBy.Text.Trim) & _
                "&t9=" & clsNotification.EncryptURL(txtCommercial.Text.Trim) & "&t10=" & clsNotification.EncryptURL(txtSupplierName.Text.Trim) & _
                "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderAppList.aspx")


            ls_Body = clsNotification.GetNotification("13", ls_URl, cboPONo.Text.Trim)

            'ls_Body = ls_Line1 & vbCr & ls_Line2 & "PO No:" & Trim(cboPONo.Text) & vbCr & vbCr & ls_Line3 & vbCr & ls_Line4 & ls_Line5 & vbCr & ls_Line6 & vbCr & ls_Line7 & vbCr & ls_Line8

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCC(Trim(cboAffiliateCode.Text), "PASI", "")
            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
                receiptEmail = receiptCCEmail
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
            mailMessage.Subject = "PASI Approval PONo: " & Trim(cboPONo.Text)

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

    Private Function GetNotification(ByVal pNotificationCode As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "select Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8 from ms_notification where notificationcode = '" & pNotificationCode & "'" & vbCrLf

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