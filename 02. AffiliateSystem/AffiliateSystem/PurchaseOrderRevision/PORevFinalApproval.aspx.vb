Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports System.Net
Imports System.Net.Mail

Public Class PORevFinalApproval
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "D04"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String, pub_Remarks As String, pub_Revision As String
    Dim pub_FinalApproval As String, pub_DeliveyBy As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim flag As Boolean = True
    Dim clsPONo As New clsPO
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
                    Session("MenuDesc") = "PO FINAL APPROVAL"
                    pub_PONo = Request.QueryString("id")
                    pub_Ship = Request.QueryString("t1")
                    pub_Commercial = Request.QueryString("t2")
                    pub_Period = Request.QueryString("t3")
                    Session("SupplierID") = Request.QueryString("t4")
                    pub_Remarks = Request.QueryString("t5")
                    pub_FinalApproval = Request.QueryString("t6")
                    pub_DeliveyBy = Request.QueryString("t7")
                    pub_Revision = Request.QueryString("t8")

                    dtPeriodFrom.Value = pub_Period
                    cboPartNo.Text = pub_PONo
                    cboPartNoRev.Text = pub_Revision
                    txtShip.Text = pub_Ship
                    txtRemarks.Text = pub_Remarks
                    txtDelivery.Text = IIf(pub_DeliveyBy = "1", "VIA PASI", "DIRECT AFFILIATE")
                    txtCommercial.Text = pub_Commercial

                    Session("Mode") = "Update"

                    txtPOKanban.Text = clsPONo.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID"))
                    rdrDiff1.Checked = True

                    bindData()
                    bindPOStatus()

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    cboPartNo.ReadOnly = True
                    cboPartNo.BackColor = Color.FromName("#CCCCCC")
                    cboPartNoRev.ReadOnly = True
                    cboPartNoRev.BackColor = Color.FromName("#CCCCCC")
                    dtPeriodFrom.ReadOnly = True
                    dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")

                    If pub_FinalApproval <> "1" Then
                        btnApprove.Enabled = False
                    Else
                        btnApprove.Enabled = True
                    End If
                ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                    Session("MenuDesc") = "PO FINAL APPROVAL"
                    pub_PONo = clsNotification.DecryptURL(Request.QueryString("id2"))
                    pub_Ship = clsNotification.DecryptURL(Request.QueryString("t1"))
                    pub_Commercial = clsNotification.DecryptURL(Request.QueryString("t2"))
                    pub_Period = clsNotification.DecryptURL(Request.QueryString("t3"))
                    Session("SupplierID") = clsNotification.DecryptURL(Request.QueryString("t4"))
                    pub_Remarks = clsNotification.DecryptURL(Request.QueryString("t5"))
                    pub_FinalApproval = clsNotification.DecryptURL(Request.QueryString("t6"))
                    pub_DeliveyBy = clsNotification.DecryptURL(Request.QueryString("t7"))
                    pub_Revision = clsNotification.DecryptURL(Request.QueryString("t8"))

                    dtPeriodFrom.Value = pub_Period
                    cboPartNo.Text = pub_PONo
                    cboPartNoRev.Text = pub_Revision
                    txtShip.Text = pub_Ship
                    txtRemarks.Text = pub_Remarks
                    txtDelivery.Text = IIf(pub_DeliveyBy = "1", "VIA PASI", "DIRECT AFFILIATE")
                    txtCommercial.Text = pub_Commercial

                    Session("Mode") = "Update"

                    txtPOKanban.Text = clsPONo.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID"))
                    rdrDiff1.Checked = True

                    bindData()
                    bindPOStatus()

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    cboPartNo.ReadOnly = True
                    cboPartNo.BackColor = Color.FromName("#CCCCCC")
                    dtPeriodFrom.ReadOnly = True
                    dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")

                    If pub_FinalApproval <> "1" Then
                        btnApprove.Enabled = False
                    Else
                        btnApprove.Enabled = True
                    End If
                Else
                    Session("MenuDesc") = "PO FINAL APPROVAL"
                    Session("Mode") = "New"
                    lblInfo.Text = ""
                    btnSubMenu.Text = "Back"                  
                    dtPeriodFrom.Value = Now
                    up_FillCombo(dtPeriodFrom.Value)
                    cboPartNo.Focus()
                    cboPartNoRev.Items.Clear()
                    'rdrKanban1.Checked = True
                End If
            Else
                Session("Mode") = "New"
                dtPeriodFrom.Value = Now
                up_FillCombo(dtPeriodFrom.Value)
                cboPartNo.Focus()
                cboPartNoRev.Items.Clear()
                'rdrKanban1.Checked = True
                rdrDiff1.Checked = True
            End If

            lblInfo.Text = ""

        ElseIf IsCallback Then
            If grid.VisibleRowCount = 0 Then Exit Sub
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/PurchaseOrderRevision/PORevFinalApprovalList.aspx")
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
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()
                    Call bindPOStatus()

                    grid.JSProperties("cpSearch") = "search"

                    'Dim TempASPxGridViewCellMerger As ASPxGridViewCellMerger = New ASPxGridViewCellMerger(grid, "NoUrut,PartNo,PartName,KanbanCls,UnitDesc,MOQ,QtyBox,Maker")
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    grid.JSProperties("cpSearch") = ""

                Case "save"
                    uf_Approve(Split(e.Parameters, "|")(2))
                    sendEmail()
                    bindPOStatus("update")

                    If clsPONo.Check_CreateKanban(Split(e.Parameters, "|")(1), Session("AffiliateID"), Session("SupplierID")) = True Then
                        'btnApprove.Enabled = False
                        grid.JSProperties("cpButton") = "YES"
                    Else
                        'btnApprove.Enabled = True
                        grid.JSProperties("cpButton") = "NO"
                    End If
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
        'uf_Approve()
        'sendEmail()
        'bindPOStatus("update")
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        grid.CollapseAll()

        cboPartNo.Text = ""
        txtShip.Text = ""
        txtRemarks.Text = ""
        txtDelivery.Text = ""
        txtCommercial.Text = ""

        cboPartNo.ReadOnly = False
        cboPartNo.BackColor = Color.FromName("#FFFFFF")
        cboPartNoRev.ReadOnly = False
        cboPartNoRev.BackColor = Color.FromName("#FFFFFF")
        dtPeriodFrom.ReadOnly = False
        dtPeriodFrom.BackColor = Color.FromName("#FFFFFF")

        dtPeriodFrom.Value = Now

        'rdrKanban1.Checked = True
        txtPOKanban.Text = ""
        rdrDiff1.Checked = True

        up_FillCombo(dtPeriodFrom.Value)

        cboPartNoRev.Items.Clear()

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
                    If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" _
                        Or e.DataColumn.FieldName = "ForecastN1" Or e.DataColumn.FieldName = "ForecastN2" _
                        Or e.DataColumn.FieldName = "ForecastN3" Or e.DataColumn.FieldName = "Maker" Then
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
                    If e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "QtyBox" _
                        Or e.DataColumn.FieldName = "ForecastN1" Or e.DataColumn.FieldName = "ForecastN2" _
                        Or e.DataColumn.FieldName = "ForecastN3" Or e.DataColumn.FieldName = "Maker" Then
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

    Private Sub cboPartNoRev_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPartNoRev.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Dim pPeriod As String = Mid(pAction, 12, 4) + "-" + clsGlobal.uf_GetShortMonth(Mid(pAction, 5, 3)) + "-" + "01"
        Dim pPONo As String = Split(e.Parameter, "|")(1)
        up_FillComboRev(pPeriod, pPONo)
    End Sub

    Private Sub cboPartNo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPartNo.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Dim pPeriod As String = Mid(pAction, 12, 4) + "-" + clsGlobal.uf_GetShortMonth(Mid(pAction, 5, 3)) + "-" + "01"
        up_FillCombo(pPeriod)
    End Sub

    Private Sub ButtonPartNo_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonPartNo.Callback
        Call up_GridLoadWhenEventChange()
        Call bindHeader(Split(e.Parameter, "|")(0), Split(e.Parameter, "|")(1))
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        'If rdrKanban2.Checked = True Then
        '    pWhereKanban = " and b.KanbanCls = '1'"
        'End If

        'If rdrKanban3.Checked = True Then
        '    pWhereKanban = " and b.KanbanCls = '0'"
        'End If

        If rdrDiff2.Checked = True Then
            pWhereDifference = " where POQty <> POQtyOld or  " & vbCrLf & _
                  " DeliveryD1 <> DeliveryD1Old or DeliveryD2 <> DeliveryD2Old or DeliveryD3 <> DeliveryD3Old or DeliveryD4 <> DeliveryD4Old or DeliveryD5 <> DeliveryD5Old or " & vbCrLf & _
                  " DeliveryD6 <> DeliveryD6Old or DeliveryD7 <> DeliveryD7Old or DeliveryD8 <> DeliveryD8Old or DeliveryD9 <> DeliveryD9Old or DeliveryD10 <> DeliveryD10Old or " & vbCrLf & _
                  " DeliveryD11 <> DeliveryD11Old or DeliveryD12 <> DeliveryD12Old or DeliveryD13 <> DeliveryD13Old or DeliveryD14 <> DeliveryD14Old or DeliveryD15 <> DeliveryD15Old or " & vbCrLf & _
                  " DeliveryD16 <> DeliveryD16Old or DeliveryD17 <> DeliveryD17Old or DeliveryD18 <> DeliveryD18Old or DeliveryD19 <> DeliveryD19Old or DeliveryD20 <> DeliveryD20Old or " & vbCrLf & _
                  " DeliveryD21 <> DeliveryD21Old or DeliveryD22 <> DeliveryD22Old or DeliveryD23 <> DeliveryD23Old or DeliveryD24 <> DeliveryD24Old or DeliveryD25 <> DeliveryD25Old or " & vbCrLf & _
                  " DeliveryD26 <> DeliveryD26Old or DeliveryD27 <> DeliveryD27Old or DeliveryD28 <> DeliveryD28Old or DeliveryD29 <> DeliveryD29Old or DeliveryD30 <> DeliveryD30Old or " & vbCrLf & _
                  " DeliveryD31 <> DeliveryD31Old "
        End If

        If rdrDiff3.Checked = True Then
            pWhereDifference = " where POQty = POQtyOld and  " & vbCrLf & _
                  " DeliveryD1 = DeliveryD1Old = DeliveryD2 = DeliveryD2Old and DeliveryD3 = DeliveryD3Old and DeliveryD4 = DeliveryD4Old and DeliveryD5 = DeliveryD5Old and " & vbCrLf & _
                  " DeliveryD6 = DeliveryD6Old = DeliveryD7 = DeliveryD7Old and DeliveryD8 = DeliveryD8Old and DeliveryD9 = DeliveryD9Old and DeliveryD10 = DeliveryD10Old and " & vbCrLf & _
                  " DeliveryD11 = DeliveryD11Old = DeliveryD12 = DeliveryD12Old and DeliveryD13 = DeliveryD13Old and DeliveryD14 = DeliveryD14Old and DeliveryD15 = DeliveryD15Old and " & vbCrLf & _
                  " DeliveryD16 = DeliveryD16Old = DeliveryD17 = DeliveryD17Old and DeliveryD18 = DeliveryD18Old and DeliveryD19 = DeliveryD19Old and DeliveryD20 = DeliveryD20Old and " & vbCrLf & _
                  " DeliveryD21 = DeliveryD21Old = DeliveryD22 = DeliveryD22Old and DeliveryD23 = DeliveryD23Old and DeliveryD24 = DeliveryD24Old and DeliveryD25 = DeliveryD25Old and " & vbCrLf & _
                  " DeliveryD26 = DeliveryD26Old = DeliveryD27 = DeliveryD27Old and DeliveryD28 = DeliveryD28Old and DeliveryD29 = DeliveryD29Old and DeliveryD30 = DeliveryD30Old and " & vbCrLf & _
                  " DeliveryD31 = DeliveryD31Old "
        End If

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
                  " 	ISNULL(DeliveryD6Old,0)DeliveryD6Old, ISNULL(DeliveryD7Old,0)DeliveryD7Old, ISNULL(DeliveryD8Old,0)DeliveryD8Old, ISNULL(DeliveryD9Old,0)DeliveryD9Old, ISNULL(DeliveryD10Old,0)DeliveryD10Old,  " & vbCrLf

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
                              " 			--row_number() over (order by b.PartNo asc) as NoUrut, " & vbCrLf & _
                              " 			b.PartNo, b.PartNo PartNo1 " & vbCrLf & _
                              " 		from PORev_Master a " & vbCrLf & _
                              " 		inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf

            ls_SQL = ls_SQL + " 		inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 		inner join MS_UnitCls d on d.UnitCls = c.UnitCls " & vbCrLf & _
                              " 		where a.PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' " & pWhereKanban & " and a.PORevNo = '" & cboPartNoRev.Text & "'" & vbCrLf & _
                              " 	)tb1b " & vbCrLf & _
                              " )tbl1 " & vbCrLf & _
                              " left join " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		convert(char,row_number() over (order by b.PartNo asc))as NoUrut, 'BY AFFILIATE' AffiliateName, '1' NoUrutDesc,  " & vbCrLf & _
                              " 		b.PartNo, b.PartNo PartNo1, c.PartName, case when c.KanbanCls = '1' then 'Yes' else 'No' end KanbanCls, d.Description UnitDesc, " & vbCrLf & _
                              " 		e.MOQ, e.QtyBox, c.Maker, b.POQty, 0 POQtyOld, " & vbCrLf

            ls_SQL = ls_SQL + " 		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
                              " 		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
                              " 		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
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
                              " 		from PORev_Master a " & vbCrLf & _
                              " 	inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	left join MS_PartMapping e on b.PartNo = e.PartNo and b.AffiliateID = e.AffiliateID and b.SupplierID = e.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 	where a.PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' and a.PORevNo = '" & cboPartNoRev.Text & "'" & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		'' NoUrut,'BY PASI' AffiliateName, '2' NoUrutDesc,  " & vbCrLf & _
                              " 		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, c.Maker, b.POQty, b.POQtyOld, " & vbCrLf & _
                              " 		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
                              " 		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
                              " 		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
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

            ls_SQL = ls_SQL + " 		from AffiliateRev_Master a " & vbCrLf & _
                              " 	inner join AffiliateRev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	where a.PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' and a.PORevNo = '" & cboPartNoRev.Text & "'" & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		'' NoUrut, 'BY SUPPLIER' AffiliateName, '3' NoUrutDesc,  " & vbCrLf & _
                              " 		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, c.Maker, b.POQty, b.POQtyOld, " & vbCrLf & _
                              " 		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf

            ls_SQL = ls_SQL + " 		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
                              " 		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0),  " & vbCrLf & _
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
                              " 		from PORev_MasterUpload a " & vbCrLf & _
                              " 	inner join PORev_DetailUpload b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	where a.PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' and a.PORevNo = '" & cboPartNoRev.Text & "' " & vbCrLf

            ls_SQL = ls_SQL + " )tbl2 on tbl2.AffiliateName = tbl1.AffiliateName and tbl1.PartNo = tbl2.PartNo1 and tbl1.NoUrutDesc = tbl2.NoUrutDesc " & vbCrLf & _
                              " " & pWhereDifference & " "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub bindPOStatus(Optional ByVal pUpdate As String = "", Optional ByVal pPONO As String = "", Optional ByVal pPORevNo As String = "")
        Dim ls_SQL As String = ""
        Dim ls_PONo As String = ""
        Dim ls_PORevNo As String = cboPartNoRev.Text

        If pPONO <> "" Then
            ls_PONo = pPONO
        Else
            ls_PONo = cboPartNo.Text
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
                              " from PORev_Master where PONo = '" & ls_PONo & "' and AffiliateID = '" & Session("AffiliateID") & "' and PORevNo = '" & ls_PORevNo & "'"


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                If IsDBNull(ds.Tables(0).Rows(0)("EntryDate")) Then
                    txtDate1.Text = "-"
                    txtUser1.Text = "-"
                Else
                    txtDate1.Text = Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser1.Text = ds.Tables(0).Rows(0)("EntryUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")) Then
                    txtDate2.Text = "-"
                    txtUser2.Text = "-"
                Else
                    txtDate2.Text = Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser2.Text = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")) Then
                    txtDate3.Text = "-"
                    txtUser3.Text = "-"
                Else
                    txtDate3.Text = Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser3.Text = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")) Then
                    txtDate4.Text = "-"
                    txtUser4.Text = "-"
                Else
                    txtDate4.Text = Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser4.Text = ds.Tables(0).Rows(0)("SupplierApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")) Then
                    txtDate5.Text = "-"
                    txtUser5.Text = "-"
                Else
                    txtDate5.Text = Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser5.Text = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")) Then
                    txtDate6.Text = "-"
                    txtUser6.Text = "-"
                Else
                    txtDate6.Text = Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser6.Text = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")) Then
                    txtDate7.Text = "-"
                    txtUser7.Text = "-"
                Else
                    txtDate7.Text = Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser7.Text = ds.Tables(0).Rows(0)("PASIApproveUser")
                End If
                If IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")) Then
                    txtDate8.Text = ""
                    txtUser8.Text = ""
                Else
                    txtDate8.Text = Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd HH:mm:ss")
                    txtUser8.Text = ds.Tables(0).Rows(0)("FinalApproveUser")
                End If

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

                    Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                    ButtonApprove.JSProperties("cpMessage") = lblInfo.Text
                Else
                    'Session("cpDate8") = txtDate8.Text
                    'Session("cpUser8") = txtUser8.Text
                    grid.JSProperties("cpDate1") = txtDate1.Text
                    grid.JSProperties("cpDate2") = txtDate2.Text
                    grid.JSProperties("cpDate3") = txtDate3.Text
                    grid.JSProperties("cpDate4") = txtDate4.Text
                    grid.JSProperties("cpDate5") = txtDate5.Text
                    grid.JSProperties("cpDate6") = txtDate6.Text
                    grid.JSProperties("cpDate7") = txtDate7.Text
                    grid.JSProperties("cpDate8") = txtDate8.Text

                    grid.JSProperties("cpUser1") = txtUser1.Text
                    grid.JSProperties("cpUser2") = txtUser2.Text
                    grid.JSProperties("cpUser3") = txtUser3.Text
                    grid.JSProperties("cpUser4") = txtUser4.Text
                    grid.JSProperties("cpUser5") = txtUser5.Text
                    grid.JSProperties("cpUser6") = txtUser6.Text
                    grid.JSProperties("cpUser7") = txtUser7.Text
                    grid.JSProperties("cpUser8") = txtUser8.Text
                End If
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub bindHeader(ByVal pPONO As String, ByVal pRevNo As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	case when DeliveryByPASICls = '1' then 'VIA PASI' else 'DIRECT TO AFFILIATE' end DeliveryByPASICls, " & vbCrLf & _
                  " 	case when pm.CommercialCls = '1' then 'YES' else 'NO' end CommercialCls, " & vbCrLf & _
                  " 	a.SupplierID, pm.ShipCls, isnull(Remarks,'')Remarks " & vbCrLf & _
                  " from PORev_Master pm left join PO_Master a on pm.PONo = a.PONo and pm.AffiliateID = a.AffiliateID and pm.SupplierID = a.SupplierID" & vbCrLf & _
                  " left join PORev_MasterUpload b on pm.PONo = b.PONo and pm.AffiliateID = b.AffiliateID and pm.PORevNo = b.PORevNo " & vbCrLf & _
                  " where pm.PONo = '" & pPONO & "' and pm.AffiliateID = '" & Session("AffiliateID") & "' and pm.PORevNo = '" & pRevNo & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtDelivery.Text = ds.Tables(0).Rows(0)("DeliveryByPASICls")
                txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                txtShip.Text = ds.Tables(0).Rows(0)("ShipCls")
                txtRemarks.Text = ds.Tables(0).Rows(0)("Remarks")
                Session("SupplierID") = ds.Tables(0).Rows(0)("SupplierID")
                txtPOKanban.Text = clsPONo.POKanban(pPONO, Session("AffiliateID"), Session("SupplierID"))

                ButtonPartNo.JSProperties("cpDelivery") = txtDelivery.Text
                ButtonPartNo.JSProperties("cpCommercial") = txtCommercial.Text
                ButtonPartNo.JSProperties("cpShip") = txtShip.Text
                ButtonPartNo.JSProperties("cpRemarks") = txtRemarks.Text
                ButtonPartNo.JSProperties("cpPOKanban") = txtPOKanban.Text
            Else
                ButtonPartNo.JSProperties("cpDelivery") = ""
                ButtonPartNo.JSProperties("cpCommercial") = ""
                ButtonPartNo.JSProperties("cpShip") = ""
                ButtonPartNo.JSProperties("cpRemarks") = ""
                Session("SupplierID") = ""
                ButtonPartNo.JSProperties("cpPOKanban") = ""
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' AffiliateName, '' PartNo, '' PartNo1, '' PartName, '' KanbanCls, '' UnitDesc, '' MOQ, '' QtyBox, '' Maker, " & vbCrLf & _
                  " 0 POQty, 0 POQtyOld, 0 ForecastN1, 0 ForecastN2, 0 ForecastN3,   " & vbCrLf & _
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

    Private Sub uf_Approve(ByVal pKanbanCls As String)
        Dim ls_sql As String        
        Dim x As Integer, i As Integer, j As Integer, k As Integer, m As Integer
        Dim ls_Time(4) As String
        Dim jlhHari As Integer = 0

        'Check Jumlah Hari
        If Month(dtPeriodFrom.Value) = "1" Or Month(dtPeriodFrom.Value) = "3" Or Month(dtPeriodFrom.Value) = "5" _
            Or Month(dtPeriodFrom.Value) = "7" Or Month(dtPeriodFrom.Value) = "8" Or Month(dtPeriodFrom.Value) = "10" _
            Or Month(dtPeriodFrom.Value) = "12" Then
            jlhHari = 31
        End If

        If Month(dtPeriodFrom.Value) = "4" Or Month(dtPeriodFrom.Value) = "6" Or Month(dtPeriodFrom.Value) = "9" _
            Or Month(dtPeriodFrom.Value) = "11" Then
            jlhHari = 31
        End If

        If Month(dtPeriodFrom.Value) = "2" Then
            If Year(dtPeriodFrom.Value) Mod 4 = 0 Then
                jlhHari = 29
            Else
                jlhHari = 28
            End If
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PORev_Master set FinalApproveDate = getdate(), FinalApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & cboPartNo.Text & "' and PORevNo = '" & cboPartNoRev.Text & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                Dim sqlCommNew As SqlCommand = sqlConn.CreateCommand
                sqlCommNew.Connection = sqlConn
                sqlCommNew.Transaction = sqlTran

                'Get KanbanTime
                'ls_sql = "select KanbanCycle, KanbanTime from MS_KanbanTime where AffiliateID = '" & Session("AffiliateID") & "'"

                'sqlCommNew.CommandText = ls_sql
                'Dim da As New SqlDataAdapter(sqlCommNew)
                'Dim ds1 As New DataSet
                'da.Fill(ds1)


                'If ds1.Tables(0).Rows.Count > 0 Then
                '    For i = 0 To ds1.Tables(0).Rows.Count - 1
                '        If ds1.Tables(0).Rows(i)("KanbanCycle") = "1" Then
                '            ls_Time(0) = ds1.Tables(0).Rows(i)("KanbanTime")
                '        ElseIf ds1.Tables(0).Rows(i)("KanbanCycle") = "2" Then
                '            ls_Time(1) = ds1.Tables(0).Rows(i)("KanbanTime")
                '        ElseIf ds1.Tables(0).Rows(i)("KanbanCycle") = "3" Then
                '            ls_Time(2) = ds1.Tables(0).Rows(i)("KanbanTime")
                '        ElseIf ds1.Tables(0).Rows(i)("KanbanCycle") = "4" Then
                '            ls_Time(3) = ds1.Tables(0).Rows(i)("KanbanTime")
                '        End If
                '    Next
                'Else
                '    ls_Time(0) = "00:00:00"
                '    ls_Time(1) = "00:00:00"
                '    ls_Time(2) = "00:00:00"
                '    ls_Time(3) = "00:00:00"
                'End If

                'da.Dispose()
                'ds1.Dispose()

                For k = 1 To jlhHari
                    Dim ls_KanbanNo As String = Format(dtPeriodFrom.Value, "yyyyMM") & IIf(k.ToString.Length = 1, "0" & k, k)
                    Dim ls_Date As String = Format(dtPeriodFrom.Value, "yyyy-MM-") & IIf(k.ToString.Length = 1, "0" & k, k)

                    'Create Kanban
                    ls_sql = " select  " & vbCrLf & _
                                  " 	a.AffiliateID, a.PONo, a.SupplierID, c.Period,  " & vbCrLf & _
                                  " 	e.DeliveryLocationCode, b.PartNo, d.UnitCls, d.QtyBox, " & vbCrLf & _
                                  " 	[DeliveryD" & k & "], " & vbCrLf & _
                                  " 	colcycle1 =  CASE WHEN (CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)) > 0  " & vbCrLf & _
                                  "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)  " & vbCrLf & _
                                  "                          ELSE 0 END) = 0 THEN DeliveryD" & k & " ELSE  " & vbCrLf & _
                                  "                          (CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)) > 0  " & vbCrLf & _
                                  "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)  " & vbCrLf & _
                                  "                          ELSE 0 END) END ,  " & vbCrLf & _
                                  "  	colcycle2 =  CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)) > 0  " & vbCrLf & _
                                  "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)  " & vbCrLf & _
                                  "                          ELSE 0 END ,  " & vbCrLf & _
                                  "  	colcycle3 = CASE WHEN (DeliveryD" & k & " - CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)) > 0  " & vbCrLf & _
                                  "                          THEN CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)  "

                    ls_sql = ls_sql + "                          ELSE 0 END ,  " & vbCrLf & _
                                      "  	colcycle4 = CASE WHEN (DeliveryD" & k & " - (CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)) * 3) > 0  " & vbCrLf & _
                                      "                          THEN DeliveryD" & k & " - ((CEILING(FLOOR(DeliveryD" & k & "/4) / isnull(d.QtyBox,0)) * isnull(d.QtyBox,0)) )*3  " & vbCrLf & _
                                      "                          ELSE 0 END " & vbCrLf & _
                                      " from PORev_MasterUpload a " & vbCrLf & _
                                      " inner join PORev_DetailUpload b on a.AffiliateID = b.AffiliateID and a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.SeqNo = b.SeqNo and a.PORevNo = b.PORevNo " & vbCrLf & _
                                      " inner join PO_Master c on a.AffiliateID = c.AffiliateID and a.PONo = c.PONo and a.SupplierID = c.SupplierID " & vbCrLf & _
                                      " left join MS_Parts d on b.PartNo = d.PartNo " & vbCrLf & _
                                      " left join MS_DeliveryPlace e on e.AffiliateID = a.AffiliateID and e.DefaultCls = '1' " & vbCrLf & _
                                      " where a.AffiliateID = '" & Session("AffiliateID") & "' and a.PONo = '" & cboPartNo.Text & "' and b.KanbanCls = 0 and DeliveryD" & k & " > 0 and a.PORevNo = '" & cboPartNoRev.Text & "'"

                    sqlCommNew.CommandText = ls_sql
                    Dim da As New SqlDataAdapter(sqlCommNew)
                    Dim ds As New DataSet
                    da.Fill(ds)

                    If ds.Tables(0).Rows.Count > 0 Then
                        'Insert Master
                        'For j = 0 To 3
                        ls_sql = " INSERT INTO [dbo].[Kanban_Master] " & vbCrLf & _
                                  "            ([KanbanNo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[KanbanCycle] " & vbCrLf & _
                                  "            ,[KanbanDate] " & vbCrLf & _
                                  "            ,[KanbanTime] " & vbCrLf & _
                                  "            ,[KanbanStatus] " & vbCrLf & _
                                  "            ,[AffiliateApproveUser] " & vbCrLf & _
                                  "            ,[AffiliateApproveDate] " & vbCrLf & _
                                  "            ,[SupplierApproveUser] "

                        ls_sql = ls_sql + "            ,[SupplierApproveDate] " & vbCrLf & _
                                          "            ,[EntryDate] " & vbCrLf & _
                                          "            ,[EntryUser] " & vbCrLf & _
                                          "            ,[DeliveryLocationCode] " & vbCrLf & _
                                          "            ,[excelcls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & ls_KanbanNo & "-1' " & vbCrLf & _
                                          "            ,'" & Session("AffiliateID") & "' "

                        ls_sql = ls_sql + "            ,'" & ds.Tables(0).Rows(0)("SupplierID") & "'" & vbCrLf & _
                                          "            ,'1'" & vbCrLf & _
                                          "            ,'" & ls_Date & "'" & vbCrLf & _
                                          "            ,'00:00:00' " & vbCrLf & _
                                          "            ,'1' " & vbCrLf & _
                                          "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("AffiliateID") & "'"

                        ls_sql = ls_sql + "            ,'" & ds.Tables(0).Rows(0)("DeliveryLocationCode") & "'" & vbCrLf & _
                                          "            ,'1')"
                        SqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        SqlComm.ExecuteNonQuery()
                        ' Next

                        'Insert Detail
                        For i = 0 To ds.Tables(0).Rows.Count - 1
                            'For j = 0 To 3
                            Dim ls_QtyBox As Integer = ds.Tables(0).Rows(i)("QtyBox")
                            Dim ls_Cycle As Integer = ds.Tables(0).Rows(i)("DeliveryD" & k)
                            Dim ls_Ulang As Integer = ls_Cycle / ls_QtyBox

                            ls_sql = " INSERT INTO [dbo].[Kanban_Detail] " & vbCrLf & _
                                      "            ([KanbanNo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[PONo] " & vbCrLf & _
                                      "            ,[DeliveryLocationCode] " & vbCrLf & _
                                      "            ,[UnitCls] " & vbCrLf & _
                                      "            ,[KanbanQty]) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & ls_KanbanNo & "-1' "

                            ls_sql = ls_sql + "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                              "            ,'" & ds.Tables(0).Rows(i)("SupplierID") & "' " & vbCrLf & _
                                              "            ,'" & ds.Tables(0).Rows(i)("PartNo") & "' " & vbCrLf & _
                                              "            ,'" & ds.Tables(0).Rows(i)("PONo") & "' " & vbCrLf & _
                                              "            ,'" & ds.Tables(0).Rows(i)("DeliveryLocationCode") & "' " & vbCrLf & _
                                              "            ,'" & ds.Tables(0).Rows(i)("UnitCls") & "' " & vbCrLf & _
                                              "            ,'" & ls_Cycle & "') "
                            SqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            SqlComm.ExecuteNonQuery()

                            For m = 0 To ls_Ulang - 1
                                Dim ls_Barcode As String = ls_KanbanNo.Trim & "-1" & ds.Tables(0).Rows(i)("SupplierID").ToString.Trim & (m + 1) & Session("AffiliateID").ToString.Trim & ds.Tables(0).Rows(i)("PartNo").ToString.Trim

                                ls_sql = " INSERT INTO [dbo].[Kanban_Barcode] " & vbCrLf & _
                                          "            ([Barcode] " & vbCrLf & _
                                          "            ,[PoNo] " & vbCrLf & _
                                          "            ,[KanbanNo] " & vbCrLf & _
                                          "            ,[Seqno] " & vbCrLf & _
                                          "            ,[AffiliateID] " & vbCrLf & _
                                          "            ,[SupplierID] " & vbCrLf & _
                                          "            ,[DeliveryLocationCode] " & vbCrLf & _
                                          "            ,[Partno] " & vbCrLf & _
                                          "            ,[Qty]) " & vbCrLf & _
                                          "      VALUES "

                                ls_sql = ls_sql + "            ('" & ls_Barcode & "' " & vbCrLf & _
                                                  "            ,'" & ds.Tables(0).Rows(i)("PONo") & "' " & vbCrLf & _
                                                  "            ,'" & ls_KanbanNo & "-1' " & vbCrLf & _
                                                  "            ,'" & m + 1 & "' " & vbCrLf & _
                                                  "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                                  "            ,'" & ds.Tables(0).Rows(i)("SupplierID") & "' " & vbCrLf & _
                                                  "            ,'" & ds.Tables(0).Rows(i)("DeliveryLocationCode") & "' " & vbCrLf & _
                                                  "            ,'" & ds.Tables(0).Rows(i)("PartNo") & "' " & vbCrLf & _
                                                  "            ,'" & ls_QtyBox & "') "

                                SqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                SqlComm.ExecuteNonQuery()
                            Next
                            'Next
                        Next
                    End If
                Next

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_FillCombo(ByVal pPeriod As String)
        Dim ls_SQL As String = ""

        ls_SQL = "select distinct RTRIM(PONo) PONo from PORev_Master where AffiliateID = '" & Session("AffiliateID") & "' and Year(Period) = '" & Year(pPeriod) & "' and month(Period) = '" & Month(pPeriod) & "' order by PONo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PONo")
                .Columns(0).Width = 180

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_FillComboRev(ByVal pPeriod As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""

        ls_SQL = "select RTRIM(PORevNo) PORevNo from PORev_Master where AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & pPONO & "' order by PORevNo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNoRev
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PORevNo")
                .Columns(0).Width = 180

                .TextField = "PORevNo"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub sendEmail()
        Dim receiptEmail As String = ""
        Dim receiptCCEmail As String = ""
        Dim fromEmail As String = ""
        Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
        Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
        Dim ls_Body As String = ""

        '*******File di Server
        Dim dsNotification As New DataSet
        dsNotification = GetNotification("14")

        If dsNotification.Tables(0).Rows.Count > 0 Then
            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line1")) Then
                ls_Line1 = ""
            Else
                ls_Line1 = dsNotification.Tables(0).Rows(0)("Line1")
            End If
            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line2")) Then
                ls_Line2 = ""
            Else
                ls_Line2 = dsNotification.Tables(0).Rows(0)("Line2")
            End If
            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line3")) Then
                ls_Line3 = ""
            Else
                ls_Line3 = dsNotification.Tables(0).Rows(0)("Line3")
            End If

            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line4")) Then
                ls_Line4 = ""
            Else
                ls_Line4 = dsNotification.Tables(0).Rows(0)("Line4")
            End If

            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line5")) Then
                ls_Line5 = ""
            Else
                ls_Line5 = dsNotification.Tables(0).Rows(0)("Line5")
            End If

            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line6")) Then
                ls_Line6 = ""
            Else
                ls_Line6 = dsNotification.Tables(0).Rows(0)("Line6")
            End If

            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line7")) Then
                ls_Line7 = ""
            Else
                ls_Line7 = dsNotification.Tables(0).Rows(0)("Line7")
            End If

            If IsDBNull(dsNotification.Tables(0).Rows(0)("Line8")) Then
                ls_Line8 = ""
            Else
                ls_Line8 = dsNotification.Tables(0).Rows(0)("Line8")
            End If
        End If

        ls_Body = ls_Line1 & vbCr & ls_Line2 & "PO No:" & cboPartNo.Text & vbCr & vbCr & "PO Revision No: " & cboPartNoRev.Text & vbCr & vbCr & ls_Line3 & vbCr & ls_Line4 & ls_Line5 & vbCr & ls_Line6 & vbCr & ls_Line7 & vbCr & ls_Line8

        Dim dsEmail As New DataSet
        dsEmail = EmailToEmailCC(Session("AffiliateID"), "PASI", "")
        '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
        For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
            If receiptCCEmail = "" Then
                receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
            Else
                receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionCC")
            End If
            If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
            End If
            If receiptEmail = "" Then
                receiptEmail = dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTO")
            Else
                receiptEmail = receiptEmail & "," & dsEmail.Tables(0).Rows(iRow)("AffiliatePORevisionTO")
            End If
        Next
        receiptCCEmail = Replace(receiptCCEmail, ",", ";")
        receiptEmail = Replace(receiptEmail, ",", ";")

        'If receiptCCEmail <> "" Then
        '    receiptCCEmail = Left(receiptCCEmail, receiptCCEmail.Length - 1)
        'End If


        If receiptEmail = "" Then
            MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
            Exit Sub
        End If

        'Make a copy of the file/Open it/Mail it/Delete it
        'If you want to change the file name then change only TempFileName


        'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
        Dim mailMessage As New Mail.MailMessage()
        mailMessage.From = New MailAddress(fromEmail)
        mailMessage.Subject = "Final Approval PO Revision No: " & cboPartNoRev.Text & " from PO No:" & cboPartNo.Text

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
        'Dim filename As String = TempFilePath & TempFileName
        'mailMessage.Attachments.Add(New Attachment(filename))
        mailMessage.IsBodyHtml = False
        Dim smtp As New SmtpClient
        'smtp.Host = "smtp.atisicloud.com"
        'smtp.Host = "mail.fast.net.id"
        'smtp.EnableSsl = False
        'smtp.UseDefaultCredentials = True
        'smtp.Port = 25
        'smtp.Send(mailMessage)

        smtp.Host = SmtpClient
        If smtp.UseDefaultCredentials = True Then
            smtp.EnableSsl = True
        Else
            smtp.EnableSsl = False
            Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
            smtp.Credentials = myCredential
        End If

        smtp.Port = portClient
        smtp.Send(mailMessage)

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
                     " select 'AFF' flag,AffiliatePORevisionCC, AffiliatePORevisionTO='',FromEmail = AffiliatePORevisionTO from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,AffiliatePORevisionCC, AffiliatePORevisionTO, FromEmail = ''  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

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