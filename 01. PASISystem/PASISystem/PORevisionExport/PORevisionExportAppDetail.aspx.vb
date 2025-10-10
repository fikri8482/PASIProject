Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports System.Net
Imports System.Net.Mail

Public Class PORevisionExportAppDetail
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
    Dim pub_Period As Date
    Dim pub_PORev As String
    Dim pub_PO As String
    Dim pub_Commercial As String
    Dim pub_AffiliateID As String
    Dim pub_AffiliateName As String
    Dim pub_SupplierID As String
    Dim pub_SupplierName As String
    Dim pub_Remarks As String
    Dim pub_Kanban As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "B02"
    Dim pSearch As Boolean = False
    Dim remain As Double
    Dim TotQtyAff As Double
    Dim TotQtyPASI As Double
    Dim ls_POqty As Double

    Dim errorBatch As Boolean
    Dim FlagGrid As Integer
    Dim UpdateSend As Boolean = False

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim DeliveryBy As String
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
                flag = False
            Else
                flag = True
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Session("M01Url") <> "" Then
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "AFFILIATE ORDER REV. APPROVAL DETAIL"
                        'up_Fillcombo()
                        pub_Period = Request.QueryString("t1")
                        pub_PORev = Request.QueryString("t2")
                        pub_PO = Request.QueryString("t3")
                        pub_Commercial = Request.QueryString("t4")
                        pub_AffiliateID = Request.QueryString("t5")
                        pub_AffiliateName = Request.QueryString("t6")
                        pub_SupplierID = Request.QueryString("t7")
                        pub_SupplierName = Request.QueryString("t8")
                        pub_Kanban = Request.QueryString("t9")
                        pub_Remarks = Request.QueryString("t10")
                        'tabIndex()
                        pSearch = False
                        bindDataHeader(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        bindDataDetail(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        'Call SaveDataMaster(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text))
                        'Call SaveDataDetail(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text))
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        pSearch = True
                        
                        'btnClear.Visible = False
                        'ScriptManager.RegisterStartupScript(AffiliateSubmit, AffiliateSubmit.GetType(), "scriptKey", "txtAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');", True)
                    ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                        Session("MenuDesc") = "AFFILIATE ORDER REV. APPROVAL DETAIL"
                        'up_Fillcombo()
                        pub_Period = clsNotification.DecryptURL(Request.QueryString("t1"))
                        pub_PORev = clsNotification.DecryptURL(Request.QueryString("t2"))
                        pub_PO = clsNotification.DecryptURL(Request.QueryString("id2"))
                        pub_Commercial = clsNotification.DecryptURL(Request.QueryString("t4"))
                        pub_AffiliateID = clsNotification.DecryptURL(Request.QueryString("t5"))
                        pub_AffiliateName = clsNotification.DecryptURL(Request.QueryString("t6"))
                        pub_SupplierID = clsNotification.DecryptURL(Request.QueryString("t7"))
                        'pub_SupplierName = clsNotification.DecryptURL(Request.QueryString("t8"))
                        pub_Kanban = clsNotification.DecryptURL(Request.QueryString("t9"))
                        'pub_Remarks = clsNotification.DecryptURL(Request.QueryString("t10"))
                        'tabIndex()
                        pSearch = False
                        bindDataHeader(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        bindDataDetail(pub_Period, pub_PORev, pub_PO, pub_AffiliateID, pub_SupplierID, pub_Kanban)
                        'Call SaveDataMaster(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text))
                        'Call SaveDataDetail(ValidasiInput(pub_AffiliateID), pub_Period, pub_PORev, pub_PO, Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), pub_Commercial, Trim(rblPOKanban.Value), Trim(txtShipBy.Text))
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        pSearch = True
                       
                        'btnClear.Visible = False
                        'ScriptManager.RegisterStartupScript(AffiliateSubmit, AffiliateSubmit.GetType(), "scriptKey", "txtAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');", True)
                    Else
                        Session("MenuDesc") = "AFFILIATE ORDER REV. APPROVAL DETAIL"
                        'tabIndex()
                        'clear()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        btnClear.Visible = True
                    End If
                Else
                    btnClear.Visible = True
                    'txtAffiliateID.Focus()
                    'tabIndex()
                    'clear()
                End If
            End If

            'If ls_AllowDelete = False Then btnDelete.Enabled = False
            'If ls_AllowUpdate = False Then btnSubmit.Enabled = False

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 16, False, clsAppearance.PagerMode.ShowAllRecord)

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    'grid.PerformCallback('load' + '|' + pDate + '|' + pPORevNo + '|' + pPONo + '|' + pAffCode + '|' + pSuppCode + '|' + pKanban);
                    Dim pDate As Date = Split(e.Parameters, "|")(1)
                    Dim pPORevNo As String = Split(e.Parameters, "|")(2)
                    Dim pPONo As String = Split(e.Parameters, "|")(3)
                    Dim pAffCode As String = Split(e.Parameters, "|")(4)
                    Dim pSuppCode As String = Split(e.Parameters, "|")(5)
                    Dim pKanban As String = Split(e.Parameters, "|")(6)

                    bindDataHeader(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
                    bindDataDetail(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)

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
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 16, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                If e.GetValue("AffiliateName") = "REV. BY AFFILIATE" Then
                    e.Cell.BackColor = Color.AliceBlue
                End If

                If e.GetValue("AffiliateName") = "REV. BY PASI" Then
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

                If e.GetValue("AffiliateName") = "REV. BY SUPPLIER" Then
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

    'Private Sub ButtonApprove_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
    '    Dim ls_MsgID As String = ""
    '    If getApp(Trim(txtPORev.Text), Trim(txtPONo.Text)) = True Then
    '        ls_MsgID = "6030"
    '        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
    '        Session("ZZ010Msg") = lblInfo.Text
    '        Exit Sub
    '    End If
    '    uf_Approve()
    '    UpdateSend = True
    '    bindDataHeader(Trim(txtPeriod.Text), Trim(txtPORev.Text), Trim(txtPONo.Text), Trim(txtAffiliateID.Text), Trim(txtSupplierCode.Text), rblPOKanban.Value)
    '    'sendEmail()
    '    sendEmailtoAffiliate()
    '    sendEmailccPASI()
    '    UpdateSend = False
    '    'bindDataDetail(pDate, pPORevNo, pPONo, pAffCode, pSuppCode, pKanban)
    'End Sub

    Private Sub btnSubMenu_Click(sender As Object, e As System.EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            'Session.Remove("M01Url")
            Response.Redirect("~/AffiliateRevision/AffiliateOrderRevAppList.aspx")
        Else
            'Session.Remove("M01Url")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindDataHeader(ByVal pDate As Date, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pAffCode As String, ByVal pSupplierID As String, ByVal pKanban As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "   SELECT DISTINCT PORM.Period,PORD.AffiliateID,AffiliateName,PORM.PORevNo,PORM.PONo  " & vbCrLf & _
                  "   ,CASE WHEN CommercialCls = '0' THEN 'NO' ELSE 'YES' END CommercialCls  " & vbCrLf & _
                  "   ,PORD.SupplierID,SupplierName,ShipCls   " & vbCrLf & _
                  "   ,PODeliveryBy   " & vbCrLf & _
                  "   ,MP.KanbanCls ,d.Remarks  " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.EntryDate,120)EntryDate,ISNULL(PORM.EntryUser,'')EntryUser --1   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.AffiliateApproveDate,120)AffiliateApproveDate,ISNULL(PORM.AffiliateApproveUser,'')AffiliateApproveUser --2   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.PASISendAffiliateDate,120)PASISendAffiliateDate,ISNULL(PORM.PASISendAffiliateUser,'')PASISendAffiliateUser --3   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.SupplierApproveDate,120)SupplierApproveDate,ISNULL(PORM.SupplierApproveUser,'')SupplierApproveUser --4   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.SupplierApprovePendingDate,120)SupplierApprovePendingDate,ISNULL(PORM.SupplierApprovePendingUser,'')SupplierApprovePendingUser --5   " & vbCrLf & _
                  "   ,CONVERT(DATETIME,PORM.SupplierUnApproveDate,120)SupplierUnApproveDate,ISNULL(PORM.SupplierUnApproveUser,'')SupplierUnApproveUser --6   " & vbCrLf

            ls_SQL = ls_SQL + "   ,CONVERT(DATETIME,PORM.PASIApproveDate,120)PASIApproveDate,ISNULL(PORM.PASIApproveUser ,'')PASIApproveUser --7   " & vbCrLf & _
                              "   ,CONVERT(DATETIME,PORM.FinalApproveDate,120)FinalApproveDate,ISNULL(PORM.FinalApproveUser,'')FinalApproveUser --8    " & vbCrLf & _
                              "   ,PORD.PartNo,PartName,CASE WHEN MP.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,UnitCls,MOQ,QtyBox,CASE WHEN MP.KanbanCls = '0' THEN 'NO' ELSE 'YES' END KanbanCls,UnitCls,MOQ,QtyBox   " & vbCrLf & _
                              "   ,PORM.CurrCls,PORD.Price,PORM.Amount,PORd.CurrCls,PORd.Amount   " & vbCrLf & _
                              "   ,ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & pDate & "'))),0) " & vbCrLf & _
                              "   ,ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & pDate & "'))),0) " & vbCrLf & _
                              "   ,ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = MSC.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & pDate & "'))),0) " & vbCrLf & _
                              "   ,PORD.DeliveryD1,PORD.DeliveryD2,PORD.DeliveryD3,PORD.DeliveryD4,PORD.DeliveryD5,PORD.DeliveryD6,PORD.DeliveryD7,PORD.DeliveryD8,PORD.DeliveryD9,PORD.DeliveryD10   " & vbCrLf & _
                              "   ,PORD.DeliveryD11,PORD.DeliveryD12,PORD.DeliveryD13,PORD.DeliveryD14,PORD.DeliveryD15,PORD.DeliveryD16,PORD.DeliveryD17,PORD.DeliveryD18,PORD.DeliveryD19,PORD.DeliveryD20   " & vbCrLf & _
                              "   ,PORD.DeliveryD21,PORD.DeliveryD22,PORD.DeliveryD23,PORD.DeliveryD24,PORD.DeliveryD25,PORD.DeliveryD26,PORD.DeliveryD27,PORD.DeliveryD28,PORD.DeliveryD29,PORD.DeliveryD30,PORD.DeliveryD31   " & vbCrLf & _
                              "   FROM dbo.PORev_Master PORM    " & vbCrLf

            ls_SQL = ls_SQL + "   LEFT JOIN dbo.PORev_Detail PORD ON PORM.PORevNo = PORD.PORevNo AND PORM.PONo = PORD.PONo AND PORM.AffiliateID = PORD.AffiliateID AND PORM.SupplierID = PORD.SupplierID " & vbCrLf & _
                              "   LEFT JOIN PO_Master POM ON PORD.AffiliateID = POM.AffiliateID AND PORD.PONo = POM.PONo AND PORD.SupplierID = POM.SupplierID  " & vbCrLf & _
                              "   LEFT JOIN dbo.PO_Detail POD ON PORM.PONo = POD.PONo AND PORM.AffiliateID = POD.AffiliateID AND PORM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_Affiliate MA ON PORD.AffiliateID = MA.AffiliateID  " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_Parts MP ON PORD.PartNo = MP.PartNo   " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_Supplier MS ON PORD.SupplierID = MS.SupplierID   " & vbCrLf & _
                              "   LEFT JOIN dbo.MS_SupplierCapacity MSC ON PORD.PartNo = MSC.PartNo AND MP.PartNo = MSC.PartNo AND PORD.SupplierID=MSC.SupplierID  " & vbCrLf & _
                              "   LEFT JOIN dbo.PO_MasterUpload d on d.PONo = PORD.PONo and d.AffiliateID = PORD.AffiliateID  " & vbCrLf

            ls_SQL = ls_SQL + " WHERE MONTH(PORM.Period) = MONTH('" & pDate & "') AND YEAR(PORM.Period) = YEAR('" & pDate & "')" & vbCrLf & _
                              " AND PORM.PORevNo = '" & pPORevNo & "' AND PORM.PONo='" & pPONo & "' AND PORM.AffiliateID='" & pAffCode & "' AND PORM.SupplierID='" & pSupplierID & "'   " & vbCrLf

            If pSearch = True Then
                If pKanban <> "2" Then
                    ls_SQL = ls_SQL + "   AND POD.KanbanCls='" & pKanban & "'  " & vbCrLf
                End If
            End If


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                dtPeriod.Value = Format(ds.Tables(0).Rows(0)("Period"), "MMM yyyy")
                cboAffiliateCode.Text = ds.Tables(0).Rows(0)("AffiliateID")
                txtAffiliateName.Text = ds.Tables(0).Rows(0)("AffiliateName")
                '.Text = ds.Tables(0).Rows(0)("PORevNo")
                'txtPONo.Text = ds.Tables(0).Rows(0)("PONo")
                txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                txtSupplierCode.Text = ds.Tables(0).Rows(0)("SupplierID")
                txtSupplierName.Text = ds.Tables(0).Rows(0)("SupplierName")

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
            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataDetail(ByVal pDate As Date, ByVal pPORevNo As String, ByVal pPONo As String, ByVal pAff As String, ByVal pSupp As String, ByVal pKanban As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  select tbl2.NoUrut, tbl1.AffiliateName, tbl2.PartNo, tbl1.PartNo1, tbl2.PartName, tbl2.KanbanCls, tbl2.UnitDesc, tbl2.MOQ, tbl2.QtyBox, tbl2.Maker,  " & vbCrLf & _
                  "  	ISNULL(POQty,0)POQty, ISNULL(POQtyOld,0)POQtyOld, CurrDesc, Price, Amount, ForecastN1, ForecastN2, ForecastN3,  " & vbCrLf & _
                  "  	ISNULL(DeliveryD1,0)DeliveryD1, ISNULL(DeliveryD2,0)DeliveryD2, ISNULL(DeliveryD3,0)DeliveryD3, ISNULL(DeliveryD4,0)DeliveryD4, ISNULL(DeliveryD5,0)DeliveryD5,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD6,0)DeliveryD6, ISNULL(DeliveryD7,0)DeliveryD7, ISNULL(DeliveryD8,0)DeliveryD8, ISNULL(DeliveryD9,0)DeliveryD9, ISNULL(DeliveryD10,0)DeliveryD10,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD11,0)DeliveryD11, ISNULL(DeliveryD12,0)DeliveryD12, ISNULL(DeliveryD13,0)DeliveryD13, ISNULL(DeliveryD14,0)DeliveryD14, ISNULL(DeliveryD15,0)DeliveryD15,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD16,0)DeliveryD16, ISNULL(DeliveryD17,0)DeliveryD17, ISNULL(DeliveryD18,0)DeliveryD18, ISNULL(DeliveryD19,0)DeliveryD19, ISNULL(DeliveryD20,0)DeliveryD20,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD21,0)DeliveryD21, ISNULL(DeliveryD22,0)DeliveryD22, ISNULL(DeliveryD23,0)DeliveryD23, ISNULL(DeliveryD24,0)DeliveryD24, ISNULL(DeliveryD25,0)DeliveryD25,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD26,0)DeliveryD26, ISNULL(DeliveryD27,0)DeliveryD27, ISNULL(DeliveryD28,0)DeliveryD28, ISNULL(DeliveryD29,0)DeliveryD29, ISNULL(DeliveryD30,0)DeliveryD30,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD31,0)DeliveryD31,  " & vbCrLf & _
                  "  	ISNULL(DeliveryD1Old,0)DeliveryD1Old, ISNULL(DeliveryD2Old,0)DeliveryD2Old, ISNULL(DeliveryD3Old,0)DeliveryD3Old, ISNULL(DeliveryD4Old,0)DeliveryD4Old, ISNULL(DeliveryD5Old,0)DeliveryD5Old,   " & vbCrLf & _
                  "  	ISNULL(DeliveryD6Old,0)DeliveryD6Old, ISNULL(DeliveryD7Old,0)DeliveryD7Old, ISNULL(DeliveryD8Old,0)DeliveryD8Old, ISNULL(DeliveryD9Old,0)DeliveryD9Old, ISNULL(DeliveryD10Old,0)DeliveryD10Old,   	ISNULL(DeliveryD11Old,0)DeliveryD11Old, ISNULL(DeliveryD12Old,0)DeliveryD12Old, ISNULL(DeliveryD13Old,0)DeliveryD13Old, ISNULL(DeliveryD14Old,0)DeliveryD14Old, ISNULL(DeliveryD15Old,0)DeliveryD15Old,   " & vbCrLf

            ls_SQL = ls_SQL + "  	ISNULL(DeliveryD16Old,0)DeliveryD16Old, ISNULL(DeliveryD17Old,0)DeliveryD17Old, ISNULL(DeliveryD18Old,0)DeliveryD18Old, ISNULL(DeliveryD19Old,0)DeliveryD19Old, ISNULL(DeliveryD20Old,0)DeliveryD20Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD21Old,0)DeliveryD21Old, ISNULL(DeliveryD22Old,0)DeliveryD22Old, ISNULL(DeliveryD23Old,0)DeliveryD23Old, ISNULL(DeliveryD24Old,0)DeliveryD24Old, ISNULL(DeliveryD25Old,0)DeliveryD25Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD26Old,0)DeliveryD26Old, ISNULL(DeliveryD27Old,0)DeliveryD27Old, ISNULL(DeliveryD28Old,0)DeliveryD28Old, ISNULL(DeliveryD29Old,0)DeliveryD29Old, ISNULL(DeliveryD30Old,0)DeliveryD30Old,   " & vbCrLf & _
                              "  	ISNULL(DeliveryD31Old,0)DeliveryD31Old  " & vbCrLf & _
                              "  from   " & vbCrLf & _
                              "  (  " & vbCrLf & _
                              "  	select * from  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select '1' NoUrutDesc, 'REV. BY AFFILIATE' AffiliateName  " & vbCrLf & _
                              "  		union all  		select '2' NoUrutDesc, 'REV. BY PASI' AffiliateName  " & vbCrLf & _
                              "  		union all  " & vbCrLf

            ls_SQL = ls_SQL + "  		select '3' NoUrutDesc, 'REV. BY SUPPLIER' AffiliateName  " & vbCrLf & _
                              "  	)tbla  " & vbCrLf & _
                              "  	cross join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select   " & vbCrLf & _
                              "  			--row_number() over (order by b.PartNo asc) as NoUrut,  " & vbCrLf & _
                              "  			b.PartNo, b.PartNo PartNo1  " & vbCrLf & _
                              "  		from PORev_Master a  " & vbCrLf & _
                              "  		LEFT join PORev_Detail b ON a.PORevNo = b.PORevNo AND  a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              "         LEFT join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  		LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls  " & vbCrLf & _
                              "  		where a.PORevNo='" & pPORevNo & "' AND  a.PONo = '" & pPONo & "' AND a.AffiliateID='" & pAff & "' AND a.SupplierID='" & pSupp & "'  " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND c.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If

            ls_SQL = ls_SQL + "  " & vbCrLf & _
                              "  	)tb1b  " & vbCrLf & _
                              "  )tbl1  " & vbCrLf & _
                              "  left join  " & vbCrLf & _
                              "  (  " & vbCrLf & _
                              "  	select   " & vbCrLf & _
                              "  		convert(char,row_number() over (order by b.PartNo asc))as NoUrut, 'REV. BY AFFILIATE' AffiliateName, '1' NoUrutDesc,   " & vbCrLf & _
                              "  		b.PartNo, b.PartNo PartNo1, c.PartName, case when c.KanbanCls = '1' then 'Yes' else 'No' end KanbanCls, d.Description UnitDesc,  " & vbCrLf & _
                              "  		c.MOQ, c.QtyBox, c.Maker, b.POQty, 0 POQtyOld, e.Description CurrDesc, b.Price, b.Amount," & vbCrLf & _
                              "         ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & pDate & "'))),0),   " & vbCrLf & _
                              "  		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & pDate & "'))),0),   " & vbCrLf & _
                              "  		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & pDate & "'))),0),   " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf & _
                              "   		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf & _
                              "   		0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old,   		0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old,   " & vbCrLf & _
                              "   		0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old,   " & vbCrLf & _
                              "   		0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old,   " & vbCrLf & _
                              "   		0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old,   " & vbCrLf

            ls_SQL = ls_SQL + "   		0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old,   " & vbCrLf & _
                              "   		0 DeliveryD31Old  " & vbCrLf & _
                              "  		from PORev_Master a  " & vbCrLf & _
                              "  	LEFT join PORev_Detail b ON a.PORevNo = b.PORevNo AND a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              "  	LEFT join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	LEFT join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  	LEFT join MS_CurrCls e on e.CurrCls = b.CurrCls  	 " & vbCrLf & _
                              "     where a.PORevNo='" & pPORevNo & "' AND  a.PONo = '" & pPONo & "' AND a.AffiliateID='" & pAff & "' AND a.SupplierID='" & pSupp & "'  " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND c.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If
            ls_SQL = ls_SQL + "   " & vbCrLf & _
                              "  	union all  " & vbCrLf & _
                              "  	select   " & vbCrLf

            ls_SQL = ls_SQL + "  		'' NoUrut,'REV. BY PASI' AffiliateName, '2' NoUrutDesc,   " & vbCrLf & _
                              "  		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, '' Maker, b.POQty, b.POQtyOld, e.Description CurrDesc, b.Price, b.Amount,  " & vbCrLf & _
                              "  		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & pDate & "'))),0),   " & vbCrLf & _
                              "  		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & pDate & "'))),0),   " & vbCrLf & _
                              "  		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & pDate & "'))),0),   " & vbCrLf & _
                              "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,    		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,   " & vbCrLf & _
                              "   		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf & _
                              "   		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   		b.DeliveryD31Old  		from AffiliateRev_Master a  " & vbCrLf & _
                              "  	LEFT join AffiliateRev_Detail b ON a.PORevNo = b.PORevNo AND a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              "  	LEFT join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	LEFT join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  	LEFT join MS_CurrCls e on e.CurrCls = b.CurrCls  " & vbCrLf

            ls_SQL = ls_SQL + "  	where a.PORevNo='" & pPORevNo & "' AND  a.PONo = '" & pPONo & "' AND a.AffiliateID='" & pAff & "' AND a.SupplierID='" & pSupp & "'  " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND c.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If

            ls_SQL = ls_SQL + "  " & vbCrLf & _
                              "  	union all  " & vbCrLf & _
                              "  	select   " & vbCrLf & _
                              "  		'' NoUrut, 'REV. BY SUPPLIER' AffiliateName, '3' NoUrutDesc,   " & vbCrLf & _
                              "  		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, '' Maker, b.POQty, b.POQtyOld, e.Description CurrDesc, b.Price, b.Amount,  " & vbCrLf & _
                              "  		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & pDate & "'))),0),  " & vbCrLf & _
                              "         ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & pDate & "'))),0),   " & vbCrLf & _
                              "  		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo AND MF.AffiliateID = b.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & pDate & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & pDate & "'))),0),   " & vbCrLf & _
                              "   		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5,  " & vbCrLf & _
                              "   		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,   " & vbCrLf & _
                              "   		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,   " & vbCrLf

            ls_SQL = ls_SQL + "   		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,   " & vbCrLf & _
                              "   		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,   " & vbCrLf & _
                              "   		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,   " & vbCrLf & _
                              "   		b.DeliveryD31,  " & vbCrLf & _
                              "   		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old,  " & vbCrLf & _
                              "   		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,    		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,   " & vbCrLf & _
                              "   		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,   " & vbCrLf & _
                              "   		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,   " & vbCrLf & _
                              "   		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,   " & vbCrLf & _
                              "   		b.DeliveryD31Old  " & vbCrLf & _
                              "  		from PORev_MasterUpload a  " & vbCrLf

            ls_SQL = ls_SQL + "  	LEFT join PORev_DetailUpload b ON a.PORevNo = b.PORevNo AND  a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID  " & vbCrLf & _
                              "  	LEFT join MS_Parts c on b.PartNo = c.PartNo  " & vbCrLf & _
                              "  	LEFT join MS_UnitCls d on c.UnitCls = d.UnitCls  " & vbCrLf & _
                              "  	LEFT join MS_CurrCls e on e.CurrCls = b.CurrCls  " & vbCrLf & _
                              "  	where a.PORevNo='" & pPORevNo & "' AND  a.PONo = '" & pPONo & "' AND a.AffiliateID='" & pAff & "' AND a.SupplierID='" & pSupp & "'  " & vbCrLf

            If pSearch = True Then
                'If pKanban <> "2" Then
                '    ls_SQL = ls_SQL + "   AND c.KanbanCls='" & pKanban & "'  " & vbCrLf
                'End If
            End If
            ls_SQL = ls_SQL + "   )tbl2 on tbl2.AffiliateName = tbl1.AffiliateName and tbl1.PartNo = tbl2.PartNo1 and tbl1.NoUrutDesc = tbl2.NoUrutDesc  " & vbCrLf & _
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Dim pDateDay As DateTime = CDate(Format(pDate, "MM") + "/01/" + Format(pDate, "yyyy"))
                Select Case Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, pDateDay)))
                    Case 28
                        grid.Columns("DeliveryD29").Visible = False
                        grid.Columns("DeliveryD30").Visible = False
                        grid.Columns("DeliveryD31").Visible = False

                    Case 29
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = False
                        grid.Columns("DeliveryD31").Visible = False

                    Case 30
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = True
                        grid.Columns("DeliveryD31").Visible = False

                    Case 31
                        grid.Columns("DeliveryD29").Visible = True
                        grid.Columns("DeliveryD30").Visible = True
                        grid.Columns("DeliveryD31").Visible = True
                End Select
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()

        End Using
    End Sub

    'Private Sub up_Fillcombo()
    '    Dim ls_SQL As String = ""
    '    'Combo Affiliate
    '    ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate" & vbCrLf
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        With cboAffiliateCode
    '            .Items.Clear()
    '            .Columns.Clear()
    '            .DataSource = ds.Tables(0)
    '            .Columns.Add("AffiliateID")
    '            .Columns(0).Width = 50
    '            .Columns.Add("AffiliateName")
    '            .Columns(1).Width = 120

    '            .TextField = "AffiliateID"
    '            .DataBind()
    '            .SelectedIndex = 0
    '            txtAffiliateName.Text = clsGlobal.gs_All
    '        End With

    '        sqlConn.Close()
    '    End Using
    '    'Combo Supplier
    '    ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierCode, '" & clsGlobal.gs_All & "' SupplierName union all select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        With cboSupplierCode
    '            .Items.Clear()
    '            .Columns.Clear()
    '            .DataSource = ds.Tables(0)
    '            .Columns.Add("SupplierCode")
    '            .Columns(0).Width = 50
    '            .Columns.Add("SupplierName")
    '            .Columns(1).Width = 120

    '            .TextField = "SupplierID"
    '            .DataBind()
    '            .SelectedIndex = 0
    '            txtSupplierName.Text = clsGlobal.gs_All
    '        End With

    '        sqlConn.Close()
    '    End Using

    'End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' as AffiliateID, '' AffiliateName, ''Address, ''City, '' PostalCode, ''Phone1, '' Phone2, ''Fax, ''NPWP, ''PODeliveryBy, ''DetailPage"

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

    'Private Sub uf_Approve()
    '    Dim ls_sql As String
    '    Dim x As Integer

    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
    '            ls_sql = " Update PORev_Master set PASIApproveDate = getdate(), PASIApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
    '                        " WHERE PORevNo='" & Trim(txtPORev.Text) & "' AND PONo='" & Trim(txtPONo.Text) & "' AND AffiliateID='" & Trim(txtAffiliateID.Text) & "' AND SupplierID='" & Trim(txtSupplierCode.Text) & "' " & vbCrLf

    '            Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
    '            x = SqlComm.ExecuteNonQuery()

    '            SqlComm.Dispose()
    '            sqlTran.Commit()
    '        End Using
    '        sqlConn.Close()
    '    End Using
    'End Sub

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

    'Private Sub GetDelivery()
    '    Dim ls_SQL As String = ""
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()
    '        ls_SQL = " SELECT DeliveryByPASICls FROM PORev_Master  PORM " & vbCrLf & _
    '              " LEFT JOIN dbo.PORev_MasterUpload POMU ON PORM.PONo=POMU.PONo AND PORM.AffiliateID=POMU.AffiliateID AND PORM.SupplierID=POMU.SupplierID  " & vbCrLf & _
    '              " LEFT JOIN PO_Master POM ON POM.PONo=PORM.PONo AND POM.AffiliateID=PORM.AffiliateID AND POM.SupplierID=PORM.SupplierID  " & vbCrLf & _
    '              " WHERE PORM.PONo='" & txtPONo.Text.Trim & "' AND PORM.PORevNo='" & txtPORev.Text.Trim & "' "

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)
    '        If ds.Tables(0).Rows.Count > 0 Then
    '            DeliveryBy = If(IsDBNull(ds.Tables(0).Rows(0)("DeliveryByPASICls")), "", ds.Tables(0).Rows(0)("DeliveryByPASICls"))
    '        End If
    '    End Using
    'End Sub

    Private Function EmailToEmailCC(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                     " select 'AFF' flag,affiliatepocc, affiliatepoto,FromEmail = '' from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,affiliatepocc,affiliatepoto='',FromEmail = affiliatepoto  from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    'Private Sub sendEmail()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
    '    Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
    '    Dim ls_Body As String = ""
    '    Dim pApproval As String
    '    If txtPASIAppDate.Text <> "" Then
    '        pApproval = "1"
    '    Else
    '        pApproval = "0"
    '    End If

    '    Call GetDelivery()
    '    '*******File di Server
    '    '"PORevEntry.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>&t5=<%#GetPORevNo(Container)%>&Session=~/PurchaseOrderRevision/PORevList.aspx"
    '    'Dim ls_URl As String = "http://" & clsNotification.pub_ServerName & "/PurchaseOrderRevision/PORevEntry.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(txtAffiliateID.Text) & _
    '    '                              "&t2=" & clsNotification.EncryptURL(txtAffiliateName.Text) & "&t3=" & clsNotification.EncryptURL(txtPeriod.Text) & "&t4=" & clsNotification.EncryptURL(txtSupplierCode.Text) & "&t5=" & clsNotification.EncryptURL(txtPORev.Text) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevList.aspx")

    '    'ls_Body = clsNotification.GetNotification("23", ls_URl, txtPONo.Text.Trim, "", "", txtPORev.Text.Trim)

    '    '"PORevFinalApproval.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>
    '    '&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>&t5=<%#GetRemarks(Container)%>
    '    '&t6=<%#GetFinalApproval(Container)%>&t7=<%#GetDeliveryBy(Container)%>&t8=<%#GetPORevNo(Container)%>&Session=~/PurchaseOrderRevision/PORevFinalApprovalList.aspx"
    '    Dim ls_URl As String = "http://" & clsNotification.pub_ServerName & "/PurchaseOrderRevision/PORevFinalApproval.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL("") & _
    '                                  "&t2=" & clsNotification.EncryptURL(txtCommercial.Text.Trim) & "&t3=" & clsNotification.EncryptURL(txtPeriod.Text.Trim) & "&t4=" & clsNotification.EncryptURL(txtSupplierCode.Text.Trim) & _
    '                                  "&t5=" & clsNotification.EncryptURL(txtRemarks.Text.Trim) & "&t6=" & clsNotification.EncryptURL(pApproval) & _
    '                                  "&t7=" & clsNotification.EncryptURL(DeliveryBy) & "&t8=" & clsNotification.EncryptURL(txtPORev.Text.Trim) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevFinalApprovalList.aspx")

    '    ls_Body = clsNotification.GetNotification("23", ls_URl, txtPONo.Text.Trim, "", "", txtPORev.Text.Trim)

    '    'ls_Body = ls_Line1 & vbCr & ls_Line2 & "PO Revision No:" & Trim(txtPORev.Text) & vbCr & vbCr & ls_Line3 & vbCr & ls_Line4 & ls_Line5 & vbCr & ls_Line6 & vbCr & ls_Line7 & vbCr & ls_Line8

    '    Dim dsEmail As New DataSet
    '    dsEmail = EmailToEmailCC(Trim(txtAffiliateID.Text), "PASI", "")
    '    '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '    For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '        If receiptCCEmail = "" Then
    '            receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '        Else
    '            receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '        End If
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
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

    '    'receiptCCEmail = "kristriyana@tos.co.id"
    '    'receiptEmail = "kristriyana@tos.co.id"
    '    If receiptEmail = "" Then
    '        MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '        Exit Sub
    '    End If

    '    'Make a copy of the file/Open it/Mail it/Delete it
    '    'If you want to change the file name then change only TempFileName


    '    'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
    '    Dim mailMessage As New Mail.MailMessage()
    '    mailMessage.From = New MailAddress(fromEmail)
    '    mailMessage.Subject = "Approval PO Revision No: " & Trim(txtPORev.Text)

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

    Private Function getApp(ByVal pPORevNo As String, ByVal pPONo As String) As Boolean
        Dim ls_SQL As String = ""
        Dim doneApp As Boolean = False
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " SELECT * FROM PORev_Master " & vbCrLf & _
                  " WHERE PORevNo='" & pPORevNo & "' AND PONO='" & pPONo & "' AND  " & vbCrLf & _
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

    'Private Sub sendEmailtoAffiliate()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
    '    Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
    '    Dim ls_Body As String = ""
    '    Dim pApproval As String
    '    If txtPASIAppDate.Text <> "" Then
    '        pApproval = "1"
    '    Else
    '        pApproval = "0"
    '    End If

    '    Call GetDelivery()
    '    '*******File di Server
    '    '"PORevEntry.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>&t5=<%#GetPORevNo(Container)%>&Session=~/PurchaseOrderRevision/PORevList.aspx"
    '    'Dim ls_URl As String = "http://" & clsNotification.pub_ServerName & "/PurchaseOrderRevision/PORevEntry.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(txtAffiliateID.Text) & _
    '    '                              "&t2=" & clsNotification.EncryptURL(txtAffiliateName.Text) & "&t3=" & clsNotification.EncryptURL(txtPeriod.Text) & "&t4=" & clsNotification.EncryptURL(txtSupplierCode.Text) & "&t5=" & clsNotification.EncryptURL(txtPORev.Text) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevList.aspx")

    '    'ls_Body = clsNotification.GetNotification("23", ls_URl, txtPONo.Text.Trim, "", "", txtPORev.Text.Trim)

    '    '"PORevFinalApproval.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>
    '    '&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>&t5=<%#GetRemarks(Container)%>
    '    '&t6=<%#GetFinalApproval(Container)%>&t7=<%#GetDeliveryBy(Container)%>&t8=<%#GetPORevNo(Container)%>&Session=~/PurchaseOrderRevision/PORevFinalApprovalList.aspx"
    '    Dim ls_URl As String = "http://" & clsNotification.pub_ServerNameAffiliate & "/PurchaseOrderRevision/PORevFinalApproval.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t1=" & clsNotification.EncryptURL("") & _
    '                                  "&t2=" & clsNotification.EncryptURL(txtCommercial.Text.Trim) & "&t3=" & clsNotification.EncryptURL(txtPeriod.Text.Trim) & "&t4=" & clsNotification.EncryptURL(txtSupplierCode.Text.Trim) & _
    '                                  "&t5=" & clsNotification.EncryptURL(txtRemarks.Text.Trim) & "&t6=" & clsNotification.EncryptURL(pApproval) & _
    '                                  "&t7=" & clsNotification.EncryptURL(DeliveryBy) & "&t8=" & clsNotification.EncryptURL(txtPORev.Text.Trim) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderRevision/PORevFinalApprovalList.aspx")

    '    ls_Body = clsNotification.GetNotification("23", ls_URl, txtPONo.Text.Trim, "", "", txtPORev.Text.Trim)

    '    'ls_Body = ls_Line1 & vbCr & ls_Line2 & "PO Revision No:" & Trim(txtPORev.Text) & vbCr & vbCr & ls_Line3 & vbCr & ls_Line4 & ls_Line5 & vbCr & ls_Line6 & vbCr & ls_Line7 & vbCr & ls_Line8

    '    Dim dsEmail As New DataSet
    '    dsEmail = EmailToEmailCCNotif(Trim(txtAffiliateID.Text), "PASI", "")
    '    '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '    For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '            fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
    '        End If
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
    '            If receiptEmail = "" Then
    '                receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '            Else
    '                receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
    '            End If
    '        End If
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
    '            If receiptCCEmail = "" Then
    '                receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '            Else
    '                receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '            End If
    '        End If
    '    Next
    '    receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '    receiptEmail = Replace(receiptEmail, ",", ";")

    '    'If receiptCCEmail <> "" Then
    '    '    receiptCCEmail = Left(receiptCCEmail, receiptCCEmail.Length - 1)
    '    'End If

    '    'receiptCCEmail = "kristriyana@tos.co.id"
    '    'receiptEmail = "kristriyana@tos.co.id"
    '    If receiptEmail = "" Then
    '        MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '        Exit Sub
    '    End If

    '    'Make a copy of the file/Open it/Mail it/Delete it
    '    'If you want to change the file name then change only TempFileName


    '    'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
    '    Dim mailMessage As New Mail.MailMessage()
    '    mailMessage.From = New MailAddress(fromEmail)
    '    mailMessage.Subject = "Approval PO Revision No: " & Trim(txtPORev.Text)

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
    '        smtp.EnableSsl = True
    '    Else
    '        smtp.EnableSsl = False
    '        Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '        smtp.Credentials = myCredential
    '    End If

    '    smtp.Port = portClient
    '    smtp.Send(mailMessage)

    'End Sub

    'Private Sub sendEmailccPASI()
    '    Dim receiptEmail As String = ""
    '    Dim receiptCCEmail As String = ""
    '    Dim fromEmail As String = ""
    '    Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
    '    Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
    '    Dim ls_Body As String = ""
    '    Dim pApproval As String
    '    If txtPASIAppDate.Text <> "" Then
    '        pApproval = "1"
    '    Else
    '        pApproval = "0"
    '    End If

    '    Call GetDelivery()

    '    '"AffiliateOrderRevAppDetail.aspx?id=<%#GetRowValue(Container)%>
    '    '&t1=<%#GetPeriod(Container)%>&t2=<%#GetPORevNo(Container)%>
    '    '&t3=<%#GetPONo(Container)%>&t4=<%#GetCommercial(Container)%>
    '    '&t5=<%#GetAffiliateID(Container)%>&t6=<%#GetAffiliateName(Container)%>
    '    '&t7=<%#GetSupplierID(Container)%>&t8=<%#GetSupplierName(Container)%>
    '    '&t9=<%#GetKanban(Container)%>&t10=<%#GetRemarks(Container)%>&Session=~/AffiliateRevision/AffiliateOrderRevAppList.aspx"

    '    Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateRevision/AffiliateOrderRevAppDetail.aspx?id2=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & _
    '                                   "&t1=" & clsNotification.EncryptURL(txtPeriod.Text.Trim) & "&t2=" & clsNotification.EncryptURL(txtPORev.Text.Trim) & _
    '                                   "&t3=" & clsNotification.EncryptURL(txtPONo.Text.Trim) & "&t4=" & clsNotification.EncryptURL(txtCommercial.Text.Trim) & _
    '                                   "&t5=" & clsNotification.EncryptURL(txtAffiliateID.Text.Trim) & "&t6=" & clsNotification.EncryptURL(txtAffiliateName.Text.Trim) & _
    '                                   "&t7=" & clsNotification.EncryptURL(txtSupplierCode.Text.Trim) & "&t8=" & clsNotification.EncryptURL(txtSupplierName.Text.Trim) & _
    '                                   "&t9=" & clsNotification.EncryptURL(rblPOKanban.Value) & "&t10=" & clsNotification.EncryptURL(txtRemarks.Text.Trim) & _
    '                                   "&Session=" & clsNotification.EncryptURL("~/AffiliateRevision/AffiliateOrderRevAppList.aspx")

    '    ls_Body = clsNotification.GetNotification("23", ls_URl, txtPONo.Text.Trim)

    '    Dim dsEmail As New DataSet
    '    dsEmail = EmailToEmailCC(Trim(txtAffiliateID.Text), "PASI", "")
    '    '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
    '    For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
    '        If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
    '            fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
    '            If receiptCCEmail = "" Then
    '                receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '            Else
    '                receiptCCEmail = receiptCCEmail & "," & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
    '            End If
    '        End If
    '        receiptEmail = receiptCCEmail
    '    Next
    '    receiptCCEmail = Replace(receiptCCEmail, ",", ";")
    '    receiptEmail = Replace(receiptEmail, ",", ";")

    '    'If receiptCCEmail <> "" Then
    '    '    receiptCCEmail = Left(receiptCCEmail, receiptCCEmail.Length - 1)
    '    'End If

    '    'receiptCCEmail = "kristriyana@tos.co.id"
    '    'receiptEmail = "kristriyana@tos.co.id"
    '    If receiptEmail = "" Then
    '        MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
    '        Exit Sub
    '    End If

    '    'Make a copy of the file/Open it/Mail it/Delete it
    '    'If you want to change the file name then change only TempFileName


    '    'Dim mailMessage As New Mail.MailMessage(fromEmail, receiptEmail)
    '    Dim mailMessage As New Mail.MailMessage()
    '    mailMessage.From = New MailAddress(fromEmail)
    '    mailMessage.Subject = "Approval PO Revision No: " & Trim(txtPORev.Text)

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
    '        smtp.EnableSsl = True
    '    Else
    '        smtp.EnableSsl = False
    '        Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
    '        smtp.Credentials = myCredential
    '    End If

    '    smtp.Port = portClient
    '    smtp.Send(mailMessage)

    'End Sub

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
#End Region

End Class