Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports OfficeOpenXml
Imports Microsoft.Office.Interop
Imports System.Net
Imports System.Net.Mail
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO

Public Class AffiliateOrderExportDetail
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
    Dim pub_Period As Date
    Dim pub_PO As String
    Dim pub_AffiliateID As String
    Dim pub_AffiliateName As String
    Dim pub_Del As String
    Dim pub_Kanban As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "J02"
#End Region

#Region "CONTROL EVENTS"
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
                        Session("MenuDesc") = "AFFILIATE ORDER DETAIL ENTRY"
                        up_Fillcombo()
                        pub_Period = Request.QueryString("t1")
                        pub_AffiliateID = Request.QueryString("t2")
                        pub_AffiliateName = Request.QueryString("t3")
                        pub_PO = Request.QueryString("t4")
                        pub_Del = Request.QueryString("t5")
                        pub_Kanban = Request.QueryString("t6")

                        'tabIndex()
                        'pSearch = False
                        'bindDataHeader(pub_Period, pub_AffiliateID, pub_PO, "", "", pub_AffiliateName)
                        'bindDataDetail(pub_Period, pub_AffiliateID, pub_PO, "", "")
                        'Call SaveDataMaster(ValidasiInput(pub_AffiliateID), pub_Period, pub_AffiliateID, pub_PO, Trim(txtSupplierCode.Text), Trim(txtCommercial.Text), Trim(rblDelivery.Value), Trim(rblPOKanban.Value), Trim(txtShipBy.Value))
                        'Call SaveDataDetail(ValidasiInput(pub_AffiliateID), pub_Period, pub_AffiliateID, pub_PO, Trim(txtSupplierCode.Text), Trim(txtCommercial.Text), Trim(rblDelivery.Value), Trim(rblPOKanban.Value), Trim(txtShipBy.Value))
                        'SaveDeliveryCls()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        'pSearch = True
                        If txtSuppAppDate.Text <> "" Or txtSuppPendDate.Text <> "" Or txtSuppUnpDate.Text <> "" Or txtPASIAppDate.Text <> "" Or txtAffFinalAppDate.Text <> "" Then
                            btnSubmit.Enabled = False
                            btnSendSupplier.Enabled = False
                        End If
                        'btnClear.Visible = False
                        'ScriptManager.RegisterStartupScript(AffiliateSubmit, AffiliateSubmit.GetType(), "scriptKey", "txtAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');", True)
                    ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                        Session("MenuDesc") = "AFFILIATE ORDER DETAIL ENTRY"
                        up_Fillcombo()
                        pub_Period = clsNotification.DecryptURL(Request.QueryString("t1"))
                        pub_AffiliateID = clsNotification.DecryptURL(Request.QueryString("t2"))
                        pub_AffiliateName = clsNotification.DecryptURL(Request.QueryString("t3"))
                        pub_PO = clsNotification.DecryptURL(Request.QueryString("id2"))
                        'pub_Del = clsNotification.DecryptURL(Request.QueryString("t5"))
                        'pub_Kanban = clsNotification.DecryptURL(Request.QueryString("t6"))

                        'tabIndex()
                        'pSearch = False
                        'bindDataHeader(pub_Period, pub_AffiliateID, pub_PO, "", "", pub_AffiliateName)
                        'bindDataDetail(pub_Period, pub_AffiliateID, pub_PO, "", "")
                        'Call SaveDataMaster(ValidasiInput(pub_AffiliateID), pub_Period, pub_AffiliateID, pub_PO, Trim(txtSupplierCode.Text), Trim(txtCommercial.Text), Trim(rblDelivery.Value), Trim(rblPOKanban.Value), Trim(txtShipBy.Value))
                        'Call SaveDataDetail(ValidasiInput(pub_AffiliateID), pub_Period, pub_AffiliateID, pub_PO, Trim(txtSupplierCode.Text), Trim(txtCommercial.Text), Trim(rblDelivery.Value), Trim(rblPOKanban.Value), Trim(txtShipBy.Value))
                        'SaveDeliveryCls()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        'pSearch = True
                        If txtSuppAppDate.Text <> "" Or txtSuppPendDate.Text <> "" Or txtSuppUnpDate.Text <> "" Or txtPASIAppDate.Text <> "" Or txtAffFinalAppDate.Text <> "" Then
                            btnSubmit.Enabled = False
                            btnSendSupplier.Enabled = False
                        End If
                    Else
                        Session("MenuDesc") = "AFFILIATE ORDER DETAIL ENTRY"
                        'tabIndex()
                        'clear()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        btnClear.Visible = True
                    End If
                Else
                    dtPeriod.Value = Now
                    btnClear.Visible = True
                    'txtAffiliateID.Focus()
                    'tabIndex()
                    'clear()
                End If
            End If

            'If ls_AllowDelete = False Then btnDelete.Enabled = False
            If ls_AllowUpdate = False Then btnSubmit.Enabled = False

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 13, False, clsAppearance.PagerMode.ShowAllRecord)

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub btnSubMenu_Click(sender As Object, e As System.EventArgs) Handles btnSubMenu.Click
        If Session("M01Url") <> "" Then
            'Session.Remove("M01Url")
            Response.Redirect("~/AffiliateOrderExport/AffiliateOrderExportList.aspx")
        Else
            'Session.Remove("M01Url")
            Response.Redirect("~/MainMenu.aspx")
        End If
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
        'Combo Affiliate
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
#End Region
End Class