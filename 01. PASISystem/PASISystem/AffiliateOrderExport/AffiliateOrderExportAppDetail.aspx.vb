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

Public Class AffiliateOrderExportAppDetail
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

#Region "CONTROL EVENTS"
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

                    'bindData(pub_Period, pub_PONo, pub_AffiliateID, pub_SupplierID, pub_Commercial, pub_DeliveyBy, pub_Ship)
                    'bindPOStatus("", pub_PONo, pub_AffiliateID)

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

                    'bindData(pub_Period, pub_PONo, pub_AffiliateID, pub_SupplierID, pub_Commercial, pub_DeliveyBy, pub_Ship)
                    'bindPOStatus("", pub_PONo, pub_AffiliateID)

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
            Response.Redirect("~/AffiliateOrderExport/AffiliateOrderExportAppList.aspx")
        Else
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
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
#End Region
End Class