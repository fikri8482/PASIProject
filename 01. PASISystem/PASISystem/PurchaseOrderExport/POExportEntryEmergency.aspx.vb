Option Explicit On
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports System.Net
Imports System.Net.Mail
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO
Imports System.Data.OleDb
Imports OfficeOpenXml

Public Class POExportEntryEmergency
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
    Dim log As String = ""
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""

    Dim pStatus As Boolean

    Dim pPeriod As Date
    Dim pCommercial As String
    Dim pDeliveryCode As String
    Dim pDeliveryName As String
    Dim pPOEmergency As String
    Dim pShipBy As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pSupplierCode As String
    Dim pSupplierName As String
    Dim pPORevNo As String
    Dim pPO As String
    Dim pRemarks As String

    Dim pFilter As String
    Dim pub_Param As String
    Dim pstatusInsert As String

    Dim serverPath As String
    Dim fullPath As String

    Dim flag As Boolean = True
    'Dim clsPO As New ClsPOEEntryHeader
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Dim filterQty As String = ""


        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                'If Not IsNothing(Request.QueryString("prm")) Then
                Session("MenuDesc") = "PO EXPORT ENTRY (EMERGENCY)"

                If Session("EmergencyUrl") <> "" Then
                    param = Session("EmergencyUrl").ToString()
                    rdrCom1.Checked = True
                    rdrShipBy2.Checked = True
                    up_FillCombo()
                    dtPeriodFrom.Value = Now
                    dt1.Value = Now
                    
                    dt6.Value = Now
                    
                    dt11.Value = Now
                    
                    dt16.Value = Now
                    
                ElseIf Session("TampungDelivery") <> "" Then
                    param = Session("TampungDelivery").ToString()
                Else
                    param = Request.QueryString("prm").ToString
                    If param = "  'back'" Then
                        btnSubMenu.Text = "BACK"
                    Else
                        If pStatus = False Then

                            pPeriod = Split(param, "|")(0)
                            pAffiliateCode = Split(param, "|")(1)
                            pAffiliateName = Split(param, "|")(2)
                            'pSupplierCode = Split(param, "|")(2)
                            'pSupplierName = Split(param, "|")(3)
                            pDeliveryCode = Split(param, "|")(5)
                            pDeliveryName = Split(param, "|")(6)
                            pCommercial = Split(param, "|")(7)
                            pPOEmergency = Split(param, "|")(8)
                            pShipBy = Split(param, "|")(9)
                            'pRemarks = Split(param, "|")(10)
                            pPO = Split(param, "|")(10)

                            If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"
                            'If Trim(pPeriod) = "01 Jan 1900" Then pPeriod = Format(Now, "dd MMM yyyy")
                            'If Trim(pPeriod) = "" Then pPeriod = Format(Now, "dd MMM yyyy")

                            dtPeriodFrom.Value = pPeriod
                            rdrCom1.Value = pCommercial
                            cboAffiliate.Text = pAffiliateCode
                            txtAffiliate.Text = pAffiliateName
                            cboDelLoc.Text = pDeliveryCode
                            txtDelLoc.Text = pDeliveryName
                            'cboSupplierCode.Text = pSupplierCode
                            'txtSupplierName.Text = pSupplierName
                            rdrShipBy2.Text = pShipBy
                            txtPOEmergency.Text = pPOEmergency
                            'txtRevisionNo.Text = pPORevNo

                            'txtRemarks.Text = pRemarks
                            pStatus = True

                            Call bindDataHeader(pPOEmergency, pAffiliateCode, pPO)
                            Call bindData(pPOEmergency, pAffiliateCode, pPO)
                            'Call InitializeComponent(pPOEmergency, pAffiliateCode, pPO)
                            Session("pFilter") = pFilter
                            Session.Remove("EmergencyUrl")
                        End If
                    End If
                    btnSubMenu.Text = "BACK"
                    'End If
                End If
                '===============================================================================
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblInfo.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            grid.JSProperties("cpMessage") = lblInfo.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try

        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    'Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    '    ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
    '    ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

    '    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
    '        Session("M01Url") = Request.QueryString("Session")
    '        flag = False
    '    Else
    '        flag = True
    '    End If

    '    If (Not IsPostBack) AndAlso (Not IsCallback) Then
    '        up_FillCombo()

    '        If Session("M01Url") <> "" Then
    '            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
    '                Session("MenuDesc") = "PO EXPORT ENTRY EMERGENCY"
    '                pub_PONo = Request.QueryString("id")
    '                pub_Ship = Request.QueryString("t1")
    '                pub_Commercial = Request.QueryString("t2")
    '                pub_Period = Request.QueryString("t3")
    '                Session("SupplierID") = Request.QueryString("t4")

    '                dtPeriodFrom.Value = pub_Period
    '                txtPOEmergency.Text = pub_PONo

    '                If pub_Commercial = "YES" Then
    '                    rdrCom1.Checked = True
    '                Else
    '                    rdrCom2.Checked = True
    '                End If

    '                If pub_Ship = "BOAT" Then
    '                    rdrShipBy2.Checked = True
    '                Else
    '                    rdrShipBy3.Checked = True
    '                End If

    '                Session("Mode") = "Update"

    '                bindData()
    '                'bindPOStatus()

    '                lblInfo.Text = ""
    '                btnSubMenu.Text = "BACK"
    '                txtPOEmergency.ReadOnly = True
    '                txtPOEmergency.BackColor = Color.FromName("#CCCCCC")
    '                dtPeriodFrom.ReadOnly = True
    '                dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")
    '                rdrCom1.ReadOnly = True
    '                rdrCom2.ReadOnly = True
    '                rdrShipBy2.ReadOnly = True
    '                rdrShipBy3.ReadOnly = True

    '                'If clsPO.H_POEmergency(pub_PONo, Session("AffiliateID"), Session("SupplierID")) = "YES" Then
    '                '    rdrEmergency2.Checked = True
    '                'Else
    '                '    rdrEmergency3.Checked = True
    '                'End If


    '                'btnCraete.Text = "UPDATE"
    '                btnClear.Enabled = False

    '                'If txtDate1.Text.Trim <> "" And txtDate2.Text <> "" Then
    '                '    btnSubmit.Enabled = False
    '                '    btnDelete.Enabled = False
    '                'Else
    '                '    btnSubmit.Enabled = True
    '                '    btnDelete.Enabled = True
    '                'End If

    '                'If txtDate2.Text.Trim = "" Then
    '                '    btnApprove.Text = "APPROVE"
    '                'Else
    '                '    btnApprove.Text = "UNAPPROVE"
    '                'End If

    '                'If txtDate3.Text.Trim <> "" Then
    '                '    btnApprove.Enabled = False
    '                'Else
    '                '    btnApprove.Enabled = True
    '                'End If

    '            ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
    '                Session("MenuDesc") = "PO EXPORT ENTRY EMERGENCY"
    '                pub_PONo = clsNotification.DecryptURL(Request.QueryString("id2"))
    '                pub_Ship = clsNotification.DecryptURL(Request.QueryString("t1"))
    '                pub_Commercial = clsNotification.DecryptURL(Request.QueryString("t2"))
    '                pub_Period = clsNotification.DecryptURL(Request.QueryString("t3"))
    '                Session("SupplierID") = clsNotification.DecryptURL(Request.QueryString("t4"))

    '                dtPeriodFrom.Value = pub_Period
    '                txtPOEmergency.Text = pub_PONo
    '                'txtShip.Text = pub_Ship
    '                If pub_Commercial = "YES" Then
    '                    rdrCom1.Checked = True
    '                Else
    '                    rdrCom2.Checked = True
    '                End If

    '                If pub_Ship = "BOAT" Then
    '                    rdrShipBy2.Checked = True
    '                Else
    '                    rdrShipBy3.Checked = True
    '                End If

    '                Session("Mode") = "Update"

    '                bindData()
    '                'bindPOStatus()

    '                lblInfo.Text = ""
    '                btnSubMenu.Text = "BACK"
    '                txtPOEmergency.ReadOnly = True
    '                txtPOEmergency.BackColor = Color.FromName("#CCCCCC")
    '                dtPeriodFrom.ReadOnly = True
    '                dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")
    '                rdrCom1.ReadOnly = True
    '                rdrCom2.ReadOnly = True

    '                'If clsPO.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID")) = "YES" Then
    '                '    rdrEmergency2.Checked = True
    '                'Else
    '                '    rdrEmergency3.Checked = True
    '                'End If

    '                btnClear.Enabled = False


    '            Else
    '                Session("MenuDesc") = "PO EXPORT ENTRY EMERGENCY"
    '                Session("Mode") = "New"
    '                lblInfo.Text = ""
    '                btnSubMenu.Text = "BACK"
    '                txtPOEmergency.Focus()
    '                dtPeriodFrom.Value = Now
    '                rdrCom1.Checked = True
    '                rdrShipBy2.Checked = True
    '                dt1.Value = Now
    '                dt6.Value = Now
    '                dt11.Value = Now
    '                dt16.Value = Now
    '                'bindData()
    '                'btnApprove.Enabled = False
    '                'btnSubmit.Enabled = True
    '                'btnDelete.Enabled = True
    '                'btnClear.Enabled = True
    '            End If
    '        Else
    '            Session("Mode") = "New"
    '            txtPOEmergency.Focus()
    '            'btnApprove.Enabled = False
    '            'btnSubmit.Enabled = True
    '            'btnDelete.Enabled = True
    '            'btnClear.Enabled = True
    '            dtPeriodFrom.Value = Now
    '            rdrCom1.Checked = True
    '            rdrShipBy2.Checked = True
    '            dt1.Value = Now
    '            dt6.Value = Now
    '            dt11.Value = Now
    '            dt16.Value = Now
    '            'bindData()
    '        End If

    '        'bindData()
    '        'ColorGrid()
    '        lblInfo.Text = ""

    '    ElseIf IsCallback Then
    '        If grid.VisibleRowCount = 0 Then Exit Sub
    '    End If

    '    'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 4, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
    '    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
    'End Sub


    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "DetailPage" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" _
             Or e.Column.FieldName = "Description" Or e.Column.FieldName = "MOQ" Or e.Column.FieldName = "QtyBox" _
             Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "ShipCls" Or e.Column.FieldName = "CommercialCls" _
             Or e.Column.FieldName = "Period") _
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
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        Else
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        End If
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        grid.JSProperties("cpMessage") = ""
        Call bindData(pPOEmergency, pAffiliateCode, pPO)
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 4, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
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

                    Call bindData(pPOEmergency, pAffiliateCode, pPO)

                    'If grid.VisibleRowCount = 0 Then
                    '    Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    'End If
                    If Session("pub_Save") = True Then
                        If Session("B02IsSubmit") = "true" Then
                            grid.PageIndex = 0
                            Session.Remove("B02IsSubmit")
                            'ls.Text = Session("PONo")
                            'grid.JSProperties("cpPONo") = Session("PONo")
                            'Session.Remove("PONo")

                            'grid.JSProperties("cpDate1") = Session("cpDate1")
                            'Session.Remove("cpDate1")

                            'grid.JSProperties("cpUser1") = Session("cpUser1")
                            'Session.Remove("cpUser1")
                        End If

                        Call bindData(pPOEmergency, pAffiliateCode, pPO)

                        If grid.VisibleRowCount = 0 Then
                            Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                            grid.JSProperties("cpMessage") = lblInfo.Text
                        End If
                    Else
                        Call bindData(pPOEmergency, pAffiliateCode, pPO)

                        If grid.VisibleRowCount = 0 Then
                            Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                            grid.JSProperties("cpMessage") = lblInfo.Text
                        End If
                    End If

                    Session("pub_Save") = False
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    'Case "loadHeijunka"
                    '    Call bindHeijunka()
                    'Case "save"
                    '    Dim pAffiliateID As String = Split(e.Parameters, "|")(3)
                    '    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    '    Call SaveData1(lb_IsUpdate, _
                    '                     Split(e.Parameters, "|")(2), _
                    '                     Split(e.Parameters, "|")(3), _
                    '                     Split(e.Parameters, "|")(4), _
                    '                     Split(e.Parameters, "|")(5), _
                    '                     Split(e.Parameters, "|")(6), _
                    '                     Split(e.Parameters, "|")(7), _
                    '                     Split(e.Parameters, "|")(8), _
                    '                     Split(e.Parameters, "|")(9), _
                    '                     Split(e.Parameters, "|")(10), _
                    '                     Split(e.Parameters, "|")(11), _
                    '                     Split(e.Parameters, "|")(12), _
                    '                     Split(e.Parameters, "|")(13), _
                    '                     Split(e.Parameters, "|")(14), _
                    '                     Split(e.Parameters, "|")(15), _
                    '                     Split(e.Parameters, "|")(16), _
                    '                     Split(e.Parameters, "|")(17), _
                    '                     Split(e.Parameters, "|")(18), _
                    '                     Split(e.Parameters, "|")(19), _
                    '                     Split(e.Parameters, "|")(20), _
                    '                     Split(e.Parameters, "|")(21), _
                    '                     Split(e.Parameters, "|")(22), _
                    '                     Split(e.Parameters, "|")(23), _
                    '                     Split(e.Parameters, "|")(24), _
                    '                     Split(e.Parameters, "|")(25), _
                    '                     Split(e.Parameters, "|")(26), _
                    '                     Split(e.Parameters, "|")(27), _
                    '                     Split(e.Parameters, "|")(28), _
                    '                     Split(e.Parameters, "|")(29), _
                    '                     Split(e.Parameters, "|")(30), _
                    '                     Split(e.Parameters, "|")(31), _
                    '                     Split(e.Parameters, "|")(32), _
                    '                     Split(e.Parameters, "|")(33), _
                    '                     Split(e.Parameters, "|")(34), _
                    '                     Split(e.Parameters, "|")(35), _
                    '                     Split(e.Parameters, "|")(36), _
                    '                     Split(e.Parameters, "|")(37), _
                    '                     Split(e.Parameters, "|")(38), _
                    '                     Split(e.Parameters, "|")(39), _
                    '                     Split(e.Parameters, "|")(40), _
                    '                     Split(e.Parameters, "|")(41), _
                    '                     Split(e.Parameters, "|")(42), _
                    '                     Split(e.Parameters, "|")(43), _
                    '                     Split(e.Parameters, "|")(44), _
                    '                     Split(e.Parameters, "|")(45), _
                    '                     Split(e.Parameters, "|")(46), _
                    '                     Split(e.Parameters, "|")(47), _
                    '                     Split(e.Parameters, "|")(48), _
                    '                     Split(e.Parameters, "|")(49))

                Case "savedata"
                    Call saveData()

                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim pSuppType As String = ""
                    'If cbSupplierType.Text <> clsGlobal.gs_All Then
                    '    If cbSupplierType.Text = "PASI SUPPLIER" Then
                    '        pSuppType = "1"
                    '    Else
                    '        pSuppType = "0"
                    '    End If
                    'Else
                    '    pSuppType = clsGlobal.gs_All
                    'End If
                    Dim dtProd As DataTable = clsPOExportEmergency.GetTableEmergency(cboAffiliate.Text, txtAffiliate.Text, dtPeriodFrom.Value, txtPOEmergency.Text, cboDelLoc.Text, txtDelLoc.Text, txtOrder1.Text, dt1.Value, dt6.Value, dt11.Value, dt16.Value)
                    FileName = "TemplatePOExportEntryMonthly.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:3", psERR)
                    End If

                Case "saveApprove"
                    Call uf_Approve()
                    'Call bindPOStatus()
                    'Case "aftersave"
                    '    bindHeijunka()
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)

        'If txtDate1.Text <> "" Then
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

        'bindPOStatus("update")
        ' End If
        'sendEmailtoAffiliate()
        'sendEmailccPASI()
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 0
        Dim pIsUpdate As Boolean
        Dim ls_activeN As Boolean

        Dim ls_PartNo As String = "", ls_Ship As String = "", ls_UOM As String = "", ls_MOQ As String = ""
        Dim ls_Qty As Double = 0, ls_emergency As String = ""
        Dim ls_Forecast1 As Double = 0, ls_Forecast2 As Double = 0, ls_Forecast3 As Double = 0
        Dim ls_W1 As Double = 0, ls_W2 As Double = 0, ls_W3 As Double = 0, ls_W4 As Double = 0, ls_W5 As Double = 0
        Dim ls_TotalPOQty As Double = 0

        Dim ls_AffiliateID As String = cboAffiliate.Text.Trim
        Dim ls_SupplierID As String = ""
        Dim ls_TempSupplierID As String = ""
        Dim ls_PODeliveryBY As String = ""

        Dim ls_PrevForecast As String = ""
        Dim ls_Var As String = ""
        Dim ls_VarPercent As String = ""
        Dim ls_F1 As String = ""
        Dim ls_F2 As String = ""
        Dim ls_F3 As String = ""
        'Dim ls_TotalCurr As String = "", ls_TotalAmount As Double = 0

        Dim a As Integer, xy As Integer = 0
        'Dim ls_tampungPO(10) As String
        'Dim wiplist As New List(Of clsPO)

        Dim sqlstring As String = ""
        Dim publi_PONO As String = ""

        Dim ls_PONO As String = ""
        Dim ls_PONOp As String = ""
        Dim ls_error As String = ""

        Dim flgTempSupplier As String = ""
        Dim flgSupplier As Boolean = False
        'Dim ls_Seq As Integer = 0

        Session("pub_Save") = False

        'Checking PONo already Exists or not
        If Session("mode") = "New" Then

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                ls_SQL = " select max(PONo) + 1 PONo,AffiliateID from PO_Master_Export WHERE AffiliateID = '" & txtAffiliate.Text.Trim & "' group by AffiliateID "
                Dim sqlCmd10 As New SqlCommand(ls_SQL, sqlConn)
                Dim sqlDA10 As New SqlDataAdapter(sqlCmd10)
                Dim ds10 As New DataSet
                sqlDA10.Fill(ds10)

                If ds10.Tables(0).Rows.Count = 0 Then
                    ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                Else
                    ls_PONO = IIf(IsDBNull(ds10.Tables(0).Rows(0)("PONo")), 0, ds10.Tables(0).Rows(0)("PONo"))

                    'If (PO.H_QtyBox Mod ls_MOQ) <> 0 Then
                    '    If ls_error = "" Then
                    '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                    '    End If
                    'End If
                End If

                Dim ls_wherePONO As String = ""

                'If flgSupplier = False Then
                ls_wherePONO = "PONo ='" & ls_PONO & "'"
                'ElseIf flgSupplier = True Then
                '    ls_wherePONO = "SUBSTRING(PONo,1," & txtPONo.Text.Trim.Length & ") ='" & txtPONo.Text & "'"
                'End If

                sqlstring = "SELECT * FROM dbo.PO_Detail_Export WHERE " & ls_wherePONO & " AND AffiliateID = '" & cboAffiliate.Text.Trim & "' AND SupplierID = '" & Session("SupplierID") & "'"

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

                ls_Active = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())

                If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"
                If ls_activeN = True Then ls_activeN = "0" Else ls_activeN = "1"

                ls_PONO = Trim(e.UpdateValues(iLoop).NewValues("PONo").ToString())
                ls_AffiliateID = Trim(e.UpdateValues(iLoop).NewValues("AffiliateID").ToString())
                ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
                ls_SupplierID = Trim(e.UpdateValues(iLoop).NewValues("SupplierID").ToString())
                ls_MOQ = Trim(e.UpdateValues(iLoop).NewValues("MOQ").ToString())
                ls_W1 = Trim(e.UpdateValues(iLoop).NewValues("Week1").ToString())
                ls_W2 = Trim(e.UpdateValues(iLoop).NewValues("Week2").ToString())
                ls_W3 = Trim(e.UpdateValues(iLoop).NewValues("Week3").ToString())
                ls_W4 = Trim(e.UpdateValues(iLoop).NewValues("Week4").ToString())
                ls_W5 = Trim(e.UpdateValues(iLoop).NewValues("Week5").ToString())
                ls_TotalPOQty = Trim(e.UpdateValues(iLoop).NewValues("TotalPOQty").ToString())

                'If rdrShipBy2.Checked = True Then
                '    ls_emergency = 1
                'Else
                '    ls_emergency = 0
                'End If

                'If Session("mode") = "New" Then
                '    Dim ls_table As String = ""

                '    'If flgSupplier = True Then
                '    '    ls_table = "tempPODetail"
                '    'Else
                '    ls_table = "PO_Detail_Export"
                '    'End If
                ls_SQL = " select PONo,AffiliateID from PO_Master_Export WHERE PONo = '" & ls_PONO & "' and AffiliateID = '" & ls_AffiliateID & "' and SupplierID = '" & ls_SupplierID & "' "
                Dim sqlCmd10 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                Dim sqlDA10 As New SqlDataAdapter(sqlCmd10)
                Dim ds10 As New DataSet
                sqlDA10.Fill(ds10)

                If ds10.Tables(0).Rows.Count = 0 Then
                    ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                Else
                    ls_PONO = IIf(IsDBNull(ds10.Tables(0).Rows(0)("PONo")), 0, ds10.Tables(0).Rows(0)("PONo"))

                End If

                Dim pub_Count As Boolean = False

                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1
                    ls_Active = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())

                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                    Dim sqlComm As New SqlCommand
                    'sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    'Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    'If sqlRdr.Read Then
                    '    pIsUpdate = True
                    'Else
                    '    pIsUpdate = False
                    'End If
                    'sqlRdr.Close()



                    If ls_Active = "1" Then
                        If ls_activeN = "0" Then

                            ls_SQL = " select max(PONo) + 1 PONo,AffiliateID from PO_Master_Export WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' group by AffiliateID "
                            Dim sqlCmd11 As New SqlCommand(ls_SQL, sqlConn)
                            Dim sqlDA11 As New SqlDataAdapter(sqlCmd11)
                            Dim ds11 As New DataSet
                            sqlDA11.Fill(ds11)

                            If ds11.Tables(0).Rows.Count = 0 Then
                                ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                            Else
                                ls_PONOp = IIf(IsDBNull(ds11.Tables(0).Rows(0)("PONo")), 0, ds11.Tables(0).Rows(0)("PONo"))

                                'If (PO.H_QtyBox Mod ls_MOQ) <> 0 Then
                                '    If ls_error = "" Then
                                '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                '    End If
                                'End If
                            End If

                            ls_SQL = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                          "            ([PONo] " & vbCrLf & _
                                          "            ,[AffiliateID] " & vbCrLf & _
                                          "            ,[SupplierID] " & vbCrLf & _
                                          "            ,[PartNo] " & vbCrLf & _
                                          "            ,[Week1] " & vbCrLf & _
                                          "            ,[Week2] " & vbCrLf & _
                                          "            ,[Week3] " & vbCrLf & _
                                          "            ,[Week4] " & vbCrLf & _
                                          "            ,[Week5] " & vbCrLf & _
                                          "            ,[TotalPOQty] " & vbCrLf 


                            ls_SQL = ls_SQL + "            ,[EntryDate] " & vbCrLf & _
                                              "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & ls_PONOp & "' " & vbCrLf & _
                                              "            ,'" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                              "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                                              "            ,'" & ls_PartNo & "' " & vbCrLf & _
                                              "            ,'" & ls_W1 & "' "

                            ls_SQL = ls_SQL + "           ,'" & ls_W2 & "' " & vbCrLf & _
                                              "            ," & ls_W3 & " " & vbCrLf & _
                                              "            ,'" & ls_W4 & "' " & vbCrLf & _
                                              "            ,'" & ls_W5 & "' " & vbCrLf & _
                                              "           ,'" & ls_TotalPOQty & "' " & vbCrLf & _
                                              "            , getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' ) "

                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                        Else

                            ls_SQL = " UPDATE [dbo].[PO_Detail_Export] " & vbCrLf & _
                                          "    SET     [Week1] = '" & ls_W1 & "' " & vbCrLf & _
                                          "            ,[Week2] = '" & ls_W2 & "' " & vbCrLf & _
                                          "            ,[Week3] = '" & ls_W3 & "' " & vbCrLf & _
                                          "            ,[Week4] = '" & ls_W4 & "' " & vbCrLf & _
                                          "            ,[Week5] = '" & ls_W5 & "' " & vbCrLf & _
                                          "            ,[TotalPOQty] = '" & ls_TotalPOQty & "' " & vbCrLf 


                            ls_SQL = ls_SQL + "            ,[EntryDate] = getdate() " & vbCrLf & _
                                              "            ,[EntryUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                                " 	 WHERE PONo ='" & ls_PONO & "' AND AffiliateID = '" & ls_AffiliateID & "' AND SupplierID = '" & ls_SupplierID & "' and PartNo = '" & ls_PartNo & "'"
                            ls_MsgID = "1002"
                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()
                        End If

                    ElseIf ls_Active = "0" And pIsUpdate = True And Session("Mode") = "Update" Then
                        'ls_SQL = "  DELETE from dbo.PO_Detail_Export" & vbCrLf & _
                        '         "  where PONo = '" & ls_PONO & "'" & vbCrLf & _
                        '         "  and PartNo = '" & ls_PartNo & "' " & vbCrLf & _
                        '         "  and AffiliateID = '" & ls_AffiliateID & "' " & vbCrLf & _
                        '         "  and SupplierID = '" & ls_SupplierID & "'"
                        'sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        'sqlComm.ExecuteNonQuery()
                        'sqlComm.Dispose()
                    End If
                    ' End If
                Next iLoop

                'Dim pPeriod As String = ""
                'pPeriod = Format(DateAdd(DateInterval.Month, 1, CDate(pPeriod)), "yyyy-MM-01")

                Dim pub_Master As Boolean = False
                ls_SQL = "select * from PO_Master_Export where PONo = '" & ls_PONO & "' and AffiliateID = '" & ls_AffiliateID & "' and SupplierID = '" & ls_SupplierID & "'"

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
                    ls_SQL = " INSERT INTO PO_Master_Export " & _
                        "(Period, PONo, AffiliateID, SupplierID, CommercialCls, EmergencyCls, ShipCls, OrderNo1, OrderNo2, OrderNo3, OrderNo4, OrderNo5, " & _
                        " ETDVendor1, ETDPort1," & vbCrLf & _
                        " ETAPort1, ETAFactory1,  " & vbCrLf & _
                        " UpdateDate, UpdateUser)" & vbCrLf & _
                        " VALUES ('" & dtPeriodFrom.Value & "','" & ls_PONOp.Trim & "','" & cboAffiliate.Text.Trim & "','" & ls_SupplierID.Trim & "','" & IIf(rdrCom1.Checked = True, "1", "0") & "'," & vbCrLf & _
                        " '" & txtPOEmergency.Text.Trim & "','" & IIf(rdrShipBy2.Checked = True, "B", "A") & "','" & txtOrder1.Text.Trim & "','" & dt1.Value & "'," & vbCrLf & _
                        " '" & dt6.Value & "'," & vbCrLf & _
                        " '" & dt11.Value & "'," & vbCrLf & _
                        " '" & dt16.Value & "'," & vbCrLf & _
                        "  GETDATE(), '" & Session("UserID").ToString & "')" & vbCrLf
                    ls_MsgID = "1001"

                Else
                    'Update
                    ls_SQL = " UPDATE [dbo].[PO_Master_Export] " & vbCrLf & _
                                          "    SET     [ForwarderID] = '" & cboDelLoc.Text.Trim & "' " & vbCrLf & _
                                          "            ,[Period] = '" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                          "            ,[CommercialCls] = '" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                          "            ,[EmergencyCls] = '" & txtPOEmergency.Text.Trim & "' " & vbCrLf & _
                                          "            ,[ShipCls] = '" & IIf(rdrShipBy2.Checked = True, "B", "A") & "' " & vbCrLf & _
                                          "            ,[OrderNo1] = '" & txtOrder1.Text.Trim & "' " & vbCrLf & _
                                          "            ,[ETDVendor1] = '" & dt1.Value & "' " & vbCrLf & _
                                          "            ,[ETDPort1] = '" & dt6.Value & "' " & vbCrLf & _
                                          "            ,[ETAPort1] = '" & dt11.Value & "' " & vbCrLf & _
                                          "            ,[ETAFactory1] = '" & dt16.Value & "' " & vbCrLf 


                    ls_SQL = ls_SQL + "            ,[EntryDate] = getdate() " & vbCrLf & _
                                      "            ,[EntryUser] = '" & Session("UserID") & "' " & vbCrLf & _
           "  WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & ls_SupplierID & "' and PONo='" & ls_PONO & "'"

                End If

                sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm1.ExecuteNonQuery()
                sqlComm1.Dispose()
                publi_PONO = ls_PONO


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
        'Call bindPOStatus("new", publi_PONO)
        'Call bindData()
        Session("Mode") = "Update"
        Session("YA010IsSubmit") = lblInfo.Text
        grid.JSProperties("cpMessage") = lblInfo.Text

    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        up_GridLoadWhenEventChange()
        'Uploader.NullText = "Click here to browse files..."
        txtPOEmergency.Text = "M"
        cboDelLoc.SelectedIndex = -1
        'txtShip.Text = ""

        txtPOEmergency.ReadOnly = True
        txtPOEmergency.BackColor = Color.FromName("#CCCCCC")
        dtPeriodFrom.ReadOnly = False
        dtPeriodFrom.BackColor = Color.FromName("#FFFFFF")
        'txtShip.ReadOnly = False
        'txtShip.BackColor = Color.FromName("#FFFFFF")
        rdrCom1.ReadOnly = True
        rdrCom2.ReadOnly = False

        dtPeriodFrom.Value = Now
        rdrShipBy2.Checked = True
        rdrShipBy3.Checked = False

        'btnCraete.Text = "CREATE"

        lblInfo.Text = ""

        Session("Mode") = "New"
        Session.Remove("SupplierID")
        Session.Remove("Period")
        Session.Remove("PONoUpload")
        bindData(pPOEmergency, pAffiliateCode, pPO)
    End Sub

    'Private Sub ButtonDelete_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonDelete.Callback
    '    Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
    '    'If AlreadyUsed(pAffiliateID) = False Then
    '    Call deleteData(pAffiliateID)
    '    'End If
    'End Sub

    Private Sub grid_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        With e.Column
            If .FieldName = "MOQ" Then
                Dim ls_MinOrderQty As String = ""
                If IsNothing(e.GetFieldValue("MOQ")) Then ls_MinOrderQty = 0 Else ls_MinOrderQty = IIf(e.GetFieldValue("MOQ").ToString().Trim() = "", 0, e.GetFieldValue("MOQ"))
                e.DisplayText = clsGlobal.FormatQty(ls_MinOrderQty)
            End If

            If .FieldName = "QtyBox" Then
                Dim ls_QtyBox As String = ""
                If IsNothing(e.GetFieldValue("QtyBox")) Then ls_QtyBox = 0 Else ls_QtyBox = IIf(e.GetFieldValue("QtyBox").ToString().Trim() = "", 0, e.GetFieldValue("QtyBox"))
                e.DisplayText = clsGlobal.FormatQty(ls_QtyBox)
            End If

            If .FieldName = "Forecast1" Then
                Dim ls_Forecast1 As String = ""
                If IsNothing(e.GetFieldValue("Forecast1")) Then ls_Forecast1 = 0 Else ls_Forecast1 = IIf(e.GetFieldValue("Forecast1").ToString().Trim() = "", 0, e.GetFieldValue("Forecast1"))
                e.DisplayText = clsGlobal.FormatQty(ls_Forecast1)
            End If

            If .FieldName = "Forecast2" Then
                Dim ls_Forecast2 As String = ""
                If IsNothing(e.GetFieldValue("Forecast2")) Then ls_Forecast2 = 0 Else ls_Forecast2 = IIf(e.GetFieldValue("Forecast2").ToString().Trim() = "", 0, e.GetFieldValue("Forecast2"))
                e.DisplayText = clsGlobal.FormatQty(ls_Forecast2)
            End If

            If .FieldName = "Forecast3" Then
                Dim ls_Forecast3 As String = ""
                If IsNothing(e.GetFieldValue("Forecast3")) Then ls_Forecast3 = 0 Else ls_Forecast3 = IIf(e.GetFieldValue("Forecast3").ToString().Trim() = "", 0, e.GetFieldValue("Forecast3"))
                e.DisplayText = clsGlobal.FormatQty(ls_Forecast3)
            End If

            'For i = 1 To 31
            '    If .FieldName = "DeliveryD" & i Then
            '        Dim ls_DeliveryD As String = ""
            '        If IsNothing(e.GetFieldValue("DeliveryD" & i)) Then ls_DeliveryD = 0 Else ls_DeliveryD = IIf(e.GetFieldValue("DeliveryD" & i).ToString().Trim() = "", 0, e.GetFieldValue("DeliveryD" & i))
            '        e.DisplayText = clsGlobal.FormatQty(ls_DeliveryD)
            '    End If
            'Next
        End With
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData(ByVal pPOEmergency As String, ByVal pAffCode As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboAffiliate.Text <> "" Then
            'If cboAffiliate.Text <> clsGlobal.gs_All Then
            pWhere = pWhere + "and a.AffiliateID = '" & cboAffiliate.Text & "' "
        End If
        'End If
        If cboDelLoc.Text <> "" Then
            'If cboAffiliate.Text <> clsGlobal.gs_All Then
            pWhere = pWhere + "and a.ForwarderID = '" & cboDelLoc.Text & "' "
        End If

        If rdrCom1.Checked = True Then
            pWhere = pWhere + " and a.CommercialCls = '1'"
        End If

        If rdrCom2.Checked = True Then
            pWhere = pWhere + " and a.CommercialCls = '0'"
        End If

        If rdrShipBy2.Checked = True Then
            pWhere = pWhere + " and a.ShipCls = 'B'"
        End If

        If rdrShipBy3.Checked = True Then
            pWhere = pWhere + " and a.ShipCls = 'A'"
        End If

        If txtPOEmergency.Text.Trim <> "M" Then
            pWhere = pWhere + " and a.EmergencyCls = '" & txtPOEmergency.Text.Trim & "' "
        End If

        If dtPeriodFrom.Text = "" Then
            ls_SQL = ls_SQL + " AND a.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM") & "' " & vbCrLf
        End If

        If Session("Mode") = "Update" Then
            pWhere = pWhere + " and a.SupplierID = '" & Session("SupplierID") & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select " & vbCrLf & _
                  "    '0' AllowAccess, " & vbCrLf & _
                  " 	row_number() over (order by a.AffiliateID, a.SupplierID ) NoUrut,  " & vbCrLf & _
                  "  	RTRIM(a.PartNo)PartNo,  " & vbCrLf & _
                  "  	RTRIM(b.PartName)PartName,  " & vbCrLf & _
                  "  	RTRIM(c.Description)Description,  " & vbCrLf & _
                  "  	RTRIM(b.MOQ)MOQ,  " & vbCrLf & _
                  "  	b.QtyBox,  " & vbCrLf & _
                              "  	'0' TotalPOQty,   " & vbCrLf & _
                              "  	'' PONo,   " & vbCrLf & _
                              "  	'' ShipCls,   " & vbCrLf & _
                              "  	'' CommercialCls,   " & vbCrLf & _
                              "  	'' ForwarderID,   " & vbCrLf & _
                              "  	'' Period,   " & vbCrLf & _
                              "  	RTRIM(a.AffiliateID)AffiliateID,  " & vbCrLf & _
                              "  	RTRIM(a.SupplierID)SupplierID  " & vbCrLf & _
                              "  from MS_PartMapping a  " & vbCrLf & _
                              "  INNER join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
                              "  LEFT join MS_UnitCls c on c.UnitCls = b.UnitCls  " & vbCrLf & _
                              "  --Left join PO_Detail_Export d on a.PartNo = d.PartNo AND A.AffiliateID = D.AffiliateID AND A.SupplierID = D.SupplierID  " & vbCrLf & _
                              "  --Left join PO_Master_Export e on d.PONo = e.PONo AND d.AffiliateID = e.AffiliateID AND d.SupplierID = e.SupplierID  " & vbCrLf & _
                              "  where a.AffiliateID = '" & cboAffiliate.Text.Trim & "'  AND NOT EXISTS " & vbCrLf & _
                              "  ( " & vbCrLf & _
                              " SELECT * FROM  PO_Detail_Export X WHERE X.PONo = '" & pPONO & "' " & vbCrLf & _
                              "  ) "
            ' and a.SupplierID = '" & cboSupplier.Text.Trim & "'
            ls_SQL = ls_SQL + "  union all " & vbCrLf & _
                              "  select '1' AllowAccess, " & vbCrLf & _
                              " 	row_number() over (order by a.AffiliateID, a.SupplierID ) NoUrut,  " & vbCrLf & _
                              "  	RTRIM(B.PartNo)PartNo,  " & vbCrLf & _
                              "  	RTRIM(C.PartName)PartName,  " & vbCrLf & _
                              "  	RTRIM(d.Description)Description,  " & vbCrLf & _
                              "  	RTRIM(C.MOQ)MOQ,  " & vbCrLf & _
                              "  	C.QtyBox,  " & vbCrLf & _
                              "  	TotalPOQty,   " & vbCrLf & _
                              "  	a.PONo,   " & vbCrLf & _
                              "  	a.ShipCls,   " & vbCrLf & _
                              "  	a.CommercialCls,   " & vbCrLf & _
                              "  	a.ForwarderID,   " & vbCrLf & _
                              "  	a.Period,   " & vbCrLf & _
                              "  	RTRIM(a.AffiliateID)AffiliateID,  " & vbCrLf & _
                              "  	RTRIM(a.SupplierID)SupplierID  " & vbCrLf & _
                              " from PO_Master_Export a  " & vbCrLf & _
                              "  INNER join PO_Detail_Export b on a.PONo = b.PONo AND a.AffiliateID = B.AffiliateID AND a.SupplierID = B.SupplierID  " & vbCrLf & _
                              "  LEFT join MS_Parts c on c.PartNo = B.PartNo  " & vbCrLf & _
                              "  LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls  " & vbCrLf & _
                              "  where a.PONo = '" & pPONO & "' and a.EmergencyCls = '" & pPOEmergency & "' and 'A' = 'A' " & pWhere & " " & vbCrLf & _
                              "   union all "

            ls_SQL = ls_SQL + "   select  " & vbCrLf & _
                              "     '0' AllowAccess,  " & vbCrLf & _
                              "  	row_number() over (order by a.AffiliateID, a.SupplierID ) NoUrut,   " & vbCrLf & _
                              "   	RTRIM(a.PartNo)PartNo,   " & vbCrLf & _
                              "   	RTRIM(b.PartName)PartName,   " & vbCrLf & _
                              "   	RTRIM(c.Description)Description,   " & vbCrLf & _
                              "   	RTRIM(b.MOQ)MOQ,   " & vbCrLf & _
                              "   	b.QtyBox,   " & vbCrLf & _
                              "   	 '0' TotalPOQty,    " & vbCrLf & _
                              "   	'' PONo,    " & vbCrLf & _
                              "   	'' ShipCls,    " & vbCrLf & _
                              "   	'' CommercialCls,    "

            ls_SQL = ls_SQL + "   	'' ForwarderID,    " & vbCrLf & _
                              "   	'' Period,     " & vbCrLf & _
                              "   	RTRIM(a.AffiliateID)AffiliateID,   " & vbCrLf & _
                              "   	RTRIM(a.SupplierID)SupplierID   " & vbCrLf & _
                              "   from MS_PartMapping a   " & vbCrLf & _
                              "   INNER join MS_Parts b on a.PartNo = b.PartNo   " & vbCrLf & _
                              "   LEFT join MS_UnitCls c on c.UnitCls = b.UnitCls   " & vbCrLf & _
                              "   --Left join PO_Detail_Export d on a.PartNo = d.PartNo AND A.AffiliateID = D.AffiliateID AND A.SupplierID = D.SupplierID   " & vbCrLf & _
                              "   --Left join PO_Master_Export e on d.PONo = e.PONo AND d.AffiliateID = e.AffiliateID AND d.SupplierID = e.SupplierID   " & vbCrLf & _
                              "   where a.PartNo Not in ( select partNo from PO_detail_export where PONo = '" & pPONO & "') " & vbCrLf & _
                              "   and affiliateid = '" & pAffiliateCode & "'  "

            'If cboAffiliate.Text <> "" Then
            '    'If cboAffiliate.Text <> clsGlobal.gs_All Then
            '    pWhere = pWhere + "and e.AffiliateID = '" & cboAffiliate.Text & "' "
            'End If
            ''End If
            'If cboDelLoc.Text <> "" Then
            '    'If cboAffiliate.Text <> clsGlobal.gs_All Then
            '    pWhere = pWhere + "and e.ForwarderID = '" & cboDelLoc.Text & "' "
            'End If

            'If rdrCom1.Checked = True Then
            '    pWhere = pWhere + " and e.CommercialCls = '1'"
            'End If

            'If rdrCom2.Checked = True Then
            '    pWhere = pWhere + " and e.CommercialCls = '0'"
            'End If

            'If rdrShipBy2.Checked = True Then
            '    pWhere = pWhere + " and e.ShipCls = 'B'"
            'End If

            'If rdrShipBy3.Checked = True Then
            '    pWhere = pWhere + " and e.ShipCls = 'A'"
            'End If

            'If txtPOEmergency.Text.Trim <> "E" Then
            '    pWhere = pWhere + " and e.EmergencyCls = '" & txtPOEmergency.Text.Trim & "' "
            'End If

            'If dtPeriodFrom.Text = "" Then
            '    ls_SQL = ls_SQL + " AND e.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM") & "' " & vbCrLf
            'End If

            'If Session("Mode") = "Update" Then
            '    pWhere = pWhere + " and SupplierID = '" & Session("SupplierID") & "'"
            'End If

            'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            '    sqlConn.Open()

            'ls_SQL = " select " & vbCrLf & _
            '      "    '0' AllowAccess, " & vbCrLf & _
            '      " 	row_number() over (order by a.AffiliateID, a.SupplierID ) NoUrut,  " & vbCrLf & _
            '      "  	RTRIM(a.PartNo)PartNo,  " & vbCrLf & _
            '      "  	RTRIM(b.PartName)PartName,  " & vbCrLf & _
            '      "  	c.Description,  " & vbCrLf & _
            '      "  	RTRIM(b.MOQ)MOQ,  " & vbCrLf & _
            '      "  	b.QtyBox,  " & vbCrLf & _
            '      "  	'0' Week1,   " & vbCrLf & _
            '      "  	'0' Week2,   " & vbCrLf & _
            '      "  	'0' Week3,   "

            'ls_SQL = ls_SQL + "  	'0' Week4,   " & vbCrLf & _
            '                  "  	'0' Week5,   " & vbCrLf & _
            '                  "  	 TotalPOQty = (Week1+Week2+Week3+Week4+Week5),   " & vbCrLf & _
            '                  "  	e.PONo,   " & vbCrLf & _
            '                  "  	e.ShipCls,   " & vbCrLf & _
            '                  "  	e.CommercialCls,   " & vbCrLf & _
            '                  "  	e.ForwarderID,   " & vbCrLf & _
            '                  "  	e.Period,   " & vbCrLf & _
            '                  "  	RTRIM(a.AffiliateID)AffiliateID,  " & vbCrLf & _
            '                  "  	RTRIM(a.SupplierID)SupplierID  " & vbCrLf & _
            '                  "  from MS_PartMapping a  " & vbCrLf & _
            '                  "  INNER join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
            '                  "  LEFT join MS_UnitCls c on c.UnitCls = b.UnitCls  " & vbCrLf & _
            '                  "  LEFT join PO_Detail_Export d on a.PartNo = d.PartNo AND A.AffiliateID = D.AffiliateID AND A.SupplierID = D.SupplierID  " & vbCrLf & _
            '                  "  Left join PO_Master_Export e on d.PONo = e.PONo AND d.AffiliateID = e.AffiliateID AND d.SupplierID = e.SupplierID  " & vbCrLf & _
            '                  "  where a.AffiliateID = '" & cboAffiliate.Text.Trim & "'  AND NOT EXISTS " & vbCrLf & _
            '                  "  ( " & vbCrLf & _
            '                  " SELECT * FROM  PO_Detail_Export X WHERE X.PONo = e.PONo and 'A' = 'A' " & pWhere & " " & vbCrLf & _
            '                  "  ) "
            '' and a.SupplierID = '" & cboSupplier.Text.Trim & "'
            'ls_SQL = ls_SQL + "  union all " & vbCrLf & _
            '                  "  select '1' AllowAccess, " & vbCrLf & _
            '                  " 	row_number() over (order by e.AffiliateID, e.SupplierID ) NoUrut,  " & vbCrLf & _
            '                  "  	RTRIM(B.PartNo)PartNo,  " & vbCrLf & _
            '                  "  	RTRIM(C.PartName)PartName,  " & vbCrLf & _
            '                  "  	RTRIM(d.Description)Description,  " & vbCrLf & _
            '                  "  	RTRIM(C.MOQ)MOQ,  " & vbCrLf & _
            '                  "  	C.QtyBox,  " & vbCrLf & _
            '                  "  	B.Week1,   " & vbCrLf & _
            '                  "  	B.Week2,   " & vbCrLf & _
            '                  "  	B.Week3,   "

            'ls_SQL = ls_SQL + "  	B.Week4,   " & vbCrLf & _
            '                  "  	B.Week5,   " & vbCrLf & _
            '                  "  	TotalPOQty = (Week1+Week2+Week3+Week4+Week5),   " & vbCrLf & _
            '                  "  	e.PONo,   " & vbCrLf & _
            '                  "  	e.ShipCls,   " & vbCrLf & _
            '                  "  	e.CommercialCls,   " & vbCrLf & _
            '                  "  	e.ForwarderID,   " & vbCrLf & _
            '                  "  	e.Period,   " & vbCrLf & _
            '                  "  	RTRIM(e.AffiliateID)AffiliateID,  " & vbCrLf & _
            '                  "  	RTRIM(e.SupplierID)SupplierID  " & vbCrLf & _
            '                  " from PO_Master_Export e  " & vbCrLf & _
            '                  "  INNER join PO_Detail_Export b on e.PONo = b.PONo AND e.AffiliateID = B.AffiliateID AND e.SupplierID = B.SupplierID  " & vbCrLf & _
            '                  "  LEFT join MS_Parts c on c.PartNo = B.PartNo  " & vbCrLf & _
            '                  "  LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls  " & vbCrLf & _
            '                  "  where 'A' = 'A' " & pWhere & " " & vbCrLf & _
            '                  "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, " & vbCrLf & _
                  " '' PartNo, '' PartName, '' UnitCls, '' MOQ, '' QtyBox, " & vbCrLf & _
                  " '' PONo, 0 POQty, '' Week1, '' Week2, '' Week3,   " & vbCrLf & _
                  " '' Week4, '' week5, '' TotalPOQty, " & vbCrLf & _
                  " '' AffiliateID, '' SupplierID " & vbCrLf

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

    'Private Sub ColorGrid()
    '    grid.VisibleColumns(12).CellStyle.BackColor = Drawing.Color.White

    'End Sub

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            Dim ls_MsgID As String = ""


            'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            '    sqlConn.Open()

            '    ls_SQL = "SELECT AffiliateID, PartNo" & vbCrLf & _
            '                " FROM MS_Price " & _
            '                " WHERE AffiliateID= '" & Trim(pAffiliate) & "'"

            '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            '    Dim ds As New DataSet
            '    sqlDA.Fill(ds)

            '    If ds.Tables(0).Rows.Count > 0 And grid.FocusedRowIndex = -1 Then
            '        ls_MsgID = "6018"
            '        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            '        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            '        flag = False
            '        Return False
            '    ElseIf ds.Tables(0).Rows.Count > 0 Then
            '        lblInfo.Text = "Affiliate ID with ID " & txtPartNo.Text & " already exists in the database."
            '        Return False
            '    End If
            '    Return True
            '    sqlConn.Close()
            'End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Function

    Private Sub saveData()
        Dim pIsNewData As Boolean
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        Dim ls_PONo As String = ""
        Dim ls_error As String = ""
        Dim i As Integer

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
        'and AffiliateID ='" & pAffiliateID & "'
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT OrderNo1 FROM PO_Master_Export WHERE OrderNo1 ='" & txtOrder1.Text.Trim & "' and AffiliateID = '" & cboAffiliate.Text.Trim & "' "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False
                Else
                    pIsNewData = True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        If txtMode.Text = "update" Then
            flag = False
        Else
            flag = True
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then

                    ls_SQL = " select max(PONo) + 1 PONo,AffiliateID from PO_Master_Export WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' group by AffiliateID "
                    Dim sqlCmd11 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlDA11 As New SqlDataAdapter(sqlCmd11)
                    Dim ds11 As New DataSet
                    sqlDA11.Fill(ds11)

                    If ds11.Tables(0).Rows.Count = 0 Then
                        ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                    Else
                        ls_PONo = IIf(IsDBNull(ds11.Tables(0).Rows(0)("PONo")), 0, ds11.Tables(0).Rows(0)("PONo"))

                        'If (PO.H_QtyBox Mod ls_MOQ) <> 0 Then
                        '    If ls_error = "" Then
                        '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                        '    End If
                        'End If
                    End If

                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                  "            ([PONo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[Period] " & vbCrLf & _
                                  "            ,[CommercialCls] " & vbCrLf & _
                                   "            ,[ForwarderID] " & vbCrLf & _
                                      "            ,[EmergencyCls] " & vbCrLf & _
                                      "            ,[ShipCls] " & vbCrLf & _
                                      "            ,[OrderNo1], " & vbCrLf & _
                                      "            ,[ETDVendor1],[ETDPort1], " & vbCrLf & _
                                      "            ,[ETAPort1],[ETAFactory1], " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & ls_PONo.Trim & "' " & vbCrLf & _
                                  "            ,'" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "SupplierID").ToString & "' " & vbCrLf & _
                                  "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                  "            ,'" & cboDelLoc.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & txtPOEmergency.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrShipBy2.Checked = True, "B", "A") & "' " & vbCrLf & _
                                  "            ,'" & txtOrder1.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & dt1.Value & "' " & vbCrLf & _
                                  "            ,'" & dt6.Value & "' " & vbCrLf & _
                                  "            ,'" & dt11.Value & "' " & vbCrLf & _
                                  "            ,'" & dt16.Value & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "') "
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                ElseIf pIsNewData = False And flag = True Then
                    ls_MsgID = "6018"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    grid.JSProperties("cpType") = "error"
                    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA

                    ls_SQL = " UPDATE [dbo].[PO_Master_Export] " & vbCrLf & _
                                          "    SET     [ForwarderID] = '" & cboDelLoc.Text.Trim & "' " & vbCrLf & _
                                          "            ,[Period] = '" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                          "            ,[CommercialCls] = '" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                          "            ,[EmergencyCls] = '" & txtPOEmergency.Text.Trim & "' " & vbCrLf & _
                                          "            ,[ShipCls] = '" & IIf(rdrShipBy2.Checked = True, "B", "A") & "' " & vbCrLf & _
                                          "            ,[OrderNo1] = '" & txtOrder1.Text.Trim & "' " & vbCrLf & _
                                          "            ,[ETDVendor1] = '" & dt1.Value & "' " & vbCrLf & _
                                          "            ,[ETDPort1] = '" & dt6.Value & "' " & vbCrLf & _
                                          "            ,[ETAPort1] = '" & dt11.Value & "' " & vbCrLf & _
                                          "            ,[ETAFactory1] = '" & dt16.Value & "' " & vbCrLf 


                    ls_SQL = ls_SQL + "            ,[EntryDate] = getdate() " & vbCrLf & _
                                      "            ,[EntryUser] = '" & Session("UserID") & "' " & vbCrLf & _
           "  WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "' and PONo='" & ls_PONo & "'"
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()


                    grid.JSProperties("cpFunction") = "update"


                End If

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        grid.JSProperties("cpType") = "info"

    End Sub

    Private Sub saveData1()
        Dim i As Integer
        Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        Dim GrandTotal As Double = 0
        Dim GrandCurr As String = ""
        Dim ls_PONo As String = ""
        Dim ls_error As String = ""
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim pStartdate As Date

        pStartdate = Format(DateAdd(DateInterval.Month, 1, CDate(pStartdate)), "yyyy-MM-dd")


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

                    'ls_Sql = "delete PO_Detail where PONo = '" & ls_PONo.Trim & "' and AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "'"

                    'SQLCom.CommandText = ls_Sql
                    'SQLCom.ExecuteNonQuery()
                    ls_Sql = " select max(PONo) + 1 PONo,AffiliateID from PO_Master_Export WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' group by AffiliateID "
                    Dim sqlCmd11 As New SqlCommand(ls_Sql, SqlCon)
                    Dim sqlDA11 As New SqlDataAdapter(sqlCmd11)
                    Dim ds11 As New DataSet
                    sqlDA11.Fill(ds11)

                    If ds11.Tables(0).Rows.Count = 0 Then
                        ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                    Else
                        ls_PONo = IIf(IsDBNull(ds11.Tables(0).Rows(0)("PONo")), 0, ds11.Tables(0).Rows(0)("PONo"))

                        'If (PO.H_QtyBox Mod ls_MOQ) <> 0 Then
                        '    If ls_error = "" Then
                        '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                        '    End If
                        'End If
                    End If

                    '2.2 Insert New Detail Data
                    For i = 0 To grid.VisibleRowCount - 1
                        If grid.GetRowValues(i, "AllowAccess").ToString = "1" Then
                            ls_Sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[Week1] " & vbCrLf & _
                                      "            ,[Week2] " & vbCrLf & _
                                      "            ,[Week3] " & vbCrLf & _
                                      "            ,[Week4] " & vbCrLf & _
                                      "            ,[Week5] " & vbCrLf & _
                                      "            ,[TotalPOQty] " & vbCrLf & _
                                      "            ,[EntryDate] " & vbCrLf & _
                                      "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & ls_PONo.Trim & "' " & vbCrLf & _
                                              "            ,'" & cboAffiliate.Text & "' " & vbCrLf & _
                                              "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week1").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week2").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week3").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week4").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week5").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "TotalPOQty").ToString & "' " & vbCrLf & _
                                              "            , getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' ) "
                            'GrandTotal = GrandTotal + grid.GetRowValues(i, "Amount").ToString
                            'GrandCurr = grid.GetRowValues(i, "CurrCls").ToString
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()
                            ls_MsgID = "1002"
                        End If
                    Next i


                    '2.3 Insert data to Master
                    Dim pub_Master As Boolean = False
                    ls_Sql = " select max(PONo) + 1 PONo,AffiliateID from PO_Master_Export WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' group by AffiliateID "
                    Dim sqlCmd12 As New SqlCommand(ls_Sql, SqlCon)
                    Dim sqlDA12 As New SqlDataAdapter(sqlCmd12)
                    Dim ds12 As New DataSet
                    sqlDA12.Fill(ds12)

                    If ds12.Tables(0).Rows.Count = 0 Then
                        ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                    Else
                        ls_PONo = IIf(IsDBNull(ds11.Tables(0).Rows(0)("PONo")), 0, ds12.Tables(0).Rows(0)("PONo"))

                        'If (PO.H_QtyBox Mod ls_MOQ) <> 0 Then
                        '    If ls_error = "" Then
                        '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                        '    End If
                        'End If
                    End If

                    ls_Sql = "select * from PO_Master_Export where PONo = '" & ls_PONo.Trim & "' and AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "'"

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
                        ls_Sql = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                  "            ([PONo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[Period] " & vbCrLf & _
                                  "            ,[CommercialCls] " & vbCrLf & _
                                   "            ,[ForwarderID] " & vbCrLf & _
                                      "            ,[EmergencyCls] " & vbCrLf & _
                                      "            ,[ShipCls] " & vbCrLf & _
                                      "            ,[OrderNo1],[OrderNo2],[OrderNo3],[OrderNo4],[OrderNo5] " & vbCrLf & _
                                      "            ,[ETDVendor1],[ETDVendor2],[ETDVendor3],[ETDVendor4],[ETDVendor5],[ETDPort1],[ETDPort2],[ETDPort3],[ETDPort4],[ETDPort5] " & vbCrLf & _
                                      "            ,[ETAPort1],[ETAPort2],[ETAPort3],[ETAPort4],[ETAPort5],[ETAFactory1],[ETAFactory2],[ETAFactory3],[ETAFactory4],[ETAFactory5] " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & ls_PONo.Trim & "' " & vbCrLf & _
                                  "            ,'" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                  "            ,'" & dtPeriodFrom.Value & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                  "            ,'" & cboDelLoc.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & txtPOEmergency.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrShipBy2.Checked = True, "B", "A") & "' " & vbCrLf & _
                                  "            ,'" & txtOrder1.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & dt1.Value & "' " & vbCrLf & _
                                  "            ,'" & dt6.Value & "' " & vbCrLf & _
                                  "            ,'" & dt11.Value & "' " & vbCrLf & _
                                  "            ,'" & dt16.Value & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "' "


                        'ls_Sql = ls_Sql + "            ,'" & GrandCurr & "' " & vbCrLf & _
                        '                  "            ,'" & GrandTotal & "' " & vbCrLf & _
                    Else
                        'Update
                        ls_Sql = " UPDATE [dbo].[PO_Master_Export] SET  WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "' and PONo='" & ls_PONo & "'"

                    End If

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()
                Else
                    '2.1 delete data 
                    Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                    SQLCom.Connection = SqlCon
                    SQLCom.Transaction = SqlTran

                    'ls_Sql = "delete PO_Detail_Export where PONo = '" & ls_PONo & "' and AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "'"

                    'SQLCom.CommandText = ls_Sql
                    'SQLCom.ExecuteNonQuery()


                    '2.2 Insert New Detail Data
                    For i = 0 To grid.VisibleRowCount - 1
                        If grid.GetRowValues(i, "AllowAccess").ToString = "1" Then
                            ls_Sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[Week1] " & vbCrLf & _
                                      "            ,[Week2] " & vbCrLf & _
                                      "            ,[Week3] " & vbCrLf & _
                                      "            ,[Week4] " & vbCrLf & _
                                      "            ,[Week5] " & vbCrLf & _
                                      "            ,[TotalPOQty] " & vbCrLf & _
                                      "            ,[PreviousForecast] " & vbCrLf & _
                                      "            ,[Variance] " & vbCrLf & _
                                      "            ,[VariancePercentage] " & vbCrLf & _
                                      "            ,[Forecast1] " & vbCrLf & _
                                      "            ,[Forecast2] " & vbCrLf & _
                                      "            ,[Forecast3] " & vbCrLf & _
                                      "            ,[EntryDate] " & vbCrLf & _
                                      "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & ls_PONo & "' " & vbCrLf & _
                                              "            ,'" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                              "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week1").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week2").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week3").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week4").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "Week5").ToString & "' " & vbCrLf & _
                                              "            ,'" & grid.GetRowValues(i, "TotalPOQty").ToString & "' " & vbCrLf & _
                                              "            , getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' ) "
                            'GrandTotal = GrandTotal + grid.GetRowValues(i, "Amount").ToString
                            'GrandCurr = grid.GetRowValues(i, "CurrCls").ToString
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()
                            ls_MsgID = "1002"
                        End If
                    Next i


                    '2.3 Insert data to Master
                    Dim pub_Master As Boolean = False
                    ls_Sql = "select * from PO_Master_Export where PONo = '" & ls_PONo & "' and AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "'"

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

                        ls_Sql = " select max(PONo) + 1 PONo,AffiliateID from PO_Master_Export WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' group by AffiliateID "
                        Dim sqlCmd11 As New SqlCommand(ls_Sql, SqlCon)
                        Dim sqlDA11 As New SqlDataAdapter(sqlCmd11)
                        Dim ds11 As New DataSet
                        sqlDA11.Fill(ds11)

                        If ds11.Tables(0).Rows.Count = 0 Then
                            ls_error = "PONo not found in PO Master Export, please check again with PASI!"
                        Else
                            ls_PONo = IIf(IsDBNull(ds11.Tables(0).Rows(0)("PONo")), 0, ds11.Tables(0).Rows(0)("PONo"))

                            'If (PO.H_QtyBox Mod ls_MOQ) <> 0 Then
                            '    If ls_error = "" Then
                            '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                            '    End If
                            'End If
                        End If

                        ls_Sql = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                  "            ([PONo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[Period] " & vbCrLf & _
                                  "            ,[CommercialCls] " & vbCrLf & _
                                  "            ,[ForwarderID] " & vbCrLf & _
                                      "            ,[EmergencyCls] " & vbCrLf & _
                                      "            ,[ShipCls] " & vbCrLf & _
                                      "            ,[OrderNo1], " & vbCrLf & _
                                      "            ,[ETDVendor1],[ETDPort1] " & vbCrLf & _
                                      "            ,[ETAPort1],[ETAFactory1] " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & ls_PONo.Trim & "' " & vbCrLf & _
                                  "            ,'" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & Session("UserID") & "' " & vbCrLf & _
                                  "            ,'" & pStartdate & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrCom1.Checked = True, "1", "0") & "' " & vbCrLf & _
                                  "            ,'" & cboDelLoc.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & txtPOEmergency.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & IIf(rdrShipBy2.Checked = True, "B", "A") & "' " & vbCrLf & _
                                  "            ,'" & txtOrder1.Text.Trim & "' " & vbCrLf & _
                                  "            ,'" & dt1.Value & "' " & vbCrLf & _
                                  "            ,'" & dt6.Value & "' " & vbCrLf & _
                                  "            ,'" & dt11.Value & "' " & vbCrLf & _
                                  "            ,'" & dt16.Value & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "') "


                        'ls_Sql = ls_Sql + "            ,'" & GrandCurr & "' " & vbCrLf & _
                        '                  "            ,'" & GrandTotal & "' " & vbCrLf & _
                        'Else
                        'Update
                        'ls_Sql = " UPDATE [dbo].[PO_Master_Export] WHERE AffiliateID = '" & cboAffiliate.Text.Trim & "' and SupplierID = '" & Session("SupplierID") & "' and PONo='" & ls_PONo.Trim & "'"

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

                sql = " delete PO_Detail_Export " & vbCrLf & _
                    " where PONo='" & txtPOEmergency.Text.Trim & "' and AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
                    " and SupplierID ='" & Session("SupplierID") & "' "

                Dim SqlComm As New SqlCommand(sql, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()

                sql = " delete PO_Master_Export " & vbCrLf & _
                    " where PONo='" & txtPOEmergency.Text.Trim & "' and AffiliateID='" & Session("AffiliateID") & "' " & vbCrLf & _
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
                ls_sql = " Update PO_Master_Export set AffiliateApproveDate = getdate(), AffiliateApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & txtPOEmergency.Text & "' and SupplierID = '" & Session("SupplierID") & "'" & vbCrLf

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
                ls_sql = " Update PO_Master_Export set AffiliateApproveDate = NULL, AffiliateApproveUser = NULL" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & txtPOEmergency.Text & "' and SupplierID = '" & Session("SupplierID") & "'" & vbCrLf

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

            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNameAffiliate & "/PurchaseOrderExport/POExportEntry.aspx?id2=" & clsNotification.EncryptURL(txtPOEmergency.Text.Trim) & "&t1=" & clsNotification.EncryptURL(IIf(rdrCom1.Checked = True, "B", "A")) & _
                                    "&t2=" & clsNotification.EncryptURL(IIf(rdrCom1.Checked = True, "YES", "NO")) & "&t3=" & clsNotification.EncryptURL(dtPeriodFrom.Value) & "&t4=" & clsNotification.EncryptURL(Session("SupplierID")) & "&Session=" & clsNotification.EncryptURL("~/PurchaseOrderExport/POExportList.aspx")

            'ls_Body = clsNotification.GetNotification("10", ls_URl, txtPONo.Text.Trim, , , , txtPONo.Text.Trim & "-" & Session("SupplierID"))

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
            mailMessage.Subject = "Issued PONo: " & Trim(txtPOEmergency.Text)

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

            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateOrder/AffiliateOrderDetail.aspx?id2=" & clsNotification.EncryptURL(txtPOEmergency.Text.Trim) & "&t1=" & clsNotification.EncryptURL(dtPeriodFrom.Value) & _
                              "&t2=" & clsNotification.EncryptURL(Session("AffiliateID")) & "&t3=" & clsNotification.EncryptURL(Session("SupplierID")) & "&t4=" & clsNotification.EncryptURL(txtPOEmergency.Text) & "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderList.aspx")
            'ls_Body = clsNotification.GetNotification("10", ls_URl, txtPONo.Text.Trim, , , , txtPONo.Text.Trim & "-" & Session("SupplierID"))

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
            mailMessage.Subject = "Issued PONo: " & Trim(txtPOEmergency.Text)

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

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        'Person In Charge
        ls_SQL = "select RTRIM(ForwarderID) ForwarderID, ForwarderName from MS_Forwarder order by ForwarderID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboDelLoc
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("ForwarderID")
                .Columns(0).Width = 75
                .Columns.Add("ForwarderName")
                .Columns(1).Width = 400

                .TextField = "ForwarderID"
                .DataBind()
                .SelectedIndex = 0

            End With

            sqlConn.Close()
        End Using

        'Person In Charge
        ls_SQL = " select RTRIM(AffiliateID)AffiliateCode, isnull(AffiliateName,'') AffiliateName from MS_Affiliate " & vbCrLf & _
      " where affiliateID IN (select affiliateID from MS_PartMapping) " & vbCrLf & _
      " order by AffiliateCode " & vbCrLf & _
      "  "
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateCode")
                .Columns(0).Width = 75
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 400

                .TextField = "AffiliateCode"
                .DataBind()
                '.SelectedIndex = 0
                'txtAffiliate.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

    End Sub
#End Region

    '#Region "Upload v.K Edi"

    '    Private Sub ASPxCallback1_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
    '        Try
    '            Dim fi As New FileInfo(Server.MapPath("~\PurchaseOrderExport\PO EXPORTS ENTRY.xlsx"))
    '            If Not fi.Exists Then
    '                lblInfo.Text = "[9999] Excel Template Not Found !"
    '                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
    '                Exit Sub
    '            End If

    '            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("PO EXPORTS ENTRY.xlsx")

    '            'lblInfo.Text = "[9998] Download template successful"
    '            'ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
    '        End Try

    '    End Sub

    '    'Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
    '    '    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

    '    '    If e.GetValue("ErrorCls") = "" Then
    '    '    Else
    '    '        e.Cell.BackColor = Color.Red
    '    '    End If
    '    'End Sub

    '    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRevise.Click
    '        Call uf_Import()
    '    End Sub

    '    'Protected Sub UploadControl_FileUploadComplete(ByVal sender As Object, ByVal e As FileUploadCompleteEventArgs)
    '    '    'serverPath = Path.Combine(MapPath(""))
    '    '    'Dim resultFilePath As String = serverPath
    '    '    'fullPath = resultFilePath & "\Import\" & e.UploadedFile.FileName
    '    '    Try
    '    '        e.CallbackData = SavePostedFiles(e.UploadedFile)
    '    '    Catch ex As Exception
    '    '        e.IsValid = False
    '    '        e.ErrorText = ex.Message
    '    '    End Try
    '    'End Sub

    '    'Private Function SavePostedFiles(ByVal uploadedFile As UploadedFile) As String
    '    '    If (Not uploadedFile.IsValid) Then
    '    '        Return String.Empty
    '    '    End If

    '    '    serverPath = Path.Combine(MapPath(""))
    '    '    fullPath = serverPath & "\Import\" & tampung.Text

    '    '    uploadedFile.SaveAs(fullPath)
    '    'End Function

    '    Private Sub uf_Import()

    '        'Dim connStr As String
    '        'Try
    '        '    serverPath = Path.Combine(MapPath(""))
    '        '    fullPath = serverPath & "\Import\" & tampung.Text
    '        '    Dim Ext As String = Right(tampung.Text, 3)
    '        '    If Ext = "xls" Then
    '        '        connStr = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & fullPath & "';Extended Properties=Excel 8.0;"
    '        '    Else
    '        '        connStr = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & fullPath & "';Extended Properties=Excel 8.0;"
    '        '    End If
    '        '    Dim MyConnection As OleDbConnection
    '        '    Dim ds As DataSet
    '        '    Dim MyCommand As OleDbDataAdapter
    '        '    MyConnection = New OleDbConnection(connStr)
    '        '    MyCommand = New OleDbDataAdapter("select * from [Sheet1$A1:U65536]", MyConnection)
    '        '    ds = New System.Data.DataSet()
    '        '    MyCommand.Fill(ds)
    '        '    MyConnection.Close()
    '        '    grid.DataSource = ds.Tables(0).DefaultView
    '        '    grid.DataBind()
    '        '    System.Threading.Thread.Sleep(1000)
    '        '    'lblTotalRec.Text = "Total : " & grid.VisibleRowCount.ToString & " Record (s)"
    '        '    'grid.JSProperties("cpTotalRec") = lblTotalRec.Text

    '        '    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS.ToString())
    '        '    tampung.ForeColor = Color.FromName("#96C8FF")
    '        '    'grid.SettingsPager.PageSize = grid.VisibleRowCount.ToString + 1
    '        'Catch ex As Exception
    '        '    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        'End Try

    '        'If (Uploader.HasFile) Then        ' CHECK IF A FILE HAS BEEN SELECTED.
    '        '    If Not IsDBNull(Uploader.PostedFile) And _
    '        '        Uploader.PostedFile.ContentLength > 0 Then

    '        '        Dim oSqlBulk As SqlBulkCopy

    '        '        ' SET A CONNECTION WITH THE EXCEL FILE.
    '        '        Dim myExcelConn As OleDbConnection = _
    '        '            New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
    '        '                Server.MapPath(".") & "\" & Uploader.FileName() & _
    '        '                ";Extended Properties=Excel 12.0;")
    '        '        Try
    '        '            myExcelConn.Open()

    '        '            ' GET DATA FROM EXCEL SHEET.
    '        '            Dim objOleDB As New OleDbCommand("SELECT *FROM [Sheet1$]", myExcelConn)

    '        '            ' READ THE DATA EXTRACTED FROM THE EXCEL FILE.
    '        '            Dim objBulkReader As OleDbDataReader
    '        '            objBulkReader = objOleDB.ExecuteReader

    '        '            ' FINALLY, LOAD DATA INTO THE DATABASE TABLE.

    '        '            Dim sCon As String = "Data Source=DNA;Persist Security Info=False;" & _
    '        '                "Integrated Security=SSPI;" & _
    '        '                "Initial Catalog=DNA_Classified;User Id=sa;Password=;" & _
    '        '                "Connect Timeout=30;"

    '        '            Using con As SqlConnection = New SqlConnection(sCon)
    '        '                con.Open()
    '        '                oSqlBulk = New SqlBulkCopy(con)
    '        '                ' TABLE, DATA WILL BE UPLOADED TO.
    '        '                oSqlBulk.DestinationTableName = "PO_Detail_Export"
    '        '                oSqlBulk.WriteToServer(objBulkReader)
    '        '            End Using

    '        '            lblInfo.Text = "DATA IMPORTED SUCCESSFULLY."
    '        '            lblInfo.Attributes.Add("style", "color:green")

    '        '        Catch ex As Exception
    '        '            lblInfo.Text = ex.Message
    '        '            lblInfo.Attributes.Add("style", "color:red")
    '        '        Finally

    '        '            ' CLEAR.
    '        '            oSqlBulk.Close() : oSqlBulk = Nothing
    '        '            myExcelConn.Close() : myExcelConn = Nothing
    '        '        End Try
    '        '    End If
    '        'End If
    '        Dim dt As New System.Data.DataTable
    '        Dim dtHeader As New System.Data.DataTable
    '        Dim dtDetail As New System.Data.DataTable
    '        Dim tempDate As Date
    '        Dim ls_MOQ As Double = 0
    '        Dim ls_sql As String = ""
    '        Dim ls_SupplierID As String = """"

    '        Try
    '            lblInfo.ForeColor = Color.Green
    '            If Uploader.HasFile Then
    '                FileName = Uploader.PostedFile.FileName
    '                FileExt = Path.GetExtension(Uploader.PostedFile.FileName)
    '                FilePath = Ext & "\Import\" & FileName
    '                Dim fi As New FileInfo(Server.MapPath("~\Import\" & FileName))
    '                If fi.Exists Then
    '                    fi.Delete()
    '                    fi = New FileInfo(Server.MapPath("~\Import\" & FileName))
    '                End If
    '                Uploader.SaveAs(FilePath)

    '                Dim connStr As String = ""
    '                Select Case FileExt
    '                    Case ".xls"
    '                        'Excel 97-03
    '                        connStr = ConfigurationManager.ConnectionStrings("Excel03ConString").ConnectionString
    '                    Case ".xlsx"
    '                        'Excel 07
    '                        connStr = ConfigurationManager.ConnectionStrings("Excel07ConString").ConnectionString
    '                End Select

    '                connStr = String.Format(connStr, FilePath, "No")

    '                Dim MyConnection As New OleDbConnection(connStr)
    '                Dim MyCommand As New OleDbCommand
    '                Dim MyAdapter As New OleDbDataAdapter
    '                MyCommand.Connection = MyConnection
    '                MyConnection.Open()

    '                Dim dtSheets As DataTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    '                Dim listSheet As New List(Of String)
    '                Dim drSheet As DataRow
    '                Dim test As Date
    '                For Each drSheet In dtSheets.Rows
    '                    listSheet.Add(drSheet("TABLE_NAME").ToString())
    '                Next

    '                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '                    sqlConn.Open()

    '                    ''==========Table EXCEL Master==========
    '                    Dim pTableCode As String = listSheet(0)

    '                    Try

    '                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A1:C5]")
    '                        MyAdapter.SelectCommand = MyCommand
    '                        MyAdapter.Fill(dt)

    '                        'PONo
    '                        If IsDBNull(dt.Rows(4).Item(2)) Then
    '                            lblInfo.Text = "[9999] Invalid column ""PONo."", please check the file again!"
    '                            grid.JSProperties("cpMessage") = lblInfo.Text
    '                            Exit Sub
    '                        End If
    '                        If dt.Rows(4).Item(2).ToString.Trim.Length > 20 Then
    '                            lblInfo.Text = "[9999] Max 20 character in column ""PONo."" , please check the file again!"
    '                            grid.JSProperties("cpMessage") = lblInfo.Text
    '                            Exit Sub
    '                        End If

    '                        If dt.Rows.Count > 0 Then
    '                            'Period
    '                            If IsDBNull(dt.Rows(0).Item(2)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""Period"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Order No1
    '                            If IsDBNull(dt.Rows(6).Item(4)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""Order No 1"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If dt.Rows(6).Item(4).ToString.Trim.Length > 20 Then
    '                                lblInfo.Text = "[9999] Max 20 character in column ""Order No 1"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Order No2
    '                            If IsDBNull(dt.Rows(6).Item(5)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""Order No 2"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If dt.Rows(6).Item(5).ToString.Trim.Length > 20 Then
    '                                lblInfo.Text = "[9999] Max 20 character in column ""Order No 2"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Order No3
    '                            If IsDBNull(dt.Rows(6).Item(6)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""Order No 3"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If dt.Rows(6).Item(6).ToString.Trim.Length > 20 Then
    '                                lblInfo.Text = "[9999] Max 20 character in column ""Order No 3"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Order No4
    '                            If IsDBNull(dt.Rows(6).Item(7)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""Order No 4"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If dt.Rows(6).Item(7).ToString.Trim.Length > 20 Then
    '                                lblInfo.Text = "[9999] Max 20 character in column ""Order No 4"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Order No5
    '                            If IsDBNull(dt.Rows(6).Item(8)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""Order No 5"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If dt.Rows(6).Item(8).ToString.Trim.Length > 20 Then
    '                                lblInfo.Text = "[9999] Max 20 character in column ""Order No 5"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Vendor 1
    '                            If IsDBNull(dt.Rows(7).Item(4)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Vendor 1"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(7).Item(4).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format  in column  ""ETD Vendor 1"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Vendor 2
    '                            If IsDBNull(dt.Rows(7).Item(5)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Vendor 2"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(7).Item(5).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Vendor 2"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Vendor 3
    '                            If IsDBNull(dt.Rows(7).Item(6)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Vendor 3"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(7).Item(6).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Vendor 3"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Vendor 4
    '                            If IsDBNull(dt.Rows(7).Item(7)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Vendor 4"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(7).Item(7).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Vendor 7"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'Vendor 5
    '                            If IsDBNull(dt.Rows(7).Item(8)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Vendor 5"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(7).Item(8).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Vendor 2"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETD Port 1
    '                            If IsDBNull(dt.Rows(8).Item(4)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Port 1"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(8).Item(4).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Port 1"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETD Port 2
    '                            If IsDBNull(dt.Rows(8).Item(5)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Port 2"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(8).Item(5).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Port 2"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETD Port 3
    '                            If IsDBNull(dt.Rows(8).Item(6)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Port 3"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(8).Item(6).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Port 3"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETD Port 4
    '                            If IsDBNull(dt.Rows(8).Item(7)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Port 4"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(8).Item(7).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Port 4"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETD Port 5
    '                            If IsDBNull(dt.Rows(8).Item(8)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETD Port 5"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(8).Item(8).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETD Port 5"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Port 1
    '                            If IsDBNull(dt.Rows(9).Item(4)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Port 1"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(9).Item(4).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Port 1"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Port 2
    '                            If IsDBNull(dt.Rows(9).Item(5)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Port 2"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(9).Item(5).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Port 2"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Port 3
    '                            If IsDBNull(dt.Rows(9).Item(6)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Port 3"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(9).Item(6).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Port 3"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Port 4
    '                            If IsDBNull(dt.Rows(9).Item(7)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Port 4"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(9).Item(7).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Port 4"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETAPort 5
    '                            If IsDBNull(dt.Rows(9).Item(8)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Port 5"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(9).Item(8).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Port 5"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Factory 1
    '                            If IsDBNull(dt.Rows(10).Item(4)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Factory 1"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(10).Item(4).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Factory 1"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Factory 2
    '                            If IsDBNull(dt.Rows(10).Item(5)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Factory 2"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(10).Item(5).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Factory 2"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Factory 3
    '                            If IsDBNull(dt.Rows(10).Item(6)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Factory 3"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(10).Item(6).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Factory 3"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Factory 4
    '                            If IsDBNull(dt.Rows(10).Item(7)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Factory 4"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(10).Item(7).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Factory 4"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If

    '                            'ETA Factory 5
    '                            If IsDBNull(dt.Rows(10).Item(8)) Then
    '                                lblInfo.Text = "[9999] Invalid column ""ETA Factory 5"", please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                            If Date.TryParseExact(dt.Rows(10).Item(8).ToString(), "yyyy/mm/dd", System.Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.None, test) Then
    '                                lblInfo.Text = "[9999] Please use yyyy/mm/dd date format in column ""ETA Factory 5"" , please check the file again!"
    '                                grid.JSProperties("cpMessage") = lblInfo.Text
    '                                Exit Sub
    '                            End If
    '                        End If

    '                        Dim dtUploadHeader As New ClsPOEEntryHeader
    '                        Dim dtUploadHeaderList As New List(Of ClsPOEEntryHeader)

    '                        'Dim dtUploadDetail As New ClsPOEEntryDetail
    '                        Dim dtUploadDetailList As New List(Of ClsPOEEntryDetail)


    '                        'Get Header Data
    '                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A2:H2]")
    '                        MyAdapter.SelectCommand = MyCommand
    '                        MyAdapter.Fill(dtHeader)

    '                        If dtHeader.Rows.Count > 0 Then
    '                            For i = 0 To dtDetail.Rows.Count - 1
    '                                dtUploadHeader.H_Period = dtHeader.Rows(i).Item(0)
    '                                dtUploadHeader.H_PONo = dtHeader.Rows(i).Item(0)
    '                                dtUploadHeader.H_Commercial = dtHeader.Rows(i).Item(0)
    '                                dtUploadHeader.H_POEmergency = dtHeader.Rows(i).Item(0)
    '                                dtUploadHeader.H_AffiliateID = dtHeader.Rows(i).Item(0)
    '                                dtUploadHeader.H_SupplierID = dtHeader.Rows(i).Item(0)
    '                                dtUploadHeader.H_ShipBy = dtHeader.Rows(i).Item(0)
    '                            Next
    '                            'Try
    '                            '    tempDate = "01-" & dtHeader.Rows(0).Item(0)
    '                            '    Session("Period") = tempDate
    '                            'Catch ex As Exception
    '                            '    lblInfo.Text = "[9999] Invalid Period, please check the file again!"
    '                            '    grid.JSProperties("cpMessage") = lblInfo.Text
    '                            '    Exit Sub
    '                            'End Try

    '                            'dtUploadHeader.H_Period = tempDate
    '                            'If dtHeader.Rows(2).Item(0).ToString.Trim.ToUpper = "YES" Then
    '                            '    dtUploadHeader.H_POKanban = 1
    '                            'Else
    '                            '    dtUploadHeader.H_POKanban = 0
    '                            'End If
    '                            'dtUploadHeader.H_POKanban = dtHeader.Rows(2).Item(0)
    '                            'dtUploadHeader.H_UnitCode = dtHeader.Columns(A2).Item(0)
    '                            'dtUploadHeader.H_Center = dtHeader.Rows(2).Item(0)
    '                            'dtUploadHeader.H_Factory = dtHeader.Rows(4).Item(0)

    '                        End If


    '                        'Get Detail Data
    '                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B14:AR65536]")
    '                        MyAdapter.SelectCommand = MyCommand
    '                        MyAdapter.Fill(dtDetail)

    '                        If dtDetail.Rows.Count > 0 Then
    '                            For i = 0 To dtDetail.Rows.Count - 1
    '                                Dim dtUploadDetail As New ClsPOEEntryDetail
    '                                dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(0)
    '                                dtUploadDetail.D_PartName = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), 0, dtDetail.Rows(i).Item(8))
    '                                dtUploadDetail.D_UOM = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), 0, dtDetail.Rows(i).Item(9))
    '                                dtUploadDetail.D_MOQ = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), 0, dtDetail.Rows(i).Item(10))
    '                                dtUploadDetail.D_Qty = IIf(IsDBNull(dtDetail.Rows(i).Item(12)), 0, dtDetail.Rows(i).Item(12))
    '                                dtUploadDetail.D_Week1 = IIf(IsDBNull(dtDetail.Rows(i).Item(13)), 0, dtDetail.Rows(i).Item(13))
    '                                dtUploadDetail.D_Week2 = IIf(IsDBNull(dtDetail.Rows(i).Item(14)), 0, dtDetail.Rows(i).Item(14))
    '                                dtUploadDetail.D_Week3 = IIf(IsDBNull(dtDetail.Rows(i).Item(15)), 0, dtDetail.Rows(i).Item(15))
    '                                dtUploadDetail.D_Week4 = IIf(IsDBNull(dtDetail.Rows(i).Item(16)), 0, dtDetail.Rows(i).Item(16))
    '                                dtUploadDetail.D_Week5 = IIf(IsDBNull(dtDetail.Rows(i).Item(17)), 0, dtDetail.Rows(i).Item(17))
    '                                dtUploadDetail.D_TotalQty = IIf(IsDBNull(dtDetail.Rows(i).Item(18)), 0, dtDetail.Rows(i).Item(18))
    '                                dtUploadDetail.ForecastN = IIf(IsDBNull(dtDetail.Rows(i).Item(19)), 0, dtDetail.Rows(i).Item(19))
    '                                dtUploadDetail.Variance = IIf(IsDBNull(dtDetail.Rows(i).Item(20)), 0, dtDetail.Rows(i).Item(20))
    '                                dtUploadDetail.VariancePercentage = IIf(IsDBNull(dtDetail.Rows(i).Item(21)), 0, dtDetail.Rows(i).Item(21))
    '                                dtUploadDetail.Forecast1 = IIf(IsDBNull(dtDetail.Rows(i).Item(22)), 0, dtDetail.Rows(i).Item(22))
    '                                dtUploadDetail.Forecast2 = IIf(IsDBNull(dtDetail.Rows(i).Item(23)), 0, dtDetail.Rows(i).Item(23))
    '                                dtUploadDetail.Forecast3 = IIf(IsDBNull(dtDetail.Rows(i).Item(24)), 0, dtDetail.Rows(i).Item(24))

    '                                dtUploadDetailList.Add(dtUploadDetail)
    '                            Next
    '                        End If

    '                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
    '                            '01. Check PO already Exists 
    '                            ls_sql = "SELECT * FROM PO_Detail_Export WHERE PONo = '" & dtUploadHeader.H_PONo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'"

    '                            Dim sqlCmd As New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                            Dim sqlDA As New SqlDataAdapter(sqlCmd)
    '                            Dim ds As New DataSet
    '                            sqlDA.Fill(ds)

    '                            If ds.Tables(0).Rows.Count > 0 Then
    '                                If Not IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")) Then
    '                                    Call clsMsg.DisplayMessage(lblInfo, "5012", clsMessage.MsgType.ErrorMessage)
    '                                    Exit Sub
    '                                End If
    '                            End If

    '                            '01.01 Delete TempoaryData
    '                            ls_sql = "delete PO_Detail_Export where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and SupplierID = '" & dtUploadHeader.H_SupplierID & "' and PONo = '" & dtUploadHeader.H_PONo & "'"
    '                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                            sqlComm9.ExecuteNonQuery()
    '                            sqlComm9.Dispose()


    '                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
    '                            For i = 0 To dtUploadDetailList.Count - 1
    '                                Dim ls_error As String = ""
    '                                Dim PO As ClsPOEEntryDetail = dtUploadDetailList(i)

    '                                '02.1 Check PartNo di MS_Part
    '                                ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & PO.D_PartNo & "' "
    '                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
    '                                Dim ds2 As New DataSet
    '                                sqlDA2.Fill(ds2)

    '                                If ds2.Tables(0).Rows.Count = 0 Then
    '                                    ls_error = "PartNo not found in Part Master, please check again with PASI!"
    '                                Else
    '                                    ls_MOQ = IIf(IsDBNull(ds2.Tables(0).Rows(0)("MOQ")), 0, ds2.Tables(0).Rows(0)("MOQ"))

    '                                    If (PO.D_Qty Mod ls_MOQ) <> 0 Then
    '                                        If ls_error = "" Then
    '                                            ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
    '                                        End If
    '                                    End If
    '                                End If


    '                                '02.2 Check PartNo di Ms_PartMapping
    '                                ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'"
    '                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
    '                                Dim ds3 As New DataSet
    '                                sqlDA3.Fill(ds3)

    '                                If ds3.Tables(0).Rows.Count = 0 Then
    '                                    If ls_error = "" Then
    '                                        ls_error = "PartNo not found in Part Mapping, please check again with PASI!"
    '                                    End If
    '                                Else
    '                                    If i = 0 Then
    '                                        ls_SupplierID = IIf(IsDBNull(ds3.Tables(0).Rows(0)("SupplierID")), "", ds3.Tables(0).Rows(0)("SupplierID"))
    '                                    End If

    '                                    If ls_SupplierID <> ds3.Tables(0).Rows(0)("SupplierID") Then
    '                                        If ls_error = "" Then
    '                                            ls_error = "Can't Upload excel more than 1 supplier, please check the file again!"
    '                                        End If
    '                                    End If
    '                                End If


    '                                '02.3 Check PartNo di MS_Part
    '                                ls_sql = "SELECT * FROM dbo.PO_Detail_Export WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and SupplierID = '" & dtUploadHeader.H_SupplierID & "' and PONo = '" & dtUploadHeader.H_PONo & "'"
    '                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
    '                                Dim ds4 As New DataSet
    '                                sqlDA4.Fill(ds4)

    '                                If ds4.Tables(0).Rows.Count > 0 Then
    '                                    ls_sql = "delete PO_Detail_Export where PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and PONo = '" & dtUploadHeader.H_PONo & "'"
    '                                    Dim sqlComm1 = New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                                    sqlComm1.ExecuteNonQuery()
    '                                    sqlComm1.Dispose()
    '                                End If

    '                                ls_sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
    '                                      "            ([PONo] " & vbCrLf & _
    '                                      "            ,[AffiliateID] " & vbCrLf & _
    '                                      "            ,[SupplierID] " & vbCrLf & _
    '                                      "            ,[PartNo] " & vbCrLf & _
    '                                      "            ,[Week1] " & vbCrLf & _
    '                                      "            ,[Week2] " & vbCrLf & _
    '                                      "            ,[Week3] " & vbCrLf & _
    '                                      "            ,[Week4] " & vbCrLf & _
    '                                      "            ,[Week5] " & vbCrLf & _
    '                                      "            ,[TotalPOQty] " & vbCrLf & _
    '                                      "            ,[EntryDate] " & vbCrLf & _
    '                                      "            ,[EntryUser]) " & vbCrLf & _
    '                                              "      VALUES " & vbCrLf & _
    '                                              "            ('" & txtPONo.Text & "' " & vbCrLf & _
    '                                              "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
    '                                              "            ,'" & Session("SupplierID") & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "Week1").ToString & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "Week2").ToString & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "Week3").ToString & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "Week4").ToString & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "Week5").ToString & "' " & vbCrLf & _
    '                                              "            ,'" & grid.GetRowValues(i, "TotalPOQty").ToString & "' " & vbCrLf & _
    '                                              "            , getdate() " & vbCrLf & _
    '                                              "            ,'" & Session("UserID") & "' ) "
    '                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
    '                                sqlComm.ExecuteNonQuery()
    '                                sqlComm.Dispose()
    '                            Next
    '                            sqlTran.Commit()

    '                            Session("PONoUpload") = dtUploadHeader.H_PONo

    '                            lblInfo.Text = "[7001] Data Checking Done!"
    '                            lblInfo.ForeColor = Color.Blue
    '                            grid.JSProperties("cpMessage") = lblInfo.Text

    '                            Call bindData()
    '                        End Using
    '                    Catch ex As Exception
    '                        lblInfo.Text = ex.Message
    '                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '                        Exit Sub
    '                    End Try
    '                    dt.Reset()
    '                    dtDetail.Reset()
    '                    dtHeader.Reset()
    '                End Using
    '                MyConnection.Close()
    '            Else
    '                If FileName = "" Then
    '                    lblInfo.Text = "[9999] Please choose the file!"
    '                    up_GridLoadWhenEventChange()
    '                    grid.JSProperties("cpMessage") = lblInfo.Text
    '                    Exit Sub
    '                End If
    '            End If
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        End Try
    '    End Sub



    '    Protected Sub Uploader_FileUploadComplete(ByVal sender As Object, ByVal e As FileUploadCompleteEventArgs)
    '        Try
    '            e.CallbackData = SavePostedFiles(e.UploadedFile)
    '        Catch ex As Exception
    '            e.IsValid = False
    '            lblInfo.Text = ex.Message
    '        End Try
    '    End Sub

    '    Private Function SavePostedFiles(ByVal uploadedFile As UploadedFile) As String
    '        If (Not uploadedFile.IsValid) Then
    '            Return String.Empty
    '        End If

    '        Ext = Path.Combine(MapPath(""))
    '        FileName = Uploader.PostedFile.FileName
    '        FilePath = Ext & "\Import\" & FileName
    '        uploadedFile.SaveAs(FilePath)

    '        Return FilePath
    '    End Function
    '#End Region

#Region "Download"

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim NewFileName As String = Server.MapPath("~\PurchaseOrderExport\" & FileName)
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets(pSheetName)
            Dim irow As Integer = 0
            Dim icol As Integer = 0

            With ws
                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count - 0
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                    Next
                Next

                ''ALIGNMENT
                ''.Cells(rowstart + 1, icol, irow, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(iRow + space, colKanbanSeqNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPartName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvQty).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvCurr).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

                'Dim rgAll As ExcelRange = .Cells('.Cells(Space() - 2, colNo, grid.VisibleRowCount + (Space() - 1), colCount - 1)
                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 6)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub EpPlusDrawAllBorders(ByVal Rg As ExcelRange)
        With Rg
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
        End With
    End Sub

#End Region

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/PurchaseOrderExport/POExportUploadPOEmergencyFromExcel.aspx")
    End Sub

    Private Sub bindDataHeader(ByVal pPOEmergency As String, ByVal pAffCode As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT  " & vbCrLf & _
                  " 	a.OrderNo1, " & vbCrLf & _
                  " 	a.ETDVendor1 ,a.ETDPort1 ,a.ETAPort1 , a.ETAFactory1 " & vbCrLf & _
                  " FROM PO_Master_Export a " & vbCrLf & _
                  " --INNER JOIN UploadPOExport b ON a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.ForwarderID = b.ForwarderID " & vbCrLf & _
                  " WHERE a.PONo = '" & pPONO & "' and a.AffiliateID = '" & pAffCode & "' and a.EmergencyCls = '" & pPOEmergency & "' " & vbCrLf & _
                  "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtOrder1.Text = ds.Tables(0).Rows(0)("OrderNo1") & ""

                dt1.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETDVendor1")), "", Format(ds.Tables(0).Rows(0)("ETDVendor1"), "yyyy-MM-dd"))

                dt6.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETDPort1")), "", Format(ds.Tables(0).Rows(0)("ETDPort1"), "yyyy-MM-dd"))

                dt11.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETAPort1")), "", Format(ds.Tables(0).Rows(0)("ETAPort1"), "yyyy-MM-dd"))

                dt16.Text = If(IsDBNull(ds.Tables(0).Rows(0)("ETAFactory1")), "", Format(ds.Tables(0).Rows(0)("ETAFactory1"), "yyyy-MM-dd"))

                Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                Session("YA010IsSubmit") = lblInfo.Text
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindDataDetail(ByVal pPOEmergency As String, ByVal pAffCode As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""
        Dim jsScript As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select " & vbCrLf & _
                      " 	convert(char,row_number() over (order by b.PartNo asc))as NoUrut, " & vbCrLf & _
                      " 	b.PartNo, d.PartName, e.Description UnitDesc, d.MOQ, d.QtyBox, " & vbCrLf & _
                      " 	b.Week1 POQty, a.ETDVendor " & vbCrLf & _
                      " from PO_Master_Export a   " & vbCrLf & _
                      " inner join PO_DetailUpload_Export b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                      " inner join PO_MasterUpload_Export c on a.PONo = c.PONo and a.SupplierID = c.SupplierID and a.AffiliateID = c.AffiliateID " & vbCrLf & _
                      " left join MS_Parts d on b.PartNo = d.PartNo " & vbCrLf & _
                      " left join MS_UnitCls e on e.UnitCls = d.UnitCls " & vbCrLf & _
                      " where a.PONo = '" & pPONO & "' and a.AffiliateID = '" & pAffCode & "'  " & vbCrLf & _
                      "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            End With
            sqlConn.Close()
        End Using

    End Sub
End Class