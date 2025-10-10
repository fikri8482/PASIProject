Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports System.Net
Imports System.Net.Mail

Public Class PORevEntry
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "C02"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String
    Dim pub_Period As Date
    Dim pub_PORevNo As String

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
            If Session("M01Url") <> "" Then
                If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                    Session("MenuDesc") = "PO REVISION ENTRY"
                    pub_PONo = Request.QueryString("id")
                    pub_Ship = Request.QueryString("t1")
                    pub_Commercial = Request.QueryString("t2")
                    pub_Period = Request.QueryString("t3")
                    Session("SupplierID") = Request.QueryString("t4")
                    pub_PORevNo = Request.QueryString("t5")

                    dtPeriodFrom.Value = pub_Period
                    cboPartNo.Text = pub_PONo
                    txtPORevNo.Text = pub_PORevNo
                    txtShip.Text = pub_Ship
                    txtCommercial.Text = pub_Commercial

                    Session("Mode") = "Update"

                    bindData()
                    bindPOStatus()

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    cboPartNo.ReadOnly = True
                    cboPartNo.BackColor = Color.FromName("#CCCCCC")
                    txtPORevNo.ReadOnly = True
                    txtPORevNo.BackColor = Color.FromName("#CCCCCC")
                    dtPeriodFrom.ReadOnly = True
                    dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")

                    If clsPO.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID")) = "YES" Then
                        rdrKanban2.Checked = True
                    Else
                        rdrKanban3.Checked = True
                    End If
                ElseIf (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                    Session("MenuDesc") = "PO REVISION ENTRY"
                    pub_PONo = clsNotification.DecryptURL(Request.QueryString("id2"))
                    pub_Ship = clsNotification.DecryptURL(Request.QueryString("t1"))
                    pub_Commercial = clsNotification.DecryptURL(Request.QueryString("t2"))
                    pub_Period = clsNotification.DecryptURL(Request.QueryString("t3"))
                    Session("SupplierID") = clsNotification.DecryptURL(Request.QueryString("t4"))
                    pub_PORevNo = clsNotification.DecryptURL(Request.QueryString("t5"))

                    dtPeriodFrom.Value = pub_Period
                    cboPartNo.Text = pub_PONo
                    txtPORevNo.Text = pub_PORevNo
                    txtShip.Text = pub_Ship
                    txtCommercial.Text = pub_Commercial

                    Session("Mode") = "Update"

                    bindData()
                    bindPOStatus()

                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"
                    cboPartNo.ReadOnly = True
                    cboPartNo.BackColor = Color.FromName("#CCCCCC")
                    txtPORevNo.ReadOnly = True
                    txtPORevNo.BackColor = Color.FromName("#CCCCCC")
                    dtPeriodFrom.ReadOnly = True
                    dtPeriodFrom.BackColor = Color.FromName("#CCCCCC")

                    If clsPO.POKanban(pub_PONo, Session("AffiliateID"), Session("SupplierID")) = "YES" Then
                        rdrKanban2.Checked = True
                    Else
                        rdrKanban3.Checked = True
                    End If
                Else
                    Session("MenuDesc") = "PO REVISION ENTRY"
                    Session("Mode") = "New"
                    lblInfo.Text = ""
                    btnSubMenu.Text = "BACK"

                    dtPeriodFrom.Value = Now
                    rdrKanban2.Checked = True
                    up_FillCombo(dtPeriodFrom.Value)
                    cboPartNo.Focus()
                    clsGlobal.HideColumTanggal(dtPeriodFrom, grid)
                    'bindData()
                End If
            Else
                Session("Mode") = "New"
                dtPeriodFrom.Value = Now

                rdrKanban2.Checked = True
                up_FillCombo(dtPeriodFrom.Value)
                cboPartNo.Focus()
                'bindData()
                clsGlobal.HideColumTanggal(dtPeriodFrom, grid)
            End If

            lblInfo.Text = ""
        ElseIf IsCallback Then
            If grid.VisibleRowCount = 0 Then Exit Sub
        End If

        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" _
             Or e.Column.FieldName = "UnitDesc" Or e.Column.FieldName = "MinOrderQty" Or e.Column.FieldName = "Maker" _
             Or e.Column.FieldName = "KanbanCls" Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "QtyBox" _
             Or e.Column.FieldName = "ForecastN1" Or e.Column.FieldName = "ForecastN2" Or e.Column.FieldName = "ForecastN3" Or e.Column.FieldName = "POQty") _
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
            Response.Redirect("~/PurchaseOrderRevision/PORevList.aspx")
        Else
            Session.Remove("M01Url")
            Session.Remove("Mode")
            Session.Remove("SupplierID")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    'Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
    '    grid.JSProperties("cpMessage") = ""
    '    'Call bindData()
    'End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"

                    'If Session("B02IsSubmit") = "true" Then
                    '    grid.PageIndex = 0
                    '    Session.Remove("B02IsSubmit")
                    '    'cboPartNo.Text = Session("PONo")
                    '    grid.JSProperties("cpPONo") = Session("PONo")
                    '    Session.Remove("PONo")

                    '    grid.JSProperties("cpDate1") = Session("cpDate1")
                    '    Session.Remove("cpDate1")
                    '    grid.JSProperties("cpUser1") = Session("cpUser1")
                    '    Session.Remove("cpUser1")
                    'End If
                    'If Session("mode") = "update" Then

                    'End If
                    'If uf_CheckPOExists() = False Then
                    '    If uf_CheckAvailablePO() = False Then
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("YA010IsSubmit") = lblInfo.Text
                    End If
                    '    Else
                    'Call clsMsg.DisplayMessage(lblInfo, "5014", clsMessage.MsgType.ErrorMessage)
                    'grid.JSProperties("cpMessage") = lblInfo.Text
                    'Session("YA010IsSubmit") = lblInfo.Text
                    '    End If
                    'Else
                    'Call clsMsg.DisplayMessage(lblInfo, "5013", clsMessage.MsgType.ErrorMessage)
                    'grid.JSProperties("cpMessage") = lblInfo.Text
                    'Session("YA010IsSubmit") = lblInfo.Text
                    'End If
                Case "loadSave"
                    If Session("B02IsSubmit") = "true" Then
                        'grid.PageIndex = 0
                        Session.Remove("B02IsSubmit")
                        'cboPartNo.Text = Session("PONo")
                        ' grid.JSProperties("cpPONo") = Session("PONo")
                        'Session.Remove("PONo")

                        grid.JSProperties("cpDate1") = Session("cpDate1")
                        Session.Remove("cpDate1")
                        grid.JSProperties("cpUser1") = Session("cpUser1")
                        Session.Remove("cpUser1")
                    End If
                    bindData()
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "loadHeijunka"
                    'Call bindHeijunka()
                Case "savedata"
                    'Call saveData()
                Case "saveApprove"
                    Call uf_Approve()
                    Call bindPOStatus()
                Case "aftersave"
                    'bindHeijunka()
            End Select

EndProcedure:
            'Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Session("YA010IsSubmit") = ""
        End Try
    End Sub

    Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
        uf_Approve()
        sendEmail()
        bindPOStatus("update")
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""

        Dim ls_PartNo As String = "", ls_POQty As Double = 0
        Dim ls_D1 As Double = 0, ls_D2 As Double = 0, ls_D3 As Double = 0, ls_D4 As Double = 0, ls_D5 As Double = 0
        Dim ls_D6 As Double = 0, ls_D7 As Double = 0, ls_D8 As Double = 0, ls_D9 As Double = 0, ls_D10 As Double = 0
        Dim ls_D11 As Double = 0, ls_D12 As Double = 0, ls_D13 As Double = 0, ls_D14 As Double = 0, ls_D15 As Double = 0
        Dim ls_D16 As Double = 0, ls_D17 As Double = 0, ls_D18 As Double = 0, ls_D19 As Double = 0, ls_D20 As Double = 0
        Dim ls_D21 As Double = 0, ls_D22 As Double = 0, ls_D23 As Double = 0, ls_D24 As Double = 0, ls_D25 As Double = 0
        Dim ls_D26 As Double = 0, ls_D27 As Double = 0, ls_D28 As Double = 0, ls_D29 As Double = 0, ls_D30 As Double = 0
        Dim ls_D31 As Double = 0

        Dim ls_AffiliateID As String = Session("AffiliateID")

        Dim a As Integer

        ' FOR VALIDATION
        If uf_CheckPOExists() = True And Session("mode") = "New" Then
            Call clsMsg.DisplayMessage(lblInfo, "5013", clsMessage.MsgType.ErrorMessage)
            grid.JSProperties("cpMessage") = lblInfo.Text
            Session("YA010IsSubmit") = lblInfo.Text
            Exit Sub
        End If
        If uf_CheckAvailablePO() = True And Session("mode") = "New" Then
            Call clsMsg.DisplayMessage(lblInfo, "5014", clsMessage.MsgType.ErrorMessage)
            grid.JSProperties("cpMessage") = lblInfo.Text
            Session("YA010IsSubmit") = lblInfo.Text
            Exit Sub
        End If

        a = e.UpdateValues.Count
        For iLoop = 0 To a - 1
            If e.UpdateValues(iLoop).NewValues("POQty") > e.UpdateValues(iLoop).NewValues("POQtyOld") Then
                Call clsMsg.DisplayMessage(lblInfo, "5015", clsMessage.MsgType.ErrorMessage)
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If
            If (e.UpdateValues(iLoop).NewValues("POQty") Mod e.UpdateValues(iLoop).NewValues("MinOrderQty")) <> 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "5005", clsMessage.MsgType.ErrorMessage)
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            Dim checkTotalDailyQty As Double = 0

            For i = 1 To 31
                checkTotalDailyQty = checkTotalDailyQty + e.UpdateValues(iLoop).NewValues("DeliveryD" & i)
            Next

            If checkTotalDailyQty <> e.UpdateValues(iLoop).NewValues("POQty") Then
                Call clsMsg.DisplayMessage(lblInfo, "5006", clsMessage.MsgType.ErrorMessage)
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            For i = 1 To 31
                If e.UpdateValues(iLoop).NewValues("DeliveryD" & i) <> 0 Then
                    If (e.UpdateValues(iLoop).NewValues("DeliveryD" & i) Mod e.UpdateValues(iLoop).NewValues("QtyBox")) <> 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "5007", clsMessage.MsgType.ErrorMessage)
                        Session("YA010IsSubmit") = lblInfo.Text
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Exit Sub
                    End If
                End If
            Next
        Next

        Dim pub_InsertNew As Boolean = uf_CheckPODetailExists()

        '''''NORMAL CASE NOT FOR HEIJUNKA'''''''''s
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("item_code")

                If grid.VisibleRowCount = 0 Then
                    'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, False, False, False)
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
                    Exit Sub
                End If

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Exit Sub
                End If

                If pub_InsertNew = False Then
                    If saveData(sqlConn, sqlTran) = False Then
                        'Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("YA010IsSubmit") = lblInfo.Text
                        Exit Sub
                    End If
                End If


                Dim pub_Count As Boolean = False

                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1
                    ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
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

                    ls_SQL = " UPDATE [dbo].[PORev_Detail] " & vbCrLf & _
                              "    SET [POQty] = '" & ls_POQty & "' " & vbCrLf & _
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
                                      " 	 WHERE PONo ='" & cboPartNo.Text & "' AND AffiliateID = '" & ls_AffiliateID & "' and PartNo = '" & ls_PartNo & "' and PORevNo = '" & txtPORevNo.Text & "'"
                    'ls_MsgID = "1002"
                    Dim SQLComm As SqlCommand '= pConStr.CreateCommand
                    SQLComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()
                Next iLoop

                sqlTran.Commit()
                Session("B02IsSubmit") = "true"
            End Using

            sqlConn.Close()
        End Using

        If Session("Mode") = "New" Then
            ls_MsgID = "1001"
        Else
            ls_MsgID = "1002"
        End If

        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        Call bindPOStatus("new")
        Session("Mode") = "Update"
        Session("YA010IsSubmit") = lblInfo.Text
        grid.JSProperties("cpMessage") = lblInfo.Text

    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        up_GridLoadWhenEventChange()
        cboPartNo.Text = ""
        txtShip.Text = ""

        cboPartNo.ReadOnly = False
        cboPartNo.BackColor = Color.FromName("#FFFFFF")
        dtPeriodFrom.ReadOnly = False
        dtPeriodFrom.BackColor = Color.FromName("#FFFFFF")
        txtShip.ReadOnly = False
        txtShip.BackColor = Color.FromName("#FFFFFF")

        dtPeriodFrom.Value = Now
        rdrKanban2.Checked = True

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

        'bindData()
    End Sub

    Private Sub ButtonDelete_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonDelete.Callback
        Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
        'If AlreadyUsed(pAffiliateID) = False Then
        Call deleteData(pAffiliateID)
        'End If
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        'e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                'If e.GetValue("AffiliateName") = "BY AFFILIATE" Then
                '    e.Cell.BackColor = Color.AliceBlue
                'End If
                If e.GetValue("AffiliateName") = "Before" Then
                    e.Cell.BackColor = Color.AliceBlue
                    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                End If
                If e.GetValue("AffiliateName") = "After" Then
                    If e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Or _
                        e.DataColumn.FieldName = "Maker" Or e.DataColumn.FieldName = "KanbanCls" Or e.DataColumn.FieldName = "UnitDesc" Or e.DataColumn.FieldName = "MinOrderQty" Or _
                        e.DataColumn.FieldName = "QtyBox" Or e.DataColumn.FieldName = "ForecastN1" Or e.DataColumn.FieldName = "ForecastN2" Or e.DataColumn.FieldName = "ForecastN3" Then
                        e.Cell.Text = ""
                    End If

                    If e.DataColumn.FieldName = "POQty" Or _
                        e.DataColumn.FieldName = "DeliveryD1" Or e.DataColumn.FieldName = "DeliveryD2" Or e.DataColumn.FieldName = "DeliveryD3" Or e.DataColumn.FieldName = "DeliveryD4" Or e.DataColumn.FieldName = "DeliveryD5" Or _
                        e.DataColumn.FieldName = "DeliveryD6" Or e.DataColumn.FieldName = "DeliveryD7" Or e.DataColumn.FieldName = "DeliveryD8" Or e.DataColumn.FieldName = "DeliveryD9" Or e.DataColumn.FieldName = "DeliveryD10" Or _
                        e.DataColumn.FieldName = "DeliveryD11" Or e.DataColumn.FieldName = "DeliveryD12" Or e.DataColumn.FieldName = "DeliveryD13" Or e.DataColumn.FieldName = "DeliveryD14" Or e.DataColumn.FieldName = "DeliveryD15" Or _
                        e.DataColumn.FieldName = "DeliveryD16" Or e.DataColumn.FieldName = "DeliveryD17" Or e.DataColumn.FieldName = "DeliveryD18" Or e.DataColumn.FieldName = "DeliveryD19" Or e.DataColumn.FieldName = "DeliveryD20" Or _
                        e.DataColumn.FieldName = "DeliveryD21" Or e.DataColumn.FieldName = "DeliveryD22" Or e.DataColumn.FieldName = "DeliveryD23" Or e.DataColumn.FieldName = "DeliveryD24" Or e.DataColumn.FieldName = "DeliveryD25" Or _
                        e.DataColumn.FieldName = "DeliveryD26" Or e.DataColumn.FieldName = "DeliveryD27" Or e.DataColumn.FieldName = "DeliveryD28" Or e.DataColumn.FieldName = "DeliveryD29" Or e.DataColumn.FieldName = "DeliveryD30" Or _
                        e.DataColumn.FieldName = "DeliveryD31" Then
                        e.Cell.BackColor = Color.White
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
            End If
        End With
    End Sub

    Private Sub cboPartNo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboPartNo.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Dim pPeriod As String = Mid(pAction, 12, 4) + "-" + clsGlobal.uf_GetShortMonth(Mid(pAction, 5, 3)) + "-" + "01"
        up_FillCombo(pPeriod)
    End Sub

    Private Sub ButtonPartNo_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonPartNo.Callback
        Call up_GridLoadWhenEventChange()
        Call bindHeader(Split(e.Parameter, "|")(0))
    End Sub
#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_Supplier = ""

        If IsNothing(Session("SupplierID")) = False Then
            ls_Supplier = Session("SupplierID")
        Else
            ls_Supplier = HF.Get("hfTest")
        End If
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select * from ( select  " & vbCrLf & _
                  " 	row_number() over (order by a.PartNo) as NoUrut,  " & vbCrLf & _
                  " 	'1' Urut, " & vbCrLf & _
                  " 	'New' Edited, " & vbCrLf & _
                  " 	'Before' AffiliateName, " & vbCrLf & _
                  " 	a.PartNo, " & vbCrLf & _
                  " 	b.PartName, " & vbCrLf & _
                  " 	case b.KanbanCls when '1' then 'YES' else 'NO' end KanbanCls, " & vbCrLf & _
                  " 	c.UnitCls, " & vbCrLf & _
                  " 	c.Description UnitDesc, " & vbCrLf & _
                  " 	e.MOQ MinOrderQty, " & vbCrLf & _
                  " 	e.QtyBox, " & vbCrLf

            ls_SQL = ls_SQL + " 	b.Maker, " & vbCrLf & _
                              " 	'' PONo, " & vbCrLf & _
                              " 	a.POQty, " & vbCrLf & _
                              " 	a.POQtyOld, " & vbCrLf & _
                              " 	ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD1,0)DeliveryD1, ISNULL(a.DeliveryD2,0)DeliveryD2, ISNULL(a.DeliveryD3,0)DeliveryD3, ISNULL(a.DeliveryD4,0)DeliveryD4, ISNULL(a.DeliveryD5,0)DeliveryD5, " & vbCrLf

            ls_SQL = ls_SQL + " 	ISNULL(a.DeliveryD6,0)DeliveryD6, ISNULL(a.DeliveryD7,0)DeliveryD7, ISNULL(a.DeliveryD8,0)DeliveryD8, ISNULL(a.DeliveryD9,0)DeliveryD9, ISNULL(a.DeliveryD10,0)DeliveryD10, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD11,0)DeliveryD11, ISNULL(a.DeliveryD12,0)DeliveryD12, ISNULL(a.DeliveryD13,0)DeliveryD13, ISNULL(a.DeliveryD14,0)DeliveryD14, ISNULL(a.DeliveryD15,0)DeliveryD15, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD16,0)DeliveryD16, ISNULL(a.DeliveryD17,0)DeliveryD17, ISNULL(a.DeliveryD18,0)DeliveryD18, ISNULL(a.DeliveryD19,0)DeliveryD19, ISNULL(a.DeliveryD20,0)DeliveryD20, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD21,0)DeliveryD21, ISNULL(a.DeliveryD22,0)DeliveryD22, ISNULL(a.DeliveryD23,0)DeliveryD23, ISNULL(a.DeliveryD24,0)DeliveryD24, ISNULL(a.DeliveryD25,0)DeliveryD25, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD26,0)DeliveryD26, ISNULL(a.DeliveryD27,0)DeliveryD27, ISNULL(a.DeliveryD28,0)DeliveryD28, ISNULL(a.DeliveryD29,0)DeliveryD29, ISNULL(a.DeliveryD30,0)DeliveryD30, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD31,0)DeliveryD31, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD1Old,0)DeliveryD1Old, ISNULL(a.DeliveryD2Old,0)DeliveryD2Old, ISNULL(a.DeliveryD3Old,0)DeliveryD3Old, ISNULL(a.DeliveryD4Old,0)DeliveryD4Old, ISNULL(a.DeliveryD5Old,0)DeliveryD5Old, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD6Old,0)DeliveryD6Old, ISNULL(a.DeliveryD7Old,0)DeliveryD7Old, ISNULL(a.DeliveryD8Old,0)DeliveryD8Old, ISNULL(a.DeliveryD9Old,0)DeliveryD9Old, ISNULL(a.DeliveryD10Old,0)DeliveryD10Old, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD11Old,0)DeliveryD11Old, ISNULL(a.DeliveryD12Old,0)DeliveryD12Old, ISNULL(a.DeliveryD13Old,0)DeliveryD13Old, ISNULL(a.DeliveryD14Old,0)DeliveryD14Old, ISNULL(a.DeliveryD15Old,0)DeliveryD15Old, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD16Old,0)DeliveryD16Old, ISNULL(a.DeliveryD17Old,0)DeliveryD17Old, ISNULL(a.DeliveryD18Old,0)DeliveryD18Old, ISNULL(a.DeliveryD19Old,0)DeliveryD19Old, ISNULL(a.DeliveryD20Old,0)DeliveryD20Old, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD21Old,0)DeliveryD21Old, ISNULL(a.DeliveryD22Old,0)DeliveryD22Old, ISNULL(a.DeliveryD23Old,0)DeliveryD23Old, ISNULL(a.DeliveryD24Old,0)DeliveryD24Old, ISNULL(a.DeliveryD25Old,0)DeliveryD25Old, " & vbCrLf

            ls_SQL = ls_SQL + " 	ISNULL(a.DeliveryD26Old,0)DeliveryD26Old, ISNULL(a.DeliveryD27Old,0)DeliveryD27Old, ISNULL(a.DeliveryD28Old,0)DeliveryD28Old, ISNULL(a.DeliveryD29Old,0)DeliveryD29Old, ISNULL(a.DeliveryD30Old,0)DeliveryD30Old, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD31Old,0)DeliveryD31Old, 0 UpdateFunction " & vbCrLf & _
                              " from PO_DetailUpload a  " & vbCrLf & _
                              " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " left join MS_UnitCls c on c.UnitCls = b.UnitCls " & vbCrLf & _
                              " left join MS_PartMapping e on a.AffiliateID = e.AffiliateID and a.SupplierID = e.SupplierID and a.PartNo = e.PartNo " & vbCrLf & _
                              " where PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & ls_Supplier & "'" & vbCrLf & _
                              " union all " & vbCrLf & _
                              " select 	 " & vbCrLf & _
                              " 	row_number() over (order by a.PartNo) as NoUrut, " & vbCrLf & _
                              " 	'2' Urut, " & vbCrLf

            ls_SQL = ls_SQL + " 	case when b.POQty is null then  'New' else 'Edit' end Edited, " & vbCrLf & _
                              " 	'After' AffiliateName, " & vbCrLf & _
                              " 	a.PartNo, " & vbCrLf & _
                              " 	c.PartName, " & vbCrLf & _
                              " 	case c.KanbanCls when '1' then 'YES' else 'NO' end KanbanCls, " & vbCrLf & _
                              " 	c.UnitCls, " & vbCrLf & _
                              " 	d.Description UnitDesc, " & vbCrLf & _
                              " 	f.MOQ MinOrderQty, " & vbCrLf & _
                              " 	f.QtyBox, " & vbCrLf & _
                              " 	c.Maker, " & vbCrLf & _
                              " 	'' PONo, " & vbCrLf & _
                              " 	ISNULL(b.POQty,a.POQty)POQty, " & vbCrLf

            ls_SQL = ls_SQL + " 	ISNULL(b.POQtyOld,a.POQty)POQtyOld, " & vbCrLf & _
                              " 	ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = a.PartNo and MF.AffiliateID = a.AffiliateID and YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD1,a.DeliveryD1)DeliveryD1, ISNULL(b.DeliveryD2,a.DeliveryD2)DeliveryD2, ISNULL(b.DeliveryD3,a.DeliveryD3)DeliveryD3, ISNULL(b.DeliveryD4,a.DeliveryD4)DeliveryD4, ISNULL(b.DeliveryD5,a.DeliveryD5)DeliveryD5, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD6,a.DeliveryD6)DeliveryD6, ISNULL(b.DeliveryD7,a.DeliveryD7)DeliveryD7, ISNULL(b.DeliveryD8,a.DeliveryD8)DeliveryD8, ISNULL(b.DeliveryD9,a.DeliveryD9)DeliveryD9, ISNULL(b.DeliveryD10,a.DeliveryD10)DeliveryD10, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD11,a.DeliveryD11)DeliveryD11, ISNULL(b.DeliveryD12,a.DeliveryD12)DeliveryD12, ISNULL(b.DeliveryD13,a.DeliveryD13)DeliveryD13, ISNULL(b.DeliveryD14,a.DeliveryD14)DeliveryD14, ISNULL(b.DeliveryD15,a.DeliveryD15)DeliveryD15, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD16,a.DeliveryD16)DeliveryD16, ISNULL(b.DeliveryD17,a.DeliveryD17)DeliveryD17, ISNULL(b.DeliveryD18,a.DeliveryD18)DeliveryD18, ISNULL(b.DeliveryD19,a.DeliveryD19)DeliveryD19, ISNULL(b.DeliveryD20,a.DeliveryD20)DeliveryD20, " & vbCrLf

            ls_SQL = ls_SQL + " 	ISNULL(b.DeliveryD21,a.DeliveryD21)DeliveryD21, ISNULL(b.DeliveryD22,a.DeliveryD22)DeliveryD22, ISNULL(b.DeliveryD23,a.DeliveryD23)DeliveryD23, ISNULL(b.DeliveryD24,a.DeliveryD24)DeliveryD24, ISNULL(b.DeliveryD25,a.DeliveryD25)DeliveryD25, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD26,a.DeliveryD26)DeliveryD26, ISNULL(b.DeliveryD27,a.DeliveryD27)DeliveryD27, ISNULL(b.DeliveryD28,a.DeliveryD28)DeliveryD28, ISNULL(b.DeliveryD29,a.DeliveryD29)DeliveryD29, ISNULL(b.DeliveryD30,a.DeliveryD30)DeliveryD30, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD31,a.DeliveryD31)DeliveryD31, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD1Old,a.DeliveryD1Old)DeliveryD1Old, ISNULL(b.DeliveryD2Old,a.DeliveryD2Old)DeliveryD2Old, ISNULL(b.DeliveryD3Old,a.DeliveryD3Old)DeliveryD3Old, ISNULL(b.DeliveryD4Old,a.DeliveryD4Old)DeliveryD4Old, ISNULL(b.DeliveryD5Old,a.DeliveryD5Old)DeliveryD5Old, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD6Old,a.DeliveryD6Old)DeliveryD6Old, ISNULL(b.DeliveryD7Old,a.DeliveryD7Old)DeliveryD7Old, ISNULL(b.DeliveryD8Old,a.DeliveryD8Old)DeliveryD8Old, ISNULL(b.DeliveryD9Old,a.DeliveryD9Old)DeliveryD9Old, ISNULL(b.DeliveryD10Old,a.DeliveryD10Old)DeliveryD10Old, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD11Old,a.DeliveryD11Old)DeliveryD11Old, ISNULL(b.DeliveryD12Old,a.DeliveryD12Old)DeliveryD12Old, ISNULL(b.DeliveryD13Old,a.DeliveryD13Old)DeliveryD13Old, ISNULL(b.DeliveryD14Old,a.DeliveryD14Old)DeliveryD14Old, ISNULL(b.DeliveryD15Old,a.DeliveryD15Old)DeliveryD15Old, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD16Old,a.DeliveryD16Old)DeliveryD16Old, ISNULL(b.DeliveryD17Old,a.DeliveryD17Old)DeliveryD17Old, ISNULL(b.DeliveryD18Old,a.DeliveryD18Old)DeliveryD18Old, ISNULL(b.DeliveryD19Old,a.DeliveryD19Old)DeliveryD19Old, ISNULL(b.DeliveryD20Old,a.DeliveryD20Old)DeliveryD20Old, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD21Old,a.DeliveryD21Old)DeliveryD21Old, ISNULL(b.DeliveryD22Old,a.DeliveryD22Old)DeliveryD22Old, ISNULL(b.DeliveryD23Old,a.DeliveryD23Old)DeliveryD23Old, ISNULL(b.DeliveryD24Old,a.DeliveryD24Old)DeliveryD24Old, ISNULL(b.DeliveryD25Old,a.DeliveryD25Old)DeliveryD25Old, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD26Old,a.DeliveryD26Old)DeliveryD26Old, ISNULL(b.DeliveryD27Old,a.DeliveryD27Old)DeliveryD27Old, ISNULL(b.DeliveryD28Old,a.DeliveryD28Old)DeliveryD28Old, ISNULL(b.DeliveryD29Old,a.DeliveryD29Old)DeliveryD29Old, ISNULL(b.DeliveryD30Old,a.DeliveryD30Old)DeliveryD30Old, " & vbCrLf & _
                              " 	ISNULL(b.DeliveryD31Old,a.DeliveryD31Old)DeliveryD31Old, 0 UpdateFunction " & vbCrLf & _
                              " from PO_DetailUpload a " & vbCrLf

            ls_SQL = ls_SQL + " left join PORev_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo and PORevNo = '" & txtPORevNo.Text & "' " & vbCrLf & _
                              " left join MS_Parts c on a.PartNo = c.PartNo " & vbCrLf & _
                              " left join MS_UnitCls d on d.UnitCls = c.UnitCls " & vbCrLf & _
                              " left join MS_PartMapping f on a.AffiliateID = f.AffiliateID and a.SupplierID = f.SupplierID and a.PartNo = f.PartNo " & vbCrLf & _
                              " where a.PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & ls_Supplier & "'" & vbCrLf & _
                              " )x where 'A' = 'A' " & pWhere & " " & vbCrLf & _
                              " order by NoUrut, Urut "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using

        clsGlobal.HideColumTanggal(dtPeriodFrom, grid)

    End Sub

    Private Sub bindPOStatus(Optional ByVal pUpdate As String = "", Optional ByVal pPONO As String = "")
        Dim ls_SQL As String = ""
        Dim ls_PONo As String = ""
        Dim ls_PORev As String = ""

        If pPONO <> "" Then
            ls_PONo = pPONO
        Else
            ls_PONo = cboPartNo.Text
        End If

        ls_PORev = txtPORevNo.Text

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
                              " from PORev_Master where PONo = '" & ls_PONo & "' and AffiliateID = '" & Session("AffiliateID") & "' and PORevNo = '" & ls_PORev & "'"


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

                    Call clsMsg.DisplayMessage(lblInfo, "1007", clsMessage.MsgType.InformationMessage)
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

            ls_SQL = " select top 0 '' Urut, '' NoUrut, '' Edited, '' AffiliateName, " & vbCrLf & _
                  " '' PartNo, '' PartName, '' KanbanCls,'' UnitDesc, '' MinOrderQty, '' QtyBox,'' Maker, " & vbCrLf & _
                  " '' PONo, 0 POQty, 0 POQtyOld, " & vbCrLf & _
                  " 0 ForecastN1, 0 ForecastN2, 0 ForecastN3, " & vbCrLf & _
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

    Private Sub deleteData(ByVal pPONo As String)
        Dim sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("DeleteRevisionMasterData")

                sql = " delete PORev_Detail " & vbCrLf & _
                    " where PONo='" & cboPartNo.Text.Trim & "' and AffiliateID='" & Session("AffiliateID") & "' and PORevNo = '" & txtPORevNo.Text & "' " & vbCrLf & _
                    " "

                Dim SqlComm As New SqlCommand(sql, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()
                SqlComm.Dispose()

                sql = " delete PORev_Master " & vbCrLf & _
                    " where PONo='" & cboPartNo.Text.Trim & "' and AffiliateID='" & Session("AffiliateID") & "' and PORevNo = '" & txtPORevNo.Text & "' " & vbCrLf & _
                    " "

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
                ls_sql = " Update PORev_Master set AffiliateApproveDate = getdate(), AffiliateApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & cboPartNo.Text & "' and PORevNo = '" & txtPORevNo.Text & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_FillCombo(ByVal pPeriod As String)
        Dim ls_SQL As String = ""

        ls_SQL = "select RTRIM(PONo) PONo, RTRIM(SupplierID) SupplierID from PO_Master where AffiliateID = '" & Session("AffiliateID") & "' and Year(Period) = '" & Year(pPeriod) & "' and month(Period) = '" & Month(pPeriod) & "' and (FinalApproveDate is null and PASIApproveDate is not null ) order by PONo " & vbCrLf
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
                .Columns(0).Width = 100
                .Columns.Add("SupplierID")
                .Columns(1).Width = 50

                .TextField = "PONo"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub bindHeader(ByVal pPONO As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	case when CommercialCls = '1' then 'YES' else 'NO' end CommercialCls, " & vbCrLf & _
                  " 	ShipCls " & vbCrLf & _
                  " from PO_Master " & vbCrLf & _
                  " where PONo = '" & pPONO & "' and AffiliateID = '" & Session("AffiliateID") & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                txtShip.Text = ds.Tables(0).Rows(0)("ShipCls")
                ButtonPartNo.JSProperties("cpCommercial") = txtCommercial.Text
                ButtonPartNo.JSProperties("cpShip") = txtShip.Text
            Else
                ButtonPartNo.JSProperties("cpCommercial") = ""
                ButtonPartNo.JSProperties("cpShip") = ""
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Function uf_CheckAvailablePO() As Boolean
        Dim ls_Sql As String
        uf_CheckAvailablePO = False
        ls_Sql = "select PORevNo from PORev_Master where PONo = '" & cboPartNo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SeqNo = (select isnull(max(SeqNo),0) from PORev_Master where PONo = '" & cboPartNo.Text & "') and FinalApproveDate is null"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_Sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                uf_CheckAvailablePO = True
            Else
                uf_CheckAvailablePO = False
            End If

            sqlConn.Close()
        End Using

    End Function

    Private Function uf_CheckPOExists() As Boolean
        Dim ls_Sql As String
        uf_CheckPOExists = False
        ls_Sql = "select PORevNo from PORev_Master where PONo = '" & cboPartNo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and PORevNo = '" & txtPORevNo.Text & "'"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_Sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                uf_CheckPOExists = True
            Else
                uf_CheckPOExists = False
            End If

            sqlConn.Close()
        End Using

    End Function

    Private Function uf_CheckPODetailExists() As Boolean
        Dim ls_Sql As String
        uf_CheckPODetailExists = False
        ls_Sql = "select PORevNo from PORev_Detail where PONo = '" & cboPartNo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and PORevNo = '" & txtPORevNo.Text & "'"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_Sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                uf_CheckPODetailExists = True
            Else
                uf_CheckPODetailExists = False
            End If

            sqlConn.Close()
        End Using

    End Function

    Private Function saveData(ByVal pConStr As SqlConnection, ByVal pTrans As SqlTransaction) As Boolean
        Dim ls_Sql As String

        Try
            Dim ls_Supplier = ""

            If IsNothing(Session("SupplierID")) = False Then
                ls_Supplier = Session("SupplierID")
            Else
                ls_Supplier = HF.Get("hfTest")
            End If

            saveData = True
            Dim SQLCom As SqlCommand = pConStr.CreateCommand
            SQLCom.Connection = pConStr
            SQLCom.Transaction = pTrans

            ls_Sql = " INSERT INTO PORev_Detail " & vbCrLf & _
                  " select  " & vbCrLf & _
                  " 	'" & txtPORevNo.Text & "' PORevNo, " & vbCrLf & _
                  " 	PONo, " & vbCrLf & _
                  " 	AffiliateID, " & vbCrLf & _
                  " 	SupplierID, " & vbCrLf & _
                  " 	a.PartNo, " & vbCrLf & _
                  " 	(select isnull(max(SeqNo),0) + 1 from PORev_Master where PONo = '" & cboPartNo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "') SeqNo, " & vbCrLf & _
                  " 	a.POQty, " & vbCrLf & _
                  " 	a.POQty, " & vbCrLf

            ls_Sql = ls_Sql + " 	ISNULL(a.DeliveryD1,0)DeliveryD1, ISNULL(a.DeliveryD1,0)DeliveryD1,  " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD2,0)DeliveryD2, ISNULL(a.DeliveryD2,0)DeliveryD2,  " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD3,0)DeliveryD3, ISNULL(a.DeliveryD3,0)DeliveryD3, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD4,0)DeliveryD4, ISNULL(a.DeliveryD4,0)DeliveryD4, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD5,0)DeliveryD5, ISNULL(a.DeliveryD5,0)DeliveryD5, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD6,0)DeliveryD6, ISNULL(a.DeliveryD6,0)DeliveryD6, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD7,0)DeliveryD7, ISNULL(a.DeliveryD7,0)DeliveryD7, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD8,0)DeliveryD8, ISNULL(a.DeliveryD8,0)DeliveryD8, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD9,0)DeliveryD9, ISNULL(a.DeliveryD9,0)DeliveryD9, "

            ls_Sql = ls_Sql + " 	ISNULL(a.DeliveryD10,0)DeliveryD10, ISNULL(a.DeliveryD10,0)DeliveryD10, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD11,0)DeliveryD11, ISNULL(a.DeliveryD11,0)DeliveryD11, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD12,0)DeliveryD12, ISNULL(a.DeliveryD12,0)DeliveryD12, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD13,0)DeliveryD13, ISNULL(a.DeliveryD13,0)DeliveryD13, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD14,0)DeliveryD14, ISNULL(a.DeliveryD14,0)DeliveryD14, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD15,0)DeliveryD15, ISNULL(a.DeliveryD15,0)DeliveryD15, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD16,0)DeliveryD16, ISNULL(a.DeliveryD16,0)DeliveryD16, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD17,0)DeliveryD17, ISNULL(a.DeliveryD17,0)DeliveryD17, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD18,0)DeliveryD18, ISNULL(a.DeliveryD18,0)DeliveryD18, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD19,0)DeliveryD19, ISNULL(a.DeliveryD19,0)DeliveryD19, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD20,0)DeliveryD20, ISNULL(a.DeliveryD20,0)DeliveryD20, "

            ls_Sql = ls_Sql + " 	ISNULL(a.DeliveryD21,0)DeliveryD21, ISNULL(a.DeliveryD21,0)DeliveryD21, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD22,0)DeliveryD22, ISNULL(a.DeliveryD22,0)DeliveryD22, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD23,0)DeliveryD23, ISNULL(a.DeliveryD23,0)DeliveryD23, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD24,0)DeliveryD24, ISNULL(a.DeliveryD24,0)DeliveryD24, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD25,0)DeliveryD25, ISNULL(a.DeliveryD25,0)DeliveryD25, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD26,0)DeliveryD26, ISNULL(a.DeliveryD26,0)DeliveryD26, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD27,0)DeliveryD27, ISNULL(a.DeliveryD27,0)DeliveryD27, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD28,0)DeliveryD28, ISNULL(a.DeliveryD28,0)DeliveryD28, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD29,0)DeliveryD29, ISNULL(a.DeliveryD29,0)DeliveryD29, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD30,0)DeliveryD30, ISNULL(a.DeliveryD30,0)DeliveryD30, " & vbCrLf & _
                              " 	ISNULL(a.DeliveryD31,0)DeliveryD31, ISNULL(a.DeliveryD31,0)DeliveryD31, "

            ls_Sql = ls_Sql + " 	GETDATE(), '" & Session("UserID") & "', GETDATE(), '" & Session("UserID") & "' " & vbCrLf & _
                              " from PO_DetailUpload a  " & vbCrLf & _
                              " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " left join MS_UnitCls c on c.UnitCls = b.UnitCls " & vbCrLf & _
                              " where PONo = '" & cboPartNo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_Supplier & "'" & vbCrLf & _
                              "  " & vbCrLf & _
                              " INSERT INTO PORev_Master(PORevNo, PONo, AffiliateID, SupplierID, SeqNo, Period, ShipCls, EntryDate, EntryUser) " & vbCrLf & _
                              " select  " & vbCrLf & _
                              " 	'" & txtPORevNo.Text & "' PORevNo, " & vbCrLf & _
                              " 	a.PONo, a.AffiliateID, a.SupplierID,  "

            ls_Sql = ls_Sql + " 	SeqNo = (select isnull(max(SeqNo),0) + 1 from PORev_Master where PONo = '" & cboPartNo.Text & "' and AffiliateID = '" & Session("AffiliateID") & "'), b.Period, b.ShipCls, " & vbCrLf & _
                              " 	GETDATE(), '" & Session("UserID") & "' " & vbCrLf & _
                              " from PO_MasterUpload a left join PO_Master b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                              " where a.PONo = '" & cboPartNo.Text & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & ls_Supplier & "'"

            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

        Catch ex As Exception
            saveData = False
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Throw ex
        End Try
    End Function

    Private Sub sendEmail()
        Dim receiptEmail As String = ""
        Dim receiptCCEmail As String = ""
        Dim fromEmail As String = ""
        Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
        Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
        Dim ls_Body As String = ""

        '"http://localhost:5832/AffiliateRevision/AffiliateOrderevEntry.aspx?id=PO-SCN-KMK-Rev1&t1=5/2/2015&t2=PO-SCN-KMK-Rev1&t3=PO-SCN-KMK&t4=YES&t5=JAI%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&t6=PT.JATIM%20AUTOCOMP%20INDONESIA&t7=KMK%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&t8=PT.%20KMK%20PLASTICS%20INDONESIA&t9=1&t10=TRUCK%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%3E&t11=1&Session=~/AffiliateRevision/AffiliateOrderRevList.aspx"

        Dim ls_URl As String = "http://" & clsNotification.pub_ServerNamePASI & "/AffiliateRevision/AffiliateOrderevEntry.aspx?id2=" & clsNotification.EncryptURL(txtPORevNo.Text.Trim) & "&t1=" & clsNotification.EncryptURL(dtPeriodFrom.Value) & _
                               "&t2=" & clsNotification.EncryptURL(txtPORevNo.Text.Trim) & "&t3=" & clsNotification.EncryptURL(cboPartNo.Text.Trim) & "&t5=" & clsNotification.EncryptURL(Session("AffiliateID")) & _
                               "&t7=" & clsNotification.EncryptURL(Session("SupplierID")) & "&t9=" & clsNotification.EncryptURL("2") & "&Session=" & clsNotification.EncryptURL("~/AffiliateOrder/AffiliateOrderList.aspx")

        ls_Body = clsNotification.GetNotification("70", ls_URl)

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
        mailMessage.Subject = "Issued PO Revision No: " & txtPORevNo.Text & " from PO No: " & cboPartNo.Text

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