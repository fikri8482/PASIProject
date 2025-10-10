Imports System.Data.SqlClient
Imports System.Transactions
Imports System.Drawing
Imports System.Collections
Imports System.Reflection

Public Class FinalInvoiceCreateDetail
    Inherits System.Web.UI.Page


#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    'Dim paramSupplier As String
    Dim paramaffiliate As String


    'parameter
    Dim pDeliverydate As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pSuratjalanNo As String
    Dim pDeliveryCode As String
    Dim pDeliveryName As String
    Dim pDriverName As String
    Dim pDriverContact As String
    Dim pNoPol As String
    Dim pJenisArmada As String
    Dim pPO As String
    Dim pKanban As String
    Dim pSupplier As String
    Dim pSupplierName As String
    Dim pStatus As Boolean
    Dim pSuratJalan As String

    Dim errorBatch As Boolean
    Dim CartonQty As Integer
#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                up_FillCombo()
                'If Not IsNothing(Request.QueryString("prm")) Then
                Session("MenuDesc") = "FINAL INVOICE CREATE DETAIL"

                If Session("POList") <> "" Then
                    param = Session("POList").ToString()
                ElseIf Session("TampungDelivery") <> "" Then
                    param = Session("TampungDelivery").ToString()
                ElseIf Session("PackingParam") <> "" Then
                    param = Session("PackingParam").ToString
                Else
                    If IsNothing(Request.QueryString("prm")) = True Then
                        lblerrmessage.Text = ""
                        Exit Sub
                    End If
                    param = Request.QueryString("prm").ToString
                End If

                Session("PrintParam") = param
                If param = "  'back'" Then
                    btnsubmenu.Text = "BACK"
                Else
                    If pStatus = False Then
                        pDeliverydate = IIf(IsNothing(Split(param, "|")(0)) = True, "", Split(param, "|")(0))
                        pAffiliateCode = IIf(IsNothing(Split(param, "|")(1)) = True, "", Split(param, "|")(1))
                        pAffiliateName = IIf(IsNothing(Split(param, "|")(2)) = True, "", Split(param, "|")(2))
                        pSuratjalanNo = IIf(IsNothing(Split(param, "|")(3)) = True, "", Split(param, "|")(3))
                        pDeliveryCode = IIf(IsNothing(Split(param, "|")(4)) = True, "", Split(param, "|")(4))
                        pDeliveryName = IIf(IsNothing(Split(param, "|")(5)) = True, "", Split(param, "|")(5))
                        pDriverName = IIf(IsNothing(Split(param, "|")(6)) = True, "", Split(param, "|")(6))
                        pDriverContact = IIf(IsNothing(Split(param, "|")(7)) = True, "", Split(param, "|")(7))
                        pNoPol = IIf(IsNothing(Split(param, "|")(8)) = True, "", Split(param, "|")(8))
                        pJenisArmada = IIf(IsNothing(Split(param, "|")(9)) = True, "", Split(param, "|")(9))
                        pPO = IIf(IsNothing(Split(param, "|")(10)) = True, "", Split(param, "|")(10))
                        pKanban = IIf(IsNothing(Split(param, "|")(11)) = True, "", Split(param, "|")(11))
                        pSupplier = IIf(IsNothing(Split(param, "|")(12)) = True, "", Split(param, "|")(12))
                        pSupplierName = IIf(IsNothing(Split(param, "|")(13)) = True, "", Split(param, "|")(13))
                        pSuratJalan = Trim(IIf(IsNothing(Split(param, "|")(14)) = True, "", Split(param, "|")(14)))

                        If Session("POList") <> "" Then pKanban = Session("KanbanList")

                        If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"
                        If Trim(pDeliverydate) = "01 Jan 1900" Then pDeliverydate = Format(Now, "dd MMM yyyy")
                        txtdeliverydate.Text = pDeliverydate
                        dtInvoiceDate.Text = pDeliverydate
                        txtaffiliatecode.Text = pAffiliateCode
                        txtaffiliatename.Text = pAffiliateName
                        txtdeliverylocationCode.Text = pDeliveryCode
                        txtdeliverylocationName.Text = pDeliveryName
                        txtSupplierCode.Text = pSupplier
                        txtSupplierName.Text = pSupplierName
                        Session("sSuppID") = pSupplier
                        Session("AFF") = pAffiliateCode

                        txttotalbox.Text = uf_SumQty(pSuratJalan, pAffiliateCode)
                        up_HeaderLoad(pSuratJalan, "", pAffiliateCode)
                        pStatus = True

                        If pSuratjalanNo <> "" Then
                            txtsuratjalanno.Text = pSuratjalanNo
                            txtdrivername.Text = pDriverName
                            txtdrivercontact.Text = pDriverContact
                            txtnopol.Text = pNoPol
                            txtjenisarmada.Text = pJenisArmada
                            Call up_IsiInvoice(pSuratjalanNo)
                        ElseIf Session("TampungDelivery") <> "" Then
                            Call up_IsiInvoice(Session("Sj"))
                        End If
                        'txttotalbox.Text = Format(pkanbandate, "dd MMM yyyy")
                        'paramDT1 = pdt1
                        'paramDT2 = pdt2
                        'paramaffiliate = pcboaffiliate
                        'paramSupplier = ptxtsupplierID

                        'Call fillHeader()
                        Call up_GridLoad(pPO, pKanban, pSuratJalan)
                        Session("PO") = pPO
                        Session("Kanban") = pKanban
                        Session("Sj") = pSuratJalan
                        'Session("TampungDelivery") = param
                        Session("TampungDelivery") = param
                    End If
                End If

                btnsubmenu.Text = "BACK"
                'End If
                'ElseIf IsPostBack Then
                '    up_HeaderLoad(txtsuratjalanno.Text, "", txtaffiliatecode.Text)
            End If
            '===============================================================================

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblerrmessage.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            lblStatus.ForeColor = Color.White
        End Try
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Session.Remove("PO")
        Session.Remove("Kanban")
        Session.Remove("sSuppID")
        Session.Remove("Sj")
        Session.Remove("TampungDelivery")

        Session.Remove("POList")
        Session.Remove("KanbanList")
        Session.Remove("PackingParam")
        Session.Remove("PrintParam")
        Session.Remove("AFF")

        'remove Request.QueryString("prm")
        '-------------------------------------------
        Dim isreadonly As PropertyInfo =
        GetType(System.Collections.Specialized.NameValueCollection).GetProperty("IsReadOnly", BindingFlags.Instance Or BindingFlags.NonPublic)

        ' make collection editable
        isreadonly.SetValue(Me.Request.QueryString, False, Nothing)

        ' remove
        If IsNothing(Me.Request.QueryString("prm")) = False Then
            Me.Request.QueryString.Remove("prm")
        End If
        '-------------------------------------------

        If btnsubmenu.Text = "BACK" Then
            Response.Redirect("~/Invoice/FinalInvoice.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If

    End Sub

    Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
        Dim ls_MsgID As String = ""
        Dim ls_sql As String = ""
        Dim iRow As Integer = 0
        Dim ls_PIC As String = Trim(Session("UserID").ToString)


        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)
        Dim ls_SupplierID As String = Session("sSuppID")
        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)
        Dim ls_DeliveryDate As Date = Trim(txtdeliverydate.Text)
        ls_DeliveryDate = Format(CDate(ls_DeliveryDate), "yyyy-MM-dd")
        Dim ls_JenisArmada As String = Trim(txtjenisarmada.Text)
        Dim ls_DriverName As String = Trim(txtdrivername.Text)
        Dim ls_DriverContact As String = Trim(txtdrivercontact.Text)
        Dim ls_NoPol As String = Trim(txtnopol.Text)
        Dim ls_TotalBox As String = Trim(txttotalbox.Text)
        Dim ls_InvoiceNo As String = Trim(txtInvoiceNo.Text)

        Dim ls_PONO As String
        Dim ls_KanbanNo As String
        Dim ls_PartNo As String
        Dim ls_PartNos As String = ""
        Dim ls_DOqty As String
        Dim ls_CartonNo As String
        Dim ls_CartonQty As String
        Dim ls_Active As String = ""
        Dim ls_combination As String = ""


        If Grid.VisibleRowCount = 0 Then Exit Sub
        'Pertama view blum ada isinya, kemudian, di union table dengan yang akan ditambahkan
        'Add Row
        For iRow = 0 To e.UpdateValues.Count - 1
            ls_Active = (e.UpdateValues(iRow).NewValues("AllowAccess").ToString())
            ls_PONO = Trim(e.UpdateValues(iRow).OldValues("colponos").ToString())
            Session("PONO") = ls_PONO
            ls_KanbanNo = Trim(e.UpdateValues(iRow).OldValues("colkanbannos").ToString())
            Session("KanbanNO") = ls_KanbanNo
            ls_PartNo = Trim(e.UpdateValues(iRow).OldValues("colstsDO").ToString())
            Session("PartNo") = ls_PartNo

            If ls_combination = "" Then
                ls_combination = "'" + ls_PONO + ls_KanbanNo + ls_PartNo + "'"
            Else
                ls_combination = ls_combination + ",'" + ls_PONO + ls_KanbanNo + ls_PartNo + "'"
            End If

            If lblStatus.Text = "addCarton" Then
                Call up_ExistCartonQty(ls_Sjno, ls_SupplierID, ls_AffiliateID, e.UpdateValues(iRow).OldValues("colponos"), e.UpdateValues(iRow).OldValues("colkanbannos"), e.UpdateValues(iRow).OldValues("colstsDO"))
                If (e.UpdateValues(iRow).OldValues("colpasideliveryqty") / e.UpdateValues(iRow).OldValues("colQtyBox")) >= CartonQty And CartonQty <> 0 Then
                    Call clsMsg.DisplayMessage(lblerrmessage, "6029", clsMessage.MsgType.ErrorMessage)
                    Session("YA010IsSubmit") = lblerrmessage.Text
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("errorBatch") = "True"
                    Exit Sub
                End If
                Session("errorBatch") = ""
                If e.UpdateValues.Count > 1 Then
                    If iRow = 0 Then
                        ls_PartNos = "'" & ls_PartNo & "'"
                    ElseIf iRow > 0 Then
                        'ls_PartNos = Replace(ls_PartNos, "'", "")
                        ls_PartNos = ls_PartNos & ",'" & Trim(e.UpdateValues(iRow).OldValues("colstsDO").ToString()) & "'"
                        Session("PartNos") = ls_PartNos
                        Session("Combination") = ls_combination
                        If e.UpdateValues.Count - 1 = iRow Then
                            up_AddRow(ls_PONO, ls_KanbanNo, ls_Sjno, ls_PartNos, ls_combination)
                        End If
                    End If
                ElseIf e.UpdateValues.Count = 1 Then
                    ls_PartNos = "'" & Trim(e.UpdateValues(iRow).OldValues("colstsDO").ToString()) & "'"
                    Session("PartNos") = ls_PartNos
                    Session("Combination") = ls_combination
                    up_AddRow(ls_PONO, ls_KanbanNo, ls_Sjno, ls_PartNos, ls_combination)
                End If
            ElseIf lblStatus.Text = "deleteData" Then
                ls_CartonNo = e.UpdateValues(iRow).NewValues("colcartonno").ToString()
                up_Delete(ls_Sjno, ls_KanbanNo, ls_PartNo, ls_CartonNo)
                Session("deletesukses") = "deleteDataSukses"
            ElseIf lblStatus.Text = "saveData" Then
                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    ls_DOqty = Trim(CDbl(IIf(e.UpdateValues(iRow).OldValues("colpasideliveryqty").ToString(), e.UpdateValues(iRow).OldValues("colpasideliveryqty").ToString(), 0)))
                    ls_CartonNo = e.UpdateValues(iRow).NewValues("colcartonno").ToString()
                    ls_CartonQty = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colcartonqty").ToString(), e.UpdateValues(iRow).NewValues("colcartonqty").ToString(), 0)))
                    ls_SupplierID = e.UpdateValues(iRow).OldValues("colsupp").ToString()
                    Dim ls_SuratJalanSupp As String = e.UpdateValues(iRow).OldValues("colSJSupp").ToString()
                    Call up_ExistCartonQty(ls_Sjno, ls_SupplierID, ls_AffiliateID, e.UpdateValues(iRow).OldValues("colponos"), e.UpdateValues(iRow).OldValues("colkanbannos"), e.UpdateValues(iRow).OldValues("colstsDO"))
                    If (e.UpdateValues(iRow).OldValues("colpasideliveryqty") / e.UpdateValues(iRow).OldValues("colQtyBox")) >= CartonQty And CartonQty <> 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "6029", clsMessage.MsgType.ErrorMessage)
                        Session("YA010IsSubmit") = lblerrmessage.Text
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        'errorBatch = True
                        Exit Sub
                    End If
                    If ls_CartonNo <> "" Then
                        ls_sql = Update_Detail(ls_Sjno, ls_SupplierID, ls_AffiliateID, ls_PONO, ls_KanbanNo, ls_PartNo, ls_DOqty, ls_CartonNo, ls_CartonQty, ls_SuratJalanSupp)
                        Dim sqlComm2 As New SqlCommand(ls_sql, sqlConn)
                        sqlComm2.ExecuteNonQuery()
                        sqlComm2.Dispose()
                        Session("savebatchsukses") = "savebatchsukses"
                    End If
                    ls_MsgID = "1001"
                    Call clsMsg.DisplayMessage(lblerrmessage, ls_MsgID, clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("YA010IsSubmit") = lblerrmessage.Text
                    sqlConn.Close()
                End Using
            End If
        Next iRow
    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Grid.JSProperties("cpMessage") = Session("YA010IsSubmit")
        Grid.JSProperties("cpIsDel") = ""

        Select Case pAction
            Case "status"
                Dim pStatus As String = Split(e.Parameters, "|")(1)
                lblStatus.Text = pStatus
            Case "gridload"
                'Call up_SaveAllDetail(Session("PO"), Session("Kanban"), Session("Sj"))
                Call up_GridLoad(Session("PO"), Session("Kanban"), Session("Sj"))
                If Grid.VisibleRowCount = 0 Then
                    Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("YA010IsSubmit") = lblerrmessage.Text
                End If
            Case "addrow"
                If Session("errorBatch") = "True" Then
                    Call up_GridLoad(Session("PO"), Session("Kanban"), Session("Sj"))
                Else
                    Call up_AddRow(Session("PONO"), Session("KanbanNO"), Session("Sj"), Session("PartNos"), Session("Combination"))
                End If
            Case "SendEDI"
                If uf_validate() Then
                    If uf_Approve() = 1 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2005", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        Grid.JSProperties("cpButton") = "1"
                    Else
                        Grid.JSProperties("cpMessage") = ""
                        Grid.JSProperties("cpButton") = "0"
                    End If
                Else
                    Call clsMsg.DisplayMessage(lblerrmessage, "8002", clsMessage.MsgType.ErrorMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("YA010IsSubmit") = lblerrmessage.Text
                    Grid.JSProperties("cpButton") = "0"
                End If
            Case "saveDataMaster"
                lblStatus.Text = "saveData"
                Dim ls_SupplierID As String = Session("sSuppID")
                Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)
                Dim ls_DeliveryDate As Date = Trim(txtdeliverydate.Text)
                ls_DeliveryDate = Format(CDate(ls_DeliveryDate), "yyyy-MM-dd")
                Dim ls_JenisArmada As String = Trim(txtjenisarmada.Text)
                Dim ls_DriverName As String = Trim(txtdrivername.Text)
                Dim ls_DriverContact As String = Trim(txtdrivercontact.Text)
                Dim ls_NoPol As String = Trim(txtnopol.Text)
                Dim ls_TotalBox As String = Trim(txttotalbox.Text)
                Dim ls_InvoiceNo As String = Trim(txtInvoiceNo.Text)
                Dim ls_InvoiceDate As String = Trim(dtInvoiceDate.Value)

                Dim ls_FromDel As String = Split(e.Parameters, "|")(1)
                Dim ls_ToDel As String = Split(e.Parameters, "|")(2)
                Dim ls_Insu As String = Split(e.Parameters, "|")(3)
                Dim ls_ViaDel As String = Split(e.Parameters, "|")(4)
                Dim ls_AboutDel As String = Split(e.Parameters, "|")(5)
                Dim ls_Privilege As String = Split(e.Parameters, "|")(6)
                Dim ls_Vessel As String = Split(e.Parameters, "|")(7)
                Dim ls_AWB As String = Split(e.Parameters, "|")(8)
                Dim ls_PayTerms As String = Split(e.Parameters, "|")(9)
                Dim ls_OnAbout As String = Split(e.Parameters, "|")(10)
                Dim ls_ContainerNo As String = Split(e.Parameters, "|")(11)
                Dim ls_Remarks As String = Split(e.Parameters, "|")(12)
                Dim ls_Place As String = Split(e.Parameters, "|")(13)
                Dim ls_Commercial As String = Split(e.Parameters, "|")(14)


                Call up_SaveMaster(Session("Sj"), "", ls_AffiliateID, ls_DeliveryDate, Session("UserID"), ls_JenisArmada, ls_DriverName, ls_DriverContact, ls_NoPol, ls_TotalBox, ls_InvoiceNo, ls_InvoiceDate,
                                   ls_FromDel, ls_ToDel, ls_Insu, ls_ViaDel, ls_AboutDel, ls_Privilege, ls_Vessel, ls_AWB, ls_PayTerms, ls_OnAbout, ls_ContainerNo, ls_Remarks, ls_Place, ls_Commercial)
                If Session("savebatchsukses") = "" Then
                    Call up_SaveDetail(Session("Sj"), Trim(ls_SupplierID), ls_AffiliateID)
                End If

                Call up_HeaderLoad(Session("Sj"), "", Trim(txtaffiliatecode.Text))
                Call up_GridLoad(Session("PO"), Session("Kanban"), Session("Sj"))
                Call clsMsg.DisplayMessage(lblerrmessage, "1001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                Session("YA010IsSubmit") = lblerrmessage.Text
            Case "gridloadAfter"
                'Call up_SaveAllDetail(Session("PO"), Session("Kanban"), Session("Sj"))
                Call up_GridLoad(Session("PO"), Session("Kanban"), Session("Sj"))
            Case "Delete"
                'Call up_Delete(txtsuratjalanno.Text)
                'lblStatus.Text = "deleteData"
                Call up_GridLoad(Session("PO"), Session("Kanban"), Session("Sj"))
                If Session("deletesukses") = "deleteDataSukses" Then
                    Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("deletesukses") = ""
                End If
            Case "DeletePL"
                'Call up_Delete(txtsuratjalanno.Text)
                lblStatus.Text = "deleteDataPL"
                Call up_DeletePL(Trim(txtsuratjalanno.Text))
                Call up_HeaderLoad(Session("Sj"), "", Trim(txtaffiliatecode.Text))
                Call up_GridLoad(Session("PO"), Session("Kanban"), Session("Sj"))

                Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
        End Select
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        lblStatus.ForeColor = Color.White
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Session("PrintSJ") = txtsuratjalanno.Text
        Session("PrintAffID") = txtaffiliatecode.Text
        Session("PrintSuppID") = txtSupplierCode.Text

        Response.Redirect("~/Invoice/FinalInvoiceViewReport.aspx")
    End Sub

    'Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
    '    If uf_Approve() = 1 Then
    '        ButtonApprove.JSProperties("cpMessage") = "[2004] Send E.D.I Successfully"
    '        ButtonApprove.JSProperties("cpButton") = "1"
    '    Else
    '        ButtonApprove.JSProperties("cpMessage") = ""
    '        ButtonApprove.JSProperties("cpButton") = "0"
    '    End If
    'End Sub

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        If e.GetValue("colno") = "" Then
            If e.DataColumn.FieldName = "colQtyBox" Or e.DataColumn.FieldName = "colsuppdelqty" Or e.DataColumn.FieldName = "colpasigoodrec" Or e.DataColumn.FieldName = "colpasidefectrec" Or e.DataColumn.FieldName = "colpasiremaining" Then
                e.Cell.Text = ""
            End If
            If Not (e.DataColumn.FieldName = "AllowAccess" Or e.DataColumn.FieldName = "colcartonno" Or e.DataColumn.FieldName = "colcartonqty") Then
                e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
            Else
                e.Cell.BackColor = Color.White
            End If
        End If

        If e.DataColumn.FieldName = "colremainingdelqty" Then
            If (CDbl(e.GetValue("colpasideliveryqty")) > e.GetValue("colpasideliveryqty")) Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If

        'Delivery Qty Not save
        If e.DataColumn.FieldName = "colpasideliveryqty" Then
            If (Trim(e.GetValue("colstsDO")) = "") Then
                e.Cell.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub Grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles Grid.HtmlRowPrepared
        Try
            Dim getRowValues As String = e.GetValue("colpasideliveryqty")
            If Not IsNothing(getRowValues) Then
                If getRowValues.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If

        Catch ex As Exception
            'Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'Session("E01Msg") = lblerrmessage.Text
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet
        Dim ls_SQL As String

        'Combo Affiliate
        With cboFrom
            ls_SQL = "FinalInvoiceCreate_FillCombo"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlcmd As New SqlCommand(ls_SQL, sqlConn)
                    sqlcmd.CommandType = CommandType.StoredProcedure

                    sqlDA = New SqlDataAdapter(sqlcmd)
                    ds = New DataSet
                    sqlDA.Fill(ds)

                    .Items.Clear()
                    .Columns.Clear()
                    .DataSource = ds.Tables(0)
                    .Columns.Add("Descript")
                    .Columns(0).Width = 150
                    .DataBind()
                End Using
            End Using
        End With

        'Combo Commercial
        With cboCommercial
            ls_SQL = "FinalInvoiceCreate_FillCombo"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlcmd As New SqlCommand(ls_SQL, sqlConn)
                    sqlcmd.CommandType = CommandType.StoredProcedure

                    sqlDA = New SqlDataAdapter(sqlcmd)
                    ds = New DataSet
                    sqlDA.Fill(ds)

                    .Items.Clear()
                    .Columns.Clear()
                    .DataSource = ds.Tables(1)
                    .Columns.Add("Descript")
                    .Columns(0).Width = 100
                    .DataBind()
                End Using
            End Using
        End With

    End Sub

    Private Function uf_validate() As Boolean
        Dim dt As New DataTable
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Dim SqlComm As New SqlCommand("sp_uf_validateSendEDI", sqlConn)
                SqlComm.CommandType = CommandType.StoredProcedure
                SqlComm.Parameters.AddWithValue("AffiliateID", txtaffiliatecode.Text)
                SqlComm.Parameters.AddWithValue("SuratJalan", txtsuratjalanno.Text)
                SqlComm.ExecuteNonQuery()

                Dim da As New SqlDataAdapter(SqlComm)
                da.Fill(dt)
                If dt.Rows.Count > 0 Then
                    Dim a, b As String
                    a = dt.Rows(0)("QtyDoPasi").ToString()
                    b = dt.Rows(0)("QtyPLPasi").ToString()

                    If dt.Rows(0)("QtyDoPasi").ToString() = dt.Rows(0)("QtyPLPasi").ToString() Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If

                SqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function uf_Approve() As Integer
        Dim ls_sql As String
        Dim x As Integer
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    ls_sql = " update PLPASI_Master set EDICls = '1'" & vbCrLf &
                                " WHERE AffiliateID = '" & txtaffiliatecode.Text & "' and SuratJalanNo = '" & txtsuratjalanno.Text & "'" & vbCrLf

                    Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            Return x
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Sub up_HeaderLoad(ByVal pSuratJalan As String, ByVal pSupplierID As String, ByVal pAffiliateID As String)
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "FinalInvoiceCreate_HeaderLoad"
            Using sqlcmd As New SqlCommand(ls_SQL, sqlConn)
                sqlcmd.CommandType = CommandType.StoredProcedure
                sqlcmd.Parameters.AddWithValue("SuratJalan", Trim(pSuratjalanNo))
                sqlcmd.Parameters.AddWithValue("AffiliateID", Trim(pAffiliateID))
                sqlcmd.Parameters.AddWithValue("SupplierID", Trim(pSupplierID))

                Dim sqlDA = New SqlDataAdapter(sqlcmd)
                Dim ds = New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    cboFrom.Value = Trim(ds.Tables(0).Rows(0)("FromDeliveryNew") & "")
                    cboCommercial.Value = Trim(ds.Tables(0).Rows(0)("Commercial") & "")
                    txtTo.Text = Trim(ds.Tables(0).Rows(0)("ToDelivery") & "")
                    txtInsurance.Text = Trim(ds.Tables(0).Rows(0)("InsurancePolicy") & "")
                    txtVia.Text = Trim(ds.Tables(0).Rows(0)("ViaDelivery") & "")
                    txtAbout.Text = Trim(ds.Tables(0).Rows(0)("AboutDelivery") & "")
                    txtPrivilege.Text = Trim(ds.Tables(0).Rows(0)("Privilege") & "")
                    txtVessel.Text = Trim(ds.Tables(0).Rows(0)("Vessel") & "")
                    txtAwb.Text = Trim(ds.Tables(0).Rows(0)("AWBBLNo") & "")
                    txtPaymentTerms.Text = Trim(ds.Tables(0).Rows(0)("PaymentTerms") & "")
                    txtOn.Text = Trim(ds.Tables(0).Rows(0)("OnAbout") & "")
                    txtContainerNo.Text = Trim(ds.Tables(0).Rows(0)("ContainerNo") & "")
                    txtRemarks.Text = Trim(ds.Tables(0).Rows(0)("Remarks") & "")
                    txtPlace.Text = Trim(ds.Tables(0).Rows(0)("Place") & "")

                    If IsDBNull(ds.Tables(0).Rows(0)("EDICls")) Then
                        btnSendEDI.Enabled = True
                    Else
                        If ds.Tables(0).Rows(0)("EDICls") = "1" Or ds.Tables(0).Rows(0)("EDICls") = "2" Then
                            btnSendEDI.Enabled = False
                        Else
                            btnSendEDI.Enabled = True
                        End If
                    End If

                    Grid.JSProperties("cpFromDelivery") = ds.Tables(0).Rows(0).Item("FromDelivery")
                    Grid.JSProperties("cpToDelivery") = ds.Tables(0).Rows(0).Item("ToDelivery")
                    Grid.JSProperties("cpInsurancePolicy") = ds.Tables(0).Rows(0).Item("InsurancePolicy")
                    Grid.JSProperties("cpViaDelivery") = ds.Tables(0).Rows(0).Item("ViaDelivery")
                    Grid.JSProperties("cpAboutDelivery") = ds.Tables(0).Rows(0).Item("AboutDelivery")
                    Grid.JSProperties("cpPrivilege") = ds.Tables(0).Rows(0).Item("Privilege")
                    Grid.JSProperties("cpVessel") = ds.Tables(0).Rows(0).Item("Vessel")
                    Grid.JSProperties("cpAWBBLNo") = ds.Tables(0).Rows(0).Item("AWBBLNo")
                    Grid.JSProperties("cpPaymentTerms") = ds.Tables(0).Rows(0).Item("PaymentTerms")
                    Grid.JSProperties("cpOnAbout") = ds.Tables(0).Rows(0).Item("OnAbout")
                    Grid.JSProperties("cpContainerNo") = ds.Tables(0).Rows(0).Item("ContainerNo")
                    Grid.JSProperties("cpRemarks") = ds.Tables(0).Rows(0).Item("Remarks")
                    Grid.JSProperties("cpPlace") = ds.Tables(0).Rows(0).Item("Place")
                Else
                    Grid.JSProperties("cpFromDelivery") = ""
                    Grid.JSProperties("cpToDelivery") = ""
                    Grid.JSProperties("cpInsurancePolicy") = ""
                    Grid.JSProperties("cpViaDelivery") = ""
                    Grid.JSProperties("cpAboutDelivery") = ""
                    Grid.JSProperties("cpPrivilege") = ""
                    Grid.JSProperties("cpVessel") = ""
                    Grid.JSProperties("cpAWBBLNo") = ""
                    Grid.JSProperties("cpPaymentTerms") = ""
                    Grid.JSProperties("cpOnAbout") = ""
                    Grid.JSProperties("cpContainerNo") = ""
                    Grid.JSProperties("cpRemarks") = ""
                    Grid.JSProperties("cpPlace") = ""
                End If
                sqlConn.Close()

            End Using

        End Using
    End Sub

    Private Sub up_GridLoadOriginal(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = "    --Part 1 ada data, Part 2 ga ada data " & vbCrLf &
                  "    IF EXISTS (SELECT * FROM dbo.PLPASI_Detail WHERE SuratJalanNo = '" & Trim(pSJ) & "' AND PONo = " & Trim(pPO) & " AND KanbanNo =" & Trim(pKanban) & " --AND PartNo IN ('7009-2190-02','7009-2191-02') " & vbCrLf &
                  "    ) " & vbCrLf &
                  "    BEGIN " & vbCrLf &
                  "    --AddRow " & vbCrLf &
                  " SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colstsDO)) FROM ( " & vbCrLf &
                  " SELECT '0'AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,   " & vbCrLf &
                  "           colpono = POM.PONo , " & vbCrLf &
                  "           colponos = POM.PONo,   " & vbCrLf &
                  "           colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'   " & vbCrLf &
                  "                              ELSE 'YES'   "

            ls_SQL = ls_SQL + "                         END ,   " & vbCrLf &
                              "           colkanbanno = CASE WHEN POD.KanbanCls = '0' THEN '-'   " & vbCrLf &
                              "                              ELSE ISNULL(KD.KanbanNo, '')   " & vbCrLf &
                              "                         END ,   " & vbCrLf &
                              "           colkanbannos = CASE WHEN POD.KanbanCls = '0' THEN '-'   " & vbCrLf &
                              "                              ELSE ISNULL(KD.KanbanNo, '')   " & vbCrLf &
                              "                         END ,                 " & vbCrLf &
                              "           colpartno = POD.PartNo ,   " & vbCrLf &
                              "           colpartname = MP.PartName ,   " & vbCrLf &
                              "           coluom = UC.Description ,   " & vbCrLf &
                              "           colCls = UC.unitcls ,  " & vbCrLf

            ls_SQL = ls_SQL + "           colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,   " & vbCrLf &
                              "           colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,   " & vbCrLf &
                              "           colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),   " & vbCrLf &
                              "           colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),   " & vbCrLf &
                              "           colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)   " & vbCrLf &
                              "                                              - ( ISNULL(PRD.GoodRecQty, 0)   " & vbCrLf &
                              "                                                  + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,   " & vbCrLf &
                              "           colpasideliveryqty = CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END ,   " & vbCrLf &
                              "           colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,   " & vbCrLf &
                              "           coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)   " & vbCrLf &
                              "                            WHEN 0 THEN 0   " & vbCrLf

            ls_SQL = ls_SQL + "                            ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)   " & vbCrLf &
                              "                          END,0)),   " & vbCrLf &
                              "           colstsDO = ISNULL(PDD.PartNo,'') ,  " & vbCrLf &
                              "           colcartonno = '',  " & vbCrLf &
                              "           colcartonqty = '',  " & vbCrLf &
                              "           sortData = 0 " & vbCrLf &
                              "    FROM   dbo.PO_Master POM   " & vbCrLf &
                              "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf &
                              "                                      AND POM.PoNo = POD.PONo   " & vbCrLf &
                              "                                      AND POM.SupplierID = POD.SupplierID   " & vbCrLf &
                              "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   "

            ls_SQL = ls_SQL + "                                             AND KD.PoNo = POD.PONo   " & vbCrLf &
                              "                                             AND KD.SupplierID = POD.SupplierID   " & vbCrLf &
                              "                                             AND KD.PartNo = POD.PartNo   " & vbCrLf &
                              "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf &
                              "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf &
                              "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf &
                              "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf &
                              "           LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf &
                              "                                                  AND KD.KanbanNo = SDD.KanbanNo   " & vbCrLf &
                              "                                                  AND KD.SupplierID = SDD.SupplierID   " & vbCrLf &
                              "                                                  AND KD.PartNo = SDD.PartNo   " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND KD.PONo = SDD.PONo   " & vbCrLf &
                              "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf &
                              "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf &
                              "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf &
                              "           LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID   " & vbCrLf &
                              "                                                   AND KD.KanbanNo = PRD.KanbanNo   " & vbCrLf &
                              "                                                   AND KD.SupplierID = PRD.SupplierID   " & vbCrLf &
                              "                                                   AND KD.PartNo = PRD.PartNo   " & vbCrLf &
                              "                                                   AND KD.PartNo = PRD.PartNo   " & vbCrLf &
                              "                                                   AND SDM.SuratJalanno = PRD.SuratJalanNo " & vbCrLf &
                              "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   " & vbCrLf &
                              "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf

            ls_SQL = ls_SQL + "                                                   AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf &
                              "           LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID   " & vbCrLf &
                              "                                              AND KD.KanbanNo = PDD.KanbanNo   " & vbCrLf &
                              "                                              AND KD.SupplierID = PDD.SupplierID   " & vbCrLf &
                              "                                              AND KD.PartNo = PDD.PartNo   " & vbCrLf &
                              "                                              AND KD.PoNo = PDD.PoNo   " & vbCrLf &
                              "                                              AND SDM.SuratJalanNoSupplier = PDD.SuratJalanNo " & vbCrLf &
                              "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID   " & vbCrLf &
                              "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf &
                              "                                              AND PDD.SupplierID = PDM.SupplierID   " & vbCrLf &
                              "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo   " & vbCrLf &
                              "           LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf

            ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID   " & vbCrLf &
                              "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf &
                              "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "    --WHERE PDD.SuratJalanNo = 'Surat Jalan PASITrie'  --AND PDD.PartNo IN ('7009-2190-02','7009-2191-02') " & vbCrLf &
                              " UNION all " & vbCrLf &
                              " SELECT '0'AllowAccess,colno =  '' ,   " & vbCrLf &
                              "        colpono = '' ,    " & vbCrLf &
                              "        colponos = POM.PONo,  " & vbCrLf &
                              "        colpokanban = '' ,    " & vbCrLf &
                              "        colkanbanno = '' ,  " & vbCrLf &
                              "        colkanbannos = CASE WHEN POD.KanbanCls = '0' THEN '-'    " & vbCrLf

            ls_SQL = ls_SQL + "                           ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf &
                              "                      END ,      " & vbCrLf &
                              "        colpartno = '' ,    " & vbCrLf &
                              "        colpartname = '' ,     " & vbCrLf &
                              "        coluom = '' ,    " & vbCrLf &
                              "        colCls = '' ,   " & vbCrLf &
                              "        colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,    " & vbCrLf &
                              "        colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,    " & vbCrLf &
                              "        colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),    " & vbCrLf &
                              "        colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),    " & vbCrLf &
                              "        colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)    " & vbCrLf

            ls_SQL = ls_SQL + "                                           - ( ISNULL(PRD.GoodRecQty, 0)    " & vbCrLf &
                              "                                               + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,    " & vbCrLf &
                              "        colpasideliveryqty = CASE ISNULL(PLD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.DOQty,0)))) END ,    " & vbCrLf &
                              "        colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,    " & vbCrLf &
                              "        coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)    " & vbCrLf &
                              "                         WHEN 0 THEN 0    " & vbCrLf &
                              "                         ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)    " & vbCrLf &
                              "                       END,0)),    " & vbCrLf &
                              "        colstsDO = ISNULL(PDD.PartNo,'') ,   " & vbCrLf &
                              "        colcartonno = ISNULL(PLD.CartonNo,'') ,   " & vbCrLf &
                              "        colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),   " & vbCrLf

            ls_SQL = ls_SQL + "        sortData = 1 " & vbCrLf &
                              " FROM   dbo.PO_Master POM    " & vbCrLf &
                              "        LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf &
                              "                                   AND POM.PoNo = POD.PONo    " & vbCrLf &
                              "                                   AND POM.SupplierID = POD.SupplierID    " & vbCrLf &
                              "        LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID    " & vbCrLf &
                              "                                          AND KD.PoNo = POD.PONo    " & vbCrLf &
                              "                                          AND KD.SupplierID = POD.SupplierID    " & vbCrLf &
                              "                                          AND KD.PartNo = POD.PartNo    " & vbCrLf &
                              "        LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf &
                              "                                          AND KD.KanbanNo = KM.KanbanNo    " & vbCrLf

            ls_SQL = ls_SQL + "                                          AND KD.SupplierID = KM.SupplierID    " & vbCrLf &
                              "                                          AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf &
                              "        LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf &
                              "                                               AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf &
                              " 							                    AND KD.SupplierID = SDD.SupplierID    " & vbCrLf &
                              "                                               AND KD.PartNo = SDD.PartNo    " & vbCrLf &
                              "                                               AND KD.PONo = SDD.PONo    " & vbCrLf &
                              "        LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID    " & vbCrLf &
                              "                                               AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf &
                              "                                               AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf &
                              "        LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID    " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.KanbanNo = PRD.KanbanNo    " & vbCrLf &
                              "                                                AND KD.SupplierID = PRD.SupplierID    " & vbCrLf &
                              "                                                AND KD.PartNo = PRD.PartNo    " & vbCrLf &
                              "                                                AND KD.PartNo = PRD.PartNo    " & vbCrLf &
                              "                                                AND SDM.SuratJalanno = PRD.SuratJalanNo " & vbCrLf &
                              "        LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf &
                              "                                                AND PRM.SuratJalanNo = PRD.SuratJalanNo    " & vbCrLf &
                              "                                                AND PRM.SupplierID = PRD.SupplierID    " & vbCrLf &
                              "        LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID    " & vbCrLf &
                              "                                           AND KD.KanbanNo = PDD.KanbanNo    " & vbCrLf &
                              "                                           AND KD.SupplierID = PDD.SupplierID    " & vbCrLf &
                              "                                           AND KD.PartNo = PDD.PartNo    " & vbCrLf &
                              "                                           AND SDM.SuratJalanNoSupplier = PDD.SuratJalanNo " & vbCrLf

            ls_SQL = ls_SQL + "                                           AND KD.PoNo = PDD.PoNo    " & vbCrLf &
                              "        LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf &
                              "                                           AND PDD.SuratJalanNo = PDM.SuratJalanNo    " & vbCrLf &
                              "                                           AND PDD.SupplierID = PDM.SupplierID    " & vbCrLf &
                              "       LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID    " & vbCrLf &
                              "                               AND KD.KanbanNo = PLD.KanbanNo    " & vbCrLf &
                              "                               AND KD.SupplierID = PLD.SupplierID    " & vbCrLf &
                              "                               AND KD.PartNo = PLD.PartNo    " & vbCrLf &
                              "                               AND KD.PoNo = PLD.PoNo                                       " & vbCrLf &
                              "        LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "        LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf

            ls_SQL = ls_SQL + "        LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf &
                              "        LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf &
                              "        LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "  --WHERE PLD.SuratJalanNo = 'Surat Jalan PASITrie' --AND PLD.PartNo IN ('7009-2190-02','7009-2191-02') " & vbCrLf &
                              "  ) data " & vbCrLf &
                              "  ORDER BY colstsDO ASC, sortData ASC " & vbCrLf &
                              "   END " & vbCrLf &
                              "   ELSE " & vbCrLf &
                              "   BEGIN " & vbCrLf &
                              " SELECT '0'AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,    " & vbCrLf &
                              "            colpono = POM.PONo ,  " & vbCrLf &
                              "            colponos = POM.PONo,    " & vbCrLf &
                              "            colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'    " & vbCrLf &
                              "                               ELSE 'YES'                            END ,    " & vbCrLf &
                              "            colkanbanno = CASE WHEN POD.KanbanCls = '0' THEN '-'    " & vbCrLf &
                              "                               ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf &
                              "                          END ,    " & vbCrLf &
                              "            colkanbannos = CASE WHEN POD.KanbanCls = '0' THEN '-'    " & vbCrLf &
                              "                               ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf &
                              "                          END ,                  " & vbCrLf

            ls_SQL = ls_SQL + "            colpartno = POD.PartNo ,    " & vbCrLf &
                              "            colpartname = MP.PartName ,    " & vbCrLf &
                              "            coluom = UC.Description ,    " & vbCrLf &
                              "            colCls = UC.unitcls ,             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,    " & vbCrLf &
                              "            colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,    " & vbCrLf &
                              "            colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),    " & vbCrLf &
                              "            colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),    " & vbCrLf &
                              "            colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)    " & vbCrLf &
                              "                                               - ( ISNULL(PRD.GoodRecQty, 0)    " & vbCrLf &
                              "                                                   + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,    " & vbCrLf &
                              "            colpasideliveryqty = CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END ,    " & vbCrLf

            ls_SQL = ls_SQL + "            colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,    " & vbCrLf &
                              "            coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)    " & vbCrLf &
                              "                             WHEN 0 THEN 0                               ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)    " & vbCrLf &
                              "                           END,0)),    " & vbCrLf &
                              "            colstsDO = ISNULL(PDD.PartNo,'') ,   " & vbCrLf &
                              "            colcartonno = '',   " & vbCrLf &
                              "            colcartonqty = '',   " & vbCrLf &
                              "            sortData = 0 ,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) " & vbCrLf &
                              "     FROM   dbo.PO_Master POM    " & vbCrLf &
                              "            LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf &
                              "                                       AND POM.PoNo = POD.PONo    " & vbCrLf

            ls_SQL = ls_SQL + "                                       AND POM.SupplierID = POD.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID                                                AND KD.PoNo = POD.PONo    " & vbCrLf &
                              "                                              AND KD.SupplierID = POD.SupplierID    " & vbCrLf &
                              "                                              AND KD.PartNo = POD.PartNo    " & vbCrLf &
                              "            LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf &
                              "                                              AND KD.KanbanNo = KM.KanbanNo    " & vbCrLf &
                              "                                              AND KD.SupplierID = KM.SupplierID    " & vbCrLf &
                              "                                              AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf &
                              "            LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf &
                              "                                                   AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf &
                              "                                                   AND KD.SupplierID = SDD.SupplierID    " & vbCrLf

            ls_SQL = ls_SQL + "                                                   AND KD.PartNo = SDD.PartNo                                                     AND KD.PONo = SDD.PONo    " & vbCrLf &
                              "            LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID    " & vbCrLf &
                              "                                                   AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf &
                              "                                                   AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID    " & vbCrLf &
                              "                                                    AND KD.KanbanNo = PRD.KanbanNo    " & vbCrLf &
                              "                                                    AND KD.SupplierID = PRD.SupplierID    " & vbCrLf &
                              "                                                    AND KD.PartNo = PRD.PartNo    " & vbCrLf &
                              "                                                    AND KD.PartNo = PRD.PartNo    " & vbCrLf &
                              "                                                    AND SDM.SuratJalanno = PRD.SuratJalanNo " & vbCrLf &
                              "            LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf &
                              "                                                    AND PRM.SuratJalanNo = PRD.SuratJalanNo           " & vbCrLf &
                              "                                                    AND PRM.SupplierID = PRD.SupplierID    " & vbCrLf

            ls_SQL = ls_SQL + "            LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID    " & vbCrLf &
                              "                                               AND KD.KanbanNo = PDD.KanbanNo    " & vbCrLf &
                              "                                               AND KD.SupplierID = PDD.SupplierID    " & vbCrLf &
                              "                                               AND KD.PartNo = PDD.PartNo    " & vbCrLf &
                              "                                               AND KD.PoNo = PDD.PoNo    " & vbCrLf &
                              "                                               AND SDM.SuratJalanNoSupplier = PDD.SuratJalanNo " & vbCrLf &
                              "            LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf &
                              "                                               AND PDD.SuratJalanNo = PDM.SuratJalanNo    " & vbCrLf &
                              "                                               AND PDD.SupplierID = PDM.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "            LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf &
                              "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf &
                              "            LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf

            ls_SQL = ls_SQL + "            LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  "

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL +
                              "    --WHERE PDD.SuratJalanNo = 'Surat Jalan PASITrie' --AND PDD.PartNo IN ('7009-2190-02','7009-2191-02') " & vbCrLf &
                              "    --) data " & vbCrLf &
                              "  ORDER BY colstsDO ASC, sortData ASC " & vbCrLf &
                              "   END "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_GridLoad_(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = " --Part 1 ada data, Part 2 ga ada data  " & vbCrLf &
                  " IF EXISTS (SELECT * FROM dbo.PLPASI_Detail WHERE SuratJalanNo = '" & Trim(pSJ) & "' --AND PONo = " & Trim(pPO) & " )  " & vbCrLf &
                  " ) BEGIN  " & vbCrLf &
                  " 	SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colstsDO)) FROM (  " & vbCrLf &
                  " --header " & vbCrLf &
                  "      select  " & vbCrLf &
                  "      AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colstsDO, colkanbannos, sortData )) ,      " & vbCrLf &
                  "      colpono, colponos, colpokanban, colkanbanno, colkanbannos, colpartno, colpartname, coluom, " & vbCrLf &
                  "      colCls, colQtyBox, colsuppdelqty, colpasigoodrec, colpasidefectrec, colpasiremaining, " & vbCrLf &
                  "      colpasideliveryqty, colremainingdelqty, coldelqtybox, colstsDO, colcartonno, colcartonqty, sortData, colsupp " & vbCrLf &
                  "      from(  " & vbCrLf &
                  " 	   --Delivery Supplier " & vbCrLf &
                  " 	   SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                  "             colpono = POM.PONo ,   " & vbCrLf &
                  "             colponos = POM.PONo,     " & vbCrLf &
                  "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                  "                                ELSE 'YES'                            END ,     " & vbCrLf &
                  "             colkanbanno = ISNULL(KD.KanbanNo, ''),    " & vbCrLf

            ls_SQL = ls_SQL + "             colkanbannos = ISNULL(KD.KanbanNo, ''),    " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf

            ls_SQL = ls_SQL + "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.RecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                                - ( ISNULL(PRD.RecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) ,     " & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.RecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.RecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf

            ls_SQL = ls_SQL + "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0 , colsupp = KD.SupplierID  " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceiveAffiliate_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceiveAffiliate_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf
            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + " 		--Delivery PASI " & vbCrLf &
                              " 		UNION ALL " & vbCrLf &
                              " 		SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                              "             colpono = POM.PONo ,   " & vbCrLf &
                              "             colponos = POM.PONo,     " & vbCrLf &
                              "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf

            ls_SQL = ls_SQL + "                                ELSE 'YES'                            END ,     " & vbCrLf &
                              "             colkanbanno = ISNULL(KD.KanbanNo, ''),     " & vbCrLf &
                              "             colkanbannos = ISNULL(KD.KanbanNo, '')  ,   " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) ,     " & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf

            ls_SQL = ls_SQL + "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0 , colsupp = KD.SupplierID  " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = SDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = SDD.SuratJalanNoSupplier " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON SDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND SDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                --AND SDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "' )header" & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf &
                              " 	-- PackingList yang sudah ada  " & vbCrLf &
                              " 	SELECT distinct '0'AllowAccess,colno =  '' ,    " & vbCrLf &
                              "         colpono = '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "         colponos = POM.PONo,   " & vbCrLf &
                              "         colpokanban = '' ,     " & vbCrLf &
                              "         colkanbanno = '' ,   " & vbCrLf &
                              "         colkanbannos = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf &
                              "         colpartno = '' ,     " & vbCrLf &
                              "         colpartname = '' ,      " & vbCrLf &
                              "         coluom = '' ,     " & vbCrLf &
                              "         colCls = '' ,    " & vbCrLf &
                              "         colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf

            ls_SQL = ls_SQL + "         colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "         colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "         colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "         colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                            - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "         colpasideliveryqty = CASE WHEN ISNULL(PDD.DOQty,0) = 0 THEN ISNULL(SDD.DOQty,0) ELSE ISNULL(PDD.DOQty,0) END,     " & vbCrLf &
                              "         colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0) - CASE ISNULL(PLD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "         coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                          WHEN 0 THEN 0     " & vbCrLf &
                              "                          ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf

            ls_SQL = ls_SQL + "                        END,0)),     " & vbCrLf &
                              "         colstsDO = CASE WHEN ISNULL(PDD.PartNo,'')= '' THEN ISNULL(SDD.PartNo,'') ELSE ISNULL(PDD.PartNo,'') END ,    " & vbCrLf &
                              "         colcartonno = ISNULL(PLD.CartonNo,'') ,    " & vbCrLf &
                              "         colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),    " & vbCrLf &
                              "         sortData = 1, colsupp = KD.SupplierID  " & vbCrLf &
                              " 	FROM   dbo.PO_Master POM     " & vbCrLf &
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                    AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                    AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                           AND KD.PoNo = POD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "                                           AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                           AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                           AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                           AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PONo = SDD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                 AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                 AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                 AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                            AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                            AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                            AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                            AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                            AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                            --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "        LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID     " & vbCrLf &
                              "                                AND KD.KanbanNo = PLD.KanbanNo     " & vbCrLf &
                              "                                --AND KD.SupplierID = PLD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                AND KD.PartNo = PLD.PartNo     " & vbCrLf &
                              "                                AND KD.PoNo = PLD.PoNo                                        " & vbCrLf &
                              "                                AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "                                AND PLD.SuratJalanNo = PDD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   where ISNULL(PLD.CartonNo,'') <> ''  " & vbCrLf
            If pSJ = "" Then
                ls_SQL = ls_SQL + "  AND  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  AND PLD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + " 	) data  " & vbCrLf &
                              " 	ORDER BY colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                              " END  " & vbCrLf

            ls_SQL = ls_SQL + " ELSE  " & vbCrLf &
                              " BEGIN  " & vbCrLf &
                              "  --Data PL masih Kosong " & vbCrLf &
                              "  --Supplier " & vbCrLf &
                              "  	SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                              "             colpono = POM.PONo ,   " & vbCrLf &
                              "             colponos = POM.PONo,     " & vbCrLf &
                              "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                              "                                ELSE 'YES'                            END ,     " & vbCrLf &
                              "             colkanbanno = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colkanbannos = ISNULL(KD.KanbanNo, ''),                   " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.RecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf

            ls_SQL = ls_SQL + "                                                - ( ISNULL(PRD.RecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) , " & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.RecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.RecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0                               ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0, colsupp = KD.SupplierID ,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo ))  " & vbCrLf

            ls_SQL = ls_SQL + "      FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID                                                 " & vbCrLf &
                              " 											  AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo                                                      " & vbCrLf &
                              " 												   AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceiveAffiliate_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceiveAffiliate_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo            " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf
            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf &
                              "     --PASI  " & vbCrLf &
                              " 	SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                              "             colpono = POM.PONo ,   " & vbCrLf &
                              "             colponos = POM.PONo,     " & vbCrLf &
                              "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                              "                                ELSE 'YES'                            END ,     " & vbCrLf &
                              "             colkanbanno = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "             colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(PDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(PDD.DOQty, 0)     " & vbCrLf &
                              "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf

            ls_SQL = ls_SQL + "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(PDD.DOQty,0) , " & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) , " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0                               ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(PDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0, colsupp = KD.SupplierID ,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo ))  " & vbCrLf &
                              "      FROM   dbo.PO_Master POM     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID                                                 " & vbCrLf &
                              " 											  AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo            " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNoSupplier = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	ORDER BY --colstsDO ASC, colkanbannos asc,    " & vbCrLf &
                              "    sortData ASC END  "



            'If pSJ = "" Then
            '    ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf & _
            '                      "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            'End If
            'ls_SQL = ls_SQL + _
            '                  "    --WHERE PDD.SuratJalanNo = 'Surat Jalan PASITrie' --AND PDD.PartNo IN ('7009-2190-02','7009-2191-02') " & vbCrLf & _
            '                  "    --) data " & vbCrLf & _
            '                  "  ORDER BY colstsDO ASC, sortData ASC " & vbCrLf & _
            '                  "   END "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_GridLoad_for1supplier(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = " --Part 1 ada data, Part 2 ga ada data  " & vbCrLf &
                  " IF EXISTS (SELECT * FROM dbo.PLPASI_Detail WHERE SuratJalanNo = '" & Trim(pSJ) & "' --AND PONo = " & Trim(pPO) & " )  " & vbCrLf &
                  " ) BEGIN  " & vbCrLf &
                  " 	SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colstsDO)) FROM (  " & vbCrLf &
                  " --header " & vbCrLf &
                  "      select  distinct " & vbCrLf &
                  "      AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colstsDO, colkanbannos, sortData )) ,      " & vbCrLf &
                  "      colpono, colponos, colpokanban, colkanbanno, colkanbannos, colpartno, colpartname, coluom, " & vbCrLf &
                  "      colCls, colQtyBox, colsuppdelqty, colpasigoodrec, colpasidefectrec, colpasiremaining, " & vbCrLf &
                  "      colpasideliveryqty, colremainingdelqty, coldelqtybox, colstsDO, colcartonno, colcartonqty, sortData, colsupp,colSJSupp " & vbCrLf &
                  "      from(  " & vbCrLf &
                  " 	   SELECT distinct 0 AllowAccess,colno = '',--CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                  "             colpono = POM.PONo ,   " & vbCrLf &
                  "             colponos = POM.PONo,     " & vbCrLf &
                  "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                  "                                ELSE 'YES'                            END ,     " & vbCrLf &
                  "             colkanbanno = ISNULL(KD.KanbanNo, ''),    " & vbCrLf

            ls_SQL = ls_SQL + "             colkanbannos = ISNULL(KD.KanbanNo, ''),    " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf

            ls_SQL = ls_SQL + "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),     " & vbCrLf &
                              "             colremainingdelqty = ISNULL(PDD.DOQty,0) ,--ROUND(CONVERT(CHAR,ISNULL(PRD.RecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.RecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf

            ls_SQL = ls_SQL + "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0 , colsupp = KD.SupplierID ,colSJSupp = SDM.SuratJalanNo  " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") )header " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' )header" & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf &
                              " 	-- PackingList yang sudah ada  " & vbCrLf &
                              " 	SELECT distinct '0'AllowAccess,colno =  '' ,    " & vbCrLf &
                              "         colpono = '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "         colponos = POM.PONo,   " & vbCrLf &
                              "         colpokanban = '' ,     " & vbCrLf &
                              "         colkanbanno = '' ,   " & vbCrLf &
                              "         colkanbannos = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf &
                              "         colpartno = '' ,     " & vbCrLf &
                              "         colpartname = '' ,      " & vbCrLf &
                              "         coluom = '' ,     " & vbCrLf &
                              "         colCls = '' ,    " & vbCrLf &
                              "         colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf

            ls_SQL = ls_SQL + "         colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "         colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "         colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "         colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                            - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "         colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),     " & vbCrLf &
                              "         colremainingdelqty = ISNULL(PDD.DOQty,0) ,--ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0) - CASE ISNULL(PLD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "         coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                          WHEN 0 THEN 0     " & vbCrLf &
                              "                          ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf

            ls_SQL = ls_SQL + "                        END,0)),     " & vbCrLf &
                              "         colstsDO = CASE WHEN ISNULL(PDD.PartNo,'')= '' THEN ISNULL(SDD.PartNo,'') ELSE ISNULL(PDD.PartNo,'') END ,    " & vbCrLf &
                              "         colcartonno = ISNULL(PLD.CartonNo,'') ,    " & vbCrLf &
                              "         colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),    " & vbCrLf &
                              "         sortData = 1, colsupp = KD.SupplierID ,colSJSupp = PLD.SuratJalanNoSupplier " & vbCrLf &
                              " 	FROM   dbo.PO_Master POM     " & vbCrLf &
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                    AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                    AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                           AND KD.PoNo = POD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "                                           AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                           AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                           AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                           AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PONo = SDD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                 AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                 AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                 AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                 AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                            AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                            AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                            AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                            AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                            AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                            --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "        LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID     " & vbCrLf &
                              "                                AND KD.KanbanNo = PLD.KanbanNo     " & vbCrLf &
                              "                                --AND KD.SupplierID = PLD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                AND KD.PartNo = PLD.PartNo     " & vbCrLf &
                              "                                AND KD.PoNo = PLD.PoNo                                        " & vbCrLf &
                              "                                AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "                                AND PLD.SuratJalanNo = PDD.SuratJalanNo " & vbCrLf &
                              "                                AND PLD.SuratJalanNoSupplier = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   where ISNULL(PLD.CartonNo,'') <> ''  " & vbCrLf
            If pSJ = "" Then
                ls_SQL = ls_SQL + "  AND  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  AND PLD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + " 	) data  " & vbCrLf &
                              " 	ORDER BY colSJSupp asc, colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                              " END  " & vbCrLf

            ls_SQL = ls_SQL + " ELSE  " & vbCrLf &
                              " BEGIN  " & vbCrLf &
                              "  --Data PL masih Kosong " & vbCrLf &
                              "  select distinct NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colkanbanno,colPartNo,colPONo,colSupp)) ,* from ( " & vbCrLf &
                              "  	SELECT distinct 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                              "             colpono = POM.PONo ,   " & vbCrLf &
                              "             colponos = POM.PONo,     " & vbCrLf &
                              "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                              "                                ELSE 'YES'                            END ,     " & vbCrLf &
                              "             colkanbanno = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colkanbannos = ISNULL(KD.KanbanNo, ''),                   " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf

            ls_SQL = ls_SQL + "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))), " & vbCrLf &
                              "             colremainingdelqty = ISNULL(PDD.DOQty,0) ,--ROUND(CONVERT(CHAR,ISNULL(PRD.DefectRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0                               ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0, colsupp = KD.SupplierID ,colSJSupp=PDD.SuratJalanNoSupplier  " & vbCrLf

            ls_SQL = ls_SQL + "      FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID                                                 " & vbCrLf &
                              " 											  AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo                                                      " & vbCrLf &
                              " 												   AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo            " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNoSupplier = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)PLKosong --ORDER BY --colstsDO ASC, colkanbannos asc,  sortData ASC   " & vbCrLf &
                              "    END  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_GridLoad(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = " --Part 1 ada data, Part 2 ga ada data   " & vbCrLf &
                     "  IF EXISTS (SELECT * FROM dbo.PLPASI_Detail WHERE SuratJalanNo = '" & txtsuratjalanno.Text & "'    " & vbCrLf &
                     "  ) BEGIN   " & vbCrLf &
                     "  	SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colstsDO)), colno = colno FROM (   " & vbCrLf &
                     "  --header  " & vbCrLf &
                     "       select  distinct  " & vbCrLf &
                     "       AllowAccess, colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colstsDO, colkanbannos, sortData )) ,       " & vbCrLf &
                     "       colpono, colponos, colpokanban, colkanbanno, colkanbannos, colpartno, colpartname, coluom,  " & vbCrLf &
                     "       colCls, colQtyBox, colsuppdelqty, colpasigoodrec, colpasidefectrec, colpasiremaining,  " & vbCrLf &
                     "       colpasideliveryqty, colremainingdelqty, coldelqtybox, colstsDO, colcartonno, colcartonqty, sortData, colsupp,colSJSupp  " & vbCrLf &
                     "       from(   " & vbCrLf

            ls_SQL = ls_SQL + "  	   SELECT distinct 0 AllowAccess,colno = '', " & vbCrLf &
                              "              colpono = POM.PONo ,    " & vbCrLf &
                              "              colponos = POM.PONo,      " & vbCrLf &
                              "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'      " & vbCrLf &
                              "                                 ELSE 'YES'                            END ,      " & vbCrLf &
                              "              colkanbanno = ISNULL(KD.KanbanNo, ''),     " & vbCrLf &
                              "              colkanbannos = ISNULL(KD.KanbanNo, ''),     " & vbCrLf &
                              "              colpartno = POD.PartNo ,      " & vbCrLf &
                              "              colpartname = MP.PartName ,      " & vbCrLf &
                              "              coluom = UC.Description ,      " & vbCrLf &
                              "              colCls = UC.unitcls ,     " & vbCrLf

            ls_SQL = ls_SQL + "              colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "              colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "              colpasigoodrec = '',      " & vbCrLf &
                              "              colpasidefectrec = '',      " & vbCrLf &
                              "              colpasiremaining = '' ,      " & vbCrLf &
                              "              colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),      " & vbCrLf &
                              "              colremainingdelqty = ISNULL(PDD.DOQty,0) , " & vbCrLf &
                              "              coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                               WHEN 0 THEN 0      " & vbCrLf &
                              "                               ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                             END,0)),      " & vbCrLf

            ls_SQL = ls_SQL + "              colstsDO = ISNULL(SDD.PartNo,'') ,     " & vbCrLf &
                              "              colcartonno = '',     " & vbCrLf &
                              "              colcartonqty = '',     " & vbCrLf &
                              "              sortData = 0 , colsupp = KD.SupplierID ,colSJSupp = pdd.SuratJalanNoSupplier    " & vbCrLf &
                              "  		FROM   dbo.PO_Master POM      " & vbCrLf &
                              "              LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                         AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                         AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "              LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID          " & vbCrLf &
                              "                                                AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                                AND KD.SupplierID = POD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "              LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                                AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                                AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "              INNER JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  						FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo          " & vbCrLf

            ls_SQL = ls_SQL + "              LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID         " & vbCrLf &
                              "                                                     AND SDM.SupplierID = SDD.SupplierID           " & vbCrLf &
                              "              LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0)), SuratJalanNoSupplier=max(SuratJalanNoSupplier)     " & vbCrLf &
                              "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo " & vbCrLf &
                              "              LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo         " & vbCrLf

            ls_SQL = ls_SQL + "              LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "              LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf


            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") )header " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf &
                                  "         AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                                  " )header" & vbCrLf
            End If

            ls_SQL = ls_SQL + " UNION ALL  " & vbCrLf &
                              "  	-- PackingList yang sudah ada   " & vbCrLf &
                              "  	SELECT distinct '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf

            ls_SQL = ls_SQL + "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "          colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "          colpasigoodrec = '',      " & vbCrLf &
                              "          colpasidefectrec = '',      " & vbCrLf &
                              "          colpasiremaining = '' ,      " & vbCrLf &
                              "          colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),      " & vbCrLf &
                              "          colremainingdelqty = ISNULL(PDD.DOQty,0) , " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf

            ls_SQL = ls_SQL + "                         END,0)),      " & vbCrLf &
                              "          colstsDO = CASE WHEN ISNULL(PDD.PartNo,'')= '' THEN ISNULL(SDD.PartNo,'') ELSE ISNULL(PDD.PartNo,'') END ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID ,colSJSupp = pdd.SuratJalanNoSupplier   " & vbCrLf &
                              "  	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                            AND KD.PoNo = POD.PONo      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          INNER JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  						FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                    AND KD.PartNo = SDD.PartNo          " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID          " & vbCrLf &
                              "          LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0)), SuratJalanNoSupplier=max(SuratJalanNoSupplier)     " & vbCrLf &
                              "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo      " & vbCrLf &
                              "         LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID      " & vbCrLf &
                              "                                 AND KD.KanbanNo = PLD.KanbanNo         " & vbCrLf &
                              "                                 AND KD.PartNo = PLD.PartNo      " & vbCrLf &
                              "                                 AND KD.PoNo = PLD.PoNo                                         " & vbCrLf &
                              "                                 AND PLD.SuratJalanNo = PDD.SuratJalanNo  " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   where ISNULL(PLD.CartonNo,'') <> ''   " & vbCrLf

            If pSJ = "" Then
                ls_SQL = ls_SQL + "  AND  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  AND PLD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf &
                                  "  AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf
            End If
            ls_SQL = ls_SQL + " 	) data  " & vbCrLf &
                              " 	ORDER BY colSJSupp asc, colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                              " END  " & vbCrLf

            ls_SQL = ls_SQL + " ELSE  " & vbCrLf &
                              " BEGIN  " & vbCrLf &
                              "  --Data PL masih Kosong " & vbCrLf &
                              " select distinct NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colkanbanno,colPartNo,colPONo,colSupp)), colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colkanbanno,colPartNo,colPONo,colSupp)) ,* from (  " & vbCrLf &
                              "   	SELECT distinct 0 AllowAccess,--colno = '' ,      " & vbCrLf &
                              "              colpono = POM.PONo ,    " & vbCrLf &
                              "              colponos = POM.PONo,      " & vbCrLf &
                              "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'      " & vbCrLf &
                              "                                 ELSE 'YES'                            END ,      " & vbCrLf &
                              "              colkanbanno = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "              colkanbannos = ISNULL(KD.KanbanNo, ''),                    " & vbCrLf &
                              "              colpartno = POD.PartNo ,      " & vbCrLf &
                              "              colpartname = MP.PartName ,      " & vbCrLf

            ls_SQL = ls_SQL + "              coluom = UC.Description ,      " & vbCrLf &
                              "              colCls = UC.unitcls ,             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "              colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "              colpasigoodrec = '',      " & vbCrLf &
                              "              colpasidefectrec ='',      " & vbCrLf &
                              "              colpasiremaining = '' ,      " & vbCrLf &
                              "              colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),  " & vbCrLf &
                              "              colremainingdelqty = ISNULL(PDD.DOQty,0) , " & vbCrLf &
                              "              coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                               WHEN 0 THEN 0                               ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                             END,0)),      " & vbCrLf

            ls_SQL = ls_SQL + "              colstsDO = ISNULL(SDD.PartNo,'') ,     " & vbCrLf &
                              "              colcartonno = '',     " & vbCrLf &
                              "              colcartonqty = '',     " & vbCrLf &
                              "              sortData = 0, colsupp = KD.SupplierID ,colSJSupp=pdd.SuratJalanNoSupplier   " & vbCrLf &
                              "       FROM   dbo.PO_Master POM      " & vbCrLf &
                              "              LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                         AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                         AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "              LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID                                                  " & vbCrLf &
                              "  											  AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                                AND KD.SupplierID = POD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "              LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                                AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                                AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "              INNER JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  						FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo        " & vbCrLf

            ls_SQL = ls_SQL + "              INNER JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID       " & vbCrLf &
                              "                                                     AND SDM.SupplierID = SDD.SupplierID          " & vbCrLf &
                              "              INNER JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0)), SuratJalanNoSupplier=max(SuratJalanNoSupplier)     " & vbCrLf &
                              "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo " & vbCrLf &
                              "              LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo       " & vbCrLf

            ls_SQL = ls_SQL + "              LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "              LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf &
                              "              LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf


            If pSJ = "" Then
                ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf &
                                  "         AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                                  "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            Else
                ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf &
                                  "         AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)PLKosong --ORDER BY --colstsDO ASC, colkanbannos asc,  sortData ASC   " & vbCrLf &
                              "    END  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 120
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_AddRowOriginal(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String, ByVal pPartNo As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = " SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colponos)) FROM (  " & vbCrLf &
                  "  SELECT '0'AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,    " & vbCrLf &
                  "            colpono = POM.PONo ,  " & vbCrLf &
                  "            colponos = POM.PONo,    " & vbCrLf &
                  "            colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'    " & vbCrLf &
                  "                               ELSE 'YES'                            END ,    " & vbCrLf &
                  "            colkanbanno = CASE WHEN POD.KanbanCls = '0' THEN '-'    " & vbCrLf &
                  "                               ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf &
                  "                          END ,    " & vbCrLf &
                  "            colkanbannos = CASE WHEN POD.KanbanCls = '0' THEN '-'    " & vbCrLf &
                  "                               ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf

            ls_SQL = ls_SQL + "                          END ,                  " & vbCrLf &
                              "            colpartno = POD.PartNo ,    " & vbCrLf &
                              "            colpartname = MP.PartName ,    " & vbCrLf &
                              "            coluom = UC.Description ,    " & vbCrLf &
                              "            colCls = UC.unitcls ,   " & vbCrLf &
                              "            colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,    " & vbCrLf &
                              "            colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,    " & vbCrLf &
                              "            colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),    " & vbCrLf &
                              "            colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),    " & vbCrLf &
                              "            colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)    " & vbCrLf &
                              "                                               - ( ISNULL(PRD.GoodRecQty, 0)    " & vbCrLf

            ls_SQL = ls_SQL + "                                                   + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,    " & vbCrLf &
                              "            colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),    " & vbCrLf &
                              "            colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,    " & vbCrLf &
                              "            coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)    " & vbCrLf &
                              "                             WHEN 0 THEN 0    " & vbCrLf &
                              "                             ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)    " & vbCrLf &
                              "                           END,0)),    " & vbCrLf &
                              "            colstsDO = ISNULL(PDD.PartNo,'') ,   " & vbCrLf &
                              "            colcartonno = '',   " & vbCrLf &
                              "            colcartonqty = '',   " & vbCrLf &
                              "            sortData = 0  " & vbCrLf

            ls_SQL = ls_SQL + "     FROM   dbo.PO_Master POM    " & vbCrLf &
                              "            LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf &
                              "                                       AND POM.PoNo = POD.PONo    " & vbCrLf &
                              "                                       AND POM.SupplierID = POD.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID        " & vbCrLf &
                              "                                              AND KD.PoNo = POD.PONo    " & vbCrLf &
                              "                                              AND KD.SupplierID = POD.SupplierID    " & vbCrLf &
                              "                                              AND KD.PartNo = POD.PartNo    " & vbCrLf &
                              "            LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf &
                              "                                              AND KD.KanbanNo = KM.KanbanNo    " & vbCrLf &
                              "                                              AND KD.SupplierID = KM.SupplierID    " & vbCrLf &
                              "                                              AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf

            ls_SQL = ls_SQL + "            LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf &
                              "                                                   AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf &
                              "                                                   AND KD.SupplierID = SDD.SupplierID    " & vbCrLf &
                              "                                                   AND KD.PartNo = SDD.PartNo    " & vbCrLf &
                              "                                                   AND KD.PONo = SDD.PONo    " & vbCrLf &
                              "            LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID    " & vbCrLf &
                              "                                                   AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf &
                              "                                                   AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID    " & vbCrLf &
                              "                                                    AND KD.KanbanNo = PRD.KanbanNo    " & vbCrLf &
                              "                                                    AND KD.SupplierID = PRD.SupplierID    " & vbCrLf

            ls_SQL = ls_SQL + "                                                    AND KD.PartNo = PRD.PartNo    " & vbCrLf &
                              "                                                    AND KD.PartNo = PRD.PartNo    " & vbCrLf &
                              "            LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf &
                              "                                                    AND PRM.SuratJalanNo = PRD.SuratJalanNo    " & vbCrLf &
                              "                                                    AND PRM.SupplierID = PRD.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID    " & vbCrLf &
                              "                                               AND KD.KanbanNo = PDD.KanbanNo    " & vbCrLf &
                              "                                               AND KD.SupplierID = PDD.SupplierID    " & vbCrLf &
                              "                                               AND KD.PartNo = PDD.PartNo    " & vbCrLf &
                              "                                               AND KD.PoNo = PDD.PoNo    " & vbCrLf &
                              "            LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND PDD.SuratJalanNo = PDM.SuratJalanNo    " & vbCrLf &
                              "                                               AND PDD.SupplierID = PDM.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "            LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf &
                              "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf &
                              "            LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf &
                              "            LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf &
                              "   WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "'  " & vbCrLf &
                              "     --WHERE PDD.SuratJalanNo = 'Surat Jalan PASITrie'  --AND PDD.PartNo IN ('7009-2190-02','7009-2191-02')  " & vbCrLf &
                              "  UNION all  " & vbCrLf &
                              "  SELECT '0'AllowAccess,colno =  '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "         colpono = '' ,     " & vbCrLf &
                              "         colponos = POM.PONo,   " & vbCrLf &
                              "         colpokanban = '' ,     " & vbCrLf &
                              "         colkanbanno = '' ,   " & vbCrLf &
                              "         colkanbannos = CASE WHEN POD.KanbanCls = '0' THEN '-'     " & vbCrLf &
                              "                            ELSE ISNULL(KD.KanbanNo, '')     " & vbCrLf &
                              "                       END ,       " & vbCrLf &
                              "         colpartno = '' ,     " & vbCrLf &
                              "         colpartname = '' ,      " & vbCrLf &
                              "         coluom = '' ,     " & vbCrLf &
                              "         colCls = '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "         colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "         colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "         colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "         colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "         colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                            - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "         colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),     " & vbCrLf &
                              "         colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "         coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                          WHEN 0 THEN 0     " & vbCrLf

            ls_SQL = ls_SQL + "                          ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                        END,0)),     " & vbCrLf &
                              "         colstsDO = ISNULL(PDD.PartNo,'') ,    " & vbCrLf &
                              "         colcartonno = ISNULL(PLD.CartonNo,'') ,    " & vbCrLf &
                              "         colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),    " & vbCrLf &
                              "         sortData = 1  " & vbCrLf &
                              "  FROM   dbo.PO_Master POM     " & vbCrLf &
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                    AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                    AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                           AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                           AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                           AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                           AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                           AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                 AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                 AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                            AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                            AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                            AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                            AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                            AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "        LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID     " & vbCrLf &
                              "                                AND KD.KanbanNo = PLD.KanbanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                AND KD.SupplierID = PLD.SupplierID     " & vbCrLf &
                              "                                AND KD.PartNo = PLD.PartNo     " & vbCrLf &
                              "                                AND KD.PoNo = PLD.PoNo                                        " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "   WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "'  " & vbCrLf &
                              "   UNION ALL  " & vbCrLf &
                              "   SELECT '0'AllowAccess,colno =  '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "         colpono = '' ,     " & vbCrLf &
                              "         colponos = POM.PONo,   " & vbCrLf &
                              "         colpokanban = '' ,     " & vbCrLf &
                              "         colkanbanno = '' ,   " & vbCrLf &
                              "         colkanbannos = CASE WHEN POD.KanbanCls = '0' THEN '-'     " & vbCrLf &
                              "                            ELSE ISNULL(KD.KanbanNo, '')     " & vbCrLf &
                              "                       END ,       " & vbCrLf &
                              "         colpartno = '' ,     " & vbCrLf &
                              "         colpartname = '' ,      " & vbCrLf &
                              "         coluom = '' ,     " & vbCrLf &
                              "         colCls = '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "         colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "         colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "         colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "         colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "         colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                            - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "         colpasideliveryqty = CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END ,     " & vbCrLf &
                              "         colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "         coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                          WHEN 0 THEN 0     " & vbCrLf

            ls_SQL = ls_SQL + "                          ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                        END,0)),     " & vbCrLf &
                              "         colstsDO = ISNULL(PDD.PartNo,'') ,    " & vbCrLf &
                              "         colcartonno = '' ,    " & vbCrLf &
                              "         colcartonqty = 0,    " & vbCrLf &
                              "         sortData = 1  " & vbCrLf &
                              "  FROM   dbo.PO_Master POM     " & vbCrLf &
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                    AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                    AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     "

            ls_SQL = ls_SQL + "                                           AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                           AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                           AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                           AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                           AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                 AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                 AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                            AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                            AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                            AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                            AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                            AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "                                   " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "   WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' AND PDD.partno IN ('" & Trim(pPartNo) & "') " & vbCrLf &
                              "   ) data  " & vbCrLf &
                              "   ORDER BY colstsDO ASC, sortData ASC " & vbCrLf



            'If pSJ = "" Then
            '    ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf & _
            '                      "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & pSJ & "' AND POD.PartNo IN ('" & pPartNo & "')" & vbCrLf
            'End If

            'ls_SQL = ls_SQL + "    ORDER BY colno,colKanbanNo DESC  " & vbCrLf & _
            '                  "   END " & vbCrLf

            'ls_SQL = ls_SQL + "  --ORDER BY KD.KanbanNo " & vbCrLf



            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_AddRowTree(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String, ByVal pPartNo As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = "  SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colponos)) FROM (   " & vbCrLf &
                   "   --Delivery Supplier " & vbCrLf &
                   " 	   SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                   "             colpono = POM.PONo ,   " & vbCrLf &
                   "             colponos = POM.PONo,     " & vbCrLf &
                   "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                   "                                ELSE 'YES'                            END ,     " & vbCrLf &
                   "             colkanbanno = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                   "             colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf

            ls_SQL = ls_SQL + "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) ," & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf

            ls_SQL = ls_SQL + "             sortData = 0   " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              " 		WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "'   " & vbCrLf &
                              " 		--Delivery PASI " & vbCrLf &
                              " 		UNION ALL " & vbCrLf &
                              " 		SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                              "             colpono = POM.PONo ,   " & vbCrLf &
                              "             colponos = POM.PONo,     " & vbCrLf &
                              "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                              "                                ELSE 'YES'                            END ,     " & vbCrLf &
                              "             colkanbanno = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colkanbannos = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf

            ls_SQL = ls_SQL + "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) , " & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0   " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = SDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = SDD.SuratJalanNoSupplier " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON SDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND SDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND SDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              " 		WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "'      " & vbCrLf &
                              "   UNION all   " & vbCrLf &
                              " 	--PackingList udah ada " & vbCrLf &
                              " 	SELECT DISTINCT * FROM ( " & vbCrLf &
                              " 	--PackingList udah ada " & vbCrLf &
                              "         SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = 0 ,      " & vbCrLf &
                              "          colsuppdelqty = 0 ,      " & vbCrLf &
                              "          colpasigoodrec = 0,      " & vbCrLf &
                              "          colpasidefectrec = 0,      " & vbCrLf

            ls_SQL = ls_SQL + "          colpasiremaining = 0 ,      " & vbCrLf &
                              "          colpasideliveryqty = ISNULL(SDD.DOQty,0), " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PLD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(POD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1   " & vbCrLf

            ls_SQL = ls_SQL + " 	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                            AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND KD.KanbanNo = SDD.KanbanNo      " & vbCrLf &
                              "   							                    AND KD.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "                                                 AND KD.PartNo = SDD.PartNo      " & vbCrLf &
                              "                                                 AND KD.PONo = SDD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDM.SuratJalanNo = SDD.SuratJalanNo      " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "                                                  AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID      " & vbCrLf
            ls_SQL = ls_SQL + "          LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                             AND KD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                             AND KD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                             AND KD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                             AND KD.PoNo = PDD.PoNo      " & vbCrLf &
                              "                                             --AND PDD.SuratJalanNoSupplier = PRD.SuratJalanNo  " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo      " & vbCrLf &
                              "                                             --AND PDD.SupplierID = PDM.SupplierID  " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Master PLM ON SDD.AffiliateID = PLM.AffiliateID      " & vbCrLf &
                              "                                             AND SDD.SuratJalanNo = PLM.SuratJalanNo      " & vbCrLf &
                              "                                             AND SDD.SupplierID = PLM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID      " & vbCrLf &
                              "                                 AND KD.KanbanNo = PLD.KanbanNo      " & vbCrLf

            ls_SQL = ls_SQL + "                                 AND KD.SupplierID = PLD.SupplierID      " & vbCrLf &
                              "                                 AND KD.PartNo = PLD.PartNo      " & vbCrLf &
                              "                                 AND KD.PoNo = PLD.PoNo  " & vbCrLf &
                              "                                 AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and SDD.DOQty <> 0 AND ISNULL(PLD.CartonNo,'') <> '' " & vbCrLf &
                              "    UNION ALL " & vbCrLf &
                              "    --Delivery PASI " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colQtyBox = 0 ,      " & vbCrLf &
                              "          colsuppdelqty = 0 ,      " & vbCrLf &
                              "          colpasigoodrec = 0,      " & vbCrLf &
                              "          colpasidefectrec = 0,      " & vbCrLf &
                              "          colpasiremaining = 0 ,      " & vbCrLf &
                              "          colpasideliveryqty = ISNULL(PDD.DOQty,0), " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf

            ls_SQL = ls_SQL + "          colstsDO = ISNULL(PDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1   " & vbCrLf &
                              " 	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                            AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf &
                              "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID " & vbCrLf &
                              " 		 LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier" & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Master PLM ON PDD.AffiliateID = PLM.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND PDD.SuratJalanNo = PLM.SuratJalanNo      " & vbCrLf &
                              "                                             AND PDD.SupplierID = PLM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID      " & vbCrLf &
                              "                                 AND KD.KanbanNo = PLD.KanbanNo      " & vbCrLf &
                              "                                 AND KD.SupplierID = PLD.SupplierID      " & vbCrLf &
                              "                                 AND KD.PartNo = PLD.PartNo      " & vbCrLf &
                              "                                 AND KD.PoNo = PLD.PoNo                                         " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and PDD.DOQty <> 0  AND ISNULL(PLD.CartonNo,'') <> '')test " & vbCrLf &
                              "    UNION ALL   " & vbCrLf &
                              "    --Delivery Supplier " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "          colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "          colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),      " & vbCrLf &
                              "          colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),      " & vbCrLf &
                              "          colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)      " & vbCrLf &
                              "                                             - ( ISNULL(PRD.GoodRecQty, 0)      " & vbCrLf

            ls_SQL = ls_SQL + "                                                 + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,      " & vbCrLf &
                              "          colpasideliveryqty = ISNULL(SDD.DOQty,0) , " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(SDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = '' ,     " & vbCrLf &
                              "          colcartonqty = 0,     " & vbCrLf &
                              "          sortData = 1   " & vbCrLf

            ls_SQL = ls_SQL + "   FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf &
                              " 										   AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND KD.KanbanNo = SDD.KanbanNo      " & vbCrLf &
                              "   							                    AND KD.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "                                                 AND KD.PartNo = SDD.PartNo      " & vbCrLf &
                              "                                                 AND KD.PONo = SDD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDM.SuratJalanNo = SDD.SuratJalanNo      " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "                                                  AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "    WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "' AND SDD.partno IN (" & Trim(pPartNo) & ")  " & vbCrLf &
                              "    UNION ALL " & vbCrLf &
                              "    --Delivery PASI " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "          colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(PDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "          colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),      " & vbCrLf &
                              "          colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),      " & vbCrLf &
                              "          colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(PDD.DOQty, 0)      " & vbCrLf &
                              "                                             - ( ISNULL(PRD.GoodRecQty, 0)      " & vbCrLf &
                              "                                                 + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,      " & vbCrLf

            ls_SQL = ls_SQL + "          colpasideliveryqty = ISNULL(PDD.DOQty,0) ,   " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(PDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = '' ,     " & vbCrLf &
                              "          colcartonqty = 0,     " & vbCrLf &
                              "          sortData = 1   " & vbCrLf &
                              "   FROM   dbo.PO_Master POM      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf &
                              " 										   AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf &
                              "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                             AND KD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                             AND KD.SupplierID = PDD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND KD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                             AND KD.PoNo = PDD.PoNo      " & vbCrLf &
                              "                                             AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo      " & vbCrLf &
                              "                                             --AND PDD.SupplierID = PDM.SupplierID      " & vbCrLf &
                              "                                     " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "    WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' AND PDD.partno IN (" & Trim(pPartNo) & ") AND PDD.KanbanNo = '" & Trim(pKanban) & "'" & vbCrLf &
                              "    ) data   " & vbCrLf &
                              "    ORDER BY colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_AddRow_(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String, ByVal pPartNo As String, ByVal pCombination As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = "  SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colponos)) FROM (   " & vbCrLf &
                     " --header " & vbCrLf &
                     "      select  " & vbCrLf &
                     "      AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colstsDO, colkanbannos, sortData )) ,      " & vbCrLf &
                     "      colpono, colponos, colpokanban, colkanbanno, colkanbannos, colpartno, colpartname, coluom, " & vbCrLf &
                     "      colCls, colQtyBox, colsuppdelqty, colpasigoodrec, colpasidefectrec, colpasiremaining, " & vbCrLf &
                     "      colpasideliveryqty, colremainingdelqty, coldelqtybox, colstsDO, colcartonno, colcartonqty, sortData, colsupp " & vbCrLf &
                     "      from(  " & vbCrLf &
                     "   --Delivery Supplier " & vbCrLf &
                     " 	   SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                     "             colpono = POM.PONo ,   " & vbCrLf &
                     "             colponos = POM.PONo,     " & vbCrLf &
                     "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                     "                                ELSE 'YES'                            END ,     " & vbCrLf &
                     "             colkanbanno = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                     "             colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf

            ls_SQL = ls_SQL + "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) ," & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf

            ls_SQL = ls_SQL + "             sortData = 0, colsupp = KD.SupplierID   " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              " 		WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "'   " & vbCrLf &
                              " 		--Delivery PASI " & vbCrLf &
                              " 		UNION ALL " & vbCrLf &
                              " 		SELECT 0 AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                              "             colpono = POM.PONo ,   " & vbCrLf &
                              "             colponos = POM.PONo,     " & vbCrLf &
                              "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                              "                                ELSE 'YES'                            END ,     " & vbCrLf &
                              "             colkanbanno = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colkanbannos = ISNULL(KD.KanbanNo, '') ,    " & vbCrLf &
                              "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf

            ls_SQL = ls_SQL + "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf &
                              "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = ISNULL(SDD.DOQty,0) , " & vbCrLf &
                              "             colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf

            ls_SQL = ls_SQL + "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf &
                              "             sortData = 0, colsupp = KD.SupplierID   " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = SDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = SDD.SuratJalanNoSupplier " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON SDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND SDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND SDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              " 		WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "' )Header     " & vbCrLf &
                              "   UNION all   " & vbCrLf &
                              " 	--PackingList udah ada " & vbCrLf &
                              " 	SELECT DISTINCT * FROM ( " & vbCrLf &
                              " 	--PackingList udah ada " & vbCrLf &
                              "         SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = 0 ,      " & vbCrLf &
                              "          colsuppdelqty = 0 ,      " & vbCrLf &
                              "          colpasigoodrec = 0,      " & vbCrLf &
                              "          colpasidefectrec = 0,      " & vbCrLf

            ls_SQL = ls_SQL + "          colpasiremaining = 0 ,      " & vbCrLf &
                              "          colpasideliveryqty = CASE WHEN ISNULL(PDD.DOQty,0) = 0 THEN ISNULL(SDD.DOQty,0) ELSE ISNULL(PDD.DOQty,0) END, " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PLD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(POD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID   " & vbCrLf

            ls_SQL = ls_SQL + " 	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                    AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                    AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                           AND KD.PoNo = POD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "                                           AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                           AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                           AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                           AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PONo = SDD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                 AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                 AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                 AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                            AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                            AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                            AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                            AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                            AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                            --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "        LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID     " & vbCrLf &
                              "                                AND KD.KanbanNo = PLD.KanbanNo     " & vbCrLf &
                              "                                --AND KD.SupplierID = PLD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                AND KD.PartNo = PLD.PartNo     " & vbCrLf &
                              "                                AND KD.PoNo = PLD.PoNo                                        " & vbCrLf &
                              "                                AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "                                AND PLD.SuratJalanNo = PDD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and SDD.DOQty <> 0 AND ISNULL(PLD.CartonNo,'') <> '' " & vbCrLf &
                              "    UNION ALL " & vbCrLf &
                              "    --Delivery PASI " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colQtyBox = 0 ,      " & vbCrLf &
                              "          colsuppdelqty = 0 ,      " & vbCrLf &
                              "          colpasigoodrec = 0,      " & vbCrLf &
                              "          colpasidefectrec = 0,      " & vbCrLf &
                              "          colpasiremaining = 0 ,      " & vbCrLf &
                              "          colpasideliveryqty = ISNULL(PDD.DOQty,0), " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf

            ls_SQL = ls_SQL + "          colstsDO = ISNULL(PDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID   " & vbCrLf &
                              " 	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                            AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf &
                              "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID " & vbCrLf &
                              " 		 LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier" & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Master PLM ON PDD.AffiliateID = PLM.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND PDD.SuratJalanNo = PLM.SuratJalanNo      " & vbCrLf &
                              "                                             AND PDD.SupplierID = PLM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID      " & vbCrLf &
                              "                                 AND KD.KanbanNo = PLD.KanbanNo      " & vbCrLf &
                              "                                 AND KD.SupplierID = PLD.SupplierID      " & vbCrLf &
                              "                                 AND KD.PartNo = PLD.PartNo      " & vbCrLf &
                              "                                 AND KD.PoNo = PLD.PoNo " & vbCrLf &
                              "                                 AND PLD.SuratJalanNo = PLM.SuratJalanNo " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and PDD.DOQty <> 0  AND ISNULL(PLD.CartonNo,'') <> '')test " & vbCrLf &
                              "    UNION ALL   " & vbCrLf &
                              "    --Delivery Supplier " & vbCrLf
            ls_SQL = ls_SQL + " SELECT AllowAccess,colno,colpono,colponos,colpokanban,colkanbanno,     " & vbCrLf &
                              " colkanbannos,colpartno,colpartname,coluom,colCls,colQtyBox,       " & vbCrLf &
                              " colsuppdelqty,colpasigoodrec,colpasidefectrec,colpasiremaining,       " & vbCrLf &
                              " colpasideliveryqty,colremainingdelqty,coldelqtybox,colstsDO ,      " & vbCrLf &
                              " colcartonno= CASE WHEN ROUND(colpasideliveryqty/colQtyBox,0) = 1 THEN 'C001' " & vbCrLf &
                              " WHEN ROUND(colpasideliveryqty/colQtyBox,0) >= 10 THEN 'C001-C0' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0)) " & vbCrLf &
                              " WHEN ROUND(colpasideliveryqty/colQtyBox,0) >= 10 THEN 'C001-C' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0)) " & vbCrLf &
                              " ELSE 'C001-C00' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0))END , " & vbCrLf &
                              " colcartonqty = round(colpasideliveryqty/colQtyBox,0),sortData, colsupp " & vbCrLf &
                              " FROM ( " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "          colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "          colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),      " & vbCrLf &
                              "          colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),      " & vbCrLf &
                              "          colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)      " & vbCrLf &
                              "                                             - ( ISNULL(PRD.GoodRecQty, 0)      " & vbCrLf

            ls_SQL = ls_SQL + "                                                 + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,      " & vbCrLf &
                              "          colpasideliveryqty = ISNULL(SDD.DOQty,0) , " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(SDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = '' ,     " & vbCrLf &
                              "          colcartonqty = 0,     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID   " & vbCrLf

            ls_SQL = ls_SQL + "   FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf &
                              " 										   AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND KD.KanbanNo = SDD.KanbanNo      " & vbCrLf &
                              "   							                    AND KD.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "                                                 AND KD.PartNo = SDD.PartNo      " & vbCrLf &
                              "                                                 AND KD.PONo = SDD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDM.SuratJalanNo = SDD.SuratJalanNo      " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "                                                  AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "    WHERE SDD.SuratJalanNo = '" & Trim(pSJ) & "' AND (Rtrim(SDD.PONO)+Rtrim(SDD.KanbanNo)+Rtrim(SDD.PartNo)) IN (" & Trim(pCombination) & ") --AND SDD.partno IN (" & Trim(pPartNo) & ")  " & vbCrLf &
                              "    UNION ALL " & vbCrLf &
                              "    --Delivery PASI " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "          colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(PDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "          colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),      " & vbCrLf &
                              "          colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),      " & vbCrLf &
                              "          colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(PDD.DOQty, 0)      " & vbCrLf &
                              "                                             - ( ISNULL(PRD.GoodRecQty, 0)      " & vbCrLf &
                              "                                                 + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,      " & vbCrLf

            ls_SQL = ls_SQL + "          colpasideliveryqty = ISNULL(PDD.DOQty,0) ,   " & vbCrLf &
                              "          colremainingdelqty = ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(PDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = '' ,     " & vbCrLf &
                              "          colcartonqty = 0,     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID   " & vbCrLf &
                              "   FROM   dbo.PO_Master POM      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf &
                              " 										   AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf &
                              "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                             AND KD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                             AND KD.SupplierID = PDD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND KD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                             AND KD.PoNo = PDD.PoNo      " & vbCrLf &
                              "                                             AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo      " & vbCrLf &
                              "                                             --AND PDD.SupplierID = PDM.SupplierID      " & vbCrLf &
                              "                                     " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "    WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' AND (Rtrim(PDD.PONO)+Rtrim(PDD.KanbanNo)+Rtrim(PDD.PartNo)) IN (" & Trim(pCombination) & ") " & vbCrLf &
                              "    )PLKosong " & vbCrLf &
                              "    ) data   " & vbCrLf &
                              "    ORDER BY colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_AddRow_1Supplier(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String, ByVal pPartNo As String, ByVal pCombination As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = "  SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colponos)) FROM (   " & vbCrLf &
                     " --header " & vbCrLf &
                     "      select distinct  " & vbCrLf &
                     "      AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colstsDO, colkanbannos, sortData )) ,      " & vbCrLf &
                     "      colpono, colponos, colpokanban, colkanbanno, colkanbannos, colpartno, colpartname, coluom, " & vbCrLf &
                     "      colCls, colQtyBox, colsuppdelqty, colpasigoodrec, colpasidefectrec, colpasiremaining, " & vbCrLf &
                     "      colpasideliveryqty, colremainingdelqty, coldelqtybox, colstsDO, colcartonno, colcartonqty, sortData, colsupp,colSJSupp " & vbCrLf &
                     "      from(  " & vbCrLf &
                     " 	   SELECT distinct 0 AllowAccess,colno = '',--CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo,POD.PartNo )) ,     " & vbCrLf &
                     "             colpono = POM.PONo ,   " & vbCrLf &
                     "             colponos = POM.PONo,     " & vbCrLf &
                     "             colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'     " & vbCrLf &
                     "                                ELSE 'YES'                            END ,     " & vbCrLf &
                     "             colkanbanno = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                     "             colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "             colpartno = POD.PartNo ,     " & vbCrLf &
                              "             colpartname = MP.PartName ,     " & vbCrLf &
                              "             coluom = UC.Description ,     " & vbCrLf &
                              "             colCls = UC.unitcls ,    " & vbCrLf &
                              "             colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,     " & vbCrLf &
                              "             colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,     " & vbCrLf &
                              "             colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),     " & vbCrLf &
                              "             colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),     " & vbCrLf &
                              "             colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)     " & vbCrLf

            ls_SQL = ls_SQL + "                                                - ( ISNULL(PRD.GoodRecQty, 0)     " & vbCrLf &
                              "                                                    + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,     " & vbCrLf &
                              "             colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0))))," & vbCrLf &
                              "             colremainingdelqty = ISNULL(PDD.DOQty,0) ,--ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,     " & vbCrLf &
                              "             coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                              WHEN 0 THEN 0     " & vbCrLf &
                              "                              ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)     " & vbCrLf &
                              "                            END,0)),     " & vbCrLf &
                              "             colstsDO = ISNULL(SDD.PartNo,'') ,    " & vbCrLf &
                              "             colcartonno = '',    " & vbCrLf &
                              "             colcartonqty = '',    " & vbCrLf

            ls_SQL = ls_SQL + "             sortData = 0, colsupp = KD.SupplierID ,colSJSupp = SDM.SuratJalanNo  " & vbCrLf &
                              " 		FROM   dbo.PO_Master POM     " & vbCrLf &
                              "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                        AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID         " & vbCrLf &
                              "                                               AND KD.PoNo = POD.PONo     " & vbCrLf &
                              "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                               AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf

            ls_SQL = ls_SQL + "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                    AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                     AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                     AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "                                                     AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                     AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                     AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "             LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf

            ls_SQL = ls_SQL + "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              " 		WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' )Header     " & vbCrLf &
                              "   UNION all   " & vbCrLf &
                              " 	--PackingList udah ada " & vbCrLf &
                              " 	SELECT DISTINCT * FROM ( " & vbCrLf &
                              " 	--PackingList udah ada " & vbCrLf &
                              "         SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf

            ls_SQL = ls_SQL + "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = 0 ,      " & vbCrLf &
                              "          colsuppdelqty = 0 ,      " & vbCrLf &
                              "          colpasigoodrec = 0,      " & vbCrLf &
                              "          colpasidefectrec = 0,      " & vbCrLf

            ls_SQL = ls_SQL + "          colpasiremaining = 0 ,      " & vbCrLf &
                              "          colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))), " & vbCrLf &
                              "          colremainingdelqty = ISNULL(PDD.DOQty,0),--ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PLD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(POD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID ,colSJSupp = SDM.SuratJalanNo  " & vbCrLf

            ls_SQL = ls_SQL + " 	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                    AND POM.PoNo = POD.PONo     " & vbCrLf &
                              "                                    AND POM.SupplierID = POD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     " & vbCrLf &
                              "                                           AND KD.PoNo = POD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "                                           AND KD.SupplierID = POD.SupplierID     " & vbCrLf &
                              "                                           AND KD.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf &
                              "                                           AND KD.KanbanNo = KM.KanbanNo     " & vbCrLf &
                              "                                           AND KD.SupplierID = KM.SupplierID     " & vbCrLf &
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf &
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PONo = SDD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND KD.KanbanNo = PRD.KanbanNo     " & vbCrLf &
                              "                                                 AND KD.SupplierID = PRD.SupplierID     " & vbCrLf &
                              "                                                 AND KD.PONo = PRD.PONo     " & vbCrLf &
                              "                                                 AND KD.PartNo = PRD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID     " & vbCrLf &
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo     " & vbCrLf &
                              "                                                 AND PRM.SupplierID = PRD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                            AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                            AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                            AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                            AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                            AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                            --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "        LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID     " & vbCrLf &
                              "                                AND KD.KanbanNo = PLD.KanbanNo     " & vbCrLf &
                              "                                --AND KD.SupplierID = PLD.SupplierID     " & vbCrLf

            ls_SQL = ls_SQL + "                                AND KD.PartNo = PLD.PartNo     " & vbCrLf &
                              "                                AND KD.PoNo = PLD.PoNo                                        " & vbCrLf &
                              "                                AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "                                AND PLD.SuratJalanNo = PDD.SuratJalanNo " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf &
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and SDD.DOQty <> 0 AND ISNULL(PLD.CartonNo,'') <> '' " & vbCrLf &
                              "    UNION ALL " & vbCrLf &
                              "    --Delivery PASI " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf &
                              "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colQtyBox = 0 ,      " & vbCrLf &
                              "          colsuppdelqty = 0 ,      " & vbCrLf &
                              "          colpasigoodrec = 0,      " & vbCrLf &
                              "          colpasidefectrec = 0,      " & vbCrLf &
                              "          colpasiremaining = 0 ,      " & vbCrLf &
                              "          colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))), " & vbCrLf &
                              "          colremainingdelqty = ISNULL(PDD.DOQty,0),--ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(PDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf

            ls_SQL = ls_SQL + "          colstsDO = ISNULL(PDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = ISNULL(PLD.CartonNo,'') ,     " & vbCrLf &
                              "          colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID ,colSJSupp = SDM.SuratJalanNo  " & vbCrLf &
                              " 	FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                            AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "  							                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PONo = SDD.PONo     " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo     " & vbCrLf &
                              "                                                AND SDM.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf &
                              "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID " & vbCrLf &
                              " 		 LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf &
                              "                                                AND KD.KanbanNo = PDD.KanbanNo     " & vbCrLf &
                              "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf &
                              "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf &
                              "                                                AND KD.PoNo = PDD.PoNo     " & vbCrLf &
                              "                                                AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier" & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf &
                              "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf &
                              "                                                --AND PDD.SupplierID = PDM.SupplierID     " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Master PLM ON PDD.AffiliateID = PLM.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND PDD.SuratJalanNo = PLM.SuratJalanNo      " & vbCrLf &
                              "                                             AND PDD.SupplierID = PLM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID      " & vbCrLf &
                              "                                 AND KD.KanbanNo = PLD.KanbanNo      " & vbCrLf &
                              "                                 AND KD.SupplierID = PLD.SupplierID      " & vbCrLf &
                              "                                 AND KD.PartNo = PLD.PartNo      " & vbCrLf &
                              "                                 AND KD.PoNo = PLD.PoNo " & vbCrLf &
                              "                                 AND PLD.SuratJalanNo = PLM.SuratJalanNo " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and PDD.DOQty <> 0  AND ISNULL(PLD.CartonNo,'') <> '')test " & vbCrLf &
                              "    UNION ALL   " & vbCrLf &
                              "    --PL Kosong " & vbCrLf
            ls_SQL = ls_SQL + " SELECT distinct AllowAccess,colno,colpono,colponos,colpokanban,colkanbanno,     " & vbCrLf &
                              " colkanbannos,colpartno,colpartname,coluom,colCls,colQtyBox,       " & vbCrLf &
                              " colsuppdelqty,colpasigoodrec,colpasidefectrec,colpasiremaining,       " & vbCrLf &
                              " colpasideliveryqty,colremainingdelqty,coldelqtybox,colstsDO ,      " & vbCrLf &
                              " colcartonno= CASE WHEN ROUND(colpasideliveryqty/colQtyBox,0) = 1 THEN 'C001' " & vbCrLf &
                              " WHEN ROUND(colpasideliveryqty/colQtyBox,0) >= 10 THEN 'C001-C0' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0)) " & vbCrLf &
                              " WHEN ROUND(colpasideliveryqty/colQtyBox,0) >= 10 THEN 'C001-C' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0)) " & vbCrLf &
                              " ELSE 'C001-C00' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0))END , " & vbCrLf &
                              " colcartonqty = round(colpasideliveryqty/colQtyBox,0),sortData, colsupp ,colSJSupp" & vbCrLf &
                              " FROM ( " & vbCrLf &
                              "    SELECT '0'AllowAccess,colno =  '' ,     " & vbCrLf &
                              "          colpono = '' ,      " & vbCrLf &
                              "          colponos = POM.PONo,    " & vbCrLf &
                              "          colpokanban = '' ,      " & vbCrLf &
                              "          colkanbanno = '' ,    " & vbCrLf &
                              "          colkanbannos = ISNULL(KD.KanbanNo, '') ,     " & vbCrLf

            ls_SQL = ls_SQL + "          colpartno = '' ,      " & vbCrLf &
                              "          colpartname = '' ,       " & vbCrLf &
                              "          coluom = '' ,      " & vbCrLf &
                              "          colCls = '' ,     " & vbCrLf &
                              "          colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "          colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "          colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),      " & vbCrLf &
                              "          colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),      " & vbCrLf &
                              "          colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(SDD.DOQty, 0)      " & vbCrLf &
                              "                                             - ( ISNULL(PRD.GoodRecQty, 0)      " & vbCrLf

            ls_SQL = ls_SQL + "                                                 + ISNULL(PRD.DefectRecQty, 0) ) )),0) ,      " & vbCrLf &
                              "          colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))), " & vbCrLf &
                              "          colremainingdelqty = ISNULL(PDD.DOQty,0) ,--ROUND(CONVERT(CHAR,ISNULL(PRD.GoodRecQty, 0) - CASE ISNULL(SDD.DOQty,0) WHEN 0 THEN CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))) ELSE CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) END,0),0) ,      " & vbCrLf &
                              "          coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                           WHEN 0 THEN 0      " & vbCrLf &
                              "                           ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                         END,0)),      " & vbCrLf &
                              "          colstsDO = ISNULL(SDD.PartNo,'') ,     " & vbCrLf &
                              "          colcartonno = '' ,     " & vbCrLf &
                              "          colcartonqty = 0,     " & vbCrLf &
                              "          sortData = 1, colsupp = KD.SupplierID ,colSJSupp = SDM.SuratJalanNo " & vbCrLf

            ls_SQL = ls_SQL + "   FROM   dbo.PO_Master POM      " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf &
                              " 										   AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND KD.KanbanNo = SDD.KanbanNo      " & vbCrLf &
                              "   							                    AND KD.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "                                                 AND KD.PartNo = SDD.PartNo      " & vbCrLf &
                              "                                                 AND KD.PONo = SDD.PONo      " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDM.SuratJalanNo = SDD.SuratJalanNo      " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo      " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND KD.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "                                                  AND KD.PartNo = PRD.PartNo      " & vbCrLf &
                              "                                                  AND KD.PONo = PRD.PONo      " & vbCrLf &
                              "                                                  AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf &
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID      " & vbCrLf &
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf &
                              "                                                  AND PRM.SupplierID = PRD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                             AND KD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                             AND KD.SupplierID = PDD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND KD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                             AND KD.PoNo = PDD.PoNo      " & vbCrLf &
                              "                                             AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo      " & vbCrLf &
                              "                                             --AND PDD.SupplierID = PDM.SupplierID      " & vbCrLf &
                              "                                     " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf

            ls_SQL = ls_SQL + "    WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' AND (Rtrim(PDD.PONO)+Rtrim(PDD.KanbanNo)+Rtrim(PDD.PartNo)) IN (" & Trim(pCombination) & ") " & vbCrLf &
                              "    )PLKosong " & vbCrLf &
                              "    ) data   " & vbCrLf &
                              "    ORDER BY colSJSupp asc, colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_AddRow(ByVal pPO As String, ByVal pKanban As String, ByVal pSJ As String, ByVal pPartNo As String, ByVal pCombination As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            ls_SQL = " SELECT *,NoUrut = CONVERT(CHAR, ROW_NUMBER() OVER (ORDER BY colponos)) FROM (    " & vbCrLf &
                  "  --header  " & vbCrLf &
                  "       select distinct   " & vbCrLf &
                  "       AllowAccess,colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colstsDO, colkanbannos, sortData )) ,       " & vbCrLf &
                  "       colpono, colponos, colpokanban, colkanbanno, colkanbannos, colpartno, colpartname, coluom,  " & vbCrLf &
                  "       colCls, colQtyBox, colsuppdelqty, colpasigoodrec, colpasidefectrec, colpasiremaining,  " & vbCrLf &
                  "       colpasideliveryqty, colremainingdelqty, coldelqtybox, colstsDO, colcartonno, colcartonqty, sortData, colsupp,colSJSupp  " & vbCrLf &
                  "       from(   " & vbCrLf &
                  "  	   SELECT distinct 0 AllowAccess,colno = '',     " & vbCrLf &
                  "              colpono = POM.PONo ,    " & vbCrLf &
                  "              colponos = POM.PONo,      " & vbCrLf

            ls_SQL = ls_SQL + "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'      " & vbCrLf &
                              "                                 ELSE 'YES' END ,      " & vbCrLf &
                              "              colkanbanno = ISNULL(KD.KanbanNo, '') ,      " & vbCrLf &
                              "              colkanbannos = ISNULL(KD.KanbanNo, '') ,      " & vbCrLf &
                              "              colpartno = POD.PartNo ,      " & vbCrLf &
                              "              colpartname = MP.PartName ,      " & vbCrLf &
                              "              coluom = UC.Description ,      " & vbCrLf &
                              "              colCls = UC.unitcls ,     " & vbCrLf &
                              "              colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,      " & vbCrLf &
                              "              colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,      " & vbCrLf &
                              "              colpasigoodrec = '',      " & vbCrLf

            ls_SQL = ls_SQL + "              colpasidefectrec = '',      " & vbCrLf &
                              "              colpasiremaining = '' ,      " & vbCrLf &
                              "              colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))), " & vbCrLf &
                              "              colremainingdelqty = ISNULL(PDD.DOQty,0), " & vbCrLf &
                              "              coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                               WHEN 0 THEN 0      " & vbCrLf &
                              "                               ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)      " & vbCrLf &
                              "                             END,0)),      " & vbCrLf &
                              "              colstsDO = ISNULL(SDD.PartNo,'') ,     " & vbCrLf &
                              "              colcartonno = '',     " & vbCrLf &
                              "              colcartonqty = '',     " & vbCrLf

            ls_SQL = ls_SQL + "              sortData = 0, colsupp = KD.SupplierID ,colSJSupp = ''  " & vbCrLf &
                              "  		FROM   dbo.PO_Master POM      " & vbCrLf &
                              "              LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                         AND POM.PoNo = POD.PONo      " & vbCrLf &
                              "                                         AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "              LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID          " & vbCrLf &
                              "                                                AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                                AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                                AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "              LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                                AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "              INNER JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  							FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo     " & vbCrLf &
                              "              LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID          " & vbCrLf &
                              "                                                     AND SDM.SupplierID = SDD.SupplierID         " & vbCrLf &
                              "              LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))     " & vbCrLf

            ls_SQL = ls_SQL + "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo   " & vbCrLf &
                              "              LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo          " & vbCrLf &
                              "              LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "              LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf

            ls_SQL = ls_SQL + "              LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "              LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf &
                              " 		WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' AND POM.AffiliateID = '" & Trim(Session("AFF")) & "' )Header     " & vbCrLf &
                              " UNION all    " & vbCrLf &
                              "  	--PackingList udah ada  " & vbCrLf &
                              "  	SELECT DISTINCT * FROM (  " & vbCrLf &
                              "  	--PackingList udah ada  " & vbCrLf &
                              "          SELECT '0'AllowAccess,colno =  '' ,      " & vbCrLf &
                              "           colpono = '' ,       " & vbCrLf &
                              "           colponos = POM.PONo,     " & vbCrLf &
                              "           colpokanban = '' ,       " & vbCrLf &
                              "           colkanbanno = '' ,     " & vbCrLf &
                              "           colkanbannos = ISNULL(KD.KanbanNo, '') ,      " & vbCrLf

            ls_SQL = ls_SQL + "           colpartno = '' ,       " & vbCrLf &
                              "           colpartname = '' ,        " & vbCrLf &
                              "           coluom = '' ,       " & vbCrLf &
                              "           colCls = '' ,      " & vbCrLf &
                              "           colQtyBox = 0 ,       " & vbCrLf &
                              "           colsuppdelqty = 0 ,       " & vbCrLf &
                              "           colpasigoodrec = 0,       " & vbCrLf &
                              "           colpasidefectrec = 0,       " & vbCrLf &
                              "           colpasiremaining = 0 ,       " & vbCrLf &
                              "           colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),  " & vbCrLf &
                              "           colremainingdelqty = ISNULL(PDD.DOQty,0),       " & vbCrLf

            ls_SQL = ls_SQL + "           coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)       " & vbCrLf &
                              "                            WHEN 0 THEN 0       " & vbCrLf &
                              "                            ELSE ISNULL(PLD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)       " & vbCrLf &
                              "                          END,0)),       " & vbCrLf &
                              "           colstsDO = ISNULL(POD.PartNo,'') ,      " & vbCrLf &
                              "           colcartonno = ISNULL(PLD.CartonNo,'') ,      " & vbCrLf &
                              "           colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),      " & vbCrLf &
                              "           sortData = 1, colsupp = KD.SupplierID ,colSJSupp = ''  " & vbCrLf &
                              "  	FROM   dbo.PO_Master POM       " & vbCrLf &
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf &
                              "                                     AND POM.PoNo = POD.PONo      "

            ls_SQL = ls_SQL + "                                     AND POM.SupplierID = POD.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID      " & vbCrLf &
                              "                                            AND KD.PoNo = POD.PONo      " & vbCrLf &
                              "                                            AND KD.SupplierID = POD.SupplierID      " & vbCrLf &
                              "                                            AND KD.PartNo = POD.PartNo      " & vbCrLf &
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID      " & vbCrLf &
                              "                                            AND KD.KanbanNo = KM.KanbanNo      " & vbCrLf &
                              "                                            AND KD.SupplierID = KM.SupplierID      " & vbCrLf &
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf &
                              "          LEFT JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  						FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf

            ls_SQL = ls_SQL + "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo       " & vbCrLf &
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID         " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID        " & vbCrLf &
                              "          LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))     " & vbCrLf &
                              "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf

            ls_SQL = ls_SQL + "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo   " & vbCrLf &
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo        " & vbCrLf &
                              "         LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID      " & vbCrLf &
                              "                                 AND KD.KanbanNo = PLD.KanbanNo         " & vbCrLf &
                              "                                 AND KD.PartNo = PLD.PartNo      " & vbCrLf &
                              "                                 AND KD.PoNo = PLD.PoNo                                         " & vbCrLf &
                              "                                 AND PLD.SuratJalanNo = PDD.SuratJalanNo  " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls      " & vbCrLf

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID      " & vbCrLf &
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and SDD.DOQty <> 0 AND ISNULL(PLD.CartonNo,'') <> '' AND POM.AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                              "    UNION ALL " & vbCrLf &
                              "    --Delivery PASI " & vbCrLf &
                              " SELECT '0'AllowAccess,colno =  '' ,      " & vbCrLf &
                              "           colpono = '' ,       " & vbCrLf &
                              "           colponos = POM.PONo,     " & vbCrLf &
                              "           colpokanban = '' ,       " & vbCrLf &
                              "           colkanbanno = '' ,     " & vbCrLf &
                              "           colkanbannos = ISNULL(KD.KanbanNo, '') ,      " & vbCrLf &
                              "           colpartno = '' ,       " & vbCrLf &
                              "           colpartname = '' ,        " & vbCrLf &
                              "           coluom = '' ,       " & vbCrLf &
                              "           colCls = '' ,      " & vbCrLf

            ls_SQL = ls_SQL + "           colQtyBox = 0 ,       " & vbCrLf &
                              "           colsuppdelqty = 0 ,       " & vbCrLf &
                              "           colpasigoodrec = 0,       " & vbCrLf &
                              "           colpasidefectrec = 0,       " & vbCrLf &
                              "           colpasiremaining = 0 ,       " & vbCrLf &
                              "           colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),  " & vbCrLf &
                              "           colremainingdelqty = ISNULL(PDD.DOQty,0), " & vbCrLf &
                              "           coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)       " & vbCrLf &
                              "                            WHEN 0 THEN 0       " & vbCrLf &
                              "                            ELSE ISNULL(PDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)       " & vbCrLf &
                              "                          END,0)),       " & vbCrLf

            ls_SQL = ls_SQL + "           colstsDO = ISNULL(PDD.PartNo,'') ,      " & vbCrLf &
                              "           colcartonno = ISNULL(PLD.CartonNo,'') ,      " & vbCrLf &
                              "           colcartonqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PLD.CartonQty,0)))),      " & vbCrLf &
                              "           sortData = 1, colsupp = KD.SupplierID ,colSJSupp =''  " & vbCrLf &
                              "  	FROM   dbo.PO_Master POM       " & vbCrLf &
                              "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID       " & vbCrLf &
                              "                                      AND POM.PoNo = POD.PONo       " & vbCrLf &
                              "                                      AND POM.SupplierID = POD.SupplierID       " & vbCrLf &
                              "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID       " & vbCrLf &
                              "                                             AND KD.PoNo = POD.PONo       " & vbCrLf &
                              "                                             AND KD.SupplierID = POD.SupplierID       " & vbCrLf

            ls_SQL = ls_SQL + "                                             AND KD.PartNo = POD.PartNo       " & vbCrLf &
                              "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID       " & vbCrLf &
                              "                                             AND KD.KanbanNo = KM.KanbanNo       " & vbCrLf &
                              "                                             AND KD.SupplierID = KM.SupplierID       " & vbCrLf &
                              "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode       " & vbCrLf &
                              "           LEFT JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  						FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo       " & vbCrLf

            ls_SQL = ls_SQL + "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID          " & vbCrLf &
                              "                                                 AND SDM.SupplierID = SDD.SupplierID      " & vbCrLf &
                              "  		 LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))     " & vbCrLf &
                              "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo  " & vbCrLf &
                              "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID      " & vbCrLf &
                              "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo       " & vbCrLf

            ls_SQL = ls_SQL + "           LEFT JOIN dbo.PLPASI_Master PLM ON PDD.AffiliateID = PLM.AffiliateID       " & vbCrLf &
                              "                                              AND PDD.SuratJalanNo = PLM.SuratJalanNo       " & vbCrLf &
                              "                                              AND PDD.SupplierID = PLM.SupplierID       " & vbCrLf &
                              "           LEFT JOIN dbo.PLPASI_Detail PLD ON KD.AffiliateID = PLD.AffiliateID       " & vbCrLf &
                              "                                  AND KD.KanbanNo = PLD.KanbanNo       " & vbCrLf &
                              "                                  AND KD.SupplierID = PLD.SupplierID       " & vbCrLf &
                              "                                  AND KD.PartNo = PLD.PartNo       " & vbCrLf &
                              "                                  AND KD.PoNo = PLD.PoNo  " & vbCrLf &
                              "                                  AND PLD.SuratJalanNo = PLM.SuratJalanNo  " & vbCrLf &
                              "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo       " & vbCrLf &
                              "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls       " & vbCrLf

            ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID       " & vbCrLf &
                              "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID       " & vbCrLf &
                              "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf &
                              "    WHERE PLD.SuratJalanNo = '" & Trim(pSJ) & "' and PDD.DOQty <> 0  AND ISNULL(PLD.CartonNo,'') <> '' AND POM.AffiliateID = '" & Trim(Session("AFF")) & "')test " & vbCrLf &
                              "    UNION ALL   " & vbCrLf &
                              "    --PL Kosong " & vbCrLf
            ls_SQL = ls_SQL + " SELECT distinct AllowAccess,colno,colpono,colponos,colpokanban,colkanbanno,      " & vbCrLf &
                              "  colkanbannos,colpartno,colpartname,coluom,colCls,colQtyBox,        " & vbCrLf &
                              "  colsuppdelqty,colpasigoodrec,colpasidefectrec,colpasiremaining,        " & vbCrLf &
                              "  colpasideliveryqty,colremainingdelqty,coldelqtybox,colstsDO ,       " & vbCrLf &
                              "  colcartonno= CASE WHEN ROUND(colpasideliveryqty/colQtyBox,0) = 1 THEN 'C001'  " & vbCrLf &
                              "  WHEN ROUND(colpasideliveryqty/colQtyBox,0) >= 10 THEN 'C001-C0' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0))  " & vbCrLf &
                              "  WHEN ROUND(colpasideliveryqty/colQtyBox,0) >= 10 THEN 'C001-C' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0))  " & vbCrLf &
                              "  ELSE 'C001-C00' + CONVERT(CHAR(5),round(colpasideliveryqty/colQtyBox,0))END ,  " & vbCrLf &
                              "  colcartonqty = round(colpasideliveryqty/colQtyBox,0),sortData, colsupp ,colSJSupp " & vbCrLf &
                              "  FROM (  " & vbCrLf

            ls_SQL = ls_SQL + "     SELECT '0'AllowAccess,colno =  '' ,      " & vbCrLf &
                              "           colpono = '' ,       " & vbCrLf &
                              "           colponos = POM.PONo,     " & vbCrLf &
                              "           colpokanban = '' ,       " & vbCrLf &
                              "           colkanbanno = '' ,     " & vbCrLf &
                              "           colkanbannos = ISNULL(KD.KanbanNo, '') ,      " & vbCrLf &
                              "           colpartno = '' ,       " & vbCrLf &
                              "           colpartname = '' ,        " & vbCrLf &
                              "           coluom = '' ,       " & vbCrLf &
                              "           colCls = '' ,      " & vbCrLf &
                              "           colQtyBox = ROUND(CONVERT(CHAR, ISNULL(POD.POQtyBox,MPM.QtyBox)),0) ,       " & vbCrLf

            ls_SQL = ls_SQL + "           colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(SDD.DOQty, 0),0),0) ,       " & vbCrLf &
                              "           colpasigoodrec = '',       " & vbCrLf &
                              "           colpasidefectrec = '',       " & vbCrLf &
                              "           colpasiremaining = '' ,       " & vbCrLf &
                              "           colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),  " & vbCrLf &
                              "           colremainingdelqty = ISNULL(PDD.DOQty,0) , " & vbCrLf &
                              "           coldelqtybox = CEILING(CONVERT(CHAR,CASE ISNULL(POD.POQtyBox,MPM.QtyBox)       " & vbCrLf &
                              "                            WHEN 0 THEN 0       " & vbCrLf &
                              "                            ELSE ISNULL(SDD.DOQty, 0) / ISNULL(POD.POQtyBox,MPM.QtyBox)       " & vbCrLf &
                              "                          END,0)),       " & vbCrLf &
                              "           colstsDO = ISNULL(SDD.PartNo,'') ,      " & vbCrLf

            ls_SQL = ls_SQL + "           colcartonno = '' ,      " & vbCrLf &
                              "           colcartonqty = 0,      " & vbCrLf &
                              "           sortData = 1, colsupp = KD.SupplierID ,colSJSupp = '' " & vbCrLf &
                              "    FROM   dbo.PO_Master POM       " & vbCrLf &
                              "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID       " & vbCrLf &
                              "                                      AND POM.PoNo = POD.PONo       " & vbCrLf &
                              "                                      AND POM.SupplierID = POD.SupplierID       " & vbCrLf &
                              "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   " & vbCrLf &
                              "  										   AND KD.PoNo = POD.PONo       " & vbCrLf &
                              "                                             AND KD.SupplierID = POD.SupplierID       " & vbCrLf &
                              "                                             AND KD.PartNo = POD.PartNo       " & vbCrLf

            ls_SQL = ls_SQL + "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID       " & vbCrLf &
                              "                                             AND KD.KanbanNo = KM.KanbanNo       " & vbCrLf &
                              "                                             AND KD.SupplierID = KM.SupplierID       " & vbCrLf &
                              "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode       " & vbCrLf &
                              "           LEFT JOIN (SELECT SupplierID, AffiliateID, PONo, KanbanNo, PartNo, SUM(DOQty) DOQty   " & vbCrLf &
                              "  						FROM DOSupplier_Detail GROUP BY SupplierID, AffiliateID, PONo, KanbanNo, PartNo) SDD ON KD.AffiliateID = SDD.AffiliateID     " & vbCrLf &
                              "                                                    AND KD.KanbanNo = SDD.KanbanNo     " & vbCrLf &
                              "                                                    AND KD.PONo = SDD.PONo     " & vbCrLf &
                              "                                                    AND KD.SupplierID = SDD.SupplierID     " & vbCrLf &
                              "                                                    AND KD.PartNo = SDD.PartNo   " & vbCrLf &
                              "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID           " & vbCrLf

            ls_SQL = ls_SQL + "                                                  AND SDM.SupplierID = SDD.SupplierID            " & vbCrLf &
                              "           LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))     " & vbCrLf &
                              "              			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD     " & vbCrLf &
                              "                                                 ON SDD.AffiliateID = PDD.AffiliateID      " & vbCrLf &
                              "                                                 AND SDD.KanbanNo = PDD.KanbanNo      " & vbCrLf &
                              "                                                 AND SDD.SupplierID = PDD.SupplierID      " & vbCrLf &
                              "                                                 AND SDD.PartNo = PDD.PartNo      " & vbCrLf &
                              "                                                 AND SDD.PoNo = PDD.PoNo   " & vbCrLf &
                              "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID       " & vbCrLf &
                              "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo       " & vbCrLf

            ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo       " & vbCrLf &
                "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf &
                              "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls       " & vbCrLf &
                              "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID       " & vbCrLf &
                              "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID       " & vbCrLf &
                              "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf

            ls_SQL = ls_SQL + "    WHERE PDD.SuratJalanNo = '" & Trim(pSJ) & "' AND POM.AffiliateID = '" & Trim(Session("AFF")) & "' " & vbCrLf

            If pCombination <> "" Then ls_SQL = ls_SQL + "   AND (Rtrim(PDD.PONO)+Rtrim(PDD.KanbanNo)+Rtrim(PDD.PartNo)) IN (" & Trim(pCombination) & ") " & vbCrLf

            ls_SQL = ls_SQL + "    )PLKosong " & vbCrLf &
                                          "    ) data   " & vbCrLf &
                                          "    ORDER BY colSJSupp asc, colstsDO ASC, colkanbannos asc, sortData ASC  " & vbCrLf &
                                          "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 300
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                lblStatus.ForeColor = Color.White
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_IsiInvoice(ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim ls_SQL1 As String = ""
        Dim ls_HT As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()


            ls_SQL = "  SELECT InvoiceNo = ISNULL(InvoiceNo,''), " & vbCrLf &
                  " 		SuratJalanNo = ISNULL(SuratJalanNo,''), " & vbCrLf &
                  " 		DriverName = ISNULL(DriverName,''), " & vbCrLf &
                  " 		DriverContact = ISNULL(DriverContact,''), " & vbCrLf &
                  " 		NoPol = ISNULL(NoPol,''), " & vbCrLf &
                  " 		JenisArmada = ISNULL(JenisArmada,''), " & vbCrLf &
                  " 		HT_Cls = ISNULL(HT_Cls,'0') " & vbCrLf &
                  "  FROM DOPASI_Master WHERE SuratJalanNo = '" & pSJ & "' " & vbCrLf &
                  "  AND AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                  "  "

            ls_SQL1 = "  SELECT InvoiceDate = ISNULL(InvoiceDate, DeliveryDate) " & vbCrLf &
                  "  FROM PLPASI_Master WHERE SuratJalanNo = '" & pSJ & "' " & vbCrLf &
                  "  AND AffiliateID = '" & Trim(Session("AFF")) & "'" & vbCrLf &
                  "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim sqlDA1 As New SqlDataAdapter(ls_SQL1, sqlConn)
            Dim ds As New DataSet
            Dim ds1 As New DataSet
            sqlDA.Fill(ds)
            sqlDA1.Fill(ds1)

            If ds.Tables(0).Rows.Count > 0 Then
                Try
                    With ds.Tables(0)
                        txtInvoiceNo.Text = Trim(.Rows(0).Item("InvoiceNo"))
                        dtInvoiceDate.Text = Trim(.Rows(0).Item("SuratJalanNo"))
                        txtsuratjalanno.Text = Trim(.Rows(0).Item("SuratJalanNo"))
                        txtdrivername.Text = Trim(.Rows(0).Item("DriverName"))
                        txtdrivercontact.Text = Trim(.Rows(0).Item("DriverContact"))
                        txtnopol.Text = Trim(.Rows(0).Item("NoPol"))
                        txtjenisarmada.Text = Trim(.Rows(0).Item("JenisArmada"))
                        ls_HT = Trim(.Rows(0).Item("HT_Cls"))
                        HF.Set("HTcls", ls_HT)
                    End With
                Catch ex As Exception

                End Try
            End If

            If ds1.Tables(0).Rows.Count > 0 Then
                Try
                    With ds1.Tables(0)
                        Dim dtInvoice As String = Trim(.Rows(0).Item("InvoiceDate"))
                        If dtInvoice = "" Then
                            dtInvoiceDate.JSProperties("cpdtinv") = dtInvoice
                            dtInvoiceDate.Date = Now
                        Else
                            dtInvoiceDate.JSProperties("cpdtinv") = dtInvoice
                            dtInvoiceDate.Date = dtInvoice
                        End If

                    End With
                Catch ex As Exception

                End Try
            End If
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_SaveMaster(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pDeliveryDate As String, ByVal pPIC As String,
                            ByVal pjenisArmada As String, ByVal pDriverName As String, ByVal pDriverContact As String, ByVal pNopol As String, ByVal pTotalBox As String,
                            ByVal pInvoiceNo As String, ByVal pInvoiceDate As String, ByVal pFromDel As String, ByVal pToDel As String, ByVal pInsu As String, ByVal pViaDel As String, ByVal pAboutDel As String,
                            ByVal pPrivilege As String, ByVal pVessel As String, ByVal pAWB As String, ByVal pPayTerms As String, ByVal pOnAbout As String, ByVal pContainerNo As String,
                            ByVal pRemarks As String, ByVal pPlace As String, ByVal pCommercial As String)
        Dim ls_sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As New TransactionScope
                'save master
                ls_sql = Save_Master(pSjno, pSupplierID, pAffiliateID, pDeliveryDate, pPIC, pjenisArmada, pDriverName, pDriverContact, pNopol, pTotalBox, pInvoiceNo, pInvoiceDate,
                                    pFromDel, pToDel, pInsu, pViaDel, pAboutDel, pPrivilege, pVessel, pAWB, pPayTerms, pOnAbout, pContainerNo, pRemarks, pPlace, pCommercial)
                Dim sqlComm As New SqlCommand(ls_sql, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlTran.Complete()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_SaveDetail(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String)

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Try
            Dim iLoop As Long = 0, jLoop As Long = 0
            Dim ls_UserID As String = ""

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("SaveDetail")
                    If Grid.VisibleRowCount = 0 Then
                        ls_MsgID = "6011"
                        Call clsMsg.DisplayMessage(lblerrmessage, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        Session("ZZ010Msg") = lblerrmessage.Text
                        Exit Sub
                    End If
                    With Grid
                        For iLoop = 0 To Grid.VisibleRowCount - 1
                            If .GetRowValues(iLoop, "colcartonno").ToString() <> "" Then
                                ls_SQL = Update_Detail(pSjno, .GetRowValues(iLoop, "colsupp").ToString(), pAffiliateID, .GetRowValues(iLoop, "colponos").ToString(),
                                    .GetRowValues(iLoop, "colkanbannos").ToString(),
                                    .GetRowValues(iLoop, "colstsDO").ToString(),
                                    .GetRowValues(iLoop, "colpasideliveryqty"),
                                    .GetRowValues(iLoop, "colcartonno").ToString(),
                                    .GetRowValues(iLoop, "colcartonqty"),
                                    .GetRowValues(iLoop, "colSJSupp").ToString())
                                ls_MsgID = "1002"

                                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            End If
                        Next iLoop
                        sqlTran.Commit()
                        Call clsMsg.DisplayMessage(lblerrmessage, ls_MsgID, clsMessage.MsgType.InformationMessage)
                        If lblerrmessage.Text = "[] " Then lblerrmessage.Text = ""
                        Session("ZZ010Msg") = lblerrmessage.Text
                    End With
                End Using

                sqlConn.Close()


            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Function uf_SumQty(ByVal pSJ As String, ByVal pAffiliate As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = ""
            'ls_SQL = ls_SQL + " SELECT  coldelqtybox =  isnull(CEILING(CONVERT(CHAR, SUM(coldelqtybox),0)),0) " & vbCrLf & _
            '                  " FROM    ( SELECT    coldelqtybox = CASE MPM.QtyBox " & vbCrLf & _
            '                  "                                      WHEN 0 THEN 0 " & vbCrLf & _
            '                  "                                      ELSE COALESCE(SDD.DOQty, PDD.DOQty) / MPM.QtyBox " & vbCrLf & _
            '                  "                                    END " & vbCrLf & _
            '                  "           FROM      dbo.PO_Master POM " & vbCrLf & _
            '                  "                     LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
            '                  "                                                AND POM.PoNo = POD.PONo " & vbCrLf & _
            '                  "                                                AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
            '                  "                     LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
            '                  "                                                       AND KD.PoNo = POD.PONo "

            'ls_SQL = ls_SQL + "                                                       AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
            '                  "                                                       AND KD.PartNo = POD.PartNo " & vbCrLf & _
            '                  "                     LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
            '                  "                                                            AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
            '                  "                                                            AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
            '                  "                                                            AND KD.PartNo = SDD.PartNo " & vbCrLf & _
            '                  "                                                            AND KD.PONo = SDD.PONo " & vbCrLf & _
            '                  "                     LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
            '                  "        LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
            '                  " LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
            '                  "                                             AND KD.KanbanNo = PDD.KanbanNo  " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = PDD.SupplierID  " & vbCrLf & _
            '                  "                                             AND KD.PartNo = PDD.PartNo  " & vbCrLf & _
            '                  "                                             AND KD.PoNo = PDD.PoNo  " & vbCrLf & _
            '                  "                                             AND PDD.SuratJalanNoSupplier = SDD.SuratJalanNo " & vbCrLf & _
            '                  "          "

            'If pSJ = "" Then
            '    ls_SQL = ls_SQL + "  WHERE  POM.PONo IN (" & pPO & ") " & vbCrLf & _
            '                      "         AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + "  WHERE PDD.SuratJalanNo = '" & pSJ & "' " & vbCrLf
            'End If
            'ls_SQL = ls_SQL + "  ) Box " & vbCrLf

            ls_SQL = "SELECT ISNULL(TotalBox,0) TotalBox FROM DOPASI_Master WHERE SuratJalanNo = '" & pSJ & "' AND AffiliateID = '" & pAffiliate & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            uf_SumQty = ds.Tables(0).Rows(0)("TotalBox")
            sqlConn.Close()


        End Using
    End Function

    Private Function Save_Master(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pDeliveryDate As String, ByVal pPIC As String,
                            ByVal pjenisArmada As String, ByVal pDriverName As String, ByVal pDriverContact As String, ByVal pNopol As String, ByVal pTotalBox As String,
                            ByVal pInvoiceNo As String, ByVal pInvoiceDate As String, ByVal pFromDel As String, ByVal pToDel As String, ByVal pInsu As String, ByVal pViaDel As String, ByVal pAboutDel As String,
                            ByVal pPrivilege As String, ByVal pVessel As String, ByVal pAWB As String, ByVal pPayTerms As String, ByVal pOnAbout As String, ByVal pContainerNo As String,
                            ByVal pRemarks As String, ByVal pPlace As String, ByVal pCommercial As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM PLPASI_Master WHERE SuratJalanNo = '" & pSjno & "' AND SupplierID = '" & pSupplierID & "' AND AffiliateID = '" & pAffiliateID & "') " & vbCrLf &
                          " BEGIN " & vbCrLf &
                          " UPDATE dbo.PLPASI_Master " & vbCrLf &
                          " SET DeliveryDate ='" & pDeliveryDate & "' " & vbCrLf &
                          " 	,PIC ='" & pPIC & "'" & vbCrLf &
                          " 	,JenisArmada ='" & pjenisArmada & "' " & vbCrLf &
                          " 	,DriverName ='" & pDriverName & "' " & vbCrLf &
                          " 	,Commercial ='" & pCommercial & "' " & vbCrLf &
                          "     ,DriverContact ='" & pDriverContact & "' " & vbCrLf &
                          "     ,NoPol ='" & pNopol & "'" & vbCrLf &
                          "     ,TotalBox ='" & pTotalBox & "'" & vbCrLf &
                          "     ,[FromDelivery] = '" & pFromDel & "' " & vbCrLf &
                          "     ,[ToDelivery] ='" & pToDel & "' " & vbCrLf &
                          "     ,[InsurancePolicy] = '" & pInsu & "' " & vbCrLf &
                          "     ,[ViaDelivery] = '" & pViaDel & "' " & vbCrLf &
                          "     ,[AboutDelivery] ='" & pAboutDel & "' " & vbCrLf &
                          "     ,[Privilege] = '" & pPrivilege & "' " & vbCrLf &
                          "     ,[Vessel] = '" & pVessel & "' " & vbCrLf &
                          "     ,[AWBBLNo] = '" & pAWB & "' " & vbCrLf &
                          "     ,[PaymentTerms] = '" & pPayTerms & "' " & vbCrLf &
                          "     ,[OnAbout] = '" & pOnAbout & "' " & vbCrLf &
                          "     ,[ContainerNo] = '" & pContainerNo & "' " & vbCrLf &
                          "     ,[Remarks] = '" & pRemarks & "' " & vbCrLf &
                          "     ,[Place] = '" & pPlace & "' " & vbCrLf &
                          "     ,InvoiceNo ='" & pInvoiceNo & "' " & vbCrLf &
                          "     ,InvoiceDate ='" & pInvoiceDate & "' " & vbCrLf &
                          "     ,UpdateDate = GETDATE() " & vbCrLf

        ls_sql = ls_sql + "     ,UpdateUser ='" & pPIC & "' " & vbCrLf &
                          " WHERE SuratJalanNo = '" & pSjno & "'  " & vbCrLf &
                          "   AND SupplierID = '" & pSupplierID & "'  " & vbCrLf &
                          "   AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf &
                          " END " & vbCrLf &
                          " ELSE " & vbCrLf &
                          " BEGIN " & vbCrLf &
                          " INSERT INTO dbo.PLPASI_Master " & vbCrLf &
                          "         ( SuratJalanNo ,SupplierID ,AffiliateID ,DeliveryDate ,PIC ,JenisArmada ,DriverName, Commercial, " & vbCrLf &
                          "           DriverContact ,NoPol ,TotalBox,[FromDelivery],[ToDelivery],[InsurancePolicy],[ViaDelivery],[AboutDelivery],[Privilege],[Vessel]" & vbCrLf &
                          "           ,[AWBBLNo],[PaymentTerms],[OnAbout],[ContainerNo],[Remarks],[Place],InvoiceNo ,InvoiceDate ,EntryDate ,EntryUser  " & vbCrLf &
                          "         ) " & vbCrLf

        ls_sql = ls_sql + " VALUES  ( '" & pSjno & "' , -- SuratJalanNo - char(20) " & vbCrLf &
                          "           '" & pSupplierID & "' , -- SupplierID - char(20) " & vbCrLf &
                          "           '" & pAffiliateID & "' , -- AffiliateID - char(20) " & vbCrLf &
                          "           '" & pDeliveryDate & "' , -- DeliveryDate - date " & vbCrLf &
                          "           '" & pPIC & "' , -- PIC - char(15) " & vbCrLf &
                          "           '" & pjenisArmada & "' , -- JenisArmada - char(15) " & vbCrLf &
                          "           '" & pDriverName & "' , -- DriverName - char(15) " & vbCrLf &
                          "           '" & pCommercial & "' , -- Commercial - varchar(5) " & vbCrLf &
                          "           '" & pDriverContact & "' , -- DriverContact - char(15) " & vbCrLf &
                          "           '" & pNopol & "'  -- NoPol - char(10) " & vbCrLf &
                          "            ," & pTotalBox & " -- TotalBox - numeric " & vbCrLf &
                          "            , '" & pFromDel & "'" & vbCrLf &
                          "            ,'" & pToDel & "' " & vbCrLf &
                          "            ,'" & pInsu & "' " & vbCrLf &
                          "            ,'" & pViaDel & "' " & vbCrLf &
                          "            ,'" & pAboutDel & "' " & vbCrLf &
                          "            ,'" & pPrivilege & "' " & vbCrLf &
                          "            ,'" & pVessel & "' " & vbCrLf &
                          "            ,'" & pAWB & "' " & vbCrLf &
                          "            ,'" & pPayTerms & "' " & vbCrLf &
                          "            ,'" & pOnAbout & "' " & vbCrLf &
                          "            , '" & pContainerNo & "' " & vbCrLf &
                          "            ,'" & pRemarks & "' " & vbCrLf &
                          "            ,'" & pPlace & "' " & vbCrLf &
                          "            ,'" & pInvoiceNo & "' " & vbCrLf &
                          "            ,'" & pInvoiceDate & "' " & vbCrLf &
                          "            ,GETDATE() -- EntryDate - datetime " & vbCrLf

        ls_sql = ls_sql + "            ,'" & pPIC & "'  -- EntryUser - char(15) " & vbCrLf &
                          "         )	 " & vbCrLf &
                          " END " & vbCrLf

        Save_Master = ls_sql
    End Function

    Private Function Save_Detail(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pPOno As String, ByVal pPOKanbanCls As String,
                            ByVal pKanbanNo As String, ByVal pPartNo As String, ByVal pUnitCls As String, ByVal pDOqty As String, ByVal pCartonNo As String, ByVal pCartonQty As Integer)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM dbo.PLPASI_Detail WHERE SuratJalanNo = '" & pSjno & "' AND SupplierID = '" & pSupplierID & "' AND AffiliateID = '" & pAffiliateID & "' AND PONo = '" & pPOno & "' AND KanbanNo ='" & pKanbanNo & "' AND PartNo = '" & pPartNo & "' AND CartonNo='" & pCartonNo & "') " & vbCrLf &
                          " BEGIN " & vbCrLf &
                          " UPDATE dbo.PLPASI_Detail " & vbCrLf &
                          " SET POKanbanCls ='" & pPOKanbanCls & "', " & vbCrLf &
                          " 	UnitCls ='" & pUnitCls & "', " & vbCrLf &
                          " 	DOQty ='" & pDOqty & "', " & vbCrLf &
                          "     CartonNo = '" & Trim(pCartonNo) & "', " & vbCrLf &
                          "     CartonQty = " & pCartonQty & ",   " & vbCrLf &
                          "     POMOQ = '" & uf_GetMOQ(pPOno, pPartNo, pSupplierID, pAffiliateID) & "',   " & vbCrLf &
                          "     POQtyBox = '" & uf_GetQtybox(pPOno, pPartNo, pSupplierID, pAffiliateID) & "'   " & vbCrLf &
                          " WHERE SuratJalanNo = '" & pSjno & "'  " & vbCrLf &
                          "   AND SupplierID = '" & pSupplierID & "'  " & vbCrLf &
                          "   AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf &
                          "   AND KanbanNo ='" & pKanbanNo & "'" & vbCrLf &
                          "   AND PONo = '" & pPOno & "' " & vbCrLf

        ls_sql = ls_sql + "   AND PartNo = '" & pPartNo & "' " & vbCrLf &
                          "   AND CartonNo = '" & Trim(pCartonNo) & "' " & vbCrLf &
                          " END " & vbCrLf &
                          " ELSE " & vbCrLf &
                          " BEGIN " & vbCrLf &
                          " INSERT INTO dbo.PLPASI_Detail " & vbCrLf &
                          "         ( SuratJalanNo ,SupplierID ,AffiliateID ,PONo ,POKanbanCls , " & vbCrLf &
                          "           KanbanNo ,PartNo ,UnitCls ,DOQty,CartonNo,CartonQty, POMOQ, POQtyBox " & vbCrLf &
                          "         ) " & vbCrLf &
                          " VALUES  ( '" & pSjno & "' , -- SuratJalanNo - char(20) " & vbCrLf &
                          "           '" & pSupplierID & "' , -- SupplierID - char(20) " & vbCrLf &
                          "           '" & pAffiliateID & "' , -- AffiliateID - char(20) " & vbCrLf

        ls_sql = ls_sql + "           '" & pPOno & "' , -- PONo - char(20) " & vbCrLf &
                          "           '" & pPOKanbanCls & "' , -- POKanbanCls - char(1) " & vbCrLf &
                          "           '" & pKanbanNo & "' , -- KanbanNo - char(20) " & vbCrLf &
                          "           '" & pPartNo & "' , -- PartNo - char(25) " & vbCrLf &
                          "           '" & pUnitCls & "' , -- UnitCls - char(2) " & vbCrLf &
                          "           " & pDOqty & " ,  -- DOQty - numeric " & vbCrLf &
                          "           '" & pCartonNo & "' , -- CartonNo - char(25) " & vbCrLf &
                          "           " & pCartonQty & " , -- CartonQty - numeric " & vbCrLf &
                          "           '" & uf_GetMOQ(pPOno, pPartNo, pSupplierID, pAffiliateID) & "',   " & vbCrLf &
                          "           '" & uf_GetQtybox(pPOno, pPartNo, pSupplierID, pAffiliateID) & "'   " & vbCrLf &
                          "         ) " & vbCrLf &
                          " END " & vbCrLf

        Save_Detail = ls_sql
    End Function

    Private Function Update_Detail(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pPono As String,
                            ByVal pKanbanNo As String, ByVal pPartNo As String, ByVal pDOqty As String, ByVal pCartonNo As String, ByVal pCartonQty As Integer, ByVal pSuratJalanSupp As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM dbo.PLPASI_Detail WHERE SuratJalanNo = '" & pSjno & "' AND SupplierID = '" & pSupplierID & "' AND AffiliateID = '" & pAffiliateID & "' AND PONo = '" & pPono & "' AND KanbanNo ='" & pKanbanNo & "' AND PartNo = '" & pPartNo & "' AND CartonNo='" & pCartonNo & "') " & vbCrLf &
                          " BEGIN " & vbCrLf &
                          " UPDATE dbo.PLPASI_Detail " & vbCrLf &
                          " SET DOQty ='" & pDOqty & "', " & vbCrLf &
                          "     CartonNo = '" & Trim(pCartonNo) & "', " & vbCrLf &
                          "     CartonQty = " & pCartonQty & " ,  " & vbCrLf &
                          "     SuratJalanNoSupplier = '" & Trim(pSuratJalanSupp) & "', " & vbCrLf &
                          "     POMOQ = " & uf_GetMOQ(pPono, pPartNo, pSupplierID, pAffiliateID) & ",   " & vbCrLf &
                          "     POQtyBox = " & uf_GetQtybox(pPono, pPartNo, pSupplierID, pAffiliateID) & "   " & vbCrLf &
                          " WHERE SuratJalanNo = '" & pSjno & "'  " & vbCrLf &
                          "   AND SupplierID = '" & pSupplierID & "'  " & vbCrLf &
                          "   AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf &
                          "   AND KanbanNo ='" & pKanbanNo & "'" & vbCrLf &
                          "   AND PONo = '" & pPono & "' " & vbCrLf

        ls_sql = ls_sql + "   AND PartNo = '" & pPartNo & "' " & vbCrLf &
                          "   AND CartonNo = '" & Trim(pCartonNo) & "' " & vbCrLf &
                          " END " & vbCrLf &
                          " ELSE " & vbCrLf &
                          " BEGIN " & vbCrLf &
                          " INSERT INTO dbo.PLPASI_Detail " & vbCrLf &
                          "         ( SuratJalanNo ,SupplierID ,AffiliateID ,PONo , " & vbCrLf &
                          "           KanbanNo ,PartNo ,DOQty,CartonNo,CartonQty,SuratJalanNoSupplier, POMOQ, POQtyBox " & vbCrLf &
                          "         ) " & vbCrLf &
                          " VALUES  ( '" & pSjno & "' , -- SuratJalanNo - char(20) " & vbCrLf &
                          "           '" & pSupplierID & "' , -- SupplierID - char(20) " & vbCrLf &
                          "           '" & pAffiliateID & "' , -- AffiliateID - char(20) " & vbCrLf

        ls_sql = ls_sql + "           '" & pPono & "' , -- PONo - char(20) " & vbCrLf &
                          "           '" & pKanbanNo & "' , -- KanbanNo - char(20) " & vbCrLf &
                          "           '" & pPartNo & "' , -- PartNo - char(25) " & vbCrLf &
                          "           " & pDOqty & " ,  -- DOQty - numeric " & vbCrLf &
                          "           '" & pCartonNo & "' , -- CartonNo - char(25) " & vbCrLf &
                          "           " & pCartonQty & " , -- CartonQty - numeric " & vbCrLf &
                          "           '" & Trim(pSuratJalanSupp) & "' , -- CartonQty - numeric " & vbCrLf &
                          "           '" & uf_GetMOQ(pPono, pPartNo, pSupplierID, pAffiliateID) & "',   " & vbCrLf &
                          "           '" & uf_GetQtybox(pPono, pPartNo, pSupplierID, pAffiliateID) & "'   " & vbCrLf &
                          "         ) " & vbCrLf &
                          " END " & vbCrLf

        Update_Detail = ls_sql
    End Function

    Private Sub up_Delete(ByVal pSJ As String, ByVal pKanban As String, ByVal pPartNo As String, ByVal pCartonno As String)
        Dim ls_SQL As String = ""

        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " DELETE dbo.PLPASI_Detail " & vbCrLf &
                    " WHERE SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf &
                    " AND KanbanNo='" & pKanban & "' AND PartNo='" & pPartNo & "' AND AffiliateID = '" & Trim(Session("AFF")) & "' " & vbCrLf

            If pCartonno <> "" Then
                ls_SQL = ls_SQL + " AND Cartonno='" & pCartonno & "'"
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)


            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()

            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_DeletePL(ByVal pSJ As String)
        Call up_DeletePLMaster(pSJ)
        Call up_DeletePLDetail(pSJ)
    End Sub

    Private Sub up_DeletePLMaster(ByVal pSJ As String)
        Dim ls_SQL As String = ""

        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " DELETE dbo.PLPASI_Master " & vbCrLf &
                    " WHERE SuratJalanNo = '" & Trim(pSJ) & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)


            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()

            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_DeletePLDetail(ByVal pSJ As String)
        Dim ls_SQL As String = ""

        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " DELETE dbo.PLPASI_Detail " & vbCrLf &
                    " WHERE SuratJalanNo = '" & Trim(pSJ) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)


            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()

            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_ExistCartonQty(ByVal pSuratJalan As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pPONO As String, ByVal pKanbanNo As String, ByVal pPartNo As String)
        Dim ls_SQL As String = ""
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT ISNULL(CartonQty,0) CartonQty FROM PLPASI_Detail WHERE SuratJalanNo='" & Trim(pSuratJalan) & "' AND SupplierID='" & Trim(pSupplierID) & "' AND AffiliateID='" & pAffiliateID & "'" & vbCrLf &
                    " AND PONo='" & Trim(pPONO) & "' AND KanbanNo='" & Trim(pKanbanNo) & "' AND PartNo='" & Trim(pPartNo) & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                CartonQty = ds.Tables(0).Rows(0)("CartonQty")
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Function uf_GetMOQ(ByVal pPoNo As String, ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(a.POMOQ,b.MOQ) MOQ FROM dbo.PO_Detail a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf &
                     "WHERE PONo='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

    Private Function uf_GetQtybox(ByVal pPoNo As String, ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim Qty As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(a.POQtyBox,b.QtyBox) Qty FROM dbo.PO_Detail a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf &
                     "WHERE PONo='" + pPoNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                Qty = dt.Rows(0)("Qty")
            End If
        End Using
        Return Qty
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(clsGlobal.ConnectionString)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function
#End Region


End Class