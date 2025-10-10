Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports DevExpress.Web.ASPxMenu

Public Class AffKanbanCreate
    Inherits System.Web.UI.Page
    Private processAddNewRow As Boolean

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramSupplier As String
    Dim paramLocation As String

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "C02"
#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
                Session("MenuDesc") = "KANBAN CREATE"
            End If

            '================ init ===================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Dim param As String = ""
                If Not IsNothing(Request.QueryString("prm")) Or Not IsNothing(Session("KCR-Param")) Then
                    If Not IsNothing(Request.QueryString("prm")) Then
                        param = Request.QueryString("prm").ToString
                    ElseIf Not IsNothing(Session("KCR-Param")) Then
                        param = Session("KCR-Param")
                    End If

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                        'dtkanban.Value = Now
                        'dt1.Value = Now
                    Else
                        'If Request.QueryString("id") <> "URL" Then
                        Session.Remove("FilterKanbanNo")
                        Dim pkanbandate As String = Split(param, "|")(0)
                        Dim psuppID As String = Split(param, "|")(1)
                        Dim psuppname As String = Split(param, "|")(2)
                        Dim paffentrydate As String = Split(param, "|")(3)
                        Dim paffentryname As String = Split(param, "|")(4)
                        Dim paffappdate As String = Split(param, "|")(5)
                        Dim paffappname As String = Split(param, "|")(6)
                        Dim psuppappdate As String = Split(param, "|")(7)
                        Dim psuppappname As String = Split(param, "|")(8)
                        Dim pdt1 As Date = Split(param, "|")(9)
                        Dim pdt2 As Date = Split(param, "|")(10)
                        Dim pcbosupplier As String = Split(param, "|")(11)
                        Dim pcbolocation As String = Split(param, "|")(13)
                        Dim pLocation As String = Split(param, "|")(14)
                        Dim pAffcode As String = Split(param, "|")(17)
                        Dim pAffName As String = Split(param, "|")(18)
                        Dim pKanbanno As String = Split(param, "|")(21)

                        If psuppID <> "" Then btnsubmenu.Text = "BACK"
                        cbosupplier.Text = psuppID
                        txtsuppliername.Text = psuppname
                        cbolocation.Text = pcbolocation
                        txtlocation.Text = pLocation
                        dtkanban.Value = pkanbandate
                        txtaffiliateappdate.Text = paffappdate
                        txtaffiliateappname.Text = paffappname
                        txtaffiliateentrydate.Text = paffentrydate
                        txtaffiliateentryname.Text = paffentryname
                        txtsupplierapprovaldate.Text = psuppappdate
                        txtsupplierapprovalname.Text = psuppappname
                        txtaffiliatecode.Text = pAffcode
                        txtaffiliatename.Text = pAffName
                        Session("FilterKanbanNo") = pKanbanno
                        'dt1.Value = Format(pkanbandate, "MMM yyyy")

                        Call fillHeader()
                        Call up_GridLoad(pkanbandate, psuppID)
                        'Else

                        '    'parameter dari URL
                        '    Dim pkanbandate As String = Request.QueryString("t1")
                        '    Dim psuppID As String = Request.QueryString("t2")
                        '    Dim pDeliveryLocation As String = Request.QueryString("t3")

                        '    If psuppID <> "" Then btnsubmenu.Text = "BACK"
                        '    dtkanban.Value = pkanbandate
                        '    cbosupplier.Text = psuppID
                        '    cbolocation.Text = pDeliveryLocation

                        '    Call fillHeader()
                        '    Call up_GridLoad(pkanbandate, psuppID)
                        'End If

                    End If

                    btnsubmenu.Text = "BACK"
                Else
                    'Clear()s
                    'dt1.Value = Format(Now, "MMM yyyy")
                    'dtkanban.Value = Now
                    If (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Or Not IsNothing(Session("URL-PARAM")) Then
                        Session("M01Url") = Request.QueryString("Session")
                        Session("MenuDesc") = "KANBAN CREATE"
                        'parameter dari URL
                        Dim pkanbandate As Date = clsNotification.DecryptURL(Request.QueryString("t0"))
                        Dim psuppID As String = clsNotification.DecryptURL(Request.QueryString("t1"))
                        Dim pDeliveryLocation As String = clsNotification.DecryptURL(Request.QueryString("t2"))

                        Session.Remove("URL-PARAM")

                        If psuppID <> "" Then btnsubmenu.Text = "BACK"
                        dtkanban.Value = Format(pkanbandate, "dd MMM yyyy")
                        cbosupplier.Text = psuppID
                        cbolocation.Text = pDeliveryLocation
                        Session("URL-PARAM") = "~/AffKanban/AffKanbanCreate.aspx?id2=URL" & "&t0=" & clsNotification.EncryptURL(pkanbandate.Date) & "&t1=" & clsNotification.EncryptURL(psuppID) & _
                                                "&t2=" & clsNotification.EncryptURL(pDeliveryLocation.Trim) & "&Session=" & clsNotification.EncryptURL("~/AffKanban/AffKanbanList.aspx")
                        Call fillHeader()
                        Call up_GridLoad(pkanbandate, psuppID)
                    Else
                        If Session("KCR-ReportCode") <> "" Then
                            Session("M01Url") = Request.QueryString("Session")
                            Session("MenuDesc") = "KANBAN CREATE"

                            dtkanban.Text = Session("KCR-KanbanDate")
                            cbosupplier.Text = Session("KCR-SupplierCode")
                            cbolocation.Text = Replace(Session("KCR-DeliveryLocation"), "'", "")
                            txtaffiliatecode.Text = Session("KCR-AffiliateCode")
                            txtaffiliatename.Text = Session("KCR-AffiliateName")
                            txtlocation.Text = Session("KCR-DeliveryLocationName")

                            Call fillHeader()
                            Call up_GridLoad(Session("KCR-KanbanDate"), Session("KCR-SupplierCode"))

                            Session.Remove("YA030ReqNo")
                            Session.Remove("KCR-KanbanNo")
                            Session.Remove("KCR-ReportCode")
                            Session.Remove("KCR-KanbanDate")
                            Session.Remove("KCR-SupplierCode")
                            Session.Remove("KCR-DeliveryLocation")
                            Session.Remove("KCR-Form")
                        End If
                    End If
                    'dt1.Value = Now
                End If
            End If
            '================ init ===================


            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblerrmessage.Text = ""
            End If

            Call colorGrid()
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Dim param As String = ""

        If Not IsNothing(Session("URL-Param")) Then
            Response.Redirect("~/AffKanban/AffKanbanList.aspx")
        End If
        If btnsubmenu.Text = "BACK" Then
            If Not IsNothing(Request.QueryString("prm")) Or Not IsNothing(Session("KCR-Param")) Then
                param = ""
                If Not IsNothing(Request.QueryString("prm")) Then
                    param = Request.QueryString("prm").ToString
                ElseIf Not IsNothing(Session("KCR-Param")) Then
                    param = Session("KCR-Param")
                End If

            End If

            'Dim param As String = Request.QueryString("prm").ToString

            Dim pdt1 As Date = Split(param, "|")(9)
            Dim pdt2 As Date = Split(param, "|")(10)

            Dim pcbosupplier As String = Split(param, "|")(11)
            Dim psuppname As String = Split(param, "|")(12)

            Dim pcboAffiliate As String = Split(param, "|")(19)
            Dim ptxtAffiliate As String = Split(param, "|")(20)

            Dim pcbolocation As String = Split(param, "|")(15)
            Dim plocationname As String = Split(param, "|")(16)
            Dim pKanbanno As String = Split(param, "|")(21)

            Session.Remove("KCR-PARAM")

            Response.Redirect("~/AffKanban/AffKanbanList.aspx?prm=" + pdt1 + "|" + pdt2 + "|" + pcbosupplier + "|" + psuppname + "|" + pcbolocation + "|" + plocationname + "|" + pcboAffiliate + "|" + ptxtAffiliate + "|" + pKanbanno + "")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Public Sub btnclear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnclear.Click
        Clear()
        up_GridLoad(dtkanban.Value, "xxx")
    End Sub

    Private Sub Grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles Grid.CellEditorInitialize

        If (e.Column.FieldName = "cols" Or e.Column.FieldName = "colkanbanqty" Or e.Column.FieldName = "colcycle1" Or e.Column.FieldName = "colcycle2" Or e.Column.FieldName = "colcycle3" Or e.Column.FieldName = "colcycle4") Then
            e.Editor.ReadOnly = False
        Else
            e.Editor.ReadOnly = True
        End If
    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)

        Try
            Select Case pAction

                Case "gridload"
                    Call fillHeader()
                    Call up_GridLoad(dtkanban.Value, Trim(cbosupplier.Text))
                    If pAction = "" Then
                        If Grid.VisibleRowCount = 0 Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        Else
                            Grid.JSProperties("cpMessage") = ""
                            lblerrmessage.Text = ""
                        End If
                    End If
                    Call colorGrid()
                Case "approve"
                    Call fillHeader()
                Case "kosong"

            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Grid.FocusedRowIndex = -1

        Finally
            'If (Not IsNothing(Session("YA010Msg"))) Then Grid.JSProperties("cpMessage") = Session("YA010Msg") : Session.Remove("YA010Msg")
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try
    End Sub

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        Dim x As Integer = CInt(e.VisibleIndex.ToString())



        Dim pDeliveryQty As Double
        Dim pSupplierCapacity As Double

        If x > Grid.VisibleRowCount Then Exit Sub
        'e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        If e.DataColumn.FieldName = "colkanbanqty" Then
            pDeliveryQty = CDbl(e.GetValue("colkanbanqty"))
            pSupplierCapacity = CDbl(e.GetValue("colremainingsupplier"))
        End If

        With Grid
            If .VisibleRowCount > 0 Then
                If pDeliveryQty > pSupplierCapacity Then
                    If e.DataColumn.FieldName = "colkanbanqty" Then
                        e.Cell.BackColor = Color.HotPink
                    End If
                End If

            End If
        End With

    End Sub

    Protected Sub btnprintcard_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnprintcard.Click
        Session.Remove("KCR-KanbanNo")
        Session.Remove("KCR-ReportCode")
        Session.Remove("KCR-KanbanDate")
        Session.Remove("KCR-SupplierCode")
        Session.Remove("KCR-DeliveryLocation")
        Session.Remove("KCR-Form")
        Session.Remove("KCR-AffiliateCode")
        Session.Remove("KCR-AffiliateName")
        Session.Remove("KCR-DeliveryLocationName")
        Session.Remove("KCR-kanbanno")
        'Session.Remove("KCR-PARAM")

        If txtaffiliateappdate.Text <> "" Then
            Session("KCR-KanbanNo") = "'" + Trim(txtkanban1.Text) + "'"
            Session("KCR-KanbanDate") = Trim(dtkanban.Text)
            Session("KCR-SupplierCode") = Trim(cbosupplier.Text)
            Session("KCR-DeliveryLocation") = "'" & Trim(cbolocation.Text) & "'"
            Session("KCR-AffiliateCode") = Trim(txtaffiliatecode.Text)
            Session("KCR-AffiliateName") = Trim(txtaffiliatename.Text)
            Session("KCR-DeliveryLocationName") = Trim(txtlocation.Text)
            Session("KCR-ReportCode") = "KanbanCard"
            Session("KCR-Form") = "KanbanCreate"
            Session("KCR-kanbanno") = "'" + Session("FilterKanbanNo") + "'"
            If Not IsNothing(Request.QueryString("prm")) Then
                Session("KCR-PARAM") = Request.QueryString("prm").ToString
            End If
            Response.Redirect("~/AffKanban/ViewReport.aspx")
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "7011", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            fillHeader()
        End If

    End Sub

    Private Sub btnprintcycle_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnprintcycle.Click
        Session.Remove("KCR-KanbanNo")
        Session.Remove("KCR-ReportCode")
        Session.Remove("KCR-KanbanDate")
        Session.Remove("KCR-SupplierCode")
        Session.Remove("KCR-DeliveryLocation")
        Session.Remove("KCR-Form")
        Session.Remove("KCR-AffiliateCode")
        Session.Remove("KCR-AffiliateName")
        Session.Remove("KCR-DeliveryLocationName")
        Session.Remove("KCR-kanbanno")
        'Session.Remove("KCR-PARAM")

        If txtaffiliateappdate.Text <> "" Then
            Session("KCR-KanbanNo") = "'" + Trim(txtkanban1.Text) + "'"
            Session("KCR-KanbanDate") = Trim(dtkanban.Text)
            Session("KCR-SupplierCode") = Trim(cbosupplier.Text)
            Session("KCR-DeliveryLocation") = Trim(cbolocation.Text)
            Session("KCR-DeliveryLocationName") = Trim(txtlocation.Text)
            Session("KCR-AffiliateCode") = Trim(txtaffiliatecode.Text)
            Session("KCR-AffiliateName") = Trim(txtaffiliatename.Text)
            Session("KCR-ReportCode") = "KanbanCycle"
            Session("KCR-Form") = "KanbanCreate"
            Session("KCR-kanbanno") = "'" + Session("FilterKanbanNo") + "'"

            If Not IsNothing(Request.QueryString("prm")) Then
                Session("KCR-PARAM") = Request.QueryString("prm").ToString
            End If

            Response.Redirect("~/AffKanban/ViewReport.aspx")
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "7011", clsMessage.MsgType.ErrorMessage)
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            fillHeader()
        End If
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub colorGrid()
        Grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(12).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(13).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(14).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(15).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(16).CellStyle.BackColor = Drawing.Color.LightYellow

        Grid.VisibleColumns(1).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(2).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(3).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(4).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(5).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(6).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(7).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(8).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(9).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(10).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(11).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(17).CellStyle.BackColor = Drawing.Color.LightYellow

        txtkanban1.BackColor = Color.LightYellow
        txtkanban2.BackColor = Color.LightYellow
        txtkanban3.BackColor = Color.LightYellow
        txtkanban4.BackColor = Color.LightYellow
        txttime1.BackColor = Color.LightYellow
        txttime2.BackColor = Color.LightYellow
        txttime3.BackColor = Color.LightYellow
        txttime4.BackColor = Color.LightYellow

    End Sub

    Private Sub Clear()
        dtkanban.Text = ""
        cbosupplier.Text = ""
        txtsuppliername.Text = ""
        txtaffiliateappdate.Text = ""
        txtaffiliateappname.Text = ""
        txtaffiliateentrydate.Text = ""
        txtaffiliateentryname.Text = ""
        txtkanban1.Text = ""
        txtkanban2.Text = ""
        txtkanban3.Text = ""
        txtkanban4.Text = ""
        txttime1.Text = ""
        txttime2.Text = ""
        txttime3.Text = ""
        txttime4.Text = ""

    End Sub

    Private Sub fillHeader()
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        i = 0
        ls_sql = ""
        ls_sql = " SELECT   " & vbCrLf & _
                  " KM.supplierID, " & vbCrLf & _
                  " MSS.SupplierName, " & vbCrLf & _
                  " KM.kanbanDate, " & vbCrLf & _
                  " KM.KanbanCycle, " & vbCrLf & _
                  " ISNULL(KM.EntryUser, '') AS EntryUser , " & vbCrLf & _
                  "         case when isnull(KM.entrydate,'') = '' then '' else convert(char(19),convert(datetime,KM.entrydate),120) end AS EntryDate , " & vbCrLf & _
                  "         ISNULL(SupplierApproveUser, '') AS SupplierUser , " & vbCrLf & _
                  "         case when isnull(SupplierApproveDate,'') = '' then '' else convert(char(19),convert(datetime,SupplierApproveDate),120) end AS supplierDate , " & vbCrLf & _
                  "         ISNULL(AffiliateApproveUser, '') AS AffiliateUser , " & vbCrLf & _
                  "         case when isnull(affiliateApproveDate,'') = '' then '' else convert(char(19),convert(datetime,affiliateApproveDate),120) end AS AffiliateDate, " & vbCrLf & _
                  "         KanbanNo , " & vbCrLf & _
                  "         isnull(KM.DeliveryLocationCode,'') as LocationCode, " & vbCrLf & _
                  "         isnull(DeliveryLocationName, '') as LocationName " & vbCrLf & _
                  "         ,AffiliateName, AffiliateCode = KM.AffiliateID " & vbCrLf & _
                  "         ,ALAMAT = RTRIM(MSA.Address) + ' ' + RTRIM(MSA.City) + ' '+ RTRIM(MSA.PostalCode), " & vbCrLf


        ls_sql = ls_sql + "         CONVERT(CHAR(5),  ISNULL(CONVERT(DATETIME, KanbanTime),'00:00:00'), 114) AS KanbanTime " & vbCrLf & _
                          " FROM    kanban_Master KM " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Supplier MSS ON KM.SupplierID = MSS.SupplierID " & vbCrLf & _
                          " LEFT JOIN MS_DeliveryPlace MDP on KM.DeliveryLocationCode = MDP.DeliveryLocationCode And KM.AffiliateID = MDP.AffiliateID " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Affiliate MSA ON MSA.AffiliateID = KM.AffiliateID " & vbCrLf & _
                          " where KM.SupplierID = '" & Trim(cbosupplier.Text) & "' and kanbanDate = '" & (dtkanban.Text) & "' " & vbCrLf
        'If cbolocation.Text <> "" Then
        ls_sql = ls_sql + " and KM.DeliveryLocationCode = '" & Trim(cbolocation.Text) & "'"
        'End If
        If (Session("FilterKanbanNo") <> "" And Session("FilterKanbanNo") <> clsGlobal.gs_All) Then
            ls_sql = ls_sql + " AND KM.Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
        End If
        ls_sql = ls_sql + " order by KanbanNo,KanbanCycle "

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtsuppliername.Text = ds.Tables(0).Rows(0)("SupplierName")
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    txtaffiliateentrydate.Text = ds.Tables(0).Rows(i)("entryDate")
                    txtsupplierapprovaldate.Text = ds.Tables(0).Rows(i)("supplierDate")
                    txtaffiliateappdate.Text = ds.Tables(0).Rows(i)("AffiliateDate")

                    Grid.JSProperties("cpAEDate") = Format(ds.Tables(0).Rows(i)("entrydate"), "yyyy-MM-dd HH:mm:ss")
                    Grid.JSProperties("cpAEName") = ds.Tables(0).Rows(i)("entryuser")
                    Grid.JSProperties("cpASName") = ds.Tables(0).Rows(i)("SupplierUser")
                    Grid.JSProperties("cpAAName") = ds.Tables(0).Rows(i)("AffiliateUser")
                    Grid.JSProperties("cpaffcode") = ds.Tables(0).Rows(i)("Affiliatecode")
                    Grid.JSProperties("cpaffname") = ds.Tables(0).Rows(i)("Affiliatecode")
                    Grid.JSProperties("cplocationcode") = ds.Tables(0).Rows(i)("locationcode")
                    Grid.JSProperties("cplocationname") = ds.Tables(0).Rows(i)("locationname")

                    If ds.Tables(0).Rows(i)("AffiliateDate") <> "1/1/1900" Then Grid.JSProperties("cpAADate") = (ds.Tables(0).Rows(i)("AffiliateDate"))
                    If ds.Tables(0).Rows(i)("supplierDate") <> "1/1/1900" Then Grid.JSProperties("cpASDate") = (ds.Tables(0).Rows(i)("supplierDate"))

                    If ds.Tables(0).Rows(i)("AffiliateDate") <> "1/1/1900" Then txtaffiliateappdate.Text = (ds.Tables(0).Rows(i)("AffiliateDate"))
                    If ds.Tables(0).Rows(i)("supplierDate") <> "1/1/1900" Then txtsupplierapprovaldate.Text = (ds.Tables(0).Rows(i)("supplierDate"))


                    txtsupplierapprovaldate.Text = (ds.Tables(0).Rows(i)("supplierDate"))
                    txtaffiliateappdate.Text = (ds.Tables(0).Rows(i)("AffiliateDate"))
                    txtaffiliateentryname.Text = ds.Tables(0).Rows(i)("entryuser")
                    txtsupplierapprovalname.Text = ds.Tables(0).Rows(i)("SupplierUser")
                    txtaffiliateappname.Text = ds.Tables(0).Rows(i)("AffiliateUser")
                    txtaffiliatecode.Text = ds.Tables(0).Rows(i)("Affiliatecode")
                    txtaffiliatename.Text = ds.Tables(0).Rows(i)("Affiliatecode")
                    txtlocation.Text = ds.Tables(0).Rows(i)("LocationName")
                    cbolocation.Text = ds.Tables(0).Rows(i)("LocationCode")

                    'If i = 0 Then
                    '    Grid.JSProperties("cpKanban1") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime1") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 1 Then
                    '    Grid.JSProperties("cpKanban2") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime2") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 2 Then
                    '    Grid.JSProperties("cpKanban3") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime3") = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 3 Then
                    '    Grid.JSProperties("cpKanban4") = ds.Tables(0).Rows(i)("kanbanno")
                    '    Grid.JSProperties("cpTime4") = ds.Tables(0).Rows(i)("kanbantime")
                    'End If

                    If ds.Tables(0).Rows(i)("KanbanCycle") = "1" Or ds.Tables(0).Rows(i)("KanbanCycle") = "5" Or ds.Tables(0).Rows(i)("KanbanCycle") = "9" Or ds.Tables(0).Rows(i)("KanbanCycle") = "13" Or ds.Tables(0).Rows(i)("KanbanCycle") = "17" Then
                        Grid.JSProperties("cpKanban1") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime1") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "2" Or ds.Tables(0).Rows(i)("KanbanCycle") = "6" Or ds.Tables(0).Rows(i)("KanbanCycle") = "10" Or ds.Tables(0).Rows(i)("KanbanCycle") = "14" Or ds.Tables(0).Rows(i)("KanbanCycle") = "18" Then
                        Grid.JSProperties("cpKanban2") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime2") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "3" Or ds.Tables(0).Rows(i)("KanbanCycle") = "7" Or ds.Tables(0).Rows(i)("KanbanCycle") = "11" Or ds.Tables(0).Rows(i)("KanbanCycle") = "15" Or ds.Tables(0).Rows(i)("KanbanCycle") = "19" Then
                        Grid.JSProperties("cpKanban3") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime3") = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "4" Or ds.Tables(0).Rows(i)("KanbanCycle") = "8" Or ds.Tables(0).Rows(i)("KanbanCycle") = "12" Or ds.Tables(0).Rows(i)("KanbanCycle") = "16" Or ds.Tables(0).Rows(i)("KanbanCycle") = "20" Then
                        Grid.JSProperties("cpKanban4") = ds.Tables(0).Rows(i)("kanbanno")
                        Grid.JSProperties("cpTime4") = ds.Tables(0).Rows(i)("kanbantime")
                    End If

                    '===========================================

                    'If i = 0 Then
                    '    txtkanban1.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime1.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 1 Then
                    '    txtkanban2.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime2.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 2 Then
                    '    txtkanban3.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime3.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'ElseIf i = 3 Then
                    '    txtkanban4.Text = ds.Tables(0).Rows(i)("kanbanno")
                    '    txttime4.Text = ds.Tables(0).Rows(i)("kanbantime")
                    'End If
                    If ds.Tables(0).Rows(i)("KanbanCycle") = "1" Or ds.Tables(0).Rows(i)("KanbanCycle") = "5" Or ds.Tables(0).Rows(i)("KanbanCycle") = "9" Or ds.Tables(0).Rows(i)("KanbanCycle") = "13" Or ds.Tables(0).Rows(i)("KanbanCycle") = "17" Then
                        txtkanban1.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime1.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "2" Or ds.Tables(0).Rows(i)("KanbanCycle") = "6" Or ds.Tables(0).Rows(i)("KanbanCycle") = "10" Or ds.Tables(0).Rows(i)("KanbanCycle") = "14" Or ds.Tables(0).Rows(i)("KanbanCycle") = "18" Then
                        txtkanban2.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime2.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "3" Or ds.Tables(0).Rows(i)("KanbanCycle") = "7" Or ds.Tables(0).Rows(i)("KanbanCycle") = "11" Or ds.Tables(0).Rows(i)("KanbanCycle") = "15" Or ds.Tables(0).Rows(i)("KanbanCycle") = "19" Then
                        txtkanban3.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime3.Text = ds.Tables(0).Rows(i)("kanbantime")
                    ElseIf ds.Tables(0).Rows(i)("KanbanCycle") = "4" Or ds.Tables(0).Rows(i)("KanbanCycle") = "8" Or ds.Tables(0).Rows(i)("KanbanCycle") = "12" Or ds.Tables(0).Rows(i)("KanbanCycle") = "16" Or ds.Tables(0).Rows(i)("KanbanCycle") = "20" Then
                        txtkanban4.Text = ds.Tables(0).Rows(i)("kanbanno")
                        txttime4.Text = ds.Tables(0).Rows(i)("kanbantime")
                    End If
                Next
            Else
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text

                txtaffiliateentrydate.Text = ""
                txtaffiliateentryname.Text = ""
                txtsupplierapprovalname.Text = ""
                txtaffiliateappname.Text = ""

                txtaffiliateappdate.Text = ""
                txtsupplierapprovaldate.Text = ""

                If dtkanban.Text <> "" Then
                    Grid.JSProperties("cpKanban1") = dtkanban.Text & "-1"
                    Grid.JSProperties("cpTime1") = "10:00"

                    Grid.JSProperties("cpKanban2") = dtkanban.Text & "-2"
                    Grid.JSProperties("cpTime2") = "13:00"

                    Grid.JSProperties("cpKanban3") = dtkanban.Text & "-3"
                    Grid.JSProperties("cpTime3") = "15:00"

                    Grid.JSProperties("cpKanban4") = dtkanban.Text & "-4"
                    Grid.JSProperties("cpTime4") = "17:00"
                End If
            End If

            cn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad(ByVal pkanbandate As Date, ByVal psuppID As String)

        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pFieldDelivery As String

        If Left(dtkanban.Text, 2) = "01" Then
            pFieldDelivery = "DeliveryD1"
        ElseIf Left(dtkanban.Text, 2) = "02" Then
            pFieldDelivery = "DeliveryD2"
        ElseIf Left(dtkanban.Text, 2) = "03" Then
            pFieldDelivery = "DeliveryD3"
        ElseIf Left(dtkanban.Text, 2) = "04" Then
            pFieldDelivery = "DeliveryD4"
        ElseIf Left(dtkanban.Text, 2) = "05" Then
            pFieldDelivery = "DeliveryD5"
        ElseIf Left(dtkanban.Text, 2) = "06" Then
            pFieldDelivery = "DeliveryD6"
        ElseIf Left(dtkanban.Text, 2) = "07" Then
            pFieldDelivery = "DeliveryD7"
        ElseIf Left(dtkanban.Text, 2) = "08" Then
            pFieldDelivery = "DeliveryD8"
        ElseIf Left(dtkanban.Text, 2) = "09" Then
            pFieldDelivery = "DeliveryD9"
        Else
            pFieldDelivery = "DeliveryD" & Left(dtkanban.Text, 2)
        End If

        If Left(dtkanban.Text, 2) = "" Then
            pFieldDelivery = "DeliveryD1"
        End If
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = " SELECT " & vbCrLf & _
                  "     idx= '', " & vbCrLf & _
                  " 	cols = cols,  " & vbCrLf & _
                  "  	colno = ROW_NUMBER() OVER(ORDER BY cols DESC),  " & vbCrLf & _
                  "  	colpartno = colpartno,  " & vbCrLf & _
                  "  	coldescription=coldescription,  " & vbCrLf & _
                  "  	colpono=colpono,  " & vbCrLf & _
                  "  	coluom=coluom,  " & vbCrLf & _
                  "  	colmoq=colmoq,  " & vbCrLf & _
                  "  	colqty=colqty,  " & vbCrLf & _
                  "  	colpoqty= colpoqty,  " & vbCrLf & _
                  "  	colremainingpo=colremainingpo ,  "

            ls_SQL = ls_SQL + "  	colremainingsupplier = isnull(colremainingsupplier,0), coldeliveryqty=isnull(coldeliveryqty,0),  " & vbCrLf & _
                              "  	colkanbanqty=colkanbanqty,  " & vbCrLf & _
                              "  	colcycle1 = colcycle1,  " & vbCrLf & _
                              "  	colcycle2 = colcycle2,  " & vbCrLf & _
                              "  	colcycle3 = colcycle3,  " & vbCrLf & _
                              "  	colcycle4 = colcycle4,  " & vbCrLf & _
                              "  	colbox = colbox,  " & vbCrLf & _
                              "  	cols1 = cols1, coluomcode = coluomcode  " & vbCrLf & _
                              "     ,kanbanno1,kanbanno2,kanbanno3,kanbanno4 " & vbCrLf & _
                              "     ,kanbantime1,kanbantime2,kanbantime3,kanbantime4 " & vbCrLf & _
                              "  FROM ( " & vbCrLf

            ls_SQL = ls_SQL + " SELECT DISTINCT  " & vbCrLf & _
                  "     cols = '1', " & vbCrLf & _
                  " 	colno = '0', " & vbCrLf & _
                  " 	colpartno = KD.partNo, " & vbCrLf & _
                  " 	coldescription = MP.partname , " & vbCrLf & _
                  " 	colpono = KD.pono,  " & vbCrLf & _
                  " 	coluom = ISNULL(MUC.Description,''), " & vbCrLf & _
                  " 	colmoq = ISNULL(KD.POMoq,MPM.MOQ), " & vbCrLf & _
                  " 	colqty = ISNULL(KD.POQtyBox,MPM.QtyBox), " & vbCrLf & _
                  " 	colpoqty = COALESCE(PRD.POQty,PD.POQty),  " & vbCrLf & _
                  " 	colremainingpo= COALESCE(PRD.POQty,PD.POQty) - (SELECT SUM(ISNULL(KanbanQty,0)) FROM dbo.Kanban_Detail  " & vbCrLf & _
                  " 										WHERE PONo = PD.PoNo "

            ls_SQL = ls_SQL + " 										AND PartNo = PD.partNo), " & vbCrLf & _
                              " 	colremainingsupplier=MSS.DailyDeliveryCapacity - (SELECT isnull(sum(KanbanQty),0) FROM dbo.Kanban_Detail A" & vbCrLf & _
                              "						                                    INNER JOIN dbo.Kanban_Master B ON A.KanbanNo = B.KanbanNo " & vbCrLf & _
                              "                             							WHERE CONVERT(char(8), CONVERT(DATETIME,KanbanDate),112) = '" & Format(dtkanban.Value, "yyyyMMdd") & "' " & vbCrLf & _
                              " 														AND B.SupplierID = '" & Trim(psuppID) & "'  AND A.PartNo = KD.PartNo) , " & vbCrLf

            ls_SQL = ls_SQL + " 	coldeliveryqty = ISNULL(PD." & pFieldDelivery & ",0), " & vbCrLf

            ls_SQL = ls_SQL + " 	colkanbanqty= ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 1 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID "

            ls_SQL = ls_SQL + " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf & _
                              " 				+ ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 2 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf & _
                              " 				+ ISNULL((SELECT KanbanQty "

            ls_SQL = ls_SQL + " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 3 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf & _
                              " 				+ ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID "

            ls_SQL = ls_SQL + " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 4 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0), " & vbCrLf & _
                              " 	colcycle1 = ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo "

            ls_SQL = ls_SQL + " 				WHERE KMI.KanbanCycle = 1 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0), " & vbCrLf & _
                              " 	colcycle2 = ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 2 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID "

            ls_SQL = ls_SQL + " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0), " & vbCrLf & _
                              " 	colcycle3 = ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 3 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0), "

            ls_SQL = ls_SQL + " 	colcycle4 = ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 4 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0), " & vbCrLf & _
                              " 	colbox = (ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 1 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID "

            ls_SQL = ls_SQL + " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf & _
                              " 				+ ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 2 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf & _
                              " 				+ ISNULL((SELECT KanbanQty "

            ls_SQL = ls_SQL + " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 3 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf & _
                              " 				+ ISNULL((SELECT KanbanQty " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 			 		INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID "

            ls_SQL = ls_SQL + " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 4 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0)) / MPM.QtyBox ," & vbCrLf & _
                              " 	cols1 = '1', coluomcode = MP.UnitCls " & vbCrLf

            ls_SQL = ls_SQL + " 	,kanbanno1= ISNULL((SELECT KMI.kanbanno " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 1 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbanno2= ISNULL((SELECT KMI.kanbanno " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 2 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbanno3= ISNULL((SELECT KMI.kanbanno " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 3 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbanno4= ISNULL((SELECT KMI.kanbanno " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 4 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbantime1= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)  " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 1 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbantime2= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)  " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 2 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbantime3= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)  " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 3 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf
            ls_SQL = ls_SQL + " 	,kanbantime4= ISNULL((SELECT CONVERT(CHAR(5),kanbantime)  " & vbCrLf & _
                              " 				FROM dbo.Kanban_Master  KMI " & vbCrLf & _
                              " 					INNER JOIN dbo.Kanban_Detail KDI ON KMI.AffiliateID = KDI.AffiliateID " & vbCrLf & _
                              " 						AND KMI.SupplierID = KDI.SupplierID " & vbCrLf & _
                              " 						AND KMI.KanbanNo = KDI.KanbanNo " & vbCrLf & _
                              " 				WHERE KMI.KanbanCycle = 4 " & vbCrLf & _
                              " 					AND KMI.AffiliateID = KD.AffiliateID " & vbCrLf & _
                              "                     AND KDI.PartNo = KD.partNo " & vbCrLf & _
                              "                     AND KDI.PONo = KD.PONo " & vbCrLf & _
                              " 					AND KMI.SupplierID = KD.SupplierID " & vbCrLf & _
                              " 					AND CONVERT(char(8), CONVERT(DATETIME,KMI.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "'),0) " & vbCrLf

            'ls_SQL = ls_SQL + " FROM dbo.Kanban_Master KM  " & vbCrLf & _
            '                  '" 	INNER JOIN dbo.Kanban_Detail  KD ON KM.KanbanNo = KD.KanbanNo " & vbCrLf & _
            '"                                         AND KM.AffiliateID = KD.AffiliateID " & vbCrLf & _
            '"                                         AND KM.SupplierID = KD.SupplierID " & vbCrLf & _
            '" 	INNER JOIN dbo.po_detailUpload PD ON KD.PartNo = PD.PartNo AND KD.PONo = PD.PONo " & vbCrLf & _
            '"     INNER JOIN PO_Master PM ON PM.PoNo = PD.PONo " & vbCrLf & _
            '"                                 and PM.AffiliateID = PD.AffiliateID " & vbCrLf & _
            '"                                 and PM.SupplierID = PD.SupplierID " & vbCrLf & _
            '"     LEFT JOIN dbo.PORev_Master PRM ON PM.AffiliateID = PRM.AffiliateID " & vbCrLf & _
            '"                                     AND PRM.PONo = PM.PONo " & vbCrLf & _
            '"                                     AND PRM.SupplierID = PM.SupplierID " & vbCrLf & _
            '"     LEFT JOIN dbo.PORev_Detail PRD ON PRD.PONo = PRM.PONo " & vbCrLf & _
            '"                                     AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
            '"                                     AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
            '"                                     AND PRM.PORevNo = PRD.PORevNo " & vbCrLf & _
            '"                                     AND PRD.PartNo = PD.PartNo " & vbCrLf & _
            '" 	INNER JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo " & vbCrLf & _
            '"     LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls " & vbCrLf & _
            '"     LEFT JOIN dbo.MS_SupplierCapacity MSS ON MSS.SupplierID = PD.SupplierID AND MSS.PartNo = PD.PartNo	" & vbCrLf & _
            ls_SQL = ls_SQL + " FROM dbo.Kanban_Master KM  " & vbCrLf & _
                                          " 	LEFT JOIN dbo.Kanban_Detail  KD ON KM.KanbanNo = KD.KanbanNo " & vbCrLf & _
                                          "                                         AND KM.AffiliateID = KD.AffiliateID " & vbCrLf & _
                                          "                                         AND KM.SupplierID = KD.SupplierID " & vbCrLf & _
                                          "                                         AND KM.DeliveryLocationCode = KD.DeliveryLocationCode " & vbCrLf & _
                                          " 	LEFT JOIN dbo.po_detailUpload PD ON KD.PartNo = PD.PartNo AND KD.PONo = PD.PONo " & vbCrLf & _
                                          "                                         AND KD.SupplierID = PD.SupplierID " & vbCrLf & _
                                          "                                         AND KD.AffiliateID = PD.AffiliateID " & vbCrLf & _
                                          "     LEFT JOIN PO_Master PM ON PM.PoNo = PD.PONo " & vbCrLf & _
                                          "                                 and PM.AffiliateID = PD.AffiliateID " & vbCrLf & _
                                          "                                 and PM.SupplierID = PD.SupplierID " & vbCrLf & _
                                          "     LEFT JOIN dbo.PORev_Master PRM ON PM.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                                          "                                     AND PRM.PONo = PM.PONo " & vbCrLf & _
                                          "                                     AND PRM.SupplierID = PM.SupplierID " & vbCrLf & _
                                          "     LEFT JOIN dbo.PORev_Detail PRD ON PRD.PONo = PRM.PONo " & vbCrLf & _
                                          "                                     AND PRD.AffiliateID = PRM.AffiliateID " & vbCrLf & _
                                          "                                     AND PRD.SupplierID = PRM.SupplierID " & vbCrLf & _
                                          "                                     AND PRD.PartNo = PD.PartNo " & vbCrLf & _
                                          "                                     AND PRD.SeqNo = (SELECT MAX(seqNO) FROM PORev_Detail A WHERE" & vbCrLf & _
                                          "                                                         A.PONo = PD.PONo " & vbCrLf & _
                                          "							                                AND A.AffiliateID = PD.AffiliateID  " & vbCrLf & _
                                          "							                                AND A.SupplierID = PD.SupplierID  " & vbCrLf & _
                                          "							                                AND A.PartNo = PD.PartNo) " & vbCrLf & _
                                          " 	LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo " & vbCrLf & _
                                          "     LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                                          "     LEFT JOIN dbo.MS_UnitCls MUC ON MUC.UnitCls = MP.UnitCls " & vbCrLf & _
                                          "     LEFT JOIN dbo.MS_SupplierCapacity MSS ON MSS.SupplierID = PD.SupplierID AND MSS.PartNo = PD.PartNo	" & vbCrLf & _
                                          " WHERE CONVERT(char(8), CONVERT(DATETIME,KM.KanbanDate),112) = '" & Format(pkanbandate, "yyyyMMdd") & "' " & vbCrLf & _
                                          "  AND KD.AffiliateID = '" & Trim(txtaffiliatecode.Text) & "' " & vbCrLf

            If psuppID <> clsGlobal.gs_All And psuppID <> "" Then
                ls_SQL = ls_SQL + " AND KM.SupplierID = '" & Trim(psuppID) & "'"
            End If

            If cbolocation.Text <> clsGlobal.gs_All And cbolocation.Text <> "" Then
                ls_SQL = ls_SQL + " AND KM.DeliveryLocationCode = '" & Trim(cbolocation.Text) & "'"
            End If

            If (Session("FilterKanbanNo") <> "" And Session("FilterKanbanNo") <> clsGlobal.gs_All) Then
                ls_SQL = ls_SQL + " AND KM.Kanbanno = '" & Trim(Session("FilterKanbanNo")) & "'" & vbCrLf
            End If

            ls_SQL = ls_SQL + ")xx"


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()

            If Grid.VisibleRowCount = 0 Then
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                Call colorGrid()
            Else
                Grid.JSProperties("cpMessage") = ""
            End If

        End Using
    End Sub


#End Region

End Class