Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView


Public Class AffKanbanList
    Inherits System.Web.UI.Page
#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_KanbanDate As String
    Dim ls_approve As Boolean
#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If
            'If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
            Session("M01Url") = Request.QueryString("Session")
            Session("MenuDesc") = "KANBAN LIST"
            'End If
            lblerrmessage.Text = ""
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                'If Session("M01Url") <> "" Then
                '    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                '        Session("MenuDesc") = "KANBAN LIST"
                '    End If
                'End If

                'If Not IsNothing(Request.QueryString("prm")) Then
                'Else
                '    Call up_fillcombo()
                '    lblerrmessage.Text = ""
                '    grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                '    grid.JSProperties("cpdt2") = Format(Now, "dd MMM yyyy")
                '    If Session("M01Url") <> "" Then
                '        Session.Remove("M01Url")
                '    End If
                'End If
                If Not IsNothing(Request.QueryString("prm")) Then
                    If Session("M01Url") <> "" Then
                        If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                            Session("MenuDesc") = "KANBAN LIST"
                        End If
                    End If

                    Dim param As String = Request.QueryString("prm").ToString

                    Dim pdt1 As Date = Split(param, "|")(0)
                    Dim pdt2 As Date = Split(param, "|")(1)
                    Dim pcbosupplier As String = Split(param, "|")(2)
                    Dim pSuppliername As String = Split(param, "|")(3)
                    Dim pcbolocation As String = Split(param, "|")(4)
                    Dim pLocationname As String = Split(param, "|")(5)
                    Dim pAffiliateCode As String = Split(param, "|")(6)
                    Dim pAffiliateName As String = Split(param, "|")(7)
                    Dim pKanbanno As String = Split(param, "|")(8)

                    Call up_fillcombo()
                    Call up_FillComboKanban(pdt1, pdt2, pcbosupplier, pcbolocation, pAffiliateCode)
                    cbokanbanno.Text = pKanbanno

                    cbosupplier.Text = pcbosupplier
                    cbolocation.Text = pcbolocation
                    txtsupplier.Text = pSuppliername
                    txtlocation.Text = pLocationname
                    cboaffiliate.Text = pAffiliateCode
                    txtaffiliate.Text = pAffiliateName
                    cbokanbanno.Text = "== ALL =="
                    dt1.Value = pdt1
                    dt2.Value = pdt2
                    grid.JSProperties("cpdt1") = pdt1
                    grid.JSProperties("cpdt2") = pdt2
                    Call up_GridLoad()
                Else
                    Clear()
                    dt1.Value = Format(Now, "dd MMM yyyy")
                    dt2.Value = Format(Now, "dd MMM yyyy")

                    If Session("KCR-Load") <> "YES" And Not IsNothing(Session("KCR-Load")) Then
                        Call up_fillcombo()
                        Call up_FillComboKanban(Session("KCR-dt1"), Session("KCR-dt2"), Session("KCR-cbosup"), Session("KCR-cboloc"), Session("KCR-AffiliateCode"))

                        dt1.Text = Session("KCR-dt1")
                        dt2.Text = Session("KCR-dt2")
                        grid.JSProperties("cpdt1") = Session("KCR-dt1")
                        grid.JSProperties("cpdt2") = Session("KCR-dt2")
                        cbosupplier.Text = Session("KCR-cbosup")
                        txtsupplier.Text = Session("KCR-txtsup")
                        cbolocation.Text = Session("KCR-cboloc")
                        txtlocation.Text = Session("KCR-txtloc")
                        cboaffiliate.Text = Session("KCR-AffiliateCode")
                        txtaffiliate.Text = Session("KCR-AffiliateName")
                        Call up_GridLoad()


                    Else
                        Call up_fillcombo()
                        lblerrmessage.Text = ""
                        grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                        grid.JSProperties("cpdt2") = Format(Now, "dd MMM yyyy")
                        If Session("M01Url") <> "" Then
                            Session.Remove("M01Url")
                        End If
                    End If
                End If
            End If

            Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        Finally
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Public Sub btnclear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnclear.Click
        Clear()
    End Sub

    Protected Sub btncreatekanban_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btncreatekanban.Click

        Response.Redirect("~/Kanban/KanbanCreate.aspx?prm= +'back'")
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", i As Long = 1
        Dim ls_Supplier As String
        Dim ls_deliveryLocation As String = ""
        Dim ls_Affiliate As String = ""
        Dim ls_KanbanNo As String = ""

        ls_Supplier = ""
        ls_KanbanDate = ""
        ls_approve = True

        Session.Remove("KCR-ReportCode")
        Session.Remove("KCR-KanbanDate")
        Session.Remove("KCR-SupplierCode")
        Session.Remove("KCR-Form")
        Session.Remove("KCR-Approve")
        Session.Remove("KCR-dt1")
        Session.Remove("KCR-Dt2")
        Session.Remove("KCR-Cbosup")
        Session.Remove("KCR-txtsup")
        Session.Remove("KCR-cboloc")
        Session.Remove("KCR-txtloc")
        Session.Remove("KCR-Load")
        Session.Remove("KCR-DeliveryLocation")
        Session.Remove("KCR-Kanbanno")
        'Session.Remove("KCR-AffiliateCode")

        Session("KCR-Form") = "KanbanList"

        With grid
            For i = 0 To e.UpdateValues.Count - 1
                If (e.UpdateValues(i).NewValues("cols").ToString()) = 1 Then
                    If ls_KanbanDate = "" Then
                        ls_KanbanDate = "'" + Trim(e.UpdateValues(i).NewValues("colkanbandate").ToString()) + "'"
                        ls_Supplier = "'" + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + "'"
                        ls_deliveryLocation = "'" + Trim(e.UpdateValues(i).NewValues("coldeliverycode").ToString()) + "'"
                        ls_KanbanNo = "'" + Trim(e.UpdateValues(i).NewValues("colkanbanno").ToString()) + "'"
                    Else
                        ls_KanbanDate = ls_KanbanDate + ",'" + Trim(e.UpdateValues(i).NewValues("colkanbandate").ToString()) + "'"
                        ls_Supplier = ls_Supplier + ",'" + Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) + "'"
                        ls_deliveryLocation = ls_deliveryLocation + ",'" + Trim(e.UpdateValues(i).NewValues("coldeliverycode").ToString()) + "'"
                        ls_KanbanNo = ls_KanbanNo + ",'" + Trim(e.UpdateValues(i).NewValues("colkanbanno").ToString()) + "'"
                    End If

                    If ls_approve = True Then
                        If Trim(e.UpdateValues(i).NewValues("colkanbanstatus2").ToString()) = 1 Then ls_approve = True Else ls_approve = False
                    End If
                End If
            Next
            Session("KCR-Load") = "YES"
            Session("KCR-KanbanDate") = ls_KanbanDate
            Session("KCR-SupplierCode") = ls_Supplier
            Session("KCR-Approve") = ls_approve
            Session("KCR-dt1") = dt1.Value
            Session("KCR-Dt2") = dt2.Value
            Session("KCR-Cbosup") = cbosupplier.Text
            Session("KCR-txtsup") = txtsupplier.Text
            Session("KCR-cboloc") = cbolocation.Text
            Session("KCR-txtloc") = txtlocation.Text
            Session("KCR-AffiliateCode") = Trim(cboaffiliate.Text)
            Session("KCR-AffiliateName") = Trim(txtaffiliate.Text)
            Session("KCR-DeliveryLocation") = ls_deliveryLocation
            Session("KCR-Kanbanno") = ls_KanbanNo
        End With

        
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)

            Select Case pAction
                Case "gridload"
                    Dim pDateFrom As Date = Split(e.Parameters, "|")(1)
                    Dim pDateTo As Date = Split(e.Parameters, "|")(2)
                    Dim pSupplier As String = Split(e.Parameters, "|")(3)

                    If cboaffiliate.Text = "" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "7003", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                        Exit Sub
                    End If

                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "kosong"

                Case "PrintCard"
                    If Session("KCR-KanbanDate") <> "" Then
                        If Session("KCR-Approve") = True Then
                            Session("KCR-ReportCode") = "KanbanCard"
                            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/AffKanban/ViewReport.aspx")
                        Else
                            Call clsMsg.DisplayMessage(lblerrmessage, "7011", clsMessage.MsgType.ErrorMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                        End If
                    Else
                        'Select Data
                        Call clsMsg.DisplayMessage(lblerrmessage, "7012", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "PrintCycle"
                    If Session("KCR-KanbanDate") <> "" Then
                        If Session("KCR-Approve") = True Then
                            Session("KCR-ReportCode") = "KanbanCycle"
                            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/AffKanban/ViewReport.aspx")
                        Else
                            Call clsMsg.DisplayMessage(lblerrmessage, "7011", clsMessage.MsgType.ErrorMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                        End If
                    Else
                        'Select Data
                        Call clsMsg.DisplayMessage(lblerrmessage, "7012", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub colorGrid()
        grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.White
    End Sub

    Private Sub up_FillComboKanban(ByVal pdate1 As Date, ByVal pdate2 As Date, ByVal psupplier As String, ByVal plocation As String, ByVal pAffiliate As String)
        Dim ls_sql As String
        ls_sql = " SELECT distinct [kanbanNo] = '" & clsGlobal.gs_All & "' from Kanban_Master union all SELECT distinct [KanbanNo] = RTRIM(KanbanNo) " & vbCrLf & _
                 " FROM Kanban_Master where AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                 " and convert(char(11), convert(datetime,KanbanDate),112) between '" & Format(pdate1, "yyyyMMdd") & "' and '" & Format(pdate2, "yyyyMMdd") & "'" & vbCrLf
        If (psupplier <> clsGlobal.gs_All And psupplier <> "") Then
            ls_sql = ls_sql + " AND SupplierID  = '" & Trim(cbosupplier.Text) & "' " & vbCrLf
        End If

        If (plocation <> clsGlobal.gs_All And plocation <> "") Then
            ls_sql = ls_sql + " AND DeliveryLocationCode = '" & Trim(cbolocation.Text) & "'" & vbCrLf
        End If

        ls_sql = ls_sql + " order by kanbanno"
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbokanbanno
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Kanbanno")
                .Columns(0).Width = 70
                '.Columns.Add("Kanbanno")
                '.Columns(1).Width = 0
                .TextField = "Kanbanno"
                .DataBind()
                .Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'SupplierCode
        ls_sql = "SELECT Distinct [Supplier Code] = '" & clsGlobal.gs_All & "' ,[Supplier Name] = '" & clsGlobal.gs_All & "' from ms_supplier union all SELECT [Supplier Code] = RTRIM(supplierID) ,[Supplier Name] = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier Code")
                .Columns(0).Width = 90
                .Columns.Add("Supplier Name")
                .Columns(1).Width = 240

                .TextField = "Supplier Code"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using

        'Affiliate Code
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(Affiliatename) FROM MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate Code")
                .Columns(0).Width = 90
                .Columns.Add("Affiliate Name")
                .Columns(1).Width = 240

                .TextField = "Affiliate Code"
                .DataBind()
                '.SelectedIndex = 0
                'txtaffiliate.Text = clsGlobal.gs_All
                'cbolocation.Text = clsGlobal.gs_All
                'txtlocation.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        'Delivery Location
        ls_sql = "SELECT [Delivery Location Code] = RTRIM(DeliveryLocationCode) ,[Delivery Location Name] = RTRIM(DeliveryLocationName) FROM MS_DeliveryPlace" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbolocation
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Delivery Location Code")
                .Columns(0).Width = 150
                .Columns.Add("Delivery Location Name")
                .Columns(1).Width = 240

                .TextField = "Delivery Location Code"
                '.DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
        
    End Sub

    Private Sub Clear()
        'grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
        'grid.JSProperties("cpdt2") = Format(Now, "dd MMM yyyy")
        cbosupplier.Text = clsGlobal.gs_All
        txtsupplier.Text = clsGlobal.gs_All
        cbolocation.Text = clsGlobal.gs_All
        txtlocation.Text = clsGlobal.gs_All
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If dt1.Text = "" Then dt1.Value = Session("KCR-dt1")
        If dt2.Text = "" Then dt2.Value = Session("KCR-dt2")

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "  SELECT " & vbCrLf & _
                     "  cols = cols,   " & vbCrLf & _
                     "  colno = ROW_NUMBER() OVER(ORDER BY cols DESC),   " & vbCrLf & _
                     "  colsuppliercode = colsuppliercode,   " & vbCrLf & _
                     "  colsuppliername = colsuppliername,   " & vbCrLf & _
                     "  coldeliverycode = coldeliverycode,   " & vbCrLf & _
                     "  coldeliveryname = coldeliveryname,   " & vbCrLf & _
                     "  colkanbandate = colkanbandate,   " & vbCrLf & _
                     "  colcreateddate = colcreateddate,   " & vbCrLf & _
                     "  colcreatedby = colcreatedby,   " & vbCrLf & _
                     "  colkanbanstatus1 = colkanbanstatus1,   " & vbCrLf & _
                     "  colkanbanstatus2 = colkanbanstatus2,   " & vbCrLf & _
                     "  colkanbanstatus3 = colkanbanstatus3, " & vbCrLf & _
                     "  coldetailname = coldetailname,  " & vbCrLf & _
                     "  coldetailurl = coldetailurl,  " & vbCrLf & _
                     "  colkanbanno = colkanbanno " & vbCrLf & _
                     "  FROM ( " & vbCrLf & _
                     "  SELECT distinct " & vbCrLf & _
                     "  colno = '',  " & vbCrLf & _
                     "  cols = '0',  " & vbCrLf & _
                     "  colsuppliercode = ISNULL(KM.SupplierID,''),  " & vbCrLf & _
                     "  colsuppliername = ISNULL(MSS1.SupplierName,''),  " & vbCrLf & _
                     "  coldeliverycode = ISNULL(KM.DeliveryLocationCode,''), " & vbCrLf & _
                     "  coldeliveryname = ISNULL(MDP.DeliveryLocationName,''), " & vbCrLf & _
                     "  colkanbandate = CONVERT(CHAR(12),CONVERT(DATETIME,KanbanDate),106),  " & vbCrLf & _
                     "  colcreateddate = (select TOP 1 CONVERT(CHAR(20),CONVERT(DATETIME,EntryDate),113)  " & vbCrLf & _
                     "                      FROM dbo.Kanban_Master WHERE SupplierID = KM.SupplierID " & vbCrLf & _
                     "                      AND KanbanDate = KM.KanbanDate and KanbanNo = km.KanbanNo), " & vbCrLf & _
                     "  colcreatedby = ISNULL(KM.EntryUser, ''),  " & vbCrLf & _
                     "  colkanbanstatus1 = CASE WHEN ISNULL(KM.EntryUser, '') = '' then '0' else '1' END, " & vbCrLf & _
                     "  colkanbanstatus2 = CASE WHEN ISNULL(AffiliateApproveUser,'') = '' THEN '0' ELSE '1' END,  " & vbCrLf & _
                     "  colkanbanstatus3 = CASE WHEN ISNULL(SupplierApproveUser,'') = '' THEN '0' ELSE '1' END, " & vbCrLf & _
                     "  coldetailname = 'DETAIL', " & vbCrLf & _
                     "  coldetailurl = 'AffKanbanCreate.aspx?prm='+CONVERT(CHAR(12),CONVERT(DATETIME,KanbanDate),106) " & vbCrLf
            ls_SQL = ls_SQL + " 									 + '|' +RTRIM(ISNULL(KM.SupplierID,'')) " & vbCrLf & _
                              " 									 + '|' +RTRIM(ISNULL(MSS1.Suppliername,'')) " & vbCrLf & _
                              "                                      + '|' +CASE WHEN ISNULL((select TOP 1 CONVERT(CHAR(20),CONVERT(DATETIME,EntryDate),113) " & vbCrLf & _
                              "									                                FROM dbo.Kanban_Master WHERE SupplierID = KM.SupplierID " & vbCrLf & _
                              "									                                AND KanbanDate = KM.KanbanDate), '') = '' " & vbCrLf & _
                              "				                                then '' else " & vbCrLf & _
                              "			                                    CONVERT(CHAR(19), CONVERT(DATETIME, ISNULL((select TOP 1 CONVERT(CHAR(20),CONVERT(DATETIME,EntryDate),113) " & vbCrLf & _
                              "											                                                FROM dbo.Kanban_Master WHERE SupplierID = KM.SupplierID" & vbCrLf & _
                              "											                                                AND KanbanDate = KM.KanbanDate),'')),120) END" & vbCrLf & vbCrLf & _
                              " 									 + '|' +RTRIM(ISNULL(MSAE.FullName,'')) " & vbCrLf & _
                              " 									 + '|' +CASE WHEN ISNULL(KM. AffiliateApproveDate,'') = '' then '' else CONVERT(CHAR(19), CONVERT(DATETIME, ISNULL(KM.AffiliateApproveDate,'')),120) end " & vbCrLf & _
                              " 									 + '|' +RTRIM(ISNULL(MSAA.AffiliateName,'')) " & vbCrLf & _
                              " 									 + '|' +CASE WHEN ISNULL(KM.SupplierApproveDate,'') = '' THEN '' else CONVERT(CHAR(19), CONVERT(DATETIME, ISNULL(KM.SupplierApproveDate,'')),120) END " & vbCrLf & _
                              " 									 + '|' +RTRIM(ISNULL(MSS2.SupplierName,'')) " & vbCrLf & _
                              "                                      + '|' +'" & dt1.Value & "' " & vbCrLf & _
                              "                                      + '|' +'" & dt2.Value & "' " & vbCrLf & _
                              "                                      + '|' +'" & cbosupplier.Text & "' " & vbCrLf & _
                              "                                      + '|' +'" & Trim(txtsupplier.Text) & "' " & vbCrLf & _
                              "                                      + '|' +RTRIM(ISNULL(KM.DeliveryLocationCode,'')) " & vbCrLf & _
                              "                                      + '|' +RTRIM(ISNULL(MDP.DeliveryLocationName,'')) " & vbCrLf & _
                              "                                      + '|' +'" & Trim(cbolocation.Text) & "' " & vbCrLf & _
                              "                                      + '|' +'" & Trim(txtlocation.Text) & "' " & vbCrLf & _
                              "                                      + '|' +RTRIM(ISNULL(KM.AffiliateID,'')) " & vbCrLf & _
                              "                                      + '|' +RTRIM(ISNULL(MSAC.AffiliateName,'')) " & vbCrLf & _
                              "                                      + '|' +'" & Trim(cboaffiliate.Text) & "' " & vbCrLf & _
                              "                                      + '|' +'" & Trim(txtaffiliate.Text) & "' " & vbCrLf & _
                              "                                      + '|' +RTRIM(ISNULL(KM.KanbanNo,'')), " & vbCrLf & _
                              "  colkanbanno = RTRIM(ISNULL(KM.KanbanNo,'')) " & vbCrLf

            ls_SQL = ls_SQL + "  FROM dbo.Kanban_Master KM  " & vbCrLf & _
                              "  LEFT JOIN dbo.MS_Supplier MSS1  ON MSS1.SupplierID = KM.SupplierID  " & vbCrLf & _
                              "  LEFT JOIN dbo.MS_Supplier MSS2 ON MSS2.SupplierID = KM.SupplierApproveUser " & vbCrLf & _
                              "  LEFT JOIN MS_DeliveryPlace MDP ON KM.DeliveryLocationCode = MDP.DeliveryLocationCode AND KM.AffiliateID = MDP.AffiliateID"

            ls_SQL = ls_SQL + "  LEFT JOIN dbo.SC_UserSetup MSAE ON MSAE.UserID = KM.EntryUser AND UserCls = 0 " & vbCrLf & _
                              "  LEFT JOIN dbo.MS_Affiliate MSAA ON MSAA.AffiliateID = KM.AffiliateApproveUser " & vbCrLf & _
                              "  LEFT JOIN dbo.MS_Affiliate MSAC ON MSAC.AffiliateID = KM.AffiliateID " & vbCrLf & _
                              "  WHERE KanbanDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' and '" & Format(dt2.Value, "yyyy-MM-dd") & "'" & vbCrLf

            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_SQL = ls_SQL + " and KM.SupplierID = '" & Trim(cbosupplier.Text) & "'"
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_SQL = ls_SQL + " and KM.AffiliateID = '" & Trim(cboaffiliate.Text) & "'"
            End If

            If cbolocation.Text <> clsGlobal.gs_All And cbolocation.Text <> "" Then
                ls_SQL = ls_SQL + " and KM.DeliveryLocationCode = '" & Trim(cbolocation.Text) & "'"
            End If

            If cbokanbanno.Text <> clsGlobal.gs_All And cbokanbanno.Text <> "" Then
                ls_SQL = ls_SQL + " AND KM.kanbanno = '" & Trim(cbokanbanno.Text) & "'" & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()

            'If grid.VisibleRowCount = 0 Then
            '    Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            '    'Session("YA010Msg") = lblInfo.Text
            '    Call colorGrid()
            'Else
            '    lblerrmessage.Text = ""
            'End If

        End Using
    End Sub

#End Region
    Private Sub cbokanbanno_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cbokanbanno.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_value1 As String = Split(e.Parameter, "|")(0)
        Dim ls_value2 As String = Split(e.Parameter, "|")(1)
        Dim ls_supplier As String = Split(e.Parameter, "|")(2)
        Dim ls_location As String = Split(e.Parameter, "|")(3)
        Dim ls_affiliate As String = IIf(Split(e.Parameter, "|")(4) = "null", "", Split(e.Parameter, "|")(4))
        Dim ls_sql As String = ""

        Call up_FillComboKanban(ls_value1, ls_value2, ls_supplier, ls_location, ls_affiliate)
    End Sub

    Private Sub cbolocation_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cbolocation.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim ls_affiliate As String = Split(e.Parameter, "|")(0)
        
        Dim ls_sql As String = ""

        'Delivery Location
        ls_sql = "SELECT DISTINCT [Delivery Location Code] = '" & clsGlobal.gs_All & "' ,[Delivery Location Name] = '" & clsGlobal.gs_All & "' From MS_DEliveryPlace UNION ALL SELECT [Delivery Location Code] = RTRIM(DeliveryLocationCode) ,[Delivery Location Name] = RTRIM(DeliveryLocationName) FROM MS_DeliveryPlace where AffiliateID = '" & ls_affiliate & "' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbolocation
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Delivery Location Code")
                .Columns(0).Width = 70
                .Columns.Add("Delivery Location Name")
                .Columns(1).Width = 240
                .TextField = "Delivery Location Code"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using

    End Sub
End Class