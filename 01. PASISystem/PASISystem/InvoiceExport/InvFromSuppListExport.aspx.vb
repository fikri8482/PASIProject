Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView

Public Class InvFromSuppListExport
    Inherits System.Web.UI.Page

    '-----------------------------------------------------
    Private grid_Renamed As ASPxGridView
    Private mergedCells As New Dictionary(Of GridViewDataColumn, TableCell)()
    Private cellRowSpans As New Dictionary(Of TableCell, Integer)()
    '-----------------------------------------------------

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdt2") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdeliveryqty") = "ALL"
                grid.JSProperties("cpinvoice") = "ALL"
                grid.JSProperties("cpsupplier") = "ALL"
                grid.JSProperties("cpaffiliate") = "ALL"

                dt1.Text = Format(Now, "01 MMM yyyy")
                dt2.Text = Format(Now, "dd MMM yyyy")
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

#Region "PROCEDURE"
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'SAffiliate
        ls_sql = "SELECT distinct Affiliate_Code = '" & clsGlobal.gs_All & "', Affiliate_Name = '" & clsGlobal.gs_All & "' from MS_AFfiliate " & vbCrLf & _
                 "UNION ALL Select Affiliate_Code = RTRIM(AffiliateID) ,Affiliate_Name = RTRIM(Affiliatename) FROM MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate_Code")
                .Columns(0).Width = 70
                .Columns.Add("Affiliate_Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "Affiliate Code"
                .DataBind()
            End With
            sqlConn.Close()
        End Using

        'SSupplier
        ls_sql = "SELECT distinct Supplier_Code = '" & clsGlobal.gs_All & "', Supplier_Name = '" & clsGlobal.gs_All & "' from MS_Supplier " & vbCrLf & _
                 "UNION ALL Select Supplier_Code = RTRIM(SupplierID) ,Supplier_Name = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier_Code")
                .Columns(0).Width = 70
                .Columns.Add("Supplier_Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "Supplier Code"
                .DataBind()
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            
            If checkbox1.Checked = True Then
                ls_filter = ls_filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(SM.DeliveryDate,'')),106) between '" & Format(dt1.Value, "dd MMM yyyy") & "' AND '" & Format(dt2.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbinvoice.Value = "YES" Then
                ls_filter = ls_filter + "AND isnull(IM.InvoiceNo, '') <> '' " & vbCrLf
            ElseIf rbinvoice.Value = "NO" Then
                ls_filter = ls_filter + " AND isnull(IM.InvoiceNo,'') = '' " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_filter = ls_filter + " AND RM.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If txtsupinvno.Text <> "" Then
                ls_filter = ls_filter + " AND IM.InvoiceNo Like '%" & Trim(txtsupinvno.Text) & "%'" & vbCrLf
            End If

            If cbosupplier.Text <> "" And cbosupplier.Text <> clsGlobal.gs_All Then
                ls_filter = ls_filter + " AND RD.SupplierID = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_filter = ls_filter + " AND RD.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_filter = ls_filter + " AND RD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
            End If


            ls_SQL = " SELECT  DISTINCT coldetail = 'InvoiceEntryExport.aspx?prm=' + RTRIM(ISNULL(IM.InvoiceNo,'')) + '|' + RTRIM(isnull(RM.SuratJalanNo,'')) + '|' + Rtrim(isnull(RM.AffiliateID,'')) + '|' + Rtrim(Isnull(RM.SupplierID,'')) , " & vbCrLf & _
                     " coldetailname = CASE WHEN ISNULL(IM.InvoiceNo,'') = '' THEN 'INVOICE' ELSE 'DETAIL' END, " & vbCrLf & _
                     "         no = ROW_NUMBER() OVER ( ORDER BY RM.OrderNo ) , " & vbCrLf & _
                     "         period = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) , " & vbCrLf & _
                     "         affiliatecode = RM.AffiliateID , " & vbCrLf & _
                     "         affiliatename = MA.AffiliateName , " & vbCrLf & _
                     "         orderno = RM.OrderNo , " & vbCrLf & _
                     "         suppliercode = RM.SupplierID , " & vbCrLf & _
                     "         suppliername = MS.SupplierName , " & vbCrLf & _
                     "         suppplandeldate = CONVERT(CHAR(12), CONVERT(DATETIME, COALESCE(PRM.ETDVendor, " & vbCrLf & _
                     "                                                               POM.ETDVendor)), 106) , " & vbCrLf & _
                     "         suppdeldate = CONVERT(CHAR(12), CONVERT(DATETIME, SM.DeliveryDate), 106) , "

            ls_SQL = ls_SQL + "         suppsj = RM.Suratjalanno , " & vbCrLf & _
                              "         forwarderreceivedate = CONVERT(CHAR(12), CONVERT(DATETIME, RM.ReceiveDate), 106) , " & vbCrLf & _
                              "         supinvno = ISNULL(IM.InvoiceNo,'') , " & vbCrLf & _
                              "         suppinvdate = ISNULL(CONVERT(CHAR(12), CONVERT(DATETIME, IM.InvoiceDate), 106),'') , " & vbCrLf & _
                              "         curr = ISNULL(MC.DESCRIPTION,''), " & vbCrLf & _
                              "         price = SUM(ISNULL(ID.Price,0)) , " & vbCrLf & _
                              "         amount = SUM(ISNULL(ID.Amount,0)) " & vbCrLf & _
                              " FROM    dbo.ReceiveForwarder_Master RM " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON RM.Suratjalanno = RD.Suratjalanno " & vbCrLf & _
                              "                                                 AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                 AND RM.SupplierID = RD.SupplierID "

            ls_SQL = ls_SQL + "                                                 AND RM.POno = RD.POno " & vbCrLf & _
                              "                                                 AND RM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "         LEFT JOIN DOSupplier_Detail_Export SD ON SD.suratjalanno = RM.suratjalanno " & vbCrLf & _
                              "                                                  AND SD.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              "                                                  AND SD.SupplierID = RM.SupplierID " & vbCrLf & _
                              "                                                  AND SD.POno = RM.POno " & vbCrLf & _
                              "                                                  AND SD.OrderNo = RM.OrderNo " & vbCrLf & _
                              "                                                  AND SD.Partno = RD.PartNo " & vbCrLf & _
                              "         LEFT JOIN DOSupplier_Master_Export SM ON SM.suratjalanno = SD.suratjalanno " & vbCrLf & _
                              "                                                  AND SM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                              "                                                  AND SM.SupplierID = SD.SupplierID "

            ls_SQL = ls_SQL + "                                                  AND SM.POno = SD.POno " & vbCrLf & _
                              "                                                  AND SM.OrderNo = SD.OrderNo " & vbCrLf & _
                              "         LEFT JOIN ( SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo1 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             ETAPort = ETAPort1 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory1 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo2 , "

            ls_SQL = ls_SQL + "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                             ETAPort = ETAPort2 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory2 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo3 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , " & vbCrLf & _
                              "                             ETAPort = ETAPort3 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory3 " & vbCrLf & _
                              "                     FROM    Po_Master_Export "

            ls_SQL = ls_SQL + "                     UNION ALL " & vbCrLf & _
                              "                     SELECT  * , " & vbCrLf & _
                              "                             OrderNO = OrderNo4 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                             ETAPort = ETAPort4 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory4 " & vbCrLf & _
                              "                     FROM    Po_Master_Export " & vbCrLf & _
                              "                   ) POM ON POM.PONO = SD.PONO " & vbCrLf & _
                              "                            AND POM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                              "                            AND POM.SupplierID = SD.SupplierID " & vbCrLf & _
                              "                            AND POM.OrderNo = SD.OrderNo "

            ls_SQL = ls_SQL + "         LEFT JOIN ( SELECT TOP 1 " & vbCrLf & _
                              "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo1 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor1 , " & vbCrLf & _
                              "                             ETAPort = ETAPort1 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory1 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT TOP 1 " & vbCrLf & _
                              "                             * , "

            ls_SQL = ls_SQL + "                             OrderNO = OrderNo2 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor2 , " & vbCrLf & _
                              "                             ETAPort = ETAPort2 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory2 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT TOP 1 " & vbCrLf & _
                              "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo3 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor3 , "

            ls_SQL = ls_SQL + "                             ETAPort = ETAPort3 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory3 " & vbCrLf & _
                              "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                     UNION ALL " & vbCrLf & _
                              "                     SELECT TOP 1 " & vbCrLf & _
                              "                             * , " & vbCrLf & _
                              "                             OrderNO = OrderNo4 , " & vbCrLf & _
                              "                             ETDVendor = ETDVendor4 , " & vbCrLf & _
                              "                             ETAPort = ETAPort4 , " & vbCrLf & _
                              "                             ETAFactory = ETAFactory4 "

            ls_SQL = ls_SQL + "                     FROM    PoRev_Master_Export " & vbCrLf & _
                              "                     ORDER BY PORevNo " & vbCrLf & _
                              "                   ) PRM ON PRM.PONO = SD.PONO " & vbCrLf & _
                              "                            AND PRM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                              "                            AND PRM.SupplierID = SD.SupplierID " & vbCrLf & _
                              "                            AND PRM.OrderNo = SD.OrderNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                                       AND IM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                       AND IM.SupplierID = RD.SupplierID " & vbCrLf & _
                              "                                                       AND IM.POno = RD.POno " & vbCrLf & _
                              "                                                       AND IM.OrderNo = RD.OrderNo "

            ls_SQL = ls_SQL + "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                              "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                              "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                              "                                                       AND ID.PartNo = RD.PartNo " & vbCrLf & _
                              " 		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = SD.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN ms_supplier MS ON MS.SupplierID = SD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                              " WHERE RM.Suratjalanno <> '' " & vbCrLf

            ls_SQL = ls_SQL + ls_filter

            ls_SQL = ls_SQL + "         GROUP BY RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8) , " & vbCrLf & _
                              "         RM.AffiliateID , " & vbCrLf & _
                              "         MA.AffiliateName , " & vbCrLf & _
                              "         RM.OrderNo , " & vbCrLf & _
                              "         RM.SupplierID , " & vbCrLf & _
                              "         MS.SupplierName , " & vbCrLf & _
                              "         CONVERT(CHAR(12), CONVERT(DATETIME, COALESCE(PRM.ETDVendor, POM.ETDVendor)), 106) , " & vbCrLf & _
                              "         CONVERT(CHAR(12), CONVERT(DATETIME, SM.DeliveryDate), 106) , " & vbCrLf & _
                              "         RM.Suratjalanno , " & vbCrLf & _
                              "         CONVERT(CHAR(12), CONVERT(DATETIME, RM.ReceiveDate), 106) , " & vbCrLf & _
                              "         ISNULL(IM.InvoiceNo,'') , " & vbCrLf & _
                              "         CONVERT(CHAR(12), CONVERT(DATETIME, IM.InvoiceDate), 106) , " & vbCrLf & _
                              "         ISNULL(MC.DESCRIPTION,'') "


            

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

#End Region

#Region "FORM EVENT"
    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            'Dim pPlan As Date = Split(e.Parameters, "|")(1)
            'Dim pSupplierDeliver As String = Split(e.Parameters, "|")(2)
            'Dim pRemaining As String = Split(e.Parameters, "|")(3)
            'Dim psj As String = Split(e.Parameters, "|")(4)
            'Dim pDateFrom As Date = Split(e.Parameters, "|")(5)
            'Dim pDateTo As Date = Split(e.Parameters, "|")(6)
            'Dim pSupplier As String = Split(e.Parameters, "|")(7)
            'Dim pPart As String = Split(e.Parameters, "|")(8)
            'Dim pPoNo As String = Split(e.Parameters, "|")(9)
            'Dim pKanban As String = Split(e.Parameters, "|")(10)

            Select Case pAction
                Case "gridload"
                    'Call up_GridLoad(pPlan, pSupplierDeliver, pRemaining, psj, pDateFrom, pDateTo, pSupplier, pPart, pPoNo, pKanban)
                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "kosong"

            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

#End Region


    'Private Sub grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
    '    If e.RowType <> GridViewRowType.Data Then Return
    '    If e.GetValue("partno").ToString = "" Then e.Row.BackColor = Drawing.Color.LightGray
    'End Sub
End Class