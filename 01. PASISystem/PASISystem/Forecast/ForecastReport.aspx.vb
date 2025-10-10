Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class ForecastReport
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim pMsgID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            dtPeriodFrom.Value = Now
            up_FillCombo()
            lblInfo.Text = ""
            up_IsiHeader()
        End If

        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("NoUrutBulan") <> "1" Then
            If e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Then
                e.Cell.Text = ""
            End If           
        End If

        If e.GetValue("DescUrut") <> "1" Then
            If e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Then
                e.Cell.Text = ""
            End If            
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()
                    Call up_IsiHeader()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("AA220Msg") = lblInfo.Text
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    Call up_IsiHeader()
                Case "loadSave"
                    grid.PageIndex = 0
                    bindData()
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = clsMaster.GetTableSummaryForecastPO(dtPeriodFrom.Value, cboPartNo.Text)
                    FileName = "TemplateSummaryForecastPO.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:10", psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Session("AA220Msg") = ""
        End Try
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call bindData()
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboPartNo.Text <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.PartNo = '" & cboPartNo.Text & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " declare @period date  " & vbCrLf & _
                  " set @period = '" & Format(dtPeriodFrom.Value, "yyyy-MM") & "-01" & "'  " & vbCrLf & _
                  "  " & vbCrLf & _
                  "  select " & vbCrLf & _
                  " 	a.NoUrut, a.NoUrutBulan, " & vbCrLf & _
                  " 	a.PartNo, " & vbCrLf & _
                  " 	a.PartName, " & vbCrLf & _
                  " 	a.thn, " & vbCrLf & _
                  " 	a.bln, " & vbCrLf & _
                  " 	BulanDesc, " & vbCrLf & _
                  " 	a.DescUrut, "

            ls_SQL = ls_SQL + " 	a.DescName, " & vbCrLf & _
                              " 	max(isnull(b.qty1,0)) qty1, " & vbCrLf & _
                              " 	max(isnull(b.qty2,0)) qty2, " & vbCrLf & _
                              " 	max(isnull(b.qty3,0)) qty3, " & vbCrLf & _
                              " 	max(isnull(b.qty4,0)) qty4, " & vbCrLf & _
                              " 	max(isnull(b.qty5,0)) qty5 " & vbCrLf & _
                              "  from  " & vbCrLf & _
                              "  (  " & vbCrLf & _
                              "  	select a.*,b.*, c.* " & vbCrLf & _
                              "  	from  " & vbCrLf & _
                              "  	(  "

            ls_SQL = ls_SQL + "  		select row_number() over (order by PartNo asc) as NoUrut, * from  " & vbCrLf & _
                              "  		(  " & vbCrLf & _
                              "  			select distinct a.partno, b.PartName from PO_Detail a  " & vbCrLf & _
                              "  			inner join PO_Master c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  " & vbCrLf & _
                              "  			left join MS_Parts b on a.PartNo = b.PartNo  " & vbCrLf & _
                              "  			where FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                              " 			--and a.partno='7184-8544' " & vbCrLf & _
                              "  		) xuz  " & vbCrLf & _
                              "  	) a  " & vbCrLf & _
                              "  	cross join  " & vbCrLf & _
                              "  	(		  "

            ls_SQL = ls_SQL + "  			select row_number() over (order by tahun, bulan asc) as NoUrutBulan, tahun thn,bulan bln,Tgl,BulanDesc   " & vbCrLf & _
                              "  			from ms_period   " & vbCrLf & _
                              "  			where tgl between @period and dateadd(month,2,@period)		  " & vbCrLf & _
                              "  	) b " & vbCrLf & _
                              " 	cross join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select '1'DescUrut, 'Total Forecast' DescName " & vbCrLf & _
                              " 		union all " & vbCrLf & _
                              " 		select '2'DescUrut, 'Total PO' DescName " & vbCrLf & _
                              " 		union all " & vbCrLf & _
                              " 		select '3'DescUrut, 'Total Delivery' DescName "

            ls_SQL = ls_SQL + " 		union all " & vbCrLf & _
                              " 		select '4'DescUrut, 'Balance PO' DescName " & vbCrLf & _
                              " 		union all " & vbCrLf & _
                              " 		select '5'DescUrut, 'Diff %' DescName " & vbCrLf & _
                              " 	)c " & vbCrLf & _
                              "  ) a  " & vbCrLf & _
                              "  left join  " & vbCrLf & _
                              "  ( " & vbCrLf & _
                              "  	select '2' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 1 then POQty end),0) qty1,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 2 then POQty end),0) qty2,  "

            ls_SQL = ls_SQL + "  		isnull(max(case when c.kd = 3 then POQty end),0) qty3,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 4 then POQty end),0) qty4,   " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 5 then POQty end),0) qty5 " & vbCrLf & _
                              "  	from  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a   " & vbCrLf & _
                              "   		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  " & vbCrLf & _
                              "  		left join  " & vbCrLf & _
                              "  		(  " & vbCrLf & _
                              "  			select a.* from PORev_Detail a  " & vbCrLf & _
                              "  			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo  "

            ls_SQL = ls_SQL + "  			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo  " & vbCrLf & _
                              "  			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI' " & vbCrLf & _
                              "  		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  " & vbCrLf & _
                              "  		left join MS_Parts d on b.PartNo = d.PartNo  " & vbCrLf & _
                              "  		where a.FinalApproveDate is not null and Period between @period and dateadd(month,4,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                              "   		group by b.partno,year(period),month(period)   " & vbCrLf & _
                              "  	) a  " & vbCrLf & _
                              "  	inner join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                              "  		where tgl between @period and dateadd(month,4,@period)		  "

            ls_SQL = ls_SQL + "  	) c  " & vbCrLf & _
                              "  		on c.tahun = a.thn and c.bulan = a.bln  " & vbCrLf & _
                              "  	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                              "   " & vbCrLf & _
                              " 	union  all " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	select '1' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 1 then Qty end),0) qty1,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 2 then Qty end),0) qty2,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 3 then Qty end),0) qty3,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 4 then Qty end),0) qty4,   "

            ls_SQL = ls_SQL + "  		isnull(max(case when c.kd = 5 then Qty end),0) qty5 " & vbCrLf & _
                              " 	from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select partno,year(period) thn,month(period) bln,sum(Qty) qty from ms_forecast " & vbCrLf & _
                              " 		where period between @period and dateadd(month,4,@period) " & vbCrLf & _
                              " 		group by PartNo, Period " & vbCrLf & _
                              " 	) a " & vbCrLf & _
                              "  	inner join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                              "  		where tgl between @period and dateadd(month,4,@period)		  "

            ls_SQL = ls_SQL + "  	) c  " & vbCrLf & _
                              "  		on c.tahun = a.thn and c.bulan = a.bln  " & vbCrLf & _
                              " 	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	select '3' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 1 then Qty end),0) qty1,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 2 then Qty end),0) qty2,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 3 then Qty end),0) qty3,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 4 then Qty end),0) qty4,   "

            ls_SQL = ls_SQL + "  		isnull(max(case when c.kd = 5 then Qty end),0) qty5 " & vbCrLf & _
                              " 	from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select partno,left(kanbanno,4) thn,substring(kanbanno,5,2) bln,sum(doqty) qty from dosupplier_detail " & vbCrLf & _
                              " 		where cast(left(kanbanno,8) as date) between @period and dateadd(month,4,@period) " & vbCrLf & _
                              " 		group by partno,left(kanbanno,4),substring(kanbanno,5,2)		 " & vbCrLf & _
                              " 	) a " & vbCrLf & _
                              "  	inner join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                              "  		where tgl between @period and dateadd(month,4,@period)		  "

            ls_SQL = ls_SQL + "  	) c  " & vbCrLf & _
                              "  		on c.tahun = a.thn and c.bulan = a.bln  " & vbCrLf & _
                              " 	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	select '4' SeqNo, a.partno,a.thn,a.bln,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 1 then poqty end),0) - isnull(max(case when c.kd = 1 then Qty end),0) qty1,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 2 then poqty end),0) - isnull(max(case when c.kd = 2 then Qty end),0) qty2,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 3 then poqty end),0) - isnull(max(case when c.kd = 3 then Qty end),0) qty3,  " & vbCrLf & _
                              "  		isnull(max(case when c.kd = 4 then poqty end),0) - isnull(max(case when c.kd = 4 then Qty end),0) qty4,   "

            ls_SQL = ls_SQL + "  		isnull(max(case when c.kd = 5 then poqty end),0) - isnull(max(case when c.kd = 5 then Qty end),0) qty5 " & vbCrLf & _
                              " 	from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select partno,left(kanbanno,4) thn,substring(kanbanno,5,2) bln,sum(doqty) qty from dosupplier_detail " & vbCrLf & _
                              " 		where cast(left(kanbanno,8) as date) between @period and dateadd(month,4,@period) " & vbCrLf & _
                              " 		group by partno,left(kanbanno,4),substring(kanbanno,5,2)		 " & vbCrLf & _
                              " 	) a " & vbCrLf & _
                              " 	left join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a   " & vbCrLf & _
                              "   		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  "

            ls_SQL = ls_SQL + "  		left join  " & vbCrLf & _
                              "  		(  " & vbCrLf & _
                              "  			select a.* from PORev_Detail a  " & vbCrLf & _
                              "  			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo  " & vbCrLf & _
                              "  			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo  " & vbCrLf & _
                              "  			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI' " & vbCrLf & _
                              "  		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  " & vbCrLf & _
                              "  		left join MS_Parts d on b.PartNo = d.PartNo  " & vbCrLf & _
                              "  		where a.FinalApproveDate is not null and Period between @period and dateadd(month,4,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                              "   		group by b.partno,year(period),month(period)   " & vbCrLf & _
                              " 	)b on a.bln = b.bln and a.PartNo = b.PartNo and a.thn = b.thn "

            ls_SQL = ls_SQL + "  	inner join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                              "  		where tgl between @period and dateadd(month,4,@period)		  " & vbCrLf & _
                              "  	) c  " & vbCrLf & _
                              "  		on c.tahun = a.thn and c.bulan = a.bln and b.bln = c.Bulan and b.thn= c.Tahun " & vbCrLf & _
                              " 	group by a.partno,a.thn,a.bln  " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              "  " & vbCrLf & _
                              " 	select '5' SeqNo, a.partno,a.thn,a.bln,  "

            ls_SQL = ls_SQL + "  		case when isnull(max(case when c.kd = 1 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 1 then poqty end),0) - isnull(max(case when c.kd = 1 then Qty end),0)) / isnull(max(case when c.kd = 1 then poqty end),0)) * 100) end  qty1, " & vbCrLf & _
                              "  		case when isnull(max(case when c.kd = 2 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 2 then poqty end),0) - isnull(max(case when c.kd = 2 then Qty end),0)) / isnull(max(case when c.kd = 2 then poqty end),0)) * 100) end  qty2, " & vbCrLf & _
                              "  		case when isnull(max(case when c.kd = 3 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 3 then poqty end),0) - isnull(max(case when c.kd = 3 then Qty end),0)) / isnull(max(case when c.kd = 3 then poqty end),0)) * 100) end  qty3, " & vbCrLf & _
                              "  		case when isnull(max(case when c.kd = 4 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 4 then poqty end),0) - isnull(max(case when c.kd = 4 then Qty end),0)) / isnull(max(case when c.kd = 4 then poqty end),0)) * 100) end  qty4, " & vbCrLf & _
                              "  		case when isnull(max(case when c.kd = 5 then poqty end),0) = 0 then 0 else abs(((isnull(max(case when c.kd = 5 then poqty end),0) - isnull(max(case when c.kd = 5 then Qty end),0)) / isnull(max(case when c.kd = 5 then poqty end),0)) * 100) end  qty5 " & vbCrLf & _
                              " 	from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select partno,year(period) thn,month(period) bln,sum(Qty) qty from ms_forecast " & vbCrLf & _
                              " 		where period between @period and dateadd(month,4,@period) " & vbCrLf & _
                              " 		group by PartNo, Period " & vbCrLf & _
                              " 	) a "

            ls_SQL = ls_SQL + " 	left join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a   " & vbCrLf & _
                              "   		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID  " & vbCrLf & _
                              "  		left join  " & vbCrLf & _
                              "  		(  " & vbCrLf & _
                              "  			select a.* from PORev_Detail a  " & vbCrLf & _
                              "  			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo  " & vbCrLf & _
                              "  			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo  " & vbCrLf & _
                              "  			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) --and a.AffiliateID = 'SUAI' " & vbCrLf & _
                              "  		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID  "

            ls_SQL = ls_SQL + "  		left join MS_Parts d on b.PartNo = d.PartNo  " & vbCrLf & _
                              "  		where a.FinalApproveDate is not null and Period between @period and dateadd(month,4,@period) --and a.AffiliateID = 'SUAI'  " & vbCrLf & _
                              "   		group by b.partno,year(period),month(period)   " & vbCrLf & _
                              " 	)b on a.bln = b.bln and a.PartNo = b.PartNo and a.thn = b.thn " & vbCrLf & _
                              "  	inner join  " & vbCrLf & _
                              "  	(  " & vbCrLf & _
                              "  		select tahun,bulan,row_number() over (order by tgl) kd from MS_Period  " & vbCrLf & _
                              "  		where tgl between @period and dateadd(month,4,@period)		  " & vbCrLf & _
                              "  	) c  " & vbCrLf & _
                              "  		on c.tahun = a.thn and c.bulan = a.bln and b.bln = c.Bulan and b.thn= c.Tahun " & vbCrLf & _
                              " 	group by a.partno,a.thn,a.bln  "

            ls_SQL = ls_SQL + " ) b   " & vbCrLf & _
                              "  	on b.SeqNo = a.DescUrut and b.partno = a.partno and b.thn = a.thn  " & vbCrLf & _
                              " 		and (b.seqno in (2,3,4,5) and b.bln = a.bln or b.seqno = 1 and b.bln >= a.bln)  " & vbCrLf & _
                              " 	group by a.NoUrut, a.NoUrutBulan, " & vbCrLf & _
                              " 	a.PartNo, " & vbCrLf & _
                              " 	a.PartName, " & vbCrLf & _
                              " 	a.thn, " & vbCrLf & _
                              " 	a.bln, " & vbCrLf & _
                              " 	BulanDesc, " & vbCrLf & _
                              " 	a.DescUrut, " & vbCrLf & _
                              " 	a.DescName "

            ls_SQL = ls_SQL + "  order by a.PartNo,	a.thn, a.bln, a.DescUrut " & vbCrLf & _
                              "   " & vbCrLf & _
                              "  " & vbCrLf & _
                              "  "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' PartNo, '' PartName, ''AffiliateID, ''SupplierID, ''MOQ, ''Project, ''PONo, ''Bln1, ''Bln2, ''Bln3, ''Bln4"

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

    Private Sub up_IsiHeader()
        Dim iWeek As Integer

        For iWeek = 1 To 5
            grid.Columns("qty" & iWeek).Caption = Format(DateAdd(DateInterval.Month, iWeek - 1, dtPeriodFrom.Value), "MMM-yy")
        Next
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all " & vbCrLf & _
                 "select distinct RTRIM(a.PartNo) PartCode, PartName from MS_Parts a" & vbCrLf & _
                 "where FinishGoodCls = '2' order by PartCode "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartCode")
                .Columns(0).Width = 85
                .Columns.Add("PartName")
                .Columns(1).Width = 180

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = 0
                txtPartNo.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Report Forecast vs PO " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
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
                '.Cells(5, 1).Value = "(" & Format(dtPeriodFrom.Value, "MMM yyyy") & " Production)"
                '.Cells(7, 1).Value = "Issue Date : " & Format(dtPeriodFrom.Value, "dd MMM yyyy")

                .Cells(8, 6).Value = Format(DateAdd(DateInterval.Month, 1 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(8, 7).Value = Format(DateAdd(DateInterval.Month, 2 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(8, 8).Value = Format(DateAdd(DateInterval.Month, 3 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(8, 9).Value = Format(DateAdd(DateInterval.Month, 4 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(8, 10).Value = Format(DateAdd(DateInterval.Month, 5 - 1, dtPeriodFrom.Value), "MMM")

                Dim NoUrut As String = pData.Rows(0)(0)
                Dim NoStart As Integer = 0 + rowstart
                Dim NoRow As Integer = 0

                For irow = 0 To pData.Rows.Count - 1
                    If pData.Rows(irow)(0) <> NoUrut Then                        
                        .Cells(NoStart, 1, NoStart + NoRow - 1, 1).Merge = True
                        .Cells(NoStart, 2, NoStart + NoRow - 1, 2).Merge = True
                        .Cells(NoStart, 3, NoStart + NoRow - 1, 3).Merge = True                        
                        NoUrut = pData.Rows(irow)(0)
                        NoStart = irow + rowstart
                        NoRow = 0
                    End If
                    NoRow = NoRow + 1                    
                    For icol = 1 To pData.Columns.Count
                        If icol = 4 Then
                            .Cells(irow + rowstart, 1, irow + rowstart, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                            .Cells(irow + rowstart, 1, irow + rowstart, 3).Style.HorizontalAlignment = Style.ExcelVerticalAlignment.Center
                        End If
                        If Trim(pData.Rows(irow)(icol - 1)) = "Diff" Then
                            .Cells(irow + rowstart, icol).Value = Trim(pData.Rows(irow)(icol - 1)) & " %"
                        Else
                            .Cells(irow + rowstart, icol).Value = Trim(pData.Rows(irow)(icol - 1))
                        End If

                    Next
                Next

                Dim rgAll As ExcelRange = .Cells(10, 1, irow + 9, 10)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

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
End Class