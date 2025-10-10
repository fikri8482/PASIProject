Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions

Imports OfficeOpenXml
Imports System.IO


Public Class TraceBackPOHistory
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "D06"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String, pub_Remarks As String, pub_Revision As String
    'Dim pub_FinalApproval As String, pub_DeliveyBy As String
    Dim pub_Period As Date
    Dim colNo As Byte = 1, colPartNo As Byte = 2, colPartName As Byte = 3, colBulan As Byte = 4
    Dim colWK1 As Byte = 5, colWK2 As Byte = 6, colWK3 As Byte = 7, colWK4 As Byte = 8, colWK5 As Byte = 9
    Dim colWK6 As Byte = 10, colWK7 As Byte = 11, colWK8 As Byte = 12, colWK9 As Byte = 13, colWK10 As Byte = 14
    Dim colWK11 As Byte = 15, colWK12 As Byte = 16
    'Dim pub_HeijunkaSttus As Boolean
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            dtPeriodFrom.Value = Now
            up_FillComboPart()
            lblInfo.Text = ""
            up_GridLoadWhenEventChange()
            up_IsiHeader()
        ElseIf IsCallback Then
            If grid.VisibleRowCount = 0 Then Exit Sub
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        grid.JSProperties("cpMessage") = ""
        Call bindData()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()
                    'Call bindPOStatus()

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
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        grid.CollapseAll()

        cboPartNo.Text = clsGlobal.gs_All

        dtPeriodFrom.Value = Now

        'up_FillCombo(dtPeriodFrom.Value)

        'cboPartNo.Items.Clear()

        lblInfo.Text = ""

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then                
                If e.GetValue("SeqNo") <> e.GetValue("bln") Then
                    If e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" _
                    Or e.DataColumn.FieldName = "NoUrut" Then
                        e.Cell.Text = ""
                    End If
                Else
                    e.Cell.BackColor = Color.GreenYellow
                End If
            End If
        End With
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pPeriod As String = Year(dtPeriodFrom.Value) & "-" & Format(Month(dtPeriodFrom.Value), "MM") & "-01"

        If cboPartNo.Text <> clsGlobal.gs_All Then
            pWhere = " where a.partno =  '" & cboPartNo.Text.Trim & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "declare @period date " & vbCrLf & _
                  " set @period = '" & Year(dtPeriodFrom.Value) & "-" & Month(dtPeriodFrom.Value) & "-01' " & vbCrLf & _
                  " select a.SeqNo, a.NoUrut ,a.PartNo,a.PartName,a.thn,a.bln,BulanDesc,isnull(qty1,0)qty1,isnull(qty2,0)qty2,isnull(qty3,0)qty3, " & vbCrLf & _
                  " isnull(qty4,0)qty4,isnull(qty5,0)qty5,isnull(qty6,0)qty6,isnull(qty7,0)qty7,isnull(qty8,0)qty8,isnull(qty9,0)qty9,isnull(qty10,0)qty10, " & vbCrLf & _
                  " isnull(qty11,0)qty11,isnull(qty12,0)qty12 " & vbCrLf & _
                  " from " & vbCrLf & _
                  " ( " & vbCrLf & _
                  " 	select '" & Month(dtPeriodFrom.Value) & "' SeqNo, a.*,b.* " & vbCrLf & _
                  " 	from " & vbCrLf & _
                  " 	( " & vbCrLf & _
                  " 		select row_number() over (order by PartNo asc) as NoUrut, * from " & vbCrLf & _
                  " 		( " & vbCrLf & _
                  " 			select distinct a.partno, b.PartName from PO_Detail a " & vbCrLf

            ls_SQL = ls_SQL + " 			inner join PO_Master c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID " & vbCrLf & _
                              " 			left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " 			where FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) and a.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                              " 		) xuz " & vbCrLf & _
                              " 	) a " & vbCrLf & _
                              " 	cross join " & vbCrLf & _
                              " 	(		 " & vbCrLf & _
                              " 			select tahun thn,bulan bln,Tgl,BulanDesc  " & vbCrLf & _
                              " 			from ms_period  " & vbCrLf & _
                              " 			where tgl between @period and dateadd(month,11,@period)		 " & vbCrLf & _
                              " 	) b " & vbCrLf

            ls_SQL = ls_SQL + " ) a " & vbCrLf & _
                              " left join " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select '2'SeqNo, a.partno,a.thn,a.bln, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 1 then POQty end),0)+isnull(sum(case when b.kd = 1 then Qty end),0) qty1, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 2 then POQty end),0)+isnull(sum(case when b.kd = 2 then Qty end),0) qty2, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 3 then POQty end),0)+isnull(sum(case when b.kd = 3 then Qty end),0) qty3, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 4 then POQty end),0)+isnull(sum(case when b.kd = 4 then Qty end),0) qty4, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 5 then POQty end),0)+isnull(sum(case when b.kd = 5 then Qty end),0) qty5, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 6 then POQty end),0)+isnull(sum(case when b.kd = 6 then Qty end),0) qty6, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 7 then POQty end),0)+isnull(sum(case when b.kd = 7 then Qty end),0) qty7, " & vbCrLf

            ls_SQL = ls_SQL + " 		isnull(max(case when c.kd = 8 then POQty end),0)+isnull(sum(case when b.kd = 8 then Qty end),0) qty8, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 9 then POQty end),0)+isnull(sum(case when b.kd = 9 then Qty end),0) qty9, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 10 then POQty end),0)+isnull(sum(case when b.kd = 10 then Qty end),0) qty10, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 11 then POQty end),0)+isnull(sum(case when b.kd = 11 then Qty end),0) qty11, " & vbCrLf & _
                              " 		isnull(max(case when c.kd = 12 then POQty end),0)+isnull(sum(case when b.kd = 12 then Qty end),0) qty12 " & vbCrLf & _
                              " 	from " & vbCrLf & _
                              " 	( " & vbCrLf
            ls_SQL = ls_SQL + " 		select b.PartNo,year(Period) thn,month(period) bln,sum(isnull(c.POQty,b.POQty)) poqty from PO_Master a  " & vbCrLf & _
                              "  		inner join PO_Detail b on a.PONo = b.PONo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                              " 		left join " & vbCrLf & _
                              " 		( " & vbCrLf & _
                              " 			select a.* from PORev_Detail a " & vbCrLf & _
                              " 			inner join PORev_Master b on a.PONo = b.PONo and a.PORevNo = b.PORevNo and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 			inner join (select MAX(SeqNo) SeqNo, PONo from PORev_Detail po group by PONo) c on a.PONo = c.PONo and a.SeqNo = c.SeqNo " & vbCrLf & _
                              " 			where b.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) and a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " 		)c on a.PONo = c.PONo and a.AffiliateID = c.AffiliateID and a.SupplierID = c.SupplierID " & vbCrLf & _
                              " 		left join MS_Parts d on b.PartNo = d.PartNo " & vbCrLf & _
                              " 		where a.FinalApproveDate is not null and Period between @period and dateadd(month,11,@period) and a.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                              "  		group by b.partno,year(period),month(period)  " & vbCrLf

            ls_SQL = ls_SQL + " 	) a " & vbCrLf & _
                              " 	inner join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select tahun,bulan,row_number() over (order by tgl) kd,tgl tgl1,dateadd(month,3,tgl) tgl2 from MS_Period " & vbCrLf & _
                              " 		where tgl between @period and dateadd(month,11,@period)		 " & vbCrLf & _
                              " 	) c " & vbCrLf & _
                              " 		on c.tahun = a.thn and c.bulan = a.bln " & vbCrLf & _
                              " 	left join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select partno,year(period) thn,month(period) bln,tgl,kd,Qty from MS_Forecast a " & vbCrLf & _
                              " 		inner join " & vbCrLf

            ls_SQL = ls_SQL + " 		( " & vbCrLf & _
                              " 			select tahun,bulan,tgl,row_number() over (order by tgl) kd from MS_Period " & vbCrLf & _
                              " 			where tgl between @period and dateadd(month,11,@period)		 " & vbCrLf & _
                              " 		) c " & vbCrLf & _
                              " 			on c.tahun = year(period) and c.bulan = month(period) " & vbCrLf & _
                              " 	) b " & vbCrLf & _
                              " 		on b.partno = a.partno and b.tgl > c.tgl1 and b.tgl <= c.tgl2 " & vbCrLf & _
                              " 	group by a.partno,a.thn,a.bln " & vbCrLf & _
                              " ) b  " & vbCrLf & _
                              " 	on b.partno = a.partno and b.thn = a.thn and b.bln = a.bln " & vbCrLf & _
                              " " & pWhere & "" & vbCrLf & _
                              " order by 3,4,5,6,8 " & vbCrLf

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

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' PartNo, '' PartName, " & vbCrLf & _
                  " 0 qty1, 0 qty2, 0 qty3, 0 qty4, 0 qty5, " & vbCrLf & _
                  " 0 qty6, 0 qty7, 0 qty8, 0 qty9, 0 qty10, " & vbCrLf & _
                  " 0 qty11, 0 qty12 "

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

        For iWeek = 1 To 12            
            grid.Columns("qty" & iWeek).Caption = Format(DateAdd(DateInterval.Month, iWeek - 1, dtPeriodFrom.Value), "MMM-yy")
        Next
    End Sub

    Private Sub up_FillComboPart()
        Dim ls_SQL As String = ""

        ls_SQL = "select '" & clsGlobal.gs_All & "'PartNo, '" & clsGlobal.gs_All & "'PartName union all select RTRIM(a.PartNo)PartNo, RTRIM(b.PartName) PartName from MS_PartMapping a left join MS_Parts b on a.PartNo = b.PartNo where AffiliateID = '" & Session("AffiliateID") & "'order by PartNo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 85
                .Columns.Add("PartName")
                .Columns(1).Width = 180

                .TextField = "PartNo"
                .DataBind()
                .SelectedIndex = 0
            End With

            txtPartName.Text = clsGlobal.gs_All

            sqlConn.Close()
        End Using
    End Sub

    Private Sub EpPlusExportExcel()
        'If grid.VisibleRowCount = 0 Then
        '    Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
        '    Session("G01Msg") = lblInfo.Text
        '    Return
        'End If

        'btnExcel.Enabled = False

        'Dim rowStart As Integer = 0, rowEnd As Integer = 0, lb_MERGE As Boolean = False

        'Dim fi As New FileInfo(Server.MapPath("~\PurchaseOrderRevision\Trace Back PO History.xlsx"))
        'If fi.Exists Then
        '    fi.Delete()
        '    fi = New FileInfo(Server.MapPath("~\PurchaseOrderRevision\Trace Back PO History.xlsx"))
        'End If
        'Dim exl As New ExcelPackage(fi)
        'Dim ws As ExcelWorksheet
        'ws = exl.Workbook.Worksheets.Add("Sheet1")
        'ws.Cells(1, 1, 100, 100).Style.Font.Name = "Calibri"
        'ws.Cells(1, 1, 100, 100).Style.Font.Size = 9

        'ws.Cells(1, 1).Value = "TRACE BACK PO HISTORY"
        'ws.Cells(1, 1).Style.Font.Size = 14
        'ws.Cells(1, 1).Style.Font.Bold = True
        'ws.Cells(1, 1).Style.Font.UnderLine = True

        'ws.Cells(3, 1).Value = "PERIOD"
        'ws.Cells(3, colPONo).Value = ": " & dtPOPeriodFrom.Text & " - " & dtPOPeriodTo.Text

        'With ws
        '    Dim space As Integer = 5

        '    'Call EpPlusFormatExcel(ws)

        '    ' Details
        '    '--------------------------------------------------------------------------------
        '    .Cells(space, colNo).Value = "NO"
        '    .Cells(space, colPeriod).Value = "PERIOD"
        '    .Cells(space, colPONo).Value = "PO NO."
        '    .Cells(space, colSupplierCode).Value = "SUPPLIER CODE"
        '    .Cells(space, colSupplierName).Value = "SUPPLIER NAME"
        '    .Cells(space, colPOKanban).Value = "PO KANBAN"
        '    .Cells(space, colKanbanNo).Value = "KANBAN NO."
        '    '.Cells(space, colKanbanSeqNo).Value = "KANBAN SEQ. NO."
        '    .Cells(space, colSupplierPlanDelDate).Value = "SUPPLIER PLAN DELIVERY DATE"
        '    .Cells(space, colSupplierDelDate).Value = "SUPPLIER DELIVERY DATE"
        '    .Cells(space, colSupplierSJNo).Value = "SUPPLIER SURAT JALAN NO."
        '    .Cells(space, colPASIRecDate).Value = "PASI RECEIVE DATE"
        '    .Cells(space, colPASIDelDate).Value = "PASI DELIVERY DATE"
        '    .Cells(space, colPASISJNo).Value = "PASI SURAT JALAN NO."
        '    .Cells(space, colAffiliateRecDate).Value = "AFFILIATE RECEIVE DATE"
        '    .Cells(space, colPartNo).Value = "PART NO."
        '    .Cells(space, colPartName).Value = "PART NAME"
        '    .Cells(space, colUOM).Value = "UOM"
        '    .Cells(space, colSupplierDeliveryQty).Value = "SUPPLIER DELIVERY QTY"
        '    .Cells(space, colPASIReceivingQty).Value = "PASI RECEIVING QTY"
        '    .Cells(space, colPASIDeliveryQty).Value = "PASI DELIVERY QTY"
        '    .Cells(space, colAffiliateReceivingQty).Value = "AFFILIATE RECEIVING QTY"
        '    .Cells(space, colPASIInvQty).Value = "PASI INVOICE QTY"
        '    .Cells(space, colPASIInvNo).Value = "PASI INVOICE NO"
        '    .Cells(space, colPASIInvDate).Value = "PASI INVOICE DATE"
        '    .Cells(space, colPASIInvCurr).Value = "PASI INVOICE"
        '    .Cells(space, colPASIInvAmount).Value = ""
        '    .Cells(space + 1, colPASIInvCurr).Value = "CURR"
        '    .Cells(space + 1, colPASIInvAmount).Value = "AMOUNT"

        '    'Merge - Wrap Text - Alignment
        '    Dim iCol As Integer = 0, iNextCol As Integer = 0
        '    For iCol = colNo To colPASIInvDate
        '        .Cells(space, (colNo + iNextCol), space + 1, (colNo + iNextCol)).Merge = True
        '        iNextCol = iNextCol + 1
        '    Next iCol
        '    iNextCol = 0
        '    ws.Cells(space, colPASIInvCurr, space, colPASIInvAmount).Merge = True
        '    .Cells(space, colNo, space + 1, colPASIInvAmount).Style.WrapText = True
        '    .Cells(space, colNo, space + 1, colPASIInvAmount).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '    .Cells(space, colNo, space + 1, colPASIInvAmount).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center


        '    space = 7
        '    For iRow = 0 To grid.VisibleRowCount - 1
        '        'fill data 
        '        .Cells(iRow + space, colNo).Value = Trim(grid.GetRowValues(iRow, "ColNo"))
        '        .Cells(iRow + space, colPeriod).Value = Trim(grid.GetRowValues(iRow, "Period"))
        '        .Cells(iRow + space, colPONo).Value = Trim(grid.GetRowValues(iRow, "PONo"))
        '        .Cells(iRow + space, colSupplierCode).Value = Trim(grid.GetRowValues(iRow, "SupplierCode"))
        '        .Cells(iRow + space, colSupplierName).Value = Trim(grid.GetRowValues(iRow, "SupplierName"))
        '        .Cells(iRow + space, colPOKanban).Value = Trim(grid.GetRowValues(iRow, "POKanban"))
        '        .Cells(iRow + space, colKanbanNo).Value = Trim(grid.GetRowValues(iRow, "KanbanNo"))
        '        '.Cells(iRow + space, colKanbanSeqNo).Value = Trim(grid.GetRowValues(iRow, "KanbanSeqNo"))
        '        .Cells(iRow + space, colSupplierPlanDelDate).Value = Trim(grid.GetRowValues(iRow, "SupplierPlanDeliveryDate"))
        '        .Cells(iRow + space, colSupplierDelDate).Value = Trim(grid.GetRowValues(iRow, "SupplierDeliveryDate"))
        '        .Cells(iRow + space, colSupplierSJNo).Value = Trim(grid.GetRowValues(iRow, "SupplierSJNo"))
        '        .Cells(iRow + space, colPASIRecDate).Value = Trim(grid.GetRowValues(iRow, "PASIReceiveDate"))
        '        .Cells(iRow + space, colPASIDelDate).Value = Trim(grid.GetRowValues(iRow, "PASIDeliveryDate"))
        '        .Cells(iRow + space, colPASISJNo).Value = Trim(grid.GetRowValues(iRow, "PASISJNo"))
        '        .Cells(iRow + space, colAffiliateRecDate).Value = Trim(grid.GetRowValues(iRow, "AffiliateReceiveDate"))
        '        .Cells(iRow + space, colPartNo).Value = Trim(grid.GetRowValues(iRow, "PartNo"))
        '        .Cells(iRow + space, colPartName).Value = Trim(grid.GetRowValues(iRow, "PartName"))
        '        .Cells(iRow + space, colUOM).Value = Trim(grid.GetRowValues(iRow, "UOM"))
        '        If Trim(grid.GetRowValues(iRow, "ColNo")) = "" Then
        '            'Detail
        '            .Cells(iRow + space, colSupplierDeliveryQty).Value = FormatNumber(IIf(Trim(grid.GetRowValues(iRow, "SupplierDeliveryQty")) = "", 0, Trim(grid.GetRowValues(iRow, "SupplierDeliveryQty"))), 0, TriState.True)
        '            .Cells(iRow + space, colPASIReceivingQty).Value = FormatNumber(IIf(Trim(grid.GetRowValues(iRow, "PASIReceivingQty")) = "", 0, Trim(grid.GetRowValues(iRow, "PASIReceivingQty"))), 0, TriState.True)
        '            .Cells(iRow + space, colPASIDeliveryQty).Value = FormatNumber(IIf(Trim(grid.GetRowValues(iRow, "PASIDeliveryQty")) = "", 0, Trim(grid.GetRowValues(iRow, "PASIDeliveryQty"))), 0, TriState.True)
        '            .Cells(iRow + space, colAffiliateReceivingQty).Value = FormatNumber(IIf(Trim(grid.GetRowValues(iRow, "AffiliateReceivingQty")) = "", 0, Trim(grid.GetRowValues(iRow, "AffiliateReceivingQty"))), 0, TriState.True)
        '            .Cells(iRow + space, colPASIInvQty).Value = FormatNumber(IIf(Trim(grid.GetRowValues(iRow, "PASIInvoiceQty")) = "", 0, Trim(grid.GetRowValues(iRow, "PASIInvoiceQty"))), 0, TriState.True)
        '            .Cells(iRow + space, colPASIInvNo).Value = Trim(grid.GetRowValues(iRow, "PASIInvoiceNo"))
        '            .Cells(iRow + space, colPASIInvDate).Value = Trim(grid.GetRowValues(iRow, "PASIInvoiceDate"))
        '            .Cells(iRow + space, colPASIInvCurr).Value = Trim(grid.GetRowValues(iRow, "PASIInvoiceCurr"))
        '            .Cells(iRow + space, colPASIInvAmount).Value = FormatNumber(IIf(Trim(grid.GetRowValues(iRow, "PASIInvoiceAmount")) = "", 0, Trim(grid.GetRowValues(iRow, "PASIInvoiceAmount"))), 0, TriState.True)
        '        Else
        '            'Header
        '            .Cells(iRow + space, colSupplierDeliveryQty).Value = ""
        '            .Cells(iRow + space, colPASIReceivingQty).Value = ""
        '            .Cells(iRow + space, colPASIDeliveryQty).Value = ""
        '            .Cells(iRow + space, colAffiliateReceivingQty).Value = ""
        '            .Cells(iRow + space, colPASIInvQty).Value = ""
        '            .Cells(iRow + space, colPASIInvNo).Value = ""
        '            .Cells(iRow + space, colPASIInvDate).Value = ""
        '            .Cells(iRow + space, colPASIInvCurr).Value = ""
        '            .Cells(iRow + space, colPASIInvAmount).Value = ""
        '        End If

        '        'ALIGNMENT
        '        .Cells(iRow + space, colNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
        '        .Cells(iRow + space, colPeriod).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colPONo, iRow + space, colSupplierName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
        '        .Cells(iRow + space, colPOKanban).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colKanbanNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
        '        '.Cells(iRow + space, colKanbanSeqNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
        '        .Cells(iRow + space, colSupplierPlanDelDate, iRow + space, colSupplierDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colSupplierSJNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
        '        .Cells(iRow + space, colPASIRecDate, iRow + space, colPASIDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colPASISJNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
        '        .Cells(iRow + space, colAffiliateRecDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colPartNo, iRow + space, colPartName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
        '        .Cells(iRow + space, colUOM).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colSupplierDeliveryQty, iRow + space, colPASIInvQty).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
        '        .Cells(iRow + space, colPASIInvNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
        '        .Cells(iRow + space, colPASIInvDate, iRow + space, colPASIInvCurr).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
        '        .Cells(iRow + space, colPASIInvAmount).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left


        '        'FORMAT NUMBER
        '        '.Cells(iRow + space, colSupplierDeliveryQty, iRow + space, colPASIInvQty).Style.Numberformat.Format = "#,###"
        '        If .Cells(iRow + space, colSupplierDeliveryQty).Value <> "0" Then
        '            .Cells(iRow + space, colSupplierDeliveryQty).Style.Numberformat.Format = "#,###"
        '        End If
        '        If .Cells(iRow + space, colPASIReceivingQty).Value <> "0" Then
        '            .Cells(iRow + space, colPASIReceivingQty).Style.Numberformat.Format = "#,###"
        '        End If
        '        If .Cells(iRow + space, colPASIDeliveryQty).Value <> "0" Then
        '            .Cells(iRow + space, colPASIDeliveryQty).Style.Numberformat.Format = "#,###"
        '        End If
        '        If .Cells(iRow + space, colAffiliateReceivingQty).Value <> "0" Then
        '            .Cells(iRow + space, colAffiliateReceivingQty).Style.Numberformat.Format = "#,###"
        '        End If
        '        If .Cells(iRow + space, colPASIInvQty).Value <> "0" Then
        '            .Cells(iRow + space, colPASIInvQty).Style.Numberformat.Format = "#,###"
        '        End If

        '        If .Cells(iRow + space, colPASIInvAmount).Value <> "0" Then
        '            .Cells(iRow + space, colPASIInvAmount).Style.Numberformat.Format = "#,###.00"
        '        End If



        '        'MERGE ROWS
        '        If .Cells(iRow + space, colNo).Value = "" Then
        '            lb_MERGE = True
        '            rowEnd = iRow + space

        '        Else
        '            If lb_MERGE = True Then
        '                For iCol = colNo To colAffiliateRecDate
        '                    .Cells(rowStart, (colNo + iNextCol), rowEnd, (colNo + iNextCol)).Merge = True
        '                    iNextCol = iNextCol + 1
        '                Next iCol
        '                iNextCol = 0

        '                .Cells(rowStart, colNo, rowEnd, colAffiliateRecDate).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Top

        '                lb_MERGE = False
        '            End If

        '            rowStart = iRow + space
        '        End If
        '    Next

        '    'MERGE END ROW
        '    If lb_MERGE = True Then
        '        For iCol = colNo To colAffiliateRecDate
        '            .Cells(rowStart, (colNo + iNextCol), rowEnd, (colNo + iNextCol)).Merge = True
        '            iNextCol = iNextCol + 1
        '        Next iCol
        '        iNextCol = 0

        '        .Cells(rowStart, colNo, rowEnd, colAffiliateRecDate).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Top

        '        lb_MERGE = False
        '    End If


        '    'BORDER
        '    Dim rgAll As ExcelRange = .Cells(space - 2, colNo, grid.VisibleRowCount + (space - 1), colCount - 1)
        '    EpPlusDrawAllBorders(rgAll)


        '    'WIDTH
        '    .Column(colNo).Width = 3.5
        '    .Column(colPONo).Width = 15
        '    .Column(colSupplierCode).Width = 9
        '    .Column(colSupplierName).Width = 27
        '    .Column(colPOKanban).Width = 8
        '    .Column(colKanbanNo).Width = 10
        '    '.Column(colKanbanSeqNo).Width = 8
        '    .Column(colSupplierPlanDelDate).Width = 12.5
        '    .Column(colSupplierDelDate).Width = 12.5
        '    .Column(colSupplierSJNo).Width = 15
        '    .Column(colPASIRecDate).Width = 12
        '    .Column(colPASIDelDate).Width = 12.5
        '    .Column(colPASISJNo).Width = 18
        '    .Column(colAffiliateRecDate).Width = 12.5
        '    .Column(colPartNo).Width = 14
        '    .Column(colPartName).Width = 28
        '    .Column(colUOM).Width = 5
        '    .Column(colSupplierDeliveryQty).Width = 12.5
        '    .Column(colPASIReceivingQty).Width = 12.5
        '    .Column(colPASIDeliveryQty).Width = 12.5
        '    .Column(colAffiliateReceivingQty).Width = 12.5
        '    .Column(colPASIInvQty).Width = 11
        '    .Column(colPASIInvNo).Width = 15
        '    .Column(colPASIInvDate).Width = 11.5
        '    .Column(colPASIInvCurr).Width = 5
        '    .Column(colPASIInvAmount).Width = 13


        '    'save to file
        '    exl.Save()

        '    'redirect to file download
        '    DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)
        'End With

    End Sub
#End Region


End Class