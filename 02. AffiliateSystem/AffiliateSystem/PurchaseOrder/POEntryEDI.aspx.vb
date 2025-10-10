Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO

Public Class POEntryEDI
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "D04"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim log As String = ""
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        dtPeriodFrom.Value = Now
        lblInfo.Text = ""

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
           

        Else
            Ext = Server.MapPath("")
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Uploader.NullText = "Click here to browse files..."
        'ProgressBar.Maximum = 0
        'ProgressBar.Minimum = 0
        'ProgressBar.Value = 0

        lblInfo.Text = ""

        Uploader.Enabled = True
        btnUpload.Enabled = True        
        btnClear.Enabled = False
        'memo.Text = ""
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        'Uploader.Enabled = False
        'btnUpload.Enabled = False
        'btnClear.Enabled = True

        'up_Import()
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        'If rdrKanban2.Checked = True Then
        '    pWhereKanban = " and b.KanbanCls = '1'"
        'End If

        'If rdrKanban3.Checked = True Then
        '    pWhereKanban = " and b.KanbanCls = '0'"
        'End If

        'If rdrDiff2.Checked = True Then
        '    pWhereDifference = " where POQty <> POQtyOld or  " & vbCrLf & _
        '          " DeliveryD1 <> DeliveryD1Old or DeliveryD2 <> DeliveryD2Old or DeliveryD3 <> DeliveryD3Old or DeliveryD4 <> DeliveryD4Old or DeliveryD5 <> DeliveryD5Old or " & vbCrLf & _
        '          " DeliveryD6 <> DeliveryD6Old or DeliveryD7 <> DeliveryD7Old or DeliveryD8 <> DeliveryD8Old or DeliveryD9 <> DeliveryD9Old or DeliveryD10 <> DeliveryD10Old or " & vbCrLf & _
        '          " DeliveryD11 <> DeliveryD11Old or DeliveryD12 <> DeliveryD12Old or DeliveryD13 <> DeliveryD13Old or DeliveryD14 <> DeliveryD14Old or DeliveryD15 <> DeliveryD15Old or " & vbCrLf & _
        '          " DeliveryD16 <> DeliveryD16Old or DeliveryD17 <> DeliveryD17Old or DeliveryD18 <> DeliveryD18Old or DeliveryD19 <> DeliveryD19Old or DeliveryD20 <> DeliveryD20Old or " & vbCrLf & _
        '          " DeliveryD21 <> DeliveryD21Old or DeliveryD22 <> DeliveryD22Old or DeliveryD23 <> DeliveryD23Old or DeliveryD24 <> DeliveryD24Old or DeliveryD25 <> DeliveryD25Old or " & vbCrLf & _
        '          " DeliveryD26 <> DeliveryD26Old or DeliveryD27 <> DeliveryD27Old or DeliveryD28 <> DeliveryD28Old or DeliveryD29 <> DeliveryD29Old or DeliveryD30 <> DeliveryD30Old or " & vbCrLf & _
        '          " DeliveryD31 <> DeliveryD31Old "
        'End If

        'If rdrDiff3.Checked = True Then
        '    pWhereDifference = " where POQty = POQtyOld and  " & vbCrLf & _
        '          " DeliveryD1 = DeliveryD1Old = DeliveryD2 = DeliveryD2Old and DeliveryD3 = DeliveryD3Old and DeliveryD4 = DeliveryD4Old and DeliveryD5 = DeliveryD5Old and " & vbCrLf & _
        '          " DeliveryD6 = DeliveryD6Old = DeliveryD7 = DeliveryD7Old and DeliveryD8 = DeliveryD8Old and DeliveryD9 = DeliveryD9Old and DeliveryD10 = DeliveryD10Old and " & vbCrLf & _
        '          " DeliveryD11 = DeliveryD11Old = DeliveryD12 = DeliveryD12Old and DeliveryD13 = DeliveryD13Old and DeliveryD14 = DeliveryD14Old and DeliveryD15 = DeliveryD15Old and " & vbCrLf & _
        '          " DeliveryD16 = DeliveryD16Old = DeliveryD17 = DeliveryD17Old and DeliveryD18 = DeliveryD18Old and DeliveryD19 = DeliveryD19Old and DeliveryD20 = DeliveryD20Old and " & vbCrLf & _
        '          " DeliveryD21 = DeliveryD21Old = DeliveryD22 = DeliveryD22Old and DeliveryD23 = DeliveryD23Old and DeliveryD24 = DeliveryD24Old and DeliveryD25 = DeliveryD25Old and " & vbCrLf & _
        '          " DeliveryD26 = DeliveryD26Old = DeliveryD27 = DeliveryD27Old and DeliveryD28 = DeliveryD28Old and DeliveryD29 = DeliveryD29Old and DeliveryD30 = DeliveryD30Old and " & vbCrLf & _
        '          " DeliveryD31 = DeliveryD31Old "
        'End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select tbl2.NoUrut, tbl1.AffiliateName, tbl2.PartNo, tbl1.PartNo1, tbl2.PartName, tbl2.KanbanCls, tbl2.UnitDesc, tbl2.MOQ, tbl2.QtyBox, tbl2.Maker, " & vbCrLf & _
                  " 	ISNULL(POQty,0)POQty, ISNULL(POQtyOld,0)POQtyOld, CurrDesc, Price, Amount, ForecastN1, ForecastN2, ForecastN3, " & vbCrLf & _
                  " 	ISNULL(DeliveryD1,0)DeliveryD1, ISNULL(DeliveryD2,0)DeliveryD2, ISNULL(DeliveryD3,0)DeliveryD3, ISNULL(DeliveryD4,0)DeliveryD4, ISNULL(DeliveryD5,0)DeliveryD5,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD6,0)DeliveryD6, ISNULL(DeliveryD7,0)DeliveryD7, ISNULL(DeliveryD8,0)DeliveryD8, ISNULL(DeliveryD9,0)DeliveryD9, ISNULL(DeliveryD10,0)DeliveryD10,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD11,0)DeliveryD11, ISNULL(DeliveryD12,0)DeliveryD12, ISNULL(DeliveryD13,0)DeliveryD13, ISNULL(DeliveryD14,0)DeliveryD14, ISNULL(DeliveryD15,0)DeliveryD15,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD16,0)DeliveryD16, ISNULL(DeliveryD17,0)DeliveryD17, ISNULL(DeliveryD18,0)DeliveryD18, ISNULL(DeliveryD19,0)DeliveryD19, ISNULL(DeliveryD20,0)DeliveryD20,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD21,0)DeliveryD21, ISNULL(DeliveryD22,0)DeliveryD22, ISNULL(DeliveryD23,0)DeliveryD23, ISNULL(DeliveryD24,0)DeliveryD24, ISNULL(DeliveryD25,0)DeliveryD25,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD26,0)DeliveryD26, ISNULL(DeliveryD27,0)DeliveryD27, ISNULL(DeliveryD28,0)DeliveryD28, ISNULL(DeliveryD29,0)DeliveryD29, ISNULL(DeliveryD30,0)DeliveryD30,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD31,0)DeliveryD31, " & vbCrLf & _
                  " 	ISNULL(DeliveryD1Old,0)DeliveryD1Old, ISNULL(DeliveryD2Old,0)DeliveryD2Old, ISNULL(DeliveryD3Old,0)DeliveryD3Old, ISNULL(DeliveryD4Old,0)DeliveryD4Old, ISNULL(DeliveryD5Old,0)DeliveryD5Old,  " & vbCrLf & _
                  " 	ISNULL(DeliveryD6Old,0)DeliveryD6Old, ISNULL(DeliveryD7Old,0)DeliveryD7Old, ISNULL(DeliveryD8Old,0)DeliveryD8Old, ISNULL(DeliveryD9Old,0)DeliveryD9Old, ISNULL(DeliveryD10Old,0)DeliveryD10Old,  "

            ls_SQL = ls_SQL + " 	ISNULL(DeliveryD11Old,0)DeliveryD11Old, ISNULL(DeliveryD12Old,0)DeliveryD12Old, ISNULL(DeliveryD13Old,0)DeliveryD13Old, ISNULL(DeliveryD14Old,0)DeliveryD14Old, ISNULL(DeliveryD15Old,0)DeliveryD15Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD16Old,0)DeliveryD16Old, ISNULL(DeliveryD17Old,0)DeliveryD17Old, ISNULL(DeliveryD18Old,0)DeliveryD18Old, ISNULL(DeliveryD19Old,0)DeliveryD19Old, ISNULL(DeliveryD20Old,0)DeliveryD20Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD21Old,0)DeliveryD21Old, ISNULL(DeliveryD22Old,0)DeliveryD22Old, ISNULL(DeliveryD23Old,0)DeliveryD23Old, ISNULL(DeliveryD24Old,0)DeliveryD24Old, ISNULL(DeliveryD25Old,0)DeliveryD25Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD26Old,0)DeliveryD26Old, ISNULL(DeliveryD27Old,0)DeliveryD27Old, ISNULL(DeliveryD28Old,0)DeliveryD28Old, ISNULL(DeliveryD29Old,0)DeliveryD29Old, ISNULL(DeliveryD30Old,0)DeliveryD30Old,  " & vbCrLf & _
                              " 	ISNULL(DeliveryD31Old,0)DeliveryD31Old " & vbCrLf & _
                              " from  " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select * from " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select '1' NoUrutDesc, 'BY AFFILIATE' AffiliateName " & vbCrLf & _
                              " 		union all "

            ls_SQL = ls_SQL + " 		select '2' NoUrutDesc, 'BY PASI' AffiliateName " & vbCrLf & _
                              " 		union all " & vbCrLf & _
                              " 		select '3' NoUrutDesc, 'BY SUPPLIER' AffiliateName " & vbCrLf & _
                              " 	)tbla " & vbCrLf & _
                              " 	cross join " & vbCrLf & _
                              " 	( " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			--row_number() over (order by b.PartNo asc) as NoUrut, " & vbCrLf & _
                              " 			b.PartNo, b.PartNo PartNo1 " & vbCrLf & _
                              " 		from PORev_Master a " & vbCrLf & _
                              " 		inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo "

            ls_SQL = ls_SQL + " 		inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 		inner join MS_UnitCls d on d.UnitCls = c.UnitCls " & vbCrLf & _
                              " 		where a.PONo = '" & ls_SQL & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' " & pWhereKanban & "" & vbCrLf & _
                              " 	)tb1b " & vbCrLf & _
                              " )tbl1 " & vbCrLf & _
                              " left join " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		convert(char,row_number() over (order by b.PartNo asc))as NoUrut, 'BY AFFILIATE' AffiliateName, '1' NoUrutDesc,  " & vbCrLf & _
                              " 		b.PartNo, b.PartNo PartNo1, c.PartName, case when c.KanbanCls = '1' then 'Yes' else 'No' end KanbanCls, d.Description UnitDesc, " & vbCrLf & _
                              " 		c.MOQ, c.QtyBox, c.Maker, b.POQty, 0 POQtyOld, e.Description CurrDesc, b.Price, b.Amount, "

            ls_SQL = ls_SQL + " 		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'2015-04-13'))),0),  " & vbCrLf & _
                              " 		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'2015-04-13'))),0),  " & vbCrLf & _
                              " 		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'2015-04-13'))),0),  " & vbCrLf & _
                              "  		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5, " & vbCrLf & _
                              "  		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,  " & vbCrLf & _
                              "  		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,  " & vbCrLf & _
                              "  		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,  " & vbCrLf & _
                              "  		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,  " & vbCrLf & _
                              "  		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,  " & vbCrLf & _
                              "  		b.DeliveryD31, " & vbCrLf & _
                              "  		0 DeliveryD1Old, 0 DeliveryD2Old, 0 DeliveryD3Old, 0 DeliveryD4Old, 0 DeliveryD5Old, "

            ls_SQL = ls_SQL + "  		0 DeliveryD6Old, 0 DeliveryD7Old, 0 DeliveryD8Old, 0 DeliveryD9Old, 0 DeliveryD10Old,  " & vbCrLf & _
                              "  		0 DeliveryD11Old, 0 DeliveryD12Old, 0 DeliveryD13Old, 0 DeliveryD14Old, 0 DeliveryD15Old,  " & vbCrLf & _
                              "  		0 DeliveryD16Old, 0 DeliveryD17Old, 0 DeliveryD18Old, 0 DeliveryD19Old, 0 DeliveryD20Old,  " & vbCrLf & _
                              "  		0 DeliveryD21Old, 0 DeliveryD22Old, 0 DeliveryD23Old, 0 DeliveryD24Old, 0 DeliveryD25Old,  " & vbCrLf & _
                              "  		0 DeliveryD26Old, 0 DeliveryD27Old, 0 DeliveryD28Old, 0 DeliveryD29Old, 0 DeliveryD30Old,  " & vbCrLf & _
                              "  		0 DeliveryD31Old " & vbCrLf & _
                              " 		from PORev_Master a " & vbCrLf & _
                              " 	inner join PORev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	inner join MS_CurrCls e on e.CurrCls = b.CurrCls "

            ls_SQL = ls_SQL + " 	where a.PONo = '" & ls_SQL & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		'' NoUrut,'BY PASI' AffiliateName, '2' NoUrutDesc,  " & vbCrLf & _
                              " 		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, c.Maker, b.POQty, b.POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              " 		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'2015-04-13'))),0),  " & vbCrLf & _
                              " 		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'2015-04-13'))),0),  " & vbCrLf & _
                              " 		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'2015-04-13'))),0),  " & vbCrLf & _
                              "  		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5, " & vbCrLf & _
                              "  		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,  " & vbCrLf & _
                              "  		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,  "

            ls_SQL = ls_SQL + "  		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,  " & vbCrLf & _
                              "  		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,  " & vbCrLf & _
                              "  		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,  " & vbCrLf & _
                              "  		b.DeliveryD31, " & vbCrLf & _
                              "  		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old, " & vbCrLf & _
                              "  		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,  " & vbCrLf & _
                              "  		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,  " & vbCrLf & _
                              "  		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,  " & vbCrLf & _
                              "  		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,  " & vbCrLf & _
                              "  		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,  " & vbCrLf & _
                              "  		b.DeliveryD31Old "

            ls_SQL = ls_SQL + " 		from AffiliateRev_Master a " & vbCrLf & _
                              " 	inner join AffiliateRev_Detail b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	inner join MS_CurrCls e on e.CurrCls = b.CurrCls " & vbCrLf & _
                              " 	where a.PONo = '" & ls_SQL & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' " & vbCrLf & _
                              " 	union all " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		'' NoUrut, 'BY SUPPLIER' AffiliateName, '3' NoUrutDesc,  " & vbCrLf & _
                              " 		'' PartNo, b.PartNo PartNo1, '' PartName, ''KanbanCls, '' UnitDesc, 0 MOQ, 0 QtyBox, c.Maker, b.POQty, b.POQtyOld, e.Description CurrDesc, b.Price, b.Amount, " & vbCrLf & _
                              " 		ForecastN1 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,1,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,1,'2015-04-13'))),0),  "

            ls_SQL = ls_SQL + " 		ForecastN2 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,2,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,2,'2015-04-13'))),0),  " & vbCrLf & _
                              " 		ForecastN3 = isnull((select qty from MS_Forecast MF where MF.PartNo = b.PartNo and YEAR(Period) = Year(DATEADD(MONTH,3,'2015-04-13')) and MONTH(Period) = MONTH(DATEADD(MONTH,3,'2015-04-13'))),0),  " & vbCrLf & _
                              "  		b.DeliveryD1, b.DeliveryD2, b.DeliveryD3, b.DeliveryD4, b.DeliveryD5, " & vbCrLf & _
                              "  		b.DeliveryD6, b.DeliveryD7, b.DeliveryD8, b.DeliveryD9, b.DeliveryD10,  " & vbCrLf & _
                              "  		b.DeliveryD11, b.DeliveryD12, b.DeliveryD13, b.DeliveryD14, b.DeliveryD15,  " & vbCrLf & _
                              "  		b.DeliveryD16, b.DeliveryD17, b.DeliveryD18, b.DeliveryD19, b.DeliveryD20,  " & vbCrLf & _
                              "  		b.DeliveryD21, b.DeliveryD22, b.DeliveryD23, b.DeliveryD24, b.DeliveryD25,  " & vbCrLf & _
                              "  		b.DeliveryD26, b.DeliveryD27, b.DeliveryD28, b.DeliveryD29, b.DeliveryD30,  " & vbCrLf & _
                              "  		b.DeliveryD31, " & vbCrLf & _
                              "  		b.DeliveryD1Old, b.DeliveryD2Old, b.DeliveryD3Old, b.DeliveryD4Old, b.DeliveryD5Old, " & vbCrLf & _
                              "  		b.DeliveryD6Old, b.DeliveryD7Old, b.DeliveryD8Old, b.DeliveryD9Old, b.DeliveryD10Old,  "

            ls_SQL = ls_SQL + "  		b.DeliveryD11Old, b.DeliveryD12Old, b.DeliveryD13Old, b.DeliveryD14Old, b.DeliveryD15Old,  " & vbCrLf & _
                              "  		b.DeliveryD16Old, b.DeliveryD17Old, b.DeliveryD18Old, b.DeliveryD19Old, b.DeliveryD20Old,  " & vbCrLf & _
                              "  		b.DeliveryD21Old, b.DeliveryD22Old, b.DeliveryD23Old, b.DeliveryD24Old, b.DeliveryD25Old,  " & vbCrLf & _
                              "  		b.DeliveryD26Old, b.DeliveryD27Old, b.DeliveryD28Old, b.DeliveryD29Old, b.DeliveryD30Old,  " & vbCrLf & _
                              "  		b.DeliveryD31Old " & vbCrLf & _
                              " 		from PORev_MasterUpload a " & vbCrLf & _
                              " 	inner join PORev_DetailUpload b on a.PONo = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.SeqNo = b.SeqNo " & vbCrLf & _
                              " 	inner join MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " 	inner join MS_UnitCls d on c.UnitCls = d.UnitCls " & vbCrLf & _
                              " 	inner join MS_CurrCls e on e.CurrCls = b.CurrCls " & vbCrLf & _
                              " 	where a.PONo = '" & ls_SQL & "' and a.AffiliateID = '" & Session("AffiliateID") & "' and a.SupplierID = '" & Session("SupplierID") & "' "

            ls_SQL = ls_SQL + " )tbl2 on tbl2.AffiliateName = tbl1.AffiliateName and tbl1.PartNo = tbl2.PartNo1 and tbl1.NoUrutDesc = tbl2.NoUrutDesc " & vbCrLf & _
                              " " & pWhereDifference & " "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            'With grid
            '    .DataSource = ds.Tables(0)
            '    .DataBind()
            '    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            'End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub bindPOStatus(Optional ByVal pUpdate As String = "", Optional ByVal pPONO As String = "", Optional ByVal pPORevNo As String = "")
        Dim ls_SQL As String = ""
        Dim ls_PONo As String = ""
        ' Dim ls_PORevNo As String = cboPartNoRev.Text

        If pPONO <> "" Then
            ls_PONo = pPONO
        Else
            'ls_PONo = cboPartNo.Text
        End If

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
                              " from PORev_Master where PONo = '" & ls_PONo & "' and AffiliateID = '" & Session("AffiliateID") & "' and PORevNo = '" & ls_SQL & "'"


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                'If IsDBNull(ds.Tables(0).Rows(0)("EntryDate")) Then
                '    txtDate1.Text = "-"
                '    txtUser1.Text = "-"
                'Else
                '    txtDate1.Text = Format(ds.Tables(0).Rows(0)("EntryDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser1.Text = ds.Tables(0).Rows(0)("EntryUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("AffiliateApproveDate")) Then
                '    txtDate2.Text = "-"
                '    txtUser2.Text = "-"
                'Else
                '    txtDate2.Text = Format(ds.Tables(0).Rows(0)("AffiliateApproveDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser2.Text = ds.Tables(0).Rows(0)("AffiliateApproveUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("PASISendAffiliateDate")) Then
                '    txtDate3.Text = "-"
                '    txtUser3.Text = "-"
                'Else
                '    txtDate3.Text = Format(ds.Tables(0).Rows(0)("PASISendAffiliateDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser3.Text = ds.Tables(0).Rows(0)("PASISendAffiliateUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("SupplierApproveDate")) Then
                '    txtDate4.Text = "-"
                '    txtUser4.Text = "-"
                'Else
                '    txtDate4.Text = Format(ds.Tables(0).Rows(0)("SupplierApproveDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser4.Text = ds.Tables(0).Rows(0)("SupplierApproveUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("SupplierApprovePendingDate")) Then
                '    txtDate5.Text = "-"
                '    txtUser5.Text = "-"
                'Else
                '    txtDate5.Text = Format(ds.Tables(0).Rows(0)("SupplierApprovePendingDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser5.Text = ds.Tables(0).Rows(0)("SupplierApprovePendingUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("SupplierUnApproveDate")) Then
                '    txtDate6.Text = "-"
                '    txtUser6.Text = "-"
                'Else
                '    txtDate6.Text = Format(ds.Tables(0).Rows(0)("SupplierUnApproveDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser6.Text = ds.Tables(0).Rows(0)("SupplierUnApproveUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("PASIApproveDate")) Then
                '    txtDate7.Text = "-"
                '    txtUser7.Text = "-"
                'Else
                '    txtDate7.Text = Format(ds.Tables(0).Rows(0)("PASIApproveDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser7.Text = ds.Tables(0).Rows(0)("PASIApproveUser")
                'End If
                'If IsDBNull(ds.Tables(0).Rows(0)("FinalApproveDate")) Then
                '    txtDate8.Text = ""
                '    txtUser8.Text = ""
                'Else
                '    txtDate8.Text = Format(ds.Tables(0).Rows(0)("FinalApproveDate"), "yyyy-MM-dd HH:mm:ss")
                '    txtUser8.Text = ds.Tables(0).Rows(0)("FinalApproveUser")
                'End If

                'If pUpdate = "update" Then
                '    ButtonApprove.JSProperties("cpDate1") = txtDate1.Text
                '    ButtonApprove.JSProperties("cpDate2") = txtDate2.Text
                '    ButtonApprove.JSProperties("cpDate3") = txtDate3.Text
                '    ButtonApprove.JSProperties("cpDate4") = txtDate4.Text
                '    ButtonApprove.JSProperties("cpDate5") = txtDate5.Text
                '    ButtonApprove.JSProperties("cpDate6") = txtDate6.Text
                '    ButtonApprove.JSProperties("cpDate7") = txtDate7.Text
                '    ButtonApprove.JSProperties("cpDate8") = txtDate8.Text

                '    ButtonApprove.JSProperties("cpUser1") = txtUser1.Text
                '    ButtonApprove.JSProperties("cpUser2") = txtUser2.Text
                '    ButtonApprove.JSProperties("cpUser3") = txtUser3.Text
                '    ButtonApprove.JSProperties("cpUser4") = txtUser4.Text
                '    ButtonApprove.JSProperties("cpUser5") = txtUser5.Text
                '    ButtonApprove.JSProperties("cpUser6") = txtUser6.Text
                '    ButtonApprove.JSProperties("cpUser7") = txtUser7.Text
                '    ButtonApprove.JSProperties("cpUser8") = txtUser8.Text

                '    Call clsMsg.DisplayMessage(lblInfo, "1008", clsMessage.MsgType.InformationMessage)
                '    ButtonApprove.JSProperties("cpMessage") = lblInfo.Text
                'Else
                '    'Session("cpDate8") = txtDate8.Text
                '    'Session("cpUser8") = txtUser8.Text
                '    grid.JSProperties("cpDate1") = txtDate1.Text
                '    grid.JSProperties("cpDate2") = txtDate2.Text
                '    grid.JSProperties("cpDate3") = txtDate3.Text
                '    grid.JSProperties("cpDate4") = txtDate4.Text
                '    grid.JSProperties("cpDate5") = txtDate5.Text
                '    grid.JSProperties("cpDate6") = txtDate6.Text
                '    grid.JSProperties("cpDate7") = txtDate7.Text
                '    grid.JSProperties("cpDate8") = txtDate8.Text

                '    grid.JSProperties("cpUser1") = txtUser1.Text
                '    grid.JSProperties("cpUser2") = txtUser2.Text
                '    grid.JSProperties("cpUser3") = txtUser3.Text
                '    grid.JSProperties("cpUser4") = txtUser4.Text
                '    grid.JSProperties("cpUser5") = txtUser5.Text
                '    grid.JSProperties("cpUser6") = txtUser6.Text
                '    grid.JSProperties("cpUser7") = txtUser7.Text
                '    grid.JSProperties("cpUser8") = txtUser8.Text
                'End If
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub bindHeader(ByVal pPONO As String, ByVal pRevNo As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	case when DeliveryByPASICls = '1' then 'VIA PASI' else 'VIA SUPPLIER' end DeliveryByPASICls, " & vbCrLf & _
                  " 	case when CommercialCls = '1' then 'YES' else 'NO' end CommercialCls, " & vbCrLf & _
                  " 	a.SupplierID, ShipCls, isnull(Remarks,'')Remarks " & vbCrLf & _
                  " from PORev_Master pm left join PO_Master a on pm.PONo = a.PONo and pm.AffiliateID = a.AffiliateID and pm.SupplierID = a.SupplierID" & vbCrLf & _
                  " left join PORev_MasterUpload b on pm.PONo = b.PONo and pm.AffiliateID = b.AffiliateID and pm.PORevNo = b.PORevNo " & vbCrLf & _
                  " where pm.PONo = '" & pPONO & "' and pm.AffiliateID = '" & Session("AffiliateID") & "' and pm.PORevNo = '" & pRevNo & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                'txtDelivery.Text = ds.Tables(0).Rows(0)("DeliveryByPASICls")
                'txtCommercial.Text = ds.Tables(0).Rows(0)("CommercialCls")
                'txtShip.Text = ds.Tables(0).Rows(0)("ShipCls")
                'txtRemarks.Text = ds.Tables(0).Rows(0)("Remarks")
                'Session("SupplierID") = ds.Tables(0).Rows(0)("SupplierID")

                'ButtonPartNo.JSProperties("cpDelivery") = txtDelivery.Text
                'ButtonPartNo.JSProperties("cpCommercial") = txtCommercial.Text
                'ButtonPartNo.JSProperties("cpShip") = txtShip.Text
                'ButtonPartNo.JSProperties("cpRemarks") = txtRemarks.Text
            Else
                'ButtonPartNo.JSProperties("cpDelivery") = ""
                'ButtonPartNo.JSProperties("cpCommercial") = ""
                'ButtonPartNo.JSProperties("cpShip") = ""
                'ButtonPartNo.JSProperties("cpRemarks") = ""
                Session("SupplierID") = ""
            End If

            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' AffiliateName, '' PartNo, '' PartNo1, '' PartName, '' KanbanCls, '' UnitDesc, '' MOQ, '' QtyBox, '' Maker, " & vbCrLf & _
                  " 0 POQty, 0 POQtyOld, '' CurrDesc, '' Price, '' Amount, 0 ForecastN1, 0 ForecastN2, 0 ForecastN3,   " & vbCrLf & _
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

            'With grid
            '    .DataSource = ds.Tables(0)
            '    .DataBind()

            'End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub uf_Approve()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                ls_sql = " Update PORev_Master set FinalApproveDate = getdate(), FinalApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & x & "' and PORevNo = '" & x & "'" & vbCrLf

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

        'ls_SQL = "select RTRIM(PONo) PONo from PORev_Master where AffiliateID = '" & Session("AffiliateID") & "' and Year(Period) = '" & Year(pPeriod) & "' and month(Period) = '" & Month(pPeriod) & "' order by PONo " & vbCrLf
        'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
        '    sqlConn.Open()

        '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
        '    Dim ds As New DataSet
        '    sqlDA.Fill(ds)

        '    With cboPartNo
        '        .Items.Clear()
        '        .Columns.Clear()
        '        .DataSource = ds.Tables(0)
        '        .Columns.Add("PONo")
        '        .Columns(0).Width = 180

        '        .TextField = "PONo"
        '        .DataBind()
        '        .SelectedIndex = -1
        '    End With

        '    sqlConn.Close()
        'End Using
    End Sub

    Private Sub up_FillComboRev(ByVal pPeriod As String, ByVal pPONO As String)
        Dim ls_SQL As String = ""

        ls_SQL = "select RTRIM(PORevNo) PORevNo from PORev_Master where AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & pPONO & "' order by PORevNo " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            'With cboPartNoRev
            '    .Items.Clear()
            '    .Columns.Clear()
            '    .DataSource = ds.Tables(0)
            '    .Columns.Add("PORevNo")
            '    .Columns(0).Width = 180

            '    .TextField = "PORevNo"
            '    .DataBind()
            '    .SelectedIndex = -1
            'End With

            sqlConn.Close()
        End Using
    End Sub

#End Region
    Private Sub up_Import()
        Try
            If Uploader.HasFile Then
                Dim strBuilder As New StringBuilder
                Dim H00 As String = ""
                Dim H10 As String = ""
                Dim H20 As String = ""
                Dim H30 As String = ""

                Dim D10 As String = ""
                Dim D11 As String = ""
                Dim D20 As String = ""

                Try
                    Dim readText() As String = File.ReadAllLines(FilePath)
                    Using sr As StreamReader = New StreamReader(FilePath)
                        Dim tempText As String
                        Dim tempA = readText.Length

                        For i = 1 To tempA
                            tempText = sr.ReadLine
                            If Left(tempText, 1) = "H" Then
                                If Left(tempText, 3) = "H00" Then
                                    H00 = "'" & Mid(tempText, 4, 8) & "','" & Mid(tempText, 12, 8) & "','" & Mid(tempText, 20, 8) & "','" _
                                          & Mid(tempText, 28, 8) & "','" & Mid(tempText, 36, 6) & "','" & Mid(tempText, 42, 15) & "'"
                                End If
                                If Left(tempText, 3) = "H10" Then
                                    H10 = "'" & Mid(tempText, 4, 2) & "','" & Mid(tempText, 6, 1) & "','" & Mid(tempText, 7, 1) & "','" _
                                          & Mid(tempText, 8, 1) & "','" & Mid(tempText, 9, 8) & "'"
                                End If
                                If Left(tempText, 3) = "H30" Then
                                    H30 = "'" & Mid(tempText, 4, 8) & "','" & Mid(tempText, 12, 8) & "','" & Mid(tempText, 20, 8) & "','" _
                                          & Mid(tempText, 28, 8) & "','" & Mid(tempText, 36, 8) & "','" & Mid(tempText, 44, 8) & "','" _
                                          & Mid(tempText, 52, 8) & "','" & Mid(tempText, 60, 8) & "'"
                                End If
                            End If
                            If Left(tempText, 1) = "D" Then
                                If Left(tempText, 3) = "D10" Then
                                    D10 = "'" & Mid(tempText, 4, 25) & "','" & Mid(tempText, 29, 3) & "','" & Mid(tempText, 32, 1) & "','" _
                                          & Mid(tempText, 33, 8) & "','" & Mid(tempText, 41, 8) & "','" & Mid(tempText, 49, 8) & "','" _
                                          & Mid(tempText, 57, 8) & "','" & Mid(tempText, 65, 8) & "'"
                                End If
                                If Left(tempText, 3) = "D11" Then
                                    D11 = "'" & Mid(tempText, 4, 30) & "'"
                                End If
                                If Left(tempText, 3) = "D20" Then
                                    D20 = "'" & Mid(tempText, 4, 9) & "','" & Mid(tempText, 13, 9) & "','" & Mid(tempText, 22, 9) & "','" _
                                          & Mid(tempText, 31, 9) & "','" & Mid(tempText, 40, 9) & "','" & Mid(tempText, 49, 9) & "','" _
                                          & Mid(tempText, 58, 9) & "','" & Mid(tempText, 67, 9) & "'"
                                End If
                            End If
                            If Left(tempText, 3) = "H30" Then
                                Dim ls_sql As String
                                Dim x As Integer

                                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                                    sqlConn.Open()

                                    Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                                        ls_sql = " INSERT INTO OES_HEADER " & vbCrLf & _
                                                 " ([H00_DATA_ID],[H00_DATA_CREATION_AFFILIATE_CODE],[H00_DATA_RECEIPT_AFFILIATE_CODE],[H00_DATA_CREATION_DATE],[H00_DATA_CREATION_TIME],[H00_DATA_DESCRIPTION], " & vbCrLf & _
                                                 " [H10_ORDER_TYPE],[H10_ORDER_REQUEST],[H10_COMM],[H10_FREIGHT],[H10_SUPPLIER], " & vbCrLf & _
                                                 " [H30_FORECAST_DATE_1],[H30_FORECAST_DATE_2],[H30_FORECAST_DATE_3],[H30_FORECAST_DATE_4],[H30_FORECAST_DATE_5],[H30_FORECAST_DATE_6],[H30_FORECAST_DATE_7],[H30_FORECAST_DATE_8]) " & vbCrLf & _
                                                 " VALUES (" & H00 & ", " & H10 & "," & H30 & ")" & vbCrLf

                                        Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        x = SqlComm.ExecuteNonQuery()

                                        SqlComm.Dispose()
                                        sqlTran.Commit()
                                    End Using
                                    sqlConn.Close()
                                End Using
                            End If

                            If Left(tempText, 3) = "D20" Then
                                Dim ls_sql As String
                                Dim x As Integer

                                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                                    sqlConn.Open()

                                    Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                                        ls_sql = " INSERT INTO OES_DETAIL " & vbCrLf & _
                                                 " ([D10_PART_NO],[D10_UOM],[D10_BO_FLAG],[D10_ORDER_QTY_1],[D10_ORDER_QTY_2],[D10_ORDER_QTY_3],[D10_ORDER_QTY_4],[D10_ORDER_QTY_5], " & vbCrLf & _
                                                 " [D11_DESCRIPTION], " & vbCrLf & _
                                                 " [D20_FORECAST_QTY_1],[D20_FORECAST_QTY_2],[D20_FORECAST_QTY_3],[D20_FORECAST_QTY_4],[D20_FORECAST_QTY_5],[D20_FORECAST_QTY_6],[D20_FORECAST_QTY_7],[D20_FORECAST_QTY_8]) " & vbCrLf & _
                                                 " VALUES (" & D10 & ", " & D11 & "," & D20 & ")" & vbCrLf

                                        Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        x = SqlComm.ExecuteNonQuery()

                                        SqlComm.Dispose()
                                        sqlTran.Commit()
                                    End Using
                                    sqlConn.Close()
                                End Using
                            End If
                        Next

                    End Using
                Catch
                    'memo.Text = "Could not read the file"
                End Try
            Else
                If FileName = "" Then
                    Call clsMsg.DisplayMessage(lblInfo, "5016", clsMessage.MsgType.ErrorMessage)
                    Uploader.Enabled = True
                    btnUpload.Enabled = True
                    btnClear.Enabled = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Uploader.Enabled = True
            btnUpload.Enabled = True
            btnClear.Enabled = True
            lblInfo.Text = ex.Message
        End Try
    End Sub

    Protected Sub Uploader_FileUploadComplete(ByVal sender As Object, ByVal e As FileUploadCompleteEventArgs)
        Try
            e.CallbackData = SavePostedFiles(e.UploadedFile)
        Catch ex As Exception
            e.IsValid = False
            lblInfo.Text = ex.Message
        End Try
    End Sub

    Private Function SavePostedFiles(ByVal uploadedFile As UploadedFile) As String
        If (Not uploadedFile.IsValid) Then
            Return String.Empty
        End If

        Ext = Path.Combine(MapPath(""))
        FileName = Uploader.PostedFile.FileName
        FilePath = Ext & "\Import\" & FileName
        uploadedFile.SaveAs(FilePath)

        Return FilePath
    End Function

End Class