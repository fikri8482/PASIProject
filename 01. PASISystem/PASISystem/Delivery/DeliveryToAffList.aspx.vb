Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing

Public Class DeliveryToAffList
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

                grid.JSProperties("cpdtfrom") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdtto") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdeliver") = "ALL"
                grid.JSProperties("cpreceive") = "ALL"
                grid.JSProperties("cpkanban") = "ALL"

            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

#Region "PROCEDURE"
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'AFFILIATE
        ls_sql = "SELECT distinct AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' from MS_Affiliate " & vbCrLf & _
                 "UNION Select AffiliateID = RTRIM(AffiliateID) ,AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "AffiliateID"
                .DataBind()
            End With
            sqlConn.Close()

            'PartNo
            ls_sql = "SELECT distinct PartNo = '" & clsGlobal.gs_All & "', PartName = '" & clsGlobal.gs_All & "' from MS_Parts " & vbCrLf & _
                "Union all SELECT PartNo = RTRIM(PartNo) ,PartName = RTRIM(PartName) FROM MS_Parts " & vbCrLf
            sqlConn.Open()

            Dim sqlDAA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds1 As New DataSet
            sqlDAA.Fill(ds1)

            With cbopart
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds1.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 70
                .Columns.Add("PartName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtpart.Text = clsGlobal.gs_All
                .TextField = "Partno"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If checkbox1.Checked = True Then
                'ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) <= '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
                ls_Filter = ls_Filter + " AND KM.KanbanDate <= '" & Format(dt1.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If
            'Supplier Already Deliver
            If rbdeliver.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(PDD.DOQty, 0) = 0 " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(PDD.DOQty,0) <> 0 " & vbCrLf
            End If

            If rbreceiving.Value = "YES" Then
                ls_Filter = ls_Filter + " AND ISNULL(SDD.DOQty, 0) - ( ISNULL(PRD.GoodRecQty, 0) + ISNULL(PRD.DefectRecQty, 0) ) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_Filter = ls_Filter + " AND ISNULL(SDD.DOQty, 0) - ( ISNULL(PRD.GoodRecQty, 0) + ISNULL(PRD.DefectRecQty, 0) ) = 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_Filter = ls_Filter + " AND (SDM.SuratJalanNo LIKE '%" & Trim(txtsj.Text) & "%' OR PDM.SuratJalanNo LIKE '%" & Trim(txtsj.Text) & "%')" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                'ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(PRM.ReceiveDate,'')),106) between '" & Format(dtfrom.Value, "dd MMM yyyy") & "' and '" & Format(dtto.Value, "dd MMM yyyy") & "'" & vbCrLf
                'ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(PRM.ReceiveDate,'')),106) <> '01 Jan 1900' " & vbCrLf
                ls_Filter = ls_Filter + " AND Convert(char(10),Convert(datetime,PRM.ReceiveDate),120) between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND KM.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_Filter = ls_Filter + "AND KD.PartNo = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If rbkanban.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(KM.KanbanStatus, '') <> '' " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(KM.KanbanStatus,'') = '' " & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_Filter = ls_Filter + "and KD.PONo LIKE '%" & txtpono.Text & "%'" & vbCrLf
            End If

            If rbPasiDelivery.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(PDD.SuratJalanNo, '') <> '' " & vbCrLf
            ElseIf rbPasiDelivery.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(PDD.SuratJalanNo,'') = '' " & vbCrLf
            End If

            'ls_SQL = ls_SQL + ls_Filter

            ls_SQL = "  SELECT  " & vbCrLf & _
                  "  	Act, coldetail = coldetail + '|' + coldetailname, coldetailname , colno =CONVERT(char,ROW_NUMBER() OVER(ORDER BY h_KanbanCls,  HpasiSJ,HSupSJ,h_poorder, h_kanbanorder,h_idxorder , urut)) ,     " & vbCrLf & _
                  "  	colperiod, colaffiliatecode, colaffiliatename, coldeliverylocationcode, coldeliverylocationname, colpono,    " & vbCrLf & _
                  "  	colsuppliercode, colsuppliername, colpokanban, colkanbanno, colplandeldate, coldeldate, colsj,   " & vbCrLf & _
                  "  	colpasideliverydate, colpasisj = ISNULL(colpasisj,''), colpartno, colpartname, coluom, coldeliveryqty, colreceiveqty, coldefect,   " & vbCrLf & _
                  "  	colremaining, colpasideliveryqty, coldeliverydate, coldeliveryby, H_POORDER, H_IDXORDER, H_KANBANORDER, H_AFFILIATEORDER,   " & vbCrLf & _
                  "  	H_KANBANCLS, HSupSJ, HPasiSJ,urut  FROM (  " & vbCrLf & _
                  "  	SELECT DISTINCT *  " & vbCrLf & _
                  "  		 FROM (    " & vbCrLf & _
                  "  		 --DATA DOPASI  " & vbCrLf & _
                  "  		 SELECT DISTINCT    "

            ls_SQL = ls_SQL + "  				  Act = 0 ,    " & vbCrLf & _
                              "  				  coldetail = 'DeliveryToAffEntry.aspx?prm='    " & vbCrLf & _
                              "  				  + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106)    " & vbCrLf & _
                              "  				  + '|' + RTRIM(KM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(PDM.SuratJalanNo,''),'&','DAN')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(KM.DeliveryLocationCode, '')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(MD.DeliveryLocationName, '')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(PDM.DriverName, '')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(PDM.DriverContact, '')) + '|'   " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(PDM.NoPol,'')) + '|'              " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(PDM.JenisArmada, '')) + '|'''    "

            ls_SQL = ls_SQL + "  				  + RTRIM(ISNULL(KD.PONo, '')) + '''|'''      " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(KD.KanbanNo, '')) + '''|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(KM.SupplierID, '')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(MS.SupplierName, '')) + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(PDD.SuratJalanNo,''),'&','DAN'))  + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(PDD.SuratJalanNo,''),'&','DAN'))  + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(SDM.SuratJalanNo,''),'&','DAN')) + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(SDM.SuratJalanNo,''),'&','DAN')) ,    " & vbCrLf & _
                              "  				  coldetailname = CASE WHEN ISNULL(PDM.SuratJalanNo, '') = ''    " & vbCrLf & _
                              "  									   THEN 'DELIVERY'    " & vbCrLf & _
                              "  									   ELSE 'DETAIL'    "

            ls_SQL = ls_SQL + "  								  END ,            colno = '' ,    " & vbCrLf & _
                              "  				  colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,KM.KanbanDate),106), 8),    " & vbCrLf & _
                              "  				  colaffiliatecode = KM.AffiliateID ,    " & vbCrLf & _
                              "  				  colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                              "  				  coldeliverylocationcode = KM.DeliveryLocationCode ,    " & vbCrLf & _
                              "  				  coldeliverylocationname = MD.DeliveryLocationName ,    " & vbCrLf & _
                              "  				  colpono = KD.PONo ,    " & vbCrLf & _
                              "  				  colsuppliercode = KM.SupplierID ,    " & vbCrLf & _
                              "  				  colsuppliername = MS.SupplierName ,    " & vbCrLf & _
                              "  				  colpokanban = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN 'NO'    " & vbCrLf & _
                              "  									 ELSE 'YES'                          END ,    "

            ls_SQL = ls_SQL + "  				  colkanbanno = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN '-'    " & vbCrLf & _
                              "  									 ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf & _
                              "  								END ,    " & vbCrLf & _
                              "  				  colplandeldate = CASE WHEN ISNULL(KM.KanbanDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(KM.KanbanDate, '')), 106) END ,    " & vbCrLf & _
                              "  		              " & vbCrLf & _
                              "  				  coldeldate = CASE WHEN ISNULL(SDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(SDM.DeliveryDate, '')), 106) END ,    " & vbCrLf & _
                              "  		             " & vbCrLf & _
                              "  				  colsj = ISNULL(SDM.SuratJalanNo, '') ,    " & vbCrLf & _
                              "  				  colpasideliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,    " & vbCrLf & _
                              "  				  colpasisj = PDM.SuratJalanNo ,              " & vbCrLf & _
                              "  				  colpartno = '' ,    "

            ls_SQL = ls_SQL + "  				  colpartname = '' ,    " & vbCrLf & _
                              "  				  coluom = '' ,    " & vbCrLf & _
                              "  				  coldeliveryqty = '' ,    " & vbCrLf & _
                              "  				  colreceiveqty = '' ,    " & vbCrLf & _
                              "  				  coldefect = '' ,    " & vbCrLf & _
                              "  				  colremaining = '' ,    " & vbCrLf & _
                              "  				  colpasideliveryqty = '' ,    " & vbCrLf & _
                              "  				  coldeliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END,    " & vbCrLf & _
                              "  				  coldeliveryby = ISNULL(PDM.EntryUser, '') ,    " & vbCrLf & _
                              "  				  KD.PONo H_POORDER ,              " & vbCrLf & _
                              "  				  H_IDXORDER = 0 ,    "

            ls_SQL = ls_SQL + "  				  H_KANBANORDER = ISNULL(KD.KanbanNo, '-') ,    " & vbCrLf & _
                              "  				  H_AFFILIATEORDER = KM.AffiliateID ,    " & vbCrLf & _
                              "  				  H_KANBANCLS = KM.KanbanStatus,    " & vbCrLf & _
                              "  				  HSupSJ = SDM.SuratJalanNo,   " & vbCrLf & _
                              "  				  HPasiSJ = isnull(PDD.SuratJalanNo,''), urut = 0 " & vbCrLf & _
                              "  		  FROM    dbo.Kanban_Detail KD " & vbCrLf & _
                              "  				  LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "  													AND KD.KanbanNo = KM.KanbanNo    " & vbCrLf & _
                              "  													AND KD.SupplierID = KM.SupplierID    " & vbCrLf & _
                              "  													AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf & _
                              "  														 AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf & _
                              "  														 AND KD.PONo = SDD.PONo    " & vbCrLf & _
                              "  														 AND KD.SupplierID = SDD.SupplierID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  "

            ls_SQL = ls_SQL + "                                                           AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf & _
                              "  														 AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf & _
                              "  				  INNER JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
                              "  														  AND KD.KanbanNo = PRD.KanbanNo    " & vbCrLf & _
                              "  														  AND KD.SupplierID = PRD.SupplierID    " & vbCrLf & _
                              "  														  AND KD.PartNo = PRD.PartNo    " & vbCrLf & _
                              "  														  AND KD.PONo = PRD.PONo    " & vbCrLf & _
                              "  														  AND SDM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
                              "  				  LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
                              "  														  AND PRM.SuratJalanNo = PRD.SuratJalanNo    " & vbCrLf & _
                              "  														  AND PRM.SupplierID = PRD.SupplierID    "

            ls_SQL = ls_SQL + "  				  LEFT JOIN dbo.DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
                              "  													 AND PRD.KanbanNo = PDD.KanbanNo    " & vbCrLf & _
                              "  													 AND PRD.SupplierID = PDD.SupplierID    " & vbCrLf & _
                              "  													 AND PRD.PartNo = PDD.PartNo    " & vbCrLf & _
                              "  													 AND PRD.PoNo = PDD.PoNo    " & vbCrLf & _
                              "  													 AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier   " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf & _
                              "  													 AND PDD.SuratJalanNo = PDM.SuratJalanNo    " & vbCrLf & _
                              "                                                      AND PDD.SupplierID = PDM.SupplierID " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = KM.SupplierID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf & _
                              "  				                                     AND MD.AffiliateID = KM.AffiliateID      " & vbCrLf & _
                              "  		WHERE   'A' = 'A'  " & vbCrLf & _
                              "  		AND ISNULL(PDD.SuratJalanNo,'') <> ''  " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + "  		--AND PRD.PartNo = '7154-0889-30'  " & vbCrLf & _
                              "  		--AND SDM.SuratJalanNo = '370/IKS-SJ/VII/2015'  " & vbCrLf & _
                              "  		--AND PRD.KanbanNo = '20150707-1'  " & vbCrLf & _
                              "  		--AND PRD.PONO = 'PC1507-IKS'  " & vbCrLf & _
                              "  		--QTY <> 0  " & vbCrLf & _
                              "  		UNION ALL  " & vbCrLf & _
                              "  		SELECT DISTINCT    " & vbCrLf & _
                              "  				  Act = 0 ,  "

            ls_SQL = ls_SQL + "  				  coldetail = 'DeliveryToAffEntry.aspx?prm='  " & vbCrLf & _
                              "  				  + ''    " & vbCrLf & _
                              "  				  + '|' + RTRIM(KM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'    " & vbCrLf & _
                              "  				  + '' + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(KM.DeliveryLocationCode, '')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(MD.DeliveryLocationName, '')) + '|'    " & vbCrLf & _
                              "  				  + '' + '|'    " & vbCrLf & _
                              "  				  + '' + '|'   " & vbCrLf & _
                              "  				  + '' + '|'  " & vbCrLf & _
                              "  				  + '' + '|'''    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(KD.PONo, '')) + '''|'''      "

            ls_SQL = ls_SQL + "  				  + RTRIM(ISNULL(KD.KanbanNo, '')) + '''|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(KM.SupplierID, '')) + '|'    " & vbCrLf & _
                              "  				  + RTRIM(ISNULL(MS.SupplierName, '')) + '|'  " & vbCrLf & _
                              "  				  + ''  + '|'    " & vbCrLf & _
                              "  				  + ''  + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(SDM.SuratJalanNo,''),'&','DAN')) + '|'    " & vbCrLf & _
                              "  				  + Rtrim(REPLACE(ISNULL(SDM.SuratJalanNo,''),'&','DAN')) ,    " & vbCrLf & _
                              "  				  coldetailname = 'DELIVERY' ,             " & vbCrLf & _
                              "  				  colno = '' ,    " & vbCrLf & _
                              "  				  colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,KM.KanbanDate),106), 8),    " & vbCrLf & _
                              "  				  colaffiliatecode = KM.AffiliateID ,    "

            ls_SQL = ls_SQL + "  				  colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
                              "  				  coldeliverylocationcode = KM.DeliveryLocationCode ,    " & vbCrLf & _
                              "  				  coldeliverylocationname = MD.DeliveryLocationName ,    " & vbCrLf & _
                              "  				  colpono = KD.PONo ,    " & vbCrLf & _
                              "  				  colsuppliercode = KM.SupplierID ,    " & vbCrLf & _
                              "  				  colsuppliername = MS.SupplierName ,    " & vbCrLf & _
                              "  				  colpokanban = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN 'NO'    " & vbCrLf & _
                              "  									 ELSE 'YES'                          END ,    " & vbCrLf & _
                              "  				  colkanbanno = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN '-'    " & vbCrLf & _
                              "  									 ELSE ISNULL(KD.KanbanNo, '')    " & vbCrLf & _
                              "  								END ,    "

            ls_SQL = ls_SQL + "  				  colplandeldate = CASE WHEN ISNULL(KM.KanbanDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(KM.KanbanDate, '')), 106) END ,    " & vbCrLf & _
                              "  		              " & vbCrLf & _
                              "  				  coldeldate = CASE WHEN ISNULL(SDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(SDM.DeliveryDate, '')), 106) END ,    " & vbCrLf & _
                              "  		             " & vbCrLf & _
                              "  				  colsj = ISNULL(SDM.SuratJalanNo, '') ,    " & vbCrLf & _
                              "  				  colpasideliverydate = '',    " & vbCrLf & _
                              "  				  colpasisj = '' ,              " & vbCrLf & _
                              "  				  colpartno = '' ,    " & vbCrLf & _
                              "  				  colpartname = '' ,    " & vbCrLf & _
                              "  				  coluom = '' ,    " & vbCrLf & _
                              "  				  coldeliveryqty = '' ,    "

            ls_SQL = ls_SQL + "  				  colreceiveqty = '' ,    " & vbCrLf & _
                              "  				  coldefect = '' ,    " & vbCrLf & _
                              "  				  colremaining = '' ,    " & vbCrLf & _
                              "  				  colpasideliveryqty = '' ,    " & vbCrLf & _
                              "  				  coldeliverydate = '', --CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END,    " & vbCrLf & _
                              "  				  coldeliveryby = '', --ISNULL(PDM.EntryUser, '') ,    " & vbCrLf & _
                              "  				  KD.PONo H_POORDER ,              " & vbCrLf & _
                              "  				  H_IDXORDER = 2 ,    " & vbCrLf & _
                              "  				  H_KANBANORDER = ISNULL(KD.KanbanNo, '-') ,    " & vbCrLf & _
                              "  				  H_AFFILIATEORDER = KM.AffiliateID ,    " & vbCrLf & _
                              "  				  H_KANBANCLS = KM.KanbanStatus,    "

            ls_SQL = ls_SQL + "  				  HSupSJ = SDM.SuratJalanNo,   " & vbCrLf & _
                              "  				  HPasiSJ = '', urut = 1 " & vbCrLf & _
                              "  		  FROM    dbo.Kanban_Detail KD LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf

            ls_SQL = ls_SQL + "  													AND KD.KanbanNo = KM.KanbanNo    " & vbCrLf & _
                              "  													AND KD.SupplierID = KM.SupplierID    " & vbCrLf & _
                              "  													AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf & _
                              "  														 AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf & _
                              "  														 AND KD.PONo = SDD.PONo    " & vbCrLf & _
                              "  														 AND KD.SupplierID = SDD.SupplierID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID                                                   AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf & _
                              "  														 AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf & _
                              "  				  INNER JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
                              "  														  AND KD.KanbanNo = PRD.KanbanNo    "

            ls_SQL = ls_SQL + "  														  AND KD.SupplierID = PRD.SupplierID    " & vbCrLf & _
                              "  														  AND KD.PartNo = PRD.PartNo    " & vbCrLf & _
                              "  														  AND KD.PONo = PRD.PONo    " & vbCrLf & _
                              "  														  AND SDM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
                              "  				  LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
                              "  														  AND PRM.SuratJalanNo = PRD.SuratJalanNo    " & vbCrLf & _
                              "  														  AND PRM.SupplierID = PRD.SupplierID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID                                                 " & vbCrLf & _
                              "  													 AND PRD.KanbanNo = PDD.KanbanNo    " & vbCrLf & _
                              "  													 AND PRD.SupplierID = PDD.SupplierID    " & vbCrLf & _
                              "  													 AND PRD.PartNo = PDD.PartNo    "

            ls_SQL = ls_SQL + "  													 AND PRD.PoNo = PDD.PoNo    " & vbCrLf & _
                              "  													 AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier   " & vbCrLf & _
                              "  				  LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf & _
                              "  													 AND PDD.SuratJalanNo = PDM.SuratJalanNo  " & vbCrLf & _
                              "                                                      AND PDD.SupplierID = PDM.SupplierID " & vbCrLf & _
                              "  				  LEFT JOIN (SELECT AffiliateID,KanbanNo,SupplierID,PartNo,PONo, DOQty = SUM(ISNULL(DOQty,0)),SuratJalanNoSupplier   " & vbCrLf & _
                              "                                     FROM DOPasi_Detail GROUP BY AffiliateID,KanbanNo,SupplierID,PartNo,PONo,SuratJalanNoSupplier) REM  " & vbCrLf & _
                              "  						ON KD.AffiliateID = REM.AffiliateID   " & vbCrLf & _
                              "  						AND KD.KanbanNo = REM.KanbanNo   " & vbCrLf & _
                              "  						AND KD.SupplierID = REM.SupplierID   " & vbCrLf & _
                              "  						AND KD.PartNo = REM.PartNo   " & vbCrLf & _
                              "  						AND KD.PoNo = REM.PoNo   " & vbCrLf & _
                              "                         AND SDM.SuratJalanNo = REM.SuratJalanNoSupplier " & vbCrLf

            ls_SQL = ls_SQL + "  				  LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = KM.SupplierID    " & vbCrLf & _
                              "  				  LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode      " & vbCrLf & _
                              "  				                                     AND MD.AffiliateID = KM.AffiliateID      " & vbCrLf & _
                              "  		WHERE   'A' = 'A'  " & vbCrLf & _
                              "  		AND (CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(SDD.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END)) > 0  " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + "  		--AND ISNULL(PDD.SuratJalanNo, '') = ''  " & vbCrLf & _
                              "  		--AND PRD.PartNo = '7154-0889-30'  " & vbCrLf & _
                              "  		--AND SDM.SuratJalanNo = '370/IKS-SJ/VII/2015'  		 " & vbCrLf & _
                              "  		--AND PRD.KanbanNo = '20150707-1'  "

            ls_SQL = ls_SQL + "  		--AND PRD.PONO = 'PC1507-IKS'  " & vbCrLf & _
                              "  		 )x  " & vbCrLf & _
                              "  	)HEADER   " & vbCrLf
            '                  "   UNION ALL   " & vbCrLf & _
            '                  "   SELECT DISTINCT   " & vbCrLf & _
            '                  "           Act = 0 ,   " & vbCrLf & _
            '                  "           coldetail = '' ,   " & vbCrLf & _
            '                  "           coldetailname = '' ,   " & vbCrLf & _
            '                  "           colno = '' ,   " & vbCrLf & _
            '                  "           colperiod = '' ,   " & vbCrLf & _
            '                  "           colaffiliatecode = '' ,   "

            'ls_SQL = ls_SQL + "           colaffiliatename = '' ,   " & vbCrLf & _
            '                  "           coldeliverylocationcode = '' ,   " & vbCrLf & _
            '                  "           coldeliverylocationname = '' ,   " & vbCrLf & _
            '                  "           colpono = '' ,   " & vbCrLf & _
            '                  "           colsuppliercode = '' ,   " & vbCrLf & _
            '                  "           colsuppliername = '' ,   " & vbCrLf & _
            '                  "           colpokanban = '' ,   " & vbCrLf & _
            '                  "           colkanbanno = '' ,   " & vbCrLf & _
            '                  "           colplandeldate = '' ,   " & vbCrLf & _
            '                  "           coldeldate = '' ,   " & vbCrLf & _
            '                  "           colsj = '' ,   "

            'ls_SQL = ls_SQL + "           colpasideliverydate = '' ,   " & vbCrLf & _
            '                  "           colpasisj = '' ,   " & vbCrLf & _
            '                  "           colpartno = KD.PartNo ,   " & vbCrLf & _
            '                  "           colpartname = MP.PartName ,   " & vbCrLf & _
            '                  "           coluom = UC.Description ,   " & vbCrLf & _
            '                  "           coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))),   " & vbCrLf & _
            '                  "           colreceiveqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),   " & vbCrLf & _
            '                  "           coldefect = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),    " & vbCrLf & _
            '                  "           --colremaining = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(SDD.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END))),   " & vbCrLf & _
            '                  "           colremaining = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),(ISNULL(SDD.DOQty,0) - (ISNULL(PRD.GoodRecQty,0))) ))), " & vbCrLf & _
            '                  "           colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),    " & vbCrLf & _
            '                  "           coldeliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,   "

            'ls_SQL = ls_SQL + "           coldeliveryby = PDM.PIC ,   " & vbCrLf & _
            '                  "           POOrder = KD.PONo ,   " & vbCrLf & _
            '                  "           idxorder = 1 ,   " & vbCrLf & _
            '                  "           kanbanorder = ISNULL(KD.KanbanNo, '-') ,   " & vbCrLf & _
            '                  "           affiliateorder = KM.AffiliateID ,   " & vbCrLf & _
            '                  "           KM.KanbanStatus,   " & vbCrLf & _
            '                  "           HSupSJ = SDM.SuratJalanNo,   " & vbCrLf & _
            '                  "           HPasiSJ = isnull(PDM.SuratJalanNo,''), urut = 0 " & vbCrLf & _
            '                  "   FROM     " & vbCrLf & _
            '                  "           dbo.Kanban_Detail KD " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND KD.KanbanNo = SDD.KanbanNo   "

            'ls_SQL = ls_SQL + "                                                  AND KD.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "                                                  AND KD.PONo = SDD.PONo   " & vbCrLf & _
            '                  "                                                  AND KD.PartNo = SDD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND SDD.KanbanNo = PRD.KanbanNo   " & vbCrLf & _
            '                  "                                                   AND SDD.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "                                                   AND SDD.PartNo = PRD.PartNo   " & vbCrLf & _
            '                  "                                                   AND SDD.PONo = PRD.PONo   "

            'ls_SQL = ls_SQL + "                                                   AND SDM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                   AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "                                              AND KD.KanbanNo = PDD.KanbanNo   " & vbCrLf & _
            '                  "                                              AND KD.SupplierID = PDD.SupplierID   " & vbCrLf & _
            '                  "                                              AND KD.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "                                              AND KD.PoNo = PDD.PoNo   " & vbCrLf & _
            '                  "                                              AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID   "

            'ls_SQL = ls_SQL + "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SupplierID = PDM.SupplierID  " & vbCrLf & _
            '                  "  		  LEFT JOIN (SELECT AffiliateID,KanbanNo,SupplierID,PartNo,PONo,SuratJalanNoSupplier, DOQty = SUM(ISNULL(DOQty,0))   " & vbCrLf & _
            '                  "                         FROM DOPasi_Detail GROUP BY AffiliateID,KanbanNo,SupplierID,PartNo,PONo,SuratJalanNoSupplier) REM  " & vbCrLf & _
            '                  "  					ON KD.AffiliateID = REM.AffiliateID  " & vbCrLf & _
            '                  "                      AND KD.KanbanNo = REM.KanbanNo   " & vbCrLf & _
            '                  "                      AND KD.SupplierID = REM.SupplierID   " & vbCrLf & _
            '                  "                      AND KD.PartNo = REM.PartNo   " & vbCrLf & _
            '                  "                      AND KD.PoNo = REM.PoNo   " & vbCrLf & _
            '                  "                      AND SDM.SuratJalanNo = REM.SuratJalanNoSupplier  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo   "

            'ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "    WHERE   'A' = 'A'   " & vbCrLf & _
            '                  "    AND ISNULL(PDD.SuratJalanNo,'') <> ''  " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter

            'ls_SQL = ls_SQL + "    --AND CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN ISNULL(SDD.DOQty,0) ELSE ISNULL(SDD.DOQty-PRD.GoodRecQty,0) END))) <> 0  " & vbCrLf & _
            '                  "    --AND PRD.PartNo = '7154-0889-30'  " & vbCrLf & _
            '                  "    --AND SDM.SuratJalanNo = '370/IKS-SJ/VII/2015'  " & vbCrLf & _
            '                  "    --AND PRD.KanbanNo = '20150707-1'     " & vbCrLf & _
            '                  "    --AND PRD.PONO = 'PC1507-IKS'  "

            'ls_SQL = ls_SQL + "   --QTY <> 0  " & vbCrLf & _
            '                  "   UNION ALL  " & vbCrLf & _
            '                  "   SELECT DISTINCT   " & vbCrLf & _
            '                  "           Act = 1 ,   " & vbCrLf & _
            '                  "           coldetail = '' ,   " & vbCrLf & _
            '                  "           coldetailname = '' ,   " & vbCrLf & _
            '                  "           colno = '' ,   " & vbCrLf & _
            '                  "           colperiod = '' ,   " & vbCrLf & _
            '                  "           colaffiliatecode = '' ,   " & vbCrLf & _
            '                  "           colaffiliatename = '' ,   " & vbCrLf & _
            '                  "           coldeliverylocationcode = '' ,   "

            'ls_SQL = ls_SQL + "           coldeliverylocationname = '' ,   " & vbCrLf & _
            '                  "           colpono = '' ,   " & vbCrLf & _
            '                  "           colsuppliercode = '' ,   " & vbCrLf & _
            '                  "           colsuppliername = '' ,   " & vbCrLf & _
            '                  "           colpokanban = '' ,   " & vbCrLf & _
            '                  "           colkanbanno = '' ,   " & vbCrLf & _
            '                  "           colplandeldate = '' ,   " & vbCrLf & _
            '                  "           coldeldate = '' ,   " & vbCrLf & _
            '                  "           colsj = '' ,   " & vbCrLf & _
            '                  "           colpasideliverydate = '' ,   " & vbCrLf & _
            '                  "           colpasisj = '' ,   "

            'ls_SQL = ls_SQL + "           colpartno = KD.PartNo ,   " & vbCrLf & _
            '                  "           colpartname = MP.PartName ,   " & vbCrLf & _
            '                  "           coluom = UC.Description ,   " & vbCrLf & _
            '                  "           coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))) ,   " & vbCrLf & _
            '                  "           colreceiveqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))),   " & vbCrLf & _
            '                  "           coldefect = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),    " & vbCrLf & _
            '                  "           --colremaining = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(SDD.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END))),   " & vbCrLf & _
            '                  "           colremaining = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),(ISNULL(SDD.DOQty,0) - (ISNULL(PRD.GoodRecQty,0))) ))), " & vbCrLf & _
            '                  "           colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(REM.DOQty,0)))),    " & vbCrLf & _
            '                  "           coldeliverydate = '' ,   " & vbCrLf & _
            '                  "           coldeliveryby = '',   " & vbCrLf & _
            '                  "           POOrder = KD.PONo ,   "

            'ls_SQL = ls_SQL + "           idxorder = 3 ,   " & vbCrLf & _
            '                  "           kanbanorder = ISNULL(KD.KanbanNo, '-') ,   " & vbCrLf & _
            '                  "           affiliateorder = KM.AffiliateID ,   " & vbCrLf & _
            '                  "           KM.KanbanStatus,   " & vbCrLf & _
            '                  "           HSupSJ = SDM.SuratJalanNo,   " & vbCrLf & _
            '                  "           HPasiSJ = '', uurt = 1 " & vbCrLf & _
            '                  "   FROM    dbo.Kanban_Detail KD  " & vbCrLf

            'ls_SQL = ls_SQL + "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND KD.KanbanNo = SDD.KanbanNo   " & vbCrLf & _
            '                  "                                                  AND KD.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "                                                  AND KD.PONo = SDD.PONo   "

            'ls_SQL = ls_SQL + "                                                  AND KD.PartNo = SDD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND SDD.KanbanNo = PRD.KanbanNo   " & vbCrLf & _
            '                  "                                                   AND SDD.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "                                                   AND SDD.PartNo = PRD.PartNo   " & vbCrLf & _
            '                  "                                                   AND SDD.PONo = PRD.PONo   " & vbCrLf & _
            '                  "                                                   AND SDM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   "

            'ls_SQL = ls_SQL + "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                   AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "                                              AND KD.KanbanNo = PDD.KanbanNo   " & vbCrLf & _
            '                  "                                              AND KD.SupplierID = PDD.SupplierID   " & vbCrLf & _
            '                  "                                              AND KD.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "                                              AND KD.PoNo = PDD.PoNo   " & vbCrLf & _
            '                  "                                              AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID   " & vbCrLf & _
            '                  "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf & _
            '                  "                                              AND PDD.SupplierID = PDM.SupplierID  "

            'ls_SQL = ls_SQL + "  		  LEFT JOIN (SELECT AffiliateID,KanbanNo,SupplierID,PartNo,PONo, DOQty = SUM(ISNULL(DOQty,0)),SuratJalanNoSupplier   " & vbCrLf & _
            '                  "                         FROM DOPasi_Detail GROUP BY AffiliateID,KanbanNo,SupplierID,PartNo,PONo,SuratJalanNoSupplier) REM  " & vbCrLf & _
            '                  "  					ON KD.AffiliateID = REM.AffiliateID  " & vbCrLf & _
            '                  "                      AND KD.KanbanNo = REM.KanbanNo   " & vbCrLf & _
            '                  "                      AND KD.SupplierID = REM.SupplierID   " & vbCrLf & _
            '                  "                      AND KD.PartNo = REM.PartNo   " & vbCrLf & _
            '                  "                      AND KD.PoNo = REM.PoNo   " & vbCrLf & _
            '                  "                      AND SDM.SuratJalanNo = REM.SuratJalanNoSupplier " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = KD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID   "

            'ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "    WHERE   'A' = 'A'   " & vbCrLf & _
            '                  "    --AND PRD.PartNo = '7254-0926-3W'  " & vbCrLf & _
            '                  "    --AND SDM.SuratJalanNo = '370/IKS-SJ/VII/2015'  " & vbCrLf & _
            '                  "    --AND PRD.KanbanNo = '20150707-1'  " & vbCrLf & _
            '                  "    --AND PRD.PONO = 'PC1507-IKS'  " & vbCrLf & _
            '                  "    AND (CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(SDD.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END)) > 0  " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + "   ORDER BY h_KanbanCls,  HpasiSJ,HSupSJ,h_poorder, h_kanbanorder,h_idxorder , urut"



            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call ColorGrid()
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoad_OLD()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If checkbox1.Checked = True Then
                ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(KM.KanbanDate,'')),106) <= '" & Format(dt1.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If
            'Supplier Already Deliver
            If rbdeliver.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(PDD.DOQty, 0) = 0 " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(PDD.DOQty,0) <> 0 " & vbCrLf
            End If

            If rbreceiving.Value = "YES" Then
                ls_Filter = ls_Filter + " AND ISNULL(SDD.DOQty, 0) - ( ISNULL(PRD.GoodRecQty, 0) + ISNULL(PRD.DefectRecQty, 0) ) > 0 " & vbCrLf
            ElseIf rbreceiving.Value = "NO" Then
                ls_Filter = ls_Filter + " AND ISNULL(SDD.DOQty, 0) - ( ISNULL(PRD.GoodRecQty, 0) + ISNULL(PRD.DefectRecQty, 0) ) = 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_Filter = ls_Filter + " AND PDM.SuratJalanNo LIKE '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(PRM.ReceiveDate,'')),106) between '" & Format(dtfrom.Value, "dd MMM yyyy") & "' and '" & Format(dtto.Value, "dd MMM yyyy") & "'" & vbCrLf
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND POM.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_Filter = ls_Filter + "AND pod.PartNo = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If rbkanban.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(POD.KanbanCls, '') <> '' " & vbCrLf
            ElseIf rbkanban.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(POD.KanbanCls,'') = '' " & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_Filter = ls_Filter + "and POM.PONo LIKE '%" & txtpono.Text & "%'" & vbCrLf
            End If

            If rbPasiDelivery.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(PDD.SuratJalanNo, '') <> '' " & vbCrLf
            ElseIf rbPasiDelivery.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(PDD.SuratJalanNo,'') = '' " & vbCrLf
            End If

            ls_SQL = "  SELECT " & vbCrLf & _
                     " Act, coldetail , coldetailname , colno = CONVERT(char,ROW_NUMBER() OVER(ORDER BY h_poorder, h_KanbanCls,h_kanbanorder, HSupSJ, h_idxorder DESC)) ,   " & vbCrLf & _
                     " colperiod, colaffiliatecode, colaffiliatename, coldeliverylocationcode, coldeliverylocationname, colpono,  " & vbCrLf & _
                     " colsuppliercode, colsuppliername, colpokanban, colkanbanno, colplandeldate, coldeldate, colsj, " & vbCrLf & _
                     " colpasideliverydate, colpasisj = ISNULL(colpasisj,''), colpartno, colpartname, coluom, coldeliveryqty, colreceiveqty, coldefect, " & vbCrLf & _
                     " colremaining, colpasideliveryqty, coldeliverydate, coldeliveryby, H_POORDER, H_IDXORDER, H_KANBANORDER, H_AFFILIATEORDER, " & vbCrLf & _
                     " H_KANBANCLS, HSupSJ, HPasiSJ  " & vbCrLf & _
                     " FROM (  " & vbCrLf
            ls_SQL = ls_SQL + " SELECT DISTINCT  " & vbCrLf & _
                  "          Act = 0 ,  " & vbCrLf & _
                  "          coldetail = 'DeliveryToAffEntry.aspx?prm='  " & vbCrLf & _
                  "          + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106)  " & vbCrLf & _
                  "          + '|' + RTRIM(POM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'  " & vbCrLf & _
                  "          + RTRIM(ISNULL(PDM.SuratJalanNo, '')) + '|'  " & vbCrLf & _
                  "          + RTRIM(ISNULL(KM.DeliveryLocationCode, '')) + '|'  " & vbCrLf & _
                  "          + RTRIM(ISNULL(MD.DeliveryLocationName, '')) + '|'  " & vbCrLf & _
                  "          + RTRIM(ISNULL(PDM.DriverName, '')) + '|'  " & vbCrLf & _
                  "          + RTRIM(ISNULL(PDM.DriverContact, '')) + '|' + RTRIM(ISNULL(PDM.NoPol,  " & vbCrLf & _
                  "                                                                '')) + '|'  "

            ls_SQL = ls_SQL + "          + RTRIM(ISNULL(PDM.JenisArmada, '')) + '|'''  " & vbCrLf & _
                              "          + RTRIM(ISNULL(POM.PONo, '')) + '''|'''    " & vbCrLf & _
                              "          + RTRIM(ISNULL(KD.KanbanNo, '')) + '''|'  " & vbCrLf & _
                              "          + RTRIM(ISNULL(POM.SupplierID, '')) + '|'  " & vbCrLf & _
                              "          + RTRIM(ISNULL(MS.SupplierName, '')) + '|'  " & vbCrLf & _
                              "          + RTRIM(ISNULL(PDD.SuratJalanNo, ''))  + '|'  " & vbCrLf & _
                              "          + RTRIM(ISNULL(PDD.SuratJalanNo, ''))  + '|'  " & vbCrLf & _
                              "          + RTRIM(ISNULL(SDM.SuratJalanNo, '')) + '|'  " & vbCrLf & _
                              "          + RTRIM(ISNULL(SDM.SuratJalanNo, '')) ,  " & vbCrLf & _
                              "          coldetailname = CASE WHEN ISNULL(PDM.SuratJalanNo, '') = ''  " & vbCrLf & _
                              "                               THEN 'DELIVERY'  " & vbCrLf & _
                              "                               ELSE 'DETAIL'  " & vbCrLf & _
                              "                          END ,  "

            ls_SQL = ls_SQL + "          colno = '' ,  " & vbCrLf & _
                              "          colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,POM.Period),106), 8),  " & vbCrLf & _
                              "          colaffiliatecode = POM.AffiliateID ,  " & vbCrLf & _
                              "          colaffiliatename = MA.AffiliateName ,  " & vbCrLf & _
                              "          coldeliverylocationcode = KM.DeliveryLocationCode ,  " & vbCrLf & _
                              "          coldeliverylocationname = MD.DeliveryLocationName ,  " & vbCrLf & _
                              "          colpono = POM.PONo ,  " & vbCrLf & _
                              "          colsuppliercode = POM.SupplierID ,  " & vbCrLf & _
                              "          colsuppliername = MS.SupplierName ,  " & vbCrLf & _
                              "          colpokanban = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN 'NO'  " & vbCrLf & _
                              "                             ELSE 'YES'  "

            ls_SQL = ls_SQL + "                        END ,  " & vbCrLf & _
                              "          colkanbanno = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN '-'  " & vbCrLf & _
                              "                             ELSE ISNULL(KD.KanbanNo, '')  " & vbCrLf & _
                              "                        END ,  " & vbCrLf & _
                              "          colplandeldate = CASE WHEN ISNULL(KM.KanbanDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(KM.KanbanDate, '')), 106) END ,  " & vbCrLf & _
                              "            " & vbCrLf & _
                              "          coldeldate = CASE WHEN ISNULL(SDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(SDM.DeliveryDate, '')), 106) END ,  " & vbCrLf & _
                              "           " & vbCrLf & _
                              "          colsj = ISNULL(SDM.SuratJalanNo, '') ,  " & vbCrLf & _
                              "          colpasideliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,  " & vbCrLf & _
                              "          colpasisj = PDM.SuratJalanNo ,  "

            ls_SQL = ls_SQL + "          colpartno = '' ,  " & vbCrLf & _
                              "          colpartname = '' ,  " & vbCrLf & _
                              "          coluom = '' ,  " & vbCrLf & _
                              "          coldeliveryqty = '' ,  " & vbCrLf & _
                              "          colreceiveqty = '' ,  " & vbCrLf & _
                              "          coldefect = '' ,  " & vbCrLf & _
                              "          colremaining = '' ,  " & vbCrLf & _
                              "          colpasideliveryqty = '' ,  " & vbCrLf & _
                              "          coldeliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END,  " & vbCrLf & _
                              "          coldeliveryby = ISNULL(PDM.EntryUser, '') ,  " & vbCrLf & _
                              "          pom.PONo H_POORDER ,  "

            ls_SQL = ls_SQL + "          H_IDXORDER = 0 ,  " & vbCrLf & _
                              "          H_KANBANORDER = ISNULL(KD.KanbanNo, '-') ,  " & vbCrLf & _
                              "          H_AFFILIATEORDER = POM.AffiliateID ,  " & vbCrLf & _
                              "          H_KANBANCLS = pod.KanbanCls,  " & vbCrLf & _
                              "          HSupSJ = SDM.SuratJalanNo, " & vbCrLf & _
                              "          HPasiSJ = isnull(PDM.SuratJalanNo,'') " & vbCrLf & _
                              "  FROM    dbo.PO_Master POM  " & vbCrLf & _
                              "          LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                              "                                     AND POM.PoNo = POD.PONo  " & vbCrLf & _
                              "                                     AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                              "                                            AND KD.PoNo = POD.PONo  " & vbCrLf & _
                              "                                            AND KD.SupplierID = POD.SupplierID  "

            ls_SQL = ls_SQL + "                                            AND KD.PartNo = POD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID  " & vbCrLf & _
                              "                                            AND KD.KanbanNo = KM.KanbanNo  " & vbCrLf & _
                              "                                            AND KD.SupplierID = KM.SupplierID  " & vbCrLf & _
                              "                                            AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                              "          LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
                              "                                                 AND KD.KanbanNo = SDD.KanbanNo  " & vbCrLf & _
                              "                                                 AND KD.PONo = SDD.PONo  " & vbCrLf & _
                              "                                                 AND KD.SupplierID = SDD.SupplierID  " & vbCrLf & _
                              "                                                 --AND KD.PartNo = SDD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  "

            ls_SQL = ls_SQL + "                                                 AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
                              "                                                 AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf & _
                              "          INNER JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                              "                                                  AND KD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
                              "                                                  AND KD.SupplierID = PRD.SupplierID  " & vbCrLf & _
                              "                                                  AND KD.PartNo = PRD.PartNo  " & vbCrLf & _
                              "                                                  AND KD.PONo = PRD.PONo  " & vbCrLf & _
                              "                                                  AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                              "          LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo  " & vbCrLf & _
                              "                                                  AND PRM.SupplierID = PRD.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID  "

            ls_SQL = ls_SQL + "                                             AND PRD.KanbanNo = PDD.KanbanNo  " & vbCrLf & _
                              "                                             AND PRD.SupplierID = PDD.SupplierID  " & vbCrLf & _
                              "                                             AND PRD.PartNo = PDD.PartNo  " & vbCrLf & _
                              "                                             AND PRD.PoNo = PDD.PoNo  " & vbCrLf & _
                              "                                             AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf & _
                              "          LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo  " & vbCrLf & _
                              "                                             --AND PDD.SupplierID = PDM.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  "

            ls_SQL = ls_SQL + "  WHERE   'A' = 'A'  " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " )Header " & vbCrLf

            ls_SQL = ls_SQL + " UNION ALL " & vbCrLf & _
                              " SELECT DISTINCT " & vbCrLf & _
                              "         Act = 0 , " & vbCrLf & _
                              "         coldetail = '' , " & vbCrLf & _
                              "         coldetailname = '' , " & vbCrLf & _
                              "         colno = '' , " & vbCrLf & _
                              "         colperiod = '' , " & vbCrLf & _
                              "         colaffiliatecode = '' , " & vbCrLf

            ls_SQL = ls_SQL + "         colaffiliatename = '' , " & vbCrLf & _
                              "         coldeliverylocationcode = '' , " & vbCrLf & _
                              "         coldeliverylocationname = '' , " & vbCrLf & _
                              "         colpono = '' , " & vbCrLf & _
                              "         colsuppliercode = '' , " & vbCrLf & _
                              "         colsuppliername = '' , " & vbCrLf & _
                              "         colpokanban = '' , " & vbCrLf & _
                              "         colkanbanno = '' , " & vbCrLf & _
                              "         colplandeldate = '' , " & vbCrLf & _
                              "         coldeldate = '' , " & vbCrLf & _
                              "         colsj = '' , " & vbCrLf

            ls_SQL = ls_SQL + "         colpasideliverydate = '' , " & vbCrLf & _
                              "         colpasisj = '' , " & vbCrLf & _
                              "         colpartno = pod.PartNo , " & vbCrLf & _
                              "         colpartname = MP.PartName , " & vbCrLf & _
                              "         coluom = UC.Description , " & vbCrLf & _
                              "         coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))), " & vbCrLf & _
                              "         colreceiveqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))), " & vbCrLf & _
                              "         coldefect = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))),  " & vbCrLf & _
                              "         colremaining = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN ISNULL(SDD.DOQty,0) ELSE ISNULL(SDD.DOQty-PRD.GoodRecQty,0) END))), " & vbCrLf

            ls_SQL = ls_SQL + "         colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),  " & vbCrLf & _
                              "         coldeliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END , " & vbCrLf & _
                              "         coldeliveryby = PDM.PIC , " & vbCrLf & _
                              "         POOrder = POM.PONo , " & vbCrLf & _
                              "         idxorder = 1 , " & vbCrLf & _
                              "         kanbanorder = ISNULL(KD.KanbanNo, '-') , " & vbCrLf & _
                              "         affiliateorder = POM.AffiliateID , " & vbCrLf & _
                              "         POD.KanbanCls, " & vbCrLf & _
                              "         HSupSJ = SDM.SuratJalanNo, " & vbCrLf & _
                              "         HPasiSJ = isnull(PDM.SuratJalanNo,'') " & vbCrLf & _
                              " FROM    dbo.PO_Master POM " & vbCrLf & _
                              "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                    AND POM.PoNo = POD.PONo " & vbCrLf

            ls_SQL = ls_SQL + "                                    AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                           AND KD.PoNo = POD.PONo " & vbCrLf & _
                              "                                           AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                           AND KD.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                              "                                           AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              "                                           AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                              "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                              "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                              "                                                AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
                              "                                                AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
                              "                                                AND KD.PONo = SDD.PONo " & vbCrLf

            ls_SQL = ls_SQL + "                                                AND KD.PartNo = SDD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                              "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf & _
                              "                                                AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
                              "         INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                              "                                                 AND SDD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                              "                                                 AND SDD.SupplierID = PRD.SupplierID " & vbCrLf & _
                              "                                                 AND SDD.PartNo = PRD.PartNo " & vbCrLf & _
                              "                                                 AND SDD.PONo = PRD.PONo " & vbCrLf & _
                              "                                                 AND SDM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                              "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                              "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                 AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID " & vbCrLf & _
                              "                                            AND KD.KanbanNo = PDD.KanbanNo " & vbCrLf & _
                              "                                            AND KD.SupplierID = PDD.SupplierID " & vbCrLf & _
                              "                                            AND KD.PartNo = PDD.PartNo " & vbCrLf & _
                              "                                            AND KD.PoNo = PDD.PoNo " & vbCrLf & _
                              "                                            AND SDM.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf & _
                              "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                              "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
                              "                                            AND PDD.SupplierID = PDM.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf

            ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                              "  WHERE   'A' = 'A' " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " ORDER BY h_poorder, h_KanbanCls,h_kanbanorder, HSupSJ,HPasiSJ, h_idxorder "

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

    Private selectedValues As List(Of Object)

    Private Sub up_Delivey()
        Dim ls_Kanban As String = ""
        'Dim ls_PO As String = ""

        'If grid.GetSelectedFieldValues.Count = 0 Then Exit Sub

        'ls_PO = Trim(grid.GetSelectedFieldValues(0, "colpono").ToString)

        'ls_Kanban = "'" & Trim(grid.GetSelectedFieldValues(0, "colkanbanno").ToString) & "'"
        'If grid.GetSelectedFieldValues.Count > 0 Then
        '    For iLoop = 0 To grid.GetSelectedFieldValues.Count - 1
        '        ls_Kanban = ls_Kanban + ",'" & Trim(grid.GetSelectedFieldValues(0, "colkanbanno").ToString) & "'"
        '    Next
        'End If

        'Session("POList") = ls_PO
        'Session("KanbanList") = ls_Kanban
        Dim fieldValues As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colpono", "colsuppliercode", "colpartno"})
        For Each item As Object() In fieldValues
            ls_Kanban = item(0).ToString
        Next item

    End Sub

#End Region

    Protected Sub btnsubmenu_Click(sender As Object, e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_BatchUpdate(sender As Object, e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_Kanban As String = ""

        Dim ls_DeliveryDate As String = ""
        Dim ls_AffiliateCode As String = ""
        Dim ls_AffiliateName As String = ""
        Dim ls_SuratJalanNo As String = ""
        Dim ls_DeliveryCode As String = ""
        Dim ls_DeliveryName As String = ""
        Dim ls_DriverName As String = ""
        Dim ls_Contact As String = ""
        Dim ls_Nopol As String = ""
        Dim ls_JenisArmada As String = ""
        Dim ls_PO As String = ""
        Dim ls_Supplier As String = ""
        Dim ls_SupplierName As String = ""
        Dim ls_SuratJalan As String = ""
        Dim ls_PSJ As String = ""
        Dim ls_SuppSJ As String = ""
        Dim ls_filter As String = ""
        Dim ls_Status As String = ""
        Dim ls_OLDAFF As String = ""

        Session.Remove("POList")

        With grid
            If e.UpdateValues.Count = 0 Then Exit Sub
            If (e.UpdateValues(0).NewValues("Act").ToString()) = 1 Then
                'ls_DeliveryDate = Trim(e.UpdateValues(0).NewValues("colpasideliverydate").ToString())
                ls_DeliveryDate = "01 Jan 1900"
                ls_AffiliateCode = Trim(e.UpdateValues(0).NewValues("colaffiliatecode").ToString())
                ls_AffiliateName = Trim(e.UpdateValues(0).NewValues("colaffiliatename").ToString())
                ls_SuratJalanNo = Trim(e.UpdateValues(0).NewValues("colpasisj").ToString())
                'ls_SuratJalanNo = ""
                If Right(Trim(e.UpdateValues(0).NewValues("coldetail").ToString()), 1) = "Y" Then
                    ls_Status = "DELIVERY"
                Else
                    ls_Status = "DETAIL"
                End If
                ls_DeliveryCode = Trim(e.UpdateValues(0).NewValues("coldeliverylocationcode").ToString())
                ls_DeliveryName = Trim(e.UpdateValues(0).NewValues("coldeliverylocationname").ToString())

                ls_PO = "'" & Trim(e.UpdateValues(0).NewValues("colpono").ToString()) & "'"
                ls_Kanban = "'" & Trim(e.UpdateValues(0).NewValues("colkanbanno").ToString()) & "'"
                ls_PSJ = "'" & Trim(e.UpdateValues(0).NewValues("colpasisj").ToString()) & "'"
                ls_Supplier = Trim(e.UpdateValues(0).NewValues("colsuppliercode").ToString())
                ls_SupplierName = Trim(e.UpdateValues(0).NewValues("colsuppliername").ToString())
                ls_SuppSJ = "'" & Trim(e.UpdateValues(0).NewValues("colsj").ToString()) & "'"
                ls_filter = "'" & Trim(e.UpdateValues(0).NewValues("colsj").ToString()) & _
                                  Trim(e.UpdateValues(0).NewValues("colpono").ToString()) & _
                                  Trim(e.UpdateValues(0).NewValues("colsuppliercode").ToString()) & _
                                  Trim(e.UpdateValues(0).NewValues("colaffiliatecode").ToString()) & _
                                  Trim(e.UpdateValues(0).NewValues("colkanbanno").ToString()) & "'"
                '(RTRIM(SDD.SuratJalanNo)+RTRIM(SDD.PONO)+RTRIM(SDD.SupplierID)+RTRIM(SDD.AffiliateID)+RTRIM(SDD.KanbanNo)+RTRIM(SDD.PartNo)
            End If

            If e.UpdateValues.Count > 1 Then
                For i = 0 To e.UpdateValues.Count - 1
                    If (e.UpdateValues(i).NewValues("Act").ToString()) = 1 Then
                        If ls_OLDAFF = "" Then
                            ls_OLDAFF = Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString())
                        Else
                            If Trim(ls_OLDAFF) <> Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) Then
                                Exit Sub
                            End If
                        End If
                        'ls_OLDAFF = Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString())

                        ls_PO = ls_PO + ",'" & Trim(e.UpdateValues(i).NewValues("colpono").ToString()) & "'"
                        ls_Kanban = ls_Kanban + ",'" & Trim(e.UpdateValues(i).NewValues("colkanbanno").ToString()) & "'"
                        ls_PSJ = ls_PSJ + ",'" & Trim(e.UpdateValues(i).NewValues("colpasisj").ToString()) & "'"
                        ls_SuppSJ = ls_SuppSJ + ",'" & Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & "'"
                        ls_filter = ls_filter + ", '" & Trim(e.UpdateValues(i).NewValues("colsj").ToString()) & _
                                                        Trim(e.UpdateValues(i).NewValues("colpono").ToString()) & _
                                                        Trim(e.UpdateValues(i).NewValues("colsuppliercode").ToString()) & _
                                                        Trim(e.UpdateValues(i).NewValues("colaffiliatecode").ToString()) & _
                                                        Trim(e.UpdateValues(i).NewValues("colkanbanno").ToString()) & "'"

                    End If
                Next
            End If
        End With
        Session("POList") = ls_DeliveryDate & "|" & ls_AffiliateCode & "|" & ls_AffiliateName & _
                            "|" & ls_SuratJalanNo & "|" & ls_DeliveryCode & _
                            "|" & ls_DeliveryName & "|" & "||||" & ls_PO & "||" & ls_Supplier & "|" & ls_SupplierName & "|" & ls_SuratJalanNo & "|" & ls_PSJ & "|" & ls_SuppSJ & "|" & ls_filter & "|" & ls_Status
        Session("KanbanList") = ls_Kanban
        HF.Set("Update", "1")

    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            If pAction <> "detail" Then
                Dim pPlan As Date = Split(e.Parameters, "|")(1)
                Dim pSupplierDeliver As String = Split(e.Parameters, "|")(2)
                Dim pRemaining As String = Split(e.Parameters, "|")(3)
                Dim psj As String = Split(e.Parameters, "|")(4)
                Dim pDateFrom As Date = Split(e.Parameters, "|")(5)
                Dim pDateTo As Date = Split(e.Parameters, "|")(6)
                Dim pSupplier As String = Split(e.Parameters, "|")(7)
                Dim pPart As String = Split(e.Parameters, "|")(8)
                Dim pPoNo As String = Split(e.Parameters, "|")(9)
                Dim pKanban As String = Split(e.Parameters, "|")(10)
            End If
            Select Case pAction
                Case "gridload"
                    'Call up_GridLoad(pPlan, pSupplierDeliver, pRemaining, psj, pDateFrom, pDateTo, pSupplier, pPart, pPoNo, pKanban)
                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "update"
                    grid.UpdateEdit()
                Case "detail"
                    If Not IsNothing(Session("POList")) = True Then
                        DevExpress.Web.ASPxGridView.ASPxGridView.RedirectOnCallback("~/Delivery/DeliveryToAffEntry.aspx")
                    Else
                        Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "coldetail" Or e.DataColumn.FieldName = "Act") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        End If

        If e.DataColumn.FieldName = "Act" Then
            If (e.GetValue("colkanbanno") = "" Or Left(e.GetValue("colpokanban"), 2) = "NO") Then
                e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If

        If (e.GetValue("colkanbanno") = "" Or Left(e.GetValue("colpokanban"), 2) = "NO") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        End If

        If (e.DataColumn.FieldName = "coldetail") Then
            If (e.GetValue("colpokanban") = "NO") Then
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If

        If e.DataColumn.FieldName = "colremaining" Then
            If (Trim(e.GetValue("coldetailname")) = "") Then
                If (e.GetValue("coldeliveryqty") > (CDbl(e.GetValue("colreceiveqty")) + CDbl(e.GetValue("coldefect")))) Then
                    e.Cell.BackColor = Color.Fuchsia
                End If
            End If
        End If
    End Sub

    Private Sub grid_HtmlRowPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
        Try
            Dim getRowValues As String = e.GetValue("colpono")
            If Not IsNothing(getRowValues) Then
                If getRowValues.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If

        Catch ex As Exception
            'Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'Session("E01Msg") = lblInfo.Text
        End Try
    End Sub

    Private Sub grid_PageIndexChanged(sender As Object, e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub

    'Private Sub btndeliver_Click(sender As Object, e As System.EventArgs) Handles btndeliver.Click
    '    Response.Redirect("~/Delivery/DeliveryToAffEntry.aspx")
    'End Sub


End Class