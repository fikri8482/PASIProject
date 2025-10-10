Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing

Public Class AffReceivingConf
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

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
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If checkbox1.Checked = True Then
                ls_Filter = ls_Filter + "AND DM.DeliveryDate BETWEEN '" & dtPasiDeliveryDateFrom.Value & "' AND '" & Format(dtPasiDeliveryDateTo.Value, "dd MMM yyyy") & "' " & vbCrLf
            End If

            If rbInvoiceByPasi.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(IPM.InvoiceNo, 0) = 0 " & vbCrLf
            ElseIf rbInvoiceByPasi.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(IPM.InvoiceNo,0) <> 0 " & vbCrLf
            End If

            If txtPasiSJ.Text <> "" Then
                ls_Filter = ls_Filter + " AND DM.SuratJalanNo LIKE '%" & txtPasiSJ.Text & "%' " & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_Filter = ls_Filter + " AND IPM.InvoiceDate BETWEEN '" & dtPasiInvoiceDateFrom.Value & "' AND '" & dtPasiInvoiceDateTo.Value & "' " & vbCrLf
            End If

            If txtPasiInvoiceNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND IPM.InvoiceNo LIKE '%" & txtPasiInvoiceNo.Text & "%' " & vbCrLf
            End If

            If cboaffiliate.Text <> "" Then
                If cboaffiliate.Text <> "== ALL ==" Then
                    ls_Filter = ls_Filter + " AND DD.AffiliateID = '" & cboaffiliate.Text & "' " & vbCrLf
                End If
            End If

            If txtpono.Text <> "" Then
                ls_Filter = ls_Filter + " AND DD.PONo LIKE '%" & txtpono.Text & "%'" & vbCrLf
            End If

            'If cbSuppDelDate.Checked = True Then
            '    ls_Filter = ls_Filter + " AND SDM.DeliveryDate = '" & dtSuppDelDate.Value & "' " & vbCrLf
            'End If

            'If txtSuppSJ.Text <> "" Then
            '    ls_Filter = ls_Filter + " AND DD.SuratJalanNo LIKE '%" & txtSuppSJ.Text & "%' " & vbCrLf
            'End If

            ls_SQL = " SELECT DISTINCT " & vbCrLf & _
                  " 	Act = 0 " & vbCrLf & _
                  " 	, coldetail = (CASE WHEN ISNULL(DM.SuratJalanno, '') = '' THEN '' else 'InvToAff.aspx?prm='    " & vbCrLf & _
                  "            + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(IPM.InvoiceDate, '')), 106) + '|'   " & vbCrLf & _
                  "            + RTRIM(DM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'''      " & vbCrLf & _
                  "            + RTRIM(ISNULL(DM.SuratJalanNo, '')) + '''|'     " & vbCrLf & _
                  "            + RTRIM(ISNULL(DM.InvoiceNo, '')) + '|'    " & vbCrLf & _
                  "            + RTRIM(ISNULL(IPM.PaymentTerm, '')) + '|'   " & vbCrLf & _
                  "            + RTRIM(ISNULL(IPM.DueDate, '')) + '|'   " & vbCrLf & _
                  "            + RTRIM(ISNULL(IPM.Notes, '')) + '|'''    " & vbCrLf & _
                  "            + '' + '''|'''   "

            ls_SQL = ls_SQL + "            + '' + '''|' END) " & vbCrLf & _
                              " 	, coldetailname = CASE WHEN ISNULL(DM.SuratJalanno, '') = ''   " & vbCrLf & _
                              " 						THEN ''    " & vbCrLf & _
                              " 							ELSE 'DETAIL'    " & vbCrLf & _
                              " 						END " & vbCrLf & _
                              " 	, colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,KM.KanbanDate),106), 8) " & vbCrLf & _
                              " 	, colno = ''  " & vbCrLf & _
                              " 	, colaffiliatecode = DM.AffiliateID " & vbCrLf & _
                              " 	, colaffiliatename = MA.AffiliateName " & vbCrLf & _
                              " 	, colpono = '' " & vbCrLf & _
                              "     , colsuppliercode = '' "

            ls_SQL = ls_SQL + "     , colsuppliername = '' " & vbCrLf & _
                              "     , colpokanban = ''    " & vbCrLf & _
                              " 	, colkanbanno = '' " & vbCrLf & _
                              " 	, colplandeldate = ''               " & vbCrLf & _
                              " 	, coldeldate = ''              " & vbCrLf & _
                              " 	, colsj = '' " & vbCrLf & _
                              " 	, colpasideliverydate = CASE WHEN ISNULL(DM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(DM.DeliveryDate, '')), 106) END " & vbCrLf & _
                              " 	, colpasisj = DM.SuratJalanNo " & vbCrLf & _
                              " 	, colpartno = '' " & vbCrLf & _
                              " 	, colpartname = '' " & vbCrLf & _
                              " 	, coluom = '' "

            ls_SQL = ls_SQL + " 	, coldeliveryqty = '' " & vbCrLf & _
                              " 	, colpasideliveryqty = '' " & vbCrLf & _
                              " 	, colAffRecQty = '' " & vbCrLf & _
                              " 	, colRemainRecQty = '' " & vbCrLf & _
                              " 	, colAffRecDate = '' " & vbCrLf & _
                              " 	, colAffRecBy = ''  " & vbCrLf & _
                              " 	, colPasiInvNo = '' " & vbCrLf & _
                              " 	, colPasiInvDate = ''              " & vbCrLf & _
                              " 	, '' H_POORDER " & vbCrLf & _
                              "     , H_IDXORDER = 0 " & vbCrLf & _
                              " 	, H_KANBANORDER = '' "

            ls_SQL = ls_SQL + "     , H_AFFILIATEORDER = DM.AffiliateID " & vbCrLf & _
                              " 	, H_KANBANCLS = DD.POKanbanCls    " & vbCrLf & _
                              " 	, H_SupSJ = ''  " & vbCrLf & _
                              " 	, H_PasiSJ = DM.SuratJalanNo   " & vbCrLf & _
                              " FROM DOPASI_Master DM " & vbCrLf & _
                              " INNER JOIN DOPASI_Detail DD ON DM.SuratJalanNo = DD.SuratJalanNo AND DM.AffiliateID = DD.AffiliateID  " & vbCrLf & _
                              " LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = DD.AffiliateID AND KD.KanbanNo = DD.KanbanNo AND KD.PartNo = DD.PartNo AND KD.SupplierID = DD.SupplierID " & vbCrLf & _
                              " LEFT JOIN Kanban_Master KM ON KM.AffiliateID = KD.AffiliateID AND KM.KanbanNo = KD.KanbanNo AND KM.SupplierID = KD.SupplierID " & vbCrLf & _
                              " LEFT JOIN InvoicePASI_Master IPM ON IPM.AffiliateID = DM.AffiliateID AND IPM.InvoiceNo =  DM.InvoiceNo AND IPM.SuratJalanNo = DM.SuratJalanNo " & vbCrLf & _
                              " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = DM.AffiliateID " & vbCrLf & _
                              " WHERE 'A' = 'A' " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " UNION ALL   " & vbCrLf & _
                              " SELECT DISTINCT    " & vbCrLf & _
                              " 	Act = 0 " & vbCrLf & _
                              " 	, coldetail = '' " & vbCrLf & _
                              " 	, coldetailname = '' " & vbCrLf & _
                              " 	, colno = '' " & vbCrLf & _
                              " 	, colperiod = '' " & vbCrLf & _
                              " 	, colaffiliatecode = '' " & vbCrLf & _
                              " 	, colaffiliatename = '' " & vbCrLf & _
                              " 	, colpono = DD.PONo   " & vbCrLf & _
                              " 	, colsuppliercode = DD.SupplierID "

            ls_SQL = ls_SQL + " 	, colsuppliername = MS.SupplierName " & vbCrLf & _
                              " 	, colpokanban = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN 'NO' ELSE 'YES' END " & vbCrLf & _
                              " 	, colkanbanno = ISNULL(KD.KanbanNo, '-') " & vbCrLf & _
                              " 	, colplandeldate = '' " & vbCrLf & _
                              " 	, coldeldate = '' " & vbCrLf & _
                              " 	, colsj = '' " & vbCrLf & _
                              " 	, colpasideliverydate = '' " & vbCrLf & _
                              " 	, colpasisj = '' " & vbCrLf & _
                              " 	, colpartno = DD.PartNo " & vbCrLf & _
                              " 	, colpartname = MP.PartName " & vbCrLf & _
                              " 	, coluom = MU.Description "

            ls_SQL = ls_SQL + " 	, coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(DD.DOQty,0)))) " & vbCrLf & _
                              " 	, colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(DD.DOQty,0)))) " & vbCrLf & _
                              " 	, colAffRecQty = 0 " & vbCrLf & _
                              " 	, colRemainRecQty = 0 " & vbCrLf & _
                              " 	, colAffRecDate = '' " & vbCrLf & _
                              " 	, colAffRecBy = '' " & vbCrLf & _
                              " 	, colPasiInvNo = DM.InvoiceNo " & vbCrLf & _
                              " 	, colPasiInvDate = CASE WHEN ISNULL(IPM.InvoiceDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(IPM.InvoiceDate, '')), 106) END " & vbCrLf & _
                              " 	, POOrder = DD.PONo " & vbCrLf & _
                              " 	, idxorder = 1 " & vbCrLf & _
                              " 	, kanbanorder = ISNULL(KD.KanbanNo, '-') "

            ls_SQL = ls_SQL + " 	, affiliateorder = DM.AffiliateID " & vbCrLf & _
                              " 	, DD.POKanbanCls " & vbCrLf & _
                              " 	, H_SupSJ = '' " & vbCrLf & _
                              " 	, H_PasiSJ = DM.SuratJalanNo   " & vbCrLf & _
                              " FROM " & vbCrLf & _
                              " DOPASI_Master DM " & vbCrLf & _
                              " INNER JOIN DOPASI_Detail DD ON DM.SuratJalanNo = DD.SuratJalanNo AND DM.AffiliateID = DD.AffiliateID  " & vbCrLf & _
                              " LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = DD.AffiliateID AND KD.KanbanNo = DD.KanbanNo AND KD.PartNo = DD.PartNo AND KD.SupplierID = DD.SupplierID " & vbCrLf & _
                              " LEFT JOIN Kanban_Master KM ON KM.AffiliateID = KD.AffiliateID AND KM.KanbanNo = KD.KanbanNo AND KM.SupplierID = KD.SupplierID " & vbCrLf & _
                              " LEFT JOIN InvoicePASI_Master IPM ON IPM.AffiliateID = DM.AffiliateID AND IPM.InvoiceNo =  DM.InvoiceNo AND IPM.SuratJalanNo = DM.SuratJalanNo " & vbCrLf & _
                              " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = DM.AffiliateID "

            ls_SQL = ls_SQL + " LEFT JOIN MS_Parts MP ON MP.PartNo = DD.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_Supplier MS ON MS.SupplierID = DD.SupplierID " & vbCrLf & _
                              " LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                              " WHERE 'A' = 'A' " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter & vbCrLf

            ls_SQL = ls_SQL + " ORDER BY h_affiliateorder, H_PasiSJ, h_poorder, h_idxorder, h_kanbanorder ,colpartno    " & vbCrLf & _
                              "  "

            'ls_SQL = " SELECT DISTINCT   " & vbCrLf & _
            '      "           Act = 0 ,   " & vbCrLf & _
            '      "           coldetail = (CASE WHEN ISNULL(RAM.SuratJalanno, '') = '' THEN '' else 'InvToAff.aspx?prm='   " & vbCrLf & _
            '      "           + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(IPM.InvoiceDate, '')), 106) + '|'  " & vbCrLf & _
            '      "           + RTRIM(POM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'''     " & vbCrLf & _
            '      "           + RTRIM(ISNULL(PDM.SuratJalanNo, '')) + '''|'    " & vbCrLf & _
            '      "           + RTRIM(ISNULL(IPD.InvoiceNo, '')) + '|'   " & vbCrLf & _
            '      "           + RTRIM(ISNULL(IPM.PaymentTerm, '')) + '|'  " & vbCrLf & _
            '      "           + RTRIM(ISNULL(IPM.DueDate, '')) + '|'  " & vbCrLf & _
            '      "           + RTRIM(ISNULL(IPM.Notes, '')) + '|'''   " & vbCrLf & _
            '      "           + '' + '''|'''  " & vbCrLf

            'ls_SQL = ls_SQL + "           + '' + '''|' END),  " & vbCrLf & _
            '                  "           coldetailname = CASE WHEN ISNULL(RAM.SuratJalanno, '') = ''  " & vbCrLf & _
            '                  "                                THEN ''   " & vbCrLf & _
            '                  "                                ELSE 'DETAIL'   " & vbCrLf & _
            '                  "                           END ,   " & vbCrLf & _
            '                  "           colno = '' ,   " & vbCrLf & _
            '                  "           colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,RAM.ReceiveDate),106), 8),   " & vbCrLf & _
            '                  "           colaffiliatecode = POM.AffiliateID ,   " & vbCrLf & _
            '                  "           colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
            '                  "           colpono = '',--POM.PONo ,   " & vbCrLf & _
            '                  "           colsuppliercode = '',--POM.SupplierID ,   " & vbCrLf

            'ls_SQL = ls_SQL + "           colsuppliername = '',--MS.SupplierName ,   " & vbCrLf & _
            '                  "           colpokanban = '' ,   " & vbCrLf & _
            '                  "           colkanbanno = '',  " & vbCrLf & _
            '                  "           colplandeldate = '',--CASE WHEN ISNULL(KM.KanbanDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(KM.KanbanDate, '')), 106) END ,   " & vbCrLf & _
            '                  "              " & vbCrLf & _
            '                  "           coldeldate = '',--CASE WHEN ISNULL(SDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(SDM.DeliveryDate, '')), 106) END ,   " & vbCrLf & _
            '                  "             " & vbCrLf & _
            '                  "           colsj = '' ,   " & vbCrLf & _
            '                  "           colpasideliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,   " & vbCrLf

            'ls_SQL = ls_SQL + "           colpasisj = PDM.SuratJalanNo ,   " & vbCrLf & _
            '                  "           colpartno = '' ,   " & vbCrLf & _
            '                  "           colpartname = '' ,   " & vbCrLf & _
            '                  "           coluom = '' ,   " & vbCrLf & _
            '                  "           coldeliveryqty = '' ,   " & vbCrLf & _
            '                  "           colpasideliveryqty = '' ,   " & vbCrLf & _
            '                  "           colAffRecQty = '',  " & vbCrLf & _
            '                  "           colRemainRecQty = '',  " & vbCrLf & _
            '                  "           colAffRecDate = '',  " & vbCrLf & _
            '                  "           colAffRecBy = '',  " & vbCrLf & _
            '                  "           colPasiInvNo = '',  " & vbCrLf

            'ls_SQL = ls_SQL + "           colPasiInvDate = '',  " & vbCrLf & _
            '                  "             " & vbCrLf & _
            '                  "           /*pom.PONo*/'' H_POORDER ,   " & vbCrLf & _
            '                  "           H_IDXORDER = 0 ,   " & vbCrLf & _
            '                  "           H_KANBANORDER = '',--ISNULL(KD.KanbanNo, '-') ,   " & vbCrLf & _
            '                  "           H_AFFILIATEORDER = POM.AffiliateID ,   " & vbCrLf & _
            '                  "           H_KANBANCLS = pod.KanbanCls   " & vbCrLf & _
            '                  "           ,H_SupSJ = '' " & vbCrLf & _
            '                  "           ,H_PasiSJ = PDD.SuratJalanNo  " & vbCrLf & _
            '                  "   FROM    dbo.PO_Master POM   " & vbCrLf & _
            '                  "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                      AND POM.PoNo = POD.PONo   " & vbCrLf & _
            '                  "                                      AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.PoNo = POD.PONo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = POD.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                                  AND KD.KanbanNo = SDD.KanbanNo   " & vbCrLf & _
            '                  "                                                  AND KD.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "                                                  AND KD.PartNo = SDD.PartNo   " & vbCrLf & _
            '                  "                                                  AND KD.PoNo = SDD.PoNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND SDD.KanbanNo = PRD.KanbanNo   " & vbCrLf & _
            '                  "                                                   AND SDD.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "                                                   AND SDD.PartNo = PRD.PartNo " & vbCrLf

            'ls_SQL = ls_SQL + "                                                   AND SDD.PONo = PRD.PONo   " & vbCrLf & _
            '                  "                                                   AND PRD.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                   --AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "           --INNER JOIN dbo.DOPASI_Detail PDD  " & vbCrLf & _
            '                  "           LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))  " & vbCrLf & _
            '                  "           			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD  " & vbCrLf & _
            '                  "                                              ON PRD.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "                                              AND PRD.KanbanNo = PDD.KanbanNo   " & vbCrLf & _
            '                  "                                              AND PRD.SupplierID = PDD.SupplierID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                              AND PRD.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "                                              AND PRD.PoNo = PDD.PoNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID   " & vbCrLf & _
            '                  "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SupplierID = PDM.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                       AND PDD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
            '                  "                                                       AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "                                                       AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
            '                  "                                                       AND PDD.PoNo = RAD.PoNo  " & vbCrLf

            'ls_SQL = ls_SQL + "           INNER JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                       AND RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                       --AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "  		 LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
            '                  "  													AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
            '                  "  													AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
            '                  "  													AND RAD.PONo = IPD.PONo  " & vbCrLf & _
            '                  "  													AND RAD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
            '                  "  		 LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
            '                  "  													AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
            '                  "  													--AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf

            'ls_SQL = ls_SQL + "  													 " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "    WHERE  'A' = 'A'   " & vbCrLf & _
            '                  "           AND POM.CommercialCls <> 0  " & vbCrLf & _
            '                  "           AND ISNULL(PDM.suratjalanno,'') <> '' " & vbCrLf & _
            '                  "           AND ISNULL(IPM.InvoiceNo,'') <> '' " & vbCrLf & _
            '                  "           --AND Isnull(PDM.suratjalanno, '') not in (Select distinct suratjalanno from invoicePasi_Detail) " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter

            'ls_SQL = ls_SQL + " UNION ALL SELECT DISTINCT   " & vbCrLf & _
            '                  "           Act = 0 ,   " & vbCrLf & _
            '                  "           coldetail = (CASE WHEN ISNULL(RAM.SuratJalanno, '') = '' THEN '' else 'InvToAff.aspx?prm='   " & vbCrLf & _
            '                  "           + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(IPM.InvoiceDate, '')), 106) + '|'  " & vbCrLf & _
            '                  "           + RTRIM(POM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'''     " & vbCrLf & _
            '                  "           + RTRIM(ISNULL(PDM.SuratJalanNo, '')) + '''|'    " & vbCrLf & _
            '                  "           + RTRIM(ISNULL(IPD.InvoiceNo, '')) + '|'   " & vbCrLf & _
            '                  "           + RTRIM(ISNULL(IPM.PaymentTerm, '')) + '|'  " & vbCrLf & _
            '                  "           + RTRIM(ISNULL(IPM.DueDate, '')) + '|'  " & vbCrLf & _
            '                  "           + RTRIM(ISNULL(IPM.Notes, '')) + '|'''   " & vbCrLf & _
            '                  "           + '' + '''|'''  " & vbCrLf

            'ls_SQL = ls_SQL + "           + '' + '''|' END),  " & vbCrLf & _
            '                  "           coldetailname = CASE WHEN ISNULL(RAM.SuratJalanno, '') = ''  " & vbCrLf & _
            '                  "                                THEN ''   " & vbCrLf & _
            '                  "                                ELSE 'DETAIL'   " & vbCrLf & _
            '                  "                           END ,   " & vbCrLf & _
            '                  "           colno = '' ,   " & vbCrLf & _
            '                  "           colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,RAM.ReceiveDate),106), 8),   " & vbCrLf & _
            '                  "           colaffiliatecode = POM.AffiliateID ,   " & vbCrLf & _
            '                  "           colaffiliatename = MA.AffiliateName ,    " & vbCrLf & _
            '                  "           colpono = '',--POM.PONo ,   " & vbCrLf & _
            '                  "           colsuppliercode = '',--POM.SupplierID ,   " & vbCrLf

            'ls_SQL = ls_SQL + "           colsuppliername = '',--MS.SupplierName ,   " & vbCrLf & _
            '                  "           colpokanban = '' ,   " & vbCrLf & _
            '                  "           colkanbanno = '',  " & vbCrLf & _
            '                  "           colplandeldate = '',--CASE WHEN ISNULL(KM.KanbanDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(KM.KanbanDate, '')), 106) END ,   " & vbCrLf & _
            '                  "              " & vbCrLf & _
            '                  "           coldeldate = '',--CASE WHEN ISNULL(SDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(SDM.DeliveryDate, '')), 106) END ,   " & vbCrLf & _
            '                  "             " & vbCrLf & _
            '                  "           colsj = '' ,   " & vbCrLf & _
            '                  "           colpasideliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,   " & vbCrLf

            'ls_SQL = ls_SQL + "           colpasisj = PDM.SuratJalanNo ,   " & vbCrLf & _
            '                  "           colpartno = '' ,   " & vbCrLf & _
            '                  "           colpartname = '' ,   " & vbCrLf & _
            '                  "           coluom = '' ,   " & vbCrLf & _
            '                  "           coldeliveryqty = '' ,   " & vbCrLf & _
            '                  "           colpasideliveryqty = '' ,   " & vbCrLf & _
            '                  "           colAffRecQty = '',  " & vbCrLf & _
            '                  "           colRemainRecQty = '',  " & vbCrLf & _
            '                  "           colAffRecDate = '',  " & vbCrLf & _
            '                  "           colAffRecBy = '',  " & vbCrLf & _
            '                  "           colPasiInvNo = '',  " & vbCrLf

            'ls_SQL = ls_SQL + "           colPasiInvDate = '',  " & vbCrLf & _
            '                  "             " & vbCrLf & _
            '                  "           /*pom.PONo*/'' H_POORDER ,   " & vbCrLf & _
            '                  "           H_IDXORDER = 0 ,   " & vbCrLf & _
            '                  "           H_KANBANORDER = '',--ISNULL(KD.KanbanNo, '-') ,   " & vbCrLf & _
            '                  "           H_AFFILIATEORDER = POM.AffiliateID ,   " & vbCrLf & _
            '                  "           H_KANBANCLS = pod.KanbanCls   " & vbCrLf & _
            '                  "           ,H_SupSJ = '' " & vbCrLf & _
            '                  "           ,H_PasiSJ = PDD.SuratJalanNo  " & vbCrLf & _
            '                  "   FROM    dbo.PO_Master POM   " & vbCrLf & _
            '                  "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                      AND POM.PoNo = POD.PONo   " & vbCrLf & _
            '                  "                                      AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.PoNo = POD.PONo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = POD.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                                  AND KD.KanbanNo = SDD.KanbanNo   " & vbCrLf & _
            '                  "                                                  AND KD.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "                                                  AND KD.PartNo = SDD.PartNo   " & vbCrLf & _
            '                  "                                                  AND KD.PoNo = SDD.PoNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND SDD.KanbanNo = PRD.KanbanNo   " & vbCrLf & _
            '                  "                                                   AND SDD.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "                                                   AND SDD.PartNo = PRD.PartNo " & vbCrLf

            'ls_SQL = ls_SQL + "                                                   AND SDD.PONo = PRD.PONo   " & vbCrLf & _
            '                  "                                                   AND PRD.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                   --AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "           --INNER JOIN dbo.DOPASI_Detail PDD  " & vbCrLf & _
            '                  "           LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))  " & vbCrLf & _
            '                  "           			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD  " & vbCrLf & _
            '                  "                                              ON PRD.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "                                              AND PRD.KanbanNo = PDD.KanbanNo   " & vbCrLf & _
            '                  "                                              AND PRD.SupplierID = PDD.SupplierID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                              AND PRD.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "                                              AND PRD.PoNo = PDD.PoNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID   " & vbCrLf & _
            '                  "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SupplierID = PDM.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                       AND PDD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
            '                  "                                                       AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "                                                       AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
            '                  "                                                       AND PDD.PoNo = RAD.PoNo  " & vbCrLf

            'ls_SQL = ls_SQL + "           INNER JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                       AND RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                       --AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "  		 LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
            '                  "  													AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
            '                  "  													AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
            '                  "  													AND RAD.PONo = IPD.PONo  " & vbCrLf & _
            '                  "  													AND RAD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
            '                  "  		 LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
            '                  "  													AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
            '                  "  													--AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf

            'ls_SQL = ls_SQL + "  													 " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "    WHERE  'A' = 'A'   " & vbCrLf & _
            '                  "           AND POM.CommercialCls <> 0  " & vbCrLf & _
            '                  "           AND ISNULL(PDM.suratjalanno,'') <> '' " & vbCrLf & _
            '                  "           AND Isnull(PDM.suratjalanno, '') not in (Select distinct suratjalanno from invoicePasi_Detail) " & vbCrLf & _
            '                  "           AND Isnull(IPM.InvoiceNo, '') = '' " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter

            'ls_SQL = ls_SQL + "  UNION ALL  " & vbCrLf
            'ls_SQL = ls_SQL + " SELECT DISTINCT   " & vbCrLf & _
            '      "           Act = 0 ,   " & vbCrLf & _
            '      "           coldetail = '' ,   " & vbCrLf & _
            '      "           coldetailname = '' ,   " & vbCrLf & _
            '      "           colno = '' ,   " & vbCrLf & _
            '      "           colperiod = '' ,   " & vbCrLf & _
            '      "           colaffiliatecode = '' ,   " & vbCrLf & _
            '      "           colaffiliatename = '' ,    " & vbCrLf & _
            '      "           colpono = POM.PONo ,   " & vbCrLf & _
            '      "           colsuppliercode = POM.SupplierID ,   " & vbCrLf & _
            '      "           colsuppliername = MS.SupplierName ,   " & vbCrLf

            'ls_SQL = ls_SQL + "           colpokanban = CASE WHEN ISNULL(KD.KanbanNo, '0') = '0' THEN 'NO' ELSE 'YES' END ,     " & vbCrLf & _
            '                  "           colkanbanno = ISNULL(KD.KanbanNo, '-') ,   " & vbCrLf & _
            '                  "           colplandeldate = '' ,   " & vbCrLf & _
            '                  "           coldeldate = '' ,   " & vbCrLf & _
            '                  "           colsj = '' ,   " & vbCrLf & _
            '                  "           colpasideliverydate = '' ,   " & vbCrLf & _
            '                  "           colpasisj = '' ,   " & vbCrLf & _
            '                  "           colpartno = pod.PartNo ,   " & vbCrLf & _
            '                  "           colpartname = MP.PartName ,   " & vbCrLf & _
            '                  "           coluom = UC.Description ,   " & vbCrLf & _
            '                  "           coldeliveryqty = '',--CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(SDD.DOQty,0)))),  " & vbCrLf

            'ls_SQL = ls_SQL + "           colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),  " & vbCrLf & _
            '                  "           colAffRecQty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(RAD.RecQty,0)))),  " & vbCrLf & _
            '                  "           colRemainRecQty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty - RAD.RecQty,0)))),  " & vbCrLf & _
            '                  "           colAffRecDate = CASE WHEN ISNULL(RAM.ReceiveDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(RAM.ReceiveDate, '')), 106) END ,  " & vbCrLf & _
            '                  "           colAffRecBy = RAM.ReceiveBy,  " & vbCrLf & _
            '                  "           colPasiInvNo = IPD.InvoiceNo,  " & vbCrLf & _
            '                  "           colPasiInvDate = CASE WHEN ISNULL(IPM.InvoiceDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(IPM.InvoiceDate, '')), 106) END,  " & vbCrLf & _
            '                  "             " & vbCrLf & _
            '                  "             " & vbCrLf & _
            '                  "           POOrder = POM.PONo ,   " & vbCrLf & _
            '                  "           idxorder = 1 ,   " & vbCrLf

            'ls_SQL = ls_SQL + "           kanbanorder = ISNULL(KD.KanbanNo, '-') ,   " & vbCrLf & _
            '                  "           affiliateorder = POM.AffiliateID ,   " & vbCrLf & _
            '                  "           POD.KanbanCls   " & vbCrLf & _
            '                  "           ,H_SupSJ = ''--SDM.SuratJalanNo  " & vbCrLf & _
            '                  "           ,H_PasiSJ = PDD.SuratJalanNo  " & vbCrLf & _
            '                  "   FROM    dbo.PO_Master POM   " & vbCrLf & _
            '                  "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
            '                  "                                      AND POM.PoNo = POD.PONo   " & vbCrLf & _
            '                  "                                      AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.PoNo = POD.PONo   " & vbCrLf

            'ls_SQL = ls_SQL + "                                             AND KD.SupplierID = POD.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
            '                  "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
            '                  "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND KD.KanbanNo = SDD.KanbanNo   " & vbCrLf & _
            '                  "                                                  AND KD.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "                                                  AND KD.PartNo = SDD.PartNo   " & vbCrLf & _
            '                  "                                                  AND KD.PoNo = SDD.PoNo  " & vbCrLf

            'ls_SQL = ls_SQL + "           LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
            '                  "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND SDD.KanbanNo = PRD.KanbanNo   " & vbCrLf & _
            '                  "                                                   AND SDD.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "                                                   AND SDD.PartNo = PRD.PartNo " & vbCrLf & _
            '                  "                                                   AND SDD.PONo = PRD.PONo   " & vbCrLf & _
            '                  "                                                   AND PRD.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
            '                  "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo   " & vbCrLf

            'ls_SQL = ls_SQL + "                                                   --AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf & _
            '                  "           --INNER JOIN dbo.DOPASI_Detail PDD  " & vbCrLf & _
            '                  "           LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))  " & vbCrLf & _
            '                  "           			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD  " & vbCrLf & _
            '                  "                                              ON PRD.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "                                              AND PRD.KanbanNo = PDD.KanbanNo   " & vbCrLf & _
            '                  "                                              AND PRD.SupplierID = PDD.SupplierID   " & vbCrLf & _
            '                  "                                              AND PRD.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "                                              AND PRD.PoNo = PDD.PoNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo  " & vbCrLf & _
            '                  "           LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf & _
            '                  "                                              --AND PDD.SupplierID = PDM.SupplierID   " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                       AND PDD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
            '                  "                                                       AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "                                                       AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
            '                  "                                                       AND PDD.PoNo = RAD.PoNo  " & vbCrLf & _
            '                  "           INNER JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                       AND RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                       --AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "  		 LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf

            'ls_SQL = ls_SQL + "  													AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
            '                  "  													AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
            '                  "  													AND RAD.PONo = IPD.PONo  " & vbCrLf & _
            '                  "  													AND RAD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
            '                  "  		 LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
            '                  "  													AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
            '                  "  													--AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
            '                  "  													 " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "    WHERE  'A' = 'A'   " & vbCrLf & _
            '                  "           AND POM.CommercialCls <> 0  " & vbCrLf & _
            '                  "           AND ISNULL(PDM.suratjalanno,'') <> '' " & vbCrLf

            'ls_SQL = ls_SQL + ls_Filter

            'ls_SQL = ls_SQL + "    ORDER BY h_affiliateorder, H_PasiSJ, h_poorder, h_idxorder, h_kanbanorder ,colpartno   " & vbCrLf & _
            '                  "  " & vbCrLf


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdtfrom") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdtto") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpAll") = "ALL"

            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 5, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_InvoiceDate As Date
        Dim ls_AffiliateCode As String = ""
        Dim ls_AffiliateName As String = ""
        Dim ls_PasiSj As String = ""
        Dim ls_PasiInvoiceno As String = ""
        Dim ls_PaymentTerm As String = ""
        Dim ls_DueDate As String = ""

        Dim ls_Notes As String = ""
        Dim ls_Po As String = ""
        Dim ls_Kanban As String = ""

        With grid
            If e.UpdateValues.Count = 0 Then Exit Sub
            If (e.UpdateValues(0).NewValues("Act").ToString()) = 1 Then
                'ls_DeliveryDate = Trim(e.UpdateValues(0).NewValues("colpasideliverydate").ToString())
                ls_InvoiceDate = "01 Jan 1900"
                ls_AffiliateCode = Trim(e.UpdateValues(0).NewValues("colaffiliatecode").ToString())
                ls_AffiliateName = Trim(e.UpdateValues(0).NewValues("colaffiliatename").ToString())

                ls_PasiSj = "'" & Trim(e.UpdateValues(0).NewValues("colpasisj").ToString()) & "'"
                If Trim(e.UpdateValues(0).NewValues("colPasiInvNo").ToString()) <> "" Then
                    ls_PasiInvoiceno = "'" & Trim(e.UpdateValues(0).NewValues("colPasiInvNo").ToString()) & "'"
                End If
                ls_PaymentTerm = ""
                ls_DueDate = ""
                ls_Notes = ""

                ls_Po = "'" & Trim(e.UpdateValues(0).NewValues("colpono").ToString()) & "'"
                ls_Kanban = "'" & Trim(e.UpdateValues(0).NewValues("colkanbanno").ToString()) & "'"
            End If

            If e.UpdateValues.Count > 1 Then
                For i = 1 To e.UpdateValues.Count - 1
                    If (e.UpdateValues(i).NewValues("Act").ToString()) = 1 Then
                        ls_Po = ls_Po + ",'" & Trim(e.UpdateValues(i).NewValues("colpono").ToString()) & "'"
                        ls_Kanban = ls_Kanban + ",'" & Trim(e.UpdateValues(i).NewValues("colkanbanno").ToString()) & "'"
                        ls_PasiSj = ls_PasiSj + ",'" & Trim(e.UpdateValues(i).NewValues("colpasisj").ToString()) & "'"
                        If ls_PasiInvoiceno <> "" Then
                            ls_PasiInvoiceno = ls_PasiInvoiceno + ",'" & Trim(e.UpdateValues(i).NewValues("colPasiInvNo").ToString()) & "'"
                        Else
                            If Trim(e.UpdateValues(i).NewValues("colPasiInvNo").ToString()) <> "" Then
                                ls_PasiInvoiceno = ls_PasiInvoiceno + ",'" & Trim(e.UpdateValues(i).NewValues("colPasiInvNo").ToString()) & "'"
                            End If
                        End If
                    End If
                Next
            End If
        End With
        Session("POListInv") = ls_InvoiceDate & "|" & ls_AffiliateCode & "|" & ls_AffiliateName & _
                            "|" & ls_PasiSj & "|" & ls_PasiInvoiceno & _
                            "|" & ls_PaymentTerm & "|" & ls_DueDate & "|" & ls_Notes & "|" & ls_Po & "|" & ls_Kanban

    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)

            Select Case pAction
                Case "gridload"
                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "kosong"

            End Select

EndProcedure:
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "coldetail" Or e.DataColumn.FieldName = "Act") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        End If

        If e.DataColumn.FieldName = "Act" Then
            If (e.GetValue("colpartno") <> "") Then
                e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If

        'If (e.GetValue("colkanbanno") = "" Or Left(e.GetValue("colpokanban"), 2) = "NO") Then
        '    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        'End If

        'If (e.DataColumn.FieldName = "coldetail") Then
        '    If (e.GetValue("colpartno") = "") Then
        '        e.Cell.Controls("0").Controls.Clear()
        '    End If
        'End If
    End Sub

    Private Sub grid_HtmlRowPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
        Try
            Dim getRowValues As String = e.GetValue("colaffiliatecode")
            If Not IsNothing(getRowValues) Then
                If getRowValues.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If

        Catch ex As Exception
            
        End Try
    End Sub

    Private Sub btnsubmenu_Click(sender As Object, e As System.EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub
End Class