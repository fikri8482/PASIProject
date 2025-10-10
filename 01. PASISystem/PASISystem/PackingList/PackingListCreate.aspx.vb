Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing

Public Class PackingListCreate
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

#End Region

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
                ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(PDD.DOQty,'')),106) between '" & Format(dtfrom.Value, "dd MMM yyyy") & "' and '" & Format(dtto.Value, "dd MMM yyyy") & "'" & vbCrLf
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
                  "         coldetail = 'PackingListCreateDetail.aspx?prm=' " & vbCrLf & _
                  "         + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) " & vbCrLf & _
                  "         + '|' + RTRIM(POM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|' " & vbCrLf & _
                  "         + RTRIM(CASE WHEN ISNULL(PDM.SuratJalanNo, '')='' THEN ISNULL(SDM.SuratJalanNo, '') ELSE ISNULL(PDM.SuratJalanNo, '') END) + '|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(KM.DeliveryLocationCode, '')) + '|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(MD.DeliveryLocationName, '')) + '|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(PDM.DriverName, '')) + '|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(PDM.DriverContact, '')) + '|' + RTRIM(ISNULL(PDM.NoPol, " & vbCrLf & _
                  "                                                               '')) + '|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(PDM.JenisArmada, '')) + '|''' " & vbCrLf & _
                  "         + RTRIM(ISNULL(POM.PONo, '')) + '''|'''   " & vbCrLf & _
                  "         + RTRIM(ISNULL(KD.KanbanNo, '')) + '''|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(POM.SupplierID, '')) + '|' " & vbCrLf & _
                  "         + RTRIM(ISNULL(MS.SupplierName, '')) + '|' " & vbCrLf & _
                  "         --+ RTRIM(ISNULL(PDD.SuratJalanNo, '')), " & vbCrLf & _
                  "         + RTRIM(CASE WHEN ISNULL(PDM.SuratJalanNo, '')='' THEN ISNULL(SDM.SuratJalanNo, '') ELSE ISNULL(PDM.SuratJalanNo, '') END), " & vbCrLf & _
                  "         coldetailname = CASE WHEN ISNULL(PDM.SuratJalanNo, '') = '' " & vbCrLf & _
                  "                              THEN '' " & vbCrLf & _
                  "                              ELSE 'DETAIL' " & vbCrLf & _
                  "                         END , " & vbCrLf

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
                              "                                            --AND PDD.SupplierID = PDM.SupplierID " & vbCrLf & _
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

    Private Sub up_GridLoad() 'update by dianiswari (Summary DOPASI)
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'Supplier Already Deliver
            If rbdeliver.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(PDD.DOQty, 0) = 0 " & vbCrLf
            ElseIf rbdeliver.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(PDD.DOQty,0) <> 0 " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_Filter = ls_Filter + " AND PDM.SuratJalanNo LIKE '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If checkbox2.Checked = True Then
                'ls_Filter = ls_Filter + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(PDD.DOQty,'')),106) between '" & Format(dtfrom.Value, "dd MMM yyyy") & "' and '" & Format(dtto.Value, "dd MMM yyyy") & "'" & vbCrLf
                ls_Filter = ls_Filter + " AND PDM.DeliveryDate between '" & Format(dtfrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtto.Value, "yyyy-MM-dd") & "'" & vbCrLf
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND PDM.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            If cbopart.Text <> clsGlobal.gs_All And cbopart.Text <> "" Then
                ls_Filter = ls_Filter + "AND PDD.PartNo = '" & Trim(cbopart.Text) & "'" & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_Filter = ls_Filter + "and PDD.PONo LIKE '%" & txtpono.Text & "%'" & vbCrLf
            End If

            If rbPasiDelivery.Value = "YES" Then
                ls_Filter = ls_Filter + "AND isnull(PDD.SuratJalanNo, '') <> '' " & vbCrLf
            ElseIf rbPasiDelivery.Value = "NO" Then
                ls_Filter = ls_Filter + " AND isnull(PDD.SuratJalanNo,'') = '' " & vbCrLf
            End If

            ls_SQL = "  SELECT   " & vbCrLf & _
                      "   Act, coldetail , coldetailname , colno = CONVERT(char,ROW_NUMBER() OVER(ORDER BY h_affiliateorder, HPasiSJ, h_poorder, h_idxorder, h_kanbanorder ,colpartno DESC)) ,     " & vbCrLf & _
                      "   colperiod, colaffiliatecode, colaffiliatename, coldeliverylocationcode, coldeliverylocationname, colpono,    " & vbCrLf & _
                      "   colsuppliercode, colsuppliername, colpokanban, colkanbanno, colplandeldate, coldeldate, colsj,   " & vbCrLf & _
                      "   colpasideliverydate, colpasisj = ISNULL(colpasisj,''), colpartno, colpartname, coluom, coldeliveryqty, colreceiveqty, coldefect,   " & vbCrLf & _
                      "   colremaining, colpasideliveryqty, coldeliverydate, coldeliveryby, H_POORDER, H_IDXORDER, H_KANBANORDER, H_AFFILIATEORDER,   " & vbCrLf & _
                      "   H_KANBANCLS, HSupSJ, HPasiSJ    " & vbCrLf & _
                      "   FROM (    " & vbCrLf

            ls_SQL = ls_SQL + " SELECT  DISTINCT     " & vbCrLf & _
                      " 					Act = 0 ,     " & vbCrLf & _
                      " 				   coldetail = 'PackingListCreateDetail.aspx?prm='    " & vbCrLf & _
                      " 				   + CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106)    " & vbCrLf & _
                      " 				   + '|' + RTRIM(PDM.AffiliateID) + '|' + RTRIM(MA.AffiliateName) + '|'    " & vbCrLf & _
                      " 				   + RTRIM(PDM.SuratJalanNo) + '|'    " & vbCrLf & _
                      " 				   + '' + '|'    " & vbCrLf & _
                      " 				   + '' + '|'    " & vbCrLf & _
                      " 				   + RTRIM(ISNULL(PDM.DriverName, '')) + '|'    " & vbCrLf & _
                      " 				   + RTRIM(ISNULL(PDM.DriverContact, '')) + '|' + RTRIM(ISNULL(PDM.NoPol,    " & vbCrLf & _
                      " 																		 '')) + '|'    " & vbCrLf

            ls_SQL = ls_SQL + " 				   + RTRIM(ISNULL(PDM.JenisArmada, '')) + '|'''    " & vbCrLf & _
                              " 				   + '' + '''|'''      " & vbCrLf & _
                              " 				   + '' + '''|'    " & vbCrLf & _
                              " 				   + '' + '|'    " & vbCrLf & _
                              " 				   + '' + '|'    " & vbCrLf & _
                              " 				   --+ RTRIM(ISNULL(PDD.SuratJalanNo, '')),    " & vbCrLf & _
                              " 				   + RTRIM(PDM.SuratJalanNo),    " & vbCrLf & _
                              " 				   coldetailname = CASE WHEN ISNULL(PDM.SuratJalanNo, '') = ''    " & vbCrLf & _
                              " 										THEN ''    " & vbCrLf & _
                              " 										ELSE 'DETAIL'    " & vbCrLf & _
                              " 								   END ,    " & vbCrLf

            ls_SQL = ls_SQL + " 					colno = '' ,     " & vbCrLf & _
                              " 					colperiod = RIGHT(CONVERT(CHAR(11),CONVERT(DATETIME,PDM.DeliveryDate),106), 8),     " & vbCrLf & _
                              " 					colaffiliatecode = PDM.AffiliateID ,     " & vbCrLf & _
                              " 					colaffiliatename = MA.AffiliateName ,     " & vbCrLf & _
                              " 					coldeliverylocationcode = '' ,     " & vbCrLf & _
                              " 					coldeliverylocationname = '' ,     " & vbCrLf & _
                              " 					colpono = '' ,     " & vbCrLf & _
                              " 					colsuppliercode = '' ,   " & vbCrLf & _
                              " 					colsuppliername = '' ,     " & vbCrLf & _
                              " 					colpokanban = '' ,     " & vbCrLf & _
                              " 					colkanbanno = '',     " & vbCrLf

            ls_SQL = ls_SQL + " 					colplandeldate = '' ,     " & vbCrLf & _
                              " 					coldeldate = '' ,                    " & vbCrLf & _
                              " 					colsj = '' ,     " & vbCrLf & _
                              " 					colpasideliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,     " & vbCrLf & _
                              " 					colpasisj = PDM.SuratJalanNo ,            colpartno = '' ,     " & vbCrLf & _
                              " 					colpartname = '' ,     " & vbCrLf & _
                              " 					coluom = '' ,     " & vbCrLf & _
                              " 					coldeliveryqty = '' ,     " & vbCrLf & _
                              " 					colreceiveqty = '' ,     " & vbCrLf & _
                              " 					coldefect = '' ,     " & vbCrLf & _
                              " 					colremaining = '' ,     " & vbCrLf

            ls_SQL = ls_SQL + " 					colpasideliveryqty = '' ,     " & vbCrLf & _
                              " 					coldeliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END,     " & vbCrLf & _
                              " 					coldeliveryby = '" & Session("UserID") & "' ,     " & vbCrLf & _
                              " 					'' H_POORDER ,            H_IDXORDER = 0 ,     " & vbCrLf & _
                              " 					H_KANBANORDER = '' ,     " & vbCrLf & _
                              " 					H_AFFILIATEORDER = PDM.AffiliateID ,     " & vbCrLf & _
                              " 					H_KANBANCLS = '',     " & vbCrLf & _
                              " 					HSupSJ = '', --SDM.SuratJalanNo,    " & vbCrLf & _
                              " 					HPasiSJ = isnull(PDM.SuratJalanNo,'') " & vbCrLf & _
                              " 		FROM DOPASI_Master PDM " & vbCrLf & _
                              "         INNER JOIN DOPASI_Detail PDD ON PDM.AffiliateID = PDD.AffiliateID and PDM.SuratJalanNo = PDD.SuratJalanNo and pdm.SupplierID=pdd.SupplierID" & vbCrLf & _
                              " 		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                              "  WHERE   'A' = 'A'    " & vbCrLf & _
                              "  AND PDM.SuratJalanNo <> ''  " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " )Header " & vbCrLf

            ls_SQL = ls_SQL + " UNION ALL  " & vbCrLf & _
                                " SELECT   DISTINCT  " & vbCrLf & _
                                  " 			   Act = 0 ,    " & vbCrLf & _
                                  " 			   coldetail = '' ,    " & vbCrLf & _
                                  " 			   coldetailname = '' ,    " & vbCrLf & _
                                  " 			   colno = '' ,    " & vbCrLf & _
                                  " 			   colperiod = '' ,    " & vbCrLf & _
                                  " 			   colaffiliatecode = '' ,    " & vbCrLf & _
                                  " 			   colaffiliatename = '' ,    " & vbCrLf & _
                                  " 			   coldeliverylocationcode = '' ,    " & vbCrLf & _
                                  " 			   coldeliverylocationname = '' ,    " & vbCrLf & _
                                  " 			   colpono = PDD.PONo ,     "

            ls_SQL = ls_SQL + " 			   colsuppliercode = PDD.SupplierID ,   " & vbCrLf & _
                              " 			   colsuppliername = '' ,     " & vbCrLf & _
                              " 			   colpokanban = '' ,   " & vbCrLf & _
                              " 			   colkanbanno = PDD.KanbanNo,    " & vbCrLf & _
                              " 			   colplandeldate = '',     " & vbCrLf & _
                              " 			   coldeldate = '',  " & vbCrLf & _
                              " 			   colsj = '',  " & vbCrLf & _
                              " 			   colpasideliverydate = '' ,    " & vbCrLf & _
                              " 			   colpasisj = '' ,    " & vbCrLf & _
                              " 			   colpartno = PDD.PartNo ,    " & vbCrLf & _
                              " 			   colpartname = MP.PartName ,    "

            ls_SQL = ls_SQL + " 			   coluom = UC.Description ,    " & vbCrLf & _
                              " 			   coldeliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),    " & vbCrLf & _
                              " 			   colreceiveqty = '',   " & vbCrLf & _
                              " 			   coldefect = '',  " & vbCrLf & _
                              " 			   colremaining = '',    " & vbCrLf & _
                              " 			   colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PDD.DOQty,0)))),     " & vbCrLf & _
                              " 			   coldeliverydate = CASE WHEN ISNULL(PDM.DeliveryDate,'') = '' THEN '' ELSE  CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(PDM.DeliveryDate, '')), 106) END ,    " & vbCrLf & _
                              " 			   coldeliveryby = '" & Session("UserID") & "' ,    " & vbCrLf & _
                              " 			   POOrder = PDD.PONo ,    " & vbCrLf & _
                              " 			   idxorder = 1 ,    " & vbCrLf & _
                              " 			   kanbanorder = ISNULL(PDD.KanbanNo, '-') ,    "

            ls_SQL = ls_SQL + " 			   affiliateorder = PDM.AffiliateID ,    " & vbCrLf & _
                              " 			   '' KanbanCls,    " & vbCrLf & _
                              " 			   HSupSJ = pdd.SuratJalanNoSupplier,    " & vbCrLf & _
                              " 			   HPasiSJ = isnull(PDM.SuratJalanNo,'')    " & vbCrLf & _
                              " 	FROM DOPASI_Master PDM " & vbCrLf & _
                              " 	INNER JOIN DOPASI_Detail PDD ON PDM.AffiliateID = PDD.AffiliateID and PDM.SuratJalanNo = PDD.SuratJalanNo and pdm.SupplierID=pdd.SupplierID" & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = PDD.PartNo " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
                              "    WHERE   'A' = 'A'   " & vbCrLf & _
                              "    AND PDM.SuratJalanNo <> '' " & vbCrLf

            ls_SQL = ls_SQL + ls_Filter

            ls_SQL = ls_SQL + " ORDER BY h_affiliateorder, HPasiSJ, h_poorder, h_idxorder, h_kanbanorder ,colpartno "

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

#Region "Control Events"
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

        Session.Remove("POList")
        Session.Remove("KanbanList")

        With grid
            If e.UpdateValues.Count = 0 Then Exit Sub
            If (e.UpdateValues(0).NewValues("Act").ToString()) = 1 Then
                ls_DeliveryDate = "01 Jan 1900"
                ls_AffiliateCode = Trim(e.UpdateValues(0).NewValues("colaffiliatecode").ToString())
                ls_AffiliateName = Trim(e.UpdateValues(0).NewValues("colaffiliatename").ToString())
                ls_SuratJalanNo = ""
                ls_DeliveryCode = Trim(e.UpdateValues(0).NewValues("coldeliverylocationcode").ToString())
                ls_DeliveryName = Trim(e.UpdateValues(0).NewValues("coldeliverylocationname").ToString())
                ls_PO = "'" & Trim(e.UpdateValues(0).NewValues("colpono").ToString()) & "'"
                ls_Kanban = "'" & Trim(e.UpdateValues(0).NewValues("colkanbanno").ToString()) & "'"
                ls_Supplier = Trim(e.UpdateValues(0).NewValues("colsuppliercode").ToString())
                ls_SupplierName = Trim(e.UpdateValues(0).NewValues("colsuppliername").ToString())
            End If

            If e.UpdateValues.Count > 1 Then
                For i = 1 To e.UpdateValues.Count - 1
                    If (e.UpdateValues(i).NewValues("Act").ToString()) = 1 Then
                        ls_PO = ls_PO + ",'" & Trim(e.UpdateValues(i).NewValues("colpono").ToString()) & "'"
                        ls_Kanban = ls_Kanban + ",'" & Trim(e.UpdateValues(i).NewValues("colkanbanno").ToString()) & "'"
                    End If
                Next
            End If
        End With
        Session("POList") = ls_DeliveryDate & "|" & ls_AffiliateCode & "|" & ls_AffiliateName & _
                            "|" & ls_SuratJalanNo & "|" & ls_DeliveryCode & _
                            "|" & ls_DeliveryName & "|" & "||||" & ls_PO & "||" & ls_Supplier & "|" & ls_SupplierName & "|" & ls_SuratJalan
        Session("KanbanList") = ls_Kanban
        HF.Set("Update", "1")
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
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

            Select Case pAction
                Case "gridload"
                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "update"
                    grid.UpdateEdit()
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
            If (e.GetValue("colpasisj") = "") Then
                e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If

        'If (e.GetValue("colkanbanno") = "" Or Left(e.GetValue("colpokanban"), 2) = "NO") Then
        '    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        'End If

        If (e.DataColumn.FieldName = "coldetail") Then
            If (e.GetValue("colpokanban") = "NO") Then
                e.Cell.Controls("0").Controls.Clear()
            End If
        End If

        If e.DataColumn.FieldName = "colremaining" Then
            'If (Trim(e.GetValue("coldetailname")) = "") Then
            '    If (e.GetValue("coldeliveryqty") > (CDbl(e.GetValue("colreceiveqty")) + CDbl(e.GetValue("coldefect")))) Then
            '        e.Cell.BackColor = Color.Fuchsia
            '    End If
            'End If
        End If
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
            'Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'Session("E01Msg") = lblInfo.Text
        End Try
    End Sub

    Private Sub grid_PageIndexChanged(sender As Object, e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

End Class