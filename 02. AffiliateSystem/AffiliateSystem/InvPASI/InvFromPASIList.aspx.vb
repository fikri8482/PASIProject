Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView

Public Class InvFromPASIList
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
                dt3.Text = Format(Now, "01 MMM yyyy")
                dt4.Text = Format(Now, "dd MMM yyyy")
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

        ''SSupplier
        'ls_sql = "SELECT distinct Supplier_Code = '" & clsGlobal.gs_All & "', Supplier_Name = '" & clsGlobal.gs_All & "' from MS_Supplier " & vbCrLf & _
        '         "UNION ALL Select Supplier_Code = RTRIM(SupplierID) ,Supplier_Name = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
        '    sqlConn.Open()

        '    Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
        '    Dim ds As New DataSet
        '    sqlDA.Fill(ds)

        '    With cbosupplier
        '        .Items.Clear()
        '        .Columns.Clear()
        '        .DataSource = ds.Tables(0)
        '        .Columns.Add("Supplier_Code")
        '        .Columns(0).Width = 70
        '        .Columns.Add("Supplier_Name")
        '        .Columns(1).Width = 240
        '        .SelectedIndex = 0
        '        txtaffiliate.Text = clsGlobal.gs_All
        '        .TextField = "Supplier Code"
        '        .DataBind()
        '    End With
        '    sqlConn.Close()
        'End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'ls_SQL = "    SELECT *      " & vbCrLf & _
            '      "     	FROM (     " & vbCrLf & _
            '      "     		  SELECT url = (case when isnull(Finvno,'') = '' then '' else url end) ,coldetail, no = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY FINVNo, FSJ)),     " & vbCrLf & _
            '      "     			     period, affiliatecode, affiliatename,pono, suppliercode, suppliername, pokanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,      " & vbCrLf & _
            '      "     				 kanbanno, suppplandeldate = SupplierPlanDeliveryDate, suppdeldate = supplierdeliverydate, suppsj = suppliersjno, pasideldate = pasideliverydate, pasisj = pasisjno, partno, partname, uom,    " & vbCrLf & _
            '      "     				 suppdelqty, pasirecqty, pasidelqty, affrecqty, affremrecqty, affrecdate, affiliaterecdate, affrecby, pasiinvno, pasiinvdate,     " & vbCrLf & _
            '      "     				 sortpono = PONo, sortkanbanno = KanbanNo, sortheader = 0, Fsuppdeldate, FInvNo, FDiffQty,FSJ, Fsupp, FAff, FPO       " & vbCrLf & _
            '      "     			FROM (     " & vbCrLf & _
            '      "     				  SELECT DISTINCT      " & vbCrLf & _
            '      "     						 period = SUBSTRING(CONVERT(CHAR,IPM.DeliveryDate,106),4,9),     " & vbCrLf & _
            '      "     						 PONo = '',         						  "

            'ls_SQL = ls_SQL + "     						 affiliatecode = ISNULL(KM.AffiliateID,''),     " & vbCrLf & _
            '                  "                              affiliatename = ISNULL(MA.AffiliateName,''),     " & vbCrLf & _
            '                  "     						 SupplierCode = '',     " & vbCrLf & _
            '                  "                              SupplierName ='',     " & vbCrLf & _
            '                  "     						 KanbanCls = ISNULL(POD.KanbanCls,'0'),     " & vbCrLf & _
            '                  "     						 KanbanNo = '',     " & vbCrLf & _
            '                  "     						 SupplierPlanDeliveryDate = '',     " & vbCrLf & _
            '                  "     						 SupplierDeliveryDate = '',     " & vbCrLf & _
            '                  "     						 SupplierSJNO = '',     " & vbCrLf & _
            '                  "     						 PASIDeliveryDate = ISNULL(CONVERT(CHAR,PDM.DeliveryDate,106),''),     " & vbCrLf & _
            '                  "     						 PASISJNo = ISNULL(PDD.SuratJalanNo,''),        						 PartNo = '',     "

            'ls_SQL = ls_SQL + "     						 PartName = '',     " & vbCrLf & _
            '                  "     						 UOM = '',     " & vbCrLf & _
            '                  "     						 suppdelqty = '',     " & vbCrLf & _
            '                  "     						 pasirecqty = '',     " & vbCrLf & _
            '                  "     						 pasidelqty = '',   " & vbCrLf & _
            '                  "     						 affrecqty = '',   " & vbCrLf & _
            '                  "     						 affremrecqty = '', affrecdate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),''), affiliaterecdate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),''),   " & vbCrLf & _
            '                  "     						 affrecby = isnull(RAM.ReceiveBy,''), pasiinvno = isnull(IPM.InvoiceNo,''), pasiinvdate = ISNULL(CONVERT(CHAR,IPM.DeliveryDate,106),''),   " & vbCrLf & _
            '                  "     						 url = 'InvFromPASIDetail.aspx?prm='+Rtrim(ISNULL(CONVERT(CHAR,IPM.DeliveryDate,106),CONVERT(CHAR,GETDATE(),106)))     " & vbCrLf & _
            '                  "     									+ '|' +Rtrim(isnull(POM.affiliateid,''))        									 " & vbCrLf & _
            '                  "     									+ '|' +Rtrim(isnull(MA.Affiliatename,''))     "

            'ls_SQL = ls_SQL + "     									+ '|' +Rtrim(ISNULL(PDM.SuratJalanNo,''))     " & vbCrLf & _
            '                  "     									+ '|' +''     " & vbCrLf & _
            '                  "     									+ '|' +''   " & vbCrLf & _
            '                  "                                      + '|' +Rtrim(ISNULL(IPM.InvoiceNo,'')),   " & vbCrLf & _
            '                  "     						 coldetail = (CASE WHEN IPM.InvoiceNo IS NULL THEN '' ELSE 'DETAIL' END),    " & vbCrLf & _
            '                  "                            Fsuppdeldate = '',   " & vbCrLf & _
            '                  "                            FInvNo = isnull(IPM.InvoiceNo,''),    " & vbCrLf & _
            '                  "                            FDiffQty = 0,   " & vbCrLf & _
            '                  "                            FSJ = ISNULL(PDD.SuratJalanNo,''),                              --FInv = isnull(INVSM.InvoiceNo,''),    " & vbCrLf & _
            '                  "                            Fsupp ='',    " & vbCrLf & _
            '                  "                            FAff = POM.AffiliateID,    "

            'ls_SQL = ls_SQL + "                            FPO = ''  " & vbCrLf & _
            '                  "     FROM    dbo.PO_Master POM    " & vbCrLf & _
            '                  "            LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "                                       AND POM.PoNo = POD.PONo    " & vbCrLf & _
            '                  "                                       AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "                                              AND KD.PoNo = POD.PONo    " & vbCrLf & _
            '                  "                                              AND KD.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "                                              AND KD.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf & _
            '                  "                                              AND KD.KanbanNo = KM.KanbanNo    "

            'ls_SQL = ls_SQL + "                                              AND KD.SupplierID = KM.SupplierID    " & vbCrLf & _
            '                  "                                              AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf & _
            '                  "                                                   AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf & _
            '                  "                                                   AND KD.SupplierID = SDD.SupplierID    " & vbCrLf & _
            '                  "                                                   AND KD.PartNo = SDD.PartNo    " & vbCrLf & _
            '                  "                                                   AND KD.PoNo = SDD.PoNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID    " & vbCrLf & _
            '                  "                                                   AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf & _
            '                  "                                                   AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf & _
            '                  "            INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID    "

            'ls_SQL = ls_SQL + "                                                    AND SDD.KanbanNo = PRD.KanbanNo    " & vbCrLf & _
            '                  "                                                    AND SDD.SupplierID = PRD.SupplierID    " & vbCrLf & _
            '                  "                                                    AND SDD.PartNo = PRD.PartNo  " & vbCrLf & _
            '                  "                                                    AND SDD.PONo = PRD.PONo    " & vbCrLf & _
            '                  "                                                    AND PRD.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
            '                  "                                                    AND PRM.SuratJalanNo = PRD.SuratJalanNo    " & vbCrLf & _
            '                  "            LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))   " & vbCrLf & _
            '                  "            			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD   " & vbCrLf & _
            '                  "                                               ON PRD.AffiliateID = PDD.AffiliateID    " & vbCrLf & _
            '                  "                                               AND PRD.KanbanNo = PDD.KanbanNo    "

            'ls_SQL = ls_SQL + "                                               AND PRD.SupplierID = PDD.SupplierID    " & vbCrLf & _
            '                  "                                               AND PRD.PartNo = PDD.PartNo    " & vbCrLf & _
            '                  "                                               AND PRD.PoNo = PDD.PoNo    " & vbCrLf & _
            '                  "                                               --AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf & _
            '                  "                                               AND PDD.SuratJalanNo = PDM.SuratJalanNo    " & vbCrLf & _
            '                  "                                               --AND PDD.SupplierID = PDM.SupplierID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "                                                        AND PDD.KanbanNo = RAD.KanbanNo   " & vbCrLf & _
            '                  "                                                        AND PDD.SupplierID = RAD.SupplierID   " & vbCrLf & _
            '                  "                                                        AND PDD.PartNo = RAD.PartNo   "

            'ls_SQL = ls_SQL + "                                                        AND PDD.PoNo = RAD.PoNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                        AND RAM.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "                                                        --AND RAM.SupplierID = RAD.SupplierID   " & vbCrLf & _
            '                  "   		 INNER JOIN dbo.PLPASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
            '                  "   													AND RAD.KanbanNo = IPD.KanbanNo   " & vbCrLf & _
            '                  "   													AND RAD.PartNo = IPD.PartNo   " & vbCrLf & _
            '                  "   													AND RAD.PONo = IPD.PONo   " & vbCrLf & _
            '                  "   													AND IPD.SuratJalanNo = RAD.SuratJalanNo    " & vbCrLf & _
            '                  "   		 INNER JOIN dbo.PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "   													AND IPD.SuratJalanNo = IPM.SuratJalanNo   " & vbCrLf & _
            '                  "   													  " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "   WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' -- AND POD.pono='PO-DEE-KMK '    " & vbCrLf & _
            '                  "                 --WHERE POD.PONO <> '' " & vbCrLf

            'If checkbox1.Checked = True Then
            '    ls_SQL = ls_SQL + " AND IPM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "' " & vbCrLf
            'End If

            'If checkbox2.Checked = True Then
            '    ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(IPM.DeliveryDate,'')),106) between '" & Format(dt3.Value, "dd MMM yyyy") & "' AND '" & Format(dt4.Value, "dd MMM yyyy") & "' " & vbCrLf
            'End If

            'If rbinvoice.Value = "YES" Then
            '    ls_SQL = ls_SQL + "AND isnull(IPM.InvoiceNo, '') <> '' " & vbCrLf
            'ElseIf rbinvoice.Value = "NO" Then
            '    ls_SQL = ls_SQL + " AND isnull(IPM.InvoiceNo,'') = '' " & vbCrLf
            'End If

            'If txtsj.Text <> "" Then
            '    ls_SQL = ls_SQL + " AND IPM.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            'End If

            'If txtpono.Text <> "" Then
            '    ls_SQL = ls_SQL + " AND IPD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
            'End If


            'ls_SQL = ls_SQL + " )hdr  " & vbCrLf & _
            '      "    UNION ALL    " & vbCrLf & _
            '      "    SELECT DISTINCT url = '',coldetail = '',no = '',     " & vbCrLf & _
            '      "     				 period = '',       " & vbCrLf & _
            '      "     				 affiliatecode = '',        				 affiliatename = '',     " & vbCrLf & _
            '      "     				 pono = POM.PONo,    " & vbCrLf & _
            '      "     				 suppliercode = PDD.SupplierID,     " & vbCrLf & _
            '      "     				 suppliername = MS.SupplierName,     " & vbCrLf & _
            '      "     				 kanbancls = '',     " & vbCrLf & _
            '      "     				 kanbanno = KM.KanbanNo,     " & vbCrLf & _
            '      "     				 suppplandeldate = KM.KanbanDate,     "

            'ls_SQL = ls_SQL + "     				 suppdeldate = '',     " & vbCrLf & _
            '                  "     				 suppsj = '',     " & vbCrLf & _
            '                  "     				 pasideldate = '',     " & vbCrLf & _
            '                  "     				 pasisj = '',        				  " & vbCrLf & _
            '                  "     				 partno = POD.PartNo,     " & vbCrLf & _
            '                  "     				 partname = MP.PartName,     " & vbCrLf & _
            '                  "     				 uom = UC.Description,     " & vbCrLf & _
            '                  "     				 suppdelqty = '', --Convert(char,Round(CONVERT(CHAR,Round(ISNULL(SDD.DOQty,0),0)),0)),     " & vbCrLf & _
            '                  "     				 pasirecqty = '', --Convert(char,Round(convert(char,Round(Isnull(PRD.GoodRecQty,0),0)),0)),    " & vbCrLf & _
            '                  "     				 pasidelqty = '', --Convert(char,Round(convert(char,Round(Isnull(PDD.DOQty,0),0)),0)),   " & vbCrLf & _
            '                  "     				 affrecqty = Convert(char,Round(convert(char,Round(Isnull(RAD.RecQty,0),0)),0)),   "

            'ls_SQL = ls_SQL + "     				 affremrecqty = '',   " & vbCrLf & _
            '                  "     				 affrecdate = '', affiliaterecdate = '', affrecby = '', pasiinvno = '', pasiinvdate = '',  " & vbCrLf & _
            '                  "     				 sortpono = POD.PONo, sortkanbanno = ISNULL(KD.KanbanNo,''), sortheader = 1,     " & vbCrLf & _
            '                  "                    Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),                      " & vbCrLf & _
            '                  "                    FInvNo = isnull(IPM.InvoiceNo,''),    " & vbCrLf & _
            '                  "                    FDiffQty = 0,   " & vbCrLf & _
            '                  "                    FSJ = ISNULL(PDD.SuratJalanNo,''),    " & vbCrLf & _
            '                  "                    --FInv = isnull(INVSM.InvoiceNo,''),    " & vbCrLf & _
            '                  "                    Fsupp = POM.SupplierID,    " & vbCrLf & _
            '                  "                    FAff = POM.AffiliateID,    " & vbCrLf & _
            '                  "                    FPO = POD.PONo   "

            'ls_SQL = ls_SQL + "     FROM    dbo.PO_Master POM    " & vbCrLf & _
            '                  "            LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "                                       AND POM.PoNo = POD.PONo    " & vbCrLf & _
            '                  "                                       AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID    " & vbCrLf & _
            '                  "                                              AND KD.PoNo = POD.PONo    " & vbCrLf & _
            '                  "                                              AND KD.SupplierID = POD.SupplierID    " & vbCrLf & _
            '                  "                                              AND KD.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID    " & vbCrLf & _
            '                  "                                              AND KD.KanbanNo = KM.KanbanNo    " & vbCrLf & _
            '                  "                                              AND KD.SupplierID = KM.SupplierID    "

            'ls_SQL = ls_SQL + "                                              AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID    " & vbCrLf & _
            '                  "                                                   AND KD.KanbanNo = SDD.KanbanNo    " & vbCrLf & _
            '                  "                                                   AND KD.SupplierID = SDD.SupplierID    " & vbCrLf & _
            '                  "                                                   AND KD.PartNo = SDD.PartNo    " & vbCrLf & _
            '                  "                                                   AND KD.PoNo = SDD.PoNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID    " & vbCrLf & _
            '                  "                                                   AND SDM.SuratJalanNo = SDD.SuratJalanNo    " & vbCrLf & _
            '                  "                                                   AND SDM.SupplierID = SDD.SupplierID    " & vbCrLf & _
            '                  "            INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
            '                  "                                                    AND SDD.KanbanNo = PRD.KanbanNo    "

            'ls_SQL = ls_SQL + "                                                    AND SDD.SupplierID = PRD.SupplierID    " & vbCrLf & _
            '                  "                                                    AND SDD.PartNo = PRD.PartNo  " & vbCrLf & _
            '                  "                                                    AND SDD.PONo = PRD.PONo    " & vbCrLf & _
            '                  "                                                    AND PRD.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID    " & vbCrLf & _
            '                  "                                                    AND PRM.SuratJalanNo = PRD.SuratJalanNo    " & vbCrLf & _
            '                  "            LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))   " & vbCrLf & _
            '                  "            			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD   " & vbCrLf & _
            '                  "                                               ON PRD.AffiliateID = PDD.AffiliateID    " & vbCrLf & _
            '                  "                                               AND PRD.KanbanNo = PDD.KanbanNo    " & vbCrLf & _
            '                  "                                               AND PRD.SupplierID = PDD.SupplierID    "

            'ls_SQL = ls_SQL + "                                               AND PRD.PartNo = PDD.PartNo    " & vbCrLf & _
            '                  "                                               AND PRD.PoNo = PDD.PoNo    " & vbCrLf & _
            '                  "                                               --AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo   " & vbCrLf & _
            '                  "            LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID    " & vbCrLf & _
            '                  "                                               AND PDD.SuratJalanNo = PDM.SuratJalanNo    " & vbCrLf & _
            '                  "                                               --AND PDD.SupplierID = PDM.SupplierID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "                                                        AND PDD.KanbanNo = RAD.KanbanNo   " & vbCrLf & _
            '                  "                                                        AND PDD.SupplierID = RAD.SupplierID   " & vbCrLf & _
            '                  "                                                        AND PDD.PartNo = RAD.PartNo   " & vbCrLf & _
            '                  "                                                        AND PDD.PoNo = RAD.PoNo   "

            'ls_SQL = ls_SQL + "            LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo   " & vbCrLf & _
            '                  "                                                        AND RAM.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "                                                        --AND RAM.SupplierID = RAD.SupplierID   " & vbCrLf & _
            '                  "   		 INNER JOIN dbo.PLPASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
            '                  "   													AND RAD.KanbanNo = IPD.KanbanNo   " & vbCrLf & _
            '                  "   													AND RAD.PartNo = IPD.PartNo   " & vbCrLf & _
            '                  "   													AND RAD.PONo = IPD.PONo   " & vbCrLf & _
            '                  "   													AND RAD.SuratJalanNo = PDD.SuratJalanNo    " & vbCrLf & _
            '                  "   		 INNER JOIN dbo.PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID   " & vbCrLf & _
            '                  "   													AND IPD.SuratJalanNo = IPM.SuratJalanNo   " & vbCrLf

            'ls_SQL = ls_SQL + "   													  " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf & _
            '                  "            LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
            '                  "     WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' --AND POD.pono='PO-DEE-KMK' --'PO20150501-KMK'    " & vbCrLf & _
            '                  "   		   --WHERE POD.PONO <> '' " & vbCrLf & _
            '                  "  "

            'If checkbox1.Checked = True Then
            '    ls_SQL = ls_SQL + " AND IPM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "' " & vbCrLf
            'End If

            'If checkbox2.Checked = True Then
            '    ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(IPM.DeliveryDate,'')),106) between '" & Format(dt3.Value, "dd MMM yyyy") & "' AND '" & Format(dt4.Value, "dd MMM yyyy") & "' " & vbCrLf
            'End If

            'If rbinvoice.Value = "YES" Then
            '    ls_SQL = ls_SQL + "AND isnull(IPM.InvoiceNo, '') <> '' " & vbCrLf
            'ElseIf rbinvoice.Value = "NO" Then
            '    ls_SQL = ls_SQL + " AND isnull(IPM.InvoiceNo,'') = '' " & vbCrLf
            'End If

            'If txtsj.Text <> "" Then
            '    ls_SQL = ls_SQL + " AND IPM.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            'End If


            'If txtpono.Text <> "" Then
            '    ls_SQL = ls_SQL + " AND IPD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
            'End If

            'ls_SQL = ls_SQL + "    ) SPDC ORDER BY FINVNo, FSJ, SortHeader    " & vbCrLf & _
            '                  "  "

            ls_SQL = " SELECT no = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY pasiinvno)), * FROM " & vbCrLf & _
                      " (SELECT DISTINCT  	" & vbCrLf & _
                      " 	url = 'InvFromPASIDetail.aspx?prm='+Rtrim(ISNULL(CONVERT(CHAR,PLM.DeliveryDate,106),CONVERT(CHAR,GETDATE(),106))) " & vbCrLf & _
                      "      									+ '|' +Rtrim(isnull(PLM.affiliateid,'')) " & vbCrLf & _
                      "      									+ '|' +Rtrim(isnull(MA.Affiliatename,'')) " & vbCrLf & _
                      " 										+ '|' +Rtrim(ISNULL(PLM.SuratJalanNo,'')) " & vbCrLf & _
                      "      									+ '|' +'' " & vbCrLf & _
                      "      									+ '|' +'' " & vbCrLf & _
                      "                                       + '|' +Rtrim(ISNULL(PLM.InvoiceNo,'')), " & vbCrLf & _
                      " 	coldetail = (CASE WHEN PLM.InvoiceNo IS NULL THEN '' ELSE 'DETAIL' END), " & vbCrLf & _
                      " 	period = SUBSTRING(CONVERT(CHAR,PLM.DeliveryDate,106),4,9), "

            ls_SQL = ls_SQL + " 	affiliatecode = ISNULL(PLM.AffiliateID,''), " & vbCrLf & _
                              " 	pasiinvno = isnull(PLM.InvoiceNo,''),  " & vbCrLf & _
                              " 	pasiinvdate = ISNULL(CONVERT(CHAR,PLM.DeliveryDate,106),''), " & vbCrLf & _
                              " 	pasisj = ISNULL(PLM.SuratJalanNo,''), " & vbCrLf & _
                              " 	pasideldate = ISNULL(CONVERT(CHAR,PLM.DeliveryDate,106),''), " & vbCrLf & _
                              " 	affrecdate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),''), " & vbCrLf & _
                              " 	affrecby = isnull(RAM.ReceiveBy,'') " & vbCrLf & _
                              " FROM PLPASI_Master PLM " & vbCrLf & _
                              " LEFT JOIN PLPASI_Detail PLD ON PLD.AffiliateID = PLM.AffiliateID and PLD.SuratJalanNo = PLM.SuratJalanNo " & vbCrLf & _
                              " LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = PLM.AffiliateID and RAM.SuratJalanNo = PLM.SuratJalanNo " & vbCrLf & _
                              " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PLM.AffiliateID "

            ls_SQL = ls_SQL + " WHERE PLM.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                              "  "

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND (PLM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "') " & vbCrLf
            End If

            If checkbox2.Checked = True Then
                ls_SQL = ls_SQL + " AND (PLM.DeliveryDate between '" & Format(dt3.Value, "yyyy-MM-dd") & "' AND '" & Format(dt4.Value, "yyyy-MM-dd") & "') " & vbCrLf
            End If

            If rbinvoice.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(PLM.InvoiceNo, '') <> '' " & vbCrLf
            ElseIf rbinvoice.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(PLM.InvoiceNo,'') = '' " & vbCrLf
            End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND PLM.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND PLD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
            End If

            ls_SQL = ls_SQL + "    ) XYZ    " & vbCrLf & _
                              "  "

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
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, true, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
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

    Private Sub grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
        If e.RowType <> GridViewRowType.Data Then Return
        'If e.GetValue("partno").ToString = "" Then e.Row.BackColor = Drawing.Color.LightGray
    End Sub
End Class