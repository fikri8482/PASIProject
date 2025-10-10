Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView

Public Class InvFromPASIListExport
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
                'Call up_fillcombo()
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

    '#Region "PROCEDURE"
    '    Private Sub up_fillcombo()
    '        Dim ls_sql As String

    '        ls_sql = ""
    '        'SAffiliate
    '        ls_sql = "SELECT distinct Affiliate_Code = '" & clsGlobal.gs_All & "', Affiliate_Name = '" & clsGlobal.gs_All & "' from MS_AFfiliate " & vbCrLf & _
    '                 "UNION ALL Select Affiliate_Code = RTRIM(AffiliateID) ,Affiliate_Name = RTRIM(Affiliatename) FROM MS_Affiliate " & vbCrLf
    '        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '            sqlConn.Open()

    '            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
    '            Dim ds As New DataSet
    '            sqlDA.Fill(ds)

    '            With cboaffiliate
    '                .Items.Clear()
    '                .Columns.Clear()
    '                .DataSource = ds.Tables(0)
    '                .Columns.Add("Affiliate_Code")
    '                .Columns(0).Width = 70
    '                .Columns.Add("Affiliate_Name")
    '                .Columns(1).Width = 240
    '                .SelectedIndex = 0
    '                txtaffiliate.Text = clsGlobal.gs_All
    '                .TextField = "Affiliate Code"
    '                .DataBind()
    '            End With
    '            sqlConn.Close()
    '        End Using

    '        ''SSupplier
    '        'ls_sql = "SELECT distinct Supplier_Code = '" & clsGlobal.gs_All & "', Supplier_Name = '" & clsGlobal.gs_All & "' from MS_Supplier " & vbCrLf & _
    '        '         "UNION ALL Select Supplier_Code = RTRIM(SupplierID) ,Supplier_Name = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
    '        'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        '    sqlConn.Open()

    '        '    Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
    '        '    Dim ds As New DataSet
    '        '    sqlDA.Fill(ds)

    '        '    With cbosupplier
    '        '        .Items.Clear()
    '        '        .Columns.Clear()
    '        '        .DataSource = ds.Tables(0)
    '        '        .Columns.Add("Supplier_Code")
    '        '        .Columns(0).Width = 70
    '        '        .Columns.Add("Supplier_Name")
    '        '        .Columns(1).Width = 240
    '        '        .SelectedIndex = 0
    '        '        txtaffiliate.Text = clsGlobal.gs_All
    '        '        .TextField = "Supplier Code"
    '        '        .DataBind()
    '        '    End With
    '        '    sqlConn.Close()
    '        'End Using
    '    End Sub

    '    Private Sub up_GridLoad()
    '        Dim ls_SQL As String = ""
    '        Dim pWhere As String = ""

    '        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '            sqlConn.Open()
    '            ls_SQL = "   SELECT *     " & vbCrLf & _
    '                  "    	FROM (    " & vbCrLf & _
    '                  "    		  SELECT url = (case when isnull(Finvno,'') = '' then '' else url end) ,coldetail, no = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, KanbanCls, KanbanNo)),    " & vbCrLf & _
    '                  "    			     period, affiliatecode, affiliatename,pono, suppliercode, suppliername, pokanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,     " & vbCrLf & _
    '                  "    				 kanbanno, suppplandeldate = SupplierPlanDeliveryDate, suppdeldate = supplierdeliverydate, suppsj = suppliersjno, pasideldate = pasideliverydate, pasisj = pasisjno, partno, partname, uom,   " & vbCrLf & _
    '                  "    				 suppdelqty, pasirecqty, pasidelqty, affrecqty, affremrecqty, affrecdate, affiliaterecdate, affrecby, pasiinvno, pasiinvdate,    " & vbCrLf & _
    '                  "    				 sortpono = PONo, sortkanbanno = KanbanNo, sortheader = 0, Fsuppdeldate, FInvNo, FDiffQty,FSJ, FInv, Fsupp, FAff, FPO      " & vbCrLf & _
    '                  "    			FROM (    " & vbCrLf & _
    '                  "    				  SELECT DISTINCT     " & vbCrLf & _
    '                  "    						 period = SUBSTRING(CONVERT(CHAR,POM.period,106),4,9),    " & vbCrLf & _
    '                  "    						 PONo = POD.PONo,     "

    '            ls_SQL = ls_SQL + "    						 affiliatecode = ISNULL(KM.AffiliateID,''),    " & vbCrLf & _
    '                              "                          affiliatename = ISNULL(MA.AffiliateName,''),    " & vbCrLf & _
    '                              "    						 SupplierCode = POM.SupplierID,    " & vbCrLf & _
    '                              "                          SupplierName = MS.SupplierName,    " & vbCrLf & _
    '                              "    						 KanbanCls = ISNULL(POD.KanbanCls,'0'),    " & vbCrLf & _
    '                              "    						 KanbanNo = ISNULL(KD.KanbanNo,''),    " & vbCrLf & _
    '                              "    						 SupplierPlanDeliveryDate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),    " & vbCrLf & _
    '                              "    						 SupplierDeliveryDate = ISNULL(CONVERT(CHAR,DSM.DeliveryDate,106),''),    " & vbCrLf & _
    '                              "    						 SupplierSJNO = ISNULL(DSD.SuratJalanNo,''),    " & vbCrLf & _
    '                              "    						 PASIDeliveryDate = ISNULL(CONVERT(CHAR,DPM.DeliveryDate,106),''),    " & vbCrLf & _
    '                              "    						 PASISJNo = ISNULL(DPD.SuratJalanNo,''),    "

    '            ls_SQL = ls_SQL + "    						 PartNo = '',    " & vbCrLf & _
    '                              "    						 PartName = '',    " & vbCrLf & _
    '                              "    						 UOM = '',    " & vbCrLf & _
    '                              "    						 suppdelqty = '',    " & vbCrLf & _
    '                              "    						 pasirecqty = '',    " & vbCrLf & _
    '                              "    						 pasidelqty = '',  " & vbCrLf & _
    '                              "    						 affrecqty = '',  " & vbCrLf & _
    '                              "    						 affremrecqty = '', affrecdate = '', affiliaterecdate = ISNULL(CONVERT(CHAR,RAM.ReceiveDate,106),''),  " & vbCrLf & _
    '                              "    						 affrecby = isnull(RAM.ReceiveBy,''), pasiinvno = isnull(IPM.InvoiceNo,''), pasiinvdate = ISNULL(CONVERT(CHAR,IPM.InvoiceDate,106),''),  " & vbCrLf & _
    '                              "    						 url = 'InvFromPASIDetail.aspx?prm='+Rtrim(ISNULL(CONVERT(CHAR,IPM.InvoiceDate,106),CONVERT(CHAR,GETDATE(),106)))    " & vbCrLf & _
    '                              "    									+ '|' +Rtrim(isnull(POM.affiliateid,''))    "

    '            ls_SQL = ls_SQL + "    									+ '|' +Rtrim(isnull(MA.Affiliatename,''))    " & vbCrLf & _
    '                              "    									+ '|' +Rtrim(ISNULL(DPM.SuratJalanNo,''))    " & vbCrLf & _
    '                              "    									+ '|' +Rtrim(ISNULL(KD.PONO,''))    " & vbCrLf & _
    '                              "    									+ '|' +Rtrim(ISNULL(KD.KanbanNo,''))    " & vbCrLf & _
    '                              "                                     + '|' +Rtrim(ISNULL(IPM.InvoiceNo,'')),  " & vbCrLf & _
    '                              "    						 coldetail = (CASE WHEN IPM.InvoiceNo IS NULL THEN '' ELSE 'DETAIL' END),   " & vbCrLf & _
    '                              "                           Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),  " & vbCrLf & _
    '                              "                           FInvNo = isnull(IPM.InvoiceNo,''),   " & vbCrLf & _
    '                              "                           FDiffQty = 0,  " & vbCrLf & _
    '                              "                           FSJ = ISNULL(DSD.SuratJalanNo,''),   "

    '            ls_SQL = ls_SQL + "                           FInv = isnull(INVSM.InvoiceNo,''),   " & vbCrLf & _
    '                              "                           Fsupp = POM.SupplierID,   " & vbCrLf & _
    '                              "                           FAff = POM.AffiliateID,   " & vbCrLf & _
    '                              "                           FPO = POD.PONo  " & vbCrLf & _
    '                              "   FROM    dbo.PO_Master POM   " & vbCrLf & _
    '                              "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
    '                              "                                      AND POM.PoNo = POD.PONo   " & vbCrLf & _
    '                              "                                      AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   " & vbCrLf & _
    '                              "                                             AND KD.PoNo = POD.PONo   " & vbCrLf & _
    '                              "                                             AND KD.SupplierID = POD.SupplierID   "

    '            ls_SQL = ls_SQL + "                                             AND KD.PartNo = POD.PartNo   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
    '                              "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
    '                              "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
    '                              "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
    '                              "           INNER JOIN dbo.DOSupplier_Detail DSD ON KD.AffiliateID = DSD.AffiliateID   " & vbCrLf & _
    '                              "                                                  AND KD.KanbanNo = DSD.KanbanNo   " & vbCrLf & _
    '                              "                                                  AND KD.SupplierID = DSD.SupplierID   " & vbCrLf & _
    '                              "                                                  --AND KD.PartNo = DSD.PartNo   " & vbCrLf & _
    '                              "                                                  AND KD.PoNo = DSD.PoNo  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.DOSupplier_Master DSM ON DSM.AffiliateID = DSD.AffiliateID   "

    '            ls_SQL = ls_SQL + "                                                  AND DSM.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
    '                              "                                                  AND DSM.SupplierID = DSD.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceivePASI_Detail RPD ON KD.AffiliateID = RPD.AffiliateID   " & vbCrLf & _
    '                              "                                                   AND KD.KanbanNo = RPD.KanbanNo   " & vbCrLf & _
    '                              "                                                   AND KD.SupplierID = RPD.SupplierID   " & vbCrLf & _
    '                              "                                                   --AND KD.PartNo = RPD.PartNo   " & vbCrLf & _
    '                              "                                                   AND KD.PONo = RPD.PONo   " & vbCrLf & _
    '                              "                                                   AND RPD.SuratJalanNo = DSM.SuratJalanNo  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = RPD.AffiliateID   " & vbCrLf & _
    '                              "                                                   AND PRM.SuratJalanNo = RPD.SuratJalanNo   " & vbCrLf & _
    '                              "                                                   AND PRM.SupplierID = RPD.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.DOPASI_Detail DPD ON KD.AffiliateID = DPD.AffiliateID   "

    '            ls_SQL = ls_SQL + "                                              AND KD.KanbanNo = DPD.KanbanNo   " & vbCrLf & _
    '                              "                                              AND KD.SupplierID = DPD.SupplierID   " & vbCrLf & _
    '                              "                                              --AND KD.PartNo = DPD.PartNo   " & vbCrLf & _
    '                              "                                              AND KD.PoNo = DPD.PoNo   " & vbCrLf & _
    '                              "                                              AND DPD.SuratJalanNoSupplier = DSM.SuratJalanno  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.DOPASI_Master DPM ON DPD.AffiliateID = DPM.AffiliateID   " & vbCrLf & _
    '                              "                                              AND DPD.SuratJalanNo = DPM.SuratJalanNo   " & vbCrLf & _
    '                              "                                              AND DPD.SupplierID = DPM.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON DPD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
    '                              "                                                       AND DPD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
    '                              "                                                       AND DPD.SupplierID = RAD.SupplierID  " & vbCrLf & _
    '                              "                                                       --AND DPD.PartNo = RAD.PartNo  "

    '            ls_SQL = ls_SQL + "                                                       AND DPD.PoNo = RAD.PoNo  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
    '                              "                                                       AND RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
    '                              "                                                       AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
    '                              "  		 LEFT JOIN dbo.InvoicePASI_Detail IPD ON KD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
    '                              "  													AND KD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
    '                              "  													--AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
    '                              "  													AND KD.PONo = IPD.PONo  " & vbCrLf & _
    '                              "  													AND DPD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
    '                              "  		 LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
    '                              "  													AND IPD.InvoiceNo = IPM.InvoiceNo  "

    '            ls_SQL = ls_SQL + "  													AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
    '                              "  		 LEFT JOIN InvoiceSupplier_Detail INVSD ON INVSD.SupplierID = KD.SupplierID   " & vbCrLf & _
    '                              "    							AND INVSD.AffiliateID = KD.AffiliateID   " & vbCrLf & _
    '                              "    							AND INVSD.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
    '                              "    							AND INVSD.KanbanNo = KD.kanbanNo   " & vbCrLf & _
    '                              "    							AND INVSD.PONo = KD.PONo   " & vbCrLf & _
    '                              "    							--AND INVSD.PartNo = KD.PartNo   " & vbCrLf & _
    '                              " 		 LEFT JOIN InvoiceSupplier_Master INVSM ON INVSM.InvoiceNo = INVSD.InvoiceNo   " & vbCrLf & _
    '                              " 												AND INVSM.SupplierID = INVSD.SupplierID   " & vbCrLf & _
    '                              " 												AND INVSM.AffiliateID = INVSD.AffiliateID   " & vbCrLf & _
    '                              " 												AND INVSM.SuratJalanNo = INVSD.SuratJalanNo 											       "

    '            ls_SQL = ls_SQL + "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
    '                              "   WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' -- AND POD.pono='PO-DEE-KMK '    " & vbCrLf & _
    '                              "                 --WHERE POD.PONO <> '' " & vbCrLf

    '            If checkbox1.Checked = True Then
    '                ls_SQL = ls_SQL + " AND DPM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "' " & vbCrLf
    '            End If

    '            If checkbox2.Checked = True Then
    '                ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(ipm.invoicedate,'')),106) between '" & Format(dt3.Value, "dd MMM yyyy") & "' AND '" & Format(dt4.Value, "dd MMM yyyy") & "' " & vbCrLf
    '            End If

    '            If rbinvoice.Value = "YES" Then
    '                ls_SQL = ls_SQL + "AND isnull(IPM.InvoiceNo, '') <> '' " & vbCrLf
    '            ElseIf rbinvoice.Value = "NO" Then
    '                ls_SQL = ls_SQL + " AND isnull(IPM.InvoiceNo,'') = '' " & vbCrLf
    '            End If

    '            If txtsj.Text <> "" Then
    '                ls_SQL = ls_SQL + " AND DPD.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
    '            End If

    '            If txtpono.Text <> "" Then
    '                ls_SQL = ls_SQL + " AND POD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
    '            End If


    '            ls_SQL = ls_SQL + "   )hdr " & vbCrLf & _
    '                              "   UNION ALL   " & vbCrLf & _
    '                              "   SELECT DISTINCT url = '',coldetail = '',no = '',    " & vbCrLf & _
    '                              "    				 period = '',      " & vbCrLf & _
    '                              "    				 affiliatecode = '',    "

    '            ls_SQL = ls_SQL + "    				 affiliatename = '',    " & vbCrLf & _
    '                              "    				 pono = '',   " & vbCrLf & _
    '                              "    				 suppliercode = '',    " & vbCrLf & _
    '                              "    				 suppliername = '',    " & vbCrLf & _
    '                              "    				 kanbancls = '',    " & vbCrLf & _
    '                              "    				 kanbanno = '',    " & vbCrLf & _
    '                              "    				 suppplandeldate = '',    " & vbCrLf & _
    '                              "    				 suppdeldate = '',    " & vbCrLf & _
    '                              "    				 suppsj = '',    " & vbCrLf & _
    '                              "    				 pasideldate = '',    " & vbCrLf & _
    '                              "    				 pasisj = '',    "

    '            ls_SQL = ls_SQL + "    				 partno = POD.PartNo,    " & vbCrLf & _
    '                              "    				 partname = MP.PartName,    " & vbCrLf & _
    '                              "    				 uom = MU.Description,    " & vbCrLf & _
    '                              "    				 suppdelqty = Convert(char,Round(CONVERT(CHAR,Round(ISNULL(DSD.DOQty,0),0)),0)),    " & vbCrLf & _
    '                              "    				 pasirecqty = Convert(char,Round(convert(char,Round(Isnull(RPD.GoodRecQty,0),0)),0)),   " & vbCrLf & _
    '                              "    				 pasidelqty = Convert(char,Round(convert(char,Round(Isnull(DPD.DOQty,0),0)),0)),  " & vbCrLf & _
    '                              "    				 affrecqty = Convert(char,Round(convert(char,Round(Isnull(RAD.RecQty,0),0)),0)),  " & vbCrLf & _
    '                              "    				 affremrecqty = '',  " & vbCrLf & _
    '                              "    				 affrecdate = '', affiliaterecdate = '', affrecby = '', pasiinvno = '', pasiinvdate = '', " & vbCrLf & _
    '                              "    				 sortpono = POD.PONo, sortkanbanno = ISNULL(KD.KanbanNo,''), sortheader = 1,    " & vbCrLf & _
    '                              "                   Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),  "

    '            ls_SQL = ls_SQL + "                   FInvNo = isnull(IPM.InvoiceNo,''),   " & vbCrLf & _
    '                              "                   FDiffQty = 0,  " & vbCrLf & _
    '                              "                   FSJ = ISNULL(DSD.SuratJalanNo,''),   " & vbCrLf & _
    '                              "                   FInv = isnull(INVSM.InvoiceNo,''),   " & vbCrLf & _
    '                              "                   Fsupp = POM.SupplierID,   " & vbCrLf & _
    '                              "                   FAff = POM.AffiliateID,   " & vbCrLf & _
    '                              "                   FPO = POD.PONo  " & vbCrLf & _
    '                              "   FROM    dbo.PO_Master POM   " & vbCrLf & _
    '                              "           LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
    '                              "                                      AND POM.PoNo = POD.PONo   " & vbCrLf & _
    '                              "                                      AND POM.SupplierID = POD.SupplierID   "

    '            ls_SQL = ls_SQL + "           LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID   " & vbCrLf & _
    '                              "                                             AND KD.PoNo = POD.PONo   " & vbCrLf & _
    '                              "                                             AND KD.SupplierID = POD.SupplierID   " & vbCrLf & _
    '                              "                                             AND KD.PartNo = POD.PartNo   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID   " & vbCrLf & _
    '                              "                                             AND KD.KanbanNo = KM.KanbanNo   " & vbCrLf & _
    '                              "                                             AND KD.SupplierID = KM.SupplierID   " & vbCrLf & _
    '                              "                                             AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
    '                              "           INNER JOIN dbo.DOSupplier_Detail DSD ON KD.AffiliateID = DSD.AffiliateID   " & vbCrLf & _
    '                              "                                                  AND KD.KanbanNo = DSD.KanbanNo   " & vbCrLf & _
    '                              "                                                  AND KD.SupplierID = DSD.SupplierID   "

    '            ls_SQL = ls_SQL + "                                                  AND KD.PartNo = DSD.PartNo   " & vbCrLf & _
    '                              "                                                  AND KD.PoNo = DSD.PoNo  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.DOSupplier_Master DSM ON DSM.AffiliateID = DSD.AffiliateID   " & vbCrLf & _
    '                              "                                                  AND DSM.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
    '                              "                                                  AND DSM.SupplierID = DSD.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceivePASI_Detail RPD ON KD.AffiliateID = RPD.AffiliateID   " & vbCrLf & _
    '                              "                                                   AND KD.KanbanNo = RPD.KanbanNo   " & vbCrLf & _
    '                              "                                                   AND KD.SupplierID = RPD.SupplierID   " & vbCrLf & _
    '                              "                                                   AND KD.PartNo = RPD.PartNo   " & vbCrLf & _
    '                              "                                                   AND KD.PONo = RPD.PONo   " & vbCrLf & _
    '                              "                                                   AND RPD.SuratJalanNo = DSM.SuratJalanNo " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = RPD.AffiliateID   "

    '            ls_SQL = ls_SQL + "                                                   AND PRM.SuratJalanNo = RPD.SuratJalanNo   " & vbCrLf & _
    '                              "                                                   AND PRM.SupplierID = RPD.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.DOPASI_Detail DPD ON KD.AffiliateID = DPD.AffiliateID   " & vbCrLf & _
    '                              "                                              AND KD.KanbanNo = DPD.KanbanNo   " & vbCrLf & _
    '                              "                                              AND KD.SupplierID = DPD.SupplierID   " & vbCrLf & _
    '                              "                                              AND KD.PartNo = DPD.PartNo   " & vbCrLf & _
    '                              "                                              AND KD.PoNo = DPD.PoNo   " & vbCrLf & _
    '                              "                                              AND DPD.SuratJalanNoSupplier = DSM.SuratJalanno  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.DOPASI_Master DPM ON DPD.AffiliateID = DPM.AffiliateID   " & vbCrLf & _
    '                              "                                              AND DPD.SuratJalanNo = DPM.SuratJalanNo   " & vbCrLf & _
    '                              "                                              AND DPD.SupplierID = DPM.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON DPD.AffiliateID = RAD.AffiliateID  "

    '            ls_SQL = ls_SQL + "                                                       AND DPD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
    '                              "                                                       AND DPD.SupplierID = RAD.SupplierID  " & vbCrLf & _
    '                              "                                                       AND DPD.PartNo = RAD.PartNo  " & vbCrLf & _
    '                              "                                                       AND DPD.PoNo = RAD.PoNo  " & vbCrLf & _
    '                              "           LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
    '                              "                                                       AND RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
    '                              "                                                       AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
    '                              "  		 LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
    '                              "  													AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
    '                              "  													AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
    '                              "  													AND RAD.PONo = IPD.PONo  "

    '            ls_SQL = ls_SQL + "  													--AND RAD.SuratJalanNo = IPD.SuratJalanNo  " & vbCrLf & _
    '                              "  		 LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
    '                              "  													AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
    '                              "  													AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
    '                              "  		LEFT JOIN InvoiceSupplier_Detail INVSD ON INVSD.SupplierID = KD.SupplierID   " & vbCrLf & _
    '                              "    							AND INVSD.AffiliateID = KD.AffiliateID   " & vbCrLf & _
    '                              "    							AND INVSD.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
    '                              "    							AND INVSD.KanbanNo = KD.kanbanNo   " & vbCrLf & _
    '                              "    							AND INVSD.PONo = KD.PONo   " & vbCrLf & _
    '                              "    							--AND INVSD.PartNo = KD.PartNo   " & vbCrLf & _
    '                              " 		 LEFT JOIN InvoiceSupplier_Master INVSM ON INVSM.InvoiceNo = INVSD.InvoiceNo   "

    '            ls_SQL = ls_SQL + " 			AND INVSM.SupplierID = INVSD.SupplierID   " & vbCrLf & _
    '                              " 			AND INVSM.AffiliateID = INVSD.AffiliateID   " & vbCrLf & _
    '                              " 			AND INVSM.SuratJalanNo = INVSD.SuratJalanNo  " & vbCrLf & _
    '                              " 													 " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID   " & vbCrLf & _
    '                              "           LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
    '                              "     WHERE POD.AffiliateID = '" & Session("AffiliateID") & "' --AND POD.pono='PO-DEE-KMK' --'PO20150501-KMK'    " & vbCrLf & _
    '                              "   		   --WHERE POD.PONO <> '' " & vbCrLf & _
    '                              "  "

    '            If checkbox1.Checked = True Then
    '                ls_SQL = ls_SQL + " AND DPM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "' " & vbCrLf
    '            End If

    '            If checkbox2.Checked = True Then
    '                ls_SQL = ls_SQL + " AND CONVERT(CHAR(12), CONVERT(DATETIME,ISNULL(ipm.invoicedate,'')),106) between '" & Format(dt3.Value, "dd MMM yyyy") & "' AND '" & Format(dt4.Value, "dd MMM yyyy") & "' " & vbCrLf
    '            End If

    '            If rbinvoice.Value = "YES" Then
    '                ls_SQL = ls_SQL + "AND isnull(IPM.InvoiceNo, '') <> '' " & vbCrLf
    '            ElseIf rbinvoice.Value = "NO" Then
    '                ls_SQL = ls_SQL + " AND isnull(IPM.InvoiceNo,'') = '' " & vbCrLf
    '            End If

    '            If txtsj.Text <> "" Then
    '                ls_SQL = ls_SQL + " AND DPD.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
    '            End If


    '            If txtpono.Text <> "" Then
    '                ls_SQL = ls_SQL + " AND POD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
    '            End If

    '            ls_SQL = ls_SQL + "    ) SPDC ORDER BY SortPONo, SortKanbanNo, FSJ, SortHeader  " & vbCrLf & _
    '                              "  "

    '            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '            Dim ds As New DataSet
    '            sqlDA.Fill(ds)
    '            With grid
    '                .DataSource = ds.Tables(0)
    '                .DataBind()
    '                'Call ColorGrid()
    '            End With
    '            sqlConn.Close()


    '        End Using
    '    End Sub

    '#End Region

#Region "FORM EVENT"
    '    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
    '        Try
    '            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    '            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, true, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
    '            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
    '            grid.JSProperties("cpMessage") = Session("AA220Msg")

    '            Dim pAction As String = Split(e.Parameters, "|")(0)
    '            'Dim pPlan As Date = Split(e.Parameters, "|")(1)
    '            'Dim pSupplierDeliver As String = Split(e.Parameters, "|")(2)
    '            'Dim pRemaining As String = Split(e.Parameters, "|")(3)
    '            'Dim psj As String = Split(e.Parameters, "|")(4)
    '            'Dim pDateFrom As Date = Split(e.Parameters, "|")(5)
    '            'Dim pDateTo As Date = Split(e.Parameters, "|")(6)
    '            'Dim pSupplier As String = Split(e.Parameters, "|")(7)
    '            'Dim pPart As String = Split(e.Parameters, "|")(8)
    '            'Dim pPoNo As String = Split(e.Parameters, "|")(9)
    '            'Dim pKanban As String = Split(e.Parameters, "|")(10)

    '            Select Case pAction
    '                Case "gridload"
    '                    'Call up_GridLoad(pPlan, pSupplierDeliver, pRemaining, psj, pDateFrom, pDateTo, pSupplier, pPart, pPoNo, pKanban)
    '                    Call up_GridLoad()
    '                    If grid.VisibleRowCount = 0 Then
    '                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    '                        grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                    End If
    '                Case "kosong"

    '            End Select

    'EndProcedure:
    '            Session("AA220Msg") = ""
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '        End Try
    '    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

#End Region

    'Private Sub grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowPrepared
    '    If e.RowType <> GridViewRowType.Data Then Return
    '    If e.GetValue("partno").ToString = "" Then e.Row.BackColor = Drawing.Color.LightGray
    'End Sub
End Class