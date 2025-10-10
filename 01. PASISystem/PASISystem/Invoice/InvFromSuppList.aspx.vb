Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView

Public Class InvFromSuppList
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
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'ls_SQL = "  SELECT *    " & vbCrLf & _
            '      "   	FROM (   " & vbCrLf & _
            '      "   		  SELECT url,coldetail, no = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, KanbanCls, KanbanNo)),   " & vbCrLf & _
            '      "   			     period, affiliatecode, affiliatename,pono, suppliercode, suppliername, pokanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,    " & vbCrLf & _
            '      "   				 kanbanno, suppplandeldate = SupplierPlanDeliveryDate, suppdeldate = supplierdeliverydate, suppsj = suppliersjno, pasideldate = pasideliverydate, pasisj = pasisjno, partno, partname, uom,  " & vbCrLf & _
            '      "   				 suppdelqty, pasirecqty, remrecqty, suppinvqty, suppinvno,   " & vbCrLf & _
            '      "   				 suppinvdate, pasireccurr = isnull(pasireccurr,''), pasirecamount = isnull(pasirecamount,''), suppinvcurr = isnull(suppinvcurr,''),suppinvamount = isnull(suppinvamount,''),   " & vbCrLf & _
            '      "   				 sortpono = PONo, sortkanbanno = KanbanNo, sortheader = 0, Fsuppdeldate, FInvNo, FDiffQty,FSJ, FInv, Fsupp, FAff, FPO     " & vbCrLf & _
            '      "   			FROM (   " & vbCrLf & _
            '      "   				  SELECT DISTINCT    " & vbCrLf & _
            '      "   						 period = SUBSTRING(CONVERT(CHAR,POM.period,106),4,9),   " & vbCrLf

            'ls_SQL = ls_SQL + "   						 PONo = POD.PONo,    " & vbCrLf & _
            '                  "   						 affiliatecode = ISNULL(KM.AffiliateID,''),   " & vbCrLf & _
            '                  "                            affiliatename = ISNULL(MA.AffiliateName,''),   " & vbCrLf & _
            '                  "   						 SupplierCode = POM.SupplierID,   " & vbCrLf & _
            '                  "                            SupplierName = MS.SupplierName,   " & vbCrLf & _
            '                  "   						 KanbanCls = ISNULL(POD.KanbanCls,'0'),   " & vbCrLf & _
            '                  "   						 KanbanNo = ISNULL(KD.KanbanNo,''),   " & vbCrLf & _
            '                  "   						 SupplierPlanDeliveryDate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),   " & vbCrLf & _
            '                  "   						 SupplierDeliveryDate = ISNULL(CONVERT(CHAR,DSM.DeliveryDate,106),''),   " & vbCrLf & _
            '                  "   						 SupplierSJNO = ISNULL(DSD.SuratJalanNo,''),   " & vbCrLf & _
            '                  "   						 PASIDeliveryDate = ISNULL(CONVERT(CHAR,DPM.DeliveryDate,106),''),   " & vbCrLf

            'ls_SQL = ls_SQL + "   						 PASISJNo = ISNULL(DPD.SuratJalanNo,''),   " & vbCrLf & _
            '                  "   						 PartNo = '',   " & vbCrLf & _
            '                  "   						 PartName = '',   " & vbCrLf & _
            '                  "   						 UOM = '',   " & vbCrLf & _
            '                  "   						 suppdelqty = '',   " & vbCrLf & _
            '                  "   						 pasirecqty = '',   " & vbCrLf & _
            '                  "   						 remrecqty = '',   " & vbCrLf & _
            '                  "   						 suppinvqty = '', --Round(convert(char, Isnull(INVSD.INVQty,0)),0), " & vbCrLf & _
            '                  "   						 suppinvno = isnull(INVSM.InvoiceNo,''),   " & vbCrLf & _
            '                  "   				         suppinvdate = ISNULL(CONVERT(CHAR,INVSM.InvoiceDate,113),''),   " & vbCrLf & _
            '                  "   				         pasireccurr = case when deliveryByPasiCls = 1 then MCPasi.Description else MCAff.Description end, --(select Description From MS_CurrCls where currcls = POD.CurrCls),   " & vbCrLf

            'ls_SQL = ls_SQL + "   				         pasirecamount = case when deliveryByPasiCls = 1 then Convert(Varchar,cast(isnull(SumPasiRec.tot,0) as money),1) else Convert(Varchar,cast(isnull(SumAffRec.tot,0) as money),1) end,  " & vbCrLf & _
            '                  "   				         suppinvcurr = (select Description From MS_CurrCls where currcls = isnull(INVSD.InvCurrCls,'')),  " & vbCrLf & _
            '                  "   				         suppinvamount = Convert(Varchar,cast(isnull(IVD.A,0) as money),1),				           " & vbCrLf & _
            '                  "   						 url = 'InvoiceEntry.aspx?prm='+Rtrim(ISNULL(CONVERT(CHAR,INVSM.InvoiceDate,106),CONVERT(CHAR,GETDATE(),106)))   " & vbCrLf & _
            '                  "   									+ '|' +Rtrim(KM.affiliateid)   " & vbCrLf & _
            '                  "   									+ '|' +Rtrim(MA.Affiliatename)   " & vbCrLf & _
            '                  "   									+ '|' +Rtrim(ISNULL(INVSM.SuratJalanNo,''))   " & vbCrLf & _
            '                  "   									+ '|' +Rtrim(ISNULL(KD.PONO,''))   " & vbCrLf & _
            '                  "   									+ '|' +Rtrim(ISNULL(KD.KanbanNo,''))   " & vbCrLf & _
            '                  "                                     + '|' +Rtrim(ISNULL(INVSM.InvoiceNo,'')) " & vbCrLf & _
            '                  "                                     + '|' +Rtrim(ISNULL(KM.SupplierID,'')), " & vbCrLf & _
            '                  "   						 coldetail = (CASE WHEN invsd.InvAmount IS NULL THEN 'INVOICE' ELSE 'DETAIL' END),  " & vbCrLf & _
            '                  "                          Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''), " & vbCrLf & _
            '                  "                          FInvNo = isnull(INVSM.InvoiceNo,''),  " & vbCrLf & _
            '                  "                          FDiffQty = 0, " & vbCrLf & _
            '                  "                          FSJ = ISNULL(DSD.SuratJalanNo,''),  " & vbCrLf & _
            '                  "                          FInv = isnull(INVSM.InvoiceNo,''),  " & vbCrLf & _
            '                  "                          Fsupp = POM.SupplierID,  " & vbCrLf & _
            '                  "                          FAff = POM.AffiliateID,  " & vbCrLf & _
            '                  "                          FPO = POD.PONo " & vbCrLf & _
            '                  "   					FROM PO_DETAIL POD   with(nolock)  " & vbCrLf

            'ls_SQL = ls_SQL + "   						 LEFT JOIN PO_Master POM  with(nolock) ON POM.AffiliateID =POD.AffiliateID   " & vbCrLf & _
            '                  "   							AND POM.SupplierID =POD.SupplierID   " & vbCrLf & _
            '                  "   							AND POM.PONO =POD.PONO   " & vbCrLf & _
            '                  "   						 LEFT JOIN Kanban_Detail KD  with(nolock) ON KD.AffiliateID =POD.AffiliateID   " & vbCrLf & _
            '                  "   							AND KD.SupplierID =POD.SupplierID   " & vbCrLf & _
            '                  "   							AND KD.PONO =POD.PONO   " & vbCrLf & _
            '                  "   							AND KD.PartNo =POD.PartNo   " & vbCrLf & _
            '                  "   						 LEFT JOIN Kanban_Master KM  with(nolock)  ON KD.AffiliateID =KM.AffiliateID   " & vbCrLf & _
            '                  "   							AND KD.SupplierID =KM.SupplierID   " & vbCrLf & _
            '                  "   							AND KD.KanbanNo =KM.KanbanNo   " & vbCrLf & _
            '                  "                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf

            'ls_SQL = ls_SQL + "   						 LEFT JOIN DOSupplier_Detail DSD   with(nolock) ON KD.AffiliateID =DSD.AffiliateID   " & vbCrLf & _
            '                  "   							AND KD.SupplierID =DSD.SupplierID   " & vbCrLf & _
            '                  "   							AND KD.PONO =DSD.PONO   " & vbCrLf & _
            '                  "   							AND KD.PartNo =DSD.PartNo   " & vbCrLf & _
            '                  "   							AND KD.KanbanNo =DSD.KanbanNo   " & vbCrLf & _
            '                  "   						 INNER JOIN DOSupplier_Master DSM   with(nolock) ON DSM.AffiliateID =DSD.AffiliateID   " & vbCrLf & _
            '                  "   							AND DSM.SupplierID =DSD.SupplierID   " & vbCrLf & _
            '                  "   							AND DSM.SuratJalanNo =DSD.SuratJalanNo   " & vbCrLf & _
            '                  "   						 LEFT JOIN DOPASI_Detail DPD   with(nolock) ON DPD.AffiliateID =KD.AffiliateID   " & vbCrLf & _
            '                  "   							AND DPD.SupplierID =KD.SupplierID   " & vbCrLf & _
            '                  "   							AND DPD.PONO =KD.PONO   " & vbCrLf

            'ls_SQL = ls_SQL + "   							--AND KD.PartNo =DPD.PartNo   " & vbCrLf & _
            '                  "   							AND KD.KanbanNo =DPD.KanbanNo   " & vbCrLf & _
            '                  "   						 LEFT JOIN DOPASI_Master DPM   with(nolock) ON DPM.AffiliateID =DPD.AffiliateID   " & vbCrLf & _
            '                  "   							AND DPM.SupplierID =DPD.SupplierID   " & vbCrLf & _
            '                  "   							AND DPM.SuratJalanNo =DPD.SuratJalanNo   " & vbCrLf & _
            '                  "   						 LEFT JOIN ReceivePASI_Detail RPD   with(nolock) ON RPD.AffiliateID = DPM.AffiliateID   " & vbCrLf & _
            '                  "   							AND RPD.SupplierID = DPM.SupplierID   " & vbCrLf & _
            '                  "   							AND RPD.PONo = POD.PONo   " & vbCrLf & _
            '                  "   							AND RPD.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "   							AND RPD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "   				         LEFT JOIN ReceiveAffiliate_Detail RAD   with(nolock) ON RAD.AffiliateID = KD.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "   					        AND RAD.SupplierID = KD.SupplierID   " & vbCrLf & _
            '                  "   					        AND RAD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "   					        AND RAD.PONo = KD.PONo   " & vbCrLf & _
            '                  "   					        AND RAD.PartNo = KD.PartNo   " & vbCrLf & _
            '                  "   				         LEFT JOIN ReceiveAffiliate_Master RAM   with(nolock) ON RAM.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "   					        AND RAM.SupplierID = RAD.SupplierID   " & vbCrLf & _
            '                  "   					        AND RAM.SuratJalanNo = RAD.SuratJalanNo   " & vbCrLf & _
            '                  "   					     LEFT JOIN InvoiceSupplier_Detail INVSD   with(nolock) ON INVSD.SupplierID = KD.SupplierID  " & vbCrLf & _
            '                  "   							AND INVSD.AffiliateID = KD.AffiliateID  " & vbCrLf & _
            '                  "   							AND INVSD.SuratJalanNo = DSD.SuratJalanNo  " & vbCrLf & _
            '                  "   							AND INVSD.KanbanNo = KD.kanbanNo  " & vbCrLf

            'ls_SQL = ls_SQL + "   							AND INVSD.PONo = KD.PONo  " & vbCrLf & _
            '                  "   							--AND INVSD.PartNo = KD.PartNo  " & vbCrLf & _
            '                  "  						 LEFT JOIN InvoiceSupplier_Master INVSM   with(nolock) ON INVSM.InvoiceNo = INVSD.InvoiceNo  " & vbCrLf & _
            '                  "  							AND INVSM.SupplierID = INVSD.SupplierID  " & vbCrLf & _
            '                  "  							AND INVSM.AffiliateID = INVSD.AffiliateID  " & vbCrLf & _
            '                  "  							AND INVSM.SuratJalanNo = INVSD.SuratJalanNo  " & vbCrLf & _
            '                  "  						 LEFT JOIn MS_CurrCls MC   with(nolock) ON MC.CurrCls = INVSD.InvCurrCls  " & vbCrLf & _
            '                  "   						 LEFT JOIN MS_Parts MP   with(nolock) ON MP.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "   						 LEFT JOIN MS_UnitCls MU   with(nolock) ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "                            LEFT JOIN MS_Affiliate MA   with(nolock) ON MA.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "                            LEFT JOIN dbo.MS_Supplier MS   with(nolock) ON MS.SupplierID = POM.SupplierID    " & vbCrLf

            'ls_SQL = ls_SQL + "                            LEFT JOIN    " & vbCrLf & _
            '                  "    				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(RecQty),0) + ISNULL(SUM(DefectQty),0)   " & vbCrLf & _
            '                  "    				            FROM ReceiveAffiliate_Detail  with(nolock) " & vbCrLf & _
            '                  "    				           GROUP BY PONo, KanbanNo   " & vbCrLf & _
            '                  "    				         ) SumKanban ON SumKanban.PONo = KD.PONo AND SumKanban.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "    				         LEFT JOIN    " & vbCrLf & _
            '                  "    				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)   " & vbCrLf & _
            '                  "    				            FROM DOSupplier_Detail   with(nolock)   " & vbCrLf & _
            '                  "    				           GROUP BY PONo, KanbanNo   " & vbCrLf & _
            '                  "    				         ) SumDSD ON SumDSD.PONo = KD.PONo AND SumDSD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "    				         LEFT JOIN    " & vbCrLf

            'ls_SQL = ls_SQL + "    				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)   " & vbCrLf & _
            '                  "    				            FROM DOPASI_Detail   with(nolock)   " & vbCrLf & _
            '                  "    				           GROUP BY PONo, KanbanNo   " & vbCrLf & _
            '                  "    				         ) SumDPD ON SumDPD.PONo = KD.PONo AND SumDPD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
            '                  "    				         LEFT JOIN  " & vbCrLf & _
            '                  "    				         (select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,tot = sum(GoodRecQty * isnull(C.Price,0)), C.currcls  " & vbCrLf & _
            '                  "  							    From ReceivePasi_Detail A   with(nolock) Left join ReceivePasi_Master B    with(nolock) " & vbCrLf & _
            '                  "  							        ON A.SuratJalanNo = B.SuratJalanNo  " & vbCrLf & _
            '                  "  							    AND A.SupplierID = B.SupplierID  " & vbCrLf & _
            '                  "  							    Left Join PO_detail D   with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo  " & vbCrLf & _
            '                  "  							    Left Join MS_Price C   with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate  " & vbCrLf

            'ls_SQL = ls_SQL + "  									 and C.AffiliateID = A.AffiliateID  " & vbCrLf & _
            '                  "  							Group by A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,C.currcls ) SumPasiRec  " & vbCrLf & _
            '                  "  						  ON SumPasiRec.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "   							AND SumPasiRec.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "   							AND SumPasiRec.PONo = POD.PONo   " & vbCrLf & _
            '                  "   							AND SumPasiRec.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "    				         LEFT JOIN  " & vbCrLf & _
            '                  "    				         (select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,tot = sum(RecQty * isnull(C.Price,0)), C.currcls  " & vbCrLf & _
            '                  "  							    From ReceiveAffiliate_Detail A   with(nolock) Left join ReceiveAffiliate_Master B    with(nolock) " & vbCrLf & _
            '                  "  							        ON A.SuratJalanNo = B.SuratJalanNo  " & vbCrLf & _
            '                  "  							    AND A.SupplierID = B.SupplierID  " & vbCrLf & _
            '                  "  							    Left Join PO_detail D   with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo  " & vbCrLf & _
            '                  "  							    Left Join MS_Price C   with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate  " & vbCrLf

            'ls_SQL = ls_SQL + "  									  AND  C.AffiliateID = A.AffiliateID  " & vbCrLf & _
            '                  "  							Group by A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,C.currcls ) SumAffRec  " & vbCrLf & _
            '                  "  						  ON SumAffRec.AffiliateID = KM.AffiliateID   " & vbCrLf & _
            '                  "   							AND SumAffRec.SupplierID = POM.SupplierID   " & vbCrLf & _
            '                  "   							AND SumAffRec.PONo = POD.PONo   " & vbCrLf & _
            '                  "   							AND SumAffRec.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "                         LEFT JOIn MS_CurrCls MCPasi   with(nolock) ON MCPasi.CurrCls = SumPasiRec.CurrCls  " & vbCrLf & _
            '                  "                         LEFT JOIn MS_CurrCls MCAff   with(nolock) ON MCAff.CurrCls = SumAffRec.CurrCls  " & vbCrLf & _
            '                  "                         LEFT JOIN (select SupplierID,AffiliateID, KanbanNo, PONo, A = sum(InvAmount) from " & vbCrLf & _
            '                  "		                                InvoiceSupplier_Detail INVSD   with(nolock) Group by SupplierID,AffiliateID, KanbanNo, PONo)IVD" & vbCrLf & _
            '                  "                                     ON IVD.SupplierID = KD.SupplierID  " & vbCrLf & _
            '                  "                                     AND IVD.AffiliateID = INVSD.AffiliateID  " & vbCrLf & _
            '                  "                                     AND IVD.KanbanNo = INVSD.kanbanNo  " & vbCrLf & _
            '                  "                                     AND IVD.PONo = INVSD.PONo " & vbCrLf & _
            '                  "   		           --WHERE POD.AffiliateID = 'JAI' AND POD.pono='PO20150501-KMK '   " & vbCrLf & _
            '                  "                 WHERE POD.PONO <> '' " & vbCrLf

            '20150714
            ls_SQL = " SELECT *     " & vbCrLf & _
                  "    	FROM (    " & vbCrLf & _
                  "    		  SELECT url = url, coldetail, no = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, KanbanCls, KanbanNo)),    " & vbCrLf & _
                  "    			     period, affiliatecode, affiliatename,pono, suppliercode, suppliername, pokanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,     " & vbCrLf & _
                  "    				 kanbanno, suppplandeldate = SupplierPlanDeliveryDate, suppdeldate = supplierdeliverydate, suppsj = suppliersjno, pasideldate = pasideliverydate, pasisj = pasisjno, partno, partname, uom,   " & vbCrLf & _
                  "    				 suppdelqty, pasirecqty, remrecqty, suppinvqty, suppinvno,    " & vbCrLf & _
                  "    				 suppinvdate, pasireccurr = isnull(pasireccurr,''), pasirecamount = isnull(pasirecamount,''), /*suppinvcurr = isnull(suppinvcurr,''),suppinvamount = isnull(suppinvamount,''),*/    " & vbCrLf & _
                  "    				 sortpono = PONo, sortkanbanno = KanbanNo, sortheader = 0, Fsuppdeldate, FInvNo, FDiffQty,FSJ, FInv, Fsupp, FAff, FPO     " & vbCrLf & _
                  "    				 ,H_SupSj,H_SupINV " & vbCrLf & _
                  "    			FROM (    " & vbCrLf & _
                  "    				  SELECT DISTINCT     "

            ls_SQL = ls_SQL + "    						 period = SUBSTRING(CONVERT(CHAR,POM.period,106),4,9),    " & vbCrLf & _
                              "    						 PONo = POD.PONo,     " & vbCrLf & _
                              "    						 affiliatecode = ISNULL(KM.AffiliateID,''),    " & vbCrLf & _
                              "                             affiliatename = ISNULL(MA.AffiliateName,''),    " & vbCrLf & _
                              "    						 SupplierCode = POM.SupplierID,    " & vbCrLf & _
                              "                             SupplierName = MS.SupplierName,    " & vbCrLf & _
                              "    						 KanbanCls = ISNULL(POD.KanbanCls,'0'),    " & vbCrLf & _
                              "    						 KanbanNo = ISNULL(KD.KanbanNo,''),    " & vbCrLf & _
                              "    						 SupplierPlanDeliveryDate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),    " & vbCrLf & _
                              "    						 SupplierDeliveryDate = ISNULL(CONVERT(CHAR,DSM.DeliveryDate,106),''),    " & vbCrLf & _
                              "    						 SupplierSJNO = ISNULL(DSM.SuratJalanNo,''),    "

            ls_SQL = ls_SQL + "    						 PASIDeliveryDate = ISNULL(CONVERT(CHAR,DPM.DeliveryDate,106),''),    " & vbCrLf & _
                              "    						 PASISJNo = ISNULL(DPD.SuratJalanNo,''),    " & vbCrLf & _
                              "    						 PartNo = '',    " & vbCrLf & _
                              "    						 PartName = '',    " & vbCrLf & _
                              "    						 UOM = '',    " & vbCrLf & _
                              "    						 suppdelqty = '',    " & vbCrLf & _
                              "    						 pasirecqty = '',    " & vbCrLf & _
                              "    						 remrecqty = '',    " & vbCrLf & _
                              "    						 suppinvqty = '', --Round(convert(char, Isnull(INVSD.INVQty,0)),0),  " & vbCrLf & _
                              "    						 suppinvno = isnull(INVSM.InvoiceNo,''),    " & vbCrLf & _
                              "    				         suppinvdate = ISNULL(CONVERT(CHAR,INVSM.InvoiceDate,113),''),    "

            ls_SQL = ls_SQL + "    				         pasireccurr = case when deliveryByPasiCls = 1 then MCPasi.Description else MCAff.Description end, --(select Description From MS_CurrCls where currcls = POD.CurrCls),    " & vbCrLf & _
                              "    				         pasirecamount = case when deliveryByPasiCls = 1 then Convert(Varchar,cast(isnull(SumPasiRec.tot,0) as money),1) else Convert(Varchar,cast(isnull(SumAffRec.tot,0) as money),1) end,   " & vbCrLf & _
                              "    				         suppinvcurr = (select Description From MS_CurrCls where currcls = isnull(INVSD.InvCurrCls,'')),   " & vbCrLf & _
                              "    				         suppinvamount = Convert(Varchar,cast(isnull(IVD.A,0) as money),1),				            " & vbCrLf & _
                              "    						 url = 'InvoiceEntry.aspx?prm='+Rtrim(ISNULL(CONVERT(CHAR,INVSM.InvoiceDate,106),CONVERT(CHAR,GETDATE(),106)))    " & vbCrLf & _
                              "    									+ '|' +Rtrim(KM.affiliateid)    " & vbCrLf & _
                              "    									+ '|' +Rtrim(MA.Affiliatename)    " & vbCrLf & _
                              "    									+ '|' +Rtrim(ISNULL(INVSM.SuratJalanNo,''))    " & vbCrLf & _
                              "    									+ '|' +Rtrim(ISNULL(KD.PONO,''))    " & vbCrLf & _
                              "    									+ '|' +Rtrim(ISNULL(KD.KanbanNo,''))    " & vbCrLf & _
                              "                                      + '|' +Rtrim(ISNULL(INVSM.InvoiceNo,''))  "

            ls_SQL = ls_SQL + "                                      + '|' +Rtrim(ISNULL(KM.SupplierID,'')),  " & vbCrLf & _
                              "    						 coldetail = (CASE WHEN invsd.InvAmount IS NULL THEN 'INVOICE' ELSE 'DETAIL' END),   " & vbCrLf & _
                              "                           Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),  " & vbCrLf & _
                              "                           FInvNo = isnull(INVSM.InvoiceNo,''),   " & vbCrLf & _
                              "                           FDiffQty = 0,  " & vbCrLf & _
                              "                           FSJ = ISNULL(DSD.SuratJalanNo,''),   " & vbCrLf & _
                              "                           FInv = isnull(INVSM.InvoiceNo,''),   " & vbCrLf & _
                              "                           Fsupp = POM.SupplierID,   " & vbCrLf & _
                              "                           FAff = POM.AffiliateID,   " & vbCrLf & _
                              "                           FPO = POD.PONo " & vbCrLf & _
                              "                           ,H_SupSj = DSM.SuratJalanNO " & vbCrLf & _
                              "                   , H_SupINV = isnull(INVSM.InvoiceNo,'') " & vbCrLf

            ls_SQL = ls_SQL + "    					FROM PO_DETAIL POD   with(nolock)   " & vbCrLf & _
                              "    						 INNER JOIN PO_Master POM  with(nolock) ON POM.AffiliateID =POD.AffiliateID    " & vbCrLf & _
                              "    							AND POM.SupplierID =POD.SupplierID    " & vbCrLf & _
                              "    							AND POM.PONO =POD.PONO    " & vbCrLf & _
                              "    						 INNER JOIN Kanban_Detail KD  with(nolock) ON KD.AffiliateID =POD.AffiliateID    " & vbCrLf & _
                              "    							AND KD.SupplierID =POD.SupplierID    " & vbCrLf & _
                              "    							AND KD.PONO =POD.PONO    " & vbCrLf & _
                              "    							AND KD.PartNo =POD.PartNo    " & vbCrLf & _
                              "    						 INNER JOIN Kanban_Master KM  with(nolock)  ON KD.AffiliateID =KM.AffiliateID    " & vbCrLf & _
                              "    							AND KD.SupplierID =KM.SupplierID    " & vbCrLf & _
                              "    							AND KD.KanbanNo =KM.KanbanNo    "

            ls_SQL = ls_SQL + "                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
                              "    						 INNER JOIN DOSupplier_Detail DSD   with(nolock) ON KD.AffiliateID =DSD.AffiliateID    " & vbCrLf & _
                              "    							AND KD.SupplierID =DSD.SupplierID    " & vbCrLf & _
                              "    							AND KD.PONO =DSD.PONO    " & vbCrLf & _
                              "    							AND KD.PartNo =DSD.PartNo    " & vbCrLf & _
                              "    							AND KD.KanbanNo =DSD.KanbanNo    " & vbCrLf & _
                              "    						 INNER JOIN DOSupplier_Master DSM   with(nolock) ON DSM.AffiliateID =DSD.AffiliateID    " & vbCrLf & _
                              "    							AND DSM.SupplierID =DSD.SupplierID    " & vbCrLf & _
                              "    							AND DSM.SuratJalanNo =DSD.SuratJalanNo    " & vbCrLf & _
                              "    						 LEFT JOIN DOPASI_Detail DPD   with(nolock) ON DPD.AffiliateID =KD.AffiliateID    " & vbCrLf & _
                              "    							AND DPD.SupplierID =KD.SupplierID    "

            ls_SQL = ls_SQL + "    							AND DPD.PONO =KD.PONO    " & vbCrLf & _
                              "    							AND KD.KanbanNo =DPD.KanbanNo " & vbCrLf & _
                              "    							AND DPD.SuratJalanNoSupplier = DSD.SuratJalanNo    " & vbCrLf & _
                              "    						 LEFT JOIN DOPASI_Master DPM   with(nolock) ON DPM.AffiliateID =DPD.AffiliateID    " & vbCrLf & _
                              "    							AND DPM.SupplierID =DPD.SupplierID    " & vbCrLf & _
                              "    							AND DPM.SuratJalanNo =DPD.SuratJalanNo    " & vbCrLf & _
                              "    						 LEFT JOIN ReceivePASI_Detail RPD   with(nolock) ON RPD.AffiliateID = DPM.AffiliateID    " & vbCrLf & _
                              "    							AND RPD.SupplierID = DPM.SupplierID    " & vbCrLf & _
                              "    							AND RPD.PONo = POD.PONo    " & vbCrLf & _
                              "    							AND RPD.PartNo = POD.PartNo    " & vbCrLf & _
                              "    							AND RPD.KanbanNo = KD.KanbanNo    "

            ls_SQL = ls_SQL + "    							AND RPD.SuratJalanNo = DSM.SuratJalanNo " & vbCrLf & _
                              "    				         LEFT JOIN ReceiveAffiliate_Detail RAD   with(nolock) ON RAD.AffiliateID = KD.AffiliateID    " & vbCrLf & _
                              "    					        AND RAD.SupplierID = KD.SupplierID    " & vbCrLf & _
                              "    					        AND RAD.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "    					        AND RAD.PONo = KD.PONo    " & vbCrLf & _
                              "    					        AND RAD.PartNo = KD.PartNo    " & vbCrLf & _
                              "    				         LEFT JOIN ReceiveAffiliate_Master RAM   with(nolock) ON RAM.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
                              "    					        AND RAM.SupplierID = RAD.SupplierID    " & vbCrLf & _
                              "    					        AND RAM.SuratJalanNo = RAD.SuratJalanNo    " & vbCrLf & _
                              "    					     LEFT JOIN InvoiceSupplier_Detail INVSD   with(nolock) ON INVSD.SupplierID = KD.SupplierID   " & vbCrLf & _
                              "    							AND INVSD.AffiliateID = KD.AffiliateID   "

            ls_SQL = ls_SQL + "    							AND INVSD.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
                              "    							AND INVSD.KanbanNo = KD.kanbanNo   " & vbCrLf & _
                              "    							AND INVSD.PONo = KD.PONo   " & vbCrLf & _
                              "   						 LEFT JOIN InvoiceSupplier_Master INVSM   with(nolock) ON INVSM.InvoiceNo = INVSD.InvoiceNo   " & vbCrLf & _
                              "   							AND INVSM.SupplierID = INVSD.SupplierID   " & vbCrLf & _
                              "   							AND INVSM.AffiliateID = INVSD.AffiliateID   " & vbCrLf & _
                              "   							AND INVSM.SuratJalanNo = INVSD.SuratJalanNo   " & vbCrLf & _
                              "   						 LEFT JOIn MS_CurrCls MC   with(nolock) ON MC.CurrCls = INVSD.InvCurrCls   " & vbCrLf & _
                              "    						 LEFT JOIN MS_Parts MP   with(nolock) ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "    						 LEFT JOIN MS_UnitCls MU   with(nolock) ON MU.UnitCls = MP.UnitCls    " & vbCrLf & _
                              "                             LEFT JOIN MS_Affiliate MA   with(nolock) ON MA.AffiliateID = KM.AffiliateID    "

            ls_SQL = ls_SQL + "                             LEFT JOIN dbo.MS_Supplier MS   with(nolock) ON MS.SupplierID = POM.SupplierID     " & vbCrLf & _
                              "                             LEFT JOIN     " & vbCrLf & _
                              "     				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(RecQty),0) + ISNULL(SUM(DefectQty),0)    " & vbCrLf & _
                              "     				            FROM ReceiveAffiliate_Detail  with(nolock)  " & vbCrLf & _
                              "     				           GROUP BY PONo, KanbanNo    " & vbCrLf & _
                              "     				         ) SumKanban ON SumKanban.PONo = KD.PONo AND SumKanban.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "     				         LEFT JOIN     " & vbCrLf & _
                              "     				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)    " & vbCrLf & _
                              "     				            FROM DOSupplier_Detail   with(nolock)    " & vbCrLf & _
                              "     				           GROUP BY PONo, KanbanNo    " & vbCrLf & _
                              "     				         ) SumDSD ON SumDSD.PONo = KD.PONo AND SumDSD.KanbanNo = KD.KanbanNo    "

            ls_SQL = ls_SQL + "     				         LEFT JOIN     " & vbCrLf & _
                              "     				         (SELECT SuratJalanNOSupplier, PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)    " & vbCrLf & _
                              "     				            FROM DOPASI_Detail   with(nolock)    " & vbCrLf & _
                              "     				           GROUP BY PONo, KanbanNo,SuratJalanNOSupplier    " & vbCrLf & _
                              "     				         ) SumDPD ON SumDPD.PONo = KD.PONo AND SumDPD.KanbanNo = KD.KanbanNo and sumDPD.SuratJalanNOSupplier = DSM.Suratjalanno " & vbCrLf & _
                              "     				         LEFT JOIN   " & vbCrLf & _
                              "     				         (select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,tot = sum(GoodRecQty * isnull(C.Price,0)), C.currcls   " & vbCrLf & _
                              "   							    From ReceivePasi_Detail A   with(nolock) Left join ReceivePasi_Master B    with(nolock)  " & vbCrLf & _
                              "   							        ON A.SuratJalanNo = B.SuratJalanNo   " & vbCrLf & _
                              "   							    AND A.SupplierID = B.SupplierID   " & vbCrLf & _
                              "   							    Left Join PO_detail D   with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo   "

            ls_SQL = ls_SQL + "   							    Left Join MS_Price C   with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate   " & vbCrLf & _
                              "   									 and C.AffiliateID = A.AffiliateID   " & vbCrLf & _
                              "   								WHERE c.price IS NOT null " & vbCrLf & _
                              "   							Group by A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,C.currcls ) SumPasiRec   " & vbCrLf & _
                              "   						  ON SumPasiRec.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "    							AND SumPasiRec.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "    							AND SumPasiRec.PONo = POD.PONo    " & vbCrLf & _
                              "    							AND SumPasiRec.KanbanNo = KD.KanbanNo " & vbCrLf & _
                              "    							AND SumPasiRec.SuratJalanNo = DSM.SuratJalanNo    " & vbCrLf & _
                              "     				         LEFT JOIN   " & vbCrLf & _
                              "     				         (select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,tot = sum(RecQty * isnull(C.Price,0)), C.currcls   "

            ls_SQL = ls_SQL + "   							    From ReceiveAffiliate_Detail A   with(nolock) Left join ReceiveAffiliate_Master B    with(nolock)  " & vbCrLf & _
                              "   							        ON A.SuratJalanNo = B.SuratJalanNo   " & vbCrLf & _
                              "   							    AND A.SupplierID = B.SupplierID   " & vbCrLf & _
                              "   							    Left Join PO_detail D   with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo   " & vbCrLf & _
                              "   							    Left Join MS_Price C   with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate   " & vbCrLf & _
                              "   									  AND  C.AffiliateID = A.AffiliateID   " & vbCrLf & _
                              "   							Group by A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,C.currcls ) SumAffRec   " & vbCrLf & _
                              "   						  ON SumAffRec.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "    							AND SumAffRec.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "    							AND SumAffRec.PONo = POD.PONo    " & vbCrLf & _
                              "    							AND SumAffRec.KanbanNo = KD.KanbanNo    "

            ls_SQL = ls_SQL + "                          LEFT JOIn MS_CurrCls MCPasi   with(nolock) ON MCPasi.CurrCls = SumPasiRec.CurrCls   " & vbCrLf & _
                              "                          LEFT JOIn MS_CurrCls MCAff   with(nolock) ON MCAff.CurrCls = SumAffRec.CurrCls   " & vbCrLf & _
                              "                          LEFT JOIN (select InvoiceNo,SupplierID,AffiliateID, KanbanNo, PONo, A = sum(InvAmount) from  " & vbCrLf & _
                              " 		                                InvoiceSupplier_Detail INVSD   with(nolock) Group by SupplierID,AffiliateID, KanbanNo, PONo, InvoiceNo)IVD " & vbCrLf & _
                              "                                      ON IVD.SupplierID = KD.SupplierID   " & vbCrLf & _
                              "                                      AND IVD.AffiliateID = INVSD.AffiliateID   " & vbCrLf & _
                              "                                      AND IVD.KanbanNo = INVSD.kanbanNo   " & vbCrLf & _
                              "                                      AND IVD.PONo = INVSD.PONo  " & vbCrLf & _
                              "                                      AND IVD.InvoiceNo = INVSD.InvoiceNo " & vbCrLf & _
                              "    		           --WHERE POD.AffiliateID = 'JAI' AND POD.pono='PO20150501-KMK '    " & vbCrLf & _
                              "                  WHERE POD.PONO <> ''  " & vbCrLf

            '20150714

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND DSM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            If rbinvoice.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(INVSM.InvoiceNo, '') <> '' " & vbCrLf
            ElseIf rbinvoice.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(INVSM.InvoiceNo,'') = '' " & vbCrLf
            End If

            'If rbreceiving.Value = "YES" Then
            '    ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) > 0 " & vbCrLf
            'ElseIf rbreceiving.Value = "NO" Then
            '    ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) = 0 " & vbCrLf
            'End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND INVSM.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If txtsupinvno.Text <> "" Then
                ls_SQL = ls_SQL + " AND INVSM.InvoiceNo Like '%" & Trim(txtsupinvno.Text) & "%'" & vbCrLf
            End If

            If cbosupplier.Text <> "" And cbosupplier.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND POM.SupplierID = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_SQL = ls_SQL + " AND POM.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND POD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
            End If


            'ls_SQL = ls_SQL + "   				) hdr   " & vbCrLf & _
            '                  "   		   UNION ALL   " & vbCrLf & _
            '                  "   		  --DETAIL   " & vbCrLf & _
            '                  "   		  SELECT DISTINCT url = '',coldetail = '',no = '',   " & vbCrLf

            'ls_SQL = ls_SQL + "   				 period = '',     " & vbCrLf & _
            '                  "   				 affiliatecode = '',   " & vbCrLf & _
            '                  "   				 affiliatename = '',   " & vbCrLf & _
            '                  "   				 pono = '',  " & vbCrLf & _
            '                  "   				 suppliercode = '',   " & vbCrLf & _
            '                  "   				 suppliername = '',   " & vbCrLf & _
            '                  "   				 kanbancls = '',   " & vbCrLf & _
            '                  "   				 kanbanno = '',   " & vbCrLf & _
            '                  "   				 suppplandeldate = '',   " & vbCrLf & _
            '                  "   				 suppdeldate = '',   " & vbCrLf & _
            '                  "   				 suppsj = '',   " & vbCrLf

            'ls_SQL = ls_SQL + "   				 pasideldate = '',   " & vbCrLf & _
            '                  "   				 pasisj = '',   " & vbCrLf & _
            '                  "   				 partno = POD.PartNo,   " & vbCrLf & _
            '                  "   				 partname = MP.PartName,   " & vbCrLf & _
            '                  "   				 uom = MU.Description,   " & vbCrLf & _
            '                  "   				 suppdelqty = Convert(char,convert(numeric(9,0),ISNULL(DSD.DOQty,'0'))),   " & vbCrLf & _
            '                  "   				 pasirecqty = Convert(char,convert(numeric(9,0),Isnull(RPD.GoodRecQty,0))),  " & vbCrLf & _
            '                  "   				 remrecqty = '',  " & vbCrLf & _
            '                  "   				 suppinvqty = Convert(char,convert(numeric(9,0),Isnull(INVSD.INVQty,0))),   " & vbCrLf & _
            '                  "   				 suppinvno = '',   " & vbCrLf & _
            '                  "   				 suppinvdate = '',  " & vbCrLf

            'ls_SQL = ls_SQL + "   				 pasireccurr = '',  " & vbCrLf & _
            '                  "   				 pasirecamount = '',   " & vbCrLf & _
            '                  "   				 suppinvcurr = '',           " & vbCrLf & _
            '                  "                   suppinvamount= '',   " & vbCrLf & _
            '                  "   				 sortpono = POD.PONo, sortkanbanno = ISNULL(KD.KanbanNo,''), sortheader = 1,   " & vbCrLf & _
            '                  "                  Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''), " & vbCrLf & _
            '                  "                  FInvNo = isnull(INVSM.InvoiceNo,''),  " & vbCrLf & _
            '                  "                  FDiffQty = 0, " & vbCrLf & _
            '                  "                  FSJ = ISNULL(DSD.SuratJalanNo,''),  " & vbCrLf & _
            '                  "                  FInv = isnull(INVSM.InvoiceNo,''),  " & vbCrLf & _
            '                  "                  Fsupp = POM.SupplierID,  " & vbCrLf & _
            '                  "                  FAff = POM.AffiliateID,  " & vbCrLf & _
            '                  "                  FPO = POD.PONo " & vbCrLf & _
            '                  "   			FROM PO_DETAIL POD with(nolock)   " & vbCrLf & _
            '                  "   				 LEFT JOIN PO_Master POM with(nolock) ON POM.AffiliateID =POD.AffiliateID   " & vbCrLf & _
            '                  "   					AND POM.SupplierID =POD.SupplierID   " & vbCrLf & _
            '                  "   					AND POM.PONO =POD.PONO   " & vbCrLf & _
            '                  "   				 LEFT JOIN Kanban_Detail KD with(nolock) ON KD.AffiliateID =POD.AffiliateID   " & vbCrLf & _
            '                  "   					AND KD.SupplierID =POD.SupplierID   " & vbCrLf

            'ls_SQL = ls_SQL + "   					AND KD.PONO =POD.PONO   " & vbCrLf & _
            '                  "   					AND KD.PartNo =POD.PartNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN Kanban_Master KM with(nolock) ON KD.AffiliateID =KM.AffiliateID   " & vbCrLf & _
            '                  "   					AND KD.SupplierID =KM.SupplierID   " & vbCrLf & _
            '                  "   					AND KD.KanbanNo =KM.KanbanNo   " & vbCrLf & _
            '                  "                       AND KD.DeliveryLocationCode = KM.DeliveryLocationCode   " & vbCrLf & _
            '                  "   				 LEFT JOIN DOSupplier_Detail DSD with(nolock) ON KD.AffiliateID =DSD.AffiliateID   " & vbCrLf & _
            '                  "   					AND KD.SupplierID =DSD.SupplierID   " & vbCrLf & _
            '                  "   					AND KD.PONO =DSD.PONO   " & vbCrLf & _
            '                  "   					AND KD.PartNo =DSD.PartNo   " & vbCrLf & _
            '                  "   					AND KD.KanbanNo =DSD.KanbanNo   " & vbCrLf

            'ls_SQL = ls_SQL + "   				 LEFT JOIN DOSupplier_Master DSM with(nolock) ON DSM.AffiliateID =DSD.AffiliateID   " & vbCrLf & _
            '                  "   					AND DSM.SupplierID =DSD.SupplierID   " & vbCrLf & _
            '                  "   					AND DSM.SuratJalanNo =DSD.SuratJalanNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN DOPASI_Detail DPD with(nolock) ON KD.AffiliateID =DPD.AffiliateID   " & vbCrLf & _
            '                  "   					AND KD.SupplierID =DPD.SupplierID   " & vbCrLf & _
            '                  "   					AND KD.PONO =DPD.PONO   " & vbCrLf & _
            '                  "   					AND KD.PartNo =DPD.PartNo   " & vbCrLf & _
            '                  "   					AND KD.KanbanNo =DPD.KanbanNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN DOPASI_Master DPM with(nolock) ON DPM.AffiliateID =DPD.AffiliateID   " & vbCrLf & _
            '                  "   					AND DPM.SupplierID =DPD.SupplierID   " & vbCrLf & _
            '                  "   					AND DPM.SuratJalanNo =DPD.SuratJalanNo   " & vbCrLf

            'ls_SQL = ls_SQL + "   				 LEFT JOIN ReceivePASI_Detail RPD with(nolock) ON RPD.AffiliateID = DPM.AffiliateID   " & vbCrLf & _
            '                  "   					AND RPD.SupplierID = DPM.SupplierID   " & vbCrLf & _
            '                  "   					AND RPD.PONo = POD.PONo   " & vbCrLf & _
            '                  "   					AND RPD.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "   					AND RPD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN ReceiveAffiliate_Detail RAD with(nolock) ON RAD.AffiliateID = KD.AffiliateID   " & vbCrLf & _
            '                  "   					AND RAD.SupplierID = KD.SupplierID   " & vbCrLf & _
            '                  "   					AND RAD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "   					AND RAD.PONo = KD.PONo   " & vbCrLf & _
            '                  "   					AND RAD.PartNo = KD.PartNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN ReceiveAffiliate_Master RAM with(nolock) ON RAM.AffiliateID = RAD.AffiliateID   " & vbCrLf

            'ls_SQL = ls_SQL + "   					AND RAM.SupplierID = RAD.SupplierID   " & vbCrLf & _
            '                  "   					AND RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "   				 LEFT JOIN InvoiceSupplier_Detail INVSD with(nolock) ON INVSD.SupplierID = KD.SupplierID  " & vbCrLf & _
            '                  "   							AND INVSD.AffiliateID = KD.AffiliateID  " & vbCrLf & _
            '                  "   							AND INVSD.SuratJalanNo = DSD.SuratJalanNo  " & vbCrLf & _
            '                  "   							AND INVSD.KanbanNo = KD.kanbanNo  " & vbCrLf & _
            '                  "   							AND INVSD.PONo = KD.PONo  " & vbCrLf & _
            '                  "   							AND INVSD.PartNo = KD.PartNo  " & vbCrLf & _
            '                  "  						 LEFT JOIN InvoiceSupplier_Master INVSM with(nolock) ON INVSM.InvoiceNo = INVSD.InvoiceNo  " & vbCrLf & _
            '                  "  							AND INVSM.SupplierID = INVSD.SupplierID  " & vbCrLf & _
            '                  "  							AND INVSM.AffiliateID = INVSD.AffiliateID  " & vbCrLf

            'ls_SQL = ls_SQL + "  							AND INVSM.SuratJalanNo = INVSD.SuratJalanNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN MS_Parts MP with(nolock) ON MP.PartNo = POD.PartNo   " & vbCrLf & _
            '                  "   				 LEFT JOIN MS_UnitCls MU with(nolock) ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
            '                  "                    LEFT JOIN    " & vbCrLf & _
            '                  "    				 (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(RecQty),0) + ISNULL(SUM(DefectQty),0)   " & vbCrLf & _
            '                  "    				    FROM ReceiveAffiliate_Detail with(nolock)   " & vbCrLf & _
            '                  "    				   GROUP BY PONo, KanbanNo   " & vbCrLf & _
            '                  "    				 ) SumKanban ON SumKanban.PONo = KD.PONo AND SumKanban.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "    				 LEFT JOIN    " & vbCrLf & _
            '                  "    				 (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)   " & vbCrLf & _
            '                  "    				    FROM DOSupplier_Detail with(nolock)    " & vbCrLf

            'ls_SQL = ls_SQL + "    				   GROUP BY PONo, KanbanNo   " & vbCrLf & _
            '                  "    				 ) SumDSD ON SumDSD.PONo = KD.PONo AND SumDSD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "    				 LEFT JOIN    " & vbCrLf & _
            '                  "    				 (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)   " & vbCrLf & _
            '                  "    				    FROM DOPASI_Detail  with(nolock)  " & vbCrLf & _
            '                  "    				   GROUP BY PONo, KanbanNo   " & vbCrLf & _
            '                  "    				 ) SumDPD ON SumDPD.PONo = KD.PONo AND SumDPD.KanbanNo = KD.KanbanNo   " & vbCrLf & _
            '                  "   		   --WHERE POD.AffiliateID = 'JAI' AND POD.pono='PO20150501-KMK'   " & vbCrLf & _
            '                  "   		   WHERE POD.PONO <> '' " & vbCrLf & _
            '                  "  "

            '20150714
            ls_SQL = ls_SQL + " ) hdr    " & vbCrLf & _
                  "    		   UNION ALL    " & vbCrLf & _
                  "    		  --DETAIL    " & vbCrLf & _
                  "    		  SELECT DISTINCT url = '',coldetail = '',no = '',    " & vbCrLf & _
                  "    				 period = '',      " & vbCrLf & _
                  "    				 affiliatecode = '',    " & vbCrLf & _
                  "    				 affiliatename = '',    " & vbCrLf & _
                  "    				 pono = '',   " & vbCrLf & _
                  "    				 suppliercode = '',    " & vbCrLf & _
                  "    				 suppliername = '',    " & vbCrLf & _
                  "    				 kanbancls = '',    "

            ls_SQL = ls_SQL + "    				 kanbanno = '',    " & vbCrLf & _
                              "    				 suppplandeldate = '',    " & vbCrLf & _
                              "    				 suppdeldate = '',    " & vbCrLf & _
                              "    				 suppsj = '',    " & vbCrLf & _
                              "    				 pasideldate = '',    " & vbCrLf & _
                              "    				 pasisj = '',    " & vbCrLf & _
                              "    				 partno = POD.PartNo,    " & vbCrLf & _
                              "    				 partname = MP.PartName,    " & vbCrLf & _
                              "    				 uom = MU.Description,    " & vbCrLf & _
                              "    				 suppdelqty = Convert(char,convert(numeric(9,0),ISNULL(DSD.DOQty,'0'))),    " & vbCrLf & _
                              "    				 pasirecqty = Convert(char,convert(numeric(9,0),Isnull(RPD.GoodRecQty,0))),   "

            ls_SQL = ls_SQL + "    				 remrecqty = '',   " & vbCrLf & _
                              "    				 suppinvqty = Convert(char,convert(numeric(9,0),Isnull(INVSD.INVQty,0))),    " & vbCrLf & _
                              "    				 suppinvno = '',    " & vbCrLf & _
                              "    				 suppinvdate = '',   " & vbCrLf & _
                              "    				 pasireccurr = '',   " & vbCrLf & _
                              "    				 pasirecamount = '',    " & vbCrLf & _
                              "    				 /*suppinvcurr = '',            " & vbCrLf & _
                              "                    suppinvamount= '',*/    " & vbCrLf & _
                              "    				 sortpono = POD.PONo, sortkanbanno = ISNULL(KD.KanbanNo,''), sortheader = 1,    " & vbCrLf & _
                              "                   Fsuppdeldate = ISNULL(CONVERT(CHAR,KM.KanbanDate,106),''),  " & vbCrLf & _
                              "                   FInvNo = isnull(INVSM.InvoiceNo,''),   "

            ls_SQL = ls_SQL + "                   FDiffQty = 0,  " & vbCrLf & _
                              "                   FSJ = ISNULL(DSD.SuratJalanNo,''),   " & vbCrLf & _
                              "                   FInv = isnull(INVSM.InvoiceNo,''),   " & vbCrLf & _
                              "                   Fsupp = POM.SupplierID,   " & vbCrLf & _
                              "                   FAff = POM.AffiliateID,   " & vbCrLf & _
                              "                   FPO = POD.PONo  " & vbCrLf & _
                              "                   ,H_SupSj = DSM.SuratJalanNO " & vbCrLf & _
                              "                   , H_SupINV = isnull(INVSM.InvoiceNo,'') " & vbCrLf & _
                              "    			FROM PO_DETAIL POD with(nolock)    " & vbCrLf & _
                              "    				 LEFT JOIN PO_Master POM with(nolock) ON POM.AffiliateID =POD.AffiliateID    " & vbCrLf & _
                              "    					AND POM.SupplierID =POD.SupplierID    " & vbCrLf & _
                              "    					AND POM.PONO =POD.PONO    "

            ls_SQL = ls_SQL + "    				 LEFT JOIN Kanban_Detail KD with(nolock) ON KD.AffiliateID =POD.AffiliateID    " & vbCrLf & _
                              "    					AND KD.SupplierID =POD.SupplierID    " & vbCrLf & _
                              "    					AND KD.PONO =POD.PONO    " & vbCrLf & _
                              "    					AND KD.PartNo =POD.PartNo    " & vbCrLf & _
                              "    				 LEFT JOIN Kanban_Master KM with(nolock) ON KD.AffiliateID =KM.AffiliateID    " & vbCrLf & _
                              "    					AND KD.SupplierID =KM.SupplierID    " & vbCrLf & _
                              "    					AND KD.KanbanNo =KM.KanbanNo    " & vbCrLf & _
                              "                        AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
                              "    				 INNER JOIN DOSupplier_Detail DSD with(nolock) ON KD.AffiliateID =DSD.AffiliateID    " & vbCrLf & _
                              "    					AND KD.SupplierID =DSD.SupplierID    " & vbCrLf & _
                              "    					AND KD.PONO =DSD.PONO    "

            ls_SQL = ls_SQL + "    					AND KD.PartNo =DSD.PartNo    " & vbCrLf & _
                              "    					AND KD.KanbanNo =DSD.KanbanNo    " & vbCrLf & _
                              "    				 INNER JOIN DOSupplier_Master DSM with(nolock) ON DSM.AffiliateID =DSD.AffiliateID    " & vbCrLf & _
                              "    					AND DSM.SupplierID =DSD.SupplierID    " & vbCrLf & _
                              "    					AND DSM.SuratJalanNo =DSD.SuratJalanNo    " & vbCrLf & _
                              "    				 LEFT JOIN DOPASI_Detail DPD with(nolock) ON KD.AffiliateID =DPD.AffiliateID    " & vbCrLf & _
                              "    					AND KD.SupplierID =DPD.SupplierID    " & vbCrLf & _
                              "    					AND KD.PONO =DPD.PONO    " & vbCrLf & _
                              "    					AND KD.PartNo =DPD.PartNo    " & vbCrLf & _
                              "    					AND KD.KanbanNo =DPD.KanbanNo  " & vbCrLf & _
                              "    					AND DPD.SuratjalanNoSupplier = DSM.SuratJalanNo   "

            ls_SQL = ls_SQL + "    				 LEFT JOIN DOPASI_Master DPM with(nolock) ON DPM.AffiliateID =DPD.AffiliateID    " & vbCrLf & _
                              "    					AND DPM.SupplierID =DPD.SupplierID    " & vbCrLf & _
                              "    					AND DPM.SuratJalanNo =DPD.SuratJalanNo    " & vbCrLf & _
                              "    				 LEFT JOIN ReceivePASI_Detail RPD with(nolock) ON RPD.AffiliateID = DPM.AffiliateID    " & vbCrLf & _
                              "    					AND RPD.SupplierID = DPM.SupplierID    " & vbCrLf & _
                              "    					AND RPD.PONo = POD.PONo    " & vbCrLf & _
                              "    					AND RPD.PartNo = POD.PartNo    " & vbCrLf & _
                              "    					AND RPD.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "    					AND RPD.SuratjalanNo = DSM.SuratJalanNo " & vbCrLf & _
                              "    				 LEFT JOIN ReceiveAffiliate_Detail RAD with(nolock) ON RAD.AffiliateID = KD.AffiliateID    " & vbCrLf & _
                              "    					AND RAD.SupplierID = KD.SupplierID    "

            ls_SQL = ls_SQL + "    					AND RAD.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "    					AND RAD.PONo = KD.PONo    " & vbCrLf & _
                              "    					AND RAD.PartNo = KD.PartNo    " & vbCrLf & _
                              "    				 LEFT JOIN ReceiveAffiliate_Master RAM with(nolock) ON RAM.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
                              "    					AND RAM.SupplierID = RAD.SupplierID    " & vbCrLf & _
                              "    					AND RAM.SuratJalanNo = RAD.SuratJalanNo   " & vbCrLf & _
                              "    				 LEFT JOIN InvoiceSupplier_Detail INVSD with(nolock) ON INVSD.SupplierID = KD.SupplierID   " & vbCrLf & _
                              "    							AND INVSD.AffiliateID = KD.AffiliateID   " & vbCrLf & _
                              "    							AND INVSD.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
                              "    							AND INVSD.KanbanNo = KD.kanbanNo   " & vbCrLf & _
                              "    							AND INVSD.PONo = KD.PONo   "

            ls_SQL = ls_SQL + "    							AND INVSD.PartNo = KD.PartNo   " & vbCrLf & _
                              "   						 LEFT JOIN InvoiceSupplier_Master INVSM with(nolock) ON INVSM.InvoiceNo = INVSD.InvoiceNo   " & vbCrLf & _
                              "   							AND INVSM.SupplierID = INVSD.SupplierID   " & vbCrLf & _
                              "   							AND INVSM.AffiliateID = INVSD.AffiliateID   " & vbCrLf & _
                              "   							AND INVSM.SuratJalanNo = INVSD.SuratJalanNo    " & vbCrLf & _
                              "    				 LEFT JOIN MS_Parts MP with(nolock) ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "    				 LEFT JOIN MS_UnitCls MU with(nolock) ON MU.UnitCls = MP.UnitCls    " & vbCrLf & _
                              "                     LEFT JOIN     " & vbCrLf & _
                              "     				 (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(RecQty),0) + ISNULL(SUM(DefectQty),0)    " & vbCrLf & _
                              "     				    FROM ReceiveAffiliate_Detail with(nolock)    " & vbCrLf & _
                              "     				   GROUP BY PONo, KanbanNo    "

            ls_SQL = ls_SQL + "     				 ) SumKanban ON SumKanban.PONo = KD.PONo AND SumKanban.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "     				 LEFT JOIN     " & vbCrLf & _
                              "     				 (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)    " & vbCrLf & _
                              "     				    FROM DOSupplier_Detail with(nolock)     " & vbCrLf & _
                              "     				   GROUP BY PONo, KanbanNo    " & vbCrLf & _
                              "     				 ) SumDSD ON SumDSD.PONo = KD.PONo AND SumDSD.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "     				 LEFT JOIN     " & vbCrLf & _
                              "     				 (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)    " & vbCrLf & _
                              "     				    FROM DOPASI_Detail  with(nolock)   " & vbCrLf & _
                              "     				   GROUP BY PONo, KanbanNo    " & vbCrLf & _
                              "     				 ) SumDPD ON SumDPD.PONo = KD.PONo AND SumDPD.KanbanNo = KD.KanbanNo    "

            ls_SQL = ls_SQL + "    		   --WHERE POD.AffiliateID = 'JAI' AND POD.pono='PO20150501-KMK'    " & vbCrLf & _
                              "    		   WHERE POD.PONO <> '' --AND isnull(INVSM.InvoiceNo,'') <> '' " & vbCrLf

            '20150714

            If checkbox1.Checked = True Then
                ls_SQL = ls_SQL + " AND DSM.DeliveryDate between '" & Format(dt1.Value, "yyyy-MM-dd") & "' AND '" & Format(dt2.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            If rbinvoice.Value = "YES" Then
                ls_SQL = ls_SQL + "AND isnull(INVSM.InvoiceNo, '') <> '' " & vbCrLf
            ElseIf rbinvoice.Value = "NO" Then
                ls_SQL = ls_SQL + " AND isnull(INVSM.InvoiceNo,'') = '' " & vbCrLf
            End If

            'If rbreceiving.Value = "YES" Then
            '    ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) > 0 " & vbCrLf
            'ElseIf rbreceiving.Value = "NO" Then
            '    ls_SQL = ls_SQL + " AND convert(numeric,H_REMAINING) = 0 " & vbCrLf
            'End If

            If txtsj.Text <> "" Then
                ls_SQL = ls_SQL + " AND INVSM.SuratJalanNo Like '%" & Trim(txtsj.Text) & "%'" & vbCrLf
            End If

            If txtsupinvno.Text <> "" Then
                ls_SQL = ls_SQL + " AND INVSM.InvoiceNo Like '%" & Trim(txtsupinvno.Text) & "%'" & vbCrLf
            End If

            If cbosupplier.Text <> "" And cbosupplier.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " AND POM.SupplierID = '" & Trim(cbosupplier.Text) & "'" & vbCrLf
            End If

            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_SQL = ls_SQL + " AND POM.AffiliateID = '" & Trim(cboaffiliate.Text) & "'" & vbCrLf
            End If

            If txtpono.Text <> "" Then
                ls_SQL = ls_SQL + " AND POD.PONo  Like '%" & Trim(txtpono.Text) & "%'"
            End If
            
            ls_SQL = ls_SQL + " ) SPDC ORDER BY SortPONo, SortKanbanNo,H_SupSJ, H_supINV, SortHeader "

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
        If e.GetValue("partno").ToString = "" Then e.Row.BackColor = Drawing.Color.LightGray
    End Sub
End Class