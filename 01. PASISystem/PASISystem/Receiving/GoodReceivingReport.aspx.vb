Imports System.Data
Imports System.Data.SqlClient

Public Class GoodReceivingReport
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = "", urlBack As String = ""
#End Region
    
#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Report As New rptGoodReceivingReport

        'ls_SQL = ""
        'ls_SQL = ls_SQL + " SELECT ColNo = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, KanbanCls, KanbanNo)), " & vbCrLf & _
        '                  "        ReceiveDate, SupplierName, SupplierSJNo, SupplierDeliveryDate, SupplierPlanDeliveryDate, " & vbCrLf & _
        '                  "        JenisArmada, NoPol, DeliveryTo, PASISJNo, PASIDeliveryDate, DriverName, TotalBox, " & vbCrLf & _
        '                  " 	   PONo, POKanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,  " & vbCrLf & _
        '                  " 	   KanbanNo, PartNo, PartName, UOM, QtyBox, " & vbCrLf & _
        '                  " 	   SupplierDeliveryQty, PASIGoodReceivingQty, PASIDefectQty, PASIDeliveryQty, GoodReceivingQty, " & vbCrLf & _
        '                  "        DefectReceivingQty, RemainingReceivingQty, ReceivingQtyBox, UnitCls, DeliveryByPASICls " & vbCrLf & _
        '                  "   FROM ( " & vbCrLf
        'ls_SQL = ls_SQL + " 		  SELECT DISTINCT " & vbCrLf & _
        '                  "                  ReceiveDate = CONVERT(CHAR,RAM.ReceiveDate,106), " & vbCrLf & _
        '                  "  				 MS.SupplierName, " & vbCrLf & _
        '                  "  				 SupplierSJNo = DSM.SuratJalanNo, " & vbCrLf & _
        '                  "  				 SupplierDeliveryDate = CONVERT(CHAR,DSM.DeliveryDate,106), " & vbCrLf & _
        '                  "  				 SupplierPlanDeliveryDate = CONVERT(CHAR,KM.KanbanDate,106), " & vbCrLf & _
        '                  "  				 RAM.JenisArmada, " & vbCrLf & _
        '                  "  				 RAM.NoPol, " & vbCrLf & _
        '                  "  				 DeliveryTo = MDP.DeliveryLocationName, " & vbCrLf & _
        '                  "  				 PASISJNo = DSM.SuratJalanNo, " & vbCrLf & _
        '                  "  				 PASIDeliveryDate = CONVERT(CHAR,DSM.DeliveryDate,106), " & vbCrLf & _
        '                  "  				 DriverName = RTRIM(RAM.DriverName) + ' (' + RTRIM(RAM.DriverContact) + ')', " & vbCrLf & _
        '                  "  				 TotalBox = RAM.TotalBox, " & vbCrLf & _
        '                  "                  PONo = POD.PONo, " & vbCrLf & _
        '                  " 				 KanbanCls = POD.KanbanCls, " & vbCrLf & _
        '                  " 				 KanbanNo = KD.KanbanNo, " & vbCrLf & _
        '                  " 				 PartNo = POD.PartNo, " & vbCrLf & _
        '                  " 				 PartName = MP.PartName, " & vbCrLf & _
        '                  " 				 UOM = MU.Description, " & vbCrLf & _
        '                  " 				 QtyBox = MPM.QtyBox, " & vbCrLf & _
        '                  " 				 SupplierDeliveryQty = DSD.DOQty, " & vbCrLf & _
        '                  " 				 PASIGoodReceivingQty = CONVERT(CHAR,ISNULL(RPD.GoodRecQty,'0')), " & vbCrLf
        'ls_SQL = ls_SQL + " 				 PASIDefectQty = CONVERT(CHAR,ISNULL(RPD.DefectRecQty,'0')), " & vbCrLf & _
        '                  " 				 PASIDeliveryQty = CONVERT(CHAR,ISNULL(DPD.DOQty,'0')), " & vbCrLf & _
        '                  " 				 GoodReceivingQty = CASE WHEN POM.DeliveryByPASICls = '1' THEN RTRIM(CONVERT(CHAR,COALESCE(RAD.RecQty,DPD.DOQty,'0'))) " & vbCrLf & _
        '                  "                                          ELSE RTRIM(CONVERT(CHAR,COALESCE(RAD.RecQty,DSD.DOQty,'0'))) END, " & vbCrLf & _
        '                  " 				 DefectReceivingQty = CONVERT(CHAR,ISNULL(RAD.DefectQty,'0')), " & vbCrLf & _
        '                  " 				 RemainingReceivingQty = CONVERT(CHAR,CASE WHEN ISNULL(POM.DeliveryByPASICls,'0') = '1' THEN RTRIM((ISNULL(DPD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0)))) " & vbCrLf & _
        '                  " 											  WHEN ISNULL(POM.DeliveryByPASICls,'0') = '0' THEN RTRIM((ISNULL(DSD.DOQty,0) - (ISNULL(RAD.RecQty,0) + ISNULL(RAD.DefectQty,0)))) END), " & vbCrLf & _
        '                  " 				 ReceivingQtyBox = CEILING(((CASE WHEN POM.DeliveryByPASICls = '1' THEN COALESCE(RAD.RecQty,DPD.DOQty,0) " & vbCrLf & _
        '                  "                                                   ELSE COALESCE(RAD.RecQty,DSD.DOQty,0) END) + ISNULL(RAD.DefectQty,0)) / MPM.QtyBox), " & vbCrLf & _
        '                  " 				 UnitCls = KD.UnitCls, " & vbCrLf & _
        '                  " 				 POM.DeliveryByPASICls " & vbCrLf
        'ls_SQL = ls_SQL + " 			FROM PO_DETAIL POD  " & vbCrLf & _
        '                  " 				 LEFT JOIN PO_Master POM ON POM.AffiliateID =POD.AffiliateID " & vbCrLf & _
        '                  " 					AND POM.SupplierID =POD.SupplierID " & vbCrLf & _
        '                  " 					AND POM.PONO =POD.PONO " & vbCrLf & _
        '                  " 				 LEFT JOIN Kanban_Detail KD ON KD.AffiliateID =POD.AffiliateID " & vbCrLf & _
        '                  " 					AND KD.SupplierID =POD.SupplierID " & vbCrLf & _
        '                  " 					AND KD.PONO =POD.PONO " & vbCrLf
        'ls_SQL = ls_SQL + " 					AND KD.PartNo =POD.PartNo " & vbCrLf & _
        '                  " 				 LEFT JOIN Kanban_Master KM ON KD.AffiliateID =KM.AffiliateID " & vbCrLf & _
        '                  " 					AND KD.SupplierID =KM.SupplierID " & vbCrLf & _
        '                  " 					AND KD.KanbanNo =KM.KanbanNo " & vbCrLf & _
        '                  "                     AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
        '                  " 				 LEFT JOIN DOSupplier_Detail DSD ON KD.AffiliateID =DSD.AffiliateID " & vbCrLf & _
        '                  " 					AND KD.SupplierID =DSD.SupplierID " & vbCrLf & _
        '                  " 					AND KD.PONO =DSD.PONO " & vbCrLf & _
        '                  " 					AND KD.PartNo =DSD.PartNo " & vbCrLf & _
        '                  " 					AND KD.KanbanNo =DSD.KanbanNo " & vbCrLf & _
        '                  " 				 LEFT JOIN DOSupplier_Master DSM ON DSM.AffiliateID =DSD.AffiliateID " & vbCrLf & _
        '                  " 					AND DSM.SupplierID =DSD.SupplierID " & vbCrLf
        'ls_SQL = ls_SQL + " 					AND DSM.SuratJalanNo =DSD.SuratJalanNo " & vbCrLf & _
        '                  " 				 LEFT JOIN DOPASI_Detail DPD ON KD.AffiliateID =DPD.AffiliateID " & vbCrLf & _
        '                  " 					AND KD.SupplierID =DPD.SupplierID " & vbCrLf & _
        '                  " 					AND KD.PONO =DPD.PONO " & vbCrLf & _
        '                  " 					AND KD.PartNo =DPD.PartNo " & vbCrLf & _
        '                  " 					AND KD.KanbanNo =DPD.KanbanNo " & vbCrLf & _
        '                  " 				 LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID =DPD.AffiliateID " & vbCrLf & _
        '                  " 					AND DPM.SupplierID =DPD.SupplierID " & vbCrLf & _
        '                  " 					AND DPM.SuratJalanNo =DPD.SuratJalanNo " & vbCrLf & _
        '                  " 				 LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = KD.AffiliateID " & vbCrLf & _
        '                  " 					AND RPD.SupplierID = KD.SupplierID " & vbCrLf
        'ls_SQL = ls_SQL + " 					AND RPD.PONo = KD.PONo " & vbCrLf & _
        '                  " 					AND RPD.PartNo = KD.PartNo " & vbCrLf & _
        '                  " 					AND RPD.KanbanNo = KD.KanbanNo " & vbCrLf & _
        '                  " 				 LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = KD.AffiliateID " & vbCrLf & _
        '                  " 					AND RAD.SupplierID = KD.SupplierID " & vbCrLf & _
        '                  " 					AND RAD.KanbanNo = KD.KanbanNo " & vbCrLf & _
        '                  " 					AND RAD.PONo = KD.PONo " & vbCrLf & _
        '                  " 					AND RAD.PartNo = KD.PartNo " & vbCrLf & _
        '                  " 				 LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
        '                  " 					AND RAM.SupplierID = RAD.SupplierID " & vbCrLf & _
        '                  " 					AND RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
        '                  " 				 LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
        '                  " 				 LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
        '                  "                  LEFT JOIN MS_Supplier MS ON MS.SupplierID = RAM.SupplierID " & vbCrLf & _
        '                  "                  LEFT JOIN MS_DeliveryPlace MDP ON MDP.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
        '                  "                     AND MDP.AffiliateID = KM.AffiliateID " & vbCrLf & _
        '" 		   WHERE RAM.AffiliateID = '" & Session("E02AffiliateID") & "' " & vbCrLf & _
        '"              AND RAM.SuratJalanNo = '" & Session("E02SupplierSJNo") & "'" & vbCrLf & _
        '"              AND KD.KanbanNo = '" & Session("E02KanbanNo") & "'" & vbCrLf
        'ls_SQL = ls_SQL + "        ) RecEntry "

        ls_SQL = "  SELECT ColNo = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PONo, KanbanCls, KanbanNo)),  " & vbCrLf & _
                  "         ReceiveBy, ReceiveDate, SupplierName, SupplierSJNo, SupplierDeliveryDate, SupplierPlanDeliveryDate,  " & vbCrLf & _
                  "         JenisArmada, NoPol, DeliveryTo, PASISJNo, PASIDeliveryDate, DriverName, TotalBox,  " & vbCrLf & _
                  "  	   PONo, POKanban = CASE WHEN ISNULL(KanbanCls,'0') = '1' THEN 'YES' ELSE 'NO' END,   " & vbCrLf & _
                  "  	   KanbanNo, PartNo, PartName, UOM, QtyBox,  " & vbCrLf & _
                  "  	   SupplierDeliveryQty, PASIGoodReceivingQty, PASIDefectQty, " & vbCrLf & _
                  "        RemainingReceivingQty, ReceivingQtyBox, UnitCls, DeliveryByPASICls,  Diff = CASE WHEN RemainingReceivingQty > 0 THEN '*' ELSE '' END   " & vbCrLf & _
                  "    FROM (  " & vbCrLf & _
                  "  		  SELECT DISTINCT  " & vbCrLf & _
                  "                   ReceiveDate = CONVERT(CHAR,RAM.ReceiveDate,106), ReceiveBy," & vbCrLf & _
                  "   				 MS.SupplierName,  "

        ls_SQL = ls_SQL + "   				 SupplierSJNo = DSM.SuratJalanNo,  " & vbCrLf & _
                          "   				 SupplierDeliveryDate = CONVERT(CHAR,DSM.DeliveryDate,106),  " & vbCrLf & _
                          "   				 SupplierPlanDeliveryDate = CONVERT(CHAR,KM.KanbanDate,106),  " & vbCrLf & _
                          "   				 RAM.JenisArmada,  " & vbCrLf & _
                          "   				 RAM.NoPol,  " & vbCrLf & _
                          "   				 DeliveryTo = MDP.DeliveryLocationName,  " & vbCrLf & _
                          "   				 PASISJNo = DSM.SuratJalanNo,  " & vbCrLf & _
                          "   				 PASIDeliveryDate = CONVERT(CHAR,DSM.DeliveryDate,106),  " & vbCrLf & _
                          "   				 DriverName = RTRIM(RAM.DriverName) + ' (' + RTRIM(RAM.DriverContact) + ')',  " & vbCrLf & _
                          "   				 TotalBox = RAM.TotalBox,  " & vbCrLf & _
                          "                   PONo = POD.PONo,  "

        ls_SQL = ls_SQL + "  				 KanbanCls = POD.KanbanCls,  " & vbCrLf & _
                          "  				 KanbanNo = KD.KanbanNo,  " & vbCrLf & _
                          "  				 PartNo = POD.PartNo,  " & vbCrLf & _
                          "  				 PartName = MP.PartName,  " & vbCrLf & _
                          "  				 UOM = MU.Description,  " & vbCrLf & _
                          "  				 QtyBox = ISNULL(POD.POQtyBox,MPM.Qtybox),  " & vbCrLf & _
                          "  				 SupplierDeliveryQty = DSD.DOQty,  " & vbCrLf & _
                          "  				 PASIGoodReceivingQty = ISNULL(RPD.GoodRecQty,'0'),  " & vbCrLf & _
                          "  				 PASIDefectQty = ISNULL(RPD.DefectRecQty,'0'),  " & vbCrLf & _
                          "  				 RemainingReceivingQty = (ISNULL(DSD.DOQty, 0) - (ISNULL(RPD.GoodRecQty, 0) + ISNULL(RPD.DefectRecQty, 0))),  " & vbCrLf & _
                          "  				 ReceivingQtyBox = CEILING( RPD.GoodRecQty / ISNULL(POD.POQtyBox,MPM.Qtybox)),  " & vbCrLf & _
                          "  				 UnitCls = KD.UnitCls,  " & vbCrLf & _
                          "  				 POM.DeliveryByPASICls  " & vbCrLf & _
                          "  			FROM PO_DETAIL POD   " & vbCrLf & _
                          "  				 LEFT JOIN PO_Master POM ON POM.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                          "  					AND POM.SupplierID =POD.SupplierID  "

        ls_SQL = ls_SQL + "  					AND POM.PONO =POD.PONO  " & vbCrLf & _
                          "  				 LEFT JOIN Kanban_Detail KD ON KD.AffiliateID =POD.AffiliateID  " & vbCrLf & _
                          "  					AND KD.SupplierID =POD.SupplierID  " & vbCrLf & _
                          "  					AND KD.PONO =POD.PONO  " & vbCrLf & _
                          "  					AND KD.PartNo =POD.PartNo  " & vbCrLf & _
                          "  				 LEFT JOIN Kanban_Master KM ON KD.AffiliateID =KM.AffiliateID  " & vbCrLf & _
                          "  					AND KD.SupplierID =KM.SupplierID  " & vbCrLf & _
                          "  					AND KD.KanbanNo =KM.KanbanNo  " & vbCrLf & _
                          "                      AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                          "  				 LEFT JOIN DOSupplier_Detail DSD ON KD.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                          "  					AND KD.SupplierID =DSD.SupplierID  "

        ls_SQL = ls_SQL + "  					AND KD.PONO =DSD.PONO  " & vbCrLf & _
                          "  					AND KD.PartNo =DSD.PartNo  " & vbCrLf & _
                          "  					AND KD.KanbanNo =DSD.KanbanNo  " & vbCrLf & _
                          "  				 LEFT JOIN DOSupplier_Master DSM ON DSM.AffiliateID =DSD.AffiliateID  " & vbCrLf & _
                          "  					AND DSM.SupplierID =DSD.SupplierID  " & vbCrLf & _
                          "  					AND DSM.SuratJalanNo =DSD.SuratJalanNo  " & vbCrLf & _
                          "  				 --LEFT JOIN DOPASI_Detail DPD ON KD.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                          "  				 --	AND KD.SupplierID =DPD.SupplierID  " & vbCrLf & _
                          "  				 --	AND KD.PONO =DPD.PONO  " & vbCrLf & _
                          "  				 --	AND KD.PartNo =DPD.PartNo  " & vbCrLf & _
                          "  				 --	AND KD.KanbanNo =DPD.KanbanNo  "

        ls_SQL = ls_SQL + "  				 --LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID =DPD.AffiliateID  " & vbCrLf & _
                          "  				 --	AND DPM.SupplierID =DPD.SupplierID  " & vbCrLf & _
                          "  				 --	AND DPM.SuratJalanNo =DPD.SuratJalanNo  " & vbCrLf & _
                          "  				 LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                          "  					AND RPD.SupplierID = KD.SupplierID  " & vbCrLf & _
                          "  					AND RPD.PONo = KD.PONo  " & vbCrLf & _
                          "  					AND RPD.PartNo = KD.PartNo  " & vbCrLf & _
                          "  					AND RPD.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                          "                     AND RPD.SuratJalanNo = DSD.SuratJalanNo " & vbCrLf & _
                          "  				 LEFT JOIN Receivepasi_Master RAM ON RAM.AffiliateID = RPD.AffiliateID  " & vbCrLf & _
                          "  					AND RAM.SupplierID = RPD.SupplierID  " & vbCrLf & _
                          "  					AND RAM.SuratJalanNo = RPD.SuratJalanNo  " & vbCrLf & _
                          "  				 LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                          "                  LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                          "  				 LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls  " & vbCrLf & _
                          "                   LEFT JOIN MS_Supplier MS ON MS.SupplierID = RAM.SupplierID  " & vbCrLf & _
                          "                   LEFT JOIN MS_DeliveryPlace MDP ON MDP.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                          "                      AND MDP.AffiliateID = KM.AffiliateID  " & vbCrLf & _
                          " 		   WHERE RAM.AffiliateID = '" & Session("E02AffiliateID") & "' " & vbCrLf & _
                          "              AND RAM.SuratJalanNo = '" & Session("E02SupplierSJNo") & "'" & vbCrLf & _
                          "              --AND KD.KanbanNo = '" & Session("E02KanbanNo") & "'" & vbCrLf & _
                          "         ) RecEntry  " & vbCrLf & _
                          "  "


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            Report.DataSource = ds.Tables(0)
            Viewer.Report = Report
            Viewer.DataBind()
        End Using
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")

        urlBack = "~/Receiving/ReceivingEntry.aspx?prm=" & Session("E02ParamPageLoad").ToString
        Response.Redirect(urlBack)
    End Sub
#End Region
    
End Class