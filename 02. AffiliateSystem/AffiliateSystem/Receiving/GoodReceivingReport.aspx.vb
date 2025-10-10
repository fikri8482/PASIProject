Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.DataAccess.ConnectionParameters
Imports DevExpress.XtraReports.UI
Imports DevExpress.DataAccess.Sql

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
        'ls_SQL = "   SELECT    ColNo = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY PONo, KanbanCls, KanbanNo )) , " & vbCrLf & _
        '          "             ReceiveDate , " & vbCrLf & _
        '          "             JenisArmada , " & vbCrLf & _
        '          "             NoPol , " & vbCrLf & _
        '          "             DeliveryTo , " & vbCrLf & _
        '          "             PASISJNo , " & vbCrLf & _
        '          "             PASIDeliveryDate , " & vbCrLf & _
        '          "             DriverName , " & vbCrLf & _
        '          "             TotalBox , " & vbCrLf & _
        '          "             PONo , " & vbCrLf & _
        '          "             POKanban = CASE WHEN ISNULL(KanbanCls, '0') = '1' THEN 'YES' "

        'ls_SQL = ls_SQL + "                             ELSE 'NO' " & vbCrLf & _
        '                  "                        END , " & vbCrLf & _
        '                  "             KanbanNo , " & vbCrLf & _
        '                  "             PartNo , " & vbCrLf & _
        '                  "             PartName , " & vbCrLf & _
        '                  "             UOM , " & vbCrLf & _
        '                  "             QtyBox , " & vbCrLf & _
        '                  "             SUM(SupplierDeliveryQty) SupplierDeliveryQty , " & vbCrLf & _
        '                  "             SUM(SupplierDeliveryQty) SupplierDeliveryQty , " & vbCrLf & _
        '                  "             SUM(PASIGoodReceivingQty) PASIGoodReceivingQty, " & vbCrLf & _
        '                  "             SUM(PASIDefectQty) PASIDefectQty , " & vbCrLf & _
        '                  "             SUM(PASIDeliveryQty) PASIDeliveryQty , "

        'ls_SQL = ls_SQL + "             SUM(GoodReceivingQty) GoodReceivingQty , " & vbCrLf & _
        '                  "             SUM(DefectReceivingQty) DefectReceivingQty , " & vbCrLf & _
        '                  "             SUM(RemainingReceivingQty) RemainingReceivingQty , " & vbCrLf & _
        '                  "             SUM(ReceivingQtyBox) ReceivingQtyBox , " & vbCrLf & _
        '                  "             UnitCls , " & vbCrLf & _
        '                  "             DeliveryByPASICls , " & vbCrLf & _
        '                  "             PerformanceCls , " & vbCrLf & _
        '                  "             Diff = CASE WHEN CONVERT(NUMERIC(18, 0), RTRIM(SUM(RemainingReceivingQty))) > 0 " & vbCrLf & _
        '                  "                         THEN '*' " & vbCrLf & _
        '                  "                         ELSE '' " & vbCrLf & _
        '                  "                    END "

        'ls_SQL = ls_SQL + "   FROM      ( SELECT DISTINCT " & vbCrLf & _
        '                  "                         ReceiveDate = CONVERT(CHAR(20), RAM.ReceiveDate, 113) , " & vbCrLf & _
        '                  "                         MS.SupplierName , " & vbCrLf & _
        '                  "                         SupplierSJNo = DSM.SuratJalanNo , " & vbCrLf & _
        '                  "                         SupplierDeliveryDate = CONVERT(CHAR, DSM.DeliveryDate, 106) , " & vbCrLf & _
        '                  "                         SupplierPlanDeliveryDate = CONVERT(CHAR, KM.KanbanDate, 106) , " & vbCrLf & _
        '                  "                         RAM.JenisArmada , " & vbCrLf & _
        '                  "                         RAM.NoPol , " & vbCrLf & _
        '                  "                         DeliveryTo = MDP.DeliveryLocationName , " & vbCrLf & _
        '                  "                         PASISJNo = ISNULL(DPM.SuratJalanNo, '-') , " & vbCrLf & _
        '                  "                         PASIDeliveryDate = CONVERT(CHAR, DSM.DeliveryDate, 106) , "

        'ls_SQL = ls_SQL + "                         DriverName = RTRIM(RAM.DriverName) + ' (' " & vbCrLf & _
        '                  "                         + RTRIM(RAM.DriverContact) + ')' , " & vbCrLf & _
        '                  "                         TotalBox = CONVERT(NUMERIC(18, 0), RAM.TotalBox) , " & vbCrLf & _
        '                  "                         PONo = POD.PONo , " & vbCrLf & _
        '                  "                         KanbanCls = POD.KanbanCls , " & vbCrLf & _
        '                  "                         KanbanNo = KD.KanbanNo , " & vbCrLf & _
        '                  "                         PartNo = POD.PartNo , " & vbCrLf & _
        '                  "                         PartName = MP.PartName , " & vbCrLf & _
        '                  "                         UOM = MU.Description , " & vbCrLf & _
        '                  "                         QtyBox = CONVERT(NUMERIC(18, 0), ISNULL(MPM.QtyBox, 0)) , " & vbCrLf & _
        '                  "                         SupplierDeliveryQty = CONVERT(NUMERIC(18, 0), ISNULL(DSD.DOQty, "

        'ls_SQL = ls_SQL + "                                                               0)) , " & vbCrLf & _
        '                  "                         PASIGoodReceivingQty = CONVERT(NUMERIC(18, 0), ISNULL(RPD.GoodRecQty, " & vbCrLf & _
        '                  "                                                               0)) , " & vbCrLf & _
        '                  "                         PASIDefectQty = CONVERT(NUMERIC(18, 0), ISNULL(RPD.DefectRecQty, " & vbCrLf & _
        '                  "                                                               0)) , " & vbCrLf & _
        '                  "                         PASIDeliveryQty = CONVERT(NUMERIC(18, 0), ISNULL(DPD.DOQty, " & vbCrLf & _
        '                  "                                                               0)) , " & vbCrLf & _
        '                  "                         GoodReceivingQty = CONVERT(NUMERIC(18, 0), CASE " & vbCrLf & _
        '                  "                                                               WHEN POM.DeliveryByPASICls = '1' " & vbCrLf & _
        '                  "                                                               THEN RTRIM(CONVERT(NUMERIC(18, " & vbCrLf & _
        '                  "                                                               0), COALESCE(RAD.RecQty, "

        'ls_SQL = ls_SQL + "                                                               DPD.DOQty, 0))) " & vbCrLf & _
        '                  "                                                               ELSE RTRIM(CONVERT(NUMERIC(18, " & vbCrLf & _
        '                  "                                                               0), COALESCE(RAD.RecQty, " & vbCrLf & _
        '                  "                                                               DSD.DOQty, 0))) " & vbCrLf & _
        '                  "                                                               END) , " & vbCrLf & _
        '                  "                         DefectReceivingQty = CONVERT(NUMERIC(18, 0), ISNULL(RAD.DefectQty, " & vbCrLf & _
        '                  "                                                               0)) , " & vbCrLf & _
        '                  "                         RemainingReceivingQty = CONVERT(NUMERIC(18, 0), CASE " & vbCrLf & _
        '                  "                                                               WHEN ISNULL(POM.DeliveryByPASICls, " & vbCrLf & _
        '                  "                                                               '0') = '1' " & vbCrLf & _
        '                  "                                                               THEN RTRIM(( ISNULL(DPD.DOQty, "

        'ls_SQL = ls_SQL + "                                                               0) " & vbCrLf & _
        '                  "                                                               - ( ISNULL(RAD.RecQty, " & vbCrLf & _
        '                  "                                                               0) " & vbCrLf & _
        '                  "                                                               + ISNULL(RAD.DefectQty, " & vbCrLf & _
        '                  "                                                               0) ) )) " & vbCrLf & _
        '                  "                                                               WHEN ISNULL(POM.DeliveryByPASICls, " & vbCrLf & _
        '                  "                                                               '0') = '0' " & vbCrLf & _
        '                  "                                                               THEN RTRIM(( ISNULL(DSD.DOQty, " & vbCrLf & _
        '                  "                                                               0) " & vbCrLf & _
        '                  "                                                               - ( ISNULL(RAD.RecQty, " & vbCrLf & _
        '                  "                                                               0) "

        'ls_SQL = ls_SQL + "                                                               + ISNULL(RAD.DefectQty, " & vbCrLf & _
        '                  "                                                               0) ) )) " & vbCrLf & _
        '                  "                                                               END) , " & vbCrLf & _
        '                  "                         ReceivingQtyBox = CEILING(( ( CASE WHEN POM.DeliveryByPASICls = '1' " & vbCrLf & _
        '                  "                                                            THEN COALESCE(RAD.RecQty, " & vbCrLf & _
        '                  "                                                               DPD.DOQty, 0) " & vbCrLf & _
        '                  "                                                            ELSE COALESCE(RAD.RecQty, " & vbCrLf & _
        '                  "                                                               DSD.DOQty, 0) " & vbCrLf & _
        '                  "                                                       END ) " & vbCrLf & _
        '                  "                                                     + ISNULL(RAD.DefectQty, 0) ) " & vbCrLf & _
        '                  "                                                   / MPM.QtyBox) , "

        'ls_SQL = ls_SQL + "                         UnitCls = KD.UnitCls , " & vbCrLf & _
        '                  "                         POM.DeliveryByPASICls , " & vbCrLf & _
        '                  "                         PerformanceCls = ISNULL(RTRIM(MPC.Description), '-') " & vbCrLf & _
        '                  "               FROM      PO_DETAIL POD " & vbCrLf & _
        '                  "                         LEFT JOIN PO_Master POM ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
        '                  "                                                    AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
        '                  "                                                    AND POM.PONO = POD.PONO " & vbCrLf & _
        '                  "                         LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
        '                  "                                                       AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
        '                  "                                                       AND KD.PONO = POD.PONO " & vbCrLf & _
        '                  "                                                       AND KD.PartNo = POD.PartNo "

        'ls_SQL = ls_SQL + "                         LEFT JOIN Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
        '                  "                                                       AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
        '                  "                                                       AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
        '                  "                                                       AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
        '                  "                         LEFT JOIN DOSupplier_Detail DSD ON KD.AffiliateID = DSD.AffiliateID " & vbCrLf & _
        '                  "                                                            AND KD.SupplierID = DSD.SupplierID " & vbCrLf & _
        '                  "                                                            AND KD.PONO = DSD.PONO " & vbCrLf & _
        '                  "                                                            AND KD.PartNo = DSD.PartNo " & vbCrLf & _
        '                  "                                                            AND KD.KanbanNo = DSD.KanbanNo " & vbCrLf & _
        '                  "                         LEFT JOIN DOSupplier_Master DSM ON DSM.AffiliateID = DSD.AffiliateID " & vbCrLf & _
        '                  "                                                            AND DSM.SupplierID = DSD.SupplierID "

        'ls_SQL = ls_SQL + "                                                            AND DSM.SuratJalanNo = DSD.SuratJalanNo " & vbCrLf & _
        '                  "                         LEFT JOIN ( SELECT  SuratJalanno , " & vbCrLf & _
        '                  "                                             SupplierID , " & vbCrLf & _
        '                  "                                             AffiliateID , " & vbCrLf & _
        '                  "                                             PONO , " & vbCrLf & _
        '                  "                                             KanbanNO , " & vbCrLf & _
        '                  "                                             Partno , " & vbCrLf & _
        '                  "                                             UnitCls , " & vbCrLf & _
        '                  "                                             DoQty = SUM(ISNULL(DoQty, 0)) " & vbCrLf & _
        '                  "                                     FROM    DOPasi_Detail " & vbCrLf & _
        '                  "                                     GROUP BY SuratJalanno , "

        'ls_SQL = ls_SQL + "                                             SupplierID , " & vbCrLf & _
        '                  "                                             AffiliateID , " & vbCrLf & _
        '                  "                                             PONO , " & vbCrLf & _
        '                  "                                             KanbanNO , " & vbCrLf & _
        '                  "                                             Partno , " & vbCrLf & _
        '                  "                                             UnitCls " & vbCrLf & _
        '                  "                                   ) DPD ON KD.AffiliateID = DPD.AffiliateID " & vbCrLf & _
        '                  "                                            AND KD.SupplierID = DPD.SupplierID " & vbCrLf & _
        '                  "                                            AND KD.PONO = DPD.PONO " & vbCrLf & _
        '                  "                                            AND KD.PartNo = DPD.PartNo " & vbCrLf & _
        '                  "                                            AND KD.KanbanNo = DPD.KanbanNo "

        'ls_SQL = ls_SQL + "                         LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID = DPD.AffiliateID " & vbCrLf & _
        '                  "                                                        AND DPM.SuratJalanNo = DPD.SuratJalanNo " & vbCrLf & _
        '                  "                         LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = KD.AffiliateID " & vbCrLf & _
        '                  "                                                             AND RPD.SupplierID = KD.SupplierID " & vbCrLf & _
        '                  "                                                             AND RPD.PONo = KD.PONo " & vbCrLf & _
        '                  "                                                             AND RPD.PartNo = KD.PartNo " & vbCrLf & _
        '                  "                                                             AND RPD.KanbanNo = KD.KanbanNo " & vbCrLf & _
        '                  "                                                             AND RPD.SuratJalanNo = DSD.SuratJalanNo " & vbCrLf & _
        '                  "                         LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = DPD.AffiliateID " & vbCrLf & _
        '                  "                                                               AND RAD.SupplierID = DPD.SupplierID " & vbCrLf & _
        '                  "                                                               AND RAD.KanbanNo = DPD.KanbanNo "

        'ls_SQL = ls_SQL + "                                                               AND RAD.PONo = DPD.PONo " & vbCrLf & _
        '                  "                                                               AND RAD.PartNo = DPD.PartNo " & vbCrLf & _
        '                  "                                                               AND RAD.SuratJalanNo = DPM.SuratJalanNo " & vbCrLf & _
        '                  "                         LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
        '                  "                                                               AND RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
        '                  "                         LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
        '                  "                         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.AffiliateID = POD.AffiliateID AND MPM.SupplierID = POD.SupplierID " & vbCrLf & _
        '                  "                         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
        '                  "                         LEFT JOIN MS_Supplier MS ON MS.SupplierID = RAM.SupplierID " & vbCrLf & _
        '                  "                         LEFT JOIN MS_DeliveryPlace MDP ON MDP.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
        '                  "                                                           AND MDP.AffiliateID = KM.AffiliateID " & vbCrLf & _
        '                  "                         LEFT JOIN MS_PerformanceCls MPC ON MPC.PerformanceCls = RAM.PerformanceCls " & vbCrLf & _
        '                  " 		   WHERE RAM.AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
        '                  "              AND RAM.SuratJalanNo = '" & Session("E02SupplierSJNo") & "'" & vbCrLf & _
        '                  "              --AND KD.KanbanNo = '" & Session("E02KanbanNo") & "'" & vbCrLf
        'ls_SQL = ls_SQL + "        ) RecEntry " & vbCrLf & _
        '                  " GROUP BY ReceiveDate,JenisArmada, NoPol, DeliveryTo, PASISJNo, PASIDeliveryDate, DriverName, TotalBox, PONo, KanbanCls, KanbanNo, PartNo, PartName, UOM, QtyBox, UnitCls, DeliveryByPASICls, PerformanceCls "

        ls_SQL = " SELECT  ColNo = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY PONo, KanbanNo )), * " & vbCrLf & _
                  " FROM " & vbCrLf & _
                  " ( " & vbCrLf & _
                  " 	SELECT DISTINCT " & vbCrLf & _
                  " 		RAM.ReceiveDate ,  " & vbCrLf & _
                  " 		RAM.JenisArmada ,  " & vbCrLf & _
                  " 		RAM.NoPol ,  " & vbCrLf & _
                  " 		DeliveryTo = MDP.DeliveryLocationName ,  " & vbCrLf & _
                  " 		PASISJNo = RAM.SuratJalanNo,  " & vbCrLf & _
                  " 		PASIDeliveryDate = DPM.DeliveryDate ,  " & vbCrLf & _
                  " 		RAM.DriverName ,  "

        ls_SQL = ls_SQL + " 		TotalBox = (SELECT SUM(ABC.RecQty / DEF.QtyBox) FROM ReceiveAffiliate_Detail ABC  " & vbCrLf & _
                          " 					LEFT JOIN MS_PartMapping DEF ON DEF.AffiliateID = ABC.AffiliateID and DEF.SupplierID = ABC.SupplierID and DEF.PartNo = ABC.PartNo " & vbCrLf & _
                          " 					WHERE ABC.SuratJalanNo = '" & Session("E02SupplierSJNo") & "' and ABC.AffiliateID = '" & Session("AffiliateID") & "'),  " & vbCrLf & _
                          " 		RAD.PONo,  " & vbCrLf & _
                          " 		POKanban = CASE WHEN ISNULL(RAD.POKanbanCls, '0') = '1' THEN 'YES' ELSE 'NO' END ,  " & vbCrLf & _
                          " 		RAD.KanbanNo ,  " & vbCrLf & _
                          " 		RAD.PartNo ,  " & vbCrLf & _
                          " 		MP.PartName ,  " & vbCrLf & _
                          " 		UOM = MU.Description,  " & vbCrLf & _
                          " 		MPM.QtyBox,  " & vbCrLf & _
                          " 		PerformanceCls = ISNULL(RTRIM(MPC.Description), '-'), "

        ls_SQL = ls_SQL + " 		PASIDeliveryQty = DPD.DOQty, " & vbCrLf & _
                          " 		AffiliateRecQty = RAD.RecQty, " & vbCrLf & _
                          " 		AffiliateDefQty = RAD.DefectQty, " & vbCrLf & _
                          " 		AffiliateRemQty = DPD.DOQty - (RAD.RecQty  + RAD.DefectQty), " & vbCrLf & _
                          " 		AffiliateRecBox = RAD.RecQty / MPM.QtyBox " & vbCrLf & _
                          " 	FROM ReceiveAffiliate_Master RAM " & vbCrLf & _
                          " 	LEFT JOIN ReceiveAffiliate_Detail RAD ON RAM.SuratJalanNo = RAD.SuratJalanNo and RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                          " 	LEFT JOIN DOPASI_Master DPM ON DPM.AffiliateID = RAM.AffiliateID and DPM.SuratJalanNo = RAM.SuratJalanNo " & vbCrLf & _
                          " 	LEFT JOIN DOPASI_Detail DPD ON DPD.AffiliateID = RAD.AffiliateID and DPD.SupplierID = RAD.SupplierID and DPD.PartNo = RAD.PartNo AND RAD.SuratJalanNo = DPD.SuratJalanNo " & vbCrLf & _
                          " 	LEFT JOIN MS_DeliveryPlace MDP ON MDP.AffiliateID = RAM.AffiliateID  " & vbCrLf & _
                          " 	LEFT JOIN MS_Parts MP ON MP.PartNo = RAD.PartNo "

        ls_SQL = ls_SQL + " 	LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                          " 	LEFT JOIN MS_PartMapping MPM ON MPM.AffiliateID = RAD.AffiliateID and MPM.SupplierID = RAD.SupplierID AND MPM.PartNo = RAD.PartNo " & vbCrLf & _
                          " 	LEFT JOIN MS_PerformanceCls MPC ON MPC.PerformanceCls = RAM.PerformanceCls  " & vbCrLf & _
                          " 	WHERE RAM.AffiliateID = '" & Session("AffiliateID") & "' and RAM.SuratJalanNo = '" & Session("E02SupplierSJNo") & "' " & vbCrLf & _
                          " )XYZ " & vbCrLf & _
                          "  "

        Report.DataSource = Nothing

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            Report.DataSource = ds.Tables(0)
            'DevExpress.XtraReports.Web.ReportViewer.WriteHtmlTo(System.Web.HttpContext.Current.Response, Report)
            Viewer.Report = Report
            Viewer.DataBind()
            'Viewer.ViewStateMode=
        End Using
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")

        urlBack = "~/Receiving/ReceivingEntry.aspx?prm=" & Session("E02ParamPageLoad").ToString
        Response.Redirect(urlBack)
    End Sub
#End Region
    
End Class