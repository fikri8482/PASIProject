Imports System.Data
Imports System.Data.SqlClient

Public Class GoodReceivingReportExport
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = "", urlBack As String = ""
#End Region

#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Report As New rptGoodReceivingReportExport

        Dim param As String = Request.QueryString("prm").ToString
        Dim pAffiliate As String = Split(param, "|")(0)
        Dim psuppID As String = Split(param, "|")(1)
        Dim pSuratjalan As String = Split(param, "|")(2)

        ls_SQL = " SELECT  " & vbCrLf & _
                  " No = CONVERT(numeric,ROW_NUMBER() OVER (ORDER BY PartNo)), Supplier,  " & vbCrLf & _
                  " SuratJalanNO, SupDeldate, Suppplandeldate, ReceiveDate, DeliveryTo, JenisArmada, Nopol,  " & vbCrLf & _
                  " DriverName,  TotalBox, OrderNo, PartNo, PartName, Uom, QtyBox, SupDelQty, GoodRecQty, DefectRecQty,  " & vbCrLf & _
                  " RemainingRecQty,RecQtyBox " & vbCrLf & _
                  "  FROM ( " & vbCrLf & _
                  " SELECT distinct " & vbCrLf & _
                  " --No = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY DSD.PartNo)), " & vbCrLf & _
                  " Supplier = MS.SupplierName, " & vbCrLf & _
                  " SuratJalanNO = DSM.SuratjalanNo, " & vbCrLf & _
                  " SupDeldate = Convert(Char(12), convert(Datetime, isnull(DSM.DeliveryDate,'')),106), " & vbCrLf & _
                  " Suppplandeldate = Convert(Char(12), convert(Datetime, COALESCE(PRM.ETDVendor, POM.ETDVendor)),106), " & vbCrLf & _
                  " ReceiveDate = Convert(Char(12), convert(Datetime, isnull(RM.ReceiveDate,'')),106), " & vbCrLf & _
                  " DeliveryTo = isnull(POM.ForwarderID,''), " & vbCrLf & _
                  " JenisArmada = DSM.JenisArmada, " & vbCrLf & _
                  " Nopol = DSM.Nopol, " & vbCrLf & _
                  " DriverName = DSM.DriverName, "

        ls_SQL = ls_SQL + " TotalBox = DSM.TotalBox, " & vbCrLf & _
                          " OrderNo = DSM.OrderNo, " & vbCrLf & _
                          " PartNo = DSD.PartNo, " & vbCrLf & _
                          " PartName = MP.PartName, " & vbCrLf & _
                          " Uom = MU.Description, " & vbCrLf & _
                          " QtyBox = ISNULL(DSD.POQtyBox,MPM.QtyBox), " & vbCrLf & _
                          " SupDelQty = DSD.DOQty, " & vbCrLf & _
                          " GoodRecQty = RD.GoodRecQty, " & vbCrLf & _
                          " DefectRecQty = RD.DefectRecQty, " & vbCrLf & _
                          " RemainingRecQty = DSD.DOQty - RD.GoodRecQty, " & vbCrLf & _
                          " RecQtyBox = CEILING(RD.GoodRecQty / ISNULL(DSD.POQtyBox,MPM.QtyBox)) "

        ls_SQL = ls_SQL + " FROM   DOSupplier_Detail_Export DSD  " & vbCrLf & _
                          "          LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo  " & vbCrLf & _
                          "                                                    AND DSM.AffiliateID = DSD.AffiliateID  " & vbCrLf & _
                          "                                                    AND DSM.SupplierID = DSD.SupplierID  " & vbCrLf & _
                          "                                                    AND DSM.PONO = DSD.PONO  " & vbCrLf & _
                          "          LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO  " & vbCrLf & _
                          "                                            AND POD.AffiliateID = DSM.AffiliateID  " & vbCrLf & _
                          "                                            AND POD.SupplierID = DSM.SupplierID  " & vbCrLf & _
                          "                                            AND POD.PartNo = DSD.PartNo  " & vbCrLf & _
                          "          LEFT JOIN ( SELECT  * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo1 ,  "

        ls_SQL = ls_SQL + "                              ETDVendor = ETDVendor1 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort1 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory1  " & vbCrLf & _
                          "                      FROM    Po_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO1,'') <> '' " & vbCrLf & _
                          "                      UNION ALL  " & vbCrLf & _
                          "                      SELECT  * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo2 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor2 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort2 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory2  " & vbCrLf & _
                          "                      FROM    Po_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO2,'') <> '' " & vbCrLf

        ls_SQL = ls_SQL + "                      UNION ALL  " & vbCrLf & _
                          "                      SELECT  * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo3 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor3 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort3 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory3  " & vbCrLf & _
                          "                      FROM    Po_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO3,'') <> '' " & vbCrLf & _
                          "                      UNION ALL  " & vbCrLf & _
                          "                      SELECT  * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo4 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor4 ,  "

        ls_SQL = ls_SQL + "                              ETAPort = ETAPort4 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory4  " & vbCrLf & _
                          "                      FROM    Po_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO4,'') <> '' " & vbCrLf & _
                          "                    ) POM ON POM.PONO = POD.PONO  " & vbCrLf & _
                          "                             AND POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "                             AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "          LEFT JOIN ( SELECT TOP 1  " & vbCrLf & _
                          "                              * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo1 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor1 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort1 ,  "

        ls_SQL = ls_SQL + "                              ETAFactory = ETAFactory1  " & vbCrLf & _
                          "                      FROM    PoRev_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO1,'') <> '' " & vbCrLf & _
                          "                      ORDER BY PORevNo  " & vbCrLf & _
                          "                      UNION ALL  " & vbCrLf & _
                          "                      SELECT TOP 1  " & vbCrLf & _
                          "                              * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo2 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor2 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort2 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory2  " & vbCrLf & _
                          "                      FROM    PoRev_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO2,'') <> '' " & vbCrLf

        ls_SQL = ls_SQL + "                      ORDER BY PORevNo  " & vbCrLf & _
                          "                      UNION ALL  " & vbCrLf & _
                          "                      SELECT TOP 1  " & vbCrLf & _
                          "                              * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo3 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor3 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort3 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory3  " & vbCrLf & _
                          "                      FROM    PoRev_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO3,'') <> '' " & vbCrLf & _
                          "                      ORDER BY PORevNo  " & vbCrLf & _
                          "                      UNION ALL  "

        ls_SQL = ls_SQL + "                      SELECT TOP 1  " & vbCrLf & _
                          "                              * ,  " & vbCrLf & _
                          "                              OrderNO = OrderNo4 ,  " & vbCrLf & _
                          "                              ETDVendor = ETDVendor4 ,  " & vbCrLf & _
                          "                              ETAPort = ETAPort4 ,  " & vbCrLf & _
                          "                              ETAFactory = ETAFactory4  " & vbCrLf & _
                          "                      FROM    PoRev_Master_Export  " & vbCrLf & _
                          "                      WHERE isnull(OrderNO4,'') <> '' " & vbCrLf & _
                          "                      ORDER BY PORevNo  " & vbCrLf & _
                          "                    ) PRM ON PRM.PONO = POD.PONO  " & vbCrLf & _
                          "                             AND PRM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "                             AND PRM.SupplierID = POD.SupplierID  "

        ls_SQL = ls_SQL + "          LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO  " & vbCrLf & _
                          "                                               AND PRD.AffiliateID = PRM.AffiliateID  " & vbCrLf & _
                          "                                               AND PRD.SupplierID = PRM.SupplierID  " & vbCrLf & _
                          "                                               AND PRD.PartNo = DSD.PartNo  " & vbCrLf & _
                          "          LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo  " & vbCrLf & _
                          "                                                  AND DSD.affiliateID = RM.affiliateID  " & vbCrLf & _
                          "                                                  AND DSD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "          LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo  " & vbCrLf & _
                          "                                                  AND RM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                          "                                                  AND RM.SupplierID = RD.SupplierID  " & vbCrLf & _
                          "                                                  AND RM.PONO = RD.PONO  "

        ls_SQL = ls_SQL + "                                                  AND DSD.PartNo = RD.PartNo  " & vbCrLf & _
                          "                                                  AND DSD.PONO = RD.PONO  " & vbCrLf & _
                          "          LEFT JOIN ( SELECT  suratjalanno ,  " & vbCrLf & _
                          "                              supplierid ,  " & vbCrLf & _
                          "                              affiliateID ,  " & vbCrLf & _
                          "                              PONO ,  " & vbCrLf & _
                          "                              partno ,  " & vbCrLf & _
                          "                              goodRecQty = SUM(ISNULL(goodRecQty, 0)) ,  " & vbCrLf & _
                          "                              DefectRecQty = SUM(ISNULL(DefectRecQty, 0))  " & vbCrLf & _
                          "                      FROM    ReceiveForwarder_Detail  " & vbCrLf & _
                          "                      GROUP BY suratjalanno ,  "

        ls_SQL = ls_SQL + "                              supplierid ,  " & vbCrLf & _
                          "                              affiliateID ,  " & vbCrLf & _
                          "                              PONO ,  " & vbCrLf & _
                          "                              partno  " & vbCrLf & _
                          "                    ) REM ON REM.SuratJalanNo = RD.SuratjalanNo  " & vbCrLf & _
                          "                             AND REM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                          "                             AND REM.SupplierID = RD.SupplierID  " & vbCrLf & _
                          "                             AND REM.PONO = RD.PONO  " & vbCrLf & _
                          "                             AND REM.PartNo = RD.PartNo  " & vbCrLf & _
                          "                             AND REM.PONO = RD.PONO  " & vbCrLf & _
                          "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID  "

        ls_SQL = ls_SQL + "          LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID  " & vbCrLf & _
                          "          LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID  " & vbCrLf & _
                          "          LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo  " & vbCrLf & _
                          "          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DSD.PartNo and MPM.AffiliateID = DSD.AffiliateID and MPM.SupplierID = DSD.SupplierID " & vbCrLf & _
                          "          LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
                          "  WHERE DSD.AffiliateID = '" & pAffiliate & "'  " & vbCrLf & _
                          "  AND DSD.SUpplierID = '" & psuppID & "'  " & vbCrLf & _
                          "  AND RM.SuratJalanNo = '" & pSuratjalan & "'  "

        ls_SQL = ls_SQL + "  )x ORDER BY NO asc"

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

        urlBack = "~/DeliveryExport/DeliveryToAffListExport.aspx"
        Response.Redirect(urlBack)
    End Sub
#End Region

End Class