Imports System.Data.SqlClient
Imports System.Transactions
Imports System.Drawing
Imports System.IO

Public Class InvToAffExport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    ''parameter

    Dim pOrderNo As String
    Dim pAffiliateCode As String
    Dim pSupplierID As String
    Dim pSJ As String
    Dim pInvoiceNo As String
    Dim pShipping As String
    Dim pNotes As String
    Dim pAffName As String

#End Region

    Private Sub up_GridLoad(ByVal pAff As String, ByVal pInvNo As String, ByVal pShippngNo As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If Replace(Replace(pInvNo, "'", ""), ",", "") = "" Then
            pInvNo = ""
        End If

        If Replace(Replace(pShippngNo, "'", ""), ",", "") = "" Then
            pShippngNo = ""
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Session.Remove("SQL")
            ls_SQL = " SELECT distinct " & vbCrLf & _
                  " colno = CONVERT(char,ROW_NUMBER() OVER(ORDER BY SD.Partno)), " & vbCrLf & _
                  " colorderno = SD.OrderNo, " & vbCrLf & _
                  " colsj = ISNULL(DSM.SuratJalanNo, ''), " & vbCrLf & _
                  " colpartno = ISNULL(SD.Partno,''), " & vbCrLf & _
                  " colpartname = ISNULL(MP.PartName,''), " & vbCrLf & _
                  " coluom = MU.Description, " & vbCrLf & _
                  " colqtybox = MPM.QtyBox, " & vbCrLf & _
                  " coldeliveryqty = DSD.DOQty, " & vbCrLf & _
                  " colinvqty = COALESCE(ID.Qty, DSD.DOQty), " & vbCrLf & _
                  " coldelqtybox = CEILING(DSD.DOQty/MPM.QtyBox), "

            ls_SQL = ls_SQL + " colInvCurr = ISNULL(IM.CurrCls,'" & txtcurr.Text & "'), " & vbCrLf & _
                              " colInvPrice = isnull(ISNULL(ID.Price,MPR.Price),0), " & vbCrLf & _
                              " colInvAmount = isnull(ISNULL(ID.Amount,(MPR.Price * COALESCE(ID.Qty, DSD.DOQty))),0), shippingno = Isnull(SM.ShippingInstructionNO, '') " & vbCrLf & _
                              " FROM    ShippingInstruction_Master SM  " & vbCrLf & _
                              "          LEFT JOIN ShippingInstruction_Detail SD ON SM.affiliateID = SD.AffiliateID  " & vbCrLf & _
                              "                                                     AND SM.forwarderID = SD.ForwarderID  " & vbCrLf & _
                              "                                                     AND SM.ShippingInstructionNo = SD.ShippingInstructionNo           " & vbCrLf & _
                              " 		 LEFT JOIN ReceiveForwarder_Master RM ON RM.AffiliateID = SM.AffiliateID  " & vbCrLf & _
                              "                                                  AND RM.ForwarderID = SM.ForwarderID  " & vbCrLf & _
                              "                                                  AND RM.OrderNo = SD.OrderNo  " & vbCrLf & _
                              "                                                  AND RM.SuratjalanNo = SD.suratJalanNo  "

            ls_SQL = ls_SQL + "          LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNo = RM.SuratJalanNo  " & vbCrLf & _
                              "                                                  AND RD.SupplierID = RM.SupplierID  " & vbCrLf & _
                              "                                                  AND RD.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                              "                                                  AND RD.OrderNo = SD.OrderNo  " & vbCrLf & _
                              "                                                  AND RD.PartNo = SD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN DOSupplier_Master_Export DSM ON DSM.affiliateID = RM.AffiliateID  " & vbCrLf & _
                              "                                                    AND DSM.orderNo = RM.OrderNo           " & vbCrLf & _
                              " 		 LEFT JOIN DOSupplier_Detail_export DSD ON DSD.SuratJalanNo = DSM.SuratJalanno " & vbCrLf & _
                              " 													AND DSD.AffiliateID =DSM.AffiliateID " & vbCrLf & _
                              " 													AND DSD.supplierID = DSM.SupplierID " & vbCrLf & _
                              " 													AND DSD.OrderNo = DSM.OrderNo "

            ls_SQL = ls_SQL + " 													AND DSD.PartNo = SD.PartNo " & vbCrLf & _
                              " 		 LEFT JOIN ( SELECT  * ,  " & vbCrLf & _
                              "                              OrderNO = OrderNo1 ,  " & vbCrLf & _
                              "                              ETDVendor = ETDVendor1 ,  " & vbCrLf & _
                              "                              ETAPort = ETAPort1 ,  " & vbCrLf & _
                              "                              ETAFactory = ETAFactory1 ,  " & vbCrLf & _
                              "                              week = 1  " & vbCrLf & _
                              "                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                      UNION ALL  " & vbCrLf & _
                              "                      SELECT  * ,  " & vbCrLf & _
                              "                              OrderNO = OrderNo2 ,  "

            ls_SQL = ls_SQL + "                              ETDVendor = ETDVendor2 ,                              ETAPort = ETAPort2 ,  " & vbCrLf & _
                              "                              ETAFactory = ETAFactory2 ,  " & vbCrLf & _
                              "                              week = 2  " & vbCrLf & _
                              "                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                      UNION ALL  " & vbCrLf & _
                              "                      SELECT  * ,  " & vbCrLf & _
                              "                              OrderNO = OrderNo3 ,  " & vbCrLf & _
                              "                              ETDVendor = ETDVendor3 ,  " & vbCrLf & _
                              "                              ETAPort = ETAPort3 ,  " & vbCrLf & _
                              "                              ETAFactory = ETAFactory3 ,  " & vbCrLf & _
                              "                              week = 3                      FROM    Po_Master_Export  "

            ls_SQL = ls_SQL + "                      UNION ALL  " & vbCrLf & _
                              "                      SELECT  * ,  " & vbCrLf & _
                              "                              OrderNO = OrderNo4 ,  " & vbCrLf & _
                              "                              ETDVendor = ETDVendor4 ,  " & vbCrLf & _
                              "                              ETAPort = ETAPort4 ,  " & vbCrLf & _
                              "                              ETAFactory = ETAFactory4 ,  " & vbCrLf & _
                              "                              week = 4  " & vbCrLf & _
                              "                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                      UNION ALL  " & vbCrLf & _
                              "                      SELECT  * ,                              OrderNO = OrderNo5 ,  " & vbCrLf & _
                              "                              ETDVendor = ETDVendor5 ,  "

            ls_SQL = ls_SQL + "                              ETAPort = ETAPort5 ,  " & vbCrLf & _
                              "                              ETAFactory = ETAFactory5 ,  " & vbCrLf & _
                              "                              week = 5  " & vbCrLf & _
                              "                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                    ) POM ON POM.AffiliateID = DSM.AffiliateID  " & vbCrLf & _
                              "                             AND POM.SupplierID = DSM.SupplierID  " & vbCrLf & _
                              "                             AND POM.orderno = DSM.OrderNo  " & vbCrLf & _
                              "  		LEFT JOIN InvoiceOverseas_Master IM ON IM.AffiliateID = SM.AffiliateID  " & vbCrLf & _
                              "  												AND IM.ShippingInstructionNo = SM.ShippingInstructionNo  		 " & vbCrLf & _
                              " 		LEFT JOIN InvoiceOverseas_Detail ID ON ID.InvoiceNo = IM.InvoiceNo  " & vbCrLf & _
                              "  												AND ID.AffiliateID = IM.AffiliateID  "

            ls_SQL = ls_SQL + "  												AND ID.ShippingInstructionNo = IM.ShippingInstructionNo  " & vbCrLf & _
                              "  												AND ID.OrderNo = SD.OrderNo  " & vbCrLf & _
                              "  												AND ID.Partno = SD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = SM.AffiliateID  " & vbCrLf & _
                              "          LEFT JOIN MS_Supplier MS ON MS.SupplierID = RM.SupplierID " & vbCrLf & _
                              "          LEFT JOIN MS_Parts MP ON MP.PartNo = SD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = RD.PartNo and MPM.AffiliateID = RD.AffiliateID and MPM.SupplierID = RD.SupplierID " & vbCrLf & _
                              "          LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
                              "          LEFT JOIN MS_Price MPR ON MPR.AffiliateID = SM.AffiliateID " & vbCrLf & _
                              "                                     AND MPR.PartNo = SD.PartNo " & vbCrLf & _
                              "                                     AND MPR.currcls = '" & cbocurr.Text & "' " & vbCrLf & _
                              "                                     AND '" & txtInvoiceDate.Text & "' between Convert(Char(12), convert(Datetime, MPR.startdate),106) and Convert(Char(12), convert(Datetime, MPR.enddate),106)" & vbCrLf & _
                              " WHERE SM.AffiliateID = '" & Trim(pAff) & "'" & vbCrLf

            If pInvNo = "" Then
                ls_SQL = ls_SQL + "   AND SM.ShippingInstructionNo IN ('" & pShippngNo & "') " & vbCrLf
            Else
                ls_SQL = ls_SQL + "   AND IM.InvoiceNo ='" & pInvNo & "' " & vbCrLf
            End If

            Session("SQL") = ls_SQL
            ''ls_SQL = ls_SQL + "   ORDER BY  SM.ShippingInstructionNo  " & vbCrLf




            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'AFFILIATE
        ls_sql = "Select CurrCls = RTRIM(CurrCls) ,Description = RTRIM(Description) FROM dbo.MS_CurrCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbocurr
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("CurrCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 240

                .TextField = "Currency"
                .DataBind()
            End With
            sqlConn.Close()
        End Using
    End Sub

    'Private Function uf_SumAmount(ByVal pPO As String, ByVal pKanban As String)
    '    Dim ls_SQL As String = ""

    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        ls_SQL = " SELECT colInvAmount = ROUND(CONVERT(Char,ISNULL(SUM(colInvAmount),0),0),0) " & vbCrLf & _
    '              " FROM ( " & vbCrLf & _
    '              "    SELECT    colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo, POD.PartNo )) ,  " & vbCrLf & _
    '              "              colpono = POM.PONo ,  " & vbCrLf & _
    '              "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'  " & vbCrLf & _
    '              "                                 ELSE 'YES'  " & vbCrLf & _
    '              "                            END ,  " & vbCrLf & _
    '              "              colkanbanno = CASE WHEN POD.KanbanCls = '0' THEN '-'  " & vbCrLf & _
    '              "                                 ELSE ISNULL(KD.KanbanNo, '')  " & vbCrLf & _
    '              "                            END ,  " & vbCrLf & _
    '              "              colpartno = POD.PartNo ,  "

    '        ls_SQL = ls_SQL + "              colpartname = MP.PartName ,  " & vbCrLf & _
    '                          "              coluom = UC.Description ,  " & vbCrLf & _
    '                          "              colCls = UC.unitcls ,  " & vbCrLf & _
    '                          "              colQtyBox = ISNULL(MP.QtyBox, 0) ,  " & vbCrLf & _
    '                          "              colpasideliveryqty = COALESCE(PDD.DOQty,  " & vbCrLf & _
    '                          "                                            ISNULL(SDD.DOQty, 0)  " & vbCrLf & _
    '                          "                                            - ( ISNULL(PRD.GoodRecQty, 0)  " & vbCrLf & _
    '                          "                                                + ISNULL(PRD.DefectRecQty, 0) ),  " & vbCrLf & _
    '                          "                                            0) ,  " & vbCrLf & _
    '                          "              colAffRecQty = ISNULL(RAD.RecQty,0) ,  " & vbCrLf & _
    '                          "              colInvoiceToAffQty = COALESCE(IPD.InvQty,RAD.RecQty,0) ,  "

    '        ls_SQL = ls_SQL + "              coldelqtybox = CASE MP.QtyBox  " & vbCrLf & _
    '                          "                               WHEN 0 THEN 0  " & vbCrLf & _
    '                          "                               ELSE ISNULL(SDD.DOQty, 0) / MP.QtyBox  " & vbCrLf & _
    '                          "                             END ,  " & vbCrLf & _
    '                          "              colInvCurr = ISNULL(MPR.CurrCls,'') ,  " & vbCrLf & _
    '                          "              colInvPrice = ISNULL(MPR.Price,0) ,  " & vbCrLf & _
    '                          "              colInvAmount = COALESCE(IPD.InvAmount,(RAD.RecQty*MPR.Price),0 ) " & vbCrLf & _
    '                          "    FROM      dbo.PO_Master POM  " & vbCrLf & _
    '                          "              LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
    '                          "                                         AND POM.PoNo = POD.PONo  " & vbCrLf & _
    '                          "                                         AND POM.SupplierID = POD.SupplierID  "

    '        ls_SQL = ls_SQL + "              LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
    '                          "                                                AND KD.PoNo = POD.PONo  " & vbCrLf & _
    '                          "                                                AND KD.SupplierID = POD.SupplierID  " & vbCrLf & _
    '                          "                                                AND KD.PartNo = POD.PartNo  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID  " & vbCrLf & _
    '                          "                                                AND KD.KanbanNo = KM.KanbanNo  " & vbCrLf & _
    '                          "                                                AND KD.SupplierID = KM.SupplierID  " & vbCrLf & _
    '                          "                                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
    '                          "                                                     AND KD.KanbanNo = SDD.KanbanNo  " & vbCrLf & _
    '                          "                                                     AND KD.SupplierID = SDD.SupplierID  "

    '        ls_SQL = ls_SQL + "                                                     AND KD.PartNo = SDD.PartNo  " & vbCrLf & _
    '                          "                                                     AND KD.PONo = SDD.PONo " & vbCrLf & _
    '                          "              LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
    '                          "                                                     AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
    '                          "                                                     AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
    '                          "                                                      AND KD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
    '                          "                                                      AND KD.SupplierID = PRD.SupplierID  " & vbCrLf & _
    '                          "                                                      AND KD.PartNo = PRD.PartNo  " & vbCrLf & _
    '                          "                                                      AND KD.PONo = PRD.PartNo  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
    '                          "                                                      AND PRM.SuratJalanNo = PRD.SuratJalanNo  "

    '        ls_SQL = ls_SQL + "                                                      AND PRM.SupplierID = PRD.SupplierID  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
    '                          "                                                 AND KD.KanbanNo = PDD.KanbanNo  " & vbCrLf & _
    '                          "                                                 AND KD.SupplierID = PDD.SupplierID  " & vbCrLf & _
    '                          "                                                 AND KD.PartNo = PDD.PartNo  " & vbCrLf & _
    '                          "                                                 AND KD.PoNo = PDD.PoNo  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
    '                          "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo  " & vbCrLf & _
    '                          "                                                 AND PDD.SupplierID = PDM.SupplierID  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
    '                          "                                                           AND PDD.KanbanNo = RAD.KanbanNo  "

    '        ls_SQL = ls_SQL + "                                                           AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
    '                          "                                                           AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
    '                          "                                                           AND PDD.PoNo = RAD.PoNo  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
    '                          "                                                           AND RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
    '                          "                                                           AND RAM.SupplierID = RAD.SupplierID  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
    '                          "                                                      AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
    '                          "                                                      AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
    '                          "                                                      AND RAD.PONo = IPD.PONo  " & vbCrLf & _
    '                          "                                                      AND RAD.SuratJalanNo = IPD.SuratJalanNo  "

    '        ls_SQL = ls_SQL + "              LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
    '                          "                                                      AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
    '                          "                                                      AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID  " & vbCrLf & _
    '                          "              LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
    '                          "               " & vbCrLf & _
    '                          "              LEFT JOIN dbo.MS_Price MPR ON MPR.AffiliateID = RAD.AffiliateID " & vbCrLf & _
    '                          " 											AND MPR.PartNo = RAD.PartNo  "

    '        ls_SQL = ls_SQL + " 											AND RAM.ReceiveDate BETWEEN MPR.StartDate AND MPR.EndDate " & vbCrLf & _
    '                          " 			 LEFT JOIN dbo.MS_CurrCls MC ON MC.CurrCls = MPR.CurrCls " & vbCrLf & _
    '                          "    WHERE     POM.PONo IN (" & pPO & ")  " & vbCrLf & _
    '                          "              AND KD.KanbanNo IN (" & pKanban & ")  " & vbCrLf & _
    '                          " ) Inv " & vbCrLf & _
    '                          "  "
    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)
    '        uf_SumAmount = ds.Tables(0).Rows(0)("colInvAmount")
    '        sqlConn.Close()

    '    End Using
    'End Function

    Private Function Save_Master(ByVal pInvNo As String, ByVal pAffiliateID As String, ByVal pShipping As String, ByVal pInvDate As Date, ByVal pDueDate As Date, _
                            ByVal pPaymentTerm As String, ByVal pTotalAmount As String, ByVal pNotes As String, ByVal pCurrCls As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM InvoiceOverseas_Master WHERE InvoiceNo = '" & pInvNo & "' AND AffiliateID = '" & pAffiliateID & "' AND ShippingInstructionNo = '" & pShipping & "') " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " UPDATE dbo.InvoiceOverseas_Master " & vbCrLf & _
                          " SET InvoiceDate ='" & pInvDate & "', " & vbCrLf & _
                          " 	DueDate ='" & pDueDate & "', " & vbCrLf & _
                          " 	PaymentTerm ='" & pPaymentTerm & "', " & vbCrLf & _
                          " 	TotalAmount ='" & pTotalAmount & "', " & vbCrLf & _
                          "     Remarks ='" & pNotes & "', " & vbCrLf & _
                          "     Currcls='" & pCurrCls & "' " & vbCrLf & _
                          " WHERE InvoiceNo = '" & pInvNo & "'  " & vbCrLf & _
                          "   AND AffiliateID = '" & pAffiliateID & "'  " & vbCrLf & _
                          "   AND ShippingInstructionNo = '" & pShipping & "' " & vbCrLf & _
                          " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " INSERT INTO dbo.InvoiceOverseas_Master " & vbCrLf & _
                          "         ( InvoiceNo , InvoiceDate, AffiliateID ,ShippingInstructionNo ,DueDate , " & vbCrLf & _
                          "           PaymentTerm ,CurrCls, TotalAmount ,remarks " & vbCrLf & _
                          "         ) " & vbCrLf

        ls_sql = ls_sql + " VALUES  ( '" & pInvNo & "' , " & vbCrLf & _
                          "           '" & pInvDate & "', " & vbCrLf & _
                          "           '" & pAffiliateID & "' , " & vbCrLf & _
                          "           '" & pShipping & "' , " & vbCrLf & _
                          "           '" & pDueDate & "' , " & vbCrLf & _
                          "           '" & pPaymentTerm & "' , " & vbCrLf & _
                          "           '" & pCurrCls & "' , " & vbCrLf & _
                          "           '" & pTotalAmount & "' , " & vbCrLf & _
                          "           '" & pNotes & "' " & vbCrLf & _
                          "         )  " & vbCrLf & _
                          " END " & vbCrLf

        Save_Master = ls_sql
    End Function

    Private Function Save_Detail(ByVal pInvNo As String, ByVal pshipping As String, ByVal pAffiliateID As String, ByVal pPOno As String, ByVal pPartNo As String, ByVal pReceiveQty As String, ByVal pInvQty As String, ByVal pInvPrice As String, ByVal pInvAmount As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM dbo.InvoiceOverseas_Detail WHERE InvoiceNo = '" & pInvNo & "' AND AffiliateID = '" & pAffiliateID & "' AND ShippingInstructionNo = '" & pshipping & "' AND OrderNo = '" & pPOno & "' AND PartNo = '" & pPartNo & "') " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " UPDATE dbo.InvoiceOverseas_Detail " & vbCrLf & _
                          " SET Qty ='" & pInvQty & "'," & vbCrLf & _
                          " 	Price ='" & pInvPrice & "', " & vbCrLf & _
                          " 	Amount ='" & pInvAmount & "'   " & vbCrLf & _
                          " WHERE InvoiceNo = '" & pInvNo & "'  " & vbCrLf & _
                          "   AND ShippingInstructionNo = '" & pshipping & "'  " & vbCrLf & _
                          "   AND AffiliateID = '" & pAffiliateID & "'  " & vbCrLf & _
                          "   AND OrderNo = '" & pPOno & "' " & vbCrLf & _
                          "   AND PartNo = '" & pPartNo & "' " & vbCrLf

        ls_sql = ls_sql + " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " INSERT INTO dbo.InvoiceOverseas_Detail " & vbCrLf & _
                          "         ( InvoiceNo,ShippingInstructionNo,AffiliateID,OrderNo, " & vbCrLf & _
                          "           PartNo,Qty,Price,Amount" & vbCrLf & _
                          "         ) " & vbCrLf & _
                          " VALUES  ( '" & pInvNo & "' , " & vbCrLf & _
                          "           '" & pshipping & "' ,  " & vbCrLf & _
                          "           '" & pAffiliateID & "' ,  " & vbCrLf

        ls_sql = ls_sql + "           '" & pPOno & "' ,  " & vbCrLf & _
                          "           '" & pPartNo & "' , " & vbCrLf & _
                          "           '" & pInvQty & "' , " & vbCrLf & _
                          "           '" & pInvPrice & "' ,  " & vbCrLf & _
                          "           " & pInvAmount & "  " & vbCrLf & _
                          "         ) " & vbCrLf & _
                          " END " & vbCrLf

        Save_Detail = ls_sql
    End Function


    Private Sub up_SaveAll()
        Dim ls_SQL As String = ""
        Dim ls_User As String = Trim(Session("UserID").ToString)

        Dim ls_InvNo As String = Trim(txtPasiInvoiceNo.Text)
        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)
        Dim ls_SJno As String = Session("shippingno")
        Dim ls_InvDate As Date = Trim(txtInvoiceDate.Text)
        ls_InvDate = Format(CDate(ls_InvDate), "yyyy-MM-dd")
        Dim ls_DueDate As Date = Trim(dtDueDate.Text)
        Dim ls_PaymentTerm As String = Trim(txtPaymentTerm.Text)
        Dim ls_TotalAmount As String = Trim(txttotalamount.Text)
        Dim ls_Notes As String = Trim(MmNotes.Text)

        Dim ls_curr As String = cbocurr.Text

        Dim ls_POno As String
        Dim ls_PartNo As String
        Dim ls_ReceiveQty As Double
        Dim ls_InvQty As Double
        Dim ls_InvPrice As Double
        Dim ls_InvAmount As Double
        Dim TotalAmount As Double = 0

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = Session("SQL")

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Using sqlScope As New TransactionScope
                    'Save Master
                    ls_SQL = Save_Master(ls_InvNo, ls_AffiliateID, ls_SJno, ls_InvDate, ls_DueDate, ls_PaymentTerm, ls_TotalAmount, ls_Notes, ls_curr)
                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    Try
                        With ds.Tables(0)
                            For iLoop = 0 To .Rows.Count - 1

                                ls_POno = Trim(.Rows(iLoop).Item("colorderno"))
                                ls_PartNo = Trim(.Rows(iLoop).Item("colpartno"))
                                ls_InvQty = Trim(.Rows(iLoop).Item("colinvqty"))
                                ls_InvPrice = Trim(.Rows(iLoop).Item("colInvPrice"))
                                ls_InvAmount = Trim(.Rows(iLoop).Item("colInvAmount"))
                                ls_SJno = Trim(.Rows(iLoop).Item("shippingno"))
                                TotalAmount = TotalAmount + ls_InvAmount

                                'Save Detail
                                ls_SQL = Save_Detail(ls_InvNo, ls_SJno, ls_AffiliateID, ls_POno, ls_PartNo, ls_ReceiveQty, ls_InvQty, ls_InvPrice, ls_InvAmount)
                                Dim sqlComm2 As New SqlCommand(ls_SQL, sqlConn)
                                sqlComm2.ExecuteNonQuery()
                                sqlComm2.Dispose()
                            Next
                            'update master
                            ls_SQL = "update InvoiceOverseas_Master set TotalAmount = " & TotalAmount & " " & vbCrLf & _
                                     " Where InvoiceNo = '" & Trim(ls_InvNo) & "' " & vbCrLf & _
                                     " AND AffiliateID = '" & Trim(ls_AffiliateID) & "' " & vbCrLf & _
                                     " AND ShippingInstructionNo = '" & Trim(ls_SJno) & "' "

                            Dim sqlComm3 As New SqlCommand(ls_SQL, sqlConn)
                            sqlComm3.ExecuteNonQuery()
                            sqlComm3.Dispose()

                            sqlScope.Complete()
                        End With
                    Catch ex As Exception

                    End Try

                End Using
            End If
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_IsiMaster(ByVal pInvNo As String, ByVal pSuratJalanNo As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Replace((Replace(pInvNo, "'", "")), ",", "") <> "" Then
                ls_SQL = " SELECT * FROM InvoicePASI_Master WHERE InvoiceNo = '" & pInvNo & "'"
            Else
                ls_SQL = "SELECT InvoiceDate, IM.AffiliateID, MA.AffiliateName, InvoiceNo, Paymentterm, Duedate, totalAmount, remarks, IM.CurrCls, Description" & vbCrLf & _
                         " FROM InvoiceOverseas_master IM " & vbCrLf & _
                         " LEFT JOIN MS_AFFILIATE MA ON IM.AffiliateID = MA.AffiliateID" & vbCrLf & _
                         " LEFT JOIN MS_CurrCls MC On MC.CurrCls = IM.CurrCls " & vbCrLf & _
                         " WHERE InvoiceNo = '" & pInvNo & "' "
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 And ds.Tables(0).Rows.Count = 1 Then
                Try
                    With ds.Tables(0)
                        txtInvoiceDate.Text = Format(.Rows(0).Item("InvoiceDate"), "dd MMM yyyy")
                        txtPasiInvoiceNo.Text = Trim(.Rows(0).Item("InvoiceNo"))
                        txtaffiliatecode.Text = Trim(.Rows(0).Item("AffiliateID"))
                        txtaffiliatename.Text = Trim(.Rows(0).Item("AffiliateName"))
                        txtPaymentTerm.Text = Trim(.Rows(0).Item("Paymentterm"))
                        cbocurr.Text = Trim(.Rows(0).Item("Description"))
                        dtDueDate.Text = Format(Trim(.Rows(0).Item("DueDate")), "dd MMM yyyy")
                        txttotalamount.Text = Trim(.Rows(0).Item("totalamount"))
                        MmNotes.Text = Trim(.Rows(0).Item("remarks"))

                    End With
                Catch ex As Exception

                End Try
            Else
                txtPasiInvoiceNo.Text = ""
                txtInvoiceDate.Text = Format(Now, "dd MMM yyyy")
                dtDueDate.Text = Format(Now, "dd MMM yyyy")
                txttotalamount.Text = 0
                MmNotes.Text = ""
                
            End If
            sqlConn.Close()

        End Using
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                'If Not IsNothing(Request.QueryString("prm")) Then

                If Session("POListInv") <> "" Then
                    param = Session("POListInv").ToString()
                ElseIf Session("TampungInv") <> "" Then
                    param = Session("TampungInv").ToString()
                Else
                    param = Request.QueryString("prm").ToString
                End If

                If param = "  'back'" Then
                    btnsubmenu.Text = "BACK"
                Else
                    'If pStatus = False Then
                    up_fillcombo()
                    cbocurr.Text = "02"
                    txtcurr.Text = "USD"

                    pOrderNo = Split(param, "|")(0)
                    pAffiliateCode = Split(param, "|")(1)
                    pSupplierID = Split(param, "|")(2)
                    pSJ = Split(param, "|")(3)
                    pInvoiceNo = Split(param, "|")(4)
                    pShipping = Split(param, "|")(5)
                    pAffName = Split(param, "|")(6)
                    

                    If Session("InvInvoice") <> "" Then pInvoiceNo = Session("InvInvoice")

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"
                    'If Trim(pInvoiceDate) = "01 Jan 1900" Then pInvoiceDate = Format(Now, "dd MMM yyyy")
                    'txtInvoiceDate.Text = pInvoiceDate
                    'txtaffiliatecode.Text = pAffiliateCode
                    'txtaffiliatename.Text = pAffiliateName
                    'pShipping = Replace(pPasiSj, "'", "")


                    'txttotalamount.Text = uf_SumAmount(pPo, pKanban)
                    dtDueDate.Text = Format(CDate(Now), "dd MMM yyyy")
                    Session("shippingno") = pShipping
                    'pStatus = True

                    If pInvoiceNo <> "" Then
                        txtPasiInvoiceNo.Text = pInvoiceNo
                        MmNotes.Text = pNotes
                        Call up_IsiMaster(pInvoiceNo, pShipping)
                    End If
                    'txttotalbox.Text = Format(pkanbandate, "dd MMM yyyy")
                    'paramDT1 = pdt1
                    'paramDT2 = pdt2
                    'paramaffiliate = pcboaffiliate
                    'paramSupplier = ptxtsupplierID
                    txtaffiliatecode.Text = pAffiliateCode
                    txtaffiliatename.Text = pAffName
                    Call up_IsiMaster(pInvoiceNo, pShipping)
                    Call up_GridLoad(pAffiliateCode, pInvoiceNo, pShipping)

                    Session("InvoiceNoInv") = pInvoiceNo
                    Session("shipping") = pShipping
                    Session("AFF") = pAffiliateCode
                    Session("TampungInv") = param
                    Session.Remove("InvInvoice")
                    'End If
                End If

                btnsubmenu.Text = "BACK"
                'End If
            End If
            '===============================================================================

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblerrmessage.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    '    Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
    '        Dim ls_MsgID As String = ""
    '        Dim ls_sql As String = ""
    '        Dim iRow As Integer = 0
    '        Dim ls_User As String = Trim(Session("UserID").ToString)


    '        Dim ls_InvNo As String = Trim(txtPasiInvoiceNo.Text)
    '        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)
    '        Dim ls_SJno As String = Trim(txtPasisuratjalanno.Text)
    '        Dim ls_InvDate As Date = Trim(txtInvoiceDate.Text)
    '        ls_InvDate = Format(CDate(ls_InvDate), "yyyy-MM-dd")
    '        Dim ls_DueDate As Date = Trim(dtDueDate.Text)
    '        Dim ls_PaymentTerm As String = Trim(txtPaymentTerm.Text)
    '        Dim ls_TotalAmount As String = Trim(txttotalamount.Text)
    '        Dim ls_Notes As String = Trim(MmNotes.Text)

    '        Dim ls_Container As String = txtContainerNo.Text
    '        Dim ls_PlaceDate As String = txtPlaceDate.Text
    '        Dim ls_ShippedPer As String = txtShipperPer.Text
    '        Dim ls_OnOrAbout As String = txtOnOrAboutCondition.Text
    '        Dim ls_DeliveryTerm As String = txtDeliveryTerm.Text
    '        Dim ls_From As String = txtFrom.Text
    '        Dim ls_To As String = txtTo.Text
    '        Dim ls_Via As String = txtVia.Text
    '        Dim ls_Freight As String = txtFreight.Text

    '        Dim ls_POno As String
    '        Dim ls_POkanbanCls As String
    '        Dim ls_KanbanNo As String
    '        Dim ls_PartNo As String
    '        Dim ls_ReceiveQty As Double
    '        Dim ls_InvQty As Double
    '        Dim ls_InvCurr As String
    '        Dim ls_InvPrice As Double
    '        Dim ls_InvAmount As Double
    '        Dim ls_InvCartonNo As String



    '        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '            sqlConn.Open()

    '            If Grid.VisibleRowCount = 0 Then Exit Sub

    '            Using sqlTran As New TransactionScope

    '                'Save Master
    '                ls_sql = Save_Master(ls_InvNo, ls_AffiliateID, ls_SJno, ls_InvDate, ls_DueDate, ls_PaymentTerm, ls_TotalAmount, ls_Notes, ls_User,
    '                                     ls_Container, ls_PlaceDate, ls_ShippedPer, ls_OnOrAbout, ls_DeliveryTerm, ls_From, ls_To, ls_Via, ls_Freight)
    '                Dim sqlComm As New SqlCommand(ls_sql, sqlConn)
    '                sqlComm.ExecuteNonQuery()
    '                sqlComm.Dispose()

    '                For iRow = 0 To e.UpdateValues.Count - 1
    '                    ls_POno = e.UpdateValues(iRow).NewValues("colpono").ToString()
    '                    ls_POkanbanCls = e.UpdateValues(iRow).NewValues("colpokanban").ToString()
    '                    If ls_POkanbanCls = "YES" Then ls_POkanbanCls = "1" Else ls_POkanbanCls = "0"
    '                    ls_KanbanNo = e.UpdateValues(iRow).NewValues("colkanbanno").ToString()
    '                    ls_PartNo = e.UpdateValues(iRow).NewValues("colpartno").ToString()
    '                    ls_ReceiveQty = e.UpdateValues(iRow).NewValues("colAffRecQty").ToString()
    '                    ls_InvQty = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colInvoiceToAffQty").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvoiceToAffQty").ToString(), 0)))
    '                    ls_InvCurr = Trim((IIf(e.UpdateValues(iRow).NewValues("colInvCurr").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvCurr").ToString(), 0)))
    '                    ls_InvPrice = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colInvPrice").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvPrice").ToString(), 0)))
    '                    ls_InvAmount = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colInvAmount").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvAmount").ToString(), 0)))
    '                    ls_InvCartonNo = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colcartonno").ToString() <> "", e.UpdateValues(iRow).NewValues("colcartonno").ToString(), 0)))

    '                    'Save Detail
    '                    ls_sql = Save_Detail(ls_InvNo, ls_SJno, ls_AffiliateID, ls_POno, ls_POkanbanCls, ls_KanbanNo, ls_PartNo, ls_ReceiveQty, ls_InvQty, ls_InvCurr, ls_InvPrice, ls_InvAmount, ls_InvCartonNo)

    '                    Dim sqlComm2 As New SqlCommand(ls_sql, sqlConn)
    '                    sqlComm2.ExecuteNonQuery()
    '                    sqlComm2.Dispose()
    '                Next iRow
    '                sqlTran.Complete()
    '            End Using
    '            sqlConn.Close()
    '        End Using
    '    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Select Case pAction
            Case "gridload"
                Call up_GridLoad(Session("AFF"), Session("InvoiceNoInv"), Session("shipping"))
                Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text

                'Case "Delete"
                '    Call up_Delete()
                '    Call up_GridLoad(Session("POInv"), Session("KanbanInv"), Session("InvoiceNoInv"), Session("PasiSJ"))

                'Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                'Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Case "EDI"
                Dim pInv As String = Split(e.Parameters, "|")(1)
                Dim pAff As String = Split(e.Parameters, "|")(2)
                'Call SendEDIFile(pInv, pAff)
                Call up_SaveAll()

        End Select
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
    End Sub

    '    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
    '        If Not (e.DataColumn.FieldName = "colInvoiceToAffQty" Or e.DataColumn.FieldName = "colcartonno") Then
    '            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    '        Else
    '            e.Cell.BackColor = Color.White
    '        End If
    '    End Sub

    '    Private Sub Grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles Grid.HtmlRowPrepared
    '        Try
    '            Dim getRowValues As String = e.GetValue("colInvoiceToAffQty")
    '            If Not IsNothing(getRowValues) Then
    '                If getRowValues.Trim() <> "" Then
    '                    e.Row.BackColor = Color.FromName("#E0E0E0")
    '                End If
    '            End If
    '            Dim getRowValues2 As String = e.GetValue("colcartonno")
    '            If Not IsNothing(getRowValues2) Then
    '                If getRowValues2.Trim() <> "" Then
    '                    e.Row.BackColor = Color.FromName("#E0E0E0")
    '                End If
    '            End If

    '        Catch ex As Exception

    '        End Try
    '    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        'Session.Remove("POInv")
        'Session.Remove("KanbanInv")
        'Session.Remove("InvoiceNoInv")
        'Session.Remove("TampungInv")

        If btnsubmenu.Text = "BACK" Then
            Response.Redirect("~/Invoice/AffReceivingConfExport.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    '    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
    '        Session("InvInvoice") = txtPasiInvoiceNo.Text
    '        Response.Redirect("~/Invoice/viewInvToAff.aspx")

    '    End Sub

    '    Private Sub SendEDIFile(ByVal pInv As String, ByVal pAff As String)
    '        'Private Sub SendEDIFile()
    '        Dim fp As StreamWriter
    '        Dim ls_sql As String

    '        ls_sql = " SELECT  *  " & vbCrLf & _
    '                  "  FROM    ( SELECT DISTINCT  " & vbCrLf & _
    '                  "                      a = 'H00' + 'VD01    ' + CONVERT(CHAR(8), IVM.AffiliateID)  " & vbCrLf & _
    '                  "                      + CONVERT(CHAR(8), IVM.AffiliateID)  " & vbCrLf & _
    '                  "                      + CONVERT(CHAR(8), CONVERT(DATETIME, GETDATE()), 112)  " & vbCrLf & _
    '                  "                      + REPLACE(CONVERT (VARCHAR(8), GETDATE(), 108), ':', '')  " & vbCrLf & _
    '                  "                      + 'INVOICE-DATA   ' + CONVERT(CHAR(19), '') ,  " & vbCrLf & _
    '                  "                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                  "                      AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                  "            FROM      InvoicePasi_Master IVM  " & vbCrLf & _
    '                  "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo                                                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  "

    '        ls_sql = ls_sql + "                                                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "            UNION ALL  " & vbCrLf & _
    '                          "            SELECT DISTINCT  " & vbCrLf & _
    '                          "                      a = 'H10'  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(6), CONVERT(DATETIME, RAM.ReceiveDate), 112) " & vbCrLf & _
    '                          "                      + CASE ShipCls WHEN 'TRUCK' THEN 'T' WHEN 'BOAT' THEN 'B' WHEN 'AIR' THEN 'A' ELSE 'C' END " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(35), '') + CONVERT(CHAR(10), '')  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(15), IVM.ContainerNo) + '2'  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(4), '') ,  " & vbCrLf & _
    '                          "                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "                      AFFILIATEID = IVM.AFFILIATEID             "

    '        ls_sql = ls_sql + "                      FROM      InvoicePasi_Master IVM  " & vbCrLf & _
    '                          "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                                                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                                                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                      LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                                                               AND RAD.KanbanNo = IVD.Kanbanno  " & vbCrLf & _
    '                          "                                                               AND RAD.PONO = IVD.PONO  " & vbCrLf & _
    '                          "                      LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
    '                          "                                                               AND RAD.SuratJalanNo = RAM.SuratJalanNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN PO_Master PM ON PM.pono = IVD.PONO "

    '        ls_sql = ls_sql + " 								AND PM.AffiliateID = IVD.AffiliateID " & vbCrLf & _
    '                          "            UNION ALL            SELECT DISTINCT  " & vbCrLf & _
    '                          "                      a = 'H20'  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), CONVERT(DATETIME, KM.KanbanDate), 112)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), CONVERT(DATETIME, KM.KanbanDate), 112)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), CONVERT(DATETIME, KM.KanbanDate), 112)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), CONVERT(DATETIME, KM.KanbanDate), 112)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(20), IVM.InvFrom)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(20), IVM.InvTo) ,  " & vbCrLf & _
    '                          "                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "                      AFFILIATEID = IVM.AFFILIATEID  "

    '        ls_sql = ls_sql + "            FROM      InvoicePasi_Master IVM                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                                                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                                                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                      LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                                                    AND KD.KanbanNo = IVD.KanbanNo  " & vbCrLf & _
    '                          "                                                    AND KD.PONo = IVD.PONo  " & vbCrLf & _
    '                          "                                                    AND KD.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN Kanban_Master KM ON KM.AffiliateID = KD.AffiliateID  " & vbCrLf & _
    '                          "                                                    AND KM.SupplierID = KD.SupplierID  " & vbCrLf & _
    '                          "                                                    AND KM.DeliveryLocationCode = KD.DeliveryLocationCode  " & vbCrLf & _
    '                          "                                                    AND KM.KanbanNo = KD.KanbanNo            UNION ALL  "

    '        ls_sql = ls_sql + "            SELECT DISTINCT  " & vbCrLf & _
    '                          "                      a = 'H21' " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(20), IVM.InvVia) " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(15), '')  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(25), '')  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(12), '') ,  " & vbCrLf & _
    '                          "                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "                      AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "            UNION ALL  " & vbCrLf & _
    '                          "            SELECT DISTINCT  "

    '        ls_sql = ls_sql + "                      a = 'H30' + CONVERT(CHAR(8), 'PASI')                       " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(15), IVM.InvoiceNo)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(15), IVM.InvoiceNo)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), CONVERT(DATETIME, IVM.InvoiceDate), 112)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), CONVERT(DATETIME, IVM.InvoiceDate), 112)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(5), IVM.PaymentTerm)  " & vbCrLf & _
    '                          "                      + CASE WHEN IVM.Invfreight = 'COLLECT' THEN 'C'  " & vbCrLf & _
    '                          "                             WHEN IVM.InvFreight = 'PREPAID' THEN 'P'  " & vbCrLf & _
    '                          "                             ELSE CONVERT(CHAR(1), '')  " & vbCrLf & _
    '                          "                        END + 'C' + CONVERT(CHAR(4), '') + '0'  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(6), '') ,  "

    '        ls_sql = ls_sql + "                      INVOICENO = IVM.INVOICENO ,                      AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "            UNION ALL  " & vbCrLf & _
    '                          "            SELECT DISTINCT  " & vbCrLf & _
    '                          "                      a = 'H40'  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(15), CONVERT(NUMERIC(32, 0), IVM.TotalAmount))  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(6), CONVERT(NUMERIC(32, 0), SUM(IVD.InvQty)  " & vbCrLf & _
    '                          "                      / SUM(MP.BoxPallet))) + CONVERT(CHAR(8), IVM.affiliateID)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), IVM.affiliateID)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), IVM.affiliateID)  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(9), CONVERT(NUMERIC(32, 0), SUM(IVD.InvQty)                      * SUM(MP.NetWeight)))  "

    '        ls_sql = ls_sql + "                      + CONVERT(CHAR(9), CONVERT(NUMERIC(32, 0), SUM(IVD.InvQty)  " & vbCrLf & _
    '                          "                      * SUM(MP.GrossWeight))) + CONVERT(CHAR(9), '') ,  " & vbCrLf & _
    '                          "                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "                      AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                                                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                                                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "            GROUP BY  IVM.INVOICENO ,                      IVM.AFFILIATEID ,  " & vbCrLf & _
    '                          "                      IVM.TotalAmount ,  "

    '        ls_sql = ls_sql + "                      IVM.affiliateID  " & vbCrLf & _
    '                          "            UNION ALL  " & vbCrLf & _
    '                          "            SELECT DISTINCT  " & vbCrLf & _
    '                          "                      a = 'H41' + CONVERT(CHAR(15), '') + CONVERT(CHAR(20), '')  " & vbCrLf & _
    '                          "                      + CONVERT(CHAR(8), '') + CONVERT(CHAR(29), '') ,  " & vbCrLf & _
    '                          "                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "                      AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "                        -------------- DETAIL ------------------  " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  "

    '        ls_sql = ls_sql + "          a = 'D10'  " & vbCrLf & _
    '                          "  			+ CONVERT(char(25),IVD.PartNo)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(3), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(2), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(15), CONVERT(numeric(32,0),IVD.InvPrice))  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(3), MC.Description)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(3), MU.Description) " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(9), MP.QtyBox)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(1), '')  			+ CONVERT(CHAR(11), '')  " & vbCrLf & _
    '                          "  		,INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "          AFFILIATEID = IVM.AFFILIATEID  "

    '        ls_sql = ls_sql + "  					  " & vbCrLf & _
    '                          "  		FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls                        " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  " & vbCrLf & _
    '                          "          a = 'D11'  "

    '        ls_sql = ls_sql + "  			+ CONVERT(char(30),MP.PartName)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(25), IVD.PartNo)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(3), 'EA')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(4), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(10), '')  " & vbCrLf & _
    '                          "  		,INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "          AFFILIATEID = IVM.AFFILIATEID  					  " & vbCrLf & _
    '                          "  		FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                          AND IVM.AffiliateID = IVD.AffiliateID  "

    '        ls_sql = ls_sql + "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  " & vbCrLf & _
    '                          "          a = 'D12' " & vbCrLf & _
    '                          " 			+ CONVERT(CHAR(3),'') " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(9), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(25), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(35), '')  " & vbCrLf & _
    '                          "  		,INVOICENO = IVM.INVOICENO ,  "

    '        ls_sql = ls_sql + "          AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "  					  " & vbCrLf & _
    '                          "  		FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                          AND IVM.AffiliateID = IVD.AffiliateID                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
    '                          "            " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  "

    '        ls_sql = ls_sql + "          a = 'D20'  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(15), IVD.PONo)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(1), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(9), CONVERT(numeric(32,0), IVD.InvQty))  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(9), CONVERT(numeric(32,0), IVD.InvQty))  			 " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(15), CONVERT(numeric(32,0), IVD.InvAmount))  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(23),'')  " & vbCrLf & _
    '                          "  		,INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "          AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "  					  " & vbCrLf & _
    '                          "  		FROM      INVOICEPASI_MASTER IVM  "

    '        ls_sql = ls_sql + "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
    '                          "                        " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  " & vbCrLf & _
    '                          "          a = 'D30'  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8), IVD.InvCartonNo)  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8), IVD.InvCartonNo)  "

    '        ls_sql = ls_sql + "  			+ CONVERT(CHAR(5), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8), CONVERT(NUMERIC(32, 0), IVD.InvQty * MP.NetWeight))  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8), CONVERT(NUMERIC(32, 0), ISNULL(IVD.InvQty,0)* ISNULL(MP.GrossWeight,0)))  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(3),'KG')  			 " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8), LEFT(((ISNULL(MP.Width,0) * ISNULL(MP.Length,0)* ISNULL(MP.Height,0)) * ISNULL(IVD.InvQty,0)),8))  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(3), 'MM3')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8),'')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(8),'')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(5),'')	  " & vbCrLf & _
    '                          "  		,INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "          AFFILIATEID = IVM.AFFILIATEID  "

    '        ls_sql = ls_sql + "  					  " & vbCrLf & _
    '                          "  		FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                          AND IVM.SuratJalanNo = IVD.SuratJalanNo                          AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
    '                          "            " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  " & vbCrLf & _
    '                          "          a = 'D31'  "

    '        ls_sql = ls_sql + "  			+ CONVERT(CHAR(10), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(10), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(7), '')  			 " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(10), '')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(1), '0')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(15),'')  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(19),'')  " & vbCrLf & _
    '                          "  		,INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "          AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "  					  " & vbCrLf & _
    '                          "  		FROM      INVOICEPASI_MASTER IVM  "

    '        ls_sql = ls_sql + "                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
    '                          "                          AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
    '                          "                          AND IVM.AffiliateID = IVD.AffiliateID                       " & vbCrLf & _
    '                          "                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
    '                          "                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
    '                          "          -------------- FOOTER --------------------  " & vbCrLf & _
    '                          "          UNION ALL  " & vbCrLf & _
    '                          "          SELECT DISTINCT  " & vbCrLf & _
    '                          "          a = 'T00'  " & vbCrLf & _
    '                          "  			+ CONVERT(CHAR(72), '')  "

    '        ls_sql = ls_sql + "  		,INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
    '                          "          AFFILIATEID = IVM.AFFILIATEID  " & vbCrLf & _
    '                          "  					  		FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
    '                          "          ) xx  "



    '        ls_sql = ls_sql + " WHERE   InvoiceNo = '" & pInv & "' " & vbCrLf & _
    '                          "         AND AffiliateID = '" & pAff & "' "



    '        Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '            cn.Open()
    '            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
    '            Dim ds As New DataSet
    '            sqlDA.Fill(ds)

    '            If ds.Tables(0).Rows.Count > 0 Then

    '                'Dim fi As New FileInfo(Server.MapPath("~\Invoice\" & txtaffiliatecode.Text & "INV.txt"))
    '                'If fi.Exists Then
    '                '    fi.Delete()
    '                '    fi = New FileInfo(Server.MapPath("~\Invoice\" & txtaffiliatecode.Text & "INV.txt"))
    '                'End If

    '                'DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

    '                fp = File.CreateText(Server.MapPath("~\Invoice\" & txtaffiliatecode.Text & "INV.txt"))

    '                For x = 0 To ds.Tables(0).Rows.Count - 1
    '                    fp.WriteLine(ds.Tables(0).Rows(x)("a") & Format(x + 1, "00000"))
    '                Next

    '                fp.Close()
    '                'fi.Delete()

    '            End If
    '        End Using




    '        Exit Sub
    'ErrHandler:
    '        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    '    End Sub
End Class