Imports System.Data.SqlClient
Imports System.Transactions
Imports System.Drawing
Imports System.IO

Public Class InvToAff
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    ''parameter

    Dim pInvoiceDate As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pPasiSj As String
    Dim pPasiInvoiceno As String
    Dim pPaymentTerm As String
    Dim pDueDate As String

    Dim pNotes As String
    Dim pPo As String
    Dim pKanban As String

#End Region

    Private Sub up_GridLoad(ByVal pPO As String, ByVal pKanban As String, ByVal pInvNo As String, ByVal pPSJ As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If Replace(Replace(pInvNo, "'", ""), ",", "") = "" Then
            pInvNo = ""
        End If

        If Replace(Replace(pPSJ, "'", ""), ",", "") = "" Then
            pPSJ = ""
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'ls_SQL = "     SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colpono,colpasisj, colkanbanno, colpartno )), * FROM( " & vbCrLf & _
            '         "     SELECT    DISTINCT  " & vbCrLf & _
            '         "              colpono = POM.PONo ,  " & vbCrLf & _
            '         "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'  " & vbCrLf & _
            '         "                                 ELSE 'YES'  " & vbCrLf & _
            '         "                            END ,  " & vbCrLf & _
            '         "              colkanbanno = ISNULL(KD.KanbanNo, '') ,  " & vbCrLf & _
            '         "              colpartno = POD.PartNo ,  " & vbCrLf & _
            '         "              colpartname = MP.PartName ,  " & vbCrLf & _
            '         "              coluom = UC.Description ,  " & vbCrLf & _
            '         "              colCls = UC.unitcls ,  " & vbCrLf

            'ls_SQL = ls_SQL + "              colQtyBox = ROUND(CONVERT(CHAR, ISNULL(MPM.QtyBox, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colpasideliveryqty = ROUND(CONVERT(CHAR, ISNULL(PDD.DoQty, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colAffRecQty = ROUND(CONVERT(CHAR, ISNULL(RAD.RecQty, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colInvoiceToAffQty = ROUND(CONVERT(CHAR, COALESCE(IPD.InvQty,  " & vbCrLf & _
            '                  "                                                                RAD.RecQty, 0), 0),  " & vbCrLf & _
            '                  "                                         0) ,  " & vbCrLf & _
            '                  "              coldelqtybox = ROUND(CONVERT(CHAR, ISNULL(PDD.DoQty, 0) / ISNULL(MPM.QtyBox, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colInvCurr = ISNULL(MC.Description, '') ,  " & vbCrLf & _
            '                  "              colInvPrice = ROUND(CONVERT(CHAR, ISNULL(MPR.Price, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colInvAmount = ROUND(CONVERT(CHAR, COALESCE(IPD.InvAmount,  " & vbCrLf & _
            '                  "                                                          ( RAD.RecQty  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                            * MPR.Price ), 0), 0),  " & vbCrLf & _
            '                  "                                   0) ,  " & vbCrLf & _
            '                  "              colcartonno = COALESCE(PCD_PASI.cartonNo,IPD.InvCartonNo),   " & vbCrLf & _
            '                  "              colpasisj = PDM.SuratJalanNo, InvoiceNo = isnull(IPD.InvoiceNo,'')  " & vbCrLf & _
            '                  "     FROM     dbo.PO_Master POM  " & vbCrLf & _
            '                  "              LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
            '                  "                                         AND POM.PoNo = POD.PONo  " & vbCrLf & _
            '                  "                                         AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
            '                  "                                                AND KD.PoNo = POD.PONo  " & vbCrLf & _
            '                  "                                                AND KD.SupplierID = POD.SupplierID  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                AND KD.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID  " & vbCrLf & _
            '                  "                                                AND KD.KanbanNo = KM.KanbanNo  " & vbCrLf & _
            '                  "                                                AND KD.SupplierID = KM.SupplierID  " & vbCrLf & _
            '                  "                                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
            '                  "                                                     AND KD.KanbanNo = SDD.KanbanNo  " & vbCrLf & _
            '                  "                                                     AND KD.SupplierID = SDD.SupplierID  " & vbCrLf & _
            '                  "                                                     AND KD.PartNo = SDD.PartNo  " & vbCrLf & _
            '                  "                                                     AND KD.PoNo = SDD.PoNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                     AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                     AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf & _
            '                  "              INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
            '                  "                                                       AND SDD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
            '                  "                                                       AND SDD.SupplierID = PRD.SupplierID  " & vbCrLf & _
            '                  "                                                       AND SDD.PartNo = PRD.PartNo  " & vbCrLf & _
            '                  "                                                       AND SDD.PONo = PRD.PONo  " & vbCrLf & _
            '                  "                                                       AND PRD.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
            '                  "                                                      AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf & _
            '                  "              LEFT JOIN ( SELECT  SuratJalanno ,  " & vbCrLf

            'ls_SQL = ls_SQL + "                                  SupplierID ,  " & vbCrLf & _
            '                  "                                  AffiliateID ,  " & vbCrLf & _
            '                  "                                  PONO ,  " & vbCrLf & _
            '                  "                                  KanbanNO ,  " & vbCrLf & _
            '                  "                                  Partno ,  " & vbCrLf & _
            '                  "                                  UnitCls ,  " & vbCrLf & _
            '                  "                                  DoQty = SUM(ISNULL(DoQty, 0))  " & vbCrLf & _
            '                  "                          FROM    DOPasi_Detail  " & vbCrLf & _
            '                  "                          GROUP BY SuratJalanno ,  " & vbCrLf & _
            '                  "                                  SupplierID ,  " & vbCrLf & _
            '                  "                                  AffiliateID ,  " & vbCrLf

            'ls_SQL = ls_SQL + "                                  PONO ,  " & vbCrLf & _
            '                  "                                  KanbanNO ,  " & vbCrLf & _
            '                  "                                  Partno ,  " & vbCrLf & _
            '                  "                                  UnitCls  " & vbCrLf & _
            '                  "                        ) PDD ON PRD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
            '                  "                                 AND PRD.KanbanNo = PDD.KanbanNo  " & vbCrLf & _
            '                  "                                 AND PRD.SupplierID = PDD.SupplierID  " & vbCrLf & _
            '                  "                                 AND PRD.PartNo = PDD.PartNo  " & vbCrLf & _
            '                  "                                 AND PRD.PoNo = PDD.PoNo     " & vbCrLf & _
            '                  "              LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
            '                  "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf

            'ls_SQL = ls_SQL + "              INNER JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                            AND PDD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
            '                  "                                                            AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "                                                            AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
            '                  "                                                            AND PDD.PoNo = RAD.PoNo  " & vbCrLf & _
            '                  "              INNER JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                            AND RAM.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
            '                  "              LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
            '                  "                                                      AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
            '                  "                                                      AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
            '                  "                                                      AND RAD.PONo = IPD.PONo  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                      AND RAD.SuratJalanNo = IPD.SuratJalanNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
            '                  "                                                      AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
            '                  "                                                      --AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "              LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "              --LEFT JOIN dbo.MS_Price MPR ON MPR.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                            --AND MPR.PartNo = RAD.PartNo  " & vbCrLf

            'ls_SQL = ls_SQL + "                                            --AND RAM.ReceiveDate BETWEEN MPR.StartDate  " & vbCrLf & _
            '                  "                                                                --AND  " & vbCrLf & _
            '                  "                                                                --MPR.EndDate  " & vbCrLf & _
            '                  "                                            --AND RAM.ReceiveDate >= MPR.Effectivedate " & vbCrLf & _
            '                  " LEFT JOIN (Select max(effectiveDate) as effectiveDate, affiliateID, Partno, currcls,PackingCls, " & vbCrLf & _
            '                  "PriceCls, StartDate, EndDate, DeliveryLocationID from MS_Price Group By affiliateID, Partno, currcls,PackingCls, " & vbCrLf & _
            '                  "PriceCls, StartDate, EndDate, DeliveryLocationID " & vbCrLf & _
            '                  ")XX ON XX.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "AND XX.PartNo = RAD.PartNo " & vbCrLf & _
            '                  "AND RAM.ReceiveDate BETWEEN XX.StartDate AND  XX.EndDate   " & vbCrLf & _
            '                  "AND RAM.ReceiveDate >= XX.Effectivedate " & vbCrLf & _
            '                  "LEFT JOIn MS_Price MPR ON MPR.AffiliateID = RAD.AffiliateID   " & vbCrLf & _
            '                  "AND MPR.PartNo = RAD.PartNo   " & vbCrLf & _
            '                  "AND  " & vbCrLf & _
            '                  "RAM.ReceiveDate BETWEEN MPR.StartDate   " & vbCrLf & _
            '                  "AND   " & vbCrLf & _
            '                  "MPR.EndDate " & vbCrLf & _
            '                  "AND XX.Effectivedate = MPR.Effectivedate " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_CurrCls MC ON MC.CurrCls = MPR.CurrCls  " & vbCrLf & _
            '                  "              LEFT JOIN PLPASI_Detail PCD_PASI ON PCD_PASI.SuratJalanNo = PDD.suratJalanNo   " & vbCrLf & _
            '                  "  											AND PCD_PASI.SupplierID = PDD.SupplierID   " & vbCrLf & _
            '                  "                           					AND PCD_PASI.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "  											AND PCD_PASI.POnO = PDD.PONo   " & vbCrLf & _
            '                  "  				                            AND PCD_PASI.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "  				                            AND PCD_PASI.KanbanNo = PDD.KanbanNo " & vbCrLf

            ls_SQL = "  SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colpono,colpasisj, colkanbanno, colpartno )), * FROM(  " & vbCrLf & _
                      "  SELECT DISTINCT   " & vbCrLf & _
                      "     colpono = DD.PONo ,   " & vbCrLf & _
                      "     colpokanban = CASE WHEN ISNULL(DD.POKanbanCls, '0') = '0' THEN 'NO'   " & vbCrLf & _
                      "                         ELSE 'YES'   " & vbCrLf & _
                      "                 END ,   " & vbCrLf & _
                      "     colkanbanno = ISNULL(DD.KanbanNo, '') ,   " & vbCrLf & _
                      "     colpartno = DD.PartNo ,   " & vbCrLf & _
                      "     colpartname = MP.PartName ,   " & vbCrLf & _
                      "     coluom = UC.Description ,   " & vbCrLf & _
                      "     colCls = UC.unitcls ,   " & vbCrLf

            ls_SQL = ls_SQL + "     colQtyBox = ROUND(CONVERT(CHAR, ISNULL(MPM.QtyBox, 0), 0), 0) ,   " & vbCrLf & _
                              "     colpasideliveryqty = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0), 0), 0) ,   " & vbCrLf & _
                              "     colAffRecQty = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0), 0), 0) ,  " & vbCrLf & _
                              "     colInvoiceToAffQty = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0), 0), 0) ,  " & vbCrLf & _
                              "     coldelqtybox = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0) / ISNULL(MPM.QtyBox, 0), 0), 0) ,   " & vbCrLf & _
                              "     colInvCurr = ISNULL(MC.Description, '') ,   " & vbCrLf & _
                              "     colInvPrice = ROUND(CONVERT(CHAR, ISNULL(DD.Price,0), 0), 0) ,   " & vbCrLf & _
                              "     colInvAmount = ROUND(CONVERT(CHAR, COALESCE(IPD.InvAmount,   " & vbCrLf & _
                              "                                                 ( DD.DoQty   " & vbCrLf & _
                              "                                                 * ISNULL(DD.Price,0)), 0), 0),   " & vbCrLf & _
                              "                         0) ,   " & vbCrLf

            ls_SQL = ls_SQL + "     colcartonno = COALESCE(PCD_PASI.cartonNo,IPD.InvCartonNo),    " & vbCrLf & _
                              "     colpasisj = DM.SuratJalanNo, InvoiceNo = isnull(IPD.InvoiceNo,'')   " & vbCrLf & _
                              " FROM DOPASI_Master DM " & vbCrLf & _
                              " INNER JOIN DOPASI_Detail DD ON DM.SuratJalanNo = DD.SuratJalanNo and DM.AffiliateID = DD.AffiliateID " & vbCrLf & _
                              " LEFT JOIN InvoicePASI_Detail IPD ON DD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
                              " 									AND DD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
                              " 									AND DD.PartNo = IPD.PartNo  " & vbCrLf & _
                              " 									AND DD.PONo = IPD.PONo  " & vbCrLf & _
                              " 									AND DD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
                              " LEFT JOIN PLPASI_Detail PCD_PASI ON PCD_PASI.SuratJalanNo = DD.suratJalanNo    " & vbCrLf & _
                              "   									AND PCD_PASI.SupplierID = DD.SupplierID    " & vbCrLf

            ls_SQL = ls_SQL + "                            			AND PCD_PASI.AffiliateID = DD.AffiliateID    " & vbCrLf & _
                              "   									AND PCD_PASI.POnO = DD.PONo    " & vbCrLf & _
                              "   				                    AND PCD_PASI.PartNo = DD.PartNo    " & vbCrLf & _
                              "   				                    AND PCD_PASI.KanbanNo = DD.KanbanNo  " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON MP.PartNo = DD.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DD.PartNo AND MPM.AffiliateID = DD.AffiliateID AND MPM.SupplierID = DD.SupplierID " & vbCrLf & _
                              " LEFT JOIN MS_Price MPR on MPR.PartNo = DD.PartNo and MPR.AffiliateID = DD.AffiliateID and (DM.DeliveryDate between MPR.EffectiveDate and MPR.EndDate) " & vbCrLf & _
                              " LEFT JOIN MS_CurrCls MC ON MC.CurrCls = MPR.CurrCls " & vbCrLf

            ls_SQL = ls_SQL + " WHERE DM.SuratJalanNo IN (" & pPSJ & ") " & vbCrLf & _
                              " )x " & vbCrLf & _
                              "  "

            'If pInvNo = "" Then
            '    ls_SQL = ls_SQL + "   WHERE     --POM.PONo IN (" & pPO & ") " & vbCrLf & _
            '                      "             --AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf & _
            '                      "             --AND " & vbCrLf & _
            '                      "             PDD.SuratJalanNo IN (" & pPSJ & ") " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + "   WHERE     IPD.InvoiceNo ='" & pInvNo & "' " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " )x "
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

    Private Function uf_SumAmount(ByVal pInvNo As String, ByVal pPSJ As String, ByVal pInvDate As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If pInvDate = "01 Jan 1900 " Or pInvDate = "01/01/1900" Then
                'ls_SQL = " SELECT colInvAmount = ROUND(CONVERT(Char,ISNULL(SUM(colInvAmount),0),0),0) " & vbCrLf & _
                '      " FROM ( " & vbCrLf & _
                '      "    SELECT    colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo, POD.PartNo )) ,  " & vbCrLf & _
                '      "              colpono = POM.PONo ,  " & vbCrLf & _
                '      "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'  " & vbCrLf & _
                '      "                                 ELSE 'YES'  " & vbCrLf & _
                '      "                            END ,  " & vbCrLf & _
                '      "              colkanbanno = CASE WHEN POD.KanbanCls = '0' THEN '-'  " & vbCrLf & _
                '      "                                 ELSE ISNULL(KD.KanbanNo, '')  " & vbCrLf & _
                '      "                            END ,  " & vbCrLf & _
                '      "              colpartno = POD.PartNo ,  " & vbCrLf

                'ls_SQL = ls_SQL + "              colpartname = MP.PartName ,  " & vbCrLf & _
                '                  "              coluom = UC.Description ,  " & vbCrLf & _
                '                  "              colCls = UC.unitcls ,  " & vbCrLf & _
                '                  "              colQtyBox = ISNULL(MPM.QtyBox, 0) ,  " & vbCrLf & _
                '                  "              colpasideliveryqty = COALESCE(PDD.DOQty,  " & vbCrLf & _
                '                  "                                            ISNULL(SDD.DOQty, 0)  " & vbCrLf & _
                '                  "                                            - ( ISNULL(PRD.GoodRecQty, 0)  " & vbCrLf & _
                '                  "                                                + ISNULL(PRD.DefectRecQty, 0) ),  " & vbCrLf & _
                '                  "                                            0) ,  " & vbCrLf & _
                '                  "              colAffRecQty = ISNULL(RAD.RecQty,0) ,  " & vbCrLf & _
                '                  "              colInvoiceToAffQty = COALESCE(IPD.InvQty,RAD.RecQty,0) ,  " & vbCrLf

                'ls_SQL = ls_SQL + "              coldelqtybox = CASE MPM.QtyBox  " & vbCrLf & _
                '                  "                               WHEN 0 THEN 0  " & vbCrLf & _
                '                  "                               ELSE ISNULL(SDD.DOQty, 0) / MPM.QtyBox  " & vbCrLf & _
                '                  "                             END ,  " & vbCrLf & _
                '                  "              colInvCurr = ISNULL(MPR.CurrCls,'') ,  " & vbCrLf & _
                '                  "              colInvPrice = ISNULL(MPR.Price,0) ,  " & vbCrLf & _
                '                  "              colInvAmount = COALESCE(IPD.InvAmount,(RAD.RecQty*MPR.Price),0 ) " & vbCrLf & _
                '                  "   FROM      dbo.PO_Master POM " & vbCrLf & _
                '                  "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                '                  "                                        AND POM.PoNo = POD.PONo " & vbCrLf & _
                '                  "                                        AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
                '                  "                                               AND KD.PoNo = POD.PONo " & vbCrLf

                'ls_SQL = ls_SQL + "                                               AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
                '                  "                                               AND KD.PartNo = POD.PartNo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                '                  "                                               AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                '                  "                                               AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                '                  "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                '                  "             LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                '                  "                                                    AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
                '                  "                                                    AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
                '                  "                                                    AND KD.PartNo = SDD.PartNo " & vbCrLf & _
                '                  "                                                    AND KD.PONo = SDD.PONo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf

                'ls_SQL = ls_SQL + "                                                    AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf & _
                '                  "                                                    AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                '                  "                                                     AND KD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                '                  "                                                     AND KD.SupplierID = PRD.SupplierID " & vbCrLf & _
                '                  "                                                     AND KD.PartNo = PRD.PartNo " & vbCrLf & _
                '                  "                                                     AND KD.PONo = PRD.PartNo " & vbCrLf & _
                '                  "                                                     AND PRD.SuratJalanNo = SDM.SuratJalanNo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                '                  "                                                     AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                '                  "                                                     AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID " & vbCrLf

                'ls_SQL = ls_SQL + "                                                AND KD.KanbanNo = PDD.KanbanNo " & vbCrLf & _
                '                  "                                                AND KD.SupplierID = PDD.SupplierID " & vbCrLf & _
                '                  "                                                AND KD.PartNo = PDD.PartNo " & vbCrLf & _
                '                  "                                                AND KD.PoNo = PDD.PoNo " & vbCrLf & _
                '                  "                                                AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                '                  "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
                '                  "                                                AND PDD.SupplierID = PDM.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                '                  "                                                          AND PDD.KanbanNo = RAD.KanbanNo " & vbCrLf & _
                '                  "                                                          AND PDD.SupplierID = RAD.SupplierID " & vbCrLf & _
                '                  "                                                          AND PDD.PartNo = RAD.PartNo " & vbCrLf

                'ls_SQL = ls_SQL + "                                                          AND PDD.PoNo = RAD.PoNo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                '                  "                                                          AND RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                '                  "                                                          AND RAM.SupplierID = RAD.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.InvoicePASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID " & vbCrLf & _
                '                  "                                                     AND PDD.KanbanNo = IPD.KanbanNo " & vbCrLf & _
                '                  "                                                     AND PDD.PartNo = IPD.PartNo " & vbCrLf & _
                '                  "                                                     AND PDD.PONo = IPD.PONo " & vbCrLf & _
                '                  "                                                     AND PDD.SuratJalanNo = IPD.SuratJalanNo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID " & vbCrLf & _
                '                  "                                                     AND IPD.InvoiceNo = IPM.InvoiceNo " & vbCrLf

                'ls_SQL = ls_SQL + "                                                     --AND IPD.SuratJalanNo = IPM.SuratJalanNo " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                '                  "              LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_Price MPR ON MPR.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                '                  "                                             AND MPR.PartNo = RAD.PartNo  " & vbCrLf & _
                '                  "                                             AND RAM.ReceiveDate BETWEEN MPR.StartDate AND MPR.EndDate " & vbCrLf & _
                '                  "             LEFT JOIN dbo.MS_CurrCls MC ON MC.CurrCls = MPR.CurrCls" & vbCrLf & _
                '                  "             LEFT JOIN PLPASI_Detail PCD_PASI ON PCD_PASI.SuratJalanNo = PDD.suratJalanNo " & vbCrLf & _
                '                  "											AND PCD_PASI.SupplierID = PDD.SupplierID " & vbCrLf & _
                '                  "                         				AND PCD_PASI.AffiliateID = PDD.AffiliateID " & vbCrLf & _
                '                  "											AND PCD_PASI.POnO = PDD.PONo " & vbCrLf & _
                '                  "				                            AND PCD_PASI.PartNo = PDD.PartNo " & vbCrLf & _
                '                  "                             			AND PDD.SuratJalanNoSupplier = SDM.SuratJalanNo   " & vbCrLf & _
                '                  "    WHERE     PDD.SuratJalanNo IN (" & pPSJ & ")  " & vbCrLf & _
                '                  " ) Inv " & vbCrLf & _
                '                  "  "
                ls_SQL = " SELECT colInvAmount = SUM(ISNULL(DD.DOQty * ISNULL(DD.Price,0),0)) " & vbCrLf & _
                          " FROM DOPASI_Master DM " & vbCrLf & _
                          " INNER JOIN DOPASI_Detail DD ON DM.AffiliateID = DD.AffiliateID AND DM.SuratJalanNo = DD.SuratJalanNo  " & vbCrLf & _
                          " LEFT JOIN MS_Price MPR on MPR.PartNo = DD.PartNo and MPR.AffiliateID = DD.AffiliateID and (DM.DeliveryDate between MPR.EffectiveDate and MPR.EndDate) " & vbCrLf & _
                          " WHERE DM.SuratJalanNo IN (" & pPSJ & ")  " & vbCrLf & _
                          "  "

            Else
                ls_SQL = "select colInvAmount= isnull(totalamount,0) from Invoicepasi_master where invoiceno = '" & pInvNo & "' " & vbCrLf
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            uf_SumAmount = ds.Tables(0).Rows(0)("colInvAmount")
            sqlConn.Close()

        End Using
    End Function

    Private Function Update_Master(ByVal pinvno As String, ByVal paffiliateID As String)
        Dim ls_sql As String

        ls_sql = " update invoicePasi_Master  " & vbCrLf & _
                 " SET totalamount = (select SUM(isnull(InvQty,0)*isnull(InvPrice,0))  " & vbCrLf & _
                 " from invoicePasi_Detail where suratjalanno = '" & Trim(pinvno) & "') " & vbCrLf & _
                 " where suratjalanno = '" & Trim(pinvno) & "' " & vbCrLf
        Update_Master = ls_sql
    End Function

    Private Function Save_Master(ByVal pInvNo As String, ByVal pAffiliateID As String, ByVal pSJno As String, ByVal pInvDate As Date, ByVal pDueDate As Date, _
                            ByVal pPaymentTerm As String, ByVal pTotalAmount As String, ByVal pNotes As String, ByVal pUser As String, _
                            ByVal pContainer As String, ByVal pPlaceDate As String, ByVal pShippedPer As String, ByVal pOnOrAbout As String, _
                            ByVal pDeliveryTerm As String, ByVal pInvFrom As String, ByVal pInvTo As String, ByVal pInvVia As String, ByVal pFreight As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM InvoicePASI_Master WHERE InvoiceNo = '" & pInvNo & "' AND AffiliateID = '" & pAffiliateID & "' AND SuratJalanNo = '" & pSJno & "') " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " UPDATE dbo.InvoicePASI_Master " & vbCrLf & _
                          " SET InvoiceDate ='" & pInvDate & "', " & vbCrLf & _
                          " 	DueDate ='" & pDueDate & "', " & vbCrLf & _
                          " 	PaymentTerm ='" & pPaymentTerm & "', " & vbCrLf & _
                          " 	TotalAmount ='" & pTotalAmount & "', " & vbCrLf & _
                          "     Notes ='" & pNotes & "', " & vbCrLf & _
                          "     ContainerNo='" & pContainer & "', " & vbCrLf & _
                          "     PlaceDate='" & pPlaceDate & "', " & vbCrLf & _
                          "     ShippedPer='" & pShippedPer & "', " & vbCrLf & _
                          "     OnOrAboutCondition='" & pOnOrAbout & "', " & vbCrLf & _
                          "     DeliveryTerm='" & pDeliveryTerm & "', " & vbCrLf & _
                          "     InvFrom='" & pInvFrom & "', " & vbCrLf & _
                          "     InvTo='" & pInvNo & "', " & vbCrLf & _
                          "     InvVia='" & pInvVia & "', " & vbCrLf & _
                          "     InvFreight='" & pFreight & "', " & vbCrLf & _
                          "     UpdateDate = GETDATE(), " & vbCrLf

        ls_sql = ls_sql + "     UpdateUser ='" & pUser & "' " & vbCrLf & _
                          " WHERE InvoiceNo = '" & pInvNo & "'  " & vbCrLf & _
                          "   AND AffiliateID = '" & pAffiliateID & "'  " & vbCrLf & _
                          "   AND SuratJalanNo = '" & pSJno & "' " & vbCrLf & _
                          " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " INSERT INTO dbo.InvoicePASI_Master " & vbCrLf & _
                          "         ( InvoiceNo ,AffiliateID ,SuratJalanNo ,InvoiceDate ,DueDate , " & vbCrLf & _
                          "           PaymentTerm ,TotalAmount ,Notes ,ContainerNo ,PlaceDate ,ShippedPer ," & vbCrLf & _
                          "           OnOrAboutCondition ,DeliveryTerm ,InvFrom ,InvTo ,InvVia ,InvFreight ,EntryDate ,EntryUser  " & vbCrLf & _
                          "         ) " & vbCrLf

        ls_sql = ls_sql + " VALUES  ( '" & pInvNo & "' , " & vbCrLf & _
                          "           '" & pAffiliateID & "' , " & vbCrLf & _
                          "           '" & pSJno & "' , " & vbCrLf & _
                          "           '" & pInvDate & "' , " & vbCrLf & _
                          "           '" & pDueDate & "' , " & vbCrLf & _
                          "           '" & pPaymentTerm & "' , " & vbCrLf & _
                          "           '" & pTotalAmount & "' , " & vbCrLf & _
                          "           '" & pNotes & "' , " & vbCrLf & _
                          "           '" & pContainer & "' , " & vbCrLf & _
                          "           '" & pPlaceDate & "' , " & vbCrLf & _
                          "           '" & pShippedPer & "' , " & vbCrLf & _
                          "           '" & pOnOrAbout & "' , " & vbCrLf & _
                          "           '" & pDeliveryTerm & "' , " & vbCrLf & _
                          "           '" & pInvFrom & "' , " & vbCrLf & _
                          "           '" & pInvTo & "' , " & vbCrLf & _
                          "           '" & pInvVia & "' , " & vbCrLf & _
                          "           '" & pFreight & "' , " & vbCrLf & _
                          "           GETDATE() , " & vbCrLf & _
                          "           '" & pUser & "' " & vbCrLf & _
                          "         )  " & vbCrLf & _
                          " END " & vbCrLf

        Save_Master = ls_sql
    End Function

    Private Function Save_Detail(ByVal pInvNo As String, ByVal pSJno As String, ByVal pAffiliateID As String, ByVal pPOno As String, ByVal pPOKanbanCls As String, _
                            ByVal pKanbanNo As String, ByVal pPartNo As String, ByVal pReceiveQty As String, ByVal pInvQty As String, ByVal pInvCurr As String, ByVal pInvPrice As String, ByVal pInvAmount As String, ByVal pCartonNo As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM dbo.InvoicePASI_Detail WHERE InvoiceNo = '" & pInvNo & "' AND AffiliateID = '" & pAffiliateID & "' AND SuratJalanNo = '" & pSJno & "' AND PONo = '" & pPOno & "' AND PartNo = '" & pPartNo & "' and kanbanno = '" & pKanbanNo & "') " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " UPDATE dbo.InvoicePASI_Detail " & vbCrLf & _
                          " SET POKanbanCls ='" & pPOKanbanCls & "', " & vbCrLf & _
                          " 	KanbanNo ='" & pKanbanNo & "', " & vbCrLf & _
                          " 	ReceiveQty ='" & pReceiveQty & "', " & vbCrLf & _
                          " 	InvQty ='" & pInvQty & "'," & vbCrLf & _
                          " 	InvCurrCls =(SELECT CurrCls FROM MS_CurrCls WHERE DESCRIPTION = '" & pInvCurr & "'), " & vbCrLf & _
                          " 	InvPrice ='" & pInvPrice & "', " & vbCrLf & _
                          " 	InvAmount ='" & pInvAmount & "',   " & vbCrLf & _
                          "     InvCartonNo = '" & pCartonNo & "'" & vbCrLf & _
                          " WHERE InvoiceNo = '" & pInvNo & "'  " & vbCrLf & _
                          "   AND SuratJalanNo = '" & pSJno & "'  " & vbCrLf & _
                          "   AND AffiliateID = '" & pAffiliateID & "'  " & vbCrLf & _
                          "   AND PONo = '" & pPOno & "' " & vbCrLf & _
                          "   AND PartNo = '" & pPartNo & "' " & vbCrLf & _
                          "   AND KanbanNo = '" & pKanbanNo & "' " & vbCrLf

        ls_sql = ls_sql + " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " INSERT INTO dbo.InvoicePASI_Detail " & vbCrLf & _
                          "         ( InvoiceNo,SuratJalanNo,AffiliateID,PONo,POKanbanCls,KanbanNo, " & vbCrLf & _
                          "           PartNo,ReceiveQty,InvQty,InvCurrCls,InvPrice,InvAmount,InvCartonNo" & vbCrLf & _
                          "         ) " & vbCrLf & _
                          " VALUES  ( '" & pInvNo & "' , " & vbCrLf & _
                          "           '" & pSJno & "' ,  " & vbCrLf & _
                          "           '" & pAffiliateID & "' ,  " & vbCrLf

        ls_sql = ls_sql + "           '" & pPOno & "' ,  " & vbCrLf & _
                          "           '" & pPOKanbanCls & "' ,  " & vbCrLf & _
                          "           '" & pKanbanNo & "' ,  " & vbCrLf & _
                          "           '" & pPartNo & "' , " & vbCrLf & _
                          "           '" & pReceiveQty & "' , " & vbCrLf & _
                          "           '" & pInvQty & "' , " & vbCrLf & _
                          "           (SELECT CurrCls FROM MS_CurrCls WHERE DESCRIPTION = '" & pInvCurr & "') ,  " & vbCrLf & _
                          "           '" & pInvPrice & "' ,  " & vbCrLf & _
                          "           " & pInvAmount & ",  " & vbCrLf & _
                          "           '" & pCartonNo & "'" & vbCrLf & _
                          "         ) " & vbCrLf & _
                          " END " & vbCrLf

        Save_Detail = ls_sql
    End Function

    Private Sub up_Delete()
        Dim ls_SQL As String = ""

        Dim ls_Invno As String = Trim(txtPasiInvoiceNo.Text)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " DELETE dbo.InvoicePASI_Master WHERE InvoiceNo = '" & ls_Invno & "'" & vbCrLf & _
                     " DELETE dbo.InvoicePASI_Detail WHERE InvoiceNo = '" & ls_Invno & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)


            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()

            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_SaveAll(ByVal pPO As String, ByVal pKanban As String, ByVal pInvNo As String, ByVal pPsj As String)
        Dim ls_SQL As String = ""
        Dim ls_User As String = Trim(Session("UserID").ToString)

        Dim ls_InvNo As String = Trim(txtPasiInvoiceNo.Text)
        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)
        Dim ls_SJno As String = Trim(txtPasisuratjalanno.Text)
        Dim ls_InvDate As Date = Trim(txtInvoiceDate.Text)
        ls_InvDate = Format(CDate(ls_InvDate), "yyyy-MM-dd")
        Dim ls_DueDate As Date = Trim(dtDueDate.Text)
        Dim ls_PaymentTerm As String = Trim(txtPaymentTerm.Text)
        Dim ls_TotalAmount As String = Trim(txttotalamount.Text)
        Dim ls_Notes As String = Trim(MmNotes.Text)

        Dim ls_Container As String = txtContainerNo.Text
        Dim ls_PlaceDate As String = txtPlaceDate.Text
        Dim ls_ShippedPer As String = txtShipperPer.Text
        Dim ls_OnOrAbout As String = txtOnOrAboutCondition.Text
        Dim ls_DeliveryTerm As String = txtDeliveryTerm.Text
        Dim ls_From As String = txtFrom.Text
        Dim ls_To As String = txtTo.Text
        Dim ls_Via As String = txtVia.Text
        Dim ls_Freight As String = txtFreight.Text

        Dim ls_POno As String
        Dim ls_POkanbanCls As String
        Dim ls_KanbanNo As String
        Dim ls_PartNo As String
        Dim ls_ReceiveQty As Double
        Dim ls_InvQty As Double
        Dim ls_InvCurr As String
        Dim ls_InvPrice As Double
        Dim ls_InvAmount As Double
        Dim ls_InvCartonNo As String

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Replace(Replace(pInvNo, "'", ""), ",", "") = "" Then
                pInvNo = ""
            End If

            If Replace(Replace(pPsj, "'", ""), ",", "") = "" Then
                pPsj = ""
            End If

            'ls_SQL = "     SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colpono,colpasisj, colkanbanno, colpartno )), * FROM( " & vbCrLf & _
            '         "     SELECT    DISTINCT  " & vbCrLf & _
            '         "              colpono = POM.PONo ,  " & vbCrLf & _
            '         "              colpokanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '0' THEN 'NO'  " & vbCrLf & _
            '         "                                 ELSE 'YES'  " & vbCrLf & _
            '         "                            END ,  " & vbCrLf & _
            '         "              colkanbanno = ISNULL(KD.KanbanNo, '') ,  " & vbCrLf & _
            '         "              colpartno = POD.PartNo ,  " & vbCrLf & _
            '         "              colpartname = MP.PartName ,  " & vbCrLf & _
            '         "              coluom = UC.Description ,  " & vbCrLf & _
            '         "              colCls = UC.unitcls ,  " & vbCrLf

            'ls_SQL = ls_SQL + "              colQtyBox = ROUND(CONVERT(CHAR, ISNULL(MPM.QtyBox, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colpasideliveryqty = '' ,  " & vbCrLf & _
            '                  "              colAffRecQty = ROUND(CONVERT(CHAR, ISNULL(RAD.RecQty, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colInvoiceToAffQty = ROUND(CONVERT(CHAR, COALESCE(IPD.InvQty,  " & vbCrLf & _
            '                  "                                                                RAD.RecQty, 0), 0),  " & vbCrLf & _
            '                  "                                         0) ,  " & vbCrLf & _
            '                  "              coldelqtybox = '' ,  " & vbCrLf & _
            '                  "              colInvCurr = ISNULL(MC.Description, '') ,  " & vbCrLf & _
            '                  "              colInvPrice = ROUND(CONVERT(CHAR, ISNULL(MPR.Price, 0), 0), 0) ,  " & vbCrLf & _
            '                  "              colInvAmount = ROUND(CONVERT(CHAR, COALESCE(IPD.InvAmount,  " & vbCrLf & _
            '                  "                                                          ( RAD.RecQty  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                            * MPR.Price ), 0), 0),  " & vbCrLf & _
            '                  "                                   0) ,  " & vbCrLf & _
            '                  "              colcartonno = COALESCE(PCD_PASI.cartonNo,IPD.InvCartonNo),   " & vbCrLf & _
            '                  "              colpasisj = PDM.SuratJalanNo  " & vbCrLf & _
            '                  "     FROM     dbo.PO_Master POM  " & vbCrLf & _
            '                  "              LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
            '                  "                                         AND POM.PoNo = POD.PONo  " & vbCrLf & _
            '                  "                                         AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
            '                  "                                                AND KD.PoNo = POD.PONo  " & vbCrLf & _
            '                  "                                                AND KD.SupplierID = POD.SupplierID  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                AND KD.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID  " & vbCrLf & _
            '                  "                                                AND KD.KanbanNo = KM.KanbanNo  " & vbCrLf & _
            '                  "                                                AND KD.SupplierID = KM.SupplierID  " & vbCrLf & _
            '                  "                                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
            '                  "                                                     AND KD.KanbanNo = SDD.KanbanNo  " & vbCrLf & _
            '                  "                                                     AND KD.SupplierID = SDD.SupplierID  " & vbCrLf & _
            '                  "                                                     AND KD.PartNo = SDD.PartNo  " & vbCrLf & _
            '                  "                                                     AND KD.PoNo = SDD.PoNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                     AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                     AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf & _
            '                  "              INNER JOIN dbo.ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
            '                  "                                                       AND SDD.KanbanNo = PRD.KanbanNo  " & vbCrLf & _
            '                  "                                                       AND SDD.SupplierID = PRD.SupplierID  " & vbCrLf & _
            '                  "                                                       AND SDD.PartNo = PRD.PartNo  " & vbCrLf & _
            '                  "                                                       AND SDD.PONo = PRD.PONo  " & vbCrLf & _
            '                  "                                                       AND PRD.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
            '                  "                                                      AND PRM.SuratJalanNo = PRD.SuratJalanNo      " & vbCrLf & _
            '                  "              LEFT JOIN ( SELECT  SuratJalanno ,  " & vbCrLf

            'ls_SQL = ls_SQL + "                                  SupplierID ,  " & vbCrLf & _
            '                  "                                  AffiliateID ,  " & vbCrLf & _
            '                  "                                  PONO ,  " & vbCrLf & _
            '                  "                                  KanbanNO ,  " & vbCrLf & _
            '                  "                                  Partno ,  " & vbCrLf & _
            '                  "                                  UnitCls ,  " & vbCrLf & _
            '                  "                                  DoQty = SUM(ISNULL(DoQty, 0))  " & vbCrLf & _
            '                  "                          FROM    DOPasi_Detail  " & vbCrLf & _
            '                  "                          GROUP BY SuratJalanno ,  " & vbCrLf & _
            '                  "                                  SupplierID ,  " & vbCrLf & _
            '                  "                                  AffiliateID ,  " & vbCrLf

            'ls_SQL = ls_SQL + "                                  PONO ,  " & vbCrLf & _
            '                  "                                  KanbanNO ,  " & vbCrLf & _
            '                  "                                  Partno ,  " & vbCrLf & _
            '                  "                                  UnitCls  " & vbCrLf & _
            '                  "                        ) PDD ON PRD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
            '                  "                                 AND PRD.KanbanNo = PDD.KanbanNo  " & vbCrLf & _
            '                  "                                 AND PRD.SupplierID = PDD.SupplierID  " & vbCrLf & _
            '                  "                                 AND PRD.PartNo = PDD.PartNo  " & vbCrLf & _
            '                  "                                 AND PRD.PoNo = PDD.PoNo     " & vbCrLf & _
            '                  "              LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
            '                  "                                                 AND PDD.SuratJalanNo = PDM.SuratJalanNo     " & vbCrLf

            'ls_SQL = ls_SQL + "              INNER JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                                            AND PDD.KanbanNo = RAD.KanbanNo  " & vbCrLf & _
            '                  "                                                            AND PDD.SupplierID = RAD.SupplierID  " & vbCrLf & _
            '                  "                                                            AND PDD.PartNo = RAD.PartNo  " & vbCrLf & _
            '                  "                                                            AND PDD.PoNo = RAD.PoNo  " & vbCrLf & _
            '                  "              INNER JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo  " & vbCrLf & _
            '                  "                                                            AND RAM.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
            '                  "              LEFT JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
            '                  "                                                      AND RAD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
            '                  "                                                      AND RAD.PartNo = IPD.PartNo  " & vbCrLf & _
            '                  "                                                      AND RAD.PONo = IPD.PONo  " & vbCrLf

            'ls_SQL = ls_SQL + "                                                      AND RAD.SuratJalanNo = IPD.SuratJalanNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID  " & vbCrLf & _
            '                  "                                                      AND IPD.InvoiceNo = IPM.InvoiceNo  " & vbCrLf & _
            '                  "                                                      --AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
            '                  "              LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_Price MPR ON MPR.AffiliateID = RAD.AffiliateID  " & vbCrLf & _
            '                  "                                            AND MPR.PartNo = RAD.PartNo  " & vbCrLf

            'ls_SQL = ls_SQL + "                                            AND RAM.ReceiveDate BETWEEN MPR.StartDate  " & vbCrLf & _
            '                  "                                                                AND  " & vbCrLf & _
            '                  "                                                                MPR.EndDate  " & vbCrLf & _
            '                  "              LEFT JOIN dbo.MS_CurrCls MC ON MC.CurrCls = MPR.CurrCls  " & vbCrLf & _
            '                  "              LEFT JOIN PLPASI_Detail PCD_PASI ON PCD_PASI.SuratJalanNo = PDD.suratJalanNo   " & vbCrLf & _
            '                  "  											AND PCD_PASI.SupplierID = PDD.SupplierID   " & vbCrLf & _
            '                  "                           					AND PCD_PASI.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
            '                  "  											AND PCD_PASI.POnO = PDD.PONo   " & vbCrLf & _
            '                  "  				                            AND PCD_PASI.PartNo = PDD.PartNo   " & vbCrLf & _
            '                  "  				                            AND PCD_PASI.KanbanNo = PDD.KanbanNo " & vbCrLf

            'If pInvNo = "" Then
            '    ls_SQL = ls_SQL + "   WHERE     --POM.PONo IN (" & pPO & ") " & vbCrLf & _
            '                      "             --AND KD.KanbanNo IN (" & pKanban & ") " & vbCrLf & _
            '                      "             --AND " & vbCrLf & _
            '                      "             PDD.SuratJalanNo IN (" & pPsj & ") " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + "   WHERE     IPD.InvoiceNo ='" & pInvNo & "' " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " )x "
            ls_SQL = "  SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colpono,colpasisj, colkanbanno, colpartno )), * FROM(  " & vbCrLf & _
                      "  SELECT DISTINCT   " & vbCrLf & _
                      "     colpono = DD.PONo ,   " & vbCrLf & _
                      "     colpokanban = CASE WHEN ISNULL(DD.POKanbanCls, '0') = '0' THEN 'NO'   " & vbCrLf & _
                      "                         ELSE 'YES'   " & vbCrLf & _
                      "                 END ,   " & vbCrLf & _
                      "     colkanbanno = ISNULL(DD.KanbanNo, '') ,   " & vbCrLf & _
                      "     colpartno = DD.PartNo ,   " & vbCrLf & _
                      "     colpartname = MP.PartName ,   " & vbCrLf & _
                      "     coluom = UC.Description ,   " & vbCrLf & _
                      "     colCls = UC.unitcls ,   " & vbCrLf

            ls_SQL = ls_SQL + "     colQtyBox = ROUND(CONVERT(CHAR, ISNULL(MPM.QtyBox, 0), 0), 0) ,   " & vbCrLf & _
                              "     colpasideliveryqty = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0), 0), 0) ,   " & vbCrLf & _
                              "     colAffRecQty = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0), 0), 0) ,  " & vbCrLf & _
                              "     colInvoiceToAffQty = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0), 0), 0) ,  " & vbCrLf & _
                              "     coldelqtybox = ROUND(CONVERT(CHAR, ISNULL(DD.DoQty, 0) / ISNULL(MPM.QtyBox, 0), 0), 0) ,   " & vbCrLf & _
                              "     colInvCurr = ISNULL(MC.Description, '') ,   " & vbCrLf & _
                              "     colInvPrice = ROUND(CONVERT(CHAR, ISNULL(DD.Price,0), 0), 0) ,   " & vbCrLf & _
                              "     colInvAmount = ROUND(CONVERT(CHAR, COALESCE(IPD.InvAmount,   " & vbCrLf & _
                              "                                                 ( DD.DoQty   " & vbCrLf & _
                              "                                                 * ISNULL(DD.Price,0)), 0), 0),   " & vbCrLf & _
                              "                         0) ,   " & vbCrLf

            ls_SQL = ls_SQL + "     colcartonno = COALESCE(PCD_PASI.cartonNo,IPD.InvCartonNo),    " & vbCrLf & _
                              "     colpasisj = DM.SuratJalanNo, InvoiceNo = isnull(IPD.InvoiceNo,'')   " & vbCrLf & _
                              " FROM DOPASI_Master DM " & vbCrLf & _
                              " INNER JOIN DOPASI_Detail DD ON DM.SuratJalanNo = DD.SuratJalanNo and DM.AffiliateID = DD.AffiliateID " & vbCrLf & _
                              " LEFT JOIN InvoicePASI_Detail IPD ON DD.AffiliateID = IPD.AffiliateID  " & vbCrLf & _
                              " 									AND DD.KanbanNo = IPD.KanbanNo  " & vbCrLf & _
                              " 									AND DD.PartNo = IPD.PartNo  " & vbCrLf & _
                              " 									AND DD.PONo = IPD.PONo  " & vbCrLf & _
                              " 									AND DD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
                              " LEFT JOIN PLPASI_Detail PCD_PASI ON PCD_PASI.SuratJalanNo = DD.suratJalanNo    " & vbCrLf & _
                              "   									AND PCD_PASI.SupplierID = DD.SupplierID    " & vbCrLf

            ls_SQL = ls_SQL + "                            			AND PCD_PASI.AffiliateID = DD.AffiliateID    " & vbCrLf & _
                              "   									AND PCD_PASI.POnO = DD.PONo    " & vbCrLf & _
                              "   				                    AND PCD_PASI.PartNo = DD.PartNo    " & vbCrLf & _
                              "   				                    AND PCD_PASI.KanbanNo = DD.KanbanNo  " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON MP.PartNo = DD.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DD.PartNo AND MPM.AffiliateID = DD.AffiliateID AND MPM.SupplierID = DD.SupplierID " & vbCrLf & _
                              " LEFT JOIN MS_Price MPR on MPR.PartNo = DD.PartNo and MPR.AffiliateID = DD.AffiliateID and (DM.DeliveryDate between MPR.EffectiveDate and MPR.EndDate) " & vbCrLf & _
                              " LEFT JOIN MS_CurrCls MC ON MC.CurrCls = MPR.CurrCls " & vbCrLf

            ls_SQL = ls_SQL + " WHERE DM.SuratJalanNo IN (" & pPsj & ") " & vbCrLf & _
                              " )x " & vbCrLf & _
                              "  "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Using sqlScope As New TransactionScope
                    'Save Master
                    ls_SQL = Save_Master(ls_InvNo, ls_AffiliateID, ls_SJno, ls_InvDate, ls_DueDate, ls_PaymentTerm, ls_TotalAmount, ls_Notes, ls_User,
                                         ls_Container, ls_PlaceDate, ls_ShippedPer, ls_OnOrAbout, ls_DeliveryTerm, ls_From, ls_To, ls_Via, ls_Freight)
                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    Try
                        With ds.Tables(0)
                            For iLoop = 0 To .Rows.Count - 1

                                ls_POno = Trim(.Rows(iLoop).Item("colpono"))
                                ls_POkanbanCls = Trim(.Rows(iLoop).Item("colpokanban"))
                                If ls_POkanbanCls = "YES" Then ls_POkanbanCls = "1" Else ls_POkanbanCls = "0"
                                ls_KanbanNo = Trim(.Rows(iLoop).Item("colkanbanno"))
                                ls_PartNo = Trim(.Rows(iLoop).Item("colpartno"))
                                ls_ReceiveQty = Trim(.Rows(iLoop).Item("colAffRecQty"))
                                ls_InvQty = Trim(.Rows(iLoop).Item("colInvoiceToAffQty"))
                                ls_InvCurr = Trim(.Rows(iLoop).Item("colInvCurr"))
                                ls_InvPrice = Trim(.Rows(iLoop).Item("colInvPrice"))
                                ls_InvAmount = Trim(.Rows(iLoop).Item("colInvAmount"))
                                ls_SJno = Trim(.Rows(iLoop).Item("colpasisj"))
                                If IsDBNull(.Rows(iLoop).Item("colcartonno")) Then
                                    ls_InvCartonNo = ""
                                Else
                                    ls_InvCartonNo = Trim(.Rows(iLoop).Item("colcartonno"))
                                End If

                                'Save Detail
                                ls_SQL = Save_Detail(ls_InvNo, ls_SJno, ls_AffiliateID, ls_POno, ls_POkanbanCls, ls_KanbanNo, ls_PartNo, ls_ReceiveQty, ls_InvQty, ls_InvCurr, ls_InvPrice, ls_InvAmount, ls_InvCartonNo)
                                Dim sqlComm2 As New SqlCommand(ls_SQL, sqlConn)
                                sqlComm2.ExecuteNonQuery()
                                sqlComm2.Dispose()
                            Next
                            ''update master
                            'ls_SQL = Update_Master(ls_InvNo, ls_AffiliateID)
                            'Dim sqlComm3 As New SqlCommand(ls_SQL, sqlConn)
                            'sqlComm3.ExecuteNonQuery()
                            'sqlComm3.Dispose()

                            sqlScope.Complete()

                            'update master
                            ls_SQL = Update_Master(ls_InvNo, ls_AffiliateID)
                            Dim sqlComm3 As New SqlCommand(ls_SQL, sqlConn)
                            sqlComm3.ExecuteNonQuery()
                            sqlComm3.Dispose()

                        End With
                    Catch ex As Exception

                    End Try

                End Using
            End If
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_IsiMaster(ByVal pInvNo As String, ByVal pSuratJalanNo As String, ByVal pInvDate As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If pInvDate = "01 Jan 1900 " Or pInvoiceDate = "01/01/1900" Then
                ls_SQL = "SELECT InvoiceNo = Isnull(DO.InvoiceNo,''), ContainerNo = ISNULL(ContainerNo,''), PlaceDate = ISNULL(Place,''), ShippedPer = '', OnOrAboutCondition = ISNULL(Onabout,''), DeliveryTerm = '', InvFrom = ISNULL(FromDelivery,''), InvTo = ISNULL(toDelivery,''), InvVia = ISNULL(viadelivery,''), InvFreight = ''" & vbCrLf & _
                         " FROM DOPasi_Master DO LEFT JOIN PLPasi_Master PL ON PL.suratjalanno = DO.SuratJalanNo " & vbCrLf & _
                         " AND PL.AffiliateID = DO.AffiliateID " & vbCrLf & _
                         " WHERE DO.SuratJalanNo IN (" & Trim(pSuratJalanNo) & ") "
            Else                
                ls_SQL = " SELECT * FROM InvoicePASI_Master WHERE InvoiceNo = '" & pInvNo & "'"
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 And ds.Tables(0).Rows.Count = 1 Then
                Try
                    With ds.Tables(0)
                        txtPasiInvoiceNo.Text = Trim(.Rows(0).Item("InvoiceNo"))
                        txtContainerNo.Text = Trim(.Rows(0).Item("ContainerNo"))
                        txtPlaceDate.Text = Trim(.Rows(0).Item("PlaceDate"))
                        txtShipperPer.Text = Trim(.Rows(0).Item("ShippedPer"))
                        txtOnOrAboutCondition.Text = Trim(.Rows(0).Item("OnOrAboutCondition"))
                        txtDeliveryTerm.Text = Trim(.Rows(0).Item("DeliveryTerm"))
                        txtFrom.Text = Trim(.Rows(0).Item("InvFrom"))
                        txtTo.Text = Trim(.Rows(0).Item("InvTo"))
                        txtVia.Text = Trim(.Rows(0).Item("InvVia"))
                        txtFreight.Text = Trim(.Rows(0).Item("InvFreight"))
                    End With
                Catch ex As Exception

                End Try
            Else
                txtPasiInvoiceNo.Text = ""
                txtContainerNo.Text = ""
                txtPlaceDate.Text = ""
                txtShipperPer.Text = ""
                txtOnOrAboutCondition.Text = ""
                txtDeliveryTerm.Text = ""
                txtFrom.Text = ""
                txtTo.Text = ""
                txtVia.Text = ""
                txtFreight.Text = ""
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
                    Session.Remove("InvInvoice")
                    pInvoiceDate = Split(param, "|")(0)
                    pAffiliateCode = Split(param, "|")(1)
                    pAffiliateName = Split(param, "|")(2)
                    pPasiSj = Split(param, "|")(3)
                    pPasiInvoiceno = Split(param, "|")(4)
                    pPaymentTerm = Split(param, "|")(5)
                    pDueDate = Split(param, "|")(6)
                    pNotes = Split(param, "|")(7)

                    pPo = Split(param, "|")(8)
                    pKanban = Split(param, "|")(9)

                    If Session("InvInvoice") <> "" Then pPasiInvoiceno = Session("InvInvoice")
                    'If Session("POListInv") <> "" Then pKanban = Session("KanbanListInv")

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"
                    If Trim(pInvoiceDate) = "01 Jan 1900" Or Trim(pInvoiceDate) = "01/01/1900" Then txtInvoiceDate.Text = Format(Now, "dd MMM yyyy") Else txtInvoiceDate.Text = pInvoiceDate
                    'txtInvoiceDate.Text = pInvoiceDate
                    txtaffiliatecode.Text = pAffiliateCode
                    txtaffiliatename.Text = pAffiliateName
                    txtPasisuratjalanno.Text = Replace(pPasiSj, "'", "")


                    txttotalamount.Text = uf_SumAmount(pPasiInvoiceno, pPasiSj, pInvoiceDate)
                    dtDueDate.Text = Format(CDate(Now), "dd MMM yyyy")
                    'pStatus = True

                    If pPasiInvoiceno <> "" Then
                        txtPasiInvoiceNo.Text = pPasiInvoiceno
                        txtPaymentTerm.Text = pPaymentTerm
                        If pDueDate <> "" Then
                            dtDueDate.Text = Format(CDate(pDueDate), "dd MMM yyyy")
                        End If
                        MmNotes.Text = pNotes
                        'Call up_IsiMaster(pPasiInvoiceno, pPasiSj)
                    End If
                    'txttotalbox.Text = Format(pkanbandate, "dd MMM yyyy")
                    'paramDT1 = pdt1
                    'paramDT2 = pdt2
                    'paramaffiliate = pcboaffiliate
                    'paramSupplier = ptxtsupplierID
                    Call up_IsiMaster(pPasiInvoiceno, pPasiSj, pInvoiceDate)
                    Call up_GridLoad(pPo, pKanban, pPasiInvoiceno, pPasiSj)
                    Session("POInv") = pPo
                    Session("KanbanInv") = pKanban
                    Session("InvoiceNoInv") = pPasiInvoiceno
                    Session("PasiSJ") = pPasiSj
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

    Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
        Dim ls_MsgID As String = ""
        Dim ls_sql As String = ""
        Dim iRow As Integer = 0
        Dim ls_User As String = Trim(Session("UserID").ToString)


        Dim ls_InvNo As String = Trim(txtPasiInvoiceNo.Text)
        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)
        Dim ls_SJno As String = Trim(txtPasisuratjalanno.Text)
        Dim ls_InvDate As Date = Trim(txtInvoiceDate.Text)
        ls_InvDate = Format(CDate(ls_InvDate), "yyyy-MM-dd")
        Dim ls_DueDate As Date = Trim(dtDueDate.Text)
        Dim ls_PaymentTerm As String = Trim(txtPaymentTerm.Text)
        Dim ls_TotalAmount As String = Trim(txttotalamount.Text)
        Dim ls_Notes As String = Trim(MmNotes.Text)

        Dim ls_Container As String = txtContainerNo.Text
        Dim ls_PlaceDate As String = txtPlaceDate.Text
        Dim ls_ShippedPer As String = txtShipperPer.Text
        Dim ls_OnOrAbout As String = txtOnOrAboutCondition.Text
        Dim ls_DeliveryTerm As String = txtDeliveryTerm.Text
        Dim ls_From As String = txtFrom.Text
        Dim ls_To As String = txtTo.Text
        Dim ls_Via As String = txtVia.Text
        Dim ls_Freight As String = txtFreight.Text

        Dim ls_POno As String
        Dim ls_POkanbanCls As String
        Dim ls_KanbanNo As String
        Dim ls_PartNo As String
        Dim ls_ReceiveQty As Double
        Dim ls_InvQty As Double
        Dim ls_InvCurr As String
        Dim ls_InvPrice As Double
        Dim ls_InvAmount As Double
        Dim ls_InvCartonNo As String



        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Grid.VisibleRowCount = 0 Then Exit Sub

            Using sqlTran As New TransactionScope

                'Save Master
                ls_sql = Save_Master(ls_InvNo, ls_AffiliateID, ls_SJno, ls_InvDate, ls_DueDate, ls_PaymentTerm, ls_TotalAmount, ls_Notes, ls_User,
                                     ls_Container, ls_PlaceDate, ls_ShippedPer, ls_OnOrAbout, ls_DeliveryTerm, ls_From, ls_To, ls_Via, ls_Freight)
                Dim sqlComm As New SqlCommand(ls_sql, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()

                For iRow = 0 To e.UpdateValues.Count - 1
                    ls_POno = e.UpdateValues(iRow).NewValues("colpono").ToString()
                    ls_POkanbanCls = e.UpdateValues(iRow).NewValues("colpokanban").ToString()
                    If ls_POkanbanCls = "YES" Then ls_POkanbanCls = "1" Else ls_POkanbanCls = "0"
                    ls_KanbanNo = e.UpdateValues(iRow).NewValues("colkanbanno").ToString()
                    ls_PartNo = e.UpdateValues(iRow).NewValues("colpartno").ToString()
                    ls_ReceiveQty = e.UpdateValues(iRow).NewValues("colAffRecQty").ToString()
                    ls_InvQty = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colInvoiceToAffQty").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvoiceToAffQty").ToString(), 0)))
                    ls_InvCurr = Trim((IIf(e.UpdateValues(iRow).NewValues("colInvCurr").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvCurr").ToString(), 0)))
                    ls_InvPrice = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colInvPrice").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvPrice").ToString(), 0)))
                    ls_InvAmount = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colInvAmount").ToString() <> "", e.UpdateValues(iRow).NewValues("colInvAmount").ToString(), 0)))
                    ls_InvCartonNo = Trim((IIf(e.UpdateValues(iRow).NewValues("colcartonno").ToString() <> "", e.UpdateValues(iRow).NewValues("colcartonno").ToString(), "")))

                    'Save Detail
                    ls_sql = Save_Detail(ls_InvNo, ls_SJno, ls_AffiliateID, ls_POno, ls_POkanbanCls, ls_KanbanNo, ls_PartNo, ls_ReceiveQty, ls_InvQty, ls_InvCurr, ls_InvPrice, ls_InvAmount, ls_InvCartonNo)

                    Dim sqlComm2 As New SqlCommand(ls_sql, sqlConn)
                    sqlComm2.ExecuteNonQuery()
                    sqlComm2.Dispose()
                Next iRow

                'update master
                ls_sql = Update_Master(ls_InvNo, ls_AffiliateID)
                Dim sqlComm3 As New SqlCommand(ls_sql, sqlConn)
                sqlComm3.ExecuteNonQuery()
                sqlComm3.Dispose()

                sqlTran.Complete()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Select Case pAction
            Case "gridload"
                Call up_SaveAll(Session("POInv"), Session("KanbanInv"), Session("InvoiceNoInv"), Session("PasiSJ"))
                Call up_GridLoad(Session("POInv"), Session("KanbanInv"), Session("InvoiceNoInv"), Session("PasiSJ"))

                Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text

            Case "Delete"
                Call up_Delete()
                Call up_GridLoad(Session("POInv"), Session("KanbanInv"), Session("InvoiceNoInv"), Session("PasiSJ"))

                Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Case "EDI"
                Dim pInv As String = Split(e.Parameters, "|")(1)
                Dim pAff As String = Split(e.Parameters, "|")(2)
                Call SendEDIFile(pInv, pAff)
                'Call SendEDIFile()
                'Call clsMsg.DisplayMessage(lblerrmessage, "2005", clsMessage.MsgType.InformationMessage)
                'Grid.JSProperties("cpMessage") = lblerrmessage.Text

        End Select
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
    End Sub

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "colInvoiceToAffQty" Or e.DataColumn.FieldName = "colcartonno") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        Else
            e.Cell.BackColor = Color.White
        End If

        'Delivery Qty Not save
        If e.DataColumn.FieldName = "colInvoiceToAffQty" Then
            If (Trim(e.GetValue("InvoiceNo")) = "") Then
                e.Cell.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub Grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles Grid.HtmlRowPrepared
        Try
            Dim getRowValues As String = e.GetValue("colInvoiceToAffQty")
            If Not IsNothing(getRowValues) Then
                If getRowValues.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If
            Dim getRowValues2 As String = e.GetValue("colcartonno")
            If Not IsNothing(getRowValues2) Then
                If getRowValues2.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Session.Remove("POInv")
        Session.Remove("KanbanInv")
        Session.Remove("InvoiceNoInv")
        Session.Remove("TampungInv")

        If btnsubmenu.Text = "BACK" Then
            Response.Redirect("~/Invoice/AffReceivingConf.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Session("InvInvoice") = txtPasiInvoiceNo.Text
        Response.Redirect("~/Invoice/viewInvToAff.aspx")

    End Sub

    Private Sub SendEDIFile(ByVal pInv As String, ByVal pAff As String)
        'Private Sub SendEDIFile()
        Dim fp As StreamWriter
        Dim ls_sql As String

        ls_sql = "    SELECT    *  " & vbCrLf & _
                  "    FROM      ( SELECT DISTINCT  " & vbCrLf & _
                  "                          a = 'H00' + CONVERT(CHAR(8), 'VD01')  " & vbCrLf & _
                  "                          + CONVERT(CHAR(8), '32G8')  " & vbCrLf & _
                  "                          + CASE WHEN RTRIM(IVM.AffiliateID) LIKE 'PEMI%'  " & vbCrLf & _
                  "                                 THEN CONVERT(CHAR(8), '32M8')  " & vbCrLf & _
                  "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI T%'  " & vbCrLf & _
                  "                                 THEN CONVERT(CHAR(8), '32M3')  " & vbCrLf & _
                  "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI JF%'  " & vbCrLf & _
                  "                                 THEN CONVERT(CHAR(8), '32CH')  " & vbCrLf & _
                  "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SUAI%'  " & vbCrLf & _
                  "                                 THEN CONVERT(CHAR(8), '32G2')  " & vbCrLf & _
                  "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAI%'  " & vbCrLf & _
                  "                                 THEN CONVERT(CHAR(8), '32M4')  " & vbCrLf & _
                  "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'JAI%'  " & vbCrLf

        ls_sql = ls_sql + "                                 THEN CONVERT(CHAR(8), '32G7')  " & vbCrLf & _
                          "                            END + CONVERT(CHAR(8), CONVERT(DATETIME, GETDATE()), 112)  " & vbCrLf & _
                          "                          + REPLACE(CONVERT (VARCHAR(8), GETDATE(), 108), ':',  " & vbCrLf & _
                          "                                    '') + CONVERT(CHAR(15), 'INVOICE-DATA')  " & vbCrLf & _
                          "                          + CONVERT(CHAR(19), '') ,  " & vbCrLf & _
                          "                          INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                          AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                          PARTNO = '' ,  " & vbCrLf & _
                          "                          KANBANNO = '' ,  " & vbCrLf & _
                          "                          IDX = 1  " & vbCrLf & _
                          "                FROM      InvoicePasi_Master IVM  " & vbCrLf

        ls_sql = ls_sql + "                          LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                              AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                              AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                UNION ALL  " & vbCrLf & _
                          "                SELECT DISTINCT  " & vbCrLf & _
                          "                          a = 'H10' + '000000' + 'T' + CONVERT(CHAR(35), 'TRUCK')  " & vbCrLf & _
                          "                          + CONVERT(CHAR(10), '')  " & vbCrLf & _
                          "                          + CONVERT(CHAR(15), IVM.ContainerNo) + '2'  " & vbCrLf & _
                          "                          + CONVERT(CHAR(4), '') ,  " & vbCrLf & _
                          "                          INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                          AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf

        ls_sql = ls_sql + "                          PARTNO = '' ,  " & vbCrLf & _
                          "                          KANBANNO = '' ,  " & vbCrLf & _
                          "                          IDX = 2  " & vbCrLf & _
                          "                FROM      InvoicePasi_Master IVM  " & vbCrLf & _
                          "                          LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                              AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                              AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                          LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                                                                AND RAD.KanbanNo = IVD.Kanbanno  " & vbCrLf & _
                          "                                                                AND RAD.PONO = IVD.PONO  " & vbCrLf & _
                          "                          LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID  "

        ls_sql = ls_sql + "                                                                AND RAD.SuratJalanNo = RAM.SuratJalanNo  " & vbCrLf & _
                          "                          LEFT JOIN PLPasi_Master PLM ON PLM.SuratJalanNo = IVM.SuratJalanNo  " & vbCrLf & _
                          "                                                         AND PLM.AffiliateID = IVM.AffiliateID  " & vbCrLf & _
                          "                                                         AND PLM.InvoiceNo = IVM.InvoiceNo  " & vbCrLf & _
                          "                          LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                          LEFT JOIN PO_Master PM ON PM.pono = IVD.PONO  " & vbCrLf & _
                          "                                                    AND PM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                UNION ALL  " & vbCrLf & _
                          "                SELECT DISTINCT  " & vbCrLf & _
                          "                          a = 'H20'  " & vbCrLf & _
                          "                          + CONVERT(CHAR(8), CONVERT(DATETIME, DPM.DeliveryDate), 112)  " & vbCrLf

        ls_sql = ls_sql + "                          + CONVERT(CHAR(8), CONVERT(DATETIME, DPM.DeliveryDate), 112)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(8), CONVERT(DATETIME, DPM.DeliveryDate), 112)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(8), CONVERT(DATETIME, DPM.DeliveryDate), 112)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(20), IVM.InvFrom)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(20), IVM.InvTo) ,  " & vbCrLf & _
                          "                          INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                          AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                          PARTNO = '' ,  " & vbCrLf & _
                          "                          KANBANNO = '' ,  " & vbCrLf & _
                          "                          IDX = 3  " & vbCrLf & _
                          "                FROM      InvoicePasi_Master IVM  " & vbCrLf

        ls_sql = ls_sql + "                          LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                              AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                              AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                          LEFT JOIN DOPASI_Master DPM On DPM.AffiliateID = IVM.AffiliateID AND IVM.SuratJalanNo = DPM.SuratJalanNo AND IVM.InvoiceNo = DPM.InvoiceNo" & vbCrLf

        ls_sql = ls_sql + "                UNION ALL  " & vbCrLf & _
                          "                SELECT DISTINCT  " & vbCrLf & _
                          "                          a = 'H30' + CONVERT(CHAR(8), '')  " & vbCrLf & _
                          "                          + CONVERT(CHAR(15), IVM.InvoiceNo)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(15), '')  " & vbCrLf & _
                          "                          + CONVERT(CHAR(8), CONVERT(DATETIME, IVM.InvoiceDate), 112)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(8), CONVERT(DATETIME, IVM.InvoiceDate), 112)  " & vbCrLf & _
                          "                          + CONVERT(CHAR(5), 'FOB') + 'C' + 'C'  " & vbCrLf & _
                          "                          + CONVERT(CHAR(4), YEAR(GETDATE())) + '0' + CONVERT(CHAR(6), '') ,  " & vbCrLf & _
                          "                          INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                          AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf

        ls_sql = ls_sql + "                          PARTNO = '' ,  " & vbCrLf & _
                          "                          KANBANNO = '' ,  " & vbCrLf & _
                          "                          IDX = 4  " & vbCrLf & _
                          "                FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                UNION ALL  " & vbCrLf & _
                          "                SELECT DISTINCT  " & vbCrLf & _
                          "                          a = 'H40' + RIGHT(RTRIM('000000000000000'  " & vbCrLf & _
                          "                                                  + CONVERT(CHAR(15), REPLACE(CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                5), IVM.TotalAmount),'.',''))),  " & vbCrLf & _
                          "                                            15) + CONVERT(CHAR(6), '000000')  " & vbCrLf & _
                          "                          + CASE WHEN RTRIM(IVM.AffiliateID) LIKE 'PEMI%'  " & vbCrLf

        ls_sql = ls_sql + "                                 THEN CONVERT(CHAR(8), '32M8')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI T%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M3')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI JF%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32CH')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SUAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32G2')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M4')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'JAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32G7')  " & vbCrLf & _
                          "                            END  " & vbCrLf & _
                          "                          + CASE WHEN RTRIM(IVM.AffiliateID) LIKE 'PEMI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M8')  " & vbCrLf 

        ls_sql = ls_sql + "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI T%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M3')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI JF%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32CH')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SUAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32G2')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M4')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'JAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32G7')  " & vbCrLf & _
                          "                            END  " & vbCrLf & _
                          "                          + CASE WHEN RTRIM(IVM.AffiliateID) LIKE 'PEMI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M8')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI T%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32M3')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI JF%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32CH')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SUAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32G2')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'SAI%'  " & vbCrLf

        ls_sql = ls_sql + "                                 THEN CONVERT(CHAR(8), '32M4')  " & vbCrLf & _
                          "                                 WHEN RTRIM(IVM.AffiliateID) LIKE 'JAI%'  " & vbCrLf & _
                          "                                 THEN CONVERT(CHAR(8), '32G7')  " & vbCrLf & _
                          "                            END + RIGHT(RTRIM('000000000'  " & vbCrLf & _
                          "                                              + CONVERT(CHAR(9), replace(CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                2), SUM(( InvQty  " & vbCrLf & _
                          "                                                                / MOQ )  " & vbCrLf & _
                          "                                                                * ( MPM.Netweight  " & vbCrLf & _
                          "                                                                / 1000 ))),'.',''))), 9)  " & vbCrLf & _
                          "                          + RIGHT(RTRIM('000000000'  " & vbCrLf & _
                          "                                        + CONVERT(CHAR(9), replace(CONVERT(NUMERIC(32, 2), SUM(( InvQty  " & vbCrLf

        ls_sql = ls_sql + "                                                                / MOQ )  " & vbCrLf & _
                          "                                                                * ( MPM.Grossweight  " & vbCrLf & _
                          "                                                                / 1000 ))),'.',''))), 9)  " & vbCrLf & _
                          "                          + ' ' + '0000' + '0000' ,  " & vbCrLf & _
                          "                          INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                          AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                          PARTNO = '' ,  " & vbCrLf & _
                          "                          KANBANNO = '' ,  " & vbCrLf & _
                          "                          IDX = 5  " & vbCrLf & _
                          "                FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                          LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf

        ls_sql = ls_sql + "                                                              AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                              AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                          LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                GROUP BY  IVM.INVOICENO ,  " & vbCrLf & _
                          "                          IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                          IVM.TotalAmount ,  " & vbCrLf & _
                          "                          IVM.affiliateID    " & vbCrLf & _
                          "              --UNION ALL    " & vbCrLf & _
                          "              --SELECT DISTINCT    " & vbCrLf & _
                          "              --          a = 'H41' + CONVERT(CHAR(15), '') + CONVERT(CHAR(20), '')    " & vbCrLf & _
                          "              --          + CONVERT(CHAR(8), '') + CONVERT(CHAR(29), '') ,              --          INVOICENO = IVM.INVOICENO ,    " & vbCrLf

        ls_sql = ls_sql + "              --          AFFILIATEID = IVM.AFFILIATEID    " & vbCrLf & _
                          "              --FROM      INVOICEPASI_MASTER IVM      " & vbCrLf & _
                          "                            -------------- DETAIL ------------------      " & vbCrLf & _
                          "                UNION ALL  " & vbCrLf & _
                          "                SELECT    *  " & vbCrLf & _
                          "                FROM      ( SELECT    a = 'D10' + CONVERT(CHAR(25), IVD.PartNo)  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(3), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(2), '')  " & vbCrLf & _
                          "                                      + RIGHT(RTRIM('000000000000000'  " & vbCrLf & _
                          "                                                    + CONVERT(CHAR(15), REPLACE(CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                5), IVD.InvPrice),'.',''))),  " & vbCrLf

        ls_sql = ls_sql + "                                              15)  " & vbCrLf & _
                          "                                      + CASE WHEN RTRIM(IVM.AffiliateID) LIKE 'PEMI%'  " & vbCrLf & _
                          "                                             THEN 'IDR'  " & vbCrLf & _
                          "                                             WHEN RTRIM(IVM.AffiliateID) LIKE 'SAMI%'  " & vbCrLf & _
                          "                                             THEN 'IDR'  " & vbCrLf & _
                          "                                             WHEN RTRIM(IVM.AffiliateID) LIKE 'SAI%'  " & vbCrLf & _
                          "                                             THEN 'IDR'  " & vbCrLf & _
                          "                                             WHEN RTRIM(IVM.AffiliateID) LIKE 'SUAI%'  " & vbCrLf & _
                          "                                             THEN 'IDR'  " & vbCrLf & _
                          "                                             WHEN RTRIM(IVM.AffiliateID) LIKE 'JAI%'  " & vbCrLf & _
                          "                                             THEN 'IDR'  " & vbCrLf & _
                          "                                        END  " & vbCrLf & _
                          "                                      + UPPER(CONVERT(char(3), MU.Description)) + RIGHT(RTRIM('000000000'  "

        ls_sql = ls_sql + "                                                          + CONVERT(CHAR(9), REPLACE(CONVERT(NUMERIC(9,  " & vbCrLf & _
                          "                                                                0), MPM.QtyBox),'.',''))),  " & vbCrLf & _
                          "                                                    9) + CONVERT(CHAR(1), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(11), '') ,  " & vbCrLf & _
                          "                                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                                      AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                                      PARTNO = IVD.PartNo ,  " & vbCrLf & _
                          "                                      KANBANNO = IVD.KANBANNO ,  " & vbCrLf & _
                          "                                      IDX = 6  " & vbCrLf & _
                          "                            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf

        ls_sql = ls_sql + "                                                                AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                                AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                                      LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
                          "                                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
                          "                            UNION ALL  " & vbCrLf & _
                          "                            SELECT    a = 'D11' + CONVERT(CHAR(30), MP.PartName)  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(25), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(3), 'EA')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(4), 'PLT')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(10), '') ,  " & vbCrLf

        ls_sql = ls_sql + "                                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                                      AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                                      PARTNO = IVD.PartNo ,  " & vbCrLf & _
                          "                                      KANBANNO = IVD.KANBANNO ,  " & vbCrLf & _
                          "                                      IDX = 7  " & vbCrLf & _
                          "                            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                                AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                                AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                                      LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf

        ls_sql = ls_sql + "                                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
                          "                            UNION ALL  " & vbCrLf & _
                          "                            SELECT    a = 'D12' + CONVERT(CHAR(3), '360')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(9), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(25), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(35), '') ,  " & vbCrLf & _
                          "                                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                                      AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                                      PARTNO = IVD.PartNo ,  " & vbCrLf & _
                          "                                      KANBANNO = IVD.KANBANNO ,  " & vbCrLf & _
                          "                                      IDX = 8  " & vbCrLf

        ls_sql = ls_sql + "                            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                                AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                                AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                                      LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
                          "                                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
                          "                            UNION ALL  " & vbCrLf & _
                          "                            SELECT    a = 'D20' + CONVERT(CHAR(15), IVD.PONo)  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(1), '')  " & vbCrLf & _
                          "                                      + RIGHT(RTRIM('000000000'  " & vbCrLf

        ls_sql = ls_sql + "                                                    + CONVERT(CHAR(9), CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                0), IVD.InvQty))),  " & vbCrLf & _
                          "                                              9) + RIGHT(RTRIM('000000000'  " & vbCrLf & _
                          "                                                               + CONVERT(CHAR(9), CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                0), IVD.InvQty))),  " & vbCrLf & _
                          "                                                         9)  " & vbCrLf & _
                          "                                      + RIGHT(RTRIM('000000000000000'  " & vbCrLf & _
                          "                                                    + CONVERT(CHAR(15), REPLACE(CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                5), IVD.InvAmount),'.',''))),  " & vbCrLf & _
                          "                                              15) + CONVERT(CHAR(23), '') ,  " & vbCrLf & _
                          "                                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf

        ls_sql = ls_sql + "                                      AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                                      PARTNO = IVD.PartNo ,  " & vbCrLf & _
                          "                                      KANBANNO = IVD.KANBANNO ,  " & vbCrLf & _
                          "                                      IDX = 9  " & vbCrLf & _
                          "                            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                                AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                                AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          "                                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                                      LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
                          "                                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf

        ls_sql = ls_sql + "                            UNION ALL  " & vbCrLf & _
                          "                            SELECT    a = 'D30'  " & vbCrLf & _
                          "                                      + CASE WHEN CHARINDEX('-',  " & vbCrLf & _
                          "                                                            RTRIM(InvCartonNo),  " & vbCrLf & _
                          "                                                            1) <> 0  " & vbCrLf & _
                          "                                             THEN CONVERT(CHAR(8), RIGHT('00000000' + REPLACE(SUBSTRING(RTRIM(InvCartonNo),  " & vbCrLf & _
                          "                                                                1,  " & vbCrLf & _
                          "                                                                CHARINDEX('-',  " & vbCrLf & _
                          "                                                                RTRIM(InvCartonNo),  " & vbCrLf & _
                          "                                                                1) - 1),LEFT(InvCartonNo,1),''),8))  " & vbCrLf & _
                          "                                             ELSE RIGHT(RTRIM('00000000' + CONVERT(CHAR(8), REPLACE(InvCartonNo,LEFT(InvCartonNo,1),''))), 8) " & vbCrLf

        ls_sql = ls_sql + "                                        END  " & vbCrLf & _
                          "                                      + CASE WHEN CHARINDEX('-',  " & vbCrLf & _
                          "                                                            RTRIM(InvCartonNo),  " & vbCrLf & _
                          "                                                            1) <> 0  " & vbCrLf & _
                          "                                             THEN CONVERT(CHAR(8), RIGHT('00000000' + REPLACE(SUBSTRING(RTRIM(InvCartonNo),  " & vbCrLf & _
                          "                                                                CHARINDEX('-',  " & vbCrLf & _
                          "                                                                RTRIM(InvCartonNo),  " & vbCrLf & _
                          "                                                                1) + 1,  " & vbCrLf & _
                          "                                                                LEN(RTRIM(InvCartonNo))  " & vbCrLf & _
                          "                                                                + 1), LEFT(InvCartonNo,1),''),8))  " & vbCrLf & _
                          "                                             ELSE RIGHT(RTRIM('00000000' + CONVERT(CHAR(8), REPLACE(InvCartonNo,'C',''))), 8)  " & vbCrLf

        ls_sql = ls_sql + "                                        END + CONVERT(CHAR(5), LEFT(InvCartonNo,1))   " & vbCrLf & _
                          "                                      + RIGHT(RTRIM('00000000'  " & vbCrLf & _
                          "                                                    + CONVERT(CHAR(8), REPLACE(CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                3), ( ISNULL(MPM.Netweight,  " & vbCrLf & _
                          "                                                                0) ) / 1000),'.',''))),  " & vbCrLf & _
                          "                                              8) + RIGHT(RTRIM('00000000'  " & vbCrLf & _
                          "                                                               + CONVERT(CHAR(8), REPLACE(CONVERT(NUMERIC(32,  " & vbCrLf & _
                          "                                                                3), ( ISNULL(MPM.Grossweight,  " & vbCrLf & _
                          "                                                                0) ) / 1000),'.',''))),  " & vbCrLf & _
                          "                                                         8)  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(3), UPPER(MU.Description))  " & vbCrLf

        ls_sql = ls_sql + "                                      --+ RIGHT(RTRIM('00000000'  " & vbCrLf & _
                          "                                                    --+ CONVERT(CHAR(8), LEFT(( ( ISNULL(MPM.Width,  " & vbCrLf & _
                          "                                                               -- 0)  " & vbCrLf & _
                          "                                                                --* ISNULL(MPM.Length,  " & vbCrLf & _
                          "                                                                --0)  " & vbCrLf & _
                          "                                                                --* ISNULL(MPM.Height,  " & vbCrLf & _
                          "                                                               -- 0) )  " & vbCrLf & _
                          "                                                               -- * ISNULL(IVD.InvQty,  " & vbCrLf & _
                          "                                                               -- 0) ), 8))), 8)  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(8), '00000000') " & vbCrLf & _
                          "                                      + CONVERT(CHAR(3), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(8), '00000000')  " & vbCrLf

        ls_sql = ls_sql + "                                      + CONVERT(CHAR(8), '00000000')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(5), '') ,  " & vbCrLf & _
                          "                                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                                      AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                                      PARTNO = IVD.PartNo ,  " & vbCrLf & _
                          "                                      KANBANNO = IVD.KANBANNO ,  " & vbCrLf & _
                          "                                      IDX = 10  " & vbCrLf & _
                          "                            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                                AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                                AND IVM.AffiliateID = IVD.AffiliateID  "

        ls_sql = ls_sql + "                                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
            "                                                    LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
                          "                                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
                          "                            UNION ALL  " & vbCrLf & _
                          "                            SELECT    a = 'D31' + CONVERT(CHAR(10), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(10), '')  " & vbCrLf & _
                          "                                      + RIGHT(RTRIM('0000000' + CONVERT(CHAR(7), REPLACE(Convert(numeric(7,0),PLD.CartonQty),'.00',''))), 7)" & vbCrLf & _
                          "                                      + CONVERT(CHAR(10), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(1), '0')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(15), '')  " & vbCrLf & _
                          "                                      + CONVERT(CHAR(19), '') ,  "

        ls_sql = ls_sql + "                                      INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                                      AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                                      PARTNO = IVD.PartNo ,  " & vbCrLf & _
                          "                                      KANBANNO = IVD.KANBANNO ,  " & vbCrLf & _
                          "                                      IDX = 11  " & vbCrLf & _
                          "                            FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "                                      LEFT JOIN InvoicePasi_Detail IVD ON IVM.InvoiceNo = IVD.InvoiceNo  " & vbCrLf & _
                          "                                                                AND IVM.SuratJalanNo = IVD.SuratJalanNo  " & vbCrLf & _
                          "                                                                AND IVM.AffiliateID = IVD.AffiliateID  " & vbCrLf & _
                          " 									 LEFT JOIN PLPasi_Master PLM ON PLM.SuratJalanNo = IVM.SuratJalanNo  " & vbCrLf & _
                          "                                                         AND PLM.AffiliateID = IVM.AffiliateID  "

        ls_sql = ls_sql + "                                                         AND PLM.InvoiceNo = IVM.InvoiceNo " & vbCrLf & _
                          " 									 LEFT JOIN PLPasi_Detail PLD ON PLD.SuratJalanno = PLM.SuratJalanno " & vbCrLf & _
                          " 														AND PLD.AffiliateID = PLM.AffiliateID " & vbCrLf & _
                          " 														AND PLD.PartNo = IVD.PartNo " & vbCrLf & _
                          " 														AND PLD.PONo = IVD.PONo " & vbCrLf & _
                          " 														AND PLD.KanbanNo = IVD.KanbanNo  " & vbCrLf & _
                          "                                      LEFT JOIN MS_Parts MP ON MP.PartNo = IVD.PartNo  " & vbCrLf & _
                          "                                      LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = IVD.PartNo AND MPM.AffiliateID = IVD.AffiliateID AND MPM.SupplierID = IVD.SupplierID " & vbCrLf & _
                          "                                      LEFT JOIN MS_CurrCls MC ON mC.CurrCls = Ivd.InvCurrCls  " & vbCrLf & _
                          "                                      LEFT JOIN MS_UnitCls MU ON MU.Unitcls = MP.UnitCls  " & vbCrLf & _
                          "                          ) x    " & vbCrLf & _
                          "    		      " & vbCrLf

        ls_sql = ls_sql + "              -------------- FOOTER --------------------      " & vbCrLf & _
                          "                UNION ALL  " & vbCrLf & _
                          "                SELECT DISTINCT  " & vbCrLf & _
                          "                          a = 'T00' + CONVERT(CHAR(72), '') ,  " & vbCrLf & _
                          "                          INVOICENO = IVM.INVOICENO ,  " & vbCrLf & _
                          "                          AFFILIATEID = IVM.AFFILIATEID ,  " & vbCrLf & _
                          "                          PARTNO = 'ZZZZZZZZZZZZZZZZZZZZ' ,  " & vbCrLf & _
                          "                          KANBANNO = 'ZZZZZZZZZZZZZZZZZZZZ' ,  " & vbCrLf & _
                          "                          IDX = 12  " & vbCrLf & _
                          "                FROM      INVOICEPASI_MASTER IVM  " & vbCrLf & _
                          "              ) xx  " & vbCrLf & _
                          " WHERE   InvoiceNo = '" & pInv & "' " & vbCrLf & _
                          "         AND AffiliateID = '" & pAff & "' " & vbCrLf & _
                          " ORDER BY PARTNO, KANBANNO, IDX "



        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()
            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            'get pono
            ls_sql = " SELECT distinct PONO FROM INVOICEPASI_DETAIL WHERE invoiceNo = '" & Trim(pInv) & "' AND affiliateID = '" & Trim(pAff) & "'"
            Dim sqlD As New SqlDataAdapter(ls_sql, cn)
            Dim ds1 As New DataSet
            sqlD.Fill(ds1)
            'get pono

            Dim namaFile As String = ""
            Dim ls_aff As String = ""

            If ds.Tables(0).Rows.Count > 0 Then
                If Trim(pAff) = "PEMI" Then ls_aff = "32M8"
                If Trim(pAff) = "SAMI T" Then ls_aff = "32M3"
                If Trim(pAff) = "SAMI JF" Then ls_aff = "32CH"
                If Trim(pAff) = "SUAI" Then ls_aff = "32G2"
                If Trim(pAff) = "SAI B" Then ls_aff = "32M4"
                If Trim(pAff) = "SAI T" Then ls_aff = "32M4"
                If Trim(pAff) = "JAI" Then ls_aff = "32G7"

                'Dim fi As New FileInfo(Server.MapPath("~\Invoice\" & txtaffiliatecode.Text & "INV.txt"))
                'If fi.Exists Then
                '    fi.Delete()
                '    fi = New FileInfo(Server.MapPath("~\Invoice\" & txtaffiliatecode.Text & "INV.txt"))
                'End If

                'DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

                namaFile = "VD01.32G8." + Trim(ls_aff) + "." + Trim(ds1.Tables(0).Rows(0)("PONO")) + "." + Format(Now, "yyyyMMdd") + "." + Format(Now, "hhmm") + "." + Format(Now, "ss") + ".txt"

                fp = File.CreateText(Server.MapPath("~\Invoice\Result\" & namaFile))

                For x = 0 To ds.Tables(0).Rows.Count - 1
                    fp.WriteLine(ds.Tables(0).Rows(x)("a") & Format(x + 1, "00000"))
                Next

                fp.Close()
                'fi.Delete()

            End If
            Dim filePath As String = "~\Invoice\Result\" & namaFile
            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Invoice\Result\" & namaFile & "")
            'Response.TransmitFile(filePath)
        End Using

        Exit Sub
ErrHandler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub
End Class