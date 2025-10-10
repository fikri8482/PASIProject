Imports System.Data.SqlClient

Public Class viewDeliveryToAff
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Report As New rptDeliveryToAff
        Dim ds As New DataSet
        'crystalReport.Load(Server.MapPath("~/Report/rptPRReportOthers.vb"))
        'crystalReport.SetDatabaseLogon(clsBudget.gs_UserDB, clsBudget.gs_PasswordDB)
        Dim ls_SQL As String

        ''ls_SQL = " SELECT  Affiliate = MA.AffiliateName," & vbCrLf & _
        ''                  "         Address = MA.Address, " & vbCrLf & _
        ''                  "         colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY KD.KanbanNo, POD.PartNo )) , " & vbCrLf & _
        ''                  "         colSJ = PDM.SuratjalanNo , " & vbCrLf & _
        ''                  "         colDate = PDM.DeliveryDate , " & vbCrLf & _
        ''                  "         colpono = POM.PONo , " & vbCrLf & _
        ''                  "         colpartno = POD.PartNo , " & vbCrLf & _
        ''                  "         colpartname = MP.PartName , " & vbCrLf & _
        ''                  "         colpasideliveryqty = ISNULL(PDD.DOQty, 0) , " & vbCrLf & _
        ''                  "         coldelqtybox = CEILING(CASE MPM.QtyBox " & vbCrLf & _
        ''                  "                          WHEN 0 THEN 0 " & vbCrLf & _
        ''                  "                          ELSE ISNULL(SDD.DOQty, 0) / MPM.QtyBox " & vbCrLf & _
        ''                  "                        END) , "

        ''ls_SQL = ls_SQL + "         colqtybox = MPM.QtyBox , " & vbCrLf & _
        ''                  "         colboxpallet = CEILING(MP.BoxPallet) , " & vbCrLf & _
        ''                  "         colTotalpalet = CEILING(CASE MP.BoxPallet " & vbCrLf & _
        ''                  "                           WHEN 0 THEN 0 " & vbCrLf & _
        ''                  "                           ELSE ( PDD.DOQty / MPM.QtyBox ) / MP.BoxPallet " & vbCrLf & _
        ''                  "                         END) , " & vbCrLf & _
        ''                  "         colNoPol = PDM.Nopol , " & vbCrLf & _
        ''                  "         colJenisArmada = PDM.JenisArmada, " & vbCrLf & _
        ''                  "         colInvoiceNo = PDM.InvoiceNo " & vbCrLf & _
        ''                  " FROM    dbo.PO_Master POM " & vbCrLf & _
        ''                  "         LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
        ''                  "                                    AND POM.PoNo = POD.PONo "

        ''ls_SQL = ls_SQL + "                                    AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
        ''                  "                                           AND KD.PoNo = POD.PONo " & vbCrLf & _
        ''                  "                                           AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
        ''                  "                                           AND KD.PartNo = POD.PartNo " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
        ''                  "                                           AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
        ''                  "                                           AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
        ''                  "                                           AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
        ''                  "                                                AND KD.KanbanNo = SDD.KanbanNo "

        ''ls_SQL = ls_SQL + "                                                AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
        ''                  "                                                AND KD.PartNo = SDD.PartNo " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf & _
        ''                  "                                                AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf & _
        ''                  "                                                AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.ReceivePASI_Detail PRD ON KD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
        ''                  "                                                 AND KD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
        ''                  "                                                 AND KD.SupplierID = PRD.SupplierID " & vbCrLf & _
        ''                  "                                                 AND KD.PartNo = PRD.PartNo " & vbCrLf & _
        ''                  "                                                 AND KD.PONo = PRD.PartNo " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID "

        ''ls_SQL = ls_SQL + "                                                 AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
        ''                  "                                                 AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.DOPASI_Detail PDD ON KD.AffiliateID = PDD.AffiliateID " & vbCrLf & _
        ''                  "                                            AND KD.KanbanNo = PDD.KanbanNo " & vbCrLf & _
        ''                  "                                            AND KD.SupplierID = PDD.SupplierID " & vbCrLf & _
        ''                  "                                            AND KD.PartNo = PDD.PartNo " & vbCrLf & _
        ''                  "                                            AND KD.PoNo = PDD.PoNo " & vbCrLf & _
        ''                  "                                            AND KD.PONo = SDD.PONo " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
        ''                  "                                            AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
        ''                  "                                            AND PDD.SupplierID = PDM.SupplierID " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo "

        ''ls_SQL = ls_SQL + "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID " & vbCrLf & _
        ''                  "         LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
        ''                  " WHERE   PDM.SuratJalanNo = '" & Session("SJAffiliate") & "' " & vbCrLf & _
        ''                  " ORDER BY KD.KanbanNo "

        'ls_SQL = "  SELECT  Affiliate = MA.AffiliateName, " & vbCrLf & _
        '          "          Address =  RTRIM(MA.Address) + ' ' + Rtrim(MA.City) + ' ' +  Rtrim(MA.PostalCode),  " & vbCrLf & _
        '          "          colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY PDD.KanbanNo, PDD.PartNo )) ,  " & vbCrLf & _
        '          "          colSJ = PDM.SuratjalanNo ,  " & vbCrLf & _
        '          "          colDate = PDM.DeliveryDate ,  " & vbCrLf & _
        '          "          colpono = PDD.PONo ,  " & vbCrLf & _
        '          "          colpartno = PDD.PartNo ,  " & vbCrLf & _
        '          "          colpartname = MP.PartName ,  " & vbCrLf & _
        '          "          colpasideliveryqty = ISNULL(PDD.DOQty, 0) ,  " & vbCrLf & _
        '          "          coldelqtybox = CEILING(CASE MPM.QtyBox  " & vbCrLf & _
        '          "                           WHEN 0 THEN 0  "

        'ls_SQL = ls_SQL + "                           ELSE ISNULL(PDD.DOQty, 0) / MPM.QtyBox  " & vbCrLf & _
        '                  "                         END) ,           " & vbCrLf & _
        '                  "          colqtybox = MPM.QtyBox ,  " & vbCrLf & _
        '                  "          colboxpallet = CEILING(MPM.BoxPallet) ,  " & vbCrLf & _
        '                  "          colTotalpalet = Isnull(PDM.TotalPalet,0) ,  " & vbCrLf & _
        '                  "          colNoPol = PDM.Nopol ,  " & vbCrLf & _
        '                  "          colJenisArmada = PDM.JenisArmada,  " & vbCrLf & _
        '                  "          colInvoiceNo = PDM.InvoiceNo  "

        'ls_SQL = ls_SQL + " From dbo.DOPASI_Detail PDD  " & vbCrLf & _
        '                  " LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
        '                  " 								AND PDD.SuratJalanNo = PDM.SuratJalanNo  " & vbCrLf & _
        '                  " 								--AND PDD.SupplierID = PDM.SupplierID  " & vbCrLf & _
        '                  " LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = PDD.PartNo " & vbCrLf & _
        '                  " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PDD.PartNo AND MPM.AffiliateID = PDD.AffiliateID AND MPM.SupplierID = PDD.SupplierID " & vbCrLf & _
        '                  " LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
        '                  " WHERE   PDM.SuratJalanNo = '" & Session("SJAffiliate") & "' " & vbCrLf & _
        '                  " AND PDM.AffiliateID = '" & Session("RPTAffiliateID") & "' " & vbCrLf & _
        '                  " ORDER BY PDD.KanbanNo "

        ls_SQL = "  SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colpartno )) ,  * FROM (SELECT DISTINCT Affiliate = MA.AffiliateName, " & vbCrLf & _
                  "  Company.Adress1 + ' ' + Company.City1 + ISNULL('. Phone : ' + Company.Phone1,'') + ISNULL(' FAX : ' + Company.Fax1,'') AS Adress1, " & vbCrLf & _
                  "  Company.Adress2 + ' ' + Company.City2 + ISNULL('. Phone : ' + Company.Phone2,'') + ISNULL(' FAX : ' + Company.Fax2,'') AS Adress2, " & vbCrLf & _
                  "          Address =  RTRIM(MA.ConsigneeAddress),  " & vbCrLf & _
                  "          PLBName = ISNULL(MA.PLB_Name,''), " & vbCrLf & _
                  "          PLBAddress = ISNULL(MA.PLB_Address,''), " & vbCrLf & _
                  "          --colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY PDD.KanbanNo, PDD.PartNo )) ,  " & vbCrLf & _
                  "          colSJ = PDM.SuratjalanNo ,  " & vbCrLf & _
                  "          colDate = PDM.DeliveryDate ,  " & vbCrLf & _
                  "          colpono = PDD.PONo ,  " & vbCrLf & _
                  "          colpartno = PDD.PartNo ,  " & vbCrLf & _
                  "          colpartname = MP.PartGroupName ,  " & vbCrLf & _
                  "          colpasideliveryqty = ISNULL(PDD.DOQty, 0) ,  " & vbCrLf & _
                  "          coldelqtybox = CEILING(CASE ISNULL(PDD.POQtyBox,MPM.QtyBox)  " & vbCrLf & _
                  "                           WHEN 0 THEN 0  " & vbCrLf

        ls_SQL = ls_SQL + "                           ELSE ISNULL(PDD.DOQty, 0) / ISNULL(PDD.POQtyBox,MPM.QtyBox)  " & vbCrLf & _
                          "                         END) ,           " & vbCrLf & _
                          "          colqtybox = ISNULL(PDD.POQtyBox,MPM.QtyBox) ,  " & vbCrLf & _
                          "          colboxpallet = CEILING(MPM.BoxPallet) ,  " & vbCrLf & _
                          "          colTotalpalet = Isnull(PDM.TotalPalet,0) ,  " & vbCrLf & _
                          "          colNoPol = PDM.Nopol ,  " & vbCrLf & _
                          "          colJenisArmada = PDM.JenisArmada,  " & vbCrLf & _
                          "          colInvoiceNo = PDM.InvoiceNo,  " & vbCrLf & _
                          "          PDD.SuratJalanNoSupplier " & vbCrLf

        ls_SQL = ls_SQL + " From dbo.DOPASI_Detail PDD  " & vbCrLf & _
                          " LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
                          " 								AND PDD.SuratJalanNo = PDM.SuratJalanNo  " & vbCrLf & _
                          " 								AND PDD.SupplierID = PDM.SupplierID  " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = PDD.PartNo " & vbCrLf & _
                          " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PDD.PartNo AND MPM.AffiliateID = PDD.AffiliateID AND MPM.SupplierID = PDD.SupplierID " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
                          " OUTER APPLY (SELECT TOP 1 * FROM dbo.CompanyProfile WHERE ActiveDate < pdm.DeliveryDate ORDER BY ActiveDate DESC) Company " & vbCrLf & _
                          " WHERE   PDM.SuratJalanNo = '" & Session("SJAffiliate") & "' " & vbCrLf & _
                          " AND PDM.AffiliateID = '" & Session("RPTAffiliateID") & "' " & vbCrLf & _
                          " )x  ORDER BY colpartno " & vbCrLf

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using

        If Not Me.Session("SJAffiliate") Is Nothing Then
            Report.Name = "Delivery " & Me.Session("SJAffiliate")
        End If

        Report.DataSource = ds.Tables(0)
        ASPxDocumentViewer1.Report = Report
        ASPxDocumentViewer1.DataBind()
    End Sub

    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Session.Remove("RPTAffiliateID")
        Session.remove("SJAffiliate")
        Response.Redirect("~/DELIVERY/DeliveryToAffEntry.aspx")
    End Sub
End Class