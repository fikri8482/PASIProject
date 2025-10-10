Imports System.Data.SqlClient

Public Class PackingListViewReport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramSupplier As String

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
#End Region


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ds As New DataSet
        Dim ls_sql As String = ""
        Dim ls_sparator As String = "|"

        Dim Report As New RptPackingList

        'ls_sql = " select distinct  " & vbCrLf & _
        '          "  Buyer = (PLM.AffiliateID),  " & vbCrLf & _
        '          "  AlamatBuyer =(Select Rtrim(Address) from MS_Affiliate where AffiliateID = PLM.AffiliateID),  " & vbCrLf & _
        '          "  Consignee = isnull((Select Rtrim(DeliveryLocationName) from ms_deliveryPlace where DeliveryLocationCode = KD.DeliveryLocationCode),''),  " & vbCrLf & _
        '          "  AlamatConsignee = isnull((Select Rtrim(Address) from ms_deliveryPlace where DeliveryLocationCode =  KD.DeliveryLocationCode),''),  " & vbCrLf & _
        '          "  Vessel = Rtrim(isnull(PLM.Vessel,'')),  " & vbCrLf & _
        '          "  [From] = Rtrim(isnull(PLM.FromDelivery,'')),  " & vbCrLf & _
        '          "  [To] = Rtrim(isnull(PLM.ToDelivery,'')),  " & vbCrLf & _
        '          "  OnAbout = Rtrim(isnull(PLM.ONAbout,'')),   "

        'ls_sql = ls_sql + "  via = Rtrim(isnull(PLM.ViaDelivery,'')),  " & vbCrLf & _
        '                  "  About = Rtrim(isnull(PLM.AboutDelivery,'')),  " & vbCrLf & _
        '                  "  InvoiceNo = Rtrim(coalesce(PLM.InvoiceNo,'-')),  " & vbCrLf & _
        '                  "  Invdate = Rtrim(Coalesce(DPM.DeliveryDate, DSM.DeliveryDate)),  " & vbCrLf & _
        '                  "  Place = Rtrim(isnull(PLM.Place,'')),  " & vbCrLf & _
        '                  "  OrderNo = '',  " & vbCrLf & _
        '                  "  Privilege = Rtrim(isnull(PLM.Privilege,'')),  " & vbCrLf & _
        '                  "  BLNo = Rtrim(isnull(PLM.AWBBLNo,'')),  " & vbCrLf & _
        '                  "  ContainerNo = Rtrim(isnull(PLM.ContainerNo,'')),  " & vbCrLf & _
        '                  "  Insurance = Rtrim(isnull(PLM.InsurancePolicy,'')),  " & vbCrLf & _
        '                  "  Remarks = Rtrim(isnull(PLM.Remarks,'')),  Paymenterm = Rtrim(isnull(PLM.PaymentTerms,'')),  "

        'ls_sql = ls_sql + "  CartonNo = Rtrim(PLD.CartonNo),  " & vbCrLf & _
        '                  "  CartonQty = convert(numeric,PLD.CartonQty),  " & vbCrLf & _
        '                  "  AffID = Rtrim(PLM.AffiliateID),  " & vbCrLf & _
        '                  "  PONo = Rtrim(PLD.PONo),  " & vbCrLf & _
        '                  "  PartNo = Rtrim(PLD.PartNo),  " & vbCrLf & _
        '                  "  QtyBox = convert(numeric,isnull(MPM.QtyBox,0)),  " & vbCrLf & _
        '                  "  Price = convert(numeric,isnull(MPR.Price,0)),  " & vbCrLf & _
        '                  "  kgm = Isnull(MPM.NetWeight,0)/1000, --convert(numeric,Isnull(MP.NetWeight,0)/1000), " & vbCrLf & _
        '                  "  gross = Isnull(MPM.GrossWeight,0)/1000 --convert(numeric,Isnull(MP.GrossWeight,0)/1000) " & vbCrLf & _
        '                  "  ,Qty =  convert(numeric,Isnull(PLD.DOQty,0)) " & vbCrLf & _
        '                  "  from  " & vbCrLf & _
        '                  "  PLPasi_Master PLM  LEFT JOIN PLPASI_Detail PLD  " & vbCrLf & _
        '                  "  	ON PLM.SuratJalanNo = PLD.SuratJalanNo  "

        'ls_sql = ls_sql + "  	--AND PLM.SupplierID = PLD.SupplierID  " & vbCrLf & _
        '                  "  	AND PLM.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
        '                  "  LEFT JOIN DOPasi_Detail DPD  " & vbCrLf & _
        '                  "  	ON DPD.SuratJalanNo = PLD.SuratJalanNo  " & vbCrLf & _
        '                  "  	AND DPD.SupplierID = PLD.SupplierID  " & vbCrLf & _
        '                  "  	AND DPD.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
        '                  "  	AND DPD.PONo = PLD.PONo  " & vbCrLf & _
        '                  "  LEFT JOIN DOPASI_Master DPM  " & vbCrLf & _
        '                  "  	ON DPM.SuratJalanNo = DPD.SuratJalanNo  	 " & vbCrLf & _
        '                  "  	AND DPD.SupplierID = DPM.SupplierID  " & vbCrLf & _
        '                  "  	AND DPD.AffiliateID = DPM.AffiliateID  "

        'ls_sql = ls_sql + "  LEFT JOIN DOSupplier_Detail DSD  " & vbCrLf & _
        '                  "  	ON DSD.SuratJalanNo = PLD.SuratJalanNo  " & vbCrLf & _
        '                  "  	AND DSD.SupplierID = PLD.SupplierID  " & vbCrLf & _
        '                  "  	AND DSD.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
        '                  "  	AND DSD.PONo = PLD.PONo  " & vbCrLf & _
        '                  "  LEFT JOIN DOSUPPLIER_Master DSM  " & vbCrLf & _
        '                  "  	ON DSM.SuratJalanNo = DSD.SuratJalanNo  	 " & vbCrLf & _
        '                  "  	AND DSD.SupplierID = DSM.SupplierID  " & vbCrLf & _
        '                  "  	AND DSD.AffiliateID = DSM.AffiliateID  " & vbCrLf & _
        '                  "  LEFT JOIN PO_Master POM  	ON POM.PONo = PLD.PONo  " & vbCrLf & _
        '                  "  	AND POM.AffiliateID = PLD.AffiliateID  "

        'ls_sql = ls_sql + "  	AND POM.SupplierID = PLD.SupplierID  " & vbCrLf & _
        '                  " LEFT JOIN KANBAN_DETAIL KD ON KD.KanbanNo = PLD.KanbanNo " & vbCrLf & _
        '                  " 	AND KD.AffiliateID = PLD.AffiliateID " & vbCrLf & _
        '                  " 	--AND KD.SupplierID = PLD.SupplierID " & vbCrLf & _
        '                  " 	AND KD.PONO = PLD.PONO " & vbCrLf & _
        '                  " 	AND KD.PARTNO = PLD.PARTNO " & vbCrLf & _
        '                  "  LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo " & vbCrLf & _
        '                  "  LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PLD.PartNo AND MPM.SupplierID = PLD.SupplierID AND MPM.AffiliateID = PLD.AffiliateID " & vbCrLf & _
        '                  "  LEFT JOIN MS_Price MPR ON MPR.AffiliateID = PLM.AffiliateID " & vbCrLf & _
        '                  "  AND MPR.PartNo = PLD.PartNo and COALESCE(DPM.DeliveryDate,DSM.DeliveryDate) between MPR.StartDate and MPR.EndDate " & vbCrLf


        'ls_sql = ls_sql + " 	Where PLM.SuratJalanNo = '" & Trim(Session("PrintSJ")) & "' " & vbCrLf & _
        '                  " --and PLM.SupplierID = '" & Trim(Session("PrintSuppID")) & "' " & vbCrLf & _
        '                  " and PLM.AffiliateID = '" & Session("AffiliateID") & "' "

        ' '' '' ''ls_sql = " SELECT  DISTINCT" & vbCrLf & _
        ' '' '' ''          " 	Buyer = RTRIM(MA.AffiliateName) " & vbCrLf & _
        ' '' '' ''          " 	, AlamatBuyer = RTRIM(MA.Address) + ' ' + Rtrim(MA.City) + ' ' +  Rtrim(MA.PostalCode) " & vbCrLf & _
        ' '' '' ''          " 	, Consignee = RTRIM(MD.DeliveryLocationName) " & vbCrLf & _
        ' '' '' ''          " 	, AlamatConsignee = RTRIM(MD.Address) " & vbCrLf & _
        ' '' '' ''          " 	, Vessel = Rtrim(isnull(PLM.Vessel,'-')) " & vbCrLf & _
        ' '' '' ''          "     , [From] = Rtrim(isnull(PLM.FromDelivery,'-')) " & vbCrLf & _
        ' '' '' ''          "     , [To] = Rtrim(isnull(PLM.ToDelivery,'-')) " & vbCrLf & _
        ' '' '' ''          "     , OnAbout = Rtrim(isnull(PLM.ONAbout,'-')) " & vbCrLf & _
        ' '' '' ''          " 	, via = Rtrim(isnull(PLM.ViaDelivery,'-')) " & vbCrLf & _
        ' '' '' ''          " 	, About = Rtrim(isnull(PLM.AboutDelivery,'-')) "

        ' '' '' ''ls_sql = ls_sql + "     , InvoiceNo = Rtrim(ISNULL(PLM.InvoiceNo,'-')) " & vbCrLf & _
        ' '' '' ''                  " 	, Invdate = DPM.DeliveryDate " & vbCrLf & _
        ' '' '' ''                  " 	, Place = Rtrim(isnull(PLM.Place,'')) " & vbCrLf & _
        ' '' '' ''                  "     , OrderNo = '' " & vbCrLf & _
        ' '' '' ''                  "     , Privilege = Rtrim(isnull(PLM.Privilege,'')) " & vbCrLf & _
        ' '' '' ''                  "     , BLNo = Rtrim(isnull(PLM.AWBBLNo,'')) " & vbCrLf & _
        ' '' '' ''                  "     , ContainerNo = Rtrim(isnull(PLM.ContainerNo,'')) " & vbCrLf & _
        ' '' '' ''                  "     , Insurance = Rtrim(isnull(PLM.InsurancePolicy,'')) " & vbCrLf & _
        ' '' '' ''                  "     , Remarks = Rtrim(isnull(PLM.Remarks,'')) " & vbCrLf & _
        ' '' '' ''                  " 	, Paymenterm = Rtrim(isnull(PLM.PaymentTerms,'')) " & vbCrLf & _
        ' '' '' ''                  " 	, CartonNo = Rtrim(PLD.CartonNo) "

        ' '' '' ''ls_sql = ls_sql + " 	, CartonQty = convert(numeric,PLD.CartonQty) " & vbCrLf & _
        ' '' '' ''                  "     , AffID = Rtrim(PLD.PartNo) " & vbCrLf & _
        ' '' '' ''                  " 	, PONo = Rtrim(PLD.PONo) " & vbCrLf & _
        ' '' '' ''                  "     , PartNo = Rtrim(MP.PartGroupName) " & vbCrLf & _
        ' '' '' ''                  "     , QtyBox = convert(numeric,isnull(PLD.POQtyBox,MPM.QtyBox)) " & vbCrLf & _
        ' '' '' ''                  "     , Price = convert(numeric,isnull(MPR.Price,0))   " & vbCrLf & _
        ' '' '' ''                  "     , kgm = Convert(DECIMAL(10,2),Isnull(MPM.NetWeight,0)/1000) " & vbCrLf & _
        ' '' '' ''                  "     , gross = Convert(DECIMAL(10,2),Isnull(MPM.GrossWeight,0)/1000)  " & vbCrLf & _
        ' '' '' ''                  "     , Qty = Isnull(PLD.DOQty,0) " & vbCrLf & _
        ' '' '' ''                  " FROM   " & vbCrLf & _
        ' '' '' ''                  " PLPasi_Master PLM "

        ' '' '' ''ls_sql = ls_sql + " LEFT JOIN PLPASI_Detail PLD ON PLM.SuratJalanNo = PLD.SuratJalanNo AND PLM.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
        ' '' '' ''                  " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PLM.AffiliateID " & vbCrLf & _
        ' '' '' ''                  " LEFT JOIN MS_DeliveryPlace MD ON MD.DeliveryLocationCode = PLM.AffiliateID " & vbCrLf & _
        ' '' '' ''                  " LEFT JOIN DOPASI_Master DPM ON DPM.SuratJalanNo = PLD.SuratJalanNo AND PLD.AffiliateID = DPM.AffiliateID     " & vbCrLf & _
        ' '' '' ''                  " LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo " & vbCrLf & _
        ' '' '' ''                  " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PLD.PartNo AND MPM.AffiliateID = PLD.AffiliateID AND MPM.SupplierID = PLD.SupplierID " & vbCrLf & _
        ' '' '' ''                  " LEFT JOIN MS_Price MPR ON MPR.AffiliateID = PLM.AffiliateID AND MPR.PartNo = PLD.PartNo and PLM.DeliveryDate between MPR.StartDate and MPR.EndDate  " & vbCrLf & _
        ' '' '' ''                  " Where PLM.SuratJalanNo = '" & Trim(Session("PrintSJ")) & "' and PLM.AffiliateID = '" & Session("AffiliateID") & "'  " & vbCrLf & _
        ' '' '' ''                  " --GROUP BY PLM.AffiliateID, MA.Address, MD.Address, MD.DeliveryLocationCode, PLM.Vessel, PLM.FromDelivery, PLM.ToDelivery, PLM.OnAbout, " & vbCrLf & _
        ' '' '' ''                  " --PLM.ViaDelivery, PLM.AboutDelivery, PLM.InvoiceNo, PLM.DeliveryDate, DPM.DeliveryDate, PLM.Place, PLM.Privilege, PLM.AWBBLNo,  " & vbCrLf & _
        ' '' '' ''                  " --PLM.ContainerNo, PLM.InsurancePolicy, PLM.Remarks, PLM.PaymentTerms, PLD.CartonNo, PLD.CartonQty, PLD.PONo, " & vbCrLf & _
        ' '' '' ''                  " --PLD.PartNo, MP.QtyBox, MPR.Price, MP.NetWeight, MP.GrossWeight "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_sql = "Exec sp_Affiliate_PackingListViewReport_LoadReport '" + Trim(Session("PrintSJ")) + "', '" + Session("AffiliateID") + "'"
            Dim cmd As New SqlCommand(ls_sql, sqlConn)
            cmd.CommandType = CommandType.Text
            'cmd.Parameters.AddWithValue("SuratJalanNo", Trim(Session("PrintSJ")))
            'cmd.Parameters.AddWithValue("AffiliateID", Session("AffiliateID"))
            cmd.ExecuteNonQuery()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Session("PackingParam") = Session("PrintParam")
        Response.Redirect("~/INVPasi/InvFromPASIDetail.aspx")
    End Sub
End Class