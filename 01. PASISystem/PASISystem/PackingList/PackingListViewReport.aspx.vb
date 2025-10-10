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

        'ls_sql = " SELECT  DISTINCT" & vbCrLf & _
        '          "     Company.Adress1 + ' ' + Company.City1 + ISNULL('. Phone : ' + Company.Phone1,'') + ISNULL(' FAX : ' + Company.Fax1,'') AS Adress1, " & vbCrLf & _
        '          "     Company.Adress2 + ' ' + Company.City2 + ISNULL('. Phone : ' + Company.Phone2,'') + ISNULL(' FAX : ' + Company.Fax2,'') AS Adress2, " & vbCrLf & _
        '          " 	Buyer = RTRIM(MA.AffiliateName) " & vbCrLf & _
        '          " 	, AlamatBuyer = RTRIM(MA.Address) + ' ' + Rtrim(MA.City) + ' ' +  Rtrim(MA.PostalCode) " & vbCrLf & _
        '          " 	, Consignee = RTRIM(MA.ConsigneeName) " & vbCrLf & _
        '          " 	, AlamatConsignee = RTRIM(MA.ConsigneeAddress) " & vbCrLf & _
        '          "     , PLBName = ISNULL(MA.PLB_Name,'') " & vbCrLf & _
        '          "     , PLBAddress = ISNULL(MA.PLB_Address,'') " & vbCrLf & _
        '          " 	, Vessel = Rtrim(isnull(PLM.Vessel,'-')) " & vbCrLf & _
        '          "     , [From] = Rtrim(isnull(PLM.FromDelivery,'-')) " & vbCrLf & _
        '          "     , [To] = Rtrim(isnull(PLM.ToDelivery,'-')) " & vbCrLf & _
        '          "     , OnAbout = Rtrim(isnull(PLM.ONAbout,'-')) " & vbCrLf & _
        '          " 	, via = Rtrim(isnull(PLM.ViaDelivery,'-')) " & vbCrLf & _
        '          " 	, About = Rtrim(isnull(PLM.AboutDelivery,'-')) "

        'ls_sql = ls_sql + "     , InvoiceNo = Rtrim(ISNULL(PLM.InvoiceNo,'-')) " & vbCrLf & _
        '                  " 	, Invdate = DPM.DeliveryDate " & vbCrLf & _
        '                  " 	, Place = Rtrim(isnull(PLM.Place,'')) " & vbCrLf & _
        '                  "     , OrderNo = '' " & vbCrLf & _
        '                  "     , Privilege = Rtrim(isnull(PLM.Privilege,'')) " & vbCrLf & _
        '                  "     , BLNo = Rtrim(isnull(PLM.AWBBLNo,'')) " & vbCrLf & _
        '                  "     , ContainerNo = Rtrim(isnull(PLM.ContainerNo,'')) " & vbCrLf & _
        '                  "     , Insurance = Rtrim(isnull(PLM.InsurancePolicy,'')) " & vbCrLf & _
        '                  "     , Remarks = Rtrim(isnull(PLM.Remarks,'')) " & vbCrLf & _
        '                  " 	, Paymenterm = Rtrim(isnull(PLM.PaymentTerms,'')) " & vbCrLf & _
        '                  " 	, CartonNo = Rtrim(PLD.CartonNo) "

        'ls_sql = ls_sql + " 	, CartonQty = convert(numeric,PLD.CartonQty) " & vbCrLf & _
        '                  "     , AffID = Rtrim(PLD.PartNo) " & vbCrLf & _
        '                  " 	, PONo = Rtrim(PLD.PONo) " & vbCrLf & _
        '                  "     , PartNo = Rtrim(MP.PartGroupName) " & vbCrLf & _
        '                  "     , QtyBox = convert(numeric,isnull(MPM.QtyBox,0)) " & vbCrLf & _
        '                  "     , Price = convert(numeric,isnull(MPR.Price,0))   " & vbCrLf & _
        '                  "     , kgm = Convert(DECIMAL(10,2),Isnull(MPM.NetWeight,0)/1000) " & vbCrLf & _
        '                  "     , gross = Convert(DECIMAL(10,2),Isnull(MPM.GrossWeight,0)/1000)  " & vbCrLf & _
        '                  "     , Qty = Isnull(PLD.DOQty,0) " & vbCrLf & _
        '                  " FROM   " & vbCrLf & _
        '                  " PLPasi_Master PLM "

        'ls_sql = ls_sql + " LEFT JOIN PLPASI_Detail PLD ON PLM.SuratJalanNo = PLD.SuratJalanNo AND PLM.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
        '                  " LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = PLM.AffiliateID " & vbCrLf & _
        '                  " LEFT JOIN MS_DeliveryPlace MD ON MD.DeliveryLocationCode = PLM.AffiliateID " & vbCrLf & _
        '                  " LEFT JOIN DOPASI_Master DPM ON DPM.SuratJalanNo = PLD.SuratJalanNo AND PLD.AffiliateID = DPM.AffiliateID     " & vbCrLf & _
        '                  " LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo " & vbCrLf & _
        '                  " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = PLD.PartNo AND MPM.AffiliateID = PLD.AffiliateID AND MPM.SupplierID = PLD.SupplierID " & vbCrLf & _
        '                  " LEFT JOIN MS_Price MPR ON MPR.AffiliateID = PLM.AffiliateID AND MPR.PartNo = PLD.PartNo and PLM.DeliveryDate between MPR.StartDate and MPR.EndDate  " & vbCrLf & _
        '                  " OUTER APPLY (SELECT TOP 1 * FROM dbo.CompanyProfile WHERE ActiveDate < DPM.DeliveryDate ORDER BY ActiveDate DESC) Company " & vbCrLf & _
        '                  " Where PLM.SuratJalanNo = '" & Trim(Session("PrintSJ")) & "' and PLM.AffiliateID = '" & Trim(Session("PrintAffID")) & "'  " & vbCrLf & _
        '                  " --GROUP BY PLM.AffiliateID, MA.Address, MD.Address, MD.DeliveryLocationCode, PLM.Vessel, PLM.FromDelivery, PLM.ToDelivery, PLM.OnAbout, " & vbCrLf & _
        '                  " --PLM.ViaDelivery, PLM.AboutDelivery, PLM.InvoiceNo, PLM.DeliveryDate, DPM.DeliveryDate, PLM.Place, PLM.Privilege, PLM.AWBBLNo,  " & vbCrLf & _
        '                  " --PLM.ContainerNo, PLM.InsurancePolicy, PLM.Remarks, PLM.PaymentTerms, PLD.CartonNo, PLD.CartonQty, PLD.PONo, " & vbCrLf & _
        '                  " --PLD.PartNo, MP.QtyBox, MPR.Price, MP.NetWeight, MP.GrossWeight "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_sql = "sp_PASI_PackingListViewReport_LoadReport"
            Dim cmd As New SqlCommand(ls_sql, sqlConn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("SuratJalanNo", Trim(Session("PrintSJ")))
            cmd.Parameters.AddWithValue("AffiliateID", Trim(Session("PrintAffID")))
            cmd.ExecuteNonQuery()

            Dim sqlDA As New SqlDataAdapter(cmd)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using

        If Not Me.Session("PrintSJ") Is Nothing Then
            Report.Name = "Packing List " & Me.Session("PrintSJ")
        End If

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Session("PackingParam") = Session("PrintParam")
        Response.Redirect("~/PackingList/PackingListCreateDetail.aspx")
    End Sub
End Class