Imports System.Data.SqlClient

Public Class viewInvToAff
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Report As New rptInfoiceToAffiliate
        Dim ds As New DataSet
        Dim ls_Invoice As String = Trim(Session("InvInvoice"))
        Dim ls_Affiliate As String = Session("InvAffiliate")
        Dim ls_SJ As String = Session("InvSJ")

        Dim ls_SQL As String = ""

        ls_SQL = ls_SQL + " SELECT No = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY IPM.InvoiceNo)) , " & vbCrLf & _
                          "        Buyer = MA.AffiliateName, " & vbCrLf & _
                          " 	   Address =  RTRIM(MA.Address) + ' ' + Rtrim(MA.City) + ' ' +  Rtrim(MA.PostalCode), " & vbCrLf & _
                          " 	   InvoiceNo = IPM.InvoiceNo,  " & vbCrLf & _
                          " 	   PONo = IPD.PONo,  " & vbCrLf & _
                          " 	   ContainerNo = IPM.ContainerNo, " & vbCrLf & _
                          " 	   PlaceDate = IPM.PlaceDate, " & vbCrLf & _
                          " 	   ShippedPer = IPM.ShippedPer, " & vbCrLf & _
                          " 	   OnOrAboutCondition = IPM.OnOrAboutCondition, " & vbCrLf & _
                          " 	   TermOfDelivery = IPM.DeliveryTerm, " & vbCrLf & _
                          " 	   [From] = IPM.InvFrom, " & vbCrLf & _
                          " 	   Via = IPM.InvVia, "

        ls_SQL = ls_SQL + " 	   [To] = IPM.InvTo, " & vbCrLf & _
                          " 	   Freight = IPM.InvFreight, " & vbCrLf & _
                          " 	   TermOfPayment = IPM.PaymentTerm,  " & vbCrLf & _
                          " 	    " & vbCrLf & _
                          " 	   CartonNo = case when isnull(IPD.InvCartonNo,'') = '' then '' else IPD.InvCartonNo END, " & vbCrLf & _
                          " 	   TotalCarton = CONVERT(NUMERIC(32,0),IPD.InvQty / QtyBox), " & vbCrLf & _
                          " 	   PartNumber = IPD.PartNo,  " & vbCrLf & _
                          " 	   PartName = MP.PartName,  " & vbCrLf & _
                          " 	   Qty = IPD.InvQty,  " & vbCrLf & _
                          " 	   Price = IPD.InvPrice,  " & vbCrLf & _
                          " 	   Amount = IPD.InvAmount, "

        ls_SQL = ls_SQL + " 	   Net = IPD.InvQty * NetWeight, " & vbCrLf & _
                          " 	   Gross = IPD.InvQty * GrossWeight	    " & vbCrLf & _
                          " FROM InvoicePASI_Master IPM " & vbCrLf & _
                          " LEFT JOIN InvoicePASI_Detail IPD ON IPM.InvoiceNo = IPD.InvoiceNo " & vbCrLf & _
                          " 								AND IPM.SuratJalanNo = IPD.SuratJalanNo " & vbCrLf & _
                          " 								AND IPM.AffiliateID = IPD.AffiliateID " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = IPD.PartNo   " & vbCrLf & _
                          " LEFT JOIN MS_partMapping MPM ON MPM.AffiliateID = IPD.AffiliateID " & vbCrLf & _
                          " AND MPM.PartNO = IPD.PartNo " & vbCrLf & _
                          " LEFT JOIN dbo.MS_Affiliate MA ON IPM.AffiliateID = MA.AffiliateID " & vbCrLf & _
                          " WHERE IPM.InvoiceNo = '" & ls_Invoice & "' " & vbCrLf & _
                          "   --AND IPM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf & _
                          "   --AND IPM.SuratJalanNo = '" & ls_SJ & "' "


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using

        Report.DataSource = ds.Tables(0)
        ASPxDocumentViewer1.Report = Report
        ASPxDocumentViewer1.DataBind()
    End Sub

    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Response.Redirect("~/Invoice/InvToAff.aspx")
    End Sub
End Class