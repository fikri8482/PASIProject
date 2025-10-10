Imports System.Data.SqlClient

Public Class viewDeliveryToFor
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Report As New rptDeliveryToFor
        Dim ds As New DataSet
        Dim ls_SQL As String

        'ls_SQL = "  SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER( ORDER BY TD.ShippingInstructionNo ASC)) ,  " & vbCrLf & _
        '           "            Affiliate = MF.ForwarderName ,   " & vbCrLf & _
        '           "            Address =  MF.Address +', City : '+ RTRIM(MF.City) +', Postal Code : '+ RTRIM(MF.PostalCode) +', Phone : '+ RTRIM(MF.Phone1) +', Fax : '+ RTRIM(MF.Fax) ,    " & vbCrLf & _
        '           "            PLBName = (MF.Attn+' ( '+ MF.MobilePhone +' )') ,   " & vbCrLf & _
        '           "            PLBAddress = '',   " & vbCrLf & _
        '           "            colSJ = DM.SuratJalanNo ,    " & vbCrLf & _
        '           "            colDate = DM.DeliveryDate ,    " & vbCrLf & _
        '           "            colpono = DE.OrderNo ,    " & vbCrLf & _
        '           "            colpartno = DE.PartNo ,    " & vbCrLf & _
        '           "            colpartname = MMP.PartName ,    " & vbCrLf & _
        '           "            colPallet = DE.PalletNo,  "

        'ls_SQL = ls_SQL + "            colpasideliveryqty = ISNULL((TD.TotalBox * MP.QtyBox), 0) ,               " & vbCrLf & _
        '                  " 		   coldelqtybox = CEILING(TD.TotalBox) ,             " & vbCrLf & _
        '                  "            colqtybox = MP.QtyBox ,    " & vbCrLf & _
        '                  "            colboxpallet = CEILING(TD.TotalBox) ,    " & vbCrLf & _
        '                  "            colTotalpalet = Isnull(TB2.TotalPallet,0) ,    " & vbCrLf & _
        '                  "            colNoPol = DM.NoPol ,    " & vbCrLf & _
        '                  "            colJenisArmada = DM.JenisArmada,    " & vbCrLf & _
        '                  "            colInvoiceNo = DM.ShippingInstructionNo,    " & vbCrLf & _
        '                  "            DM.SuratJalanNo " & vbCrLf & _
        '                  "    FROM dbo.Tally_Master TM  " & vbCrLf & _
        '                  "  LEFT JOIN dbo.Tally_Detail TD ON   "

        'ls_SQL = ls_SQL + "  TD.AffiliateID = TM.AffiliateID AND TD.ForwarderID = TM.ForwarderID AND TD.ContainerNo = TM.ContainerNo   " & vbCrLf & _
        '                  "  AND TD.ShippingInstructionNo = TM.ShippingInstructionNo  " & vbCrLf & _
        '                  "  LEFT JOIN dbo.DOPASI_Detail_Export DE ON  " & vbCrLf & _
        '                  "  DE.AffiliateID = TD.AffiliateID AND DE.ForwarderID = TD.ForwarderID and  DE.CaseNo = TD.CaseNo AND DE.ContainerNo = TD.ContainerNo AND DE.OrderNo = TD.OrderNo  " & vbCrLf & _
        '                  "  LEFT JOIN dbo.DOPASI_Master_Export DM ON DM.AffiliateID = DE.AffiliateID AND DM.ForwarderID = DE.ForwarderID  " & vbCrLf & _
        '                  "  AND DM.ContainerNo = DE.ContainerNo   " & vbCrLf & _
        '                  "  LEFT JOIN dbo.MS_PartMapping MP ON MP.AffiliateID = TD.AffiliateID AND MP.PartNo = TD.PartNo  " & vbCrLf & _
        '                  "  INNER JOIN dbo.MS_Parts MMP ON MMP.PartNo = DE.PartNo  " & vbCrLf & _
        '                  "  INNER JOIN dbo.MS_Forwarder MF ON MF.ForwarderID = DE.ForwarderID " & vbCrLf & _
        '                  "  inner JOIN (  " & vbCrLf & _
        '                  "  SELECT COUNT(TB0.PalletNo)TotalPallet,TB0.ContainerNo FROM (  " & vbCrLf & _
        '                  "  SELECT DISTINCT PalletNo,ContainerNo FROM dbo.Tally_Detail)TB0  " & vbCrLf & _
        '                  "  GROUP BY TB0.ContainerNo)TB2 ON TB2.ContainerNo = TD.ContainerNo  "

        'ls_SQL = ls_SQL + "  INNER JOIN (  " & vbCrLf & _
        '                  "  SELECT SUM(TB1.QTY)QTY,TB1.ContainerNo FROM (  " & vbCrLf & _
        '                  "  SELECT QTY =(TTD.TotalBox * MPP.QtyBox),TTD.ContainerNo FROM dbo.Tally_Detail TTD  INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo)TB1  " & vbCrLf & _
        '                  "  GROUP BY TB1.ContainerNo)TB3 ON TB3.ContainerNo = TD.ContainerNo  " & vbCrLf & _
        '                  "  INNER JOIN (  " & vbCrLf & _
        '                  "  SELECT SUM(TB4.TotalBox)SumTotalBox,TB4.ContainerNo FROM (  " & vbCrLf & _
        '                  "  SELECT TTD.TotalBox,TTD.ContainerNo FROM dbo.Tally_Detail TTD  " & vbCrLf & _
        '                  "  INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo  " & vbCrLf & _
        '                  "  )TB4  " & vbCrLf & _
        '                  "  GROUP BY TB4.ContainerNo)TB5 ON TB5.ContainerNo = TD.ContainerNo  " & vbCrLf & _
        '                  "  WHERE 'A' = 'A' AND ISNULL(DM.SuratJalanNo,'') <> ''  "

        'ls_SQL = ls_SQL + " And   DM.SuratJalanNo = '" & Session("SJForwarder") & "' "

        ls_SQL = " SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER( ORDER BY DE.PalletNo, DE.PartNo)) ,   " & vbCrLf & _
                  "             Affiliate = MF.ForwarderName,    " & vbCrLf & _
                  "             Address =  MF.Address +', City : '+ RTRIM(MF.City) +', Postal Code : '+ RTRIM(MF.PostalCode) +', Phone : '+ RTRIM(MF.Phone1) +', Fax : '+ RTRIM(MF.Fax) ,     " & vbCrLf & _
                  "             PLBName = (MF.Attn+' ( '+ MF.MobilePhone +' )') , " & vbCrLf & _
                  "             PLBAddress = '', " & vbCrLf & _
                  "             colSJ = DM.SuratJalanNo , " & vbCrLf & _
                  "             colDate = DM.DeliveryDate ,     " & vbCrLf & _
                  "             colpono = DE.PalletNo , " & vbCrLf & _
                  "             colpartno = DE.PartNo , " & vbCrLf & _
                  "             colpartname = MMP.PartName , " & vbCrLf & _
                  "             colPallet = DE.OrderNo, "

        ls_SQL = ls_SQL + " 			colpasideliveryqty = SUM(ISNULL((TD.TotalBox * MP.QtyBox), 0)) , " & vbCrLf & _
                          "  		   coldelqtybox = SUM(CEILING(TD.TotalBox)) ,        " & vbCrLf & _
                          "             colqtybox = SUM(MP.QtyBox) , " & vbCrLf & _
                          "             colboxpallet = SUM(CEILING(TD.TotalBox)) , " & vbCrLf & _
                          "             colTotalpalet = Isnull(TB2.TotalPallet,0) , " & vbCrLf & _
                          "             colNoPol = RTRIM(DM.NoPol) , " & vbCrLf & _
                          "             colJenisArmada = RTRIM(DM.JenisArmada), " & vbCrLf & _
                          "             colInvoiceNo = RTRIM(DM.ShippingInstructionNo), " & vbCrLf & _
                          "             DM.SuratJalanNo  " & vbCrLf & _
                          "   FROM dbo.Tally_Master TM   " & vbCrLf & _
                          "   LEFT JOIN dbo.Tally_Detail TD ON     TD.AffiliateID = TM.AffiliateID AND TD.ForwarderID = TM.ForwarderID AND TD.ContainerNo = TM.ContainerNo    "

        ls_SQL = ls_SQL + "   AND TD.ShippingInstructionNo = TM.ShippingInstructionNo   " & vbCrLf & _
                          "   LEFT JOIN dbo.DOPASI_Detail_Export DE ON   " & vbCrLf & _
                          "   DE.AffiliateID = TD.AffiliateID AND DE.ForwarderID = TD.ForwarderID and  DE.CaseNo = TD.CaseNo AND DE.ContainerNo = TD.ContainerNo AND DE.OrderNo = TD.OrderNo   " & vbCrLf & _
                          "   LEFT JOIN dbo.DOPASI_Master_Export DM ON DM.AffiliateID = DE.AffiliateID AND DM.ForwarderID = DE.ForwarderID   " & vbCrLf & _
                          "   AND DM.ContainerNo = DE.ContainerNo " & vbCrLf & _
                          "   LEFT JOIN dbo.MS_PartMapping MP ON MP.AffiliateID = TD.AffiliateID AND MP.PartNo = TD.PartNo " & vbCrLf & _
                          "   INNER JOIN dbo.MS_Parts MMP ON MMP.PartNo = DE.PartNo " & vbCrLf & _
                          "   INNER JOIN dbo.MS_Forwarder MF ON MF.ForwarderID = DE.ForwarderID " & vbCrLf & _
                          "   inner JOIN (   " & vbCrLf & _
                          "   SELECT COUNT(TB0.PalletNo)TotalPallet,TB0.ContainerNo FROM (   " & vbCrLf & _
                          "   SELECT DISTINCT PalletNo,ContainerNo FROM dbo.Tally_Detail)TB0   "

        ls_SQL = ls_SQL + "   GROUP BY TB0.ContainerNo)TB2 ON TB2.ContainerNo = TD.ContainerNo    INNER JOIN (   " & vbCrLf & _
                          "   SELECT SUM(TB1.QTY)QTY,TB1.ContainerNo FROM (   " & vbCrLf & _
                          "   SELECT QTY =(TTD.TotalBox * MPP.QtyBox),TTD.ContainerNo FROM dbo.Tally_Detail TTD  INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo)TB1   " & vbCrLf & _
                          "   GROUP BY TB1.ContainerNo)TB3 ON TB3.ContainerNo = TD.ContainerNo " & vbCrLf & _
                          "   INNER JOIN (   " & vbCrLf & _
                          "   SELECT SUM(TB4.TotalBox)SumTotalBox,TB4.ContainerNo FROM (   " & vbCrLf & _
                          "   SELECT TTD.TotalBox,TTD.ContainerNo FROM dbo.Tally_Detail TTD   " & vbCrLf & _
                          "   INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo   " & vbCrLf & _
                          "   )TB4   " & vbCrLf & _
                          "   GROUP BY TB4.ContainerNo)TB5 ON TB5.ContainerNo = TD.ContainerNo   " & vbCrLf & _
                          "   WHERE 'A' = 'A'  "

        ls_SQL = ls_SQL + "   AND ISNULL(DM.SuratJalanNo,'') <> '' " & vbCrLf & _
                          "   AND DM.SuratJalanNo = '" & Session("SJForwarder") & "' " & vbCrLf & _
                          "   GROUP BY MF.ForwarderName, MF.Address +', City : '+ RTRIM(MF.City) +', Postal Code : '+ RTRIM(MF.PostalCode) +', Phone : '+ RTRIM(MF.Phone1) +', Fax : '+ RTRIM(MF.Fax) , " & vbCrLf & _
                          "    (MF.Attn+' ( '+ MF.MobilePhone +' )') , DM.SuratJalanNo ,DM.DeliveryDate ,DE.OrderNo ,DE.PartNo ,MMP.PartName ,DE.PalletNo,DM.NoPol,DM.JenisArmada, " & vbCrLf & _
                          "    DM.ShippingInstructionNo,TB2.TotalPallet " & vbCrLf & _
                          "   ORDER BY DE.PalletNo, DE.PartNo "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using

        If Not Me.Session("SJForwarder") Is Nothing Then
            Report.Name = "Delivery Export [" & Me.Session("SJForwarder") & "] " & Format(Now, "dd-MMM-yyyy")
        End If

        Report.DataSource = ds.Tables(0)
        ASPxDocumentViewer1.Report = Report
        ASPxDocumentViewer1.DataBind()
    End Sub

    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Session.Remove("SJForwarder")
        Response.Redirect("~/DeliveryExport/DeliveryExportForm.aspx")
    End Sub
End Class