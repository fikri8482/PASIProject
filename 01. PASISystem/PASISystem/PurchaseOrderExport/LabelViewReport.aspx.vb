Imports System.Data.SqlClient
Imports DevExpress.XtraPrinting.BarCode
Imports DevExpress.XtraReports.UI

Public Class LabelListViewReport
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

        Dim Report As New LabelExport

        'ls_sql = " SELECT " & vbCrLf & _
        '                  " 	a.PartNo, " & vbCrLf & _
        '                  " 	a.PONo, " & vbCrLf & _
        '                  " 	PartNo1 = SUBSTRING(RTRIM(a.PartNo),1,2), " & vbCrLf & _
        '                  " 	PartNo2 = SUBSTRING(RTRIM(a.PartNo),3,8), " & vbCrLf & _
        '                  " 	PartNo3 = SUBSTRING(RTRIM(a.PartNo),11,10), " & vbCrLf & _
        '                  " 	a.LabelNo, " & vbCrLf & _
        '                  " 	c.DestinationPort, " & vbCrLf & _
        '                  " 	d.DeliveryPoint, " & vbCrLf & _
        '                  " 	a.AffiliateID, " & vbCrLf & _
        '                  " 	c.ConsigneeCode, "

        'ls_sql = ls_sql + " 	QtyBox, " & vbCrLf & _
        '                  " 	BarcodeNo = 'P' + RTRIM(a.PartNo) + ',Q' + REPLACE(RTRIM(CONVERT(varchar(30),QtyBox,128)),'.00','') + ',K' + RTRIM(a.PONo) + ',1C' + RTRIM(a.LabelNo) + ',8V' + RTRIM(c.ConsigneeCode) + ',22V' + RTRIM(a.SupplierID), " & vbCrLf & _
        '                  " 	SrvDate = GETDATE() " & vbCrLf & _
        '                  " FROM PrintLabelExport a " & vbCrLf & _
        '                  " LEFT JOIN PO_DetailUpload_Export b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo and a.PONo = b.PONo and a.OrderNo = b.OrderNo1 " & vbCrLf & _
        '                  " LEFT JOIN MS_Affiliate c on a.AffiliateID = c.AffiliateID " & vbCrLf & _
        '                  " LEFT JOIN MS_Forwarder d on d.ForwarderID = b.ForwarderID " & vbCrLf & _
        '                  " LEFT JOIN MS_Parts e on e.PartNo = a.PartNo " & vbCrLf & _
        '                  " LEFT JOIN MS_PartMapping f on f.PartNo = a.PartNo and a.SupplierID = f.SupplierID and a.AffiliateID = f.AffiliateID " & vbCrLf & _
        '                  " WHERE Rtrim(a.PONo)+Rtrim(a.AffiliateID)+Rtrim(a.SupplierID)+RTRIM(a.OrderNo) in (" & Trim(Session("eFilter")) & ") " & vbCrLf & _
        '                  " Order by OrderNo, LabelNo "

        ls_sql = "Exec PrintLabelExportReport '" & Trim(Session("eFilter")) & "' "

        ' + ',22V'
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

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
        Response.Redirect("~/PurchaseOrderExport/PrintLabelExportList.aspx")
    End Sub
End Class