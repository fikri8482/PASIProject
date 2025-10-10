Imports System.Data.SqlClient

Public Class ShippingViewReportExport
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

        If Session("REPORT") = "PL" Then
            Dim Report As New RptInvoice
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
        ElseIf Session("REPORT") = "SI" Then
            Dim Report As New RptShipping
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
        ElseIf Session("REPORT") = "TALLY" Then
            Dim Report As New RptTally
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
        ElseIf Session("REPORT") = "INV-EX" Then
            Dim Report As New RptInvoice
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            ASPxDocumentViewer1.DataBind()
        End If
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        'Session("PackingParam") = Session("PrintParam")
        Response.Redirect("~/ShippingInstruction/ShippingInstructionToForwarder.aspx")
    End Sub
End Class