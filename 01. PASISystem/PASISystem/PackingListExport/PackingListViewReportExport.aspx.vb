Imports System.Data.SqlClient

Public Class PackingListViewReportExport
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

        If Session("REPORT") = "packing" Then
            Dim Report As New RptPackingListExport
            'Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
            'ASPxDocumentViewer1.DataBind()
        Else
            Dim Report As New RptPackingListExportDetail
            'Report.DataSource = ds.Tables(0)
            ASPxDocumentViewer1.Report = Report
        End If
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        'Session("PackingParam") = Session("PrintParam")
        Response.Redirect("~/PackingListExport/PackingListExportEntry.aspx")
    End Sub
End Class