Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class ShippingViewReportExportCR
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
    Dim ls_ConsString As String = clsGlobal.ConnectionString
    Private crystalReport As CrystalDecisions.CrystalReports.Engine.ReportDocument = Nothing

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Not IsPostBack Then
        Dim ds As New DataSet
        Dim ls_sql As String = ""
        Dim ls_sparator As String = "|"

        ls_ConsString = Replace(ls_ConsString, "=", "")
        ls_ConsString = Replace(ls_ConsString, "Data Source", "")
        ls_ConsString = Replace(ls_ConsString, "Initial Catalog", "")
        ls_ConsString = Replace(ls_ConsString, "Persist Security Info", "")
        ls_ConsString = Replace(ls_ConsString, "User ID", "")
        ls_ConsString = Replace(ls_ConsString, "Password", "")

        Dim gs_DBserver As String = Trim(Split(ls_ConsString, ";")(0))
        Dim gs_DBdatabase As String = Trim(Split(ls_ConsString, ";")(1))
        Dim gs_DBuser As String = Trim(Split(ls_ConsString, ";")(3))
        Dim gs_DBpass As String = Trim(Split(ls_ConsString, ";")(4))

        If Session("REPORT") = "PL" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/Invoice.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        ElseIf Session("REPORT") = "SI" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/rptShippingInstruction.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        ElseIf Session("REPORT") = "TALLY" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/TallyData.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        ElseIf Session("REPORT") = "INV-EX" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/Invoice.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        End If
        'End If
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        'Session("PackingParam") = Session("PrintParam")
        Response.Redirect("~/ShippingInstruction/ShippingInstructionToForwarder.aspx")
    End Sub

    Private Sub Page_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
        If Me.crystalReport IsNot Nothing Then
            Me.crystalReport.Close()
            Me.crystalReport.Dispose()
        End If
    End Sub

    'Protected Sub ExportPDF(sender As Object, e As EventArgs)
    '    Dim crystalReport As New ReportDocument()
    '    'BindReport(crystalReport)
    '    Dim ls_sql As String = ""
    '    Dim ds As New DataSet
    '    Dim gs_DBserver As String = Trim(Split(ls_ConsString, ";")(0))
    '    Dim gs_DBdatabase As String = Trim(Split(ls_ConsString, ";")(1))
    '    Dim gs_DBuser As String = Trim(Split(ls_ConsString, ";")(3))
    '    Dim gs_DBpass As String = Trim(Split(ls_ConsString, ";")(4))
    '    ls_sql = Session("Query")
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()
    '        Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
    '        sqlDA.Fill(ds)
    '        sqlConn.Close()
    '    End Using
    '    crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
    '    crystalReport.Load(Server.MapPath("~/Invoice.rpt"))
    '    crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
    '    crystalReport.SetDataSource(ds.Tables(0))
    '    CrystalReportViewer1.ReportSource = crystalReport



    '    Dim formatType As ExportFormatType = ExportFormatType.NoFormat

    '    formatType = ExportFormatType.PortableDocFormat

    '    crystalReport.ExportToHttpResponse(formatType, Response, True, "Crystal")
    '    'Response.[End]()



    'End Sub

    Private Sub BindReport(crystalReport As ReportDocument)
        Dim ds As New DataSet
        Dim ls_sql As String = ""
        Dim ls_sparator As String = "|"

        ls_ConsString = Replace(ls_ConsString, "=", "")
        ls_ConsString = Replace(ls_ConsString, "Data Source", "")
        ls_ConsString = Replace(ls_ConsString, "Initial Catalog", "")
        ls_ConsString = Replace(ls_ConsString, "Persist Security Info", "")
        ls_ConsString = Replace(ls_ConsString, "User ID", "")
        ls_ConsString = Replace(ls_ConsString, "Password", "")

        Dim gs_DBserver As String = Trim(Split(ls_ConsString, ";")(0))
        Dim gs_DBdatabase As String = Trim(Split(ls_ConsString, ";")(1))
        Dim gs_DBuser As String = Trim(Split(ls_ConsString, ";")(3))
        Dim gs_DBpass As String = Trim(Split(ls_ConsString, ";")(4))

        If Session("REPORT") = "PL" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/Invoice.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        ElseIf Session("REPORT") = "SI" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/rptShippingInstruction.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        ElseIf Session("REPORT") = "TALLY" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/TallyData.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        ElseIf Session("REPORT") = "INV-EX" Then
            ls_sql = Session("Query")
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                sqlDA.Fill(ds)
                sqlConn.Close()
            End Using

            crystalReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            crystalReport.Load(Server.MapPath("~/Invoice.rpt"))
            crystalReport.SetDatabaseLogon(gs_DBuser, gs_DBpass, gs_DBserver, gs_DBdatabase)
            crystalReport.SetDataSource(ds.Tables(0))

            CrystalReportViewer1.ReportSource = crystalReport
        End If
    End Sub
End Class