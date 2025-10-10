Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing

Public Class PackingListExportEntry
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

#End Region
#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_fillcombo()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdtfrom") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdtto") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdt1") = Format(Now, "01 MMM yyyy")
                grid.JSProperties("cpdeliver") = "ALL"
                grid.JSProperties("cpreceive") = "ALL"
                grid.JSProperties("cpkanban") = "ALL"
            End If
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(sender As Object, e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub
#End Region
#Region "PROCEDURE"
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'AFFILIATE
        ls_sql = "SELECT distinct AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' from MS_Affiliate " & vbCrLf & _
                 "UNION Select AffiliateID = RTRIM(AffiliateID) ,AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "AffiliateID"
                .DataBind()
            End With
            sqlConn.Close()

            'PartNo
            ls_sql = "SELECT distinct PartNo = '" & clsGlobal.gs_All & "', PartName = '" & clsGlobal.gs_All & "' from MS_Parts " & vbCrLf & _
                "Union all SELECT PartNo = RTRIM(PartNo) ,PartName = RTRIM(PartName) FROM MS_Parts " & vbCrLf
            sqlConn.Open()

            Dim sqlDAA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds1 As New DataSet
            sqlDAA.Fill(ds1)

            With cbopart
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds1.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 70
                .Columns.Add("PartName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtpart.Text = clsGlobal.gs_All
                .TextField = "Partno"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub
#End Region


    Protected Sub btnPrint_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPrint.Click
        Session("REPORT") = "packing"
        Response.Redirect("~/PackingListExport/PackingListViewReportExport.aspx")
    End Sub

    Private Sub btnPrint0_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint0.Click
        Session("REPORT") = "detail"
        Response.Redirect("~/PackingListExport/PackingListViewReportExport.aspx")
    End Sub
End Class