Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Public Class AffiliateOrderExportList
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
#End Region

#Region "CONTROL EVENTS"
    Private Sub btnSubMenu_Click(sender As Object, e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            up_Fillcombo()
            dtPeriodFrom.Value = Now
            dtPeriodTo.Value = Now
            If Session("M01Url") <> "" Then
                'Call bindData()

                Session.Remove("M01Url")
            End If
            lblInfo.Text = ""
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub
#End Region

    

#Region "PROCEDURE"

    Private Sub up_Fillcombo()
        Dim ls_SQL As String = ""
        'Combo Affiliate
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName UNION ALL SELECT RTRIM(AffiliateID) AffiliateID,AffiliateName FROM dbo.MS_Affiliate" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 50
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 120

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliateName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
        'Combo Supplier
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierCode, '" & clsGlobal.gs_All & "' SupplierName union all select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 50
                .Columns.Add("SupplierName")
                .Columns(1).Width = 120

                .TextField = "SupplierID"
                .DataBind()
                .SelectedIndex = 0
                txtSupplierName.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

    End Sub

    Protected Function GetRowValue(ByVal container As GridViewDataItemTemplateContainer) As String
        If Not IsNothing(container.KeyValue) Then
            Return container.Grid.GetRowValuesByKeyValue(container.KeyValue, "PONo")
        End If
    End Function

    't1
    Protected Function GetPeriod(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPeriod = container.Grid.GetRowValues(container.ItemIndex, "Period")
    End Function
    't2
    Protected Function GetAffiliateID(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateID = container.Grid.GetRowValues(container.ItemIndex, "AffiliateID")
    End Function
    't3
    Protected Function GetAffiliateName(ByVal container As GridViewDataItemTemplateContainer) As String
        GetAffiliateName = container.Grid.GetRowValues(container.ItemIndex, "SupplierID")
    End Function
    't4
    Protected Function GetPO(ByVal container As GridViewDataItemTemplateContainer) As String
        GetPO = container.Grid.GetRowValues(container.ItemIndex, "PONo")
    End Function
    't5
    Protected Function GetDeliveryBy(ByVal container As GridViewDataItemTemplateContainer) As String
        GetDeliveryBy = container.Grid.GetRowValues(container.ItemIndex, "PONo")
    End Function

    Protected Function GetKanban(ByVal container As GridViewDataItemTemplateContainer) As String
        GetKanban = container.Grid.GetRowValues(container.ItemIndex, "PONo")
    End Function
#End Region

End Class