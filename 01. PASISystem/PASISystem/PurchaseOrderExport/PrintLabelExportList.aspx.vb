Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView


Public Class PrintLabelExportList
    Inherits System.Web.UI.Page
#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_KanbanDate As String
    Dim ls_approve As Boolean
#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try            
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            Session("M01Url") = Request.QueryString("Session")
            Session("MenuDesc") = "PRINT LABEL EXPORT"

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Clear()
                dtPeriodFrom.Text = Format(Now, "yyyy-MM")
                Call up_fillcombo()
            End If

            Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try
    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Public Sub btnclear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnclear.Click
        Clear()
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", i As Long = 1
        Dim ls_Supplier As String = ""
        Dim ls_deliveryLocation As String = ""
        Dim ls_Affiliate As String = ""
        Dim ls_KanbanNo As String = ""

        Dim ls_filter As String = ""

        Session.Remove("eFilter")

        With grid
            For i = 0 To e.UpdateValues.Count - 1
                If (e.UpdateValues(i).NewValues("cols").ToString()) = 1 Then
                    If ls_filter = "" Then
                        ls_filter = "''" + Trim(e.UpdateValues(i).NewValues("OrderNo").ToString()) + Trim(e.UpdateValues(i).NewValues("AffiliateID").ToString()) + Trim(e.UpdateValues(i).NewValues("SupplierID").ToString()) + Trim(e.UpdateValues(i).NewValues("OrderNo1").ToString()) + "''"
                    Else
                        ls_filter = ls_filter + ",''" + Trim(e.UpdateValues(i).NewValues("OrderNo").ToString()) + Trim(e.UpdateValues(i).NewValues("AffiliateID").ToString()) + Trim(e.UpdateValues(i).NewValues("SupplierID").ToString()) + Trim(e.UpdateValues(i).NewValues("OrderNo1").ToString()) + "''"
                    End If
                End If
            Next

            Session("eFilter") = ls_filter
        End With
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)

            Select Case pAction
                Case "gridload"
                    Dim pDateFrom As String = Split(e.Parameters, "|")(1)
                    Dim pAffiliate As String = Split(e.Parameters, "|")(2)
                    Dim pSupplier As String = Split(e.Parameters, "|")(3)

                    Call up_GridLoad()
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "PrintCard"
                    If Session("eFilter") <> "" Then
                        DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/PurchaseOrderExport/LabelViewReport.aspx")
                    Else
                        Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub colorGrid()
        grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.White
    End Sub

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'SupplierCode
        ls_sql = "SELECT Distinct [Supplier Code] = '" & clsGlobal.gs_All & "' ,[Supplier Name] = '" & clsGlobal.gs_All & "' from ms_supplier union all SELECT [Supplier Code] = RTRIM(supplierID) ,[Supplier Name] = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier Code")
                .Columns(0).Width = 90
                .Columns.Add("Supplier Name")
                .Columns(1).Width = 240

                .TextField = "Supplier Code"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using

        'Affiliate Code
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(Affiliatename) FROM MS_Affiliate where isnull(overseascls, '0') = '1'" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate Code")
                .Columns(0).Width = 90
                .Columns.Add("Affiliate Name")
                .Columns(1).Width = 240

                .TextField = "Affiliate Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub Clear()
        cbosupplier.Text = clsGlobal.gs_All
        txtsupplier.Text = clsGlobal.gs_All
        cboaffiliate.Text = clsGlobal.gs_All
        txtaffiliate.Text = clsGlobal.gs_All
        lblerrmessage.Text = ""
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + " and a.AffiliateID = '" & cboaffiliate.Text & "'"
            End If

            If cbosupplier.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + " and a.SupplierID = '" & cbosupplier.Text & "'"
            End If

            ls_SQL = " SELECT cols = '0', colno = ROW_NUMBER() OVER(ORDER BY OrderNo DESC), * FROM " & vbCrLf & _
                  " ( " & vbCrLf & _
                  " 	SELECT distinct a.PONo OrderNo, a.OrderNo OrderNo1, a.AffiliateID, a.SupplierID " & vbCrLf & _
                  " 	FROM PrintLabelExport a " & vbCrLf & _
                  " 	LEFT JOIN PO_DetailUpload_Export b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo and a.PONo = b.PONo and a.OrderNo = b.OrderNo1 " & vbCrLf & _
                  "     LEFT JOIN PO_Master_Export PME ON PME.AffiliateID = a.AffiliateID and PME.SupplierID = a.SupplierID and PME.PONo = a.PONo and PME.OrderNo1 = a.OrderNo " & vbCrLf & _
                  " 	LEFT JOIN MS_Affiliate c on a.AffiliateID = c.AffiliateID " & vbCrLf & _
                  " 	LEFT JOIN MS_Forwarder d on d.ForwarderID = b.ForwarderID " & vbCrLf & _
                  " 	LEFT JOIN MS_Parts e on e.PartNo = a.PartNo " & vbCrLf & _
                  " 	LEFT JOIN MS_PartMapping f on f.PartNo = a.PartNo and a.SupplierID = f.SupplierID and a.AffiliateID = f.AffiliateID " & vbCrLf & _
                  " 	WHERE PME.Period = '" & Format(dtPeriodFrom.Value, "yyyy-MM") & "-01' " & pWhere & "" & vbCrLf & _
                  " )xyz "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)                
            End With
            sqlConn.Close()
        End Using
    End Sub

#End Region

End Class