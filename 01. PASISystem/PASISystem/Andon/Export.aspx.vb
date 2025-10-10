Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView

Public Class Export
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim cmdTimeOut As Int32 = 300

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

#Region "Events"

    Protected Sub JamTick_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs)
        Dim ls_SQL As String = "Exec Andon_GetJam"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim dt As New DataTable
            sqlDA.SelectCommand.CommandTimeout = cmdTimeOut
            sqlDA.Fill(dt)
            JamTick.JSProperties("cpMessage") = dt.Rows(0)(0)
            sqlConn.Close()
        End Using
    End Sub

    Protected Sub LegendTick_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs)
        Dim ls_SQL As String = "Exec AndonExport_GetDataCount_ALL"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = cmdTimeOut
            sqlDA.Fill(ds)

            LegendTick.JSProperties("cpMessage_CustOrder") = ds.Tables(0).Rows(0)(0)
            LegendTick.JSProperties("cpMessage_DNPasi") = ds.Tables(1).Rows(0)(0)
            LegendTick.JSProperties("cpMessage_InvoicePasi") = ds.Tables(2).Rows(0)(0)
            LegendTick.JSProperties("cpMessage_SuppOrder") = ds.Tables(3).Rows(0)(0)
            LegendTick.JSProperties("cpMessage_DNSupp") = ds.Tables(4).Rows(0)(0)
            LegendTick.JSProperties("cpMessage_InvSupp") = ds.Tables(5).Rows(0)(0)

            sqlConn.Close()
        End Using
    End Sub

    Private Sub grid_DelayPasi_CustomCallback(sender As Object, e As ASPxGridViewCustomCallbackEventArgs) Handles grid_DelayPasi.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad_DelayPasi()
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grid_DelaySupplier_CustomCallback(sender As Object, e As ASPxGridViewCustomCallbackEventArgs) Handles grid_DelaySupplier.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad_DelaySupplier()
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grid_ReceivePasi_CustomCallback(sender As Object, e As ASPxGridViewCustomCallbackEventArgs) Handles grid_ReceivePasi.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad_ReceivePasi()
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grid_DeliveryPasi_CustomCallback(sender As Object, e As ASPxGridViewCustomCallbackEventArgs) Handles grid_DeliveryPasi.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad_DeliveryPasi()
            End Select
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Functions"

    Private Sub up_GridLoad_DelayPasi()
        Dim ls_SQL As String = "Exec AndonExport_GetData_DelayPasi"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim dt As New DataTable
            sqlDA.SelectCommand.CommandTimeout = cmdTimeOut
            sqlDA.Fill(dt)
            With grid_DelayPasi
                .DataSource = dt
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad_DelaySupplier()
        Dim ls_SQL As String = "Exec AndonExport_GetData_DelaySupplier"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim dt As New DataTable
            sqlDA.SelectCommand.CommandTimeout = cmdTimeOut
            sqlDA.Fill(dt)
            With grid_DelaySupplier
                .DataSource = dt
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad_ReceivePasi()
        Dim ls_SQL As String = "Exec AndonExport_GetData_PlanPenerimaanPasi"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim dt As New DataTable
            sqlDA.SelectCommand.CommandTimeout = cmdTimeOut
            sqlDA.Fill(dt)
            With grid_ReceivePasi
                .DataSource = dt
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad_DeliveryPasi()
        Dim ls_SQL As String = "Exec AndonExport_GetData_PlanPengirimanPasi"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim dt As New DataTable
            sqlDA.SelectCommand.CommandTimeout = cmdTimeOut
            sqlDA.Fill(dt)
            With grid_DeliveryPasi
                .DataSource = dt
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            End With
            sqlConn.Close()
        End Using
    End Sub

#End Region

End Class