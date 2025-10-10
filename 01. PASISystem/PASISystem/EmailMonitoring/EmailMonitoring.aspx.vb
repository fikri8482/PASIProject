Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class EmailMonitoring
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim pMsgID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            dtPeriodFrom.Value = Now
            up_FillCombo()
            lblInfo.Text = ""
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_BatchUpdate(sender As Object, e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim a As Integer
        Dim ls_sql As String = ""

        a = e.UpdateValues.Count

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("remainingItem")
                Dim sqlComm As New SqlCommand

                For iLoop = 0 To a - 1
                    If e.UpdateValues(iLoop).NewValues("POCls").ToString() <> e.UpdateValues(iLoop).OldValues("POCls").ToString() Then
                        If e.UpdateValues(iLoop).NewValues("POCls").ToString() = "1" Then
                            ls_sql = "Update Affiliate_Master set ExcelCls = 2 " & vbCrLf & _
                                     "where PONo = '" & e.UpdateValues(iLoop).NewValues("PONo").ToString() & "'" & vbCrLf & _
                                     "and AffiliateID = '" & e.UpdateValues(iLoop).NewValues("AffiliateID").ToString() & "'" & vbCrLf & _
                                     "and SupplierID = '" & e.UpdateValues(iLoop).NewValues("SupplierID").ToString() & "'"
                        Else
                            ls_sql = "Update Affiliate_Master set ExcelCls = 1 " & vbCrLf & _
                                     "where PONo = '" & e.UpdateValues(iLoop).NewValues("PONo").ToString() & "'" & vbCrLf & _
                                     "and AffiliateID = '" & e.UpdateValues(iLoop).NewValues("AffiliateID").ToString() & "'" & vbCrLf & _
                                     "and SupplierID = '" & e.UpdateValues(iLoop).NewValues("SupplierID").ToString() & "'"
                        End If

                        sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()
                    End If

                    If e.UpdateValues(iLoop).NewValues("DeliveryCls").ToString() <> e.UpdateValues(iLoop).OldValues("DeliveryCls").ToString() Then
                        If e.UpdateValues(iLoop).NewValues("DeliveryCls").ToString() = "1" Then
                            ls_sql = " Update Kanban_Master set excelcls = '2' " & vbCrLf & _
                                      " where exists ( " & vbCrLf & _
                                      " select *  " & vbCrLf & _
                                      " from Kanban_Detail a  " & vbCrLf & _
                                      " where a.AffiliateID = Kanban_Master.AffiliateID and a.KanbanNo = Kanban_Master.Kanbanno " & vbCrLf & _
                                      " 	  and a.SupplierID = Kanban_Master.SupplierID " & vbCrLf & _
                                      "       and a.PONo = '" & e.UpdateValues(iLoop).NewValues("PONo").ToString() & "' " & vbCrLf & _
                                      "       and a.AffiliateID = '" & e.UpdateValues(iLoop).NewValues("AffiliateID").ToString() & "' " & vbCrLf & _
                                      "       and a.SupplierID = '" & e.UpdateValues(iLoop).NewValues("SupplierID").ToString() & "' " & vbCrLf & _
                                      " ) "
                        Else
                            ls_sql = " Update Kanban_Master set excelcls = '1' " & vbCrLf & _
                                      " where exists ( " & vbCrLf & _
                                      " select *  " & vbCrLf & _
                                      " from Kanban_Detail a  " & vbCrLf & _
                                      " where a.AffiliateID = Kanban_Master.AffiliateID and a.KanbanNo = Kanban_Master.Kanbanno " & vbCrLf & _
                                      " 	  and a.SupplierID = Kanban_Master.SupplierID " & vbCrLf & _
                                      "       and a.PONo = '" & e.UpdateValues(iLoop).NewValues("PONo").ToString() & "' " & vbCrLf & _
                                      "       and a.AffiliateID = '" & e.UpdateValues(iLoop).NewValues("AffiliateID").ToString() & "' " & vbCrLf & _
                                      "       and a.SupplierID = '" & e.UpdateValues(iLoop).NewValues("SupplierID").ToString() & "' " & vbCrLf & _
                                      " ) "
                        End If

                        sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()
                    End If

                    If e.UpdateValues(iLoop).NewValues("SendReceivingTOSupplier").ToString() <> e.UpdateValues(iLoop).OldValues("SendReceivingTOSupplier").ToString() Then
                        If e.UpdateValues(iLoop).NewValues("SendReceivingTOSupplier").ToString() = "1" Then
                            ls_sql = " Update ReceivePASI_Master set excelcls = '2' " & vbCrLf & _
                                      " where exists ( " & vbCrLf & _
                                      " select *  " & vbCrLf & _
                                      " from ReceivePASI_Detail a  " & vbCrLf & _
                                      " where a.AffiliateID = ReceivePASI_Master.AffiliateID and a.SuratJalanNo = ReceivePASI_Master.SuratJalanNo " & vbCrLf & _
                                      " 	  and a.SupplierID = ReceivePASI_Master.SupplierID " & vbCrLf & _
                                      "       and a.PONo = '" & e.UpdateValues(iLoop).NewValues("PONo").ToString() & "' " & vbCrLf & _
                                      "       and a.AffiliateID = '" & e.UpdateValues(iLoop).NewValues("AffiliateID").ToString() & "' " & vbCrLf & _
                                      "       and a.SupplierID = '" & e.UpdateValues(iLoop).NewValues("SupplierID").ToString() & "' " & vbCrLf & _
                                      " ) "
                        Else
                            ls_sql = " Update ReceivePASI_Master set excelcls = '1' " & vbCrLf & _
                                      " where exists ( " & vbCrLf & _
                                      " select *  " & vbCrLf & _
                                      " from ReceivePASI_Detail a  " & vbCrLf & _
                                      " where a.AffiliateID = ReceivePASI_Master.AffiliateID and a.SuratJalanNo = ReceivePASI_Master.SuratJalanNo " & vbCrLf & _
                                      " 	  and a.SupplierID = ReceivePASI_Master.SupplierID " & vbCrLf & _
                                      "       and a.PONo = '" & e.UpdateValues(iLoop).NewValues("PONo").ToString() & "' " & vbCrLf & _
                                      "       and a.AffiliateID = '" & e.UpdateValues(iLoop).NewValues("AffiliateID").ToString() & "' " & vbCrLf & _
                                      "       and a.SupplierID = '" & e.UpdateValues(iLoop).NewValues("SupplierID").ToString() & "' " & vbCrLf & _
                                      " ) "
                        End If

                        sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()
                    End If
                Next
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using    
    End Sub

    'Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
    '    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    'End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "Period" Or e.Column.FieldName = "AffiliateID" Or e.Column.FieldName = "SupplierID" _
             Or e.Column.FieldName = "PONo" Or e.Column.FieldName = "SuratJalanNo") _
             And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("AA220Msg") = lblInfo.Text
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    'Call up_IsiHeader()
                Case "loadSave"
                    grid.PageIndex = 0
                    bindData()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Session("AA220Msg") = ""
        End Try
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call bindData()
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboAffiliate.Text <> clsGlobal.gs_All Then
            pWhere = pWhere & " and a.AffiliateID = '" & cboAffiliate.Text & "'"
        End If

        If cboSupplier.Text <> clsGlobal.gs_All Then
            pWhere = pWhere & " and a.SupplierID = '" & cboSupplier.Text & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " select distinct " & vbCrLf & _
                      " 	a.Period ,a.AffiliateID, a.SupplierID, a.PONo, ISNULL(d.SuratJalanNo,'') SuratJalanNo, ISNULL(KanbanNo,'')KanbanNo," & vbCrLf & _
                      " 	case when b.ExcelCls = 2 then 1 else 0 end POCls, case when c.excelcls = 2  then 1 else 0 end DeliveryCls, case when d.ExcelCls = 2 then 1 else 0 end SendReceivingTOSupplier " & vbCrLf & _
                      " from PO_Master a " & vbCrLf & _
                      " inner join Affiliate_Master b on a.AffiliateID =b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo " & vbCrLf & _
                      " left join  " & vbCrLf & _
                      " ( " & vbCrLf & _
                      " 	select b.AffiliateID, b.SupplierID, b.PONo, excelcls, a.KanbanNo from Kanban_Master a " & vbCrLf & _
                      " 	inner join Kanban_Detail b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.KanbanNo = b.KanbanNo " & vbCrLf & _
                      "  " & vbCrLf & _
                      " )c on a.AffiliateID =c.AffiliateID and a.SupplierID = c.SupplierID and a.PONo = c.PONo "

            ls_SQL = ls_SQL + " left join " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " 	select distinct a.SuratJalanNo, a.SupplierID, a.AffiliateID, b.PONo, a.ExcelCls " & vbCrLf & _
                              " 	from ReceivePASI_Master a " & vbCrLf & _
                              " 	inner join ReceivePASI_Detail b on a.SuratJalanNo = b.SuratJalanNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                              " )d on a.AffiliateID = d.AffiliateID and a.SupplierID = d.SupplierID and a.PONo = d.PONo " & vbCrLf & _
                              " where year(a.Period) = " & Year(dtPeriodFrom.Value) & " and month(a.Period) = " & Month(dtPeriodFrom.Value) & " " & pWhere & "" & vbCrLf & _
                              " order by 1,2,3,4 " & vbCrLf & _
                              "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' Period, ''AffiliateID, ''SupplierID, ''SuratJalanNo, ''POCls, ''DeliveryCls, ''SendReceivingTOSupplier"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateID, '" & clsGlobal.gs_All & "' AffiliateName union all " & vbCrLf & _
                 "select distinct RTRIM(AffiliateID) AffiliateID, AffiliateName from MS_Affiliate" & vbCrLf & _
                 "order by AffiliateID "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 85
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 180

                .TextField = "AffiliateID"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliate.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using



        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierID, '" & clsGlobal.gs_All & "' SupplierName union all " & vbCrLf & _
                 "select distinct RTRIM(SupplierID) SupplierID, SupplierName from MS_Supplier" & vbCrLf & _
                 "order by SupplierID "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 85
                .Columns.Add("SupplierName")
                .Columns(1).Width = 180

                .TextField = "SupplierID"
                .DataBind()
                .SelectedIndex = 0
                txtSupplier.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using
    End Sub
#End Region
End Class