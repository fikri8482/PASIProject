Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class DeliveryExportForm
    Inherits System.Web.UI.Page

    '-----------------------------------------------------
    Private grid_Renamed As ASPxGridView
    Private mergedCells As New Dictionary(Of GridViewDataColumn, TableCell)()
    Private cellRowSpans As New Dictionary(Of TableCell, Integer)()
    '-----------------------------------------------------


#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance


    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pOrderNo As String
    Dim pSupplierCode As String
    Dim pSupplierName As String
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""

#End Region
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim param As String = ""
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Session("MenuDesc") = "Delivery Export"
                Call up_fillcombo()
                Call up_fillcombocreateupdate()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdtFDeliveryDateFrom") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdtFDeliveryDateEnd") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdtDeliveryDate") = Format(Now, "dd MMM yyyy")
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try
    End Sub
    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""
        'AFFILIATE
        ls_sql = "SELECT distinct ForwarderID = '" & clsGlobal.gs_All & "', ForwarderName = '" & clsGlobal.gs_All & "' from MS_Forwarder " & vbCrLf & _
                 "UNION Select ForwarderID = RTRIM(ForwarderID) ,ForwarderName = RTRIM(ForwarderName) FROM dbo.MS_Forwarder " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboFForwarder
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("ForwarderID")
                .Columns(0).Width = 70
                .Columns.Add("ForwarderName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtFForwarder.Text = clsGlobal.gs_All
                .TextField = "ForwarderID"
                .DataBind()
            End With
            sqlConn.Close()


            'FORWARDER
            sqlConn.Open()
            ls_sql = "SELECT distinct AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' from MS_Affiliate " & vbCrLf & _
                     "UNION Select AffiliateID = RTRIM(AffiliateID) ,AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate  where isnull(overseascls, '0') = '1'" & vbCrLf
            Dim sqlDA2 As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds2 As New DataSet
            sqlDA2.Fill(ds2)

            With cboFaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds2.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtFaffiliate.Text = clsGlobal.gs_All
                .TextField = "AffiliateID"
                .DataBind()
            End With
            sqlConn.Close()


            'Container
            sqlConn.Open()
            ls_sql = " SELECT DISTINCT TD.ContainerNo as ContainerNo FROM dbo.Tally_Detail TD " & vbCrLf & _
                     " left JOIN dbo.DOPASI_Master_Export DM ON DM.AffiliateID = TD.AffiliateID AND DM.ContainerNo = TD.ContainerNo " & vbCrLf & _
                     " AND DM.ForwarderID = TD.ForwarderID AND DM.ShippingInstructionNo = TD.ShippingInstructionNo " & vbCrLf & _
                     " WHERE ISNULL(DM.SuratJalanNo,'') = '' AND ISNULL(TD.ContainerNo,'') <> '' "
            Dim sqlDA3 As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds3 As New DataSet
            sqlDA3.Fill(ds3)

            With cboContainer
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds3.Tables(0)
                .Columns.Add("ContainerNo")
                .Columns(0).Width = 70
                .SelectedIndex = 0
                .TextField = "ContainerNo"
                .DataBind()
            End With
            sqlConn.Close()
        End Using
    End Sub
    Private Sub up_fillcombocreateupdate()
        Dim ls_sql As String

        ls_sql = ""

        'Shipping No
        ls_sql = "SELECT 'SEARCH' As CreateSearch UNION ALL SELECT 'CREATE' As CreateSearch " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboCreate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("CreateSearch")
                .Columns(0).Width = 100
                .SelectedIndex = 0

                .TextField = "CreateSearch"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub
    Private Sub up_GridLoadClear()
        Dim ls_Sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_Sql = " SELECT TOP 0 (SELECT '')ACT, (SELECT '') AS colInvoice, (SELECT '') as colForwarder, " & vbCrLf & _
                   " (SELECT '') AS colAffiliate, (SELECT '') colpallet, (SELECT '') colLength, (SELECT '') colWidth, " & vbCrLf & _
                   " (SELECT '') colHeight, (SELECT '') colM3, (SELECT '') colHeightPallet, (SELECT '') colOrder, " & vbCrLf & _
                   " (SELECT '') colPart, (SELECT '') colContainer, (SELECT '') colBoxFrom, (SELECT '') colBoxTo, " & vbCrLf & _
                   " (SELECT '') colSuratJalan, (SELECT '') colDeliveryDate, (SELECT '') colPIC, (SELECT '') colJenisArmada,(SELECT '') colDriverName, " & vbCrLf & _
                   " (SELECT '') colDriverContact, (SELECT '') colNoPol, (SELECT '') colTotalBox, (SELECT '') colTotalPallet, (SELECT '') AS colQTY, " & vbCrLf & _
                   " (SELECT '') colTotalQty, (SELECT '') NoUrut "

            Dim sqlDA As New SqlDataAdapter(ls_Sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
            sqlConn.Close()

        End Using
    End Sub
    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim ls_Group As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'Affiliate
            If cboFaffiliate.Text <> clsGlobal.gs_All And cboFaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND TD.AffiliateID = '" & Trim(cboFaffiliate.Text) & "' " & vbCrLf
            End If

            'Forwarder
            If cboFForwarder.Text <> clsGlobal.gs_All And cboFForwarder.Text <> "" Then
                ls_Filter = ls_Filter + " AND TD.ForwarderID = '" & Trim(cboFForwarder.Text) & "' " & vbCrLf
            End If

            'Surat Jalan
            If txtFSJNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND DM.SuratJalanNo = '" & txtFSJNo.Text & "' " & vbCrLf
            End If

            'Container
            If txtFContainerNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND TM.ContainerNo  = '" & txtFContainerNo.Text & "' " & vbCrLf
            End If

            'DELIVERY DATE from - DELIVERY DATE end
            If checkboxdt.Checked = True And dtFDeliveryDateFrom.Text <> "" And dtFDeliveryDateEnd.Text <> "" Then
                ls_Filter = ls_Filter + " AND CONVERT(date,DM.DeliveryDate) BETWEEN '" & Format(dtFDeliveryDateFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtFDeliveryDateEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            ls_SQL = " SELECT  " & vbCrLf & _
                  " colInvoice =  TD.ShippingInstructionNo, " & vbCrLf & _
                  " colForwarder = TM.ForwarderID, " & vbCrLf & _
                  " colAffiliate = TM.AffiliateID, " & vbCrLf & _
                  " colpallet = TD.PalletNo, " & vbCrLf & _
                  " colLength = TD.Length, " & vbCrLf & _
                  " colWidth = TD.Width, " & vbCrLf & _
                  " colHeight = TD.Height, " & vbCrLf & _
                  " colM3 = TD.M3, " & vbCrLf & _
                  " colHeightPallet = Td.WeightPallet , " & vbCrLf & _
                  " colOrder = TD.OrderNo, "

            ls_SQL = ls_SQL + " colPart = TD.PartNo, " & vbCrLf & _
                              " colContainer = TD.ContainerNo, " & vbCrLf & _
                              " colBoxFrom= TD.CaseNo, " & vbCrLf & _
                              " colBoxTo = TD.CaseNo2, " & vbCrLf & _
                              " colSuratJalan = ISNULL(DM.SuratJalanNo,''), " & vbCrLf & _
                              " colDeliveryDate = CASE WHEN DM.DeliveryDate = '' THEN '' ELSE DM.DeliveryDate END, " & vbCrLf & _
                              " colPIC = ISNULL(DM.PIC,''), " & vbCrLf & _
                              " colJenisArmada = ISNULL(DM.JenisArmada,'') , " & vbCrLf & _
                              " colDriverName = ISNULL(DM.DriverName,''), " & vbCrLf & _
                              " colDriverContact = ISNULL(DM.DriverContact,''), " & vbCrLf & _
                              " colNoPol = ISNULL(DM.NoPol,''), "

            ls_SQL = ls_SQL + " colTotalBox = TD.TotalBox, " & vbCrLf & _
                              " colSumTotalBox =TB5.SumTotalBox, " & vbCrLf & _
                              " colTotalPallet = TB2.TotalPallet, " & vbCrLf & _
                              " colQTY = TD.TotalBox * ISNULL(TD.POQtyBox,MP.QtyBox), " & vbCrLf & _
                              " colTotalQty = TB3.QTY, " & vbCrLf & _
                              " ROW_NUMBER() OVER( ORDER BY TD.ShippingInstructionNo ASC) AS NoUrut " & vbCrLf & _
                              " FROM dbo.Tally_Master TM " & vbCrLf & _
                              " LEFT JOIN dbo.Tally_Detail TD ON  " & vbCrLf & _
                              " TD.AffiliateID = TM.AffiliateID AND TD.ForwarderID = TM.ForwarderID AND TD.ContainerNo = TM.ContainerNo  " & vbCrLf & _
                              " AND TD.ShippingInstructionNo = TM.ShippingInstructionNo " & vbCrLf & _
                              " LEFT JOIN dbo.DOPASI_Detail_Export DE ON " & vbCrLf & _
                              " DE.AffiliateID = TD.AffiliateID AND DE.ForwarderID = TD.ForwarderID and "

            ls_SQL = ls_SQL + " DE.CaseNo = TD.CaseNo AND DE.ContainerNo = TD.ContainerNo AND DE.OrderNo = TD.OrderNo " & vbCrLf & _
                              " LEFT JOIN dbo.DOPASI_Master_Export DM ON DM.AffiliateID = DE.AffiliateID AND DM.ForwarderID = DE.ForwarderID " & vbCrLf & _
                              " AND DM.ContainerNo = DE.ContainerNo  " & vbCrLf & _
                              " LEFT JOIN dbo.MS_PartMapping MP ON MP.AffiliateID = TD.AffiliateID AND MP.PartNo = TD.PartNo " & vbCrLf & _
                              " inner JOIN ( " & vbCrLf & _
                              " SELECT COUNT(TB0.PalletNo)TotalPallet,TB0.ContainerNo FROM ( " & vbCrLf & _
                              " SELECT DISTINCT PalletNo,ContainerNo FROM dbo.Tally_Detail)TB0 " & vbCrLf & _
                              " GROUP BY TB0.ContainerNo)TB2 ON TB2.ContainerNo = TD.ContainerNo " & vbCrLf & _
                              " INNER JOIN ( " & vbCrLf & _
                              " SELECT SUM(TB1.QTY)QTY,TB1.ContainerNo FROM ( " & vbCrLf & _
                              " SELECT QTY =(TTD.TotalBox * ISNULL(TTD.POQtyBox,MPP.QtyBox)),TTD.ContainerNo FROM dbo.Tally_Detail TTD "

            ls_SQL = ls_SQL + " INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo)TB1 " & vbCrLf & _
                              " GROUP BY TB1.ContainerNo)TB3 ON TB3.ContainerNo = TD.ContainerNo " & vbCrLf & _
                              " INNER JOIN ( " & vbCrLf & _
                              " SELECT SUM(TB4.TotalBox)SumTotalBox,TB4.ContainerNo FROM ( " & vbCrLf & _
                              " SELECT TTD.TotalBox,TTD.ContainerNo FROM dbo.Tally_Detail TTD " & vbCrLf & _
                              " INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo " & vbCrLf & _
                              " )TB4 " & vbCrLf & _
                              " GROUP BY TB4.ContainerNo)TB5 ON TB5.ContainerNo = TD.ContainerNo " & vbCrLf & _
                              " WHERE 'A' = 'A' AND ISNULL(DM.SuratJalanNo,'') <> '' " & vbCrLf & _
                              "  "

            ls_SQL = ls_SQL + ls_Filter

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                    End With
            sqlConn.Close()

        End Using
    End Sub
    Private Sub up_GridLoad_Insert()
        Dim ls_SQL As String = ""
        Dim ls_Filter As String = ""
        Dim ls_Group As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'Affiliate
            If cboFaffiliate.Text <> clsGlobal.gs_All And cboFaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND TD.AffiliateID = '" & Trim(cboFaffiliate.Text) & "' " & vbCrLf
            End If

            'Forwarder
            If cboFForwarder.Text <> clsGlobal.gs_All And cboFForwarder.Text <> "" Then
                ls_Filter = ls_Filter + " AND TD.ForwarderID = '" & Trim(cboFForwarder.Text) & "' " & vbCrLf
            End If

            ''Invoice
            'If txtFSJNo.Text <> "" Then
            '    ls_Filter = ls_Filter + " AND DM.SuratJalanNo = '" & txtFSJNo.Text & "' " & vbCrLf
            'End If

            'Container
            If cboContainer.Text <> "" Then
                ls_Filter = ls_Filter + " AND TM.ContainerNo  = '" & cboContainer.Text & "' " & vbCrLf
            End If

            ''DELIVERY DATE from - DELIVERY DATE end
            'If checkboxdt.Checked = True And dtFDeliveryDateFrom.Text <> "" And dtFDeliveryDateEnd.Text <> "" Then
            '    ls_Filter = ls_Filter + " AND CONVERT(date,DM.DeliveryDate) BETWEEN '" & Format(dtFDeliveryDateFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtFDeliveryDateEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf
            'End If

            ls_SQL = " SELECT  " & vbCrLf & _
                  " colInvoice =  TD.ShippingInstructionNo, " & vbCrLf & _
                  " colForwarder = TM.ForwarderID, " & vbCrLf & _
                  " colAffiliate = TM.AffiliateID, " & vbCrLf & _
                  " colpallet = TD.PalletNo, " & vbCrLf & _
                  " colLength = TD.Length, " & vbCrLf & _
                  " colWidth = TD.Width, " & vbCrLf & _
                  " colHeight = TD.Height, " & vbCrLf & _
                  " colM3 = TD.M3, " & vbCrLf & _
                  " colHeightPallet = Td.WeightPallet , " & vbCrLf & _
                  " colOrder = TD.OrderNo, "

            ls_SQL = ls_SQL + " colPart = TD.PartNo, " & vbCrLf & _
                              " colContainer = TD.ContainerNo, " & vbCrLf & _
                              " colBoxFrom= TD.CaseNo, " & vbCrLf & _
                              " colBoxTo = TD.CaseNo2, " & vbCrLf & _
                              " colSuratJalan = '" & txtSJNo.Text & "', " & vbCrLf & _
                              " colDeliveryDate = CASE WHEN DM.DeliveryDate = '' THEN '' ELSE DM.DeliveryDate END, " & vbCrLf & _
                              " colPIC = Rtrim(ISNULL(DM.PIC,'')), " & vbCrLf & _
                              " colJenisArmada = Rtrim(ISNULL(DM.JenisArmada,'')) , " & vbCrLf & _
                              " colDriverName = Rtrim(ISNULL(DM.DriverName,'')), " & vbCrLf & _
                              " colDriverContact = Rtrim(ISNULL(DM.DriverContact,'')), " & vbCrLf & _
                              " colNoPol = Rtrim(ISNULL(DM.NoPol,'')), "

            ls_SQL = ls_SQL + " colTotalBox = TD.TotalBox, " & vbCrLf & _
                              " colSumTotalBox =TB5.SumTotalBox, " & vbCrLf & _
                              " colTotalPallet = TB2.TotalPallet, " & vbCrLf & _
                              " colQTY = TD.TotalBox * ISNULL(DE.POQtyBox,MP.QtyBox), " & vbCrLf & _
                              " colTotalQty = TB3.QTY, " & vbCrLf & _
                              " ROW_NUMBER() OVER( ORDER BY TD.ShippingInstructionNo ASC) AS NoUrut " & vbCrLf & _
                              " FROM dbo.Tally_Master TM " & vbCrLf & _
                              " LEFT JOIN dbo.Tally_Detail TD ON  " & vbCrLf & _
                              " TD.AffiliateID = TM.AffiliateID AND TD.ForwarderID = TM.ForwarderID AND TD.ContainerNo = TM.ContainerNo  " & vbCrLf & _
                              " AND TD.ShippingInstructionNo = TM.ShippingInstructionNo " & vbCrLf & _
                              " LEFT JOIN dbo.DOPASI_Detail_Export DE ON " & vbCrLf & _
                              " DE.AffiliateID = TD.AffiliateID AND DE.ForwarderID = TD.ForwarderID and "

            ls_SQL = ls_SQL + " DE.CaseNo = TD.CaseNo AND DE.ContainerNo = TD.ContainerNo AND DE.OrderNo = TD.OrderNo " & vbCrLf & _
                              " LEFT JOIN dbo.DOPASI_Master_Export DM ON DM.AffiliateID = DE.AffiliateID AND DM.ForwarderID = DE.ForwarderID " & vbCrLf & _
                              " AND DM.ContainerNo = DE.ContainerNo  " & vbCrLf & _
                              " LEFT JOIN dbo.MS_PartMapping MP ON MP.AffiliateID = TD.AffiliateID AND MP.PartNo = TD.PartNo " & vbCrLf & _
                              " inner JOIN ( " & vbCrLf & _
                              " SELECT COUNT(TB0.PalletNo)TotalPallet,TB0.ContainerNo FROM ( " & vbCrLf & _
                              " SELECT DISTINCT PalletNo,ContainerNo FROM dbo.Tally_Detail)TB0 " & vbCrLf & _
                              " GROUP BY TB0.ContainerNo)TB2 ON TB2.ContainerNo = TD.ContainerNo " & vbCrLf & _
                              " INNER JOIN ( " & vbCrLf & _
                              " SELECT SUM(TB1.QTY)QTY,TB1.ContainerNo FROM ( " & vbCrLf & _
                              " SELECT QTY =(TTD.TotalBox * ISNULL(TTD.POQtyBox,MPP.QtyBox)),TTD.ContainerNo FROM dbo.Tally_Detail TTD "

            ls_SQL = ls_SQL + " INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo)TB1 " & vbCrLf & _
                              " GROUP BY TB1.ContainerNo)TB3 ON TB3.ContainerNo = TD.ContainerNo " & vbCrLf & _
                              " INNER JOIN ( " & vbCrLf & _
                              " SELECT SUM(TB4.TotalBox)SumTotalBox,TB4.ContainerNo FROM ( " & vbCrLf & _
                              " SELECT TTD.TotalBox,TTD.ContainerNo FROM dbo.Tally_Detail TTD " & vbCrLf & _
                              " INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo " & vbCrLf & _
                              " )TB4 " & vbCrLf & _
                              " GROUP BY TB4.ContainerNo)TB5 ON TB5.ContainerNo = TD.ContainerNo " & vbCrLf & _
                              " WHERE 'A' = 'A' AND ISNULL(DM.SuratJalanNo,'') = '' " & vbCrLf & _
                              "  "

            ls_SQL = ls_SQL + ls_Filter

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid

                .DataSource = ds.Tables(0)
                .DataBind()
            End With
            sqlConn.Close()

        End Using
    End Sub
    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""
        Dim ls_group As String = ""
        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                'Affiliate
                If cboFaffiliate.Text <> clsGlobal.gs_All And cboFaffiliate.Text <> "" Then
                    ls_filter = ls_filter + " AND TD.AffiliateID = '" & Trim(cboFaffiliate.Text) & "' " & vbCrLf
                End If

                'Forwarder
                If cboFForwarder.Text <> clsGlobal.gs_All And cboFForwarder.Text <> "" Then
                    ls_filter = ls_filter + " AND TD.ForwarderID = '" & Trim(cboFForwarder.Text) & "' " & vbCrLf
                End If

                'Invoice
                If txtFSJNo.Text <> "" Then
                    ls_filter = ls_filter + " AND TD.ShippingInstructionNo = '" & txtFSJNo.Text & "' " & vbCrLf
                End If



                'Container
                If txtFContainerNo.Text <> "" Then
                    ls_filter = ls_filter + " AND TM.ContainerNo  = '" & txtFContainerNo.Text & "' " & vbCrLf
                End If

                'DELIVERY DATE from - DELIVERY DATE end
                If checkboxdt.Checked = True And dtFDeliveryDateFrom.Text <> "" And dtFDeliveryDateEnd.Text <> "" Then
                    ls_filter = ls_filter + " AND CONVERT(date,DM.DeliveryDate) BETWEEN '" & Format(dtFDeliveryDateFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtFDeliveryDateEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                ls_sql = " SELECT  " & vbCrLf & _
                      " colInvoice =  TD.ShippingInstructionNo, " & vbCrLf & _
                      " colForwarder = TM.ForwarderID, " & vbCrLf & _
                      " colAffiliate = TM.AffiliateID, " & vbCrLf & _
                      " colpallet = TD.PalletNo, " & vbCrLf & _
                      " colLength = TD.Length, " & vbCrLf & _
                      " colWidth = TD.Width, " & vbCrLf & _
                      " colHeight = TD.Height, " & vbCrLf & _
                      " colM3 = TD.M3, " & vbCrLf & _
                      " colHeightPallet = Td.WeightPallet , " & vbCrLf & _
                      " colOrder = TD.OrderNo, "

                ls_sql = ls_sql + " colPart = TD.PartNo, " & vbCrLf & _
                                  " colContainer = TD.ContainerNo, " & vbCrLf & _
                                  " colBoxFrom= TD.CaseNo, " & vbCrLf & _
                                  " colBoxTo = TD.CaseNo2, " & vbCrLf & _
                                  " colSuratJalan = ISNULL(DM.SuratJalanNo,''), " & vbCrLf & _
                                  " colDeliveryDate = CASE WHEN DM.DeliveryDate = '' THEN '' ELSE DM.DeliveryDate END, " & vbCrLf & _
                                  " colPIC = ISNULL(DM.PIC,''), " & vbCrLf & _
                                  " colJenisArmada = ISNULL(DM.JenisArmada,'') , " & vbCrLf & _
                                  " colDriverName = ISNULL(DM.DriverName,''), " & vbCrLf & _
                                  " colDriverContact = ISNULL(DM.DriverContact,''), " & vbCrLf & _
                                  " colNoPol = ISNULL(DM.NoPol,''), "

                ls_sql = ls_sql + " colTotalBox = TD.TotalBox, " & vbCrLf & _
                                  " colQTY = TD.TotalBox * ISNULL(TD.POQtyBox,MP.QtyBox), " & vbCrLf & _
                                  " colTotalPallet = TB2.TotalPallet, " & vbCrLf & _
                                  " colTotalQty = TB3.QTY, " & vbCrLf & _
                                  " ROW_NUMBER() OVER( ORDER BY TD.ShippingInstructionNo ASC) AS NoUrut " & vbCrLf & _
                                  " FROM dbo.Tally_Master TM " & vbCrLf & _
                                  " LEFT JOIN dbo.Tally_Detail TD ON  " & vbCrLf & _
                                  " TD.AffiliateID = TM.AffiliateID AND TD.ForwarderID = TM.ForwarderID AND TD.ContainerNo = TM.ContainerNo  " & vbCrLf & _
                                  " AND TD.ShippingInstructionNo = TM.ShippingInstructionNo " & vbCrLf & _
                                  " LEFT JOIN dbo.DOPASI_Detail_Export DE ON " & vbCrLf & _
                                  " DE.AffiliateID = TD.AffiliateID AND DE.ForwarderID = TD.ForwarderID and "

                ls_sql = ls_sql + " DE.CaseNo = TD.CaseNo AND DE.ContainerNo = TD.ContainerNo AND DE.OrderNo = TD.OrderNo " & vbCrLf & _
                                  " LEFT JOIN dbo.DOPASI_Master_Export DM ON DM.AffiliateID = DE.AffiliateID AND DM.ForwarderID = DE.ForwarderID " & vbCrLf & _
                                  " AND DM.ContainerNo = DE.ContainerNo  " & vbCrLf & _
                                  " LEFT JOIN dbo.MS_PartMapping MP ON MP.AffiliateID = TD.AffiliateID AND MP.PartNo = TD.PartNo " & vbCrLf & _
                                  " inner JOIN ( " & vbCrLf & _
                                  " SELECT COUNT(TB0.PalletNo)TotalPallet,TB0.ContainerNo FROM ( " & vbCrLf & _
                                  " SELECT DISTINCT PalletNo,ContainerNo FROM dbo.Tally_Detail)TB0 " & vbCrLf & _
                                  " GROUP BY TB0.ContainerNo)TB2 ON TB2.ContainerNo = TD.ContainerNo " & vbCrLf & _
                                  " INNER JOIN ( " & vbCrLf & _
                                  " SELECT SUM(TB1.QTY)QTY,TB1.ContainerNo FROM ( " & vbCrLf & _
                                  " SELECT QTY =(TTD.TotalBox * ISNULL(TTD.POQtyBox,MPP.QtyBox)),TTD.ContainerNo FROM dbo.Tally_Detail TTD "

                ls_sql = ls_sql + " INNER JOIN dbo.MS_PartMapping MPP ON MPP.AffiliateID = TTD.AffiliateID AND MPP.PartNo = TTD.PartNo)TB1 " & vbCrLf & _
                                  " GROUP BY TB1.ContainerNo)TB3 ON TB3.ContainerNo = TD.ContainerNo " & vbCrLf & _
                                  " WHERE 'A' = 'A' AND ISNULL(TD.ShippingInstructionNo,'') <> '' " & vbCrLf & _
                                  "  "

                ls_sql = ls_sql + ls_filter

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 200
                da.Fill(dt)

                Return dt
            End Using

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Try

            grid.JSProperties("cpMessage") = Session("AA220Msg")

            If pAction = "gridExcel" Or pAction = "Delete" Then GoTo keluar

            If pAction <> "send" Or pAction <> "gridExcel" Then
                Dim pAffiliate As String = Split(e.Parameters, "|")(1)
                Dim pForwarder As String = Split(e.Parameters, "|")(2)
                Dim pSJNo As String = Split(e.Parameters, "|")(3)
                Dim pContainer As String = Split(e.Parameters, "|")(4)
                Dim pDeliveryDateFrom As String = Split(e.Parameters, "|")(5)
                Dim pDeliveryDateEnd As String = Split(e.Parameters, "|")(6)
            End If
keluar:

            Select Case pAction

                Case "gridload"

                    If cboCreate.Text = "CREATE" Then
                        Call up_GridLoad_Insert()
                    Else
                        Call up_GridLoad()
                    End If
                    If grid.VisibleRowCount = 0 Then

                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "Delete"
                    Dim table As DataTable = Nothing
                    table = DirectCast(Session("table"), DataTable)
                    Dim selectItems As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colSuratJalan;colForwarder;colAffiliate;colInvoice;colContainer;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colTotalBox;colQTY;NoUrut"})

                    If selectItems.Count = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.InformationMessage)
                    Else
                        Dim ls_sql As String
                        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                            sqlConn.Open()
                            For Each selectItemId As Object In selectItems

                                Dim pSuratJalan As String = Split(selectItemId, "|")(0)
                                'Dim pForwarder As String = Split(selectItemId, "|")(1)
                                'Dim pAffiliate As String = Split(selectItemId, "|")(2)
                                Dim pContainer As String = Split(selectItemId, "|")(4)

                                '[ShippingInstructionNo], [ForwarderID], [AffiliateID], [ContainerNo]
                                ls_sql = " DELETE dbo.DOPASI_Master_Export " & vbCrLf & _
                                         " Where SuratJalanNo = '" & pSuratJalan & "' " & vbCrLf & _
                                         " and ContainerNo = '" & pContainer & "' " & vbCrLf & _
                                         " " & vbCrLf & _
                                         " Delete dbo.DOPASI_Detail_Export WHERE SuratJalanNo = '" & pSuratJalan & "' " & vbCrLf & _
                                         " AND ContainerNo = '" & pContainer & "' "

                                Dim sqlConnDelete As New SqlCommand(ls_sql, sqlConn)
                                sqlConnDelete.ExecuteNonQuery()
                                sqlConnDelete.Dispose()
                            Next selectItemId
                            sqlConn.Close()

                            Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                            grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Call up_GridLoad()
                        End Using
                    End If

                Case "gridExcel"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplateDeliveryExport.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:8", psERR)
                    End If

            End Select

EndProcedure:
            'Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
    Private Sub SaveSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles SaveSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"
                    Dim lb_IsUpdate As Boolean = False
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameter, "|")(2), _
                                     Split(e.Parameter, "|")(3), _
                                     Split(e.Parameter, "|")(4), _
                                     Split(e.Parameter, "|")(5), _
                                     Split(e.Parameter, "|")(6), _
                                     Split(e.Parameter, "|")(7), _
                                     Split(e.Parameter, "|")(8), _
                                     Split(e.Parameter, "|")(9), _
                                     Split(e.Parameter, "|")(10))
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "DeliveryExport" & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\ProgressReport\Import\" & tempFile & "")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets(pSheetName)
            Dim irow As Integer = 0
            Dim icol As Integer = 0

            With ws
                .Cells(3, 4).Value = ": " & Format(dtFDeliveryDateFrom.Value, "dd MMM yyyy") & " - " & Format(dtFDeliveryDateEnd.Value, "dd MMM yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 26).AutoFitColumns()

                .Cells(8, 16, pData.Rows.Count + 7, 16).Style.Numberformat.Format = "dd-mmm-yy"

                '.Cells(8, 6, pData.Rows.Count + 7, 6).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 7, pData.Rows.Count + 7, 7).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 20, pData.Rows.Count + 7, 20).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 21, pData.Rows.Count + 7, 21).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 22, pData.Rows.Count + 7, 22).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 23, pData.Rows.Count + 7, 23).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 24, pData.Rows.Count + 7, 24).Style.Numberformat.Format = "#,##0"
                '.Cells(8, 25, pData.Rows.Count + 7, 25).Style.Numberformat.Format = "#,##0"


                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 26)
                EpPlusDrawAllBorders(rgAll)
                .DeleteColumn(26)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub
    Private Sub EpPlusDrawAllBorders(ByVal Rg As ExcelRange)
        With Rg
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
        End With
    End Sub
    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub
    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnclear.Click
        clear()
    End Sub
    Private Sub clear()
        txtFContainerNo.Text = ""
        txtFSJNo.Text = ""
        Call up_fillcombo()
        Call up_fillcombocreateupdate()
        grid.JSProperties("cpdtFDeliveryDateFrom") = Format(Now, "dd MMM yyyy")
        grid.JSProperties("cpdtFDeliveryDateEnd") = Format(Now, "dd MMM yyyy")
        grid.JSProperties("cpdtDeliveryDate") = Format(Now, "dd MMM yyyy")
        checkboxdt.Checked = True
        txtFSJNo.Text = ""
        txtFContainerNo.Text = ""

        txtSJNo.Text = ""
        txtPIC.Text = ""
        txtJenisArmada.Text = ""
        txtDriverName.Text = ""
        txtDriverContact.Text = ""
        txtNoPol.Text = ""
        txtTotalBox.Text = ""
        txtTotalPallet.Text = ""

        grid.FocusedRowIndex = -1
        up_GridLoadClear()
        lblerrmessage.Text = ""
    End Sub
    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pSJNo As String = "", _
                         Optional ByVal pDeliveryDate As String = "", _
                         Optional ByVal pPIC As String = "", _
                         Optional ByVal pJenisArmada As String = "", _
                         Optional ByVal pDriverName As String = "", _
                         Optional ByVal pDriverContact As String = "", _
                         Optional ByVal pNoPol As String = "", _
                         Optional ByVal pTotalBox As String = "", _
                         Optional ByVal pTotalPallet As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""

        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName


        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT SuratJalanNo,ContainerNo,ShippingInstructionNo,AffiliateID,ForwarderID " & vbCrLf & _
                         " FROM dbo.DOPASI_Master_Export WHERE SuratJalanNo = '" & Trim(pSJNo) & "' "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False
                Else
                    pIsNewData = True
                End If
                sqlConn.Close()
            End Using

            '=============================================================================
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CostCenter")

                    Dim sqlComm As New SqlCommand()

                    If pIsNewData = True Then

                        Dim table As DataTable = Nothing
                        table = DirectCast(Session("table"), DataTable)
                        Dim selectItems As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colSuratJalan;colForwarder;colAffiliate;colInvoice;colContainer;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colTotalBox;colQTY;NoUrut"})

                        If selectItems.Count = 0 Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.ErrorMessage)
                            SaveSubmit.JSProperties("cpMessage") = lblerrmessage.Text
                            Exit Sub
                        Else
                            'DO PASI DETAIL
                            Dim ls_for, ls_aff, ls_ship, ls_cont As String

                            For Each selectItemId As Object In selectItems
                                'colSuratJalan;colForwarder;colAffiliate;colInvoice;colContainer;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colTotalBox;colQTY
                                Dim pSuratJalan As String = Split(selectItemId, "|")(0)
                                Dim pForwarder As String = Split(selectItemId, "|")(1)
                                Dim pAffiliate As String = Split(selectItemId, "|")(2)
                                Dim pInvoice As String = Split(selectItemId, "|")(3)
                                Dim pContainer As String = Split(selectItemId, "|")(4)
                                Dim pPallet As String = Split(selectItemId, "|")(5)
                                Dim pOrder As String = Split(selectItemId, "|")(6)
                                Dim pPart As String = Split(selectItemId, "|")(7)
                                Dim pBoxFrom As String = Split(selectItemId, "|")(8)
                                Dim pBoxTo As String = Split(selectItemId, "|")(9)
                                Dim pTotalPerBox As String = Split(selectItemId, "|")(10)
                                Dim pQTY As String = Split(selectItemId, "|")(11)

                                ls_for = Trim(pForwarder)
                                ls_aff = Trim(pAffiliate)
                                ls_ship = Trim(pInvoice)
                                ls_cont = Trim(pContainer)

                                ls_SQL = " INSERT INTO dbo.DOPASI_Detail_Export " & vbCrLf & _
                                         "         ( SuratJalanNo ,ForwarderID ,AffiliateID ,ShippingInstructionNo ,ContainerNo ,PalletNo ,OrderNo , " & vbCrLf & _
                                         "           PartNo ,CaseNo ,CaseNo2 ,TotalBox ,Qty, POMOQ, POQtyBox ) " & vbCrLf & _
                                         " VALUES  ( '" & Trim(pSJNo) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                                         "           '" & Trim(pForwarder) & "' , -- ForwarderID - char(20) " & vbCrLf & _
                                         "           '" & Trim(pAffiliate) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                         "           '" & Trim(pInvoice) & "' , -- ShippingInstructionNo - char(20) " & vbCrLf & _
                                         "           '" & Trim(pContainer) & "' , -- ContainerNo - char(30) " & vbCrLf & _
                                         "           '" & Trim(pPallet) & "' , -- PalletNo - char(20) " & vbCrLf & _
                                         "           '" & Trim(pOrder) & "' , -- OrderNo - char(20) " & vbCrLf & _
                                         "           '" & Trim(pPart) & "' , -- PartNo - char(20) " & vbCrLf & _
                                         "           '" & Trim(pBoxFrom) & "' , -- CaseNo - char(10) " & vbCrLf & _
                                         "           '" & Trim(pBoxTo) & "' , -- CaseNo2 - char(10) " & vbCrLf & _
                                         "           " & Trim(pTotalPerBox) & " , -- TotalBox - numeric " & vbCrLf & _
                                         "           " & Trim(pQTY) & ",  -- Qty - numeric " & vbCrLf & _
                                         "           '" & uf_GetMOQ(Trim(pOrder), Trim(pPart), Trim(pAffiliate), Trim(pForwarder)) & "',  -- PO MOQ - numeric " & vbCrLf & _
                                         "           '" & uf_GetQtybox(Trim(pOrder), Trim(pPart), Trim(pAffiliate), Trim(pForwarder)) & "'  -- PO QtyBox - numeric " & vbCrLf & _
                                         "         ) "

                                Dim sqlConnDetail2 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlConnDetail2.ExecuteNonQuery()
                                sqlConnDetail2.Dispose()
                            Next selectItemId

                            '#INSERT NEW DATA 
                            'DO PASI MASTER
                            ls_SQL = " INSERT INTO dbo.DOPASI_Master_Export " & vbCrLf & _
                                     "         ( SuratJalanNo ,ForwarderID ,AffiliateID ,ShippingInstructionNo ,ContainerNo ,DeliveryDate ,PIC , " & vbCrLf & _
                                     "           JenisArmada ,DriverName ,DriverContact ,NoPol ,TotalBox ,TotalPalet ,EntryDate ,EntryUser  " & vbCrLf & _
                                     "         ) " & vbCrLf & _
                                     " VALUES  ( '" & Trim(pSJNo) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                                     "           '" & Trim(ls_for) & "' , -- ForwarderID - char(20) " & vbCrLf & _
                                     "           '" & Trim(ls_aff) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                     "           '" & Trim(ls_ship) & "' , -- ShippingInstructionNo - char(20) " & vbCrLf & _
                                     "           '" & Trim(ls_cont) & "' , -- ContainerNo - char(30) " & vbCrLf & _
                                     "           GETDATE() , -- DeliveryDate - date " & vbCrLf & _
                                     "           '" & Trim(pPIC) & "' , -- PIC - char(15) " & vbCrLf & _
                                     "           '" & Trim(pJenisArmada) & "' , -- JenisArmada - char(15) " & vbCrLf & _
                                     "           '" & Trim(pDriverName) & "' , -- DriverName - char(15) " & vbCrLf & _
                                     "           '" & Trim(pDriverContact) & "' , -- DriverContact - char(15) " & vbCrLf & _
                                     "           '" & Trim(pNoPol) & "' , -- NoPol - char(10) " & vbCrLf & _
                                     "           " & Trim(pTotalBox) & " , -- TotalBox - numeric " & vbCrLf & _
                                     "           " & Trim(pTotalPallet) & " , -- TotalPalet - numeric " & vbCrLf & _
                                     "           GETDATE() , -- EntryDate - datetime " & vbCrLf & _
                                     "           '" & Trim(admin) & "' -- EntryUser - char(15) " & vbCrLf & _
                                     "         ) "

                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                        End If
                        ls_MsgID = "1001"

                    ElseIf pIsNewData = False Then

                        Call clsMsg.DisplayMessage(lblerrmessage, "6018", clsMessage.MsgType.ErrorMessage)
                        SaveSubmit.JSProperties("cpMessage") = lblerrmessage.Text
                        Exit Sub

                        '#UPDATE DATA
                        'ls_SQL = " UPDATE dbo.Tally_Master SET " & vbCrLf & _
                        '        " SealNo = '" & Trim(pSealNo) & "' ," & vbCrLf & _
                        '        " Tare  = " & (pTare) & " , " & vbCrLf & _
                        '        " Vessel = '" & Trim(pVesselNo) & "' , " & vbCrLf & _
                        '        " ContainerSize = '" & Trim(pSizeContainer) & "' , " & vbCrLf & _
                        '        " ETD = GETDATE() , " & vbCrLf & _
                        '        " ShippingLine = '" & Trim(pShippingLine) & "' , " & vbCrLf & _
                        '        " NamaKapal = '" & Trim(pVesselName) & "' , " & vbCrLf & _
                        '        " StuffingDate = '" & Trim(pStuffingDate) & "', " & vbCrLf & _
                        '        " DestinationPort = '" & Trim(pLocation) & "' " & vbCrLf & _
                        '        " where ShippingInstructionNo = '" & Trim(pInvoiceNo) & "' " & vbCrLf & _
                        '        " and ForwarderID = (SELECT DISTINCT ForwarderID FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') " & vbCrLf & _
                        '        " and AffiliateID = (SELECT DISTINCT AffiliateID FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') " & vbCrLf & _
                        '        " and ContainerNo = '" & Trim(pContainerNo) & "'"

                        'ls_MsgID = "1002"

                        'sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        'sqlComm.ExecuteNonQuery()

                    End If
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using

            If ls_MsgID = "1001" Then
                If cboCreate.Text = "CREATE" Then
                    Call up_GridLoad()
                    ls_MsgID = "1001"
                End If
                'If checkInvoice.Checked = True Then
                '    Call up_GridLoad_Insert()
                'End If
            End If
            Call clsMsg.DisplayMessage(lblerrmessage, ls_MsgID, clsMessage.MsgType.InformationMessage)
            SaveSubmit.JSProperties("cpMessage") = lblerrmessage.Text
            SaveSubmit.JSProperties("cpType") = "info"


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub btnprint_Click(sender As Object, e As EventArgs) Handles btnprint.Click
        Dim selectItems As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colSuratJalan;colForwarder;colAffiliate;colInvoice;colContainer;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colTotalBox;colQTY;NoUrut"})
        If selectItems.Count = 0 Then
            clear()
            Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.ErrorMessage)
            SaveSubmit.JSProperties("cpMessage") = lblerrmessage.Text
            Exit Sub
        Else
            For Each selectItemId As Object In selectItems

                Dim pSuratJalan As String = Split(selectItemId, "|")(0)
                If pSuratJalan <> "" Then
                    Session("SJForwarder") = Trim(pSuratJalan)
                End If
                Response.Redirect("~/DeliveryExport/viewDeliveryToFor.aspx")
            Next
        End If
    End Sub

    Private Function uf_GetMOQ(ByVal pOrderNo As String, ByVal pPartNo As String, ByVal pAffiliateID As String, ByVal pForwarderID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(a.POMOQ,b.MOQ) MOQ FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE OrderNo1 ='" + pOrderNo + "' AND a.PartNo = '" + pPartNo + "' AND a.AffiliateID = '" + pAffiliateID + "' AND a.ForwarderID = '" + pForwarderID + "' "
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

    Private Function uf_GetQtybox(ByVal pOrderNo As String, ByVal pPartNo As String, ByVal pAffiliateID As String, ByVal pForwarderID As String) As Integer
        Dim Qty As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(a.POQtyBox,b.QtyBox) Qty FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE OrderNo1 ='" + pOrderNo + "' AND a.PartNo = '" + pPartNo + "' AND a.AffiliateID = '" + pAffiliateID + "' AND a.ForwarderID = '" + pForwarderID + "' "
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                Qty = dt.Rows(0)("Qty")
            End If
        End Using
        Return Qty
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(clsGlobal.ConnectionString)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function
End Class