Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class TallyDataForm
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
                Session("MenuDesc") = "TALLY DATA"
                Call up_fillcombo()
                Call up_fillcombocreateupdate()
                lblerrmessage.Text = ""
                grid.JSProperties("cpdtFStuffingDateFrom") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdtFStuffingDateEnd") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdtStuffingDate") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdtBLAWBDate") = Format(Now, "dd MMM yyyy")
                grid.JSProperties("cpdtPEBDate") = Format(Now, "dd MMM yyyy")
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
                   " (SELECT '') AS colAffiliate, (SELECT '') colContainer, (SELECT '') colSeal, (SELECT '') colTare, " & vbCrLf & _
                   " (SELECT '') colGross, (SELECT '') colVesselNo, (SELECT '') colVesselName, (SELECT '') colContainerSize, " & vbCrLf & _
                   " (SELECT '') colETD, (SELECT '') colShippingLine, (SELECT '') colDestination, (SELECT '') colStuffing, " & vbCrLf & _
                   " (SELECT '') colpallet, (SELECT '') colOrder, (SELECT '') colPart, (SELECT '') colBoxFrom,(SELECT '') colBoxTo, " & vbCrLf & _
                   " (SELECT '') colLength, (SELECT '') colWidth, (SELECT '') colHeight, (SELECT '') colM3, (SELECT '') AS colHeightPallet, " & vbCrLf & _
                   " (SELECT '') colTotalBox, (SELECT '') NoUrut, (SELECT '') colBLAWBNo, (SELECT '') colBLAWBDate, (SELECT '') colPEBNo, (SELECT '') colPEBDate, (SELECT '') colType "

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

            'Invoice
            If txtFInvoiceNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND TD.ShippingInstructionNo = '" & txtFInvoiceNo.Text & "' " & vbCrLf
            End If

            'Container
            If txtFContainerNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND TM.ContainerNo  = '" & txtFContainerNo.Text & "' " & vbCrLf
            End If

            'Pallet
            If txtFPalletNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND TD.PalletNo = '" & txtFPalletNo.Text & "' " & vbCrLf
            End If

            'RECEIVE DATE from - RECEIVE DATE end
            If checkboxdt.Checked = True And dtFStuffingDateFrom.Text <> "" And dtFStuffingDateEnd.Text <> "" Then
                ls_Filter = ls_Filter + " AND CONVERT(date,TM.StuffingDate) BETWEEN '" & Format(dtFStuffingDateFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtFStuffingDateEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            ls_SQL = " SELECT (SELECT '')ACT, TD.ShippingInstructionNo AS colInvoice, TD.ForwarderID as colForwarder, " & vbCrLf & _
                     " TD.AffiliateID AS colAffiliate, TM.ContainerNo AS colContainer, TM.SealNo AS colSeal, TM.Tare AS colTare, " & vbCrLf & _
                     " TM.Gross AS colGross, TM.Vessel AS colVesselNo, TM.NamaKapal as colVesselName, TM.ContainerSize AS colContainerSize, " & vbCrLf & _
                     " TM.ETD AS colETD, TM.ShippingLine AS colShippingLine, TM.DestinationPort AS colDestination, TM.StuffingDate AS colStuffing, " & vbCrLf & _
                     " TD.PalletNo AS colpallet, TD.OrderNo AS colOrder, TD.PartNo AS colPart, TD.CaseNo as colBoxFrom,TD.CaseNo2 as colBoxTo, " & vbCrLf & _
                     " TD.Length AS colLength, TD.Width AS colWidth, TD.Height AS colHeight, TD.M3 AS colM3, TD.WeightPallet AS colHeightPallet, " & vbCrLf & _
                     " SUM(ISNULL(TD.TotalBox, 0 )) AS colTotalBox, " & vbCrLf & _
                     " ROW_NUMBER() OVER( ORDER BY TD.ShippingInstructionNo ASC) AS NoUrut, " & vbCrLf & _
      " RTRIM(ISNULL(TM.BL_AWB_No,'')) colBLAWBNo, TM.BL_AWB_Date colBLAWBDate, RTRIM(ISNULL(TM.PEB_No,'')) colPEBNo, TM.PEB_Date colPEBDate, RTRIM(ISNULL(TM.Type,'')) colType " & vbCrLf & _
                     " FROM dbo.Tally_Detail TD " & vbCrLf & _
                     " INNER JOIN dbo.Tally_Master TM ON TM.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                     " AND TM.ForwarderID = TD.ForwarderID AND TM.ContainerNo = TD.ContainerNo" & vbCrLf & _
                     " AND TM.ShippingInstructionNo = TD.ShippingInstructionNo " & vbCrLf & _
                     " WHERE ISNULL(TD.ShippingInstructionNo, '') <> '' "

            ls_Group = " GROUP BY TD.ShippingInstructionNo, TD.ForwarderID,TD.AffiliateID,TM.ContainerNo,TM.SealNo,TM.Tare,TM.Gross,TM.Vessel,TM.ContainerSize, " & vbCrLf & _
                       " TM.ETD,TM.ShippingLine,TM.DestinationPort,TM.StuffingDate,TD.PalletNo,TD.OrderNo,TD.PartNo,TD.Length,TD.Width,TD.Height,TD.M3,TD.CaseNo,TD.CaseNo2,TD.WeightPallet, TM.NamaKapal, " & vbCrLf & _
                       " TM.BL_AWB_No, TM.BL_AWB_Date, TM.PEB_No, TM.PEB_Date, TM.Type " & vbCrLf & _
                       " ORDER BY TD.ShippingInstructionNo "

            ls_SQL = ls_SQL + ls_Filter + ls_Group

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

            'Invoice
            If txtFInvoiceNo.Text <> "" Then
                ls_SQL = " IF EXISTS (SELECT DISTINCT ShippingInstructionNo FROM dbo.ShippingInstruction_DetailPallet " & vbCrLf & _
                         " WHERE  ShippingInstructionNo = '" & txtFInvoiceNo.Text & "') " & vbCrLf & _
                         " BEGIN " & vbCrLf & _
                         " EXEC sp_Tally_Data_Null2	'" & txtFInvoiceNo.Text & "' " & vbCrLf & _
                         " END " & vbCrLf & _
                         " ELSE " & vbCrLf & _
                         " BEGIN" & vbCrLf & _
                         " SELECT TOP 0 (SELECT '')ACT, (SELECT '') AS colInvoice, (SELECT '') as colForwarder, " & vbCrLf & _
                         " (SELECT '') AS colAffiliate, (SELECT '') colContainer, (SELECT '') colSeal, (SELECT '') colTare,  " & vbCrLf & _
                         " (SELECT '') colGross, (SELECT '') colVesselNo, (SELECT '') colVesselName, (SELECT '') colContainerSize,  " & vbCrLf & _
                         " (SELECT '') colETD, (SELECT '') colShippingLine, (SELECT '') colDestination, (SELECT '') colStuffing,  " & vbCrLf & _
                         " (SELECT '') colpallet, (SELECT '') colOrder, (SELECT '') colPart, (SELECT '') colBoxFrom,(SELECT '') colBoxTo,  " & vbCrLf & _
                         " (SELECT '') colLength, (SELECT '') colWidth, (SELECT '') colHeight, (SELECT '') colM3, (SELECT '') AS colHeightPallet,  " & vbCrLf & _
                         " (SELECT '') colTotalBox, (SELECT '') NoUrut,(SELECT '') colBLAWBNo, (SELECT '') colBLAWBDate, (SELECT '') colPEBNo, (SELECT '') colPEBDate, (SELECT '') colType " & vbCrLf & _
                         " End "
            End If

            ''Pallet
            'If txtFPalletNo.Text <> "" Then
            '    ls_Filter = ls_Filter + " AND TD.PalletNo = '" & txtFPalletNo.Text & "' " & vbCrLf
            'End If

            ''RECEIVE DATE from - RECEIVE DATE end
            'If checkboxdt.Checked = True And dtFStuffingDateFrom.Text <> "" And dtFStuffingDateEnd.Text <> "" Then
            '    ls_Filter = ls_Filter + " AND CONVERT(date,TM.StuffingDate) BETWEEN '" & Format(dtFStuffingDateFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtFStuffingDateEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf
            'End If

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
                If txtFInvoiceNo.Text <> "" Then
                    ls_filter = ls_filter + " AND TD.ShippingInstructionNo = '" & txtFInvoiceNo.Text & "' " & vbCrLf
                End If

                'Container
                If txtFContainerNo.Text <> "" Then
                    ls_filter = ls_filter + " AND TM.ContainerNo  = '" & txtFContainerNo.Text & "' " & vbCrLf
                End If

                'Pallet
                If txtFPalletNo.Text <> "" Then
                    ls_filter = ls_filter + " AND TD.PalletNo = '" & txtFPalletNo.Text & "' " & vbCrLf
                End If

                'RECEIVE DATE from - RECEIVE DATE end
                If checkboxdt.Checked = True And dtFStuffingDateFrom.Text <> "" And dtFStuffingDateEnd.Text <> "" Then
                    ls_filter = ls_filter + " AND CONVERT(date,TM.StuffingDate) BETWEEN '" & Format(dtFStuffingDateFrom.Value, "yyyy-MM-dd") & "' and '" & Format(dtFStuffingDateEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                ls_sql = " SELECT TD.ShippingInstructionNo AS colInvoice, TD.ForwarderID as colForwarder, " & vbCrLf & _
                                  " TD.AffiliateID AS colAffiliate, TM.ContainerNo AS colContainer, TM.SealNo AS colSeal, TM.Tare AS colTare, " & vbCrLf & _
                                  " TM.Gross AS colGross, TM.Vessel AS colVesselNo,TM.NamaKapal as colVesselName, TM.ContainerSize AS colContainerSize, " & vbCrLf & _
                                  " TM.ETD AS colETD, TM.ShippingLine AS colShippingLine, TM.DestinationPort AS colDestination, TM.StuffingDate AS colStuffing, " & vbCrLf & _
                                  " TD.PalletNo AS colpallet, TD.OrderNo AS colOrder, TD.PartNo AS colPart, TD.CaseNo as colBoxFrom,TD.CaseNo2 as colBoxTo, " & vbCrLf & _
                                  " TD.Length AS colLength, TD.Width AS colWidth, TD.Height AS colHeight, TD.M3 AS colM3, TD.WeightPallet AS colHeightPallet, " & vbCrLf & _
                                  " SUM(ISNULL(TD.TotalBox, 0 )) AS colTotalBox, TD.TotalBox QtyCarton, TD.TotalBox * SD.QtyBox QtyPack, " & vbCrLf & _
                                  " TM.BL_AWB_No colBLAWBNo, TM.BL_AWB_Date colBLAWBDate, TM.PEB_No colPEBNo, TM.PEB_Date colPEBDate, TM.Type colType " & vbCrLf & _
                                  " FROM dbo.Tally_Detail TD " & vbCrLf & _
                                  " INNER JOIN dbo.Tally_Master TM ON TM.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                                  " AND TM.ForwarderID = TD.ForwarderID AND TM.ContainerNo = TD.ContainerNo " & vbCrLf & _
                                  " AND TM.ShippingInstructionNo = TD.ShippingInstructionNo " & vbCrLf & _
                                  " INNER JOIN ShippingInstruction_Detail SD ON TD.PartNo = SD.PartNo and TD.ShippingInstructionNo =  SD.ShippingInstructionNo and TD.ForwarderID = SD.ForwarderID AND TD.AffiliateID = SD.AffiliateID " & vbCrLf & _
                                  " WHERE ISNULL(TD.ShippingInstructionNo, '') <> '' "

                ls_group = " GROUP BY TD.ShippingInstructionNo, TD.ForwarderID,TD.AffiliateID,TM.ContainerNo,TM.SealNo,TM.Tare,TM.Gross,TM.Vessel,TM.ContainerSize, " & vbCrLf & _
                           " TM.ETD,TM.ShippingLine,TM.DestinationPort,TM.StuffingDate,TD.PalletNo,TD.OrderNo,TD.PartNo,TD.Length,TD.Width,TD.Height,TD.M3, TD.CaseNo, TD.CaseNo2,TD.WeightPallet, TM.NamaKapal, " & vbCrLf & _
                           " TM.BL_AWB_No, TM.BL_AWB_Date, TM.PEB_No, TM.PEB_Date, TM.Type, TD.TotalBox, SD.QtyBox " & vbCrLf & _
                           " ORDER BY TD.ShippingInstructionNo "


                ls_sql = ls_sql + ls_filter + ls_group

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
    Private Sub grid_CellEditorInitialize(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "colInvoice" Or e.Column.FieldName = "colForwarder" Or e.Column.FieldName = "colAffiliate" Or _
            e.Column.FieldName = "colContainer" Or e.Column.FieldName = "colSeal" Or e.Column.FieldName = "colTare" Or _
            e.Column.FieldName = "colGross" Or e.Column.FieldName = "colVesselNo" Or e.Column.FieldName = "colVesselName" Or _
            e.Column.FieldName = "colContainerSize" Or e.Column.FieldName = "colETD" Or e.Column.FieldName = "colShippingLine" Or _
            e.Column.FieldName = "colDestination" Or e.Column.FieldName = "colStuffing" Or e.Column.FieldName = "colpallet" Or e.Column.FieldName = "colOrder" Or _
            e.Column.FieldName = "colPart" Or e.Column.FieldName = "colBoxFrom" Or e.Column.FieldName = "colBoxTo" Or e.Column.FieldName = "colLength" Or _
            e.Column.FieldName = "colWidth" Or e.Column.FieldName = "colHeight" Or e.Column.FieldName = "colM3" Or e.Column.FieldName = "colHeightPallet" Or _
            e.Column.FieldName = "colTotalBox" Or e.Column.FieldName = "colBLAWBNo" Or e.Column.FieldName = "colBLAWBDate" Or e.Column.FieldName = "colPEBNo" _
            Or e.Column.FieldName = "colPEBDate" Or e.Column.FieldName = "colType") _
          And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub
    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Try

            grid.JSProperties("cpMessage") = Session("AA220Msg")

            If pAction = "gridExcel" Or pAction = "Delete" Then GoTo keluar

            If pAction <> "send" Or pAction <> "gridExcel" Then
                Dim pAffiliate As String = Split(e.Parameters, "|")(1)
                Dim pForwarder As String = Split(e.Parameters, "|")(2)
                Dim pInvoice As String = Split(e.Parameters, "|")(3)
                Dim pContainer As String = Split(e.Parameters, "|")(4)
                Dim pPaletNo As String = Split(e.Parameters, "|")(5)
                Dim pStuffingDateFrom As String = Split(e.Parameters, "|")(6)
                Dim pStuffingDateEnd As String = Split(e.Parameters, "|")(7)
            End If
keluar:

            Select Case pAction

                Case "gridload"
                    'If checkInvoice.Checked = True Then
                    '    Call up_GridLoad_Insert()
                    'Else
                    '    Call up_GridLoad()
                    'End If
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
                    Dim selectItems As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colInvoice;colForwarder;colAffiliate;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colContainer;colSeal;NoUrut;colLength;colWidth;colHeight;colM3;colHeightPallet;colTotalBox"})

                    If selectItems.Count = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.InformationMessage)
                    Else
                        Dim ls_sql As String
                        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                            sqlConn.Open()
                            For Each selectItemId As Object In selectItems

                                Dim pInvoice As String = Split(selectItemId, "|")(0)
                                Dim pForwarder As String = Split(selectItemId, "|")(1)
                                Dim pAffiliate As String = Split(selectItemId, "|")(2)
                                Dim pContainer As String = Split(selectItemId, "|")(8)

                                '[ShippingInstructionNo], [ForwarderID], [AffiliateID], [ContainerNo]
                                ls_sql = " DELETE dbo.Tally_Master " & vbCrLf & _
                                         " Where ShippingInstructionNo = '" & pInvoice & "' " & vbCrLf & _
                                         " and ForwarderID = '" & pForwarder & "' " & vbCrLf & _
                                         " and AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                                         " and ContainerNo = '" & pContainer & "' " & vbCrLf & _
                                         " " & vbCrLf & _
                                         " Delete dbo.Tally_Detail WHERE ShippingInstructionNo = '" & pInvoice & "' " & vbCrLf & _
                                         " AND ForwarderID = '" & pForwarder & "' AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
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
                    FileName = "TemplateTallyData.xlsx"
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
                                     Split(e.Parameter, "|")(10), _
                                     Split(e.Parameter, "|")(11), _
                                     Split(e.Parameter, "|")(12), _
                                     Split(e.Parameter, "|")(13), _
                                     Split(e.Parameter, "|")(14), _
                                     Split(e.Parameter, "|")(15), _
                                     Split(e.Parameter, "|")(16), _
                                     Split(e.Parameter, "|")(17), _
                                     Split(e.Parameter, "|")(18))
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "TallyData" & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                .Cells(3, 4).Value = ": " & Format(dtFStuffingDateFrom.Value, "MMM yyyy") & " - " & Format(dtFStuffingDateEnd.Value, "MMM yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 26).AutoFitColumns()

                .Cells(8, 11, pData.Rows.Count + 7, 11).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 14, pData.Rows.Count + 7, 14).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 6, pData.Rows.Count + 7, 6).Style.Numberformat.Format = "#,##0"
                .Cells(8, 7, pData.Rows.Count + 7, 7).Style.Numberformat.Format = "#,##0"
                .Cells(8, 20, pData.Rows.Count + 7, 20).Style.Numberformat.Format = "#,##0"
                .Cells(8, 21, pData.Rows.Count + 7, 21).Style.Numberformat.Format = "#,##0"
                .Cells(8, 22, pData.Rows.Count + 7, 22).Style.Numberformat.Format = "#,##0"
                .Cells(8, 23, pData.Rows.Count + 7, 23).Style.Numberformat.Format = "#,##0"
                .Cells(8, 24, pData.Rows.Count + 7, 24).Style.Numberformat.Format = "#,##0"
                .Cells(8, 25, pData.Rows.Count + 7, 25).Style.Numberformat.Format = "#,##0"


                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 28)
                EpPlusDrawAllBorders(rgAll)
                .DeleteColumn(28)
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
        txtFInvoiceNo.Text = ""
        txtFPalletNo.Text = ""
        Call up_fillcombo()
        Call up_fillcombocreateupdate()
        grid.JSProperties("cpdtFStuffingDateFrom") = Format(Now, "dd MMM yyyy")
        grid.JSProperties("cpdtFStuffingDateEnd") = Format(Now, "dd MMM yyyy")
        grid.JSProperties("cpdtStuffingDate") = Format(Now, "dd MMM yyyy")
        grid.JSProperties("cpdtBLAWBDate") = Format(Now, "dd MMM yyyy")
        grid.JSProperties("cpdtPEBDate") = Format(Now, "dd MMM yyyy")
        checkboxdt.Checked = True
        'checkInvoice.Checked = False
        txtInvoiceNo.Text = ""
        txtContainerNo.Text = ""
        txtSealNo.Text = ""
        txtTare.Text = ""
        txtVesselNo.Text = ""
        txtSizeContainer.Text = ""
        txtShippingLine.Text = ""
        txtBLAWBNo.Text = ""
        txtPEBNo.Text = ""
        txtType.Text = ""
        'txtVesselName.Text = ""
        txtLocation.Text = ""

        grid.FocusedRowIndex = -1
        up_GridLoadClear()
        lblerrmessage.Text = ""
    End Sub
    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pInvoiceNo As String = "", _
                         Optional ByVal pForwarder As String = "", _
                         Optional ByVal pAffiliate As String = "", _
                         Optional ByVal pContainerNo As String = "", _
                         Optional ByVal pSealNo As String = "", _
                         Optional ByVal pTare As String = "", _
                         Optional ByVal pVesselNo As String = "", _
                         Optional ByVal pSizeContainer As String = "", _
                         Optional ByVal pShippingLine As String = "", _
                         Optional ByVal pVesselName As String = "", _
                         Optional ByVal pStuffingDate As String = "", _
                         Optional ByVal pLocation As String = "", _
                         Optional ByVal pBLAWBNo As String = "", _
                         Optional ByVal pBLAWBDate As String = "", _
                         Optional ByVal pPEBNo As String = "", _
                         Optional ByVal pPEBDate As String = "", _
                         Optional ByVal pType As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""

        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName
        '[ShippingInstructionNo], [ForwarderID], [AffiliateID], [ContainerNo]
        pTare = pTare.Replace(",", "")
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT ShippingInstructionNo FROM dbo.Tally_Master where  ShippingInstructionNo = '" & Trim(pInvoiceNo) & "' " & vbCrLf & _
                         " and ForwarderID = (SELECT DISTINCT ForwarderID FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') " & vbCrLf & _
                         " and AffiliateID = (SELECT DISTINCT AffiliateID FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') " & vbCrLf & _
                         " and ContainerNo = '" & Trim(pContainerNo) & "'"

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


                    pIsNewData = True
                    If pIsNewData = True Then

                        Dim table As DataTable = Nothing
                        table = DirectCast(Session("table"), DataTable)
                        Dim selectItems As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colInvoice;colForwarder;colAffiliate;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colContainer;colSeal;NoUrut;colLength;colWidth;colHeight;colM3;colHeightPallet;colTotalBox"})

                        If selectItems.Count = 0 Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.ErrorMessage)
                            SaveSubmit.JSProperties("cpMessage") = lblerrmessage.Text
                            Exit Sub
                        Else
                            '#INSERT NEW DATA 
                            'TALLY MASTER
                            Dim pForwdr_Header As String = ""
                            Dim pAff_Header As String = ""

                            Dim pContainerNoOld As String
                            For Each selectItemId As Object In selectItems
                                pForwdr_Header = Split(selectItemId, "|")(1)
                                pAff_Header = Split(selectItemId, "|")(2)
                                pContainerNoOld = Split(selectItemId, "|")(8)
                                Exit For
                            Next
                            ls_SQL = " IF EXISTS (SELECT TOP 1 1 FROM dbo.Tally_Master WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "' AND ContainerNo = '" & Trim(pContainerNoOld) & "' )  " & vbCrLf & _
                                     " BEGIN " & vbCrLf

                            If cboCreate.Text = "CREATE" Then
                                ls_SQL = ls_SQL & " DELETE dbo.Tally_Detail WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "' AND ContainerNo = '" & Trim(pContainerNo) & "' " & vbCrLf
                            End If

                            ls_SQL = ls_SQL & " DELETE dbo.Tally_Master WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "' AND ContainerNo = '" & Trim(pContainerNoOld) & "' " & vbCrLf & _
                                    " END " & vbCrLf & _
                                    "  " & vbCrLf & _
                                    " INSERT INTO dbo.Tally_Master " & vbCrLf & _
                                    " ( ShippingInstructionNo ,ForwarderID ,AffiliateID ,ContainerNo ,SealNo ,Tare ,Gross ,TotalCarton , " & vbCrLf & _
                                    " Vessel ,DNNo , ContainerSize ,ETD , ShippingLine ,DestinationPort ,NamaKapal , StuffingDate , TallyCls2,BL_AWB_No,BL_AWB_Date,PEB_No,PEB_Date,Type) " & vbCrLf & _
                                    " VALUES  ( '" & Trim(pInvoiceNo) & "', " & vbCrLf & _
                                    " '" & Trim(pForwdr_Header) & "', " & vbCrLf & _
                                    " '" & Trim(pAff_Header) & "', " & vbCrLf & _
                                    " '" & Trim(pContainerNo) & "', " & vbCrLf & _
                                    " '" & Trim(pSealNo) & "', " & vbCrLf & _
                                    " " & (pTare) & " ,  " & vbCrLf & _
                                    " (SELECT DISTINCT SUM(GrossWeight)countGross FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') , " & vbCrLf & _
                                    " (SELECT DISTINCT Count(LabelNo)CountBox FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') , " & vbCrLf & _
                                    " '" & Trim(pVesselNo) & "', " & vbCrLf & _
                                    " (SELECT DISTINCT TOP 1 SuratJalanNo FROM dbo.ShippingInstruction_DetailPallet WHERE ShippingInstructionNo = '" & Trim(pInvoiceNo) & "') , " & vbCrLf & _
                                    " '" & Trim(pSizeContainer) & "', " & vbCrLf & _
                                    " GETDATE() , " & vbCrLf & _
                                    " '" & Trim(pShippingLine) & "', " & vbCrLf & _
                                    " '" & Trim(pLocation) & "', " & vbCrLf & _
                                    " '" & Trim(pVesselName) & "', " & vbCrLf & _
                                    " '" & Trim(pStuffingDate) & "', " & vbCrLf & _
                                    " '2', " & vbCrLf & _
                                    " '" & Trim(pBLAWBNo) & "', " & vbCrLf & _
                                    " '" & Trim(pBLAWBDate) & "', " & vbCrLf & _
                                    " '" & Trim(pPEBNo) & "', " & vbCrLf & _
                                    " '" & Trim(pPEBDate) & "', " & vbCrLf & _
                                    " '" & Trim(pType) & "')  " & vbCrLf & _
                                    " "
                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                            'TALLY DETAIL
                            For Each selectItemId As Object In selectItems

                                Dim pForwdr As String = Split(selectItemId, "|")(1)
                                Dim pAff As String = Split(selectItemId, "|")(2)
                                Dim pPallet As String = Split(selectItemId, "|")(3)
                                Dim pOrder As String = Split(selectItemId, "|")(4)
                                Dim pPart As String = Split(selectItemId, "|")(5)
                                Dim pBoxFrom As String = Split(selectItemId, "|")(6)
                                Dim pBoxTo As String = Split(selectItemId, "|")(7)
                                Dim pLength As String = Split(selectItemId, "|")(11)
                                Dim pWidth As String = Split(selectItemId, "|")(12)
                                Dim pHeight As String = Split(selectItemId, "|")(13)
                                Dim pM3 As String = Split(selectItemId, "|")(14)
                                Dim pHeightPallet As String = Split(selectItemId, "|")(15)
                                Dim pTotalBox As String = Split(selectItemId, "|")(16)

                                If cboCreate.Text = "CREATE" Then
                                    ls_SQL = " INSERT INTO dbo.Tally_Detail " & vbCrLf & _
                                    "         ( ShippingInstructionNo ,ForwarderID ,AffiliateID ,PalletNo ,OrderNo ,PartNo ,CaseNo ,Length ,Width ,Height ,M3 ,WeightPallet ,CaseNo2 ,TotalBox, ContainerNo, POMOQ, POQtyBox) " & vbCrLf & _
                                    " VALUES  ( '" & Trim(pInvoiceNo) & "' , -- ShippingInstructionNo - char(20) " & vbCrLf & _
                                    "           '" & Trim(pForwdr) & "' , -- ForwarderID - char(20) " & vbCrLf & _
                                    "           '" & Trim(pAff) & "' , -- AffiliateID - char(20) " & vbCrLf & _
                                    "           '" & pPallet.Trim & "' , -- PalletNo - char(20) " & vbCrLf & _
                                    "           '" & pOrder.Trim & "' , -- OrderNo - char(20) " & vbCrLf & _
                                    "           '" & pPart.Trim & "' , -- PartNo - char(20) " & vbCrLf & _
                                    "           '" & pBoxFrom.Trim & "' , -- CaseNo - char(10) " & vbCrLf & _
                                    "           " & pLength & " , -- Length - numeric " & vbCrLf & _
                                    "           " & pWidth & " , -- Width - numeric " & vbCrLf & _
                                    "           " & pHeight & " , -- Height - numeric " & vbCrLf & _
                                    "           " & pM3 & " , -- M3 - numeric " & vbCrLf & _
                                    "           " & pHeightPallet & " , -- WeightPallet - numeric " & vbCrLf & _
                                    "           '" & pBoxTo.Trim & "' , -- CaseNo2 - char(10) " & vbCrLf & _
                                    "           " & pTotalBox & ",  -- TotalBox - numeric " & vbCrLf & _
                                    "           '" & pContainerNo.Trim & "', --ContainerNo  " & vbCrLf & _
                                    "           '" & uf_GetMOQ(Trim(pInvoiceNo), pOrder.Trim, pPart.Trim, pAff.Trim) & "',  -- POMOQ - numeric " & vbCrLf & _
                                    "           '" & uf_GetQtyBox(Trim(pInvoiceNo), pOrder.Trim, pPart.Trim, pAff.Trim) & "')  -- POQtyBox - numeric "
                                Else
                                    ls_SQL = " UPDATE dbo.Tally_Detail SET " & vbCrLf & _
                                        " ContainerNo = '" & Trim(pContainerNo) & "' " & vbCrLf & _
                                        " where ShippingInstructionNo = '" & Trim(pInvoiceNo) & "' " & vbCrLf & _
                                        " and ForwarderID = '" & Trim(pForwarder) & "' " & vbCrLf & _
                                        " and AffiliateID = '" & Trim(pAffiliate) & "' " & vbCrLf & _
                                        " and ContainerNo = '" & Trim(pContainerNoOld) & "' "
                                End If

                                Dim sqlConnDetail2 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlConnDetail2.ExecuteNonQuery()
                                sqlConnDetail2.Dispose()

                            Next selectItemId
                        End If
                        ls_MsgID = "1001"

                    ElseIf pIsNewData = False Then

                        'Call clsMsg.DisplayMessage(lblerrmessage, "6018", clsMessage.MsgType.ErrorMessage)
                        'SaveSubmit.JSProperties("cpMessage") = lblerrmessage.Text
                        'Exit Sub



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
                    Call up_GridLoad_Insert()
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

    Private Function uf_GetQtyBox(ByVal pShipNo As String, ByVal pPoNo As String, ByVal pPartNo As String, ByVal pAffiliateID As String) As Integer
        Dim QTY As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Cn.Open()
            Dim ls_SQL As String
            ls_SQL = "select ISNULL(a.POQtyBox,(ISNULL(b.POQtyBox,ISNULL(c.QtyBox,0)))) Qty from ShippingInstruction_Detail a " & vbCrLf & _
                     " left join PO_Detail_Export b on a.AffiliateID = b.AffiliateID and a.SupplierID =  b.SupplierID and a.PartNo = b.PartNo and a.OrderNo = b.PONo " & vbCrLf & _
                     " left join MS_PartMapping c on a.AffiliateID = c.AffiliateID and a.SupplierID =  c.SupplierID and a.PartNo = c.PartNo " & vbCrLf & _
                     " where a.ShippingInstructionNo = '" + pShipNo + "' and a.OrderNo = '" + pPoNo + "' and a.PartNo = '" + pPartNo + "' and a.AffiliateID = '" + pAffiliateID + "' "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, Cn)
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(dt)
            Cn.Close()

            If dt.Rows.Count > 0 Then
                QTY = dt.Rows(0)("Qty")
            End If
        End Using
        Return QTY
    End Function

    Private Function uf_GetMOQ(ByVal pShipNo As String, ByVal pPoNo As String, ByVal pPartNo As String, ByVal pAffiliateID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Cn.Open()
            Dim ls_SQL As String
            ls_SQL = "select ISNULL(a.POMOQ,(ISNULL(b.POMOQ,ISNULL(c.MOQ,0)))) MOQ from ShippingInstruction_Detail a " & vbCrLf & _
                     " left join PO_Detail_Export b on a.AffiliateID = b.AffiliateID and a.SupplierID =  b.SupplierID and a.PartNo = b.PartNo and a.OrderNo = b.PONo " & vbCrLf & _
                     " left join MS_PartMapping c on a.AffiliateID = c.AffiliateID and a.SupplierID =  c.SupplierID and a.PartNo = c.PartNo " & vbCrLf & _
                     " where a.ShippingInstructionNo = '" + pShipNo + "' and a.OrderNo = '" + pPoNo + "' and a.PartNo = '" + pPartNo + "' and a.AffiliateID = '" + pAffiliateID + "' "
            Dim sqlDA As New SqlDataAdapter(ls_SQL, Cn)
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(dt)
            Cn.Close()

            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

End Class