Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports DevExpress.Web.ASPxMenu
Imports OfficeOpenXml
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Net

Imports DevExpress.XtraPrinting
Imports DevExpress.XtraPrintingLinks
Imports DevExpress.XtraCharts.Native


Public Class PrintForecast
    Inherits System.Web.UI.Page
#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_KanbanDate As String
    Dim ls_approve As Boolean

    Private grid_Renamed As ASPxGridView
    Private mergedCells As New Dictionary(Of GridViewDataColumn, TableCell)()
    Private cellRowSpans As New Dictionary(Of TableCell, Integer)()

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtHeader As DataTable
    Dim dtHeader2 As DataTable
    Dim dtDetail As DataTable

#End Region

#Region "CONTROL EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Or (Not String.IsNullOrEmpty(Request.QueryString("id2"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            Session("M01Url") = Request.QueryString("Session")
            Session("MenuDesc") = "PRINT FORECAST"


            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Clear()
                dtPeriodFrom.Text = Format(Now, "yyyy-MM")
                Call up_fillcombo()
                Call up_GridHeader()
                'Call up_GridLoad()
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

    'Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
    '    Dim ls_SQL As String = "", ls_MsgID As String = ""
    '    Dim ls_Active As String = "", i As Long = 1
    '    Dim ls_Supplier As String = ""
    '    Dim ls_deliveryLocation As String = ""
    '    Dim ls_Affiliate As String = ""
    '    Dim ls_KanbanNo As String = ""

    '    Dim ls_filter As String = ""

    '    Session.Remove("eFilter")

    '    With grid
    '        For i = 0 To e.UpdateValues.Count - 1
    '            If (e.UpdateValues(i).NewValues("cols").ToString()) = 1 Then
    '                If ls_filter = "" Then
    '                    ls_filter = "'" + Trim(e.UpdateValues(i).NewValues("OrderNo").ToString()) + Trim(e.UpdateValues(i).NewValues("AffiliateID").ToString()) + Trim(e.UpdateValues(i).NewValues("SupplierID").ToString()) + "'"
    '                Else
    '                    ls_filter = ls_filter + ",'" + Trim(e.UpdateValues(i).NewValues("OrderNo").ToString()) + Trim(e.UpdateValues(i).NewValues("AffiliateID").ToString()) + Trim(e.UpdateValues(i).NewValues("SupplierID").ToString()) + "'"
    '                End If
    '            End If
    '        Next

    '        Session("eFilter") = ls_filter
    '    End With
    'End Sub

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
                    DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/PurchaseOrderExport/LabelViewReport.aspx")
            End Select

EndProcedure:
            Session("AA220Msg") = ""
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    'Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
    '    'add the attribute that will be used to find which column the cell belongs to
    '    e.Cell.Attributes.Add("no", e.DataColumn.VisibleIndex.ToString())

    '    If cellRowSpans.ContainsKey(e.Cell) Then
    '        e.Cell.RowSpan = cellRowSpans(e.Cell)
    '    End If
    'End Sub

    'Protected Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As ASPxGridViewTableDataCellEventArgs)
    '    If e.DataColumn.FieldName <> "no" Then
    '        Return
    '    End If
    '    Dim position As String = Convert.ToString(e.CellValue)
    '    'Dim positionIcon As ASPxImage = CType(grid.FindRowCellTemplateControl(e.VisibleIndex, e.DataColumn, "PositionIcon"), ASPxImage)
    '    'positionIcon.Caption = position
    '    'positionIcon.EmptyImage.IconID = GetIconIDByPosition(position)
    'End Sub


    'Protected Sub OnDataBound(ByVal sender As Object, ByVal e As EventArgs)
    '    For i As Integer = grid.VisibleRowCount - 1 To 1 Step -1
    '        Dim row As ASPxGridView = grid.GetRowValues(i)
    '        Dim previousRow As ASPxGridView = grid.GetRowValues(i - 1)
    '        For j As Integer = 0 To grid.VisibleRowCount - 1
    '            If row.GetRow(j).Text = previousRow.GetRow(j).Text Then
    '                If previousRow.GetRow(j).RowSpan = 0 Then
    '                    If row.GetRow(j).RowSpan = 0 Then
    '                        previousRow.GetRow(j).RowSpan += 2
    '                    Else
    '                        previousRow.GetRow(j).RowSpan = row.GetRow(j).RowSpan + 1
    '                    End If
    '                    row.GetRow(j).Visible = False
    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub

    'Protected Sub OnDataBound(ByVal sender As Object, ByVal e As EventArgs)
    '    For i As Integer = grid.VisibleRowCount - 1 To 1 Step -1
    '        Dim row As  = grid.GetRowValues(i)
    '        Dim previousRow As ASPxGridView = grid.GetRowValues(i - 1)
    '        For j As Integer = 0 To grid.VisibleRowCount - 1
    '            If row.GetRowValues(j).Text = previousRow.GetRowValues(j).Text Then
    '                If previousRow.GetRowValues(j).RowSpan = 0 Then
    '                    If row.GetRowValues(j).RowSpan = 0 Then
    '                        previousRow.GetRowValues(j).RowSpan += 2
    '                    Else
    '                        previousRow.GetRowValues(j).RowSpan = row.GetRowValues(j).RowSpan + 1
    '                    End If
    '                    row.GetRowValues(j).Visible = False
    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub
    'Private Sub grid_HtmlRowCreated(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles grid.HtmlRowCreated
    '    If grid.GetRowLevel(e.VisibleIndex) = 0 Then
    '        If grid.GetRowLevel(e.VisibleIndex) <> grid.GroupCount Then
    '            Return
    '        End If
    '        For i As Integer = e.Row.Cells.Count - 1 To 0 Step -1
    '            Dim dataCell As DevExpress.Web.ASPxGridView.Rendering.GridViewTableDataCell = TryCast(e.Row.Cells(i), DevExpress.Web.ASPxGridView.Rendering.GridViewTableDataCell)
    '            If dataCell IsNot Nothing Then
    '                MergeCells(dataCell.DataColumn, e.VisibleIndex, dataCell)
    '            End If
    '        Next i
    '    End If
    'End Sub
#End Region

#Region "PROCEDURE"

    'Private Sub MergeCells(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer, ByVal cell As TableCell)
    '    Dim isNextTheSame As Boolean = IsNextRowHasSameData(column, visibleIndex)
    '    If isNextTheSame Then
    '        If Not mergedCells.ContainsKey(column) Then
    '            mergedCells(column) = cell
    '        End If
    '    End If
    '    If IsPrevRowHasSameData(column, visibleIndex) Then
    '        CType(cell.Parent, TableRow).Cells.Remove(cell)
    '        If mergedCells.ContainsKey(column) Then
    '            Dim mergedCell As TableCell = mergedCells(column)
    '            If Not cellRowSpans.ContainsKey(mergedCell) Then
    '                cellRowSpans(mergedCell) = 1
    '            End If
    '            cellRowSpans(mergedCell) = cellRowSpans(mergedCell) + 1
    '        End If
    '    End If
    '    If Not isNextTheSame Then
    '        mergedCells.Remove(column)
    '    End If
    'End Sub

    Private Function IsNextRowHasSameData(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer) As Boolean
        'is it the last visible row
        If visibleIndex >= grid.VisibleRowCount - 1 Then
            Return False
        End If

        Return IsSameData(column.FieldName, visibleIndex, visibleIndex + 1)
    End Function

    Private Function IsPrevRowHasSameData(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer) As Boolean

        Dim grid_Renamed As ASPxGridView = column.Grid
        'is it the first visible row
        If visibleIndex <= Grid.VisibleStartIndex Then
            Return False
        End If

        Return IsSameData(column.FieldName, visibleIndex, visibleIndex - 1)
    End Function

    Private Function IsSameData(ByVal fieldName As String, ByVal visibleIndex1 As Integer, ByVal visibleIndex2 As Integer) As Boolean
        ' is it a group row?
        If Grid.GetRowLevel(visibleIndex2) <> Grid.GroupCount Then
            Return False
        End If

        Return Object.Equals(Grid.GetRowValues(visibleIndex1, fieldName), Grid.GetRowValues(visibleIndex2, fieldName))
    End Function

    Private Sub colorGrid()
        'grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.White
    End Sub

    Private Sub up_fillcombo()
        Dim ls_sql As String

        'Affiliate Code
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(Affiliatename) FROM MS_Affiliate " & vbCrLf
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
        cboaffiliate.Text = clsGlobal.gs_All
        txtaffiliate.Text = clsGlobal.gs_All
        lblerrmessage.Text = ""
    End Sub

    Private Sub up_GridHeader()

        With grid
            .Columns.Clear()

            'ACCOUNT CODE
            Dim col As New GridViewDataTextColumn
            col.Caption = "NO"
            col.FieldName = "no"
            col.Index = 0
            col.VisibleIndex = 0
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(30)
            .Columns.Add(col)

            col = New GridViewDataTextColumn
            col.Caption = "Part No."
            col.FieldName = "HPartNo"
            col.Index = 1
            col.VisibleIndex = 1
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(120)
            .Columns.Add(col)

            col = New GridViewDataTextColumn
            col.Caption = "Part Name"
            col.FieldName = "PartName"
            col.Index = 2
            col.VisibleIndex = 2
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(200)
            .Columns.Add(col)

            col = New GridViewDataTextColumn
            col.Caption = "Period"
            col.FieldName = "Period"
            col.Index = 3
            col.VisibleIndex = 3
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(90)
            .Columns.Add(col)

            col = New GridViewDataTextColumn
            col.Caption = " "
            col.FieldName = "caption"
            col.Index = 4
            col.VisibleIndex = 4
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(170)
            .Columns.Add(col)

            'F1
            col = New GridViewDataTextColumn
            col.Caption = Format(dtPeriodFrom.Value, "MMM")
            col.FieldName = "F1"
            col.Index = 5
            col.VisibleIndex = 5
            col.CellStyle.HorizontalAlign = HorizontalAlign.Right
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(90)
            .Columns.Add(col)

            'F2
            col = New GridViewDataTextColumn
            'col.Caption = Format(DateDiff(DateInterval.Month, 1, dtPeriodFrom.Value), "MMM")
            col.Caption = Format(DateAdd("m", 1, dtPeriodFrom.Value), "MMM")
            col.FieldName = "F2"
            col.Index = 6
            col.VisibleIndex = 6
            col.CellStyle.HorizontalAlign = HorizontalAlign.Right
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(90)
            .Columns.Add(col)

            'F3
            col = New GridViewDataTextColumn
            'col.Caption = Format(DateDiff(DateInterval.Month, 2, dtPeriodFrom.Value), "MMM")
            col.Caption = Format(DateAdd("m", 2, dtPeriodFrom.Value), "MMM")
            col.FieldName = "F3"
            col.Index = 7
            col.VisibleIndex = 7
            col.CellStyle.HorizontalAlign = HorizontalAlign.Right
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(90)
            .Columns.Add(col)

            'F4
            col = New GridViewDataTextColumn
            'col.Caption = Format(DateDiff(DateInterval.Month, 3, dtPeriodFrom.Value), "MMM")
            col.Caption = Format(DateAdd("m", 3, dtPeriodFrom.Value), "MMM")
            col.FieldName = "F4"
            col.Index = 8
            col.VisibleIndex = 8
            col.CellStyle.HorizontalAlign = HorizontalAlign.Right
            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
            col.Width = Unit.Pixel(90)
            .Columns.Add(col)
        End With

    End Sub

    Private Sub up_GridLoad_OLD()
        Dim ls_SQL As String = ""

        Dim ls_Affiliate As String = ""
        Dim ls_period As String = ""
        Dim ls_Period1 As String = ""
        Dim ls_period2 As String = ""
        Dim ls_Period3 As String = ""
        Dim ls_Period4 As String = ""

        'cboaffiliate.Text = "JAI"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_period = Format(dtPeriodFrom.Value, "yyyy-MM") & "-01"
            ls_Period1 = Format(DateAdd("m", -1, dtPeriodFrom.Value), "yyy-MM") & "-01"
            ls_period2 = Format(DateAdd("m", 1, dtPeriodFrom.Value), "yyy-MM") & "-01"
            ls_Period3 = Format(DateAdd("m", 2, dtPeriodFrom.Value), "yyy-MM") & "-01"
            ls_Period4 = Format(DateAdd("m", 3, dtPeriodFrom.Value), "yyy-MM") & "-01"

            ls_Affiliate = Trim(cboaffiliate.Text)

            ls_SQL = ls_SQL + " if exists ( select Pono from po_master_export where Period = '" & ls_period & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + " UNION ALL select Pono from po_master where Period = '" & ls_period & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + ") BEGIN" & vbCrLf
            ls_SQL = ls_SQL + " --CONTOH UNTUK BULAN 3 " & vbCrLf & _
                     "  Select no,URUT,AffiliateID,HPartno,PartName,Period,caption,F1,F2,F3,F4,PartNo from( " & vbCrLf & _
                     " -- XXX.TOTAL FORECAST -- " & vbCrLf & _
                     " select no,URUT, AffiliateID, HPartNo, PartName, Period = Right(Convert(char(11), convert(datetime,Period),106), 8), caption," & vbCrLf & _
                     " F1 = Case when caption like '%Diff%' then  Replace(F1,'.00','') + '%' else Replace(F1,'.00','') END ," & vbCrLf & _
                     " F2 = Case when caption like '%Diff%' then  Replace(F2,'.00','') + '%' else Replace(F2,'.00','') END ," & vbCrLf & _
                     " F3 = Case when caption like '%Diff%' then  Replace(F3,'.00','') + '%' else Replace(F3,'.00','') END ," & vbCrLf & _
                     " F4 = Case when caption like '%Diff%' then  Replace(F4,'.00','') + '%' else Replace(F4,'.00','') END ," & vbCrLf & _
                     " PartNo from ( " & vbCrLf & _
                     " Select * from ( " & vbCrLf & _
                     " Select " & vbCrLf & _
                     " Convert(char,ROW_NUMBER() over (order by x.PartNo,Period)) as no," & vbCrLf & _
                     " URUT = 1, " & vbCrLf & _
                     " AffiliateID = AffiliateID, " & vbCrLf & _
                     " PartNo = x.PartNo, " & vbCrLf & _
                     " PartName = MP.PartName, " & vbCrLf & _
                     " Period = Right(Convert(char(11), convert(datetime,Period),106), 8),  " & vbCrLf & _
                     " caption = Caption, " & vbCrLf

            ls_SQL = ls_SQL + " F1 = SUM(F1), " & vbCrLf & _
                              " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf & _
                              " F4 = SUM(F4), HPartNo = x.PartNo  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	select distinct " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf

            ls_SQL = ls_SQL + " 		F1 = isnull(POD2.Forecast1,0), " & vbCrLf & _
                              " 		F2 = isnull(POD.ForecastN1,0), " & vbCrLf & _
                              " 		F3 = isnull(ForecastN2,0), " & vbCrLf & _
                              " 		F4 = isnull(ForecastN3,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN (Select distinct Forecast1 =isnull(ForecastN1,0), a.AffiliateID, PartNo  " & vbCrLf & _
                              " 				from PO_Detail b with(nolock) left join PO_master a with(nolock) ON a.AffiliateID = b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 				and a.SupplierID = b.SupplierID and a.POno = b.Pono " & vbCrLf & _
                              " 				where Period = '" & ls_Period1 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and a.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 		AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select distinct " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = isnull(POD2.Forecast1,0), " & vbCrLf & _
                              " 		F2 = isnull(POD.Forecast1,0), " & vbCrLf & _
                              " 		F3 = isnull(Forecast2,0), " & vbCrLf & _
                              " 		F4 = isnull(Forecast3,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN (Select distinct Forecast1 = isnull(Forecast1,0), a.AffiliateID, PartNo  " & vbCrLf & _
                              " 				from po_detail_Export b with(nolock) left join PO_Master_Export a with(nolock) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				and a.SupplierID = b.SupplierID and a.POno = b.Pono and a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              " 				where Period = '" & ls_Period1 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and a.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 		AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " where Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName,x.caption " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " no = '', URUT = 1, " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = dateadd(month,1,Period), " & vbCrLf & _
                              " caption = Caption, " & vbCrLf & _
                              " F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf & _
                              " F4 = SUM(F4), HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F2 = isnull(POD.ForecastN1,0), " & vbCrLf & _
                              " 		F3 = isnull(ForecastN2,0), " & vbCrLf & _
                              " 		F4 = isnull(ForecastN3,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.Forecast1,0), " & vbCrLf & _
                              " 		F3 = isnull(Forecast2,0), " & vbCrLf & _
                              " 		F4 = isnull(Forecast3,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " where Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName,x.caption " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " no = '', URUT = 1, " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = dateadd(month,2,Period), " & vbCrLf & _
                              " caption = Caption, " & vbCrLf & _
                              " F1 = 0, " & vbCrLf & _
                              " F2 = 0, " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf

            ls_SQL = ls_SQL + " F4 = SUM(F4), HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.ForecastN1,0), " & vbCrLf & _
                              " 		F3 = isnull(ForecastN2,0), " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = isnull(ForecastN3,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(Forecast2,0), " & vbCrLf & _
                              " 		F4 = isnull(Forecast3,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " where Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName,x.caption " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select " & vbCrLf

            ls_SQL = ls_SQL + " no = '', URUT = 1, " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = dateadd(month,3,Period), " & vbCrLf & _
                              " caption = Caption, " & vbCrLf & _
                              " F1 = 0, " & vbCrLf & _
                              " F2 = 0, " & vbCrLf & _
                              " F3 = 0, " & vbCrLf & _
                              " F4 = SUM(F4), HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf

            ls_SQL = ls_SQL + " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(ForecastN3,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(Forecast3,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " where Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName,x.caption " & vbCrLf & _
                              " )x where Period = '" & ls_period & "' " & vbCrLf & _
                              " -- XXX.TOTAL PO -- " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select  " & vbCrLf

            ls_SQL = ls_SQL + " no = '', " & vbCrLf & _
                              " URUT = 2, " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = Period, " & vbCrLf & _
                              " caption = caption, " & vbCrLf & _
                              " F1 = SUM(F1), " & vbCrLf & _
                              " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf & _
                              " F4 = SUM(F4), HPartNo = ''  " & vbCrLf

            ls_SQL = ls_SQL + " FROM( " & vbCrLf & _
                              " --F1 " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F2 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F3 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf

            ls_SQL = ls_SQL + " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F4 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where x.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption " & vbCrLf & _
                              "  " & vbCrLf

            ls_SQL = ls_SQL + " -- XXX.TOTAL PO KANBAN -- " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " no = '', " & vbCrLf & _
                              " URUT = 3, " & vbCrLf & _
                              " AffiliateID = isnull(AffiliateID,''), " & vbCrLf & _
                              " PartNo = isnull(x.PartNo,''), " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = isnull(Period, '" & ls_period & "'), " & vbCrLf & _
                              " caption = caption, " & vbCrLf & _
                              " F1 = isnull(SUM(F1),0), " & vbCrLf

            ls_SQL = ls_SQL + " F2 = isnull(SUM(F2),0), " & vbCrLf & _
                              " F3 = isnull(SUM(F3),0), " & vbCrLf & _
                              " F4 = isnull(SUM(F4),0),HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " --F1 " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = PartNo, PartName = '', period = '" & ls_period & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = isnull(POD.POQty,0), " & vbCrLf

            ls_SQL = ls_SQL + " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "     AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F2 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_period2 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf

            ls_SQL = ls_SQL + " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F3 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_Period3 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F4 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_Period4 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf

            ls_SQL = ls_SQL + " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption " & vbCrLf & _
                              "  " & vbCrLf & _
                              " -- XXX.Total PO Delivery -- " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " no = '', " & vbCrLf & _
                              " URUT = 4, " & vbCrLf & _
                              " AffiliateID = isnull(AffiliateID,''), " & vbCrLf & _
                              " PartNo = isnull(x.PartNo,''), " & vbCrLf & _
                              " PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " Period = isnull(Period, '" & ls_period & "'), " & vbCrLf & _
                              " caption = caption, " & vbCrLf & _
                              " F1 = isnull(SUM(F1),0), " & vbCrLf & _
                              " F2 = isnull(SUM(F2),0), " & vbCrLf & _
                              " F3 = isnull(SUM(F3),0), " & vbCrLf & _
                              " F4 = isnull(SUM(F4),0), HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " --F1 " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = PartNo, PartName = '', period = '" & ls_period & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F2 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + " Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_period2 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F3 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_Period3 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(DSD.DOQty,0), " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F4 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_Period4 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO Delivery', " & vbCrLf

            ls_SQL = ls_SQL + " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 	ON DSD.PONo = POM.PONo " & vbCrLf

            ls_SQL = ls_SQL + " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO Delivery', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 	AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 	AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 	AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption " & vbCrLf & _
                              "  " & vbCrLf & _
                              " --XXX.DIFF PO vs FORECAST-- " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select  " & vbCrLf & _
                              " no = '', " & vbCrLf & _
                              " URUT = 5, " & vbCrLf & _
                              " AffiliateID = PO.AffiliateID, " & vbCrLf & _
                              " PartNo = PO.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = PO.Period, " & vbCrLf

            ls_SQL = ls_SQL + " caption = 'Diff PO vs Forecast', " & vbCrLf & _
                              " F1 = Case when PO.F1 = 0 or FR.F1 = 0 then 0 else PO.F1/FR.F1-1 END, " & vbCrLf & _
                              " F2 = Case when PO.F2 = 0 or FR.F2 = 0 then 0 else PO.F2/FR.F2-1 END, " & vbCrLf & _
                              " F3 = Case when PO.F3 = 0 or FR.F3 = 0 then 0 else PO.F3/FR.F3-1 END, " & vbCrLf & _
                              " F4 = Case when PO.F4 = 0 or FR.F4 = 0 then 0 else PO.F4/FR.F4-1 END, HPartNo = '' " & vbCrLf & _
                              " FROM " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = MP.PartName, " & vbCrLf

            ls_SQL = ls_SQL + " Period = Period, " & vbCrLf & _
                              " F1 = SUM(F1), " & vbCrLf & _
                              " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf & _
                              " F4 = SUM(F4)  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		F1 = isnull(POD2.Forecast1,0), " & vbCrLf & _
                              " 		F2 = isnull(POD.ForecastN1,0), " & vbCrLf & _
                              " 		F3 = isnull(ForecastN2,0), " & vbCrLf & _
                              " 		F4 = isnull(ForecastN3,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN (Select Forecast1 =isnull(ForecastN1,0), a.AffiliateID, PartNo  " & vbCrLf & _
                              " 				from PO_Detail b with(nolock) left join PO_master a with(nolock) ON a.AffiliateID = b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 				and a.SupplierID = b.SupplierID and a.POno = b.Pono " & vbCrLf & _
                              " 				where Period = '" & ls_Period1 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and a.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + ") POD2 " & vbCrLf & _
                              " 	ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 		AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select distinct " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 		Period = Period, " & vbCrLf & _
                              " 		F1 = isnull(POD2.Forecast1,0), " & vbCrLf & _
                              " 		F2 = isnull(POD.Forecast1,0), " & vbCrLf & _
                              " 		F3 = isnull(Forecast2,0), " & vbCrLf & _
                              " 		F4 = isnull(Forecast3,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN (Select distinct Forecast1 = isnull(Forecast1,0), a.AffiliateID, PartNo  " & vbCrLf

            ls_SQL = ls_SQL + " 				from po_detail_Export b with(nolock) left join PO_Master_Export a with(nolock) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				and a.SupplierID = b.SupplierID and a.POno = b.Pono and a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              " 				where Period = '" & ls_Period1 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and a.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + ") POD2 " & vbCrLf & _
                              " 	ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 		AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " where Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName " & vbCrLf & _
                              " ) FR " & vbCrLf & _
                              " LEFT JOIN ( " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = MP.PartName, " & vbCrLf & _
                              " Period = Period, " & vbCrLf & _
                              " F1 = SUM(F1), " & vbCrLf & _
                              " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf

            ls_SQL = ls_SQL + " F4 = SUM(F4)  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " --F1 " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		F1 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F2 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F3 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 		Period = POM.Period, " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F4 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName) PO " & vbCrLf & _
                              " ON FR.AffiliateID = PO.AffiliateID and FR.PartNo = PO.PartNo " & vbCrLf & _
                              "  " & vbCrLf & _
                              " --XXX.DIFF PO KANBAN vs PO " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select distinct " & vbCrLf & _
                              " no = '', " & vbCrLf & _
                              " URUT = 6, " & vbCrLf & _
                              " AffiliateID = PO.AffiliateID, " & vbCrLf & _
                              " PartNo = PO.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " PartName = '', " & vbCrLf & _
                              " Period = PO.Period, " & vbCrLf & _
                              " caption = 'Diff PO Kanban vs PO', " & vbCrLf & _
                              " F1 = Case when PO.F1 = 0 then 0 else POK.F1/PO.F1-1 END, " & vbCrLf & _
                              " F2 = Case when PO.F2 = 0 then 0 else POK.F2/PO.F2-1 END, " & vbCrLf & _
                              " F3 = Case when PO.F3 = 0 then 0 else POK.F3/PO.F3-1 END, " & vbCrLf & _
                              " F4 = Case when PO.F4 = 0 then 0 else POK.F4/PO.F4-1 END, HPartNo = ''  " & vbCrLf & _
                              " FROM " & vbCrLf & _
                              " ( " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " PartNo = x.PartNo, " & vbCrLf & _
                              " PartName = MP.PartName, " & vbCrLf & _
                              " Period = Period, " & vbCrLf & _
                              " F1 = SUM(F1), " & vbCrLf & _
                              " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf & _
                              " F4 = SUM(F4)  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " --F1 " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf

            ls_SQL = ls_SQL + " 		F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F2 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F3 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.Week1,0), " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F4 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf

            ls_SQL = ls_SQL + " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption " & vbCrLf & _
                              " ) PO " & vbCrLf & _
                              " LEFT JOIN ( " & vbCrLf & _
                              " Select " & vbCrLf & _
                              " URUT = 3, " & vbCrLf & _
                              " AffiliateID = isnull(AffiliateID,''), " & vbCrLf

            ls_SQL = ls_SQL + " PartNo = isnull(x.PartNo,''), " & vbCrLf & _
                              " PartName = isnull(MP.PartName,''), " & vbCrLf & _
                              " Period = isnull(Period, '" & ls_period & "'), " & vbCrLf & _
                              " caption = caption, " & vbCrLf & _
                              " F1 = isnull(SUM(F1),0), " & vbCrLf & _
                              " F2 = isnull(SUM(F2),0), " & vbCrLf & _
                              " F3 = isnull(SUM(F3),0), " & vbCrLf & _
                              " F4 = isnull(SUM(F4),0)  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " --F1 " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + "     PartNo = PartNo, PartName = '', period = '" & ls_period & "',caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F2 " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " 	Select " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " PartNo = partNo, PartName = '', period = '" & ls_period2 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf

            ls_SQL = ls_SQL + " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F3 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_Period3 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf

            ls_SQL = ls_SQL + " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 		F3 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 		F4 = 0 " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " --F4 " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	PartNo = partNo, PartName = '', period = '" & ls_Period4 & "',  " & vbCrLf & _
                              " 		   caption = 'Total PO Kanban', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 	From po_detail POD " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf

            ls_SQL = ls_SQL + " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	AND KanbanCls = '1' " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total PO KANBAN', " & vbCrLf & _
                              " 		F1 = 0, " & vbCrLf & _
                              " 		F2 = 0, " & vbCrLf & _
                              " 		F3 = 0, " & vbCrLf & _
                              " 		F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 	where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption) POK " & vbCrLf & _
                              " ON PO.AffiliateID = POK.AffiliateID and PO.PartNo = POK.PartNo and PO.Period = POK.Period " & vbCrLf & _
                              "  " & vbCrLf & _
                              " --XXX.DIFF DELIVERY vs PO " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + " Select distinct " & vbCrLf & _
                              " no = '', " & vbCrLf & _
                              " URUT = 7, " & vbCrLf & _
                              " AffiliateID = PO.AffiliateID, " & vbCrLf & _
                              " PartNo = PO.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = PO.Period, " & vbCrLf & _
                              " caption = 'Diff Delivery vs PO', " & vbCrLf & _
                              " F1 = Case when PO.F1 = 0 then 0 else POD.F1/PO.F1-1 END, " & vbCrLf & _
                              " F2 = Case when PO.F2 = 0 then 0 else POD.F2/PO.F2-1 END, " & vbCrLf & _
                              " F3 = Case when PO.F3 = 0 then 0 else POD.F3/PO.F3-1 END, " & vbCrLf

            ls_SQL = ls_SQL + " F4 = Case when PO.F4 = 0 then 0 else POD.F4/PO.F4-1 END, HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	(Select " & vbCrLf & _
                              " 	AffiliateID = isnull(AffiliateID,''), " & vbCrLf & _
                              " 	PartNo = isnull(x.PartNo,''), " & vbCrLf & _
                              " 	PartName = isnull(MP.PartName,''), " & vbCrLf & _
                              " 	Period = isnull(Period, '" & ls_period & "'), " & vbCrLf & _
                              " 	F1 = isnull(SUM(F1),0), " & vbCrLf & _
                              " 	F2 = isnull(SUM(F2),0), " & vbCrLf & _
                              " 	F3 = isnull(SUM(F3),0), " & vbCrLf & _
                              " 	F4 = isnull(SUM(F4),0)  " & vbCrLf

            ls_SQL = ls_SQL + " 	FROM( " & vbCrLf & _
                              " 	--F1 " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = Isnull(POM.AffiliateID,''), " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = PartNo, PartName = '', period = '" & ls_period & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F2 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = Isnull(POM.AffiliateID,''), " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_period2 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf

            ls_SQL = ls_SQL + " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F3 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = Isnull(POM.AffiliateID,''), " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_Period3 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F4 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = Isnull(POM.AffiliateID,''), " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_Period4 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf

            ls_SQL = ls_SQL + " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)x " & vbCrLf & _
                              " 	LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where x.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption ) POD " & vbCrLf & _
                              " 	LEFT JOIN ( " & vbCrLf & _
                              " 	Select  " & vbCrLf & _
                              " 	URUT = 2, " & vbCrLf & _
                              " 	AffiliateID = AffiliateID, " & vbCrLf & _
                              " 	PartNo = x.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 	PartName = MP.PartName, " & vbCrLf & _
                              " 	Period = Period, " & vbCrLf & _
                              " 	caption = caption, " & vbCrLf & _
                              " 	F1 = SUM(F1), " & vbCrLf & _
                              " 	F2 = SUM(F2), " & vbCrLf & _
                              " 	F3 = SUM(F3), " & vbCrLf & _
                              " 	F4 = SUM(F4)  " & vbCrLf & _
                              " 	FROM( " & vbCrLf & _
                              " 	--F1 " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf

            ls_SQL = ls_SQL + " 			F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F2 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F3 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf

            ls_SQL = ls_SQL + " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(POD.Week1,0), " & vbCrLf

            ls_SQL = ls_SQL + " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F4 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf

            ls_SQL = ls_SQL + " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)x " & vbCrLf & _
                              " 	LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " 	" & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption " & vbCrLf & _
                              " 	) PO " & vbCrLf & _
                              " ON PO.AffiliateID = POD.AffiliateID and PO.PartNo = POD.PartNo and PO.Period = POD.Period) " & vbCrLf & _
                              "  " & vbCrLf & _
                              " --XXX.DIFF PO DELIVERY vs FORECAST " & vbCrLf & _
                              " UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + " Select distinct " & vbCrLf & _
                              " no = '', " & vbCrLf & _
                              " URUT = 8, " & vbCrLf & _
                              " AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " PartNo = POD.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = POD.Period, " & vbCrLf & _
                              " caption = 'Diff Delivery vs Forecast', " & vbCrLf & _
                              " F1 = Case when POD.F1 = 0 or FR.F1 = 0 then 0 else POD.F1/FR.F1-1 END, " & vbCrLf & _
                              " F2 = Case when POD.F2 = 0 or FR.F2 = 0 then 0 else POD.F2/FR.F2-1 END, " & vbCrLf & _
                              " F3 = Case when POD.F3 = 0 or FR.F3 = 0 then 0 else POD.F3/FR.F3-1 END, " & vbCrLf

            ls_SQL = ls_SQL + " F4 = Case when POD.F4 = 0 or FR.F4 = 0 then 0 else POD.F4/FR.F4-1 END, HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	(Select " & vbCrLf & _
                              " 	AffiliateID = isnull(AffiliateID,''), " & vbCrLf & _
                              " 	PartNo = isnull(x.PartNo,''), " & vbCrLf & _
                              " 	PartName = isnull(MP.PartName,''), " & vbCrLf & _
                              " 	Period = isnull(Period, '" & ls_period & "'), " & vbCrLf & _
                              " 	F1 = isnull(SUM(F1),0), " & vbCrLf & _
                              " 	F2 = isnull(SUM(F2),0), " & vbCrLf & _
                              " 	F3 = isnull(SUM(F3),0), " & vbCrLf & _
                              " 	F4 = isnull(SUM(F4),0)  " & vbCrLf

            ls_SQL = ls_SQL + " 	FROM( " & vbCrLf & _
                              " 	--F1 " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = PartNo, PartName = '', period = '" & ls_period & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F2 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_period2 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf

            ls_SQL = ls_SQL + " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F3 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_Period3 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F4 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_Period4 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf

            ls_SQL = ls_SQL + " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "UNION ALL" & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)x " & vbCrLf & _
                              " 	LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption ) POD " & vbCrLf & _
                              " LEFT JOIN ( " & vbCrLf & _
                              " 	Select " & vbCrLf & _
                              " URUT = 1, " & vbCrLf & _
                              " AffiliateID = AffiliateID, " & vbCrLf & _
                              " PartNo = x.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " PartName = MP.PartName, " & vbCrLf & _
                              " Period = Period, " & vbCrLf & _
                              " caption = Caption, " & vbCrLf & _
                              " F1 = SUM(F1), " & vbCrLf & _
                              " F2 = SUM(F2), " & vbCrLf & _
                              " F3 = SUM(F3), " & vbCrLf & _
                              " F4 = SUM(F4)  " & vbCrLf & _
                              " FROM( " & vbCrLf & _
                              " 	select  " & vbCrLf & _
                              " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 		PartName = '', " & vbCrLf & _
                              " 		Period = POM.Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = isnull(POD2.Forecast1,0), " & vbCrLf & _
                              " 		F2 = isnull(POD.ForecastN1,0), " & vbCrLf & _
                              " 		F3 = isnull(ForecastN2,0), " & vbCrLf & _
                              " 		F4 = isnull(ForecastN3,0) " & vbCrLf & _
                              " 	From po_detail POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN (Select Forecast1 =isnull(ForecastN1,0), a.AffiliateID, PartNo  " & vbCrLf & _
                              " 				from PO_Detail b with(nolock) left join PO_master a with(nolock) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				and a.SupplierID = b.SupplierID and a.POno = b.Pono " & vbCrLf & _
                              " 				where Period = '" & ls_Period1 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and a.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + ") POD2 " & vbCrLf & _
                              " 	ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 		AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            ls_SQL = ls_SQL + " 	UNION ALL " & vbCrLf & _
                              " 	select distinct " & vbCrLf

            ls_SQL = ls_SQL + " 		AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 		PartNo = POD.PartNo, " & vbCrLf & _
                              " 		PartName = '', " & vbCrLf & _
                              " 		Period = Period, " & vbCrLf & _
                              " 		caption = 'Total Forecast', " & vbCrLf & _
                              " 		F1 = isnull(POD2.Forecast1,0), " & vbCrLf & _
                              " 		F2 = isnull(POD.Forecast1,0), " & vbCrLf & _
                              " 		F3 = isnull(Forecast2,0), " & vbCrLf & _
                              " 		F4 = isnull(Forecast3,0) " & vbCrLf & _
                              " 	From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 	LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 	ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 	and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 	AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN (Select distinct Forecast1 = isnull(Forecast1,0), a.AffiliateID, PartNo  " & vbCrLf & _
                              " 				from po_detail_Export b with(nolock) left join PO_Master_Export a with(nolock) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " 				and a.SupplierID = b.SupplierID and a.POno = b.Pono and a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              " 				where Period = '" & ls_Period1 & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and a.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + ") POD2 " & vbCrLf & _
                              " 	ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 		AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              " 	where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )x " & vbCrLf & _
                              " LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              " where Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName,x.caption) FR " & vbCrLf & _
                              " ON FR.AffiliateID = POD.AffiliateID and FR.PartNo = POD.PartNo ) " & vbCrLf & _
                              "  " & vbCrLf & _
                              " --XXX.BALANCE PO " & vbCrLf & _
                              " UNION ALL " & vbCrLf & _
                              " Select distinct " & vbCrLf & _
                              " no = '', " & vbCrLf

            ls_SQL = ls_SQL + " URUT = 9, " & vbCrLf & _
                              " AffiliateID = PO.AffiliateID, " & vbCrLf & _
                              " PartNo = PO.PartNo, " & vbCrLf & _
                              " PartName = '', " & vbCrLf & _
                              " Period = PO.Period, " & vbCrLf & _
                              " caption = 'Diff Delivery vs PO', " & vbCrLf & _
                              " F1 = Case when PO.F1 = 0 then 0 else PO.F1-POD.F1 END, " & vbCrLf & _
                              " F2 = Case when PO.F2 = 0 then 0 else PO.F2-POD.F2 END, " & vbCrLf & _
                              " F3 = Case when PO.F3 = 0 then 0 else PO.F3-POD.F3 END, " & vbCrLf & _
                              " F4 = Case when PO.F4 = 0 then 0 else PO.F4-POD.F4 END, HPartNo = ''  " & vbCrLf & _
                              " FROM( " & vbCrLf

            ls_SQL = ls_SQL + " 	(Select " & vbCrLf & _
                              " 	AffiliateID = isnull(AffiliateID,''), " & vbCrLf & _
                              " 	PartNo = isnull(x.PartNo,''), " & vbCrLf & _
                              " 	PartName = isnull(MP.PartName,''), " & vbCrLf & _
                              " 	Period = isnull(Period, '" & ls_period & "'), " & vbCrLf & _
                              " 	F1 = isnull(SUM(F1),0), " & vbCrLf & _
                              " 	F2 = isnull(SUM(F2),0), " & vbCrLf & _
                              " 	F3 = isnull(SUM(F3),0), " & vbCrLf & _
                              " 	F4 = isnull(SUM(F4),0)  " & vbCrLf & _
                              " 	FROM( " & vbCrLf & _
                              " 	--F1 " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = PartNo, PartName = '', period = '" & ls_period & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf

            ls_SQL = ls_SQL + " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "UNION ALL" & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F2 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_period2 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(DSD.DOQty,0), " & vbCrLf

            ls_SQL = ls_SQL + " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F3 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_Period3 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf

            ls_SQL = ls_SQL + " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf

            ls_SQL = ls_SQL + " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(DSD.DOQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F4 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "Select AffiliateID = '" & ls_Affiliate & "', " & vbCrLf
            Else
                ls_SQL = ls_SQL + "Select AffiliateID = '', " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		PartNo = partNo, PartName = '', period = '" & ls_Period4 & "',  " & vbCrLf & _
                              " 			   caption = 'Total PO Delivery', F1 = 0, F2 = 0, F3 = 0, F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO Delivery', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(DSD.DOQty,0) " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		LEFT JOIN DOSUPPLIER_Detail_Export DSD with(nolock) " & vbCrLf & _
                              " 		ON DSD.PONo = POM.PONo " & vbCrLf & _
                              " 		AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              " 		AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              " 		AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)x " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + " where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption ) POD " & vbCrLf & _
                              " 	LEFT JOIN ( " & vbCrLf & _
                              " 	Select  " & vbCrLf & _
                              " 	URUT = 2, " & vbCrLf & _
                              " 	AffiliateID = AffiliateID, " & vbCrLf & _
                              " 	PartNo = x.PartNo, " & vbCrLf & _
                              " 	PartName = MP.PartName, " & vbCrLf & _
                              " 	Period = Period, " & vbCrLf

            ls_SQL = ls_SQL + " 	caption = caption, " & vbCrLf & _
                              " 	F1 = SUM(F1), " & vbCrLf & _
                              " 	F2 = SUM(F2), " & vbCrLf & _
                              " 	F3 = SUM(F3), " & vbCrLf & _
                              " 	F4 = SUM(F4)  " & vbCrLf & _
                              " 	FROM( " & vbCrLf & _
                              " 	--F1 " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F2 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf

            ls_SQL = ls_SQL + " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + " 		where POM.Period = '" & ls_period2 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F3 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(POD.POQty,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = isnull(POD.Week1,0), " & vbCrLf & _
                              " 			F4 = 0 " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf

            ls_SQL = ls_SQL + " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period3 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	--F4 " & vbCrLf & _
                              " 	UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + " 			PartName = '', " & vbCrLf & _
                              " 			Period = POM.Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf & _
                              " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(POD.POQty,0) " & vbCrLf & _
                              " 		From po_detail POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 		UNION ALL " & vbCrLf & _
                              " 		select  " & vbCrLf & _
                              " 			AffiliateID = POD.AffiliateID, " & vbCrLf & _
                              " 			PartNo = POD.PartNo, " & vbCrLf & _
                              " 			PartName = '', " & vbCrLf & _
                              " 			Period = Period, " & vbCrLf & _
                              " 			caption = 'Total PO', " & vbCrLf & _
                              " 			F1 = 0, " & vbCrLf

            ls_SQL = ls_SQL + " 			F2 = 0, " & vbCrLf & _
                              " 			F3 = 0, " & vbCrLf & _
                              " 			F4 = isnull(POD.Week1,0) " & vbCrLf & _
                              " 		From po_detail_Export POD with(nolock) " & vbCrLf & _
                              " 		LEFT JOIN PO_Master_Export POM with(nolock) " & vbCrLf & _
                              " 		ON POM.Pono = POD.PONO " & vbCrLf & _
                              " 		and POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 		AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 		where POM.Period = '" & ls_Period4 & "' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and POM.AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	)x " & vbCrLf & _
                              " 	LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " 	Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName, x.caption " & vbCrLf & _
                              " 	) PO " & vbCrLf & _
                              " ON PO.AffiliateID = POD.AffiliateID and PO.PartNo = POD.PartNo and PO.Period = POD.Period) " & vbCrLf & _
                              " )ccc " & vbCrLf & _
                              " --ORDER BY partNo,Convert(char(10), convert(datetime,Period),112),urut " & vbCrLf & _
                              " )x order by PartNo,Period,AffiliateID,Urut,No  " & vbCrLf

            ls_SQL = ls_SQL + " END ELSE BEGIN " & vbCrLf & _
                              " select no = '',URUT = '', AffiliateID = '', HPartNo = '', PartName = '',  " & vbCrLf & _
                              " Period = '', caption = '', F1 = '',F2 = '',F3 = '',F4 = '',PartNo = '' from   " & vbCrLf & _
                              " po_master_export where Period = '" & ls_period & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + " UNION ALL " & vbCrLf & _
                              " select no = '',URUT = '', AffiliateID = '', HPartNo = '', PartName = '',  " & vbCrLf & _
                              " Period = '', caption = '', F1 = '',F2 = '',F3 = '',F4 = '',PartNo = '' from   " & vbCrLf & _
                              " po_master where Period = '" & ls_period & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "And AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "                End " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
            End With
            dtHeader = ds.Tables(0)
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""

        Dim ls_Affiliate As String = ""
        Dim ls_period As String = ""
        Dim ls_Period1 As String = ""
        Dim ls_period2 As String = ""
        Dim ls_Period3 As String = ""
        Dim ls_Period4 As String = ""

        'cboaffiliate.Text = "JAI"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_period = Format(dtPeriodFrom.Value, "yyyy-MM") & "-01"
            ls_Period1 = Format(DateAdd("m", -1, dtPeriodFrom.Value), "yyy-MM") & "-01"
            ls_period2 = Format(DateAdd("m", 1, dtPeriodFrom.Value), "yyy-MM") & "-01"
            ls_Period3 = Format(DateAdd("m", 2, dtPeriodFrom.Value), "yyy-MM") & "-01"
            ls_Period4 = Format(DateAdd("m", 3, dtPeriodFrom.Value), "yyy-MM") & "-01"

            ls_Affiliate = Trim(cboaffiliate.Text)
            ls_SQL = ls_SQL + " if exists ( select Pono from po_master_export WITH(NOLOCK) where Period = '" & ls_period & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + " UNION ALL select Pono from po_master WITH(NOLOCK) where Period = '" & ls_period & "' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "     )BEGIN  " & vbCrLf & _
                              "   --CONTOH UNTUK BULAN 3   " & vbCrLf & _
                              "         SELECT  no , " & vbCrLf & _
                              "                 URUT , " & vbCrLf

            ls_SQL = ls_SQL + "                 AffiliateID , " & vbCrLf & _
                              "                 PartNo , " & vbCrLf & _
                              "                 PartName , " & vbCrLf & _
                              "                 Period , " & vbCrLf & _
                              "                 caption , " & vbCrLf & _
                              "                 F1 , " & vbCrLf & _
                              "                 F2 , " & vbCrLf & _
                              "                 F3 , " & vbCrLf & _
                              "                 F4 , " & vbCrLf & _
                              "                 HPartNo " & vbCrLf & _
                              "         FROM    (  " & vbCrLf

            ls_SQL = ls_SQL + "   -- 1XXX.TOTAL FORECAST --   " & vbCrLf & _
                              "                   SELECT    no = no, " & vbCrLf & _
                              "                             URUT , " & vbCrLf & _
                              "                             AffiliateID , " & vbCrLf & _
                              "                             HPartNo , " & vbCrLf & _
                              "                             PartName , " & vbCrLf & _
                              "                             Period = RIGHT(CONVERT(CHAR(11), CONVERT(DATETIME, Period), 106), " & vbCrLf & _
                              "                                            8) , " & vbCrLf & _
                              "                             caption , " & vbCrLf & _
                              "                             F1 = CASE WHEN caption LIKE '%Diff%' " & vbCrLf & _
                              "                                       THEN REPLACE(F1, '.00', '') + '%' " & vbCrLf

            ls_SQL = ls_SQL + "                                       ELSE REPLACE(F1, '.00', '') " & vbCrLf & _
                              "                                  END , " & vbCrLf & _
                              "                             F2 = CASE WHEN caption LIKE '%Diff%' " & vbCrLf & _
                              "                                       THEN REPLACE(F2, '.00', '') + '%' " & vbCrLf & _
                              "                                       ELSE REPLACE(F2, '.00', '') " & vbCrLf & _
                              "                                  END , " & vbCrLf & _
                              "                             F3 = CASE WHEN caption LIKE '%Diff%' " & vbCrLf & _
                              "                                       THEN REPLACE(F3, '.00', '') + '%' " & vbCrLf & _
                              "                                       ELSE REPLACE(F3, '.00', '') " & vbCrLf & _
                              "                                  END , " & vbCrLf & _
                              "                             F4 = CASE WHEN caption LIKE '%Diff%' " & vbCrLf

            ls_SQL = ls_SQL + "                                       THEN REPLACE(F4, '.00', '') + '%' " & vbCrLf & _
                              "                                       ELSE REPLACE(F4, '.00', '') " & vbCrLf & _
                              "                                  END , " & vbCrLf & _
                              "                             PartNo " & vbCrLf & _
                              "                   FROM      (   " & vbCrLf & _
                              "                         SELECT    no = no ,  " & vbCrLf & _
                              "                             URUT , AffiliateID , HPartNo , PartName , Period,  " & vbCrLf & _
                              "                             caption , F1, F2, F3, F4, PartNo FROM ( " & vbCrLf & _
                              "   --FORECAST PERIOD  " & vbCrLf & _
                              "                               SELECT    CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY x.PartNo, Period )) AS no , " & vbCrLf & _
                              "                                         URUT = 1 , " & vbCrLf & _
                              "                                         AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                         PartNo = x.PartNo , " & vbCrLf & _
                              "                                         PartName = MP.PartName , " & vbCrLf

            ls_SQL = ls_SQL + "                                         Period = RIGHT(CONVERT(CHAR(11), CONVERT(DATETIME, Period), 106), " & vbCrLf & _
                              "                                                        8) , " & vbCrLf & _
                              "                                         caption = Caption , " & vbCrLf & _
                              "                                         F1 = SUM(F1) , " & vbCrLf & _
                              "                                         F2 = SUM(F2) , " & vbCrLf & _
                              "                                         F3 = SUM(F3) , " & vbCrLf & _
                              "                                         F4 = SUM(F4) , " & vbCrLf & _
                              "                                         HPartNo = x.PartNo " & vbCrLf & _
                              "                               FROM      ( SELECT DISTINCT " & vbCrLf & _
                              "                                                     AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                     PartNo = POD.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     PartName = '' , " & vbCrLf & _
                              "                                                     Period = POM.Period , " & vbCrLf & _
                              "                                                     caption = 'Total Forecast' , " & vbCrLf & _
                              "                                                     F1 = ISNULL(POD2.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                     F2 = ISNULL(POD.ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                     F3 = ISNULL(ForecastN2, 0) , " & vbCrLf & _
                              "                                                     F4 = ISNULL(ForecastN3, 0) " & vbCrLf & _
                              "                                           FROM      po_detail POD WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                     LEFT JOIN PO_Master POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                     WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                     LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                     PartNo = POD.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     PartName = '' , " & vbCrLf & _
                              "                                                     Period = Period , " & vbCrLf & _
                              "                                                     caption = 'Total Forecast' , " & vbCrLf & _
                              "                                                     F1 = ISNULL(POD2.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                     F2 = ISNULL(POD.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                     F3 = ISNULL(Forecast2, 0) , " & vbCrLf & _
                              "                                                     F4 = ISNULL(Forecast3, 0) " & vbCrLf & _
                              "                                           FROM      po_detail_Export POD WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                     LEFT JOIN PO_Master_Export POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                     WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                     LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                         ) x " & vbCrLf & _
                              "                                         LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                               WHERE     Period = '" & ls_period & "' " & vbCrLf & _
                              "                                         --AND AffiliateID = 'HESTO' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "                                         --AND x.PartNo = '7154-0708-30' " & vbCrLf & _
                              "                               GROUP BY  x.AffiliateID , " & vbCrLf & _
                              "                                         x.PartNo , " & vbCrLf & _
                              "                                         x.Period , " & vbCrLf & _
                              "                                         MP.PartName , " & vbCrLf & _
                              "                                         x.caption " & vbCrLf & _
                              "                               UNION ALL --FORECAST +1 " & vbCrLf & _
                              "   --FORECAST PERIOD  " & vbCrLf & _
                              "                               SELECT    no='' , " & vbCrLf

            ls_SQL = ls_SQL + "                                         URUT = 1 , " & vbCrLf & _
                              "                                         AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                         PartNo = x.PartNo , " & vbCrLf & _
                              "                                         PartName = '' , " & vbCrLf & _
                              "                                         Period = DATEADD(month, 1, Period) , " & vbCrLf & _
                              "                                         caption = Caption , " & vbCrLf & _
                              "                                         F1 = SUM(F1) , " & vbCrLf & _
                              "                                         F2 = SUM(F2) , " & vbCrLf & _
                              "                                         F3 = SUM(F3) , " & vbCrLf & _
                              "                                         F4 = SUM(F4) , " & vbCrLf & _
                              "                                         HPartNo = '' " & vbCrLf

            ls_SQL = ls_SQL + "                               FROM      ( SELECT DISTINCT " & vbCrLf & _
                              "                                                     AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                     PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                     PartName = '' , " & vbCrLf & _
                              "                                                     Period = POM.Period , " & vbCrLf & _
                              "                                                     caption = 'Total Forecast' , " & vbCrLf & _
                              "                                                     F1 = 0 , " & vbCrLf & _
                              "                                                     F2 = ISNULL(POD.ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                     F3 = ISNULL(ForecastN2, 0) , " & vbCrLf & _
                              "                                                     F4 = ISNULL(ForecastN3, 0) " & vbCrLf

            ls_SQL = ls_SQL + "                                           FROM      po_detail POD WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                     LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                     WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                     LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                     PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                     PartName = '' , " & vbCrLf & _
                              "                                                     Period = Period , " & vbCrLf & _
                              "                                                     caption = 'Total Forecast' , " & vbCrLf & _
                              "                                                     F1 = 0 , " & vbCrLf & _
                              "                                                     F2 = ISNULL(POD.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                     F3 = ISNULL(Forecast2, 0) , " & vbCrLf & _
                              "                                                     F4 = ISNULL(Forecast3, 0) " & vbCrLf & _
                              "                                           FROM      po_detail_Export POD WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                     LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                     WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                     LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                         ) x " & vbCrLf

            ls_SQL = ls_SQL + "                                         LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                               WHERE     Period = '" & ls_period & "' " & vbCrLf & _
                              "                                         --AND AffiliateID = 'HESTO' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "                                         --AND x.PartNo = '7154-0708-30' " & vbCrLf & _
                              "                               GROUP BY  x.AffiliateID , " & vbCrLf & _
                              "                                         x.PartNo , " & vbCrLf & _
                              "                                         x.Period , " & vbCrLf & _
                              "                                         MP.PartName , " & vbCrLf & _
                              "                                         x.caption " & vbCrLf & _
                              "                               UNION ALL --FORECAST +2 " & vbCrLf & _
                              "   --FORECAST PERIOD  " & vbCrLf

            ls_SQL = ls_SQL + "                               SELECT    no='' , " & vbCrLf & _
                              "                                         URUT = 1 , " & vbCrLf & _
                              "                                         AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                         PartNo = x.PartNo , " & vbCrLf & _
                              "                                         PartName = '' , " & vbCrLf & _
                              "                                         Period = DATEADD(month, 2, Period) , " & vbCrLf & _
                              "                                         caption = Caption , " & vbCrLf & _
                              "                                         F1 = SUM(F1) , " & vbCrLf & _
                              "                                         F2 = SUM(F2) , " & vbCrLf & _
                              "                                         F3 = SUM(F3) , " & vbCrLf & _
                              "                                         F4 = SUM(F4) , " & vbCrLf

            ls_SQL = ls_SQL + "                                         HPartNo = '' " & vbCrLf & _
                              "                               FROM      ( SELECT DISTINCT " & vbCrLf & _
                              "                                                     AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                     PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                     PartName = '' , " & vbCrLf & _
                              "                                                     Period = POM.Period , " & vbCrLf & _
                              "                                                     caption = 'Total Forecast' , " & vbCrLf & _
                              "                                                     F1 = 0 , " & vbCrLf & _
                              "                                                     F2 = 0 , " & vbCrLf & _
                              "                                                     F3 = ISNULL(ForecastN2, 0) , " & vbCrLf & _
                              "                                                     F4 = ISNULL(ForecastN3, 0) " & vbCrLf

            ls_SQL = ls_SQL + "                                           FROM      po_detail POD WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                     LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                     WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                     LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                     PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                     PartName = '' , " & vbCrLf & _
                              "                                                     Period = Period , " & vbCrLf & _
                              "                                                     caption = 'Total Forecast' , " & vbCrLf & _
                              "                                                     F1 = 0 , " & vbCrLf & _
                              "                                                     F2 = 0 , " & vbCrLf & _
                              "                                                     F3 = ISNULL(Forecast2, 0) , " & vbCrLf & _
                              "                                                     F4 = ISNULL(Forecast3, 0) " & vbCrLf & _
                              "                                           FROM      po_detail_Export POD WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                     LEFT JOIN PO_Master_Export POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                     WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                     LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                         ) x " & vbCrLf & _
                              "                                         LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                               WHERE     Period = '" & ls_period & "' " & vbCrLf & _
                              "                                         --AND AffiliateID = 'HESTO' " & vbCrLf

            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "and AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "                                         --AND x.PartNo = '7154-0708-30' " & vbCrLf & _
                              "                               GROUP BY  x.AffiliateID , " & vbCrLf & _
                              "                                         x.PartNo , " & vbCrLf & _
                              "                                         x.Period , " & vbCrLf & _
                              "                                         MP.PartName , " & vbCrLf & _
                              "                                         x.caption  " & vbCrLf & _
                              "   --UNION ALL --FORECAST +3 " & vbCrLf & _
                              "   ----FORECAST PERIOD  " & vbCrLf & _
                              "   --Select   " & vbCrLf

            ls_SQL = ls_SQL + "   --Convert(char,ROW_NUMBER() over (order by x.PartNo,Period)) as no,  " & vbCrLf & _
                              "   --URUT = 1,   " & vbCrLf & _
                              "   --AffiliateID = AffiliateID,   " & vbCrLf & _
                              "   --PartNo = x.PartNo,   " & vbCrLf & _
                              "   --PartName = MP.PartName,   " & vbCrLf & _
                              "   --Period = dateadd(month,3,Period),    " & vbCrLf & _
                              "   --caption = Caption,   " & vbCrLf & _
                              "   --F1 = SUM(F1),     " & vbCrLf & _
                              "   --F2 = SUM(F2),   " & vbCrLf & _
                              "   --F3 = SUM(F3),   " & vbCrLf & _
                              "   --F4 = SUM(F4), HPartNo = x.PartNo    " & vbCrLf

            ls_SQL = ls_SQL + "   --FROM(   " & vbCrLf & _
                              "   --		select distinct   " & vbCrLf & _
                              "   --			AffiliateID = POD.AffiliateID,   " & vbCrLf & _
                              "   --			PartNo = POD.PartNo,   " & vbCrLf & _
                              "   --			PartName = '',   " & vbCrLf & _
                              "   --			Period = POM.Period,   " & vbCrLf & _
                              "   --			caption = 'Total Forecast',   " & vbCrLf & _
                              "   --			F1 = 0,    		 " & vbCrLf & _
                              "   --			F2 = 0,   " & vbCrLf & _
                              "   --			F3 = 0,   " & vbCrLf & _
                              "   --			F4 = isnull(ForecastN3,0)   " & vbCrLf

            ls_SQL = ls_SQL + "   --		From po_detail POD with(nolock)   " & vbCrLf & _
                              "   --		LEFT JOIN PO_Master POM with(nolock)   " & vbCrLf & _
                              "   --		ON POM.Pono = POD.PONO   " & vbCrLf & _
                              "   --		and POM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "   --		AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "   --		LEFT JOIN (Select distinct Forecast1 =isnull(ForecastN1,0), a.AffiliateID, PartNo    " & vbCrLf & _
                              "   --					from PO_Detail b with(nolock) left join PO_master a with(nolock) ON a.AffiliateID = b.AffiliateID   " & vbCrLf & _
                              "   --					and a.SupplierID = b.SupplierID and a.POno = b.Pono    				 " & vbCrLf & _
                              "   --					where Period = '" & ls_Period1 & "'   " & vbCrLf & _
                              "   --		) POD2 ON POD2.AffiliateID = POD.AffiliateID   " & vbCrLf & _
                              "   --			AND POD2.PartNo = POD.PartNo   " & vbCrLf

            ls_SQL = ls_SQL + "   --		UNION ALL   " & vbCrLf & _
                              "   --		select distinct   " & vbCrLf & _
                              "   --			AffiliateID = POD.AffiliateID,   " & vbCrLf & _
                              "   --			PartNo = POD.PartNo,   " & vbCrLf & _
                              "   --			PartName = '',   " & vbCrLf & _
                              "   --			Period = Period,   " & vbCrLf & _
                              "   --			caption = 'Total Forecast',    		 " & vbCrLf & _
                              "   --			F1 = 0,   " & vbCrLf & _
                              "   --			F2 = 0,   " & vbCrLf & _
                              "   --			F3 = 0,   " & vbCrLf & _
                              "   --			F4 = isnull(Forecast3,0)   "

            ls_SQL = ls_SQL + "   --		From po_detail_Export POD with(nolock)   " & vbCrLf & _
                              "   --		LEFT JOIN PO_Master_Export POM with(nolock)   " & vbCrLf & _
                              "   --		ON POM.Pono = POD.PONO   " & vbCrLf & _
                              "   --		and POM.AffiliateID = POM.AffiliateID   " & vbCrLf & _
                              "   --		AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                              "   --		LEFT JOIN (Select distinct Forecast1 = isnull(Forecast1,0), a.AffiliateID, PartNo    " & vbCrLf & _
                              "   --					from po_detail_Export b with(nolock)  " & vbCrLf & _
                              "   --					left join PO_Master_Export a with(nolock) ON a.AffiliateID = b.AffiliateID    				 " & vbCrLf & _
                              "   --					and a.SupplierID = b.SupplierID and a.POno = b.Pono and a.OrderNo1 = b.OrderNo1   " & vbCrLf & _
                              "   --					where Period = '" & ls_Period1 & "'   " & vbCrLf & _
                              "   --		) POD2 ON POD2.AffiliateID = POD.AffiliateID   " & vbCrLf

            ls_SQL = ls_SQL + "   --			AND POD2.PartNo = POD.PartNo    " & vbCrLf & _
                              "   --)x   " & vbCrLf & _
                              "   --LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo   " & vbCrLf & _
                              "   -- where Period = '" & ls_period & "' and AffiliateID = 'HESTO' and x.PartNo = '7154-0708-30'  " & vbCrLf & _
                              "   --Group BY x.AffiliateID, x.PartNo, x.Period, MP.PartName,x.caption  " & vbCrLf & _
                              "     )Header " & vbCrLf & _
                              "                               UNION ALL " & vbCrLf & _
                              "                               SELECT    no = '' , " & vbCrLf & _
                              "                                         URUT , " & vbCrLf & _
                              "                                         AffiliateID , " & vbCrLf & _
                              "                                         HPartNo = '' , " & vbCrLf & _
                              "                                         PartName , "

            ls_SQL = ls_SQL + "                                         Period , " & vbCrLf & _
                              "                                         caption , " & vbCrLf & _
                              "                                         F1 , " & vbCrLf & _
                              "                                         F2 , " & vbCrLf & _
                              "                                         F3 , " & vbCrLf & _
                              "                                         F4 , " & vbCrLf & _
                              "                                         PartNo " & vbCrLf & _
                              "                               FROM      ( --DETAIL DATA " & vbCrLf & _
                              "   -- 2XXX.TOTAL PO --   " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf & _
                              "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo "

            ls_SQL = ls_SQL + "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 2 , " & vbCrLf & _
                              "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , "

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' "

            ls_SQL = ls_SQL + "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, "

            ls_SQL = ls_SQL + "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT "

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , "

            ls_SQL = ls_SQL + "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT "

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo "

            ls_SQL = ls_SQL + "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf & _
                              " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                       SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 2 , " & vbCrLf & _
                              "                                                               AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period , " & vbCrLf & _
                              "                                                               caption , " & vbCrLf & _
                              "                                                               F1 , " & vbCrLf & _
                              "                                                               F2 , " & vbCrLf & _
                              "                                                               F3 , " & vbCrLf & _
                              "                                                               F4 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = PO.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               no = '' , " & vbCrLf & _
                              "                                                               URUT = 2 , " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = caption , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = SUM(F3) , " & vbCrLf & _
                              "                                                               F4 = SUM(F4) , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              " 		  --F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , "

            ls_SQL = ls_SQL + "                                                               F1 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'    --F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , "

            ls_SQL = ls_SQL + "                                                               F2 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.POQty, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               --UNION ALL " & vbCrLf

            'ls_SQL = ls_SQL + "                                                               --SELECT " & vbCrLf & _
            '                  "                                                               --AffiliateID = POD.AffiliateID , " & vbCrLf & _
            '                  "                                                               --PartNo = POD.PartNo , " & vbCrLf & _
            '                  "                                                               --PartName = '' , " & vbCrLf & _
            '                  "                                                               --Period = POM.Period , " & vbCrLf & _
            '                  "                                                               --caption = 'Total PO' , " & vbCrLf & _
            '                  "                                                               --F1 = 0 , " & vbCrLf & _
            '                  "                                                               --F2 = 0 , " & vbCrLf & _
            '                  "                                                               --F3 = 0 , " & vbCrLf & _
            '                  "                                                               --F4 = ISNULL(POD.POQty, " & vbCrLf & _
            '                  "                                                               --0) "

            'ls_SQL = ls_SQL + "                                                               --FROM " & vbCrLf & _
            '                  "                                                               --po_detail POD " & vbCrLf & _
            '                  "                                                               --WITH ( NOLOCK ) " & vbCrLf & _
            '                  "                                                               --LEFT JOIN PO_Master POM " & vbCrLf & _
            '                  "                                                               --WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
            '                  "                                                               --AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
            '                  "                                                               --AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
            '                  "                                                               --WHERE " & vbCrLf & _
            '                  "                                                               --POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
            '                  "                                                               --UNION ALL " & vbCrLf & _
            '                  "                                                               --SELECT " & vbCrLf

            'ls_SQL = ls_SQL + "                                                               --AffiliateID = POD.AffiliateID , " & vbCrLf & _
            '                  "                                                               --PartNo = POD.PartNo , " & vbCrLf & _
            '                  "                                                               --PartName = '' , " & vbCrLf & _
            '                  "                                                              --Period = Period , " & vbCrLf & _
            '                  "                                                              --caption = 'Total PO' , " & vbCrLf & _
            '                  "                                                              --F1 = 0 , " & vbCrLf & _
            '                  "                                                               --F2 = 0 , " & vbCrLf & _
            '                  "                                                               --F3 = 0 , " & vbCrLf & _
            '                  "                                                               --F4 = ISNULL(POD.Week1, " & vbCrLf & _
            '                  "                                                               --0) " & vbCrLf & _
            '                  "                                                              -- FROM "

            'ls_SQL = ls_SQL + "                                                               --po_detail_Export POD " & vbCrLf & _
            '                  "                                                               --WITH ( NOLOCK ) " & vbCrLf & _
            '                  "                                                               --LEFT JOIN PO_Master_Export POM " & vbCrLf & _
            '                  "                                                               --WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
            '                  "                                                               --AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
            '                  "                                                               --AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
            '                  "                                                               --WHERE " & vbCrLf & _
            '                  "                                                               --POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
            ls_SQL = ls_SQL + "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) PO " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON PO.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "       -- 3XXX.TOTAL PO KANBAN --   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf & _
                              "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 3 , " & vbCrLf & _
                              "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , "

            ls_SQL = ls_SQL + "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono "

            ls_SQL = ls_SQL + "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , "

            ls_SQL = ls_SQL + "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf & _
                              " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 3 , " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = ISNULL(PO.PartNo, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               '') , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = ISNULL(Period, " & vbCrLf & _
                              "                                                               '" & ls_period & "') , " & vbCrLf & _
                              "                                                               caption = caption , " & vbCrLf & _
                              "                                                               F1 = ISNULL(SUM(F1), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(SUM(F2), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(SUM(F3), " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F4 = ISNULL(SUM(F4), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period2 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period3 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , "

            ls_SQL = ls_SQL + "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period4 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) PO " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON PO.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                       GROUP BY PO.AffiliateID , " & vbCrLf & _
                              "                                                               PO.PartNo , " & vbCrLf & _
                              "                                                               PO.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PO.caption " & vbCrLf & _
                              "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo     " & vbCrLf & _
                              "   --4XXX.Total PO Delivery --   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption = 'Total Delivery' , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , "

            ls_SQL = ls_SQL + "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 4 , " & vbCrLf & _
                              "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' "

            ls_SQL = ls_SQL + "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf

            ls_SQL = ls_SQL + " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 4 , " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = ISNULL(PO.PartNo, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = ISNULL(Period, " & vbCrLf & _
                              "                                                               '" & ls_period & "') , "

            ls_SQL = ls_SQL + "                                                               caption = caption , " & vbCrLf & _
                              "                                                               F1 = ISNULL(SUM(F1), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(SUM(F2), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(SUM(F3), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(SUM(F4), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    (   " & vbCrLf

            ls_SQL = ls_SQL + "   --F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf & _
                              "   --F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period2 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'   " & vbCrLf & _
                              "   --F3   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period3 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'    --F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period4 & "' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO Delivery' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) PO " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON PO.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                       GROUP BY PO.AffiliateID , " & vbCrLf & _
                              "                                                               PO.PartNo , " & vbCrLf & _
                              "                                                               PO.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf & _
                              "                                                               PO.caption " & vbCrLf & _
                              "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              " 	     " & vbCrLf & _
                              "   --5XXX.DIFF PO vs FORECAST--   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption = 'Diff PO vs Forecast' , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf & _
                              "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               URUT = 5 , " & vbCrLf & _
                              "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf

            ls_SQL = ls_SQL + "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf & _
                              " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT  no = '' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               URUT = 5 , " & vbCrLf & _
                              "                                                               AffiliateID = POx.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POx.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POx.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff PO vs Forecast' , " & vbCrLf & _
                              "                                                               F1 = CASE " & vbCrLf & _
                              "                                                               WHEN POx.F1 = 0 " & vbCrLf & _
                              "                                                               OR FR.F1 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POx.F1 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               / FR.F1 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F2 = CASE " & vbCrLf & _
                              "                                                               WHEN POx.F2 = 0 " & vbCrLf & _
                              "                                                               OR FR.F2 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POx.F2 " & vbCrLf & _
                              "                                                               / FR.F2 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F3 = CASE " & vbCrLf & _
                              "                                                               WHEN POx.F3 = 0 "

            ls_SQL = ls_SQL + "                                                               OR FR.F3 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POx.F3 " & vbCrLf & _
                              "                                                               / FR.F3 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F4 = CASE " & vbCrLf & _
                              "                                                               WHEN POx.F4 = 0 " & vbCrLf & _
                              "                                                               OR FR.F4 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POx.F4 " & vbCrLf & _
                              "                                                               / FR.F4 - 1 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               END , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , " & vbCrLf & _
                              "                                                               F3 = SUM(F3) , " & vbCrLf & _
                              "                                                               F4 = SUM(F4) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD2.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(ForecastN2, " & vbCrLf

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(ForecastN3, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F1 = ISNULL(POD2.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(Forecast2, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(Forecast3, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName " & vbCrLf & _
                              "                                                               ) FR " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , " & vbCrLf & _
                              "                                                               F3 = SUM(F3) , " & vbCrLf & _
                              "                                                               F4 = SUM(F4) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F2 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName " & vbCrLf & _
                              "                                                               ) POx ON FR.AffiliateID = POx.AffiliateID " & vbCrLf & _
                              "                                                               AND FR.PartNo = POx.PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "      " & vbCrLf & _
                              "   --6XXX.Diff Delivery vs PO   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 6 , " & vbCrLf & _
                              "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf

            ls_SQL = ls_SQL + "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf

            ls_SQL = ls_SQL + " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT DISTINCT " & vbCrLf & _
                              "                                                               no = '' , " & vbCrLf & _
                              "                                                               URUT = 6 , " & vbCrLf & _
                              "                                                               AffiliateID = PO.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = PO.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = PO.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = CASE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               WHEN PO.F1 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POK.F1 " & vbCrLf & _
                              "                                                               / PO.F1 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F2 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F2 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POK.F2 " & vbCrLf & _
                              "                                                               / PO.F2 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F3 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POK.F3 " & vbCrLf & _
                              "                                                               / PO.F3 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F4 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F4 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POK.F4 " & vbCrLf & _
                              "                                                               / PO.F4 - 1 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               END , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , " & vbCrLf & _
                              "                                                               F3 = SUM(F3) , " & vbCrLf & _
                              "                                                               F4 = SUM(F4) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   --F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf

            ls_SQL = ls_SQL + "   --F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'   " & vbCrLf & _
                              "   --F3   " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   " & vbCrLf & _
                              "   --F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf

            ls_SQL = ls_SQL + "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) PO " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf & _
                              "                                                               URUT = 3 , " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = ISNULL(x.PartNo, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               PartName = ISNULL(MP.PartName, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               Period = ISNULL(Period, " & vbCrLf & _
                              "                                                               '" & ls_period & "') , " & vbCrLf & _
                              "                                                               caption = caption , " & vbCrLf & _
                              "                                                               F1 = ISNULL(SUM(F1), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(SUM(F2), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(SUM(F3), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               F4 = ISNULL(SUM(F4), " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   --F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period & "' , " & vbCrLf

            ls_SQL = ls_SQL + "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.POQty, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.Week1, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf & _
                              "   --F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period2 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , "

            ls_SQL = ls_SQL + "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'   " & vbCrLf & _
                              "   --F3   " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period3 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   " & vbCrLf & _
                              "   --F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, "

            ls_SQL = ls_SQL + "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period4 & "' , " & vbCrLf & _
                              "                                                               caption = 'Total PO Kanban' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , "

            ls_SQL = ls_SQL + "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               AND KanbanCls = '1' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , "

            ls_SQL = ls_SQL + "                                                               caption = 'Total PO KANBAN' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf & _
                              "                                                               x.caption "

            ls_SQL = ls_SQL + "                                                               ) POK ON PO.AffiliateID = POK.AffiliateID " & vbCrLf & _
                              "                                                               AND PO.PartNo = POK.PartNo " & vbCrLf & _
                              "                                                               AND PO.Period = POK.Period " & vbCrLf & _
                              "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo "

            ls_SQL = ls_SQL + "      " & vbCrLf & _
                              "   --7XXX.DIFF DELIVERY vs PO   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption = 'Diff Delivery vs PO' , "

            ls_SQL = ls_SQL + "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf & _
                              "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 7 , " & vbCrLf & _
                              "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , "

            ls_SQL = ls_SQL + "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , "

            ls_SQL = ls_SQL + "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , "

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , "

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' "

            ls_SQL = ls_SQL + "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, "

            ls_SQL = ls_SQL + "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT "

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono "

            ls_SQL = ls_SQL + "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , "

            ls_SQL = ls_SQL + "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf & _
                              " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT DISTINCT " & vbCrLf & _
                              "                                                               no = '' , " & vbCrLf & _
                              "                                                               URUT = 7 , " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(PO.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PO.PartNo , "

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = PO.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F1 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F1 " & vbCrLf & _
                              "                                                               / PO.F1 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F2 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F2 = 0 "

            ls_SQL = ls_SQL + "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F2 " & vbCrLf & _
                              "                                                               / PO.F2 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F3 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F3 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F3 " & vbCrLf & _
                              "                                                               / PO.F3 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F4 = CASE "

            ls_SQL = ls_SQL + "                                                               WHEN PO.F4 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F4 " & vbCrLf & _
                              "                                                               / PO.F4 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    ( ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = ISNULL(x.PartNo, " & vbCrLf & _
                              "                                                               '') , "

            ls_SQL = ls_SQL + "                                                               PartName = ISNULL(MP.PartName, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               Period = ISNULL(Period, " & vbCrLf & _
                              "                                                               '" & ls_period & "') , " & vbCrLf & _
                              "                                                               F1 = ISNULL(SUM(F1), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(SUM(F2), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(SUM(F3), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(SUM(F4), "

            ls_SQL = ls_SQL + "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   	--F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POM.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period & "' , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , "

            ls_SQL = ls_SQL + "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   "

            ls_SQL = ls_SQL + "   	--F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POM.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period2 & "' , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , "

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , "

            ls_SQL = ls_SQL + "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'    	--F3   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = ISNULL(POM.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period3 & "' , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo "

            ls_SQL = ls_SQL + "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   " & vbCrLf & _
                              "   	--F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POM.AffiliateID, " & vbCrLf & _
                              "                                                               '') , "

            ls_SQL = ls_SQL + "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period4 & "' , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf & _
                              "                                                               x.caption "

            ls_SQL = ls_SQL + "                                                               ) POD " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf & _
                              "                                                               URUT = 2 , " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = caption , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , " & vbCrLf & _
                              "                                                               F3 = SUM(F3) , "

            ls_SQL = ls_SQL + "                                                               F4 = SUM(F4) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   	--F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.POQty, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf & _
                              "   	--F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.POQty, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'    	--F3   " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   " & vbCrLf & _
                              "   	--F4   " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.Week1, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) PO ON PO.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND PO.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               AND PO.Period = POD.Period " & vbCrLf & _
                              "                                                               ) " & vbCrLf & _
                              "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , "

            ls_SQL = ls_SQL + "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo   " & vbCrLf & _
                              "       --8XXX.DIFF PO DELIVERY vs FORECAST   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , "

            ls_SQL = ls_SQL + "                                                     HPartNo = '' , " & vbCrLf & _
                              "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption = 'Diff Delivery vs Forecast' , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf & _
                              "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 8 , "

            ls_SQL = ls_SQL + "                                                               z.AffiliateID , " & vbCrLf & _
                              "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs Forecast' "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , "

            ls_SQL = ls_SQL + "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , "

            ls_SQL = ls_SQL + "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf & _
                              " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT DISTINCT " & vbCrLf & _
                              "                                                               no = '' , "

            ls_SQL = ls_SQL + "                                                               URUT = 8 , " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POD.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs Forecast' , " & vbCrLf & _
                              "                                                               F1 = CASE " & vbCrLf & _
                              "                                                               WHEN POD.F1 = 0 " & vbCrLf & _
                              "                                                               OR FR.F1 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F1 "

            ls_SQL = ls_SQL + "                                                               / FR.F1 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F2 = CASE " & vbCrLf & _
                              "                                                               WHEN POD.F2 = 0 " & vbCrLf & _
                              "                                                               OR FR.F2 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F2 " & vbCrLf & _
                              "                                                               / FR.F2 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F3 = CASE " & vbCrLf & _
                              "                                                               WHEN POD.F3 = 0 "

            ls_SQL = ls_SQL + "                                                               OR FR.F3 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F3 " & vbCrLf & _
                              "                                                               / FR.F3 - 1 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F4 = CASE " & vbCrLf & _
                              "                                                               WHEN POD.F4 = 0 " & vbCrLf & _
                              "                                                               OR FR.F4 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE POD.F4 " & vbCrLf & _
                              "                                                               / FR.F4 - 1 "

            ls_SQL = ls_SQL + "                                                               END , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    ( ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = ISNULL(x.PartNo, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartName = ISNULL(MP.PartName, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               Period = ISNULL(Period, " & vbCrLf & _
                              "                                                               '" & ls_period & "') , "

            ls_SQL = ls_SQL + "                                                               F1 = ISNULL(SUM(F1), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(SUM(F2), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(SUM(F3), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(SUM(F4), " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   	--F1   " & vbCrLf

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf & _
                              "   	--F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               period = '" & ls_period2 & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'   " & vbCrLf & _
                              "   	--F3   " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period3 & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 "

            ls_SQL = ls_SQL + "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   " & vbCrLf & _
                              "   	--F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , "

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period4 & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , "

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , "

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' "

            ls_SQL = ls_SQL + "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName " & vbCrLf & _
                              "                                                               ) POD " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf & _
                              "                                                               URUT = 1 , " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , "

            ls_SQL = ls_SQL + "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , " & vbCrLf & _
                              "                                                               F3 = SUM(F3) , " & vbCrLf & _
                              "                                                               F4 = SUM(F4) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD2.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(ForecastN2, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(ForecastN3, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD2.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(Forecast2, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(Forecast3, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT "

            ls_SQL = ls_SQL + "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono "

            ls_SQL = ls_SQL + "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName " & vbCrLf & _
                              "                                                               ) FR ON FR.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND FR.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) " & vbCrLf & _
                              "                                                     ) gab " & vbCrLf & _
                              "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo , "

            ls_SQL = ls_SQL + "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo     " & vbCrLf & _
                              "   --9XXX.BALANCE PO   " & vbCrLf & _
                              "                                           UNION ALL " & vbCrLf & _
                              "                                           SELECT DISTINCT " & vbCrLf & _
                              "                                                     no = '' , " & vbCrLf & _
                              "                                                     URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo = '' , "

            ls_SQL = ls_SQL + "                                                     PartName='' , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption = 'Balance PO' , " & vbCrLf & _
                              "                                                     F1 = SUM(F1) , " & vbCrLf & _
                              "                                                     F2 = SUM(F2) , " & vbCrLf & _
                              "                                                     F3 = SUM(F3) , " & vbCrLf & _
                              "                                                     F4 = SUM(F4) , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                           FROM      ( SELECT  no = '' , " & vbCrLf & _
                              "                                                               URUT = 9 , " & vbCrLf & _
                              "                                                               z.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               HPartNo = '' , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               z.Period , " & vbCrLf & _
                              "                                                               z.caption , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 , " & vbCrLf & _
                              "                                                               PartNo = z.PartNo " & vbCrLf & _
                              "                                                       FROM    ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               0, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs Forecast' " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               1, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID "

            ls_SQL = ls_SQL + "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' "

            ls_SQL = ls_SQL + "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               Period = DATEADD(month, " & vbCrLf & _
                              "                                                               2, Period) , " & vbCrLf & _
                              "                                                               caption = Caption " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(ForecastN1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , "

            ls_SQL = ls_SQL + "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               PO_Detail b WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_master a " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT DISTINCT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT DISTINCT " & vbCrLf & _
                              "                                                               Forecast1 = ISNULL(Forecast1, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               a.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export b " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export a "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              "                                                               AND a.SupplierID = b.SupplierID " & vbCrLf & _
                              "                                                               AND a.POno = b.Pono " & vbCrLf & _
                              "                                                               AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_Period1 & "' " & vbCrLf & _
                              "                                                               ) POD2 ON POD2.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND POD2.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) z " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON z.PartNo = MP.PartNo  " & vbCrLf & _
                              " 	 --DATA PO " & vbCrLf & _
                              "                                                       UNION ALL " & vbCrLf & _
                              "                                                       SELECT DISTINCT " & vbCrLf & _
                              "                                                               no = '' , " & vbCrLf & _
                              "                                                               URUT = 9 , "

            ls_SQL = ls_SQL + "                                                               AffiliateID = PO.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = PO.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = PO.Period , " & vbCrLf & _
                              "                                                               caption = 'Diff Delivery vs PO' , " & vbCrLf & _
                              "                                                               F1 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F1 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE PO.F1 " & vbCrLf & _
                              "                                                               - POD.F1 " & vbCrLf & _
                              "                                                               END , "

            ls_SQL = ls_SQL + "                                                               F2 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F2 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE PO.F2 " & vbCrLf & _
                              "                                                               - POD.F2 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               F3 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F3 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE PO.F3 " & vbCrLf & _
                              "                                                               - POD.F3 "

            ls_SQL = ls_SQL + "                                                               END , " & vbCrLf & _
                              "                                                               F4 = CASE " & vbCrLf & _
                              "                                                               WHEN PO.F4 = 0 " & vbCrLf & _
                              "                                                               THEN 0 " & vbCrLf & _
                              "                                                               ELSE PO.F4 " & vbCrLf & _
                              "                                                               - POD.F4 " & vbCrLf & _
                              "                                                               END , " & vbCrLf & _
                              "                                                               HPartNo = '' " & vbCrLf & _
                              "                                                       FROM    ( ( SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(AffiliateID, " & vbCrLf & _
                              "                                                               '') , "

            ls_SQL = ls_SQL + "                                                               PartNo = ISNULL(x.PartNo, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartName = ISNULL(MP.PartName, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               Period = ISNULL(Period, " & vbCrLf & _
                              "                                                               '" & ls_period & "') , " & vbCrLf & _
                              "                                                               F1 = ISNULL(SUM(F1), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = ISNULL(SUM(F2), " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = ISNULL(SUM(F3), "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = ISNULL(SUM(F4), " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   	--F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               period = '" & ls_period & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , "

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , "

            ls_SQL = ls_SQL + "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf & _
                              "   	--F2   " & vbCrLf

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_period2 & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 "

            ls_SQL = ls_SQL + "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT "

            ls_SQL = ls_SQL + "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 "

            ls_SQL = ls_SQL + "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "'   " & vbCrLf & _
                              "   	--F3   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , "

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period3 & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO "

            ls_SQL = ls_SQL + "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , "

            ls_SQL = ls_SQL + "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , "

            ls_SQL = ls_SQL + "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM "

            ls_SQL = ls_SQL + "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   "

            ls_SQL = ls_SQL + "   	--F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = ISNULL(POD.AffiliateID, " & vbCrLf & _
                              "                                                               '') , " & vbCrLf & _
                              "                                                               PartNo = partNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               period = '" & ls_Period4 & "' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL "

            ls_SQL = ls_SQL + "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM "

            ls_SQL = ls_SQL + "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = ISNULL(DSD.DOQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               LEFT JOIN DOSUPPLIER_Detail_Export DSD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON DSD.PONo = POM.PONo "

            ls_SQL = ls_SQL + "                                                               AND DSD.OrderNo = POM.OrderNo1 " & vbCrLf & _
                              "                                                               AND DSD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND DSD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                               AND DSD.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x " & vbCrLf & _
                              "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , "

            ls_SQL = ls_SQL + "                                                               MP.PartName " & vbCrLf & _
                              "                                                               ) POD " & vbCrLf & _
                              "                                                               LEFT JOIN ( SELECT " & vbCrLf & _
                              "                                                               URUT = 2 , " & vbCrLf & _
                              "                                                               AffiliateID = AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = x.PartNo , " & vbCrLf & _
                              "                                                               PartName = MP.PartName , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = caption , " & vbCrLf & _
                              "                                                               F1 = SUM(F1) , " & vbCrLf & _
                              "                                                               F2 = SUM(F2) , "

            ls_SQL = ls_SQL + "                                                               F3 = SUM(F3) , " & vbCrLf & _
                              "                                                               F4 = SUM(F4) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               (   " & vbCrLf & _
                              "   	--F1   " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , "

            ls_SQL = ls_SQL + "                                                               F1 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID "

            ls_SQL = ls_SQL + "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = ISNULL(POD.Week1, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period & "'   " & vbCrLf & _
                              "   	--F2   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , "

            ls_SQL = ls_SQL + "                                                               F2 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID "

            ls_SQL = ls_SQL + "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_period2 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = ISNULL(POD.Week1, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_period2 & "'   " & vbCrLf & _
                              "   	--F3   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , "

            ls_SQL = ls_SQL + "                                                               F3 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE "

            ls_SQL = ls_SQL + "                                                               POM.Period = '" & ls_Period3 & "' " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = ISNULL(POD.Week1, "

            ls_SQL = ls_SQL + "                                                               0) , " & vbCrLf & _
                              "                                                               F4 = 0 " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period3 & "'   "

            ls_SQL = ls_SQL + "   	--F4   " & vbCrLf & _
                              "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = POM.Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , "

            ls_SQL = ls_SQL + "                                                               F4 = ISNULL(POD.POQty, " & vbCrLf & _
                              "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' "

            ls_SQL = ls_SQL + "                                                               UNION ALL " & vbCrLf & _
                              "                                                               SELECT " & vbCrLf & _
                              "                                                               AffiliateID = POD.AffiliateID , " & vbCrLf & _
                              "                                                               PartNo = POD.PartNo , " & vbCrLf & _
                              "                                                               PartName = '' , " & vbCrLf & _
                              "                                                               Period = Period , " & vbCrLf & _
                              "                                                               caption = 'Total PO' , " & vbCrLf & _
                              "                                                               F1 = 0 , " & vbCrLf & _
                              "                                                               F2 = 0 , " & vbCrLf & _
                              "                                                               F3 = 0 , " & vbCrLf & _
                              "                                                               F4 = ISNULL(POD.Week1, "

            ls_SQL = ls_SQL + "                                                               0) " & vbCrLf & _
                              "                                                               FROM " & vbCrLf & _
                              "                                                               po_detail_Export POD " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) " & vbCrLf & _
                              "                                                               LEFT JOIN PO_Master_Export POM " & vbCrLf & _
                              "                                                               WITH ( NOLOCK ) ON POM.Pono = POD.PONO " & vbCrLf & _
                              "                                                               AND POM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                               AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              "                                                               WHERE " & vbCrLf & _
                              "                                                               POM.Period = '" & ls_Period4 & "' " & vbCrLf & _
                              "                                                               ) x "

            ls_SQL = ls_SQL + "                                                               LEFT JOIN MS_Parts MP ON x.PartNo = MP.PartNo " & vbCrLf & _
                              "                                                               GROUP BY x.AffiliateID , " & vbCrLf & _
                              "                                                               x.PartNo , " & vbCrLf & _
                              "                                                               x.Period , " & vbCrLf & _
                              "                                                               MP.PartName , " & vbCrLf & _
                              "                                                               x.caption " & vbCrLf & _
                              "                                                               ) PO ON PO.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                                               AND PO.PartNo = POD.PartNo " & vbCrLf & _
                              "                                                               AND PO.Period = POD.Period " & vbCrLf & _
                              "                                                               ) " & vbCrLf & _
                              "                                                     ) gab "

            ls_SQL = ls_SQL + "                                           GROUP BY  URUT , " & vbCrLf & _
                              "                                                     AffiliateID , " & vbCrLf & _
                              "                                                     HPartNo , " & vbCrLf & _
                              "                                                     PartName , " & vbCrLf & _
                              "                                                     Period , " & vbCrLf & _
                              "                                                     caption , " & vbCrLf & _
                              "                                                     PartNo " & vbCrLf & _
                              "                                         ) Detail " & vbCrLf & _
                              "                               --WHERE     AffiliateID = 'HESTO' " & vbCrLf
            If cboaffiliate.Text <> clsGlobal.gs_All Then
                ls_SQL = ls_SQL + "where AffiliateID = '" & ls_Affiliate & "' " & vbCrLf
            End If
            ls_SQL = ls_SQL + "                                         --AND partNo = '7154-0708-30' " & vbCrLf & _
                              "                             ) gabungan " & vbCrLf

            ls_SQL = ls_SQL + "                 ) x WHERE x.Partno <> ''" & vbCrLf & _
                              "         ORDER BY PartNo , " & vbCrLf & _
                              "                 CONVERT(CHAR(10), CONVERT(DATETIME, Period), 112) , " & vbCrLf & _
                              "                 AffiliateID , " & vbCrLf & _
                              "                 Urut " & vbCrLf & _
                              "     END " & vbCrLf & _
                              "   ELSE  " & vbCrLf & _
                              "     BEGIN   " & vbCrLf & _
                              "         SELECT  no = '' , " & vbCrLf & _
                              "                 URUT = '' , " & vbCrLf & _
                              "                 AffiliateID = '' , "

            ls_SQL = ls_SQL + "                 HPartNo = '' , " & vbCrLf & _
                              "                 PartName = '' , " & vbCrLf & _
                              "                 Period = '' , " & vbCrLf & _
                              "                 caption = '' , " & vbCrLf & _
                              "                 F1 = '' , " & vbCrLf & _
                              "                 F2 = '' , " & vbCrLf & _
                              "                 F3 = '' , " & vbCrLf & _
                              "                 F4 = '' , " & vbCrLf & _
                              "                 PartNo = '' " & vbCrLf & _
                              "         FROM    po_master_export WITH(NOLOCK) " & vbCrLf & _
                              "         WHERE   Period = '" & ls_period & "' "

            ls_SQL = ls_SQL + "         UNION ALL " & vbCrLf & _
                              "         SELECT  no = '' , " & vbCrLf & _
                              "                 URUT = '' , " & vbCrLf & _
                              "                 AffiliateID = '' , " & vbCrLf & _
                              "                 HPartNo = '' , " & vbCrLf & _
                              "                 PartName = '' , " & vbCrLf & _
                              "                 Period = '' , " & vbCrLf & _
                              "                 caption = '' , " & vbCrLf & _
                              "                 F1 = '' , " & vbCrLf & _
                              "                 F2 = '' , " & vbCrLf & _
                              "                 F3 = '' , "

            ls_SQL = ls_SQL + "                 F4 = '' , " & vbCrLf & _
                              "                 PartNo = '' " & vbCrLf & _
                              "         FROM    po_master WITH(NOLOCK) " & vbCrLf & _
                              "         WHERE   Period = '" & ls_period & "'   " & vbCrLf & _
                              "     END "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
            End With
            dtHeader = ds.Tables(0)
            sqlConn.Close()
        End Using
    End Sub

#End Region

#Region "EXCEL"
    Private Sub GetExcel()
        Call up_GridLoad()
        FileName = "PRE-BOOKING.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        If dtHeader.Rows.Count - 1 > 0 Then
            Call epplusExportHeaderExcel(FilePath, "", dtHeader, dtDetail, "C:7", "")
        Else
            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.ErrorMessage)
            grid.JSProperties("cpMessage") = lblerrmessage.Text
        End If
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

    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData1 As DataTable, ByVal pData2 As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try

            Dim NewFileName As String = Server.MapPath("~\PurchaseOrderExport\PreBooking.xlsx")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet
            Dim sumQtyBox1 As Long = 0
            Dim sumPalletBox1 As Long = 0
            Dim sumQtyBox2 As Long = 0
            Dim sumPalletBox2 As Long = 0
            Dim sumQtyBox3 As Long = 0
            Dim sumPalletBox3 As Long = 0
            Dim sumQtyBox4 As Long = 0
            Dim sumPalletBox4 As Long = 0
            Dim sumCbm1 As Long = 0
            Dim sumCbm2 As Long = 0
            Dim sumCbm3 As Long = 0
            Dim sumCbm4 As Long = 0

            ws = exl.Workbook.Worksheets("PRE-BOOKING")
            Dim irow As Long = 0
            Dim iRowTmp As Long = 0
            Dim icol As Long = 0

            With ws
                For irow = 0 To pData1.Rows.Count - 1
                    If pData1.Rows.Count > 0 Then
                        If pData1.Rows(irow)("Week") = "1" Then
                            ws.Cells("C7").Value = Trim(pData1.Rows(irow)("OrderNo1"))
                            ws.Cells("C8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("C9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("C10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("C11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("C12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("C13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 3).Style.Font.Size = 11
                            .Cells(7, 3).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "2" Then
                            ws.Cells("D7").Value = Trim(pData1.Rows(irow)("OrderNo2"))
                            ws.Cells("D8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("D9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("D10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("D11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("D12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("D13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 4).Style.Font.Size = 11
                            .Cells(7, 4).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "3" Then
                            ws.Cells("E7").Value = Trim(pData1.Rows(irow)("OrderNo3"))
                            ws.Cells("E8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("E9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("E10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("E11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("E12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("E13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 5).Style.Font.Size = 11
                            .Cells(7, 5).Style.Font.Name = "Calibri"
                        ElseIf pData1.Rows(irow)("Week") = "4" Then
                            ws.Cells("F7").Value = Trim(pData1.Rows(irow)("OrderNo4"))
                            ws.Cells("F8").Value = pData1.Rows(irow)("ETAForwarder")
                            ws.Cells("F9").Value = pData1.Rows(irow)("ETDPort")
                            ws.Cells("F10").Value = pData1.Rows(irow)("ETAPort")
                            ws.Cells("F11").Value = pData1.Rows(irow)("ETAFactory")
                            ws.Cells("F12").Value = pData1.Rows(irow)("Vessel1")
                            ws.Cells("F13").Value = pData1.Rows(irow)("Vessel2")
                            .Cells(7, 6).Style.Font.Size = 11
                            .Cells(7, 6).Style.Font.Name = "Calibri"
                        End If
                    End If
                Next
            End With

            iRowTmp = 18
            For irow = 0 To pData2.Rows.Count - 1
                If pData2.Rows.Count > 0 Then
                    ws.Cells("B" & iRowTmp).Value = Trim(pData2.Rows(irow)("SupplierID"))
                    ws.Cells("C" & iRowTmp).Value = pData2.Rows(irow)("QtyBox1")
                    ws.Cells("D" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet1")
                    ws.Cells("E" & iRowTmp).Value = pData2.Rows(irow)("QtyBox2")
                    ws.Cells("F" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet2")
                    ws.Cells("G" & iRowTmp).Value = pData2.Rows(irow)("QtyBox3")
                    ws.Cells("H" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet3")
                    ws.Cells("I" & iRowTmp).Value = pData2.Rows(irow)("QtyBox4")
                    ws.Cells("J" & iRowTmp).Value = pData2.Rows(irow)("QtyPallet4")
                    'ws.Cells("C18" & ":J" & iRowTmp).Style.Numberformat.Format = "#,###"

                    sumQtyBox1 = sumQtyBox1 + pData2.Rows(irow)("QtyBox1")
                    sumPalletBox1 = sumPalletBox1 + pData2.Rows(irow)("QtyPallet1")
                    sumQtyBox2 = sumQtyBox2 + pData2.Rows(irow)("QtyBox2")
                    sumPalletBox2 = sumPalletBox2 + pData2.Rows(irow)("QtyPallet2")
                    sumQtyBox3 = sumQtyBox3 + pData2.Rows(irow)("QtyBox3")
                    sumPalletBox3 = sumPalletBox3 + pData2.Rows(irow)("QtyPallet3")
                    sumQtyBox4 = sumQtyBox4 + pData2.Rows(irow)("QtyBox4")
                    sumPalletBox4 = sumPalletBox4 + pData2.Rows(irow)("QtyPallet4")

                    sumCbm1 = sumCbm1 + pData2.Rows(irow)("CBM1")
                    sumCbm2 = sumCbm2 + pData2.Rows(irow)("CBM2")
                    sumCbm3 = sumCbm3 + pData2.Rows(irow)("CBM3")
                    sumCbm4 = sumCbm4 + pData2.Rows(irow)("CBM4")

                    'ALIGNMENT
                    ws.Cells("B" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                    ws.Cells("C" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("D" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("E" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("F" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("G" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("H" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("I" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    ws.Cells("J" & iRowTmp).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'FORMAT
                    ws.Cells("C" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("D" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("E" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("F" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("G" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("H" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("I" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                    ws.Cells("J" & iRowTmp).Style.Numberformat.Format = "###,##0.00"
                End If
                iRowTmp = iRowTmp + 1
            Next

            iRowTmp = iRowTmp
            ws.Cells("B" & iRowTmp).Value = "TOTAL CBM"
            ws.Cells("D" & iRowTmp).Value = sumCbm1
            ws.Cells("F" & iRowTmp).Value = sumCbm2
            ws.Cells("H" & iRowTmp).Value = sumCbm3
            ws.Cells("J" & iRowTmp).Value = sumCbm4

            ws.Cells("B" & iRowTmp + 1).Value = "TOTAL PALLET"
            ws.Cells("D" & iRowTmp + 1).Value = sumPalletBox1
            ws.Cells("F" & iRowTmp + 1).Value = sumPalletBox2
            ws.Cells("H" & iRowTmp + 1).Value = sumPalletBox3
            ws.Cells("J" & iRowTmp + 1).Value = sumPalletBox4

            ws.Cells("A" & iRowTmp + 2).Value = "CONTAINER"

            'rumus 40FT
            Dim ls_40FT As Single
            ls_40FT = 15

            'rumus 20FT
            Dim ls_20FT As Single
            ls_20FT = 7.5

            Dim JmlPallet As Long = 0
            Dim JmlContainer40 As Long = 0
            Dim JmlContainer20 As Long = 0

            '=============WEEK 1=================='
            If CDbl(sumPalletBox1) >= ls_40FT Then
                'menggunakan 40FT
                JmlContainer40 = Int(sumPalletBox1 / ls_40FT)
                ws.Cells("D" & iRowTmp + 2).Value = JmlContainer40
                JmlPallet = sumPalletBox1 Mod ls_40FT
                If CDbl(JmlPallet) > 0 Then JmlContainer20 = JmlPallet / ls_20FT : ws.Cells("D" & iRowTmp + 3).Value = JmlContainer20

            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox1 / ls_20FT)
                ws.Cells("D" & iRowTmp + 3).Value = JmlContainer20
            End If

            '=============WEEK 2=================='
            If CDbl(sumPalletBox2) >= ls_40FT Then
                'menggunakan 40FT
                JmlContainer40 = Int(sumPalletBox2 / ls_40FT)
                ws.Cells("F" & iRowTmp + 2).Value = JmlContainer40
                JmlPallet = sumPalletBox2 Mod ls_40FT
                If CDbl(JmlPallet) > 0 Then JmlContainer20 = JmlPallet / ls_20FT : ws.Cells("F" & iRowTmp + 3).Value = JmlContainer20

            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox1 / ls_20FT)
                ws.Cells("F" & iRowTmp + 3).Value = JmlContainer20
            End If

            '=============WEEK 3=================='
            If CDbl(sumPalletBox3) >= ls_40FT Then
                'menggunakan 40FT
                JmlContainer40 = Int(sumPalletBox3 / ls_40FT)
                ws.Cells("H" & iRowTmp + 2).Value = JmlContainer40
                JmlPallet = sumPalletBox3 Mod ls_40FT
                If CDbl(JmlPallet) > 0 Then JmlContainer20 = JmlPallet / ls_20FT : ws.Cells("H" & iRowTmp + 3).Value = JmlContainer20

            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox1 / ls_20FT)
                ws.Cells("H" & iRowTmp + 3).Value = JmlContainer20
                ws.Cells("H" & iRowTmp + 3).Style.Numberformat.Format = "#,##0"
            End If

            '=============WEEK 4=================='
            If CDbl(sumPalletBox4) >= ls_40FT Then
                'menggunakan 40FT
                JmlContainer40 = Int(sumPalletBox4 / ls_40FT)
                ws.Cells("J" & iRowTmp + 2).Value = JmlContainer40
                ws.Cells("J" & iRowTmp + 2).Style.Numberformat.Format = "#,##0"
                JmlPallet = sumPalletBox4 Mod ls_40FT
                If CDbl(JmlPallet) > 0 Then JmlContainer20 = JmlPallet / ls_20FT : ws.Cells("J" & iRowTmp + 3).Value = JmlContainer20 : ws.Cells("J" & iRowTmp + 3).Style.Numberformat.Format = "#,##0"

            Else
                JmlContainer20 = Math.Ceiling(sumPalletBox1 / ls_20FT)
                ws.Cells("J" & iRowTmp + 3).Value = JmlContainer20
                ws.Cells("J" & iRowTmp + 3).Style.Numberformat.Format = "#,##0"
                'If CDbl(JmlPallet) > 0 Then JmlContainer20 = JmlPallet / ls_40FT
            End If

            ws.Cells("B" & iRowTmp + 2).Value = "20FT"
            ws.Cells("B" & iRowTmp + 3).Value = "40FT"

            Dim rgAll As ExcelRange = ws.Cells(18, 2, iRowTmp + 2, 10)
            EpPlusDrawAllBorders(rgAll)

            Dim rgAll2 As ExcelRange = ws.Cells(iRowTmp + 2, 1, iRowTmp + 3, 10)
            EpPlusDrawAllBorders(rgAll2)

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

#End Region

    Protected Sub btnprintcard_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnprintcard.Click

        up_GridLoad()
        'Dim f As GridViewBandColumn = TryCast(grid.Columns("TahunAwal"), GridViewBandColumn)
        'f.Caption = Year(dtFrom.Value).ToString
        'Dim t As GridViewBandColumn = TryCast(grid.Columns("TahunAkhir"), GridViewBandColumn)
        't.Caption = Year(dtTo.Value).ToString


        Dim ps As New PrintingSystem()

        Dim link1 As New PrintableComponentLink(ps)
        link1.Component = GridExporter

        Dim compositeLink As New CompositeLink(ps)
        compositeLink.Links.AddRange(New Object() {link1})

        compositeLink.CreateDocument()
        Using stream As New MemoryStream()
            compositeLink.PrintingSystem.ExportToXlsx(stream)
            Response.Clear()
            Response.Buffer = False
            Response.AppendHeader("Content-Type", "application/xlsx")
            Response.AppendHeader("Content-Transfer-Encoding", "binary")
            Response.AppendHeader("Content-Disposition", "attachment; filename=Report Forecast vs PO.xlsx")
            Response.BinaryWrite(stream.ToArray())
            Response.End()
        End Using
        ps.Dispose()

    End Sub

    Protected Sub btnsearch_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsearch.Click
        up_GridHeader()
        up_GridLoad()
        If grid.VisibleRowCount = 0 Then
            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblerrmessage.Text
        End If
    End Sub
End Class