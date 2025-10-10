Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class SummaryForecast
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
            up_IsiHeader()
        End If

        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager, False, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()
                    Call up_IsiHeader()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        Session("AA220Msg") = lblInfo.Text
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                    Call up_IsiHeader()
                Case "loadSave"
                    grid.PageIndex = 0
                    bindData()
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = clsMaster.GetTableSummaryForecast(dtPeriodFrom.Value, cboAffiliate.Text, cboSupplier.Text, cboPartNo.Text)
                    FileName = "TemplateSummaryForecast.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:10", psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Session("AA220Msg") = ""
        End Try
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData()
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboPartNo.Text <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.PartNo = '" & cboPartNo.Text & "'"
        End If

        If cboAffiliate.Text <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.AffiliateID = '" & cboAffiliate.Text & "'"
        End If

        If cboSupplier.Text <> clsGlobal.gs_All Then
            pWhere = pWhere & " and b.SupplierID = '" & cboSupplier.Text & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " SELECT  " & vbCrLf & _
                              "     row_number() over (order by b.PartNo) as NoUrut, " & vbCrLf & _
                              " 	b.PartNo,  " & vbCrLf & _
                              " 	c.PartName,  " & vbCrLf & _
                              " 	b.AffiliateID, " & vbCrLf & _
                              " 	b.SupplierID, " & vbCrLf & _
                              " 	b.POMOQ, " & vbCrLf & _
                              " 	c.Project, " & vbCrLf & _
                              " 	b.PONo, " & vbCrLf & _
                              " 	ISNULL(b.POQty,0) Bln1, " & vbCrLf & _
                              " 	ISNULL(b.ForecastN1,0) Bln2, " & vbCrLf & _
                              " 	ISNULL(b.ForecastN2,0) Bln3, " & vbCrLf & _
                              " 	ISNULL(b.ForecastN3,0) Bln4 " & vbCrLf

            ls_SQL = ls_SQL + " FROM PO_Master a with (nolock)" & vbCrLf & _
                              " INNER JOIN PO_Detail b with (nolock) on a.PONO = b.PONo and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID " & vbCrLf & _
                              " LEFT JOIN MS_Parts c on b.PartNo = c.PartNo " & vbCrLf & _
                              " LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = b.PartNo AND MPM.AffiliateID = b.AffiliateID AND MPM.SupplierID = b.SupplierID " & vbCrLf & _
                              " WHERE YEAR(Period) = " & Year(dtPeriodFrom.Value) & " and MONTH(Period) = " & Month(dtPeriodFrom.Value) & " and FinalApproveDate IS NOT NULL " & pWhere & "" & vbCrLf & _
                              " ORDER BY b.PartNo "

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

            ls_SQL = " select top 0  '' NoUrut, '' PartNo, '' PartName, ''AffiliateID, ''SupplierID, ''MOQ, ''Project, ''PONo, ''Bln1, ''Bln2, ''Bln3, ''Bln4"

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

    Private Sub up_IsiHeader()
        Dim iWeek As Integer

        For iWeek = 1 To 4
            grid.Columns("Bln" & iWeek).Caption = Format(DateAdd(DateInterval.Month, iWeek - 1, dtPeriodFrom.Value), "MMM-yy")
        Next
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all " & vbCrLf & _
                 "select distinct RTRIM(a.PartNo) PartCode, PartName from MS_Parts a" & vbCrLf & _
                 "where FinishGoodCls = '2' order by PartCode "

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartCode")
                .Columns(0).Width = 85
                .Columns.Add("PartName")
                .Columns(1).Width = 180

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = 0
                txtPartNo.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using



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

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Summary Forecast " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
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
                .Cells(5, 1).Value = "(" & Format(dtPeriodFrom.Value, "MMM yyyy") & " Production)"
                .Cells(7, 1).Value = "Issue Date : " & Format(dtPeriodFrom.Value, "dd MMM yyyy")

                .Cells(9, 8).Value = Format(DateAdd(DateInterval.Month, 1 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(9, 9).Value = Format(DateAdd(DateInterval.Month, 2 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(9, 10).Value = Format(DateAdd(DateInterval.Month, 3 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(9, 11).Value = Format(DateAdd(DateInterval.Month, 4 - 1, dtPeriodFrom.Value), "MMM")

                'For irow = 0 To pData.Rows.Count - 1
                '    For icol = 1 To pData.Columns.Count
                '        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                '    Next
                'Next

                .Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(10, 1, pData.Rows.Count + 9, 12).AutoFitColumns()

                Dim rgAll As ExcelRange = .Cells(10, 1, pData.Rows.Count + 9, 12)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub epplusExportExcelOld(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Summary Forecast " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
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
                .Cells(5, 1).Value = "(" & Format(dtPeriodFrom.Value, "MMM yyyy") & " Production)"
                .Cells(7, 1).Value = "Issue Date : " & Format(dtPeriodFrom.Value, "dd MMM yyyy")

                .Cells(9, 8).Value = Format(DateAdd(DateInterval.Month, 1 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(9, 9).Value = Format(DateAdd(DateInterval.Month, 2 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(9, 10).Value = Format(DateAdd(DateInterval.Month, 3 - 1, dtPeriodFrom.Value), "MMM")
                .Cells(9, 11).Value = Format(DateAdd(DateInterval.Month, 4 - 1, dtPeriodFrom.Value), "MMM")

                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                    Next
                Next

                Dim rgAll As ExcelRange = .Cells(10, 1, irow + 9, 12)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

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
#End Region
End Class