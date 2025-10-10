Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO

Public Class SummaryInvoiceSupplierExp
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim dt As DataTable

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False

    Dim FilePath = "", ls_SQL = "", FileName As String = ""
#End Region

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_FillCombo()
                Call up_GridLoadWhenEventChange()
                Call up_Initialize()
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("N06Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowPager)

        Try
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("N06Msg") = lblInfo.Text
                    Else
                        grid.PageIndex = 0
                    End If
                Case "clear"
                    Call up_GridLoadWhenEventChange()

                Case "excel"
                    Dim psERR As String = ""
                    dt = CType(Session("N06dTable"), DataTable)
                    FileName = "TemplateSummaryInvoiceSupplierExport.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dt.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dt, "A:8", psERR)
                    Else
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("N06Msg") = lblInfo.Text
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("N06Msg") = lblInfo.Text
        End Try

        If (Not IsNothing(Session("N06Msg"))) Then grid.JSProperties("cpMessage") = Session("N06Msg") : Session.Remove("N06Msg")

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "if (cboAffiliateCode.GetItemCount() > 1) { " & vbCrLf & _
            "   txtAffiliateName.SetText('==ALL=='); " & vbCrLf & _
            "   cboAffiliateCode.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "if (cboSupplierCode.GetItemCount() > 1) { " & vbCrLf & _
            "   txtSupplierName.SetText('==ALL=='); " & vbCrLf & _
            "   cboSupplierCode.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriodFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtPOPeriodTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierDelDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierDelDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "chkSupplierDelDate.SetChecked(false); " & vbCrLf & _
            "dtSupplierDelDateFrom.SetEnabled(false); " & vbCrLf & _
            "dtSupplierDelDateTo.SetEnabled(false); " & vbCrLf & _
            " " & vbCrLf & _
            "txtPONo.SetText(''); " & vbCrLf & _
            " " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(chkPOPeriod, chkPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Affiliate
        With cboAffiliateCode
            ls_SQL = "SELECT AffiliateID = '==ALL==', AffiliateName = '==ALL=='" & vbCrLf & _
                     "UNION ALL " & vbCrLf & _
                     "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate Where isnull(overseascls, '0') = '1'"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 90
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240

                .TextField = "AffiliateID"
                .DataBind()
            End Using
        End With

        'Combo Supplier
        With cboSupplierCode
            ls_SQL = "SELECT SupplierID = '==ALL==', SupplierName = '==ALL=='" & vbCrLf & _
                     "UNION ALL " & vbCrLf & _
                     "SELECT SupplierID = RTRIM(SupplierID), SupplierName = RTRIM(SupplierName) FROM dbo.MS_supplier Where isnull(overseas, '0') = '0'"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 90
                .Columns.Add("SupplierName")
                .Columns(1).Width = 240

                .TextField = "SupplierID"
                .DataBind()
            End Using
        End With

    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim ls_filter As String = ""
            ls_SQL = ""

            Dim ls_End As String = ""
            ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

            ls_SQL = "sp_PASI_SummaryInvoiceSupplierExport_GridLoad"
            Dim cmd As New SqlCommand(ls_SQL, sqlConn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@AffiliateID", Trim(cboAffiliateCode.Text))
            cmd.Parameters.AddWithValue("@SupplierCode", Trim(cboSupplierCode.Text))
            cmd.Parameters.AddWithValue("@PONumber", Trim(txtPONo.Text))

            'AFFILIATE PO PERIOD
            If chkPOPeriod.Checked = True Then
                cmd.Parameters.AddWithValue("@AffiliatePOPeriodFrom", Format(dtPOPeriodFrom.Value, "yyyy-MM-01"))
                cmd.Parameters.AddWithValue("@AffiliatePOPeriodTo", Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End))
            End If

            'SUPPLIER DELIVERY DATE
            If chkSupplierDelDate.Checked = True Then
                cmd.Parameters.AddWithValue("@SupplierDeliveryDateFrom", Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@SupplierDeliveryDateTo", Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd"))
            End If

            cmd.CommandTimeout = 300
            dt = New DataTable()
            Dim sqlDA As New SqlDataAdapter
            sqlDA.SelectCommand = cmd
            sqlDA.SelectCommand.CommandTimeout = 300
            sqlDA.Fill(dt)
            Session("N06dTable") = dt
            With grid
                .DataSource = dt
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "TemplateSummaryInvoiceSupplierExp " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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

            With (ws)
                .Cells(8, 8).Worksheet.View.FreezePanes(8, 8)
                .Cells(3, 3).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                .Cells(4, 3).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 18).AutoFitColumns()
                .Cells(8, 1, pData.Rows.Count + 7, 18).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

                .Cells(8, 1, pData.Rows.Count + 7, 1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(8, 10, pData.Rows.Count + 7, 10).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(8, 13, pData.Rows.Count + 7, 13).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(8, 16, pData.Rows.Count + 7, 16).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

                .Cells(8, 2, pData.Rows.Count + 7, 8).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                .Cells(8, 11, pData.Rows.Count + 7, 11).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                .Cells(8, 15, pData.Rows.Count + 7, 15).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

                .Cells(8, 9, pData.Rows.Count + 7, 9).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(8, 12, pData.Rows.Count + 7, 12).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(8, 14, pData.Rows.Count + 7, 14).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(8, 17, pData.Rows.Count + 7, 17).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(8, 18, pData.Rows.Count + 7, 18).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                .Cells(8, 1, pData.Rows.Count + 7, 1).Style.Numberformat.Format = "mmm-yy"

                .Cells(8, 2, pData.Rows.Count + 7, 3).Style.Numberformat.Format = "@"
                .Cells(8, 11, pData.Rows.Count + 7, 11).Style.Numberformat.Format = "@"
                .Cells(8, 15, pData.Rows.Count + 7, 15).Style.Numberformat.Format = "@"

                .Cells(8, 10, pData.Rows.Count + 7, 10).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 13, pData.Rows.Count + 7, 13).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 16, pData.Rows.Count + 7, 16).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 9, pData.Rows.Count + 7, 9).Style.Numberformat.Format = "#,##0"
                .Cells(8, 12, pData.Rows.Count + 7, 12).Style.Numberformat.Format = "#,##0"
                .Cells(8, 14, pData.Rows.Count + 7, 14).Style.Numberformat.Format = "#,##0"
                .Cells(8, 17, pData.Rows.Count + 7, 18).Style.Numberformat.Format = "#,##0"

                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 18)
                EpPlusDrawAllBorders(rgAll)
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

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT Top 0" & vbCrLf & _
                  " Period = '' " & vbCrLf & _
                  ",PONo = '' " & vbCrLf & _
                  ",OrderNo = '' " & vbCrLf & _
                  ",AffiliateID = '' " & vbCrLf & _
                  ",SupplierID = '' " & vbCrLf & _
                  ",PartNo = '' " & vbCrLf & _
                  ",PartName = '' " & vbCrLf & _
                  ",QtyPO = '' " & vbCrLf & _
                  ",SupplierDeliveryDate = '' " & vbCrLf & _
                  ",SupplierSuratJalanNo = '' " & vbCrLf & _
                  ",SupplierDeliveryQty = '' " & vbCrLf & _
                  ",PASIReceiveDate = '' " & vbCrLf & _
                  ",PASIReceivingQty = '' " & vbCrLf & _
                  ",InvoiceNo = '' " & vbCrLf & _
                  ",InvoiceDate = '' " & vbCrLf & _
                  ",Price = '' " & vbCrLf & _
                  ",Total = '' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            dt = New DataTable()
            sqlDA.Fill(dt)
            Session("N06dTable") = dt
            With grid
                .DataSource = dt
                .DataBind()
            End With
        End Using
    End Sub
#End Region

End Class