Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO

Public Class PDPLBDomestic
    Inherits System.Web.UI.Page

#Region "Declaration"

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Dim ls_SQL As String = ""
    Dim menuID As String = "R01"

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""

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

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowPager)

        Try
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("G01Msg") = lblInfo.Text
                    Else
                        grid.PageIndex = 0
                    End If
                Case "clear"
                    Call up_GridLoadWhenEventChange()
                    Call up_FillCombo()
                    Session("G01Msg") = ""

                Case "excel"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetSummaryOutStandingExcel()
                    FileName = "TemplatePDPLB.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("G01Msg") = lblInfo.Text
        End Try

        If (Not IsNothing(Session("G01Msg"))) Then grid.JSProperties("cpMessage") = Session("G01Msg") : Session.Remove("G01Msg")

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub

    Private Sub cboSuratJalan_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboSuratJalan.Callback
        If String.IsNullOrEmpty(e.Parameter) Then
            Return
        End If

        Dim AffID As String = Split(e.Parameter, "|")(0)
        Dim DeliveryDate As String = Split(e.Parameter, "|")(1)


        Dim datex As String = Format(Convert.ToDateTime(DeliveryDate), "yyyy-MM-dd")

        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        ls_SQL = " Exec PDPLB_FillCombo '3', '" + AffID + "' , '" + datex + "' "
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
            ds = New DataSet
            sqlDA.Fill(ds)

            With cboSuratJalan
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Descrip")
                .Columns(0).Width = 200
                .DataBind()
            End With
        End Using
    End Sub

#End Region


#Region "Functions"
    Private Sub up_Initialize()
        Dim script As String =
            "clear()"

        ScriptManager.RegisterStartupScript(cboAffiliateCode, cboAffiliateCode.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        ls_SQL = "Exec PDPLB_FillCombo '1'"
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
            ds = New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Code")
                .Columns(0).Width = 120
                .Columns.Add("Descrip")
                .Columns(1).Width = 250

                .DataBind()
                .TextField = "Code"
            End With

            With cboRemark
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(1)
                .Columns.Add("Descrip")
                .Columns(0).Width = 150
                .DataBind()
            End With

            With cboKategory
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(2)
                .Columns.Add("Descrip")
                .Columns(0).Width = 150
                .DataBind()
            End With
        End Using

        Dim datex As String = Format(Now, "yyyy-MM-dd")

        ls_SQL = " Exec PDPLB_FillCombo '3', '" + "" + "' , '" + datex + "' "
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
            ds = New DataSet
            sqlDA.Fill(ds)

            With cboSuratJalan
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Descrip")
                .Columns(0).Width = 200
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim dt As DataTable = New DataTable()
        dt = GetSummaryOutStanding()

        With grid
            .DataSource = dt
            .DataBind()
        End With
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " Exec PDPLB_Sel_Blank "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
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

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "SO PASI PBPLB_" & Format(Now, "yyyymmdd") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\PDPLB\Import\" & tempFile & "")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim fi As New FileInfo(NewFileName)
            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets(pSheetName)
            With ws
                .Cells("A2").LoadFromDataTable(DirectCast(pData, DataTable), False)

                .Cells(2, 2, pData.Rows.Count + 1, 2).Style.Numberformat.Format = "yyyy-MM-dd"

                .Cells(2, 6, pData.Rows.Count + 1, 6).Style.Numberformat.Format = "#,##0"
                .Cells(2, 8, pData.Rows.Count + 1, 8).Style.Numberformat.Format = "#,##0"
                .Cells(2, 9, pData.Rows.Count + 1, 9).Style.Numberformat.Format = "#,##0"

                Dim rgAll As ExcelRange = .Cells(2, 1, pData.Rows.Count + 1, 16)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\PDPLB\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                ls_sql = "PDPLB_Sel_Dom"
                Dim Cmd As New SqlCommand(ls_sql, cn)
                Cmd.CommandType = CommandType.StoredProcedure

                Cmd.Parameters.AddWithValue("InvoiceNo", txtInvoice.Text)
                Cmd.Parameters.AddWithValue("RemarkOrder", cboRemark.Text)
                Cmd.Parameters.AddWithValue("Kategory", cboKategory.Text)
                Cmd.Parameters.AddWithValue("SuratJalan", cboSuratJalan.Text)
                Cmd.Parameters.AddWithValue("AffiliateID", cboAffiliateCode.Text)
                Cmd.Parameters.AddWithValue("DeliveryDate", Format(dtDeliveryDate.Value, "yyyy-MM-dd"))

                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 300
                da.Fill(dt)

                Return dt
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function GetSummaryOutStandingExcel() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                ls_sql = "PDPLB_Sel_Dom_Excel"
                Dim Cmd As New SqlCommand(ls_sql, cn)
                Cmd.CommandType = CommandType.StoredProcedure

                Cmd.Parameters.AddWithValue("InvoiceNo", txtInvoice.Text)
                Cmd.Parameters.AddWithValue("RemarkOrder", cboRemark.Text)
                Cmd.Parameters.AddWithValue("Kategory", cboKategory.Text)
                Cmd.Parameters.AddWithValue("SuratJalan", cboSuratJalan.Text)
                Cmd.Parameters.AddWithValue("AffiliateID", cboAffiliateCode.Text)
                Cmd.Parameters.AddWithValue("DeliveryDate", Format(dtDeliveryDate.Value, "yyyy-MM-dd"))

                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 300
                da.Fill(dt)

                Return dt
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

End Class