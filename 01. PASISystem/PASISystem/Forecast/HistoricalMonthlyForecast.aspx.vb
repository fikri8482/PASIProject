Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel
Imports DevExpress.XtraPrinting
Imports DevExpress.Utils


Public Class HistoricalMonthlyForecast
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "G01"

    Const colNo As Byte = 1
    Const colPeriod As Byte = 2
    Const colPONo As Byte = 3
    Const colAffiliateCode As Byte = 4
    Const colSupplierCode As Byte = 5
    Const colPOKanban As Byte = 6
    Const colKanbanNo As Byte = 7
    Const colSupplierPlanDelDate As Byte = 8
    Const colPartNo As Byte = 9
    Const colPartName As Byte = 10
    Const colQtyPO As Byte = 11
    Const colSupplierDelDate As Byte = 12
    Const colSupplierSJNo As Byte = 13
    Const colSupplierDeliveryQty As Byte = 14
    Const colPASIRecDate As Byte = 15
    Const colPASIReceivingQty As Byte = 16
    Const colInvoiceNoFromSupplier As Byte = 17
    Const colInvoiceDateFromSupplier As Byte = 18
    Const colInvoiceFromSupplierCurr As Byte = 19
    Const colInvoiceFromSupplierAmount As Byte = 20
    Const colPASIDelDate As Byte = 21
    Const colPASISJNo As Byte = 22
    Const colPASIDeliveryQty As Byte = 23
    Const colAffiliateRecDate As Byte = 24
    Const colAffiliateReceivingQty As Byte = 25
    Const colInvoiceNoToAffiliate As Byte = 26
    Const colInvoiceDateToAffiliate As Byte = 27
    Const colInvoiceToAffiliateCurr As Byte = 28
    Const colInvoiceToAffiliateAmount As Byte = 29
    Const colCount As Byte = 30

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "if (cboAffiliateID.GetItemCount() > 1) { " & vbCrLf & _
            "   cboAffiliateID.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "if (cboSupplier.GetItemCount() > 1) { " & vbCrLf & _
            "   cboSupplier.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "if (cbopart.GetItemCount() > 1) { " & vbCrLf & _
            "   cbopart.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "if (cboRev.GetItemCount() > 1) { " & vbCrLf & _
            "   cboRev.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            " " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriod.SetDate(PeriodTo); " & vbCrLf & _
            "lblInfo.SetText(''); "

        '"if (cboSupplier.GetItemCount() > 1) { " & vbCrLf & _
        '"   cboSupplier.SetValue('==ALL=='); " & vbCrLf & _
        '"} " & vbCrLf & _
        '" " & vbCrLf & _

        ScriptManager.RegisterStartupScript(dtPOPeriod, dtPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Function GetHeaderBandGrid(pPeriod As String, ByVal pConStr As String, Optional ByRef pErr As String = "") As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Try
            Using con = New SqlConnection(pConStr)
                con.Open()
                'sql = "SELECT SeqNo,Nama FROM dbo.fn_GetMonthYear('" & pPeriod & "') A "
                sql = "sp_HistoricalMonthlyForecast_Select_GetHeaderBandGrid"

                cmd = New SqlCommand
                cmd.Parameters.AddWithValue("Period", pPeriod)
                cmd.CommandText = sql
                cmd.Connection = con
                cmd.CommandType = CommandType.StoredProcedure

                da = New SqlDataAdapter(cmd)
                da.Fill(dt)
                cmd.Dispose()
            End Using
            Return dt
        Catch ex As Exception
            Throw New Exception("Generate Month  --> " & ex.Message)
        End Try
    End Function


    Private Sub up_GridHeader(pPeriod As String)
        Dim dt As DataTable
        dt = GetHeaderBandGrid(pPeriod, clsGlobal.ConnectionString)
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                grid.VisibleColumns(8 + i).Caption = dt.Rows(i).Item("Nama").ToString
            Next

        End If


        '    grid.VisibleColumns(8).Caption = "JUL "
        '    grid.VisibleColumns(9).Caption = "AUG "
        '    grid.VisibleColumns(10).Caption = "SEP "
        '    grid.VisibleColumns(11).Caption = "OCT "
        '    grid.VisibleColumns(12).Caption = "NOV "
        '    grid.VisibleColumns(13).Caption = "DEC "

        '    grid.VisibleColumns(14).Caption = "JAN "
        '    grid.VisibleColumns(15).Caption = "FEB "
        '    grid.VisibleColumns(16).Caption = "MAR "
        '    grid.VisibleColumns(17).Caption = "APR "
        '    grid.VisibleColumns(18).Caption = "MAY "
        '    grid.VisibleColumns(19).Caption = "JUN "
    End Sub

    Private Sub up_GridLoad()
        up_GridHeader(Format(dtPOPeriod.Value, "yyyyMM") & "01")

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim ls_filter As String = ""
            ls_SQL = ""

            ls_SQL = "EXEC sp_HistoricalMonthlyForecast_Select '" & (Format(dtPOPeriod.Value, "yyyyMM") & "01") & "','" & cboAffiliateID.Text & "','" & cboSupplier.Text & "', '" & cbopart.Text & "','" & cboRev.Text & "' "

            Dim cmd As New SqlCommand(ls_SQL, sqlConn)
            cmd.CommandTimeout = 300
            Dim sqlDA As New SqlDataAdapter
            sqlDA.SelectCommand = cmd
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 300
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Function GetData() As DataSet
        Try

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Dim ls_filter As String = ""
                ls_SQL = ""

                ls_SQL = "EXEC sp_HistoricalMonthlyForecast_Select_Trial '" & (Format(dtPOPeriod.Value, "yyyyMM") & "01") & "','" & cboAffiliateID.Text & "','" & cboSupplier.Text & "', '" & cbopart.Text & "','" & cboRev.Text & "' "

                Dim cmd As New SqlCommand(ls_SQL, sqlConn)
                cmd.CommandTimeout = 300
                Dim sqlDA As New SqlDataAdapter
                sqlDA.SelectCommand = cmd
                Dim ds As New DataSet
                sqlDA.SelectCommand.CommandTimeout = 300
                sqlDA.Fill(ds)

                Return ds
            End Using
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    'Private Sub up_FillCombo(ByVal pYear As String)
    '    ls_SQL = ""
    '    ls_SQL = "SELECT Distinct RTRIM(FM.PartNo) PartCode, PartName  " & vbCrLf & _
    '             "from ForecastMonthly FM" & vbCrLf & _
    '             "left join MS_Parts MP ON FM.PartNo = MP.PartNo" & vbCrLf & _
    '             "Where Year = '" & pYear & "'" & vbCrLf & _
    '             "order by PartCode" & vbCrLf
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        With cbopart
    '            .Items.Clear()
    '            .Columns.Clear()
    '            .DataSource = ds.Tables(0)
    '            .Columns.Add("PartCode")
    '            .Columns(0).Width = 90
    '            .Columns.Add("PartName")
    '            .Columns(1).Width = 400

    '            .TextField = "PartCode"
    '            .DataBind()
    '            .SelectedIndex = 0
    '            'txtPartNo.Text = clsGlobal.gs_All
    '        End With

    '        sqlConn.Close()
    '    End Using
    'End Sub

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Affiliate
        With cboAffiliateID
            ls_SQL = "SELECT AffiliateID = '==ALL==', AffiliateName = '==ALL=='" & vbCrLf & _
                     "UNION ALL " & vbCrLf & _
                     "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate " 'Where isnull(overseascls, '0') = '0'"
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

        'Combo supplier
        With cboSupplier
            '"--SELECT SupplierID = '==ALL==', SupplierName = '==ALL=='" & vbCrLf & _
            '"--UNION ALL " & vbCrLf & _

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

        'Combo partno
        With cbopart
            ls_SQL = "SELECT PartNo = '==ALL==', PartName = '==ALL=='" & vbCrLf & _
                     "UNION ALL " & vbCrLf & _
                     "SELECT RTRIM(PartNo) PartNo,RTRIM(PartName) PartName FROM dbo.MS_Parts"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 240

                .TextField = "PartNo"
                .DataBind()
            End Using
        End With

        'Combo revision
        With cboRev
            ls_SQL = " SELECT Rev = '==ALL==' " & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = '0'" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = '1'" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = '2'" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = '3' " & vbCrLf

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Rev")
                .Columns(0).Width = 90
                '.Columns.Add("SupplierName")
                '.Columns(1).Width = 240

                .TextField = "Rev"
                .DataBind()
            End Using
        End With

    End Sub

    Private Function uf_ColorCls(ByVal pPeriod As Date, ByVal pAffiliate As String, ByVal pRev As Integer, ByVal pPartNo As String, ByVal pTgl As Integer) As Boolean
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = "Select C" & pTgl & " From ForecastDaily Where Period = '" & pPeriod & "' And AffiliateID = '" & pAffiliate & "' And Rev = '" & pRev & "' And PartNo = '" & pPartNo & "' "
            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlDA As New SqlDataAdapter(sqlCmd)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("C" & pTgl & "") = "1" Then
                    Return True
                End If
            Else
                Return False
            End If
        End Using
    End Function

    'Private Function GetSummaryOutStanding() As DataTable
    '    Dim ls_sql As String = ""
    '    Dim ls_filter As String = ""

    '    Try
    '        Dim clsGlobal As New clsGlobal
    '        Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '            cn.Open()
    '            Dim sql As String = ""

    '            'SUPPLIER CODE
    '            If Trim(cbopart.Text) <> "==ALL==" And Trim(cbopart.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND FD.PartNo = '" & Trim(cbopart.Text) & "' " & vbCrLf
    '            End If

    '            ls_sql = " Select FD.* " & vbCrLf & _
    '                  " From ForecastMonthly FD " & vbCrLf & _
    '                  " Left Join MS_PartMapping MPM ON FD.PartNo = MPM.PartNo And FD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
    '                  " Left Join MS_Parts MP ON FD.PartNo = MP.PartNo " & vbCrLf & _
    '                  " Where Year = '" & dtPOPeriod.Text & "' " & vbCrLf & _
    '                  "  "


    '            ls_sql = ls_sql + ls_filter & vbCrLf

    '            ls_sql = ls_sql + " " & vbCrLf & _
    '                              " " & vbCrLf


    '            Dim Cmd As New SqlCommand(ls_sql, cn)
    '            Dim da As New SqlDataAdapter(Cmd)
    '            Dim dt As New DataTable
    '            da.SelectCommand.CommandTimeout = 300
    '            da.Fill(dt)

    '            Return dt
    '        End Using
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function

    'Private Function GetSummaryOutStanding2() As DataSet
    '    Dim ls_sql As String = ""
    '    Dim ls_filter As String = ""

    '    Try
    '        Dim clsGlobal As New clsGlobal
    '        Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '            cn.Open()
    '            Dim sql As String = ""

    '            ls_sql = "EXEC sp_SelectForecastMonthly '" & dtPOPeriod.Text & "','" & cbopart.Text & "'"

    '            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
    '            Dim ds As New DataSet
    '            sqlDA.SelectCommand.CommandTimeout = 200
    '            sqlDA.Fill(ds)
    '            Return ds
    '        End Using
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function

    Private Function uf_NoChar(ByVal iNo As Integer)
        Dim ls_char As String = ""

        If iNo <= 25 Then
            ls_char = Chr(65 + iNo)
        Else
            ls_char = Chr(65 + (iNo - 26))
        End If

        uf_NoChar = ls_char
    End Function

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT  Top 0 " & vbCrLf & _
                  "  	 Period = '' " & vbCrLf & _
                  "  	,PONo = '' " & vbCrLf & _
                  "  	,AffiliateID = '' " & vbCrLf & _
                  "  	,SupplierID = '' " & vbCrLf & _
                  "  	,POKanban = '' " & vbCrLf & _
                  "  	,PASISendAffiliateDate = '' " & vbCrLf & _
                  "  	,PartNo = '' " & vbCrLf & _
                  "  	,PartName = '' " & vbCrLf & _
                  "  	,QtyPO = '' " & vbCrLf & _
                  " 	,QtyBox = '' "

            ls_SQL = ls_SQL + " 	,BoxPallet = '' " & vbCrLf & _
                              " 	,VolumePallet = '' " & vbCrLf & _
                              "  	,ETDSupp = '' " & vbCrLf & _
                              "  	,ETAAff = '' " & vbCrLf & _
                              "  	,SupplierDeliveryDate = '' " & vbCrLf & _
                              "  	,SupplierSuratJalanNo = '' " & vbCrLf & _
                              "  	,SupplierDeliveryQty = '' " & vbCrLf & _
                              " 	,PASIReceiveDate = '' " & vbCrLf & _
                              "  	,PASIReceivingQty = '' " & vbCrLf & _
                              " 	,Remaining = '' " & vbCrLf & _
                              " 	,StatusPO = '' "


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

    Private Sub InsertHeader(ByVal pExl As ExcelWorksheet)
        With pExl
            .Cells(5, 1, 5, 1).Value = "HISTORICAL MONTHLY FORESCAST"
            .Cells(5, 1, 5, 1).Style.HorizontalAlignment = HorzAlignment.Default
            .Cells(5, 1, 5, 1).Style.VerticalAlignment = VertAlignment.Center
            .Cells(5, 1, 5, 1).Style.Font.Bold = True
            .Cells(5, 1, 5, 1).Style.Font.Size = 14
            .Cells(5, 1, 5, 1).Style.Font.Name = "Arial"

            .Cells(7, 1, 7, 1).Value = "Download Date , " & Format(Now.Date, "dd MMM yyyy")
            .Cells(7, 1, 7, 1).Style.HorizontalAlignment = HorzAlignment.Default
            .Cells(7, 1, 7, 1).Style.VerticalAlignment = VertAlignment.Default
            .Cells(7, 1, 7, 1).Style.Font.Bold = False
            .Cells(7, 1, 7, 1).Style.Font.Size = 10
            .Cells(7, 1, 7, 1).Style.Font.Name = "Arial"


        End With
    End Sub

    Private Sub FormatExcel(ByVal pExl As ExcelWorksheet, ByVal ds As DataSet)
        With pExl
            '.Column(1).Style.Fill.BackgroundColor.SetColor(Color.AliceBlue)
            .Column(1).Width = 8
            .Column(2).Width = 15
            .Column(3).Width = 15
            .Column(4).Width = 15
            .Column(5).Width = 20
            .Column(6).Width = 12
            .Column(7).Width = 15
            .Column(8).Width = 28
            .Column(9).Width = 12
            .Column(10).Width = 12 '1
            '.Column(11).Width = 12 '2
            '.Column(12).Width = 12 '3
            '.Column(13).Width = 12 '4
            '.Column(14).Width = 12 '5
            '.Column(15).Width = 12 '6
            '.Column(16).Width = 12 '7
            '.Column(17).Width = 12 '8
            '.Column(18).Width = 12 '9
            '.Column(19).Width = 12 '10
            '.Column(20).Width = 12 '11
            '.Column(21).Width = 12 '12

            Dim rgAll As ExcelRange = .Cells(9, 1, ds.Tables(0).Rows.Count + 9, 10) '21)
            EpPlusDrawAllBorders(rgAll)

            Dim rgAF As ExcelRange = .Cells(9, 1, ds.Tables(0).Rows.Count + 9, 10) '21)
            rgAF.AutoFilter = True

            Dim rgHeader As ExcelRange = .Cells(9, 1, 9, 10) '21)
            rgHeader.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            rgHeader.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            rgHeader.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            rgHeader.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 210, 166))

        End With
    End Sub

#End Region

#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                'Call up_GridLoadWhenEventChange()

                Call up_Initialize()
                Call up_FillCombo()
                Call up_GridHeader(Format(Now.Date, "yyyyMM") & "01")
                'GridViewFeaturesHelper.SetupGlobalGridViewBehavior(grid)

            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 5, False, clsAppearance.PagerMode.ShowPager)

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
                    Call up_GridHeader(Format(dtPOPeriod.Value, "yyyyMM") & "01")

                Case "excel"
                    'up_LoadExcel()

                    'Dim psERR As String = ""
                    'Dim dtProd As DataSet = GetData()
                    'FileName = "HistoricalMonthlyForecast.xlsx"
                    'FilePath = Server.MapPath("~\Template\" & FileName)
                    If grid.VisibleRowCount > 0 Then
                        Call epplusExportExcel()
                        '    'Call epplusExportExcelNew(FilePath, "Sheet1", dtProd, "A:9", psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("G01Msg") = lblInfo.Text
        End Try

        If (Not IsNothing(Session("G01Msg"))) Then grid.JSProperties("cpMessage") = Session("G01Msg") : Session.Remove("G01Msg")

    End Sub

    Private Sub grid_CustomColumnDisplayText(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewColumnDisplayTextEventArgs) Handles grid.CustomColumnDisplayText
        'With e.Column
        '    For i = 1 To 12
        '        If .FieldName = "F" & i Then
        '            If e.GetFieldValue("Data") = "Diff Firm vs Act" Then
        '                e.DisplayText = e.GetFieldValue("F" & i) & " %"
        '            End If
        '            If e.GetFieldValue("Data") = "Diff Firm vs Last FC" Then
        '                e.DisplayText = e.GetFieldValue("F" & i) & " %"
        '            End If
        '        End If
        '    Next
        'End With
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        'rev
        With e.DataColumn
            If .FieldName = "1" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("1")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If

                End If

            End If
            If .FieldName = "2" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("2")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "3" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("3")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "4" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("4")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "5" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("5")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "6" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("6")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "7" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("7")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "8" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("8")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "9" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("9")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "10" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("10")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "11" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("11")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If
            If .FieldName = "12" Then
                If e.GetValue("Rev").ToString() = "0" Then
                    If Not IsDBNull(e.GetValue("12")) Then
                        e.Cell.BackColor = Color.LimeGreen
                    End If
                End If
            End If

            'actual order and delivery order
            If .FieldName = "1" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("1")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "2" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("2")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "3" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("3")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "4" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("4")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "5" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("5")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "6" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("6")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "7" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("7")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "8" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("8")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "9" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("9")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "10" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("10")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "11" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("11")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

            If .FieldName = "12" Then
                If e.GetValue("Data").ToString() = "Actual Order" Or e.GetValue("Data").ToString() = "Actual Delivery" Then
                    If Not IsDBNull(e.GetValue("12")) Then
                        e.Cell.BackColor = Color.MediumSeaGreen
                    End If
                End If
            End If

        End With
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

    'Private Sub cboPart_Callback(sender As Object, e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cbopart.Callback
    '    Dim pAction As String = Split(e.Parameter, "|")(0)

    '    Call up_FillCombo(pAction)
    '    Call up_GridHeader(pAction)
    'End Sub

    'ByVal pFilename As String, ByVal pSheetName As String,ByVal pData As DataTable, ByVal pCellStart As String,
    Private Sub epplusExportExcel( Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "HistoricalMonthlyForecast.xlsx" '" & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
            Dim fi As New FileInfo(Server.MapPath("~\Forecast\Import\HistoricalMonthlyForecast.xlsx"))
            If fi.Exists Then
                fi.Delete()
                fi = New FileInfo(Server.MapPath("~\Forecast\Import\HistoricalMonthlyForecast.xlsx"))
            End If
            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet
            ws = exl.Workbook.Worksheets.Add("Historical Monthly Forecast")
            ws.View.ShowGridLines = True

            ws.View.FreezePanes(10, 1)

            With ws
                .Cells(9, 1, 9, 21).Style.Font.Bold = False

                .Cells(9, 1, 9, 1).Value = "No"
                .Cells(9, 1, 9, 1).Style.HorizontalAlignment = HorzAlignment.Default
                .Cells(9, 1, 9, 1).Style.VerticalAlignment = VertAlignment.Center
                '.Cells9, 1, 9, 1).Merge = True
                .Cells(9, 1, 9, 1).Style.Font.Size = 10
                .Cells(9, 1, 9, 1).Style.Font.Name = "Arial"



                .Cells(9, 2, 9, 2).Value = "Affiliate"
                .Cells(9, 2, 9, 2).Style.HorizontalAlignment = HorzAlignment.Default
                .Cells(9, 2, 9, 2).Style.VerticalAlignment = VertAlignment.Center
                '.Cells9, 2, 9, 2).Merge = True
                .Cells(9, 2, 9, 2).Style.Font.Size = 10
                .Cells(9, 2, 9, 2).Style.Font.Name = "Arial"


                .Cells(9, 3, 9, 3).Value = "Supplier"
                .Cells(9, 3, 9, 3).Style.HorizontalAlignment = HorzAlignment.Far
                '.Cells9, 3, 9, 3).Merge = True
                .Cells(9, 3, 9, 3).Style.VerticalAlignment = VertAlignment.Center
                .Cells(9, 3, 9, 3).Style.Font.Size = 10
                .Cells(9, 3, 9, 3).Style.Font.Name = "Arial"


                .Cells(9, 4, 9, 4).Value = "Part No"
                .Cells(9, 4, 9, 4).Style.HorizontalAlignment = HorzAlignment.Far
                .Cells(9, 4, 9, 4).Style.VerticalAlignment = VertAlignment.Center
                '.Cells(4,4,9, 4).Merge = True
                .Cells(9, 4, 9, 4).Style.Font.Size = 10
                .Cells(9, 4, 9, 4).Style.Font.Name = "Arial"


                .Cells(9, 5, 9, 5).Value = "Project"
                .Cells(9, 5, 9, 5).Style.HorizontalAlignment = HorzAlignment.Far
                .Cells(9, 5, 9, 5).Style.VerticalAlignment = VertAlignment.Center
                '.Cells94, 5,94, 5).Merge = True
                .Cells(9, 5, 9, 5).Style.Font.Size = 10
                .Cells(9, 5, 9, 5).Style.Font.Name = "Arial"


                .Cells(9, 6, 9, 6).Value = "MPQ"
                .Cells(9, 6, 9, 6).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(9, 6, 9, 6).Style.VerticalAlignment = VertAlignment.Center
                '.Cells94, 6,9, 6).Merge = True
                .Cells(9, 6, 9, 6).Style.Font.Size = 10
                .Cells(9, 6, 9, 6).Style.Font.Name = "Arial"


                .Cells(9, 7, 9, 7).Value = "Issue Date"
                .Cells(9, 7, 9, 7).Style.HorizontalAlignment = HorzAlignment.Far
                .Cells(9, 7, 9, 7).Style.VerticalAlignment = VertAlignment.Center
                '.Cells94, 7,94, 7).Merge = True
                .Cells(9, 7, 9, 7).Style.Font.Size = 10
                .Cells(9, 7, 9, 7).Style.Font.Name = "Arial"


                .Cells(9, 8, 9, 8).Value = "Data"
                .Cells(9, 8, 9, 8).Style.HorizontalAlignment = HorzAlignment.Far
                .Cells(9, 8, 9, 8).Style.VerticalAlignment = VertAlignment.Center
                '.Cells94, 8,94, 8).Merge = True
                .Cells(9, 8, 9, 8).Style.Font.Size = 10
                .Cells(9, 8, 9, 8).Style.Font.Name = "Arial"


                .Cells(9, 9, 9, 9).Value = "Revision"
                .Cells(9, 9, 9, 9).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(9, 9, 9, 9).Style.VerticalAlignment = VertAlignment.Center
                ' .Cell9(4, 99 4, 9).Merge = True
                .Cells(9, 9, 9, 9).Style.Font.Size = 10
                .Cells(9, 9, 9, 9).Style.Font.Name = "Arial"


                Dim dt As DataTable
                dt = GetHeaderBandGrid(Format(dtPOPeriod.Value, "yyyyMM") & "01", clsGlobal.ConnectionString)
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        .Cells(9, 10 + i, 9, 10 + i).Value = dt.Rows(i).Item("Nama").ToString
                        .Cells(9, 10 + i, 9, 10 + i).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                        .Cells(9, 10 + i, 9, 10 + i).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
                        .Cells(9, 10 + i, 9, 10 + i).Style.Font.Size = 10
                        .Cells(9, 10 + i, 9, 10 + i).Style.Font.Name = "Arial"


                    Next

                End If
               
                Dim ErrMsg As String = ""
                'Dim dtFrom = Format(dtDateFrom.Value, "yyyy-MM-dd")
                'Dim dtTo = Format(dtDateTo.Value, "yyyy-MM-dd")

                Dim ds As DataSet
                ds = GetData()

                Dim rgHeader As ExcelRange
                Dim No As Integer, RNo As Integer = 0
                Dim PartNoTemp As String = ""
                No = 1
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    'no
                    If PartNoTemp <> ds.Tables(0).Rows(i)("PartNoSort").ToString().Trim Then
                        RNo = RNo + 1
                        .Cells(i + 10, 1, i + 10, 1).Value = RNo
                    Else
                        .Cells(i + 10, 1, i + 10, 1).Value = ""
                    End If


                    .Cells(i + 10, 2, i + 10, 2).Value = ds.Tables(0).Rows(i)("AffiliateID").ToString().Trim
                    .Cells(i + 10, 3, i + 10, 3).Value = ds.Tables(0).Rows(i)("SupplierID").ToString().Trim
                    .Cells(i + 10, 4, i + 10, 4).Value = ds.Tables(0).Rows(i)("PartNo").ToString().Trim
                    .Cells(i + 10, 5, i + 10, 5).Value = ds.Tables(0).Rows(i)("Project").ToString().Trim

                    If (IsDBNull(ds.Tables(0).Rows(i)("MPQ"))) Then
                        .Cells(i + 10, 6, i + 10, 6).Value = ds.Tables(0).Rows(i)("MPQ").ToString()
                    Else
                        .Cells(i + 10, 6, i + 10, 6).Value = Format(CDbl(ds.Tables(0).Rows(i)("MPQ").ToString), "###,##")
                    End If
                    .Cells(i + 10, 6, i + 10, 6).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    '.Cells(i + 10, 6, i + 10, 6).Value = IIf(IsDBNull(ds.Tables(0).Rows(i)("MPQ")), ds.Tables(0).Rows(i)("MPQ").ToString, Format(CDbl(ds.Tables(0).Rows(i)("MPQ").ToString), "###,##"))
                    .Cells(i + 10, 7, i + 10, 7).Value = ds.Tables(0).Rows(i)("IssueDate")
                    .Cells(i + 10, 8, i + 10, 8).Value = ds.Tables(0).Rows(i)("Data").ToString().Trim
                    .Cells(i + 10, 9, i + 10, 9).Value = ds.Tables(0).Rows(i)("Rev").ToString().Trim
                    .Cells(i + 10, 9, i + 10, 9).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center

                    Dim bln As String
                    For icol = 0 To 0'11
                        bln = CStr(icol + 1)
                        If (IsDBNull(ds.Tables(0).Rows(i)(bln))) Then
                            .Cells(i + 10, 10 + icol, i + 10, 10 + icol).Value = ds.Tables(0).Rows(i)(bln).ToString()
                        Else
                            .Cells(i + 10, 10 + icol, i + 10, 10 + icol).Value = IIf(ds.Tables(0).Rows(i)(bln).ToString = "0", "0", Format(CDbl(ds.Tables(0).Rows(i)(bln).ToString), "###,##"))
                        End If
                        .Cells(i + 10, 10 + icol, i + 10, 10 + icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    Next

                    'If (IsDBNull(ds.Tables(0).Rows(i)("1"))) Then
                    '    .Cells(i + 10, 10, i + 10, 10).Value = ds.Tables(0).Rows(i)("1").ToString()
                    'Else
                    '    .Cells(i + 10, 10, i + 10, 10).Value = Format(CDbl(ds.Tables(0).Rows(i)("1").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 10, i + 10, 10).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("2"))) Then
                    '    .Cells(i + 10, 11, i + 10, 11).Value = ds.Tables(0).Rows(i)("2").ToString()
                    'Else
                    '    .Cells(i + 10, 11, i + 10, 11).Value = Format(CDbl(ds.Tables(0).Rows(i)("2").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 11, i + 10, 11).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("3"))) Then
                    '    .Cells(i + 10, 12, i + 10, 12).Value = ds.Tables(0).Rows(i)("3").ToString()
                    'Else
                    '    .Cells(i + 10, 12, i + 10, 12).Value = Format(CDbl(ds.Tables(0).Rows(i)("3").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 12, i + 10, 12).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("4"))) Then
                    '    .Cells(i + 10, 13, i + 10, 13).Value = ds.Tables(0).Rows(i)("4").ToString()
                    'Else
                    '    .Cells(i + 10, 13, i + 10, 13).Value = Format(CDbl(ds.Tables(0).Rows(i)("4").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 13, i + 10, 13).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("5"))) Then
                    '    .Cells(i + 10, 14, i + 10, 14).Value = ds.Tables(0).Rows(i)("5").ToString()
                    'Else
                    '    .Cells(i + 10, 14, i + 10, 14).Value = Format(CDbl(ds.Tables(0).Rows(i)("5").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 14, i + 10, 14).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("6"))) Then
                    '    .Cells(i + 10, 15, i + 10, 15).Value = ds.Tables(0).Rows(i)("6").ToString()
                    'Else
                    '    .Cells(i + 10, 15, i + 10, 15).Value = Format(CDbl(ds.Tables(0).Rows(i)("6").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 15, i + 10, 15).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("7"))) Then
                    '    .Cells(i + 10, 16, i + 10, 16).Value = ds.Tables(0).Rows(i)("7").ToString()
                    'Else
                    '    .Cells(i + 10, 16, i + 10, 16).Value = Format(CDbl(ds.Tables(0).Rows(i)("7").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 16, i + 10, 16).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("8"))) Then
                    '    .Cells(i + 10, 17, i + 10, 17).Value = ds.Tables(0).Rows(i)("8").ToString()
                    'Else
                    '    .Cells(i + 10, 17, i + 10, 17).Value = Format(CDbl(ds.Tables(0).Rows(i)("8").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 17, i + 10, 17).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("9"))) Then
                    '    .Cells(i + 10, 18, i + 10, 18).Value = ds.Tables(0).Rows(i)("9").ToString()
                    'Else
                    '    .Cells(i + 10, 18, i + 10, 18).Value = Format(CDbl(ds.Tables(0).Rows(i)("9").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 18, i + 10, 18).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("10"))) Then
                    '    .Cells(i + 10, 19, i + 10, 19).Value = ds.Tables(0).Rows(i)("10").ToString()
                    'Else
                    '    .Cells(i + 10, 19, i + 10, 19).Value = Format(CDbl(ds.Tables(0).Rows(i)("10").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 19, i + 10, 19).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("11"))) Then
                    '    .Cells(i + 10, 20, i + 10, 20).Value = ds.Tables(0).Rows(i)("11").ToString()
                    'Else
                    '    .Cells(i + 10, 20, i + 10, 20).Value = Format(CDbl(ds.Tables(0).Rows(i)("11").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 20, i + 10, 20).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    'If (IsDBNull(ds.Tables(0).Rows(i)("12"))) Then
                    '    .Cells(i + 10, 21, i + 10, 21).Value = ds.Tables(0).Rows(i)("12").ToString()
                    'Else
                    '    .Cells(i + 10, 21, i + 10, 21).Value = Format(CDbl(ds.Tables(0).Rows(i)("12").ToString), "###,##")
                    'End If
                    '.Cells(i + 10, 21, i + 10, 21).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    
                    'coloring cell
                    Dim dt2 As DataTable
                    Dim bulan As String
                    dt2 = GetHeaderBandGrid(Format(dtPOPeriod.Value, "yyyyMM") & "01", clsGlobal.ConnectionString)
                    If dt.Rows.Count > 0 Then
                        For i2 = 0 To dt2.Rows.Count - 1
                            bulan = CStr(i2 + 1)
                            If dt2.Rows(i2).Item("SeqNo").ToString() = bulan Then
                                If ds.Tables(0).Rows(i)("Rev").ToString().Trim = "0" Then
                                    If Not IsDBNull(ds.Tables(0).Rows(i)(bulan)) Then
                                        rgHeader = .Cells(i + 10, 10 + i2, i + 10, 10 + i2)
                                        rgHeader.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                        rgHeader.Style.Fill.BackgroundColor.SetColor(Color.LimeGreen)
                                    End If
                                ElseIf ds.Tables(0).Rows(i)("Rev").ToString().Trim = "1" Or _
                                       ds.Tables(0).Rows(i)("Rev").ToString().Trim = "2" Or _
                                       ds.Tables(0).Rows(i)("Rev").ToString().Trim = "3" Then

                                    If Not IsDBNull(ds.Tables(0).Rows(i)(bulan)) Then
                                        rgHeader = .Cells(i + 10, 10 + i2, i + 10, 10 + i2)
                                        rgHeader.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                        rgHeader.Style.Fill.BackgroundColor.SetColor(Color.LemonChiffon)

                                    End If
                                Else
                                    If (ds.Tables(0).Rows(i)("Data").ToString().Trim = "Capacity" Or _
                                        ds.Tables(0).Rows(i)("Data").ToString().Trim = "LTF 78-1" Or _
                                        ds.Tables(0).Rows(i)("Data").ToString().Trim = "LTF 78-2") Then
                                        If Not IsDBNull(ds.Tables(0).Rows(i)(bulan)) Then
                                            rgHeader = .Cells(i + 10, 10 + i2, i + 10, 10 + i2)
                                            rgHeader.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                            rgHeader.Style.Fill.BackgroundColor.SetColor(Color.LemonChiffon)

                                        End If
                                    ElseIf (ds.Tables(0).Rows(i)("Data").ToString().Trim = "Actual Order" Or _
                                            ds.Tables(0).Rows(i)("Data").ToString().Trim = "Actual Delivery") Then
                                        If Not IsDBNull(ds.Tables(0).Rows(i)(bulan)) Then
                                            rgHeader = .Cells(i + 10, 10 + i2, i + 10, 10 + i2)
                                            rgHeader.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                            rgHeader.Style.Fill.BackgroundColor.SetColor(Color.LimeGreen)
                                        End If

                                    End If
                                End If
                            End If


                        Next
                    End If

                    No = No + 1
                    PartNoTemp = ds.Tables(0).Rows(i)("PartNoSort").ToString().Trim

                Next


                FormatExcel(ws, ds)
                InsertHeader(ws)
            End With

            exl.Save()
            'DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("Download/" & fi.Name)

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub up_LoadExcel(Optional ByRef pErr As String = "")
        up_GridLoad()
        Dim ps As New PrintingSystem()

        Dim link1 As New PrintableComponentLink(ps)
        link1.Component = GridExporter

        Dim compositeLink As New XtraPrintingLinks.CompositeLink(ps)
        compositeLink.Links.AddRange(New Object() {link1})

        compositeLink.CreateDocument()
        Using stream As New MemoryStream()
            compositeLink.PrintingSystem.ExportToXlsx(stream)
            Response.Clear()
            Response.Buffer = False
            Response.AppendHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            Response.AppendHeader("Content-Disposition", "attachment; filename=SummaryPurchaseRequest_" & Format(CDate(Now), "yyyyMMdd_hhmmss") & ".xlsx")
            Response.BinaryWrite(stream.ToArray())
            Response.End()
        End Using

        ps.Dispose()

    End Sub

    'Protected Sub btnexcel_Click(sender As Object, e As EventArgs) Handles btnexcel.Click
    '    up_LoadExcel()
    'End Sub
End Class