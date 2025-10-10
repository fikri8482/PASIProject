Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing
Imports System.Transactions
Imports OfficeOpenXml
Imports System.IO

Public Class Palletizing
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
                Session("MenuDesc") = "PALLETIZING"

                    Call up_fillcombo()
                    lblerrmessage.Text = ""
                    grid.JSProperties("cpdtReceivingDate") = Format(Now, "dd MMM yyyy")

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
        ls_sql = "SELECT distinct SupplierID = '" & clsGlobal.gs_All & "', SupplierName = '" & clsGlobal.gs_All & "' from ms_supplier " & vbCrLf & _
                 "UNION Select SupplierID = RTRIM(SupplierID) ,SupplierName = RTRIM(SupplierName) FROM dbo.ms_supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbosupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 70
                .Columns.Add("SupplierName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtsupplier.Text = clsGlobal.gs_All
                .TextField = "SupplierID"
                .DataBind()
            End With
            sqlConn.Close()

            
            'AFFILIATE
            sqlConn.Open()
            ls_sql = "SELECT distinct AffiliateID = '" & clsGlobal.gs_All & "', AffiliateName = '" & clsGlobal.gs_All & "' from MS_Affiliate " & vbCrLf & _
                     "UNION Select AffiliateID = RTRIM(AffiliateID) ,AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate  where isnull(overseascls, '0') = '1'" & vbCrLf
            Dim sqlDA2 As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds2 As New DataSet
            sqlDA2.Fill(ds2)

            With cboaffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds2.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 70
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                txtaffiliate.Text = clsGlobal.gs_All
                .TextField = "AffiliateID"
                .DataBind()
            End With
            sqlConn.Close()

        End Using
    End Sub
    Private Sub up_GridLoadClear()
        Dim ls_Sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_Sql = "SELECT TOP 0 (SELECT '')ACT,(SELECT '')colperiod,(SELECT '') as colaffiliatecode, (SELECT '') as colaffiliatename, " & vbCrLf & _
                     "(SELECT '') as colsuppliercode, (SELECT '') AS colsuppliername, (SELECT '') as colorderno, (SELECT '')as colinvoiveno,  " & vbCrLf & _
                     "(SELECT '') as colpallet,(SELECT '') as colLabelNo,(SELECT '') as colpartno,(SELECT '') as colqty,(SELECT '') as collocation,(SELECT '') as colpallettype, " & vbCrLf & _
                     "(SELECT '') AS colHForwarder,(SELECT '') AS colHsuratjalan "

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
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
       
            'AFF
            If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                ls_Filter = ls_Filter + " AND D.AffiliateID = '" & Trim(cboaffiliate.Text) & "' " & vbCrLf
            End If

            'SUPP
            If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                ls_Filter = ls_Filter + " AND D.SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf
            End If

            'RECEIVE DATE
            If checkboxdt.Checked = True And dtReceivingDate.Text <> "" Then
                ls_Filter = ls_Filter + " AND CONVERT(date,M.ShippingInstructionDate) = '" & Format(dtReceivingDate.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'INVOICENO
            If txtInvoiceNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND D.ShippingInstructionNo = '" & txtInvoiceNo.Text & "' " & vbCrLf
            End If

            'ALREADY SHIPPING
            If rbshipping.Value = "Already" Then
                ls_Filter = ls_Filter + " AND ISNULL(D.ShippingInstructionNo,'') <> '' " & vbCrLf
            ElseIf rbshipping.Value = "Progress" Then
                ls_Filter = ls_Filter + " AND ISNULL(D.ShippingInstructionNo,'') = '' " & vbCrLf
            End If

            'ORDERNO
            If txtorderno.Text <> "" Then
                ls_Filter = ls_Filter + " AND D.OrderNo LIKE '%" & txtorderno.Text & "%' " & vbCrLf
            End If

            'PALLETNO
            If txtPalletNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND DP.PalletNo = '" & txtPalletNo.Text & "' " & vbCrLf
            End If
            'BOXNO
            If txtBoxNo.Text <> "" Then
                ls_Filter = ls_Filter + " AND D.BoxNo = '" & txtBoxNo.Text & "' " & vbCrLf
            End If


            ls_SQL = " SELECT (SELECT '')ACT,CONVERT(date,M.ShippingInstructionDate)colperiod,D.AffiliateID as colaffiliatecode, MA.AffiliateName as colaffiliatename, " & vbCrLf & _
                  " D.SupplierID as colsuppliercode, MS.SupplierName as colsuppliername, D.OrderNo as colorderno, D.ShippingInstructionNo as colinvoiveno,  " & vbCrLf & _
                  " DP.PalletNo as colpallet,DP.LabelNo as colLabelNo,D.PartNo as colpartno,D.QtyBox as colqty,DP.Location as collocation,DP.PalletType as colpallettype," & vbCrLf & _
                  " D.ForwarderID AS colHForwarder,D.SuratJalanNo AS colHsuratjalan " & vbCrLf & _
                  " FROM dbo.ShippingInstruction_DetailPallet DP " & vbCrLf & _
                  " INNER JOIN dbo.ShippingInstruction_Master M ON M.AffiliateID = DP.AffiliateID  " & vbCrLf & _
                  " AND M.ForwarderID = DP.ForwarderID AND M.ShippingInstructionNo = DP.ShippingInstructionNo " & vbCrLf & _
                  " INNER JOIN dbo.ShippingInstruction_Detail D ON D.AffiliateID = DP.AffiliateID  " & vbCrLf & _
                  " AND D.ForwarderID = M.ForwarderID AND D.ShippingInstructionNo = DP.ShippingInstructionNo " & vbCrLf & _
                  " INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = D.SupplierID " & vbCrLf & _
                  " INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = D.AffiliateID  " & vbCrLf & _
                  " Where ISNULL(DP.PalletNo,'') <> ''"

            ls_SQL = ls_SQL + ls_Filter

            'ls_SQL = "Select * from ShippingInstruction_Detail"
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

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                'AFF
                If cboaffiliate.Text <> clsGlobal.gs_All And cboaffiliate.Text <> "" Then
                    ls_filter = ls_filter + " AND D.AffiliateID = '" & Trim(cboaffiliate.Text) & "' " & vbCrLf
                End If

                'SUPP
                If cbosupplier.Text <> clsGlobal.gs_All And cbosupplier.Text <> "" Then
                    ls_filter = ls_filter + " AND D.SupplierID = '" & Trim(cbosupplier.Text) & "' " & vbCrLf
                End If

                'RECEIVE DATE
                If checkboxdt.Checked = True And dtReceivingDate.Text <> "" Then
                    ls_filter = ls_filter + " AND CONVERT(date,M.ShippingInstructionDate) = '" & Format(dtReceivingDate.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                'INVOICENO
                If txtInvoiceNo.Text <> "" Then
                    ls_filter = ls_filter + " AND D.ShippingInstructionNo = '" & txtInvoiceNo.Text & "' " & vbCrLf
                End If

                'ALREADY SHIPPING
                If rbshipping.Value = "Already" Then
                    ls_filter = ls_filter + " AND ISNULL(D.ShippingInstructionNo,'') <> '' " & vbCrLf
                ElseIf rbshipping.Value = "Progress" Then
                    ls_filter = ls_filter + " AND ISNULL(D.ShippingInstructionNo,'') = '' " & vbCrLf
                End If

                'ORDERNO
                If txtorderno.Text <> "" Then
                    ls_filter = ls_filter + " AND D.OrderNo LIKE '%" & txtorderno.Text & "%' " & vbCrLf
                End If

                'PALLETNO
                If txtPalletNo.Text <> "" Then
                    ls_filter = ls_filter + " AND DP.PalletNo = '" & txtPalletNo.Text & "' " & vbCrLf
                End If
                'BOXNO
                If txtBoxNo.Text <> "" Then
                    ls_filter = ls_filter + " AND D.BoxNo = '" & txtBoxNo.Text & "' " & vbCrLf
                End If


                ls_sql = " SELECT CONVERT(date,M.ShippingInstructionDate)colperiod,D.AffiliateID as colaffiliatecode, MA.AffiliateName as colaffiliatename, " & vbCrLf & _
                  " D.SupplierID as colsuppliercode, MS.SupplierName as colsuppliername, D.OrderNo as colorderno, D.ShippingInstructionNo as colinvoiveno,  " & vbCrLf & _
                  " DP.PalletNo as colpallet,DP.LabelNo as colLabelNo,D.PartNo as colpartno,D.QtyBox as colqty,DP.Location as collocation,DP.PalletType as colpallettype," & vbCrLf & _
                  " D.ForwarderID AS colHForwarder,D.SuratJalanNo AS colHsuratjalan " & vbCrLf & _
                  " FROM dbo.ShippingInstruction_DetailPallet DP " & vbCrLf & _
                  " INNER JOIN dbo.ShippingInstruction_Master M ON M.AffiliateID = DP.AffiliateID  " & vbCrLf & _
                  " AND M.ForwarderID = DP.ForwarderID AND M.ShippingInstructionNo = DP.ShippingInstructionNo " & vbCrLf & _
                  " INNER JOIN dbo.ShippingInstruction_Detail D ON D.AffiliateID = DP.AffiliateID  " & vbCrLf & _
                  " AND D.ForwarderID = M.ForwarderID AND D.ShippingInstructionNo = DP.ShippingInstructionNo " & vbCrLf & _
                  " INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = D.SupplierID " & vbCrLf & _
                  " INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = D.AffiliateID  " & vbCrLf & _
                  " Where ISNULL(DP.PalletNo,'') <> ''"

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
    Private Sub grid_CellEditorInitialize(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "colperiod" Or e.Column.FieldName = "colpartno" Or e.Column.FieldName = "colaffiliatecode" Or _
            e.Column.FieldName = "colHForwarder" Or e.Column.FieldName = "colinvoiveno" Or e.Column.FieldName = "colHsuratjalan" Or _
            e.Column.FieldName = "colsuppliercode" Or e.Column.FieldName = "colpartno" Or e.Column.FieldName = "colqty" Or _
            e.Column.FieldName = "colpallettype" Or e.Column.FieldName = "colaffiliatename" Or e.Column.FieldName = "colsuppliername" Or _
            e.Column.FieldName = "colorderno" Or e.Column.FieldName = "colLabelNo" Or e.Column.FieldName = "colpallet" Or e.Column.FieldName = "collocation") _
          And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub
    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)

            If pAction = "gridExcel" Or pAction = "Delete" Then GoTo keluar

            If pAction <> "send" Or pAction <> "gridExcel" Then
                Dim pAffiliate As String = Split(e.Parameters, "|")(1)
                Dim pSupplier As String = Split(e.Parameters, "|")(2)
                Dim pReceivingDate As String = Split(e.Parameters, "|")(3)
                Dim pInvoice As String = Split(e.Parameters, "|")(4)
                Dim pshipping As String = Split(e.Parameters, "|")(5)
                Dim pOrderNo As String = Split(e.Parameters, "|")(6)
                Dim pPaletNo As String = Split(e.Parameters, "|")(7)
                Dim pBoxNo As String = Split(e.Parameters, "|")(8)

            End If
keluar:
            Select Case pAction
                Case "gridload"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                Case "Delete"
                    Dim table As DataTable = Nothing
                    table = DirectCast(Session("table"), DataTable)
                    Dim selectItems As List(Of Object) = grid.GetSelectedFieldValues(New String() {"colperiod;colpartno;colaffiliatecode;colHForwarder;colinvoiveno;colHsuratjalan;colorderno;colLabelNo;colpallet"})

                    If selectItems.Count = 0 Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "6010", clsMessage.MsgType.InformationMessage)
                    Else
                        Dim ls_sql As String
                        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                            sqlConn.Open()
                            For Each selectItemId As Object In selectItems

                                Dim pAffiliate As String = Split(selectItemId, "|")(2)
                                Dim pForwarder As String = Split(selectItemId, "|")(3)
                                Dim pInvoice As String = Split(selectItemId, "|")(4)
                                Dim pSuratJalan As String = Split(selectItemId, "|")(5)
                                Dim pOrderNo As String = Split(selectItemId, "|")(6)
                                Dim pLabel As String = Split(selectItemId, "|")(7)
                                Dim pPalletNo As String = Split(selectItemId, "|")(8)

                                ls_sql = " DELETE dbo.ShippingInstruction_DetailPallet where " & vbCrLf & _
                                        "  AffiliateID = '" & pAffiliate & "' AND ForwarderID = '" & pForwarder & "' " & vbCrLf & _
                                        "  AND ShippingInstructionNo = '" & pInvoice & "' AND SuratJalanNo = '" & pSuratJalan & "' " & vbCrLf & _
                                        "  AND OrderNo = '" & pOrderNo & "' AND LabelNo = '" & pLabel & "'AND PalletNo = '" & pPalletNo & "' "
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
                Case "print"
                    'Call up_GridLoad()
                Case "gridExcel"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplatePalletizing.xlsx"
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
    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Palletizing" & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                '.Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 15).AutoFitColumns()

                .Cells(8, 1, pData.Rows.Count + 7, 1).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 11, pData.Rows.Count + 7, 11).Style.Numberformat.Format = "#,##0"

                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 15)
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
        txtInvoiceNo.Text = ""
        txtorderno.Text = ""
        txtPalletNo.Text = ""
        txtBoxNo.Text = ""
        checkboxdt.Checked = True
        rbshipping.SelectedIndex = 0
        Call up_fillcombo()
        up_GridLoadClear()
        grid.JSProperties("cpdtReceivingDate") = Format(Now, "dd MMM yyyy")
        grid.FocusedRowIndex = -1
        lblerrmessage.Text = ""
    End Sub
End Class