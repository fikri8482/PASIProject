Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports OfficeOpenXml
Imports System.IO
Imports System.Drawing

Public Class PartMaster
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), "A05")
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), "A05")

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "PART MASTER"
            Call bindData()
            DeleteHistory()
            ScriptManager.RegisterStartupScript(grid, grid.GetType(), "init", "grid.SetFocusedRowIndex(-1);", True)

            lblInfo.Text = ""
        End If

        If ls_AllowDownload = False Then btnDownload.Enabled = False
        If ls_AllowUpdate = False Then btnUpload.Enabled = False
        If ls_AllowUpdate = False Then btnADD.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnADD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnADD.Click
        Session("M01Url") = "~/Master/PartMaster.aspx"
        Response.Redirect("~/Master/PartMasterDetail.aspx")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadPart.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowPager)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim pSuppType As String = ""
                    Dim dtProd As DataTable = clsMaster.GetTablePart(txtPartCode.Text, txtPartName.Text, txtMaker.Text, txtProject.Text)
                    FileName = "TemplateMSPart.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:3", psERR)
                    End If
            End Select

EndProcedure:
            Session("AA220Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        If grid.VisibleRowCount > 0 Then
            If e.GetValue("DeleteCls") = "1" Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If txtPartCode.Text.Trim <> "" Then
            pWhere = pWhere + " and PartNo like '%" & txtPartCode.Text.Trim & "%' "
        End If

        If txtPartName.Text.Trim <> "" Then
            pWhere = pWhere + "and PartName like '%" & txtPartName.Text.Trim & "%' "
        End If

        If txtMaker.Text.Trim <> "" Then
            pWhere = pWhere + "and Maker like '%" & txtMaker.Text.Trim & "%' "
        End If

        If txtProject.Text.Trim <> "" Then
            pWhere = pWhere + "and Project like '%" & txtProject.Text.Trim & "%' "
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select row_number() over (order by PartNo) as NoUrut, * from " & vbCrLf & _
                      "    (select " & vbCrLf & _
                      " 	RTRIM(PartNo)PartNo, " & vbCrLf & _
                      " 	RTRIM(PartName)PartName, " & vbCrLf & _
                      " 	ISNULL(RTRIM(PartCarMaker),'')PartCarMaker, " & vbCrLf & _
                      " 	ISNULL(RTRIM(PartCarName),'')PartCarName, " & vbCrLf & _
                      " 	ISNULL(RTRIM(PartGroupName),'')PartGroupName, " & vbCrLf & _
                      " 	ISNULL(RTRIM(HSCode),'')HSCode, " & vbCrLf & _
                      " 	case when FinishGoodCls = '1' then 'FG' else 'PART' end  FinishGoodCls, " & vbCrLf & _
                      "     b.Description UOM, " & vbCrLf & _
                      " 	case when KanbanCls = '1' then 'YES' else 'NO' end KanbanCls, " & vbCrLf & _
                      " 	Maker, " & vbCrLf & _
                      " 	Project, " & vbCrLf & _
                      " 	EntryDate, " & vbCrLf & _
                      " 	EntryUser, " & vbCrLf & _
                      " 	UpdateDate, " & vbCrLf & _
                      " 	UpdateUser, 0 DeleteCls, " & vbCrLf & _
                      " 'DETAIL' DetailPage " & vbCrLf & _
                      " from MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls " & vbCrLf

            ls_SQL = ls_SQL + "UNION ALL" & vbCrLf & _
                      " select " & vbCrLf & _                      
                      " 	RTRIM(PartNo)PartNo, " & vbCrLf & _
                      " 	RTRIM(PartName)PartName, " & vbCrLf & _
                      " 	ISNULL(RTRIM(PartCarMaker),'')PartCarMaker, " & vbCrLf & _
                      " 	ISNULL(RTRIM(PartCarName),'')PartCarName, " & vbCrLf & _
                      " 	ISNULL(RTRIM(PartGroupName),'')PartGroupName, " & vbCrLf & _
                      " 	ISNULL(RTRIM(HSCode),'')HSCode, " & vbCrLf & _
                      " 	case when FinishGoodCls = '1' then 'FG' else 'PART' end  FinishGoodCls, " & vbCrLf & _
                      "     b.Description UOM, " & vbCrLf & _
                      " 	case when KanbanCls = '1' then 'YES' else 'NO' end KanbanCls, " & vbCrLf & _
                      " 	Maker, " & vbCrLf & _
                      " 	Project, " & vbCrLf & _
                      " 	EntryDate, " & vbCrLf & _
                      " 	EntryUser, " & vbCrLf & _
                      " 	UpdateDate, " & vbCrLf & _
                      " 	UpdateUser, 1 DeleteCls, " & vbCrLf & _
                      " 'DETAIL' DetailPage " & vbCrLf & _
                      " from MS_Parts_History a left join MS_UnitCls b on a.UnitCls = b.UnitCls " & vbCrLf

            ls_SQL = ls_SQL + ")xyz where 'A' = 'A' " & pWhere & " Order by DeleteCls, PartNo"

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

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' as PartNo, '' PartName, '' PartCarMaker, '' PartCarName, '' PartGroupName, '' HSCode, '' FinishGoodCls, '' UOM, '' KanbanCls, '' Maker, '' Project, '' DetailPage"

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

    Private Sub DeleteHistory()
        Dim ls_sql As String

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                    Dim sqlComm As New SqlCommand()

                    ls_sql = " delete MS_Parts_History " & vbCrLf & _
                              " where exists " & vbCrLf & _
                              " (select * from MS_Parts a where MS_Parts_History.PartNo = a.PartNo) "

                    sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    sqlComm.Dispose()
                    sqlTran.Commit()
                End Using

                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Protected Function GetRowValue(ByVal container As GridViewDataItemTemplateContainer) As String
        If Not IsNothing(container.KeyValue) Then
            Return container.Grid.GetRowValuesByKeyValue(container.KeyValue, "PartNo").ToString()
        End If
    End Function

    Protected Function GetAffiliateID() As String
        GetAffiliateID = txtPartCode.Text.Trim
    End Function

    Protected Function GetAffiliateName() As String
        GetAffiliateName = txtPartName.Text.Trim
    End Function

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Part Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Template\Result\" & tempFile & "")

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
                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count - 0
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                    Next
                Next

                ''ALIGNMENT
                ''.Cells(rowstart + 1, icol, irow, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(iRow + space, colKanbanSeqNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPartName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvQty).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                ''.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvCurr).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

                'Dim rgAll As ExcelRange = .Cells('.Cells(Space() - 2, colNo, grid.VisibleRowCount + (Space() - 1), colCount - 1)
                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 10)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Template\Result\" & tempFile & "")

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