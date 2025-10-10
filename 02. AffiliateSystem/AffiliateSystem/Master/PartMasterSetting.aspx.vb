Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO

Public Class PartMasterSetting
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "PART MASTER SETTING"
            up_FillCombo()
            'If Session("M01Url") <> "" Then
            Call bindData()
            'Session.Remove("M01Url")
            'End If
            ScriptManager.RegisterStartupScript(grid, grid.GetType(), "init", "grid.SetFocusedRowIndex(-1);", True)
            lblInfo.Text = ""
            'txtMode.Text = "new"
        End If

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate


        Dim pIsUpdate As Boolean
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 0
        'Dim ls_ShowCls As Boolean
        Dim ls_ShowCls As String = ""
        Dim ls_PartNo As String = ""
        Dim ls_LocationID As String = ""
        Dim ls_AffiliateId As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("PartSetting")

                If grid.VisibleRowCount = 0 Then
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager, False, False)
                    Exit Sub
                End If

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
                    Session("YA010Msg") = lblInfo.Text
                    Exit Sub
                End If


                Dim a As Integer
                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1

                    ls_ShowCls = (e.UpdateValues(iLoop).NewValues("ShowCls").ToString())
                    If ls_ShowCls = True Then ls_ShowCls = "1" Else ls_ShowCls = "0"

                    ls_PartNo = (e.UpdateValues(iLoop).NewValues("PartNo").ToString())

                    If (e.UpdateValues(iLoop).NewValues("LocationID")) Is Nothing Then
                        ls_LocationID = ""
                    Else
                        ls_LocationID = (e.UpdateValues(iLoop).NewValues("LocationID").ToString())
                    End If

                    Dim sqlstring As String
                    sqlstring = " SELECT PartNo FROM MS_PartSetting WHERE PartNo='" & ls_PartNo & "' and AffiliateID = '" & Session("AffiliateID").ToString & "'"

                    Dim sqlComm As New SqlCommand(sqlstring, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If pIsUpdate = False Then
                        'INSERT DATA
                        ls_SQL = " INSERT INTO MS_PartSetting " & _
                            "(AffiliateID,PartNo,ShowCls)" & _
                            " VALUES ( '" & Session("AffiliateID").ToString & "','" & ls_PartNo & "','" & ls_ShowCls & "')" & vbCrLf
                        ls_MsgID = "1001"
                    Else
                        ls_SQL = " 	UPDATE MS_PartSetting " & vbCrLf & _
                                 " 	   SET ShowCls = '" & ls_ShowCls & "' " & vbCrLf & _
                                 " 	 WHERE PartNo ='" & ls_PartNo & "' and AffiliateID = '" & Session("AffiliateID").ToString & "'"
                        ls_MsgID = "1002"
                    End If

                    ls_SQL = ls_SQL & "UPDATE MS_PartMapping SET LocationID = '" & ls_LocationID & "' " & vbCrLf & _
                        "WHERE AffiliateID = '" & Session("AffiliateID").ToString & "' " & vbCrLf & _
                        "AND PartNo = '" & ls_PartNo & "' "

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

NextLoop:
                Next iLoop

                sqlTran.Commit()

            End Using

            sqlConn.Close()
        End Using

        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        Session("YA010Msg") = lblInfo.Text
        grid.JSProperties("cpType") = "info"
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName") _
            And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnUpload_Click(sender As Object, e As System.EventArgs) Handles btnUpload.Click
        Session.Remove("M01Url")
        Response.Redirect("~/Master/PartMasterUpload.aspx")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Session("YA010Msg") = ""
        Session("YA010Msg") = lblInfo.Text

        bindData()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call bindData()

                    grid.JSProperties("cpMessage") = Session("YA010Msg")
                    Session("YA010Msg") = lblInfo.Text

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                        grid.JSProperties("cpMessage") = Session("YA010Msg")
                    End If
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetTableAffiliate()
                    FileName = "TemplateMSPart.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:3", psERR)
                    End If
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "upload"
                    Response.Redirect("~/Master/PartMasterUpload.aspx", False)
EndProcedure:
                    Session("AA220Msg") = ""
            End Select




        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Function GetTableAffiliate(Optional ByRef pErr As String = "") As DataTable
        Dim ls_sql As String = ""
        Dim pWhere As String = ""
        pErr = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""
                ls_sql = " select row_number() over (order by PartNo) NoUrut, * " & vbCrLf & _
                          " from ( " & vbCrLf & _
                          " 	select  " & vbCrLf & _
                          " 		distinct " & vbCrLf & _
                          " 		RTRIM(a.PartNo) PartNo, " & vbCrLf & _
                          " 		RTRIM(c.PartName) PartName, ISNULL(a.MOQ,0) MOQ, " & vbCrLf & _
                          " 		CASE WHEN ISNULL(b.ShowCls,0) = '1' THEN 'YES' ELSE 'NO' END ShowCls, ISNULL(a.LocationID, '') LocationID " & vbCrLf & _
                          " 	from  MS_PartMapping a " & vbCrLf & _
                          " 	left join MS_PartSetting b on a.AffiliateID = b.AffiliateID and a.PartNo = b.PartNo " & vbCrLf & _
                          " 	left join MS_Parts c on a.PartNo = c.PartNo " & vbCrLf & _
                          " 	where a.AffiliateID = '" & Session("AffiliateID") & "'" & pWhere & " " & vbCrLf & _
                          " )x "

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                Return dt
            End Using
        Catch ex As Exception
            pErr = ex.Message
            Return Nothing
        End Try
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

                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 6)
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

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim id As String = Session("AffiliateID").ToString


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If cboPartNo.Text <> "" Then
                If cboPartNo.Text <> clsGlobal.gs_All Then
                    pWhere = " and a.PartNo = '" & cboPartNo.Text & "' "
                End If
            End If

            ls_SQL = " select row_number() over (order by PartNo) NoUrut, * " & vbCrLf & _
                  " from ( " & vbCrLf & _
                  " 	select  " & vbCrLf & _
                  " 		distinct " & vbCrLf & _
                  " 		RTRIM(a.PartNo) PartNo, " & vbCrLf & _
                  " 		RTRIM(c.PartName) PartName, " & vbCrLf & _
                  " 		ISNULL(b.ShowCls,0) ShowCls, ISNULL(a.MOQ,0) MOQ, ISNULL(a.LocationID, '') LocationID " & vbCrLf & _
                  " 	from  MS_PartMapping a " & vbCrLf & _
                  " 	left join MS_PartSetting b on a.AffiliateID = b.AffiliateID and a.PartNo = b.PartNo " & vbCrLf & _
                  " 	left join MS_Parts c on a.PartNo = c.PartNo " & vbCrLf & _
                  " 	where a.AffiliateID = '" & Session("AffiliateID") & "'" & pWhere & " " & vbCrLf & _
                  " )x "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = " select top 0 '' NoUrut, '' PartNo, '' PartName, '' ShowCls, '' MOQ, '' LocationID"

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

        'Person In Charge

        ls_SQL = " SELECT '== ALL ==' PartCode, '== ALL ==' PartName union all select RTRIM(a.PartNo) PartCode, b.PartName from MS_PartMapping a  left join MS_Parts b on a.PartNo = b.PartNo  where a.AffiliateID = '" & Session("AffiliateID").ToString & "' " & vbCrLf & _
                  "  order by PartCode " & vbCrLf

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
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 400

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = 0
                txtPartNo.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using


    End Sub
#End Region

End Class