Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions

Public Class ForecastEntry
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim pMsgID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            dtPeriodFrom.Value = Now
            up_FillCombo()
            lblInfo.Text = ""
            up_GridLoadWhenEventChange()
            up_IsiHeader()
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim sql As String = ""
        Dim a As Integer
        Dim iWeek As Integer = 1
        Dim pInsert As String
        Dim ls_PartNo As String
        Dim ls_Qty As Double

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("InsertDetailData")
                a = e.UpdateValues.Count

                For iLoop = 0 To a - 1
                    For iWeek = 1 To 12
                        Dim pFullDate As String = "01-" & grid.Columns("Bln" & iWeek).ToString
                        pFullDate = "01-" & Format(DateAdd(DateInterval.Month, iWeek - 1, dtPeriodFrom.Value), "MMM-yy")

                        ls_PartNo = e.UpdateValues(iLoop).NewValues("PartNo")
                        ls_Qty = e.UpdateValues(iLoop).NewValues("Bln" & iWeek)

                        pInsert = getInsertOrUpdate(ls_PartNo, pFullDate)

                        If pInsert = True Then
                            sql = " UPDATE [dbo].[MS_Forecast] " & vbCrLf & _
                                    "    SET [Qty] = " & ls_Qty & " " & vbCrLf & _
                                    "       ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                    "       ,[UpdateDate] = GETDATE() " & vbCrLf & _
                                    "  WHERE [PartNo] = '" & ls_PartNo & "' and [Period] = '" & pFullDate & "' and [AffiliateID] = '" & Session("AffiliateID") & "'"

                            pMsgID = "1002"

                        Else
                            sql = " INSERT INTO [MS_Forecast] " & vbCrLf & _
                                    "            ([AffiliateID] " & vbCrLf & _
                                    "            ,[PartNo] " & vbCrLf & _
                                    "            ,[Period] " & vbCrLf & _
                                    "            ,[Qty] " & vbCrLf & _
                                    "            ,[EntryDate] " & vbCrLf & _
                                    "            ,[EntryUser]) " & vbCrLf & _
                                    "      VALUES " & vbCrLf & _
                                    "            ('" & Session("AffiliateID") & "','" & ls_PartNo & "' " & vbCrLf & _
                                    "            ,'" & pFullDate & "' " & vbCrLf & _
                                    "            ,'" & ls_Qty & "' " & vbCrLf & _
                                    "            ,getdate() " & vbCrLf & _
                                    "            ,'" & Session("UserID") & "') "

                            pMsgID = "1001"
                        End If

                        Dim SqlComm As New SqlCommand(sql, sqlConn, sqlTran)
                        SqlComm.ExecuteNonQuery()
                        SqlComm.Dispose()
                    Next
                Next

                sqlTran.Commit()
            End Using

            sqlConn.Close()

        End Using

        Call clsMsg.DisplayMessage(lblInfo, pMsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        Session("AA220Msg") = lblInfo.Text
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        'e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        If x > grid.VisibleRowCount Then Exit Sub

        With grid
            If .VisibleRowCount > 0 Then
                If e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Then
                    e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
                End If

                If e.DataColumn.FieldName = "Bln1" Or e.DataColumn.FieldName = "Bln2" Or e.DataColumn.FieldName = "Bln3" Or _
                    e.DataColumn.FieldName = "Bln4" Or e.DataColumn.FieldName = "Bln5" Or e.DataColumn.FieldName = "Bln6" Or _
                    e.DataColumn.FieldName = "Bln7" Or e.DataColumn.FieldName = "Bln8" Or e.DataColumn.FieldName = "Bln9" Or _
                    e.DataColumn.FieldName = "Bln10" Or e.DataColumn.FieldName = "Bln11" Or e.DataColumn.FieldName = "Bln12" Then
                    e.Cell.BackColor = Color.White
                End If
            End If
        End With
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        grid.JSProperties("cpMessage") = ""
        up_GridLoadWhenEventChange()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
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
                    'Call up_IsiHeader()
                Case "loadSave"
                    grid.PageIndex = 0
                    bindData()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        Finally
            Session("AA220Msg") = ""
        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim pFullDate As String = ""
        Dim pYear As String = ""
        Dim pMonth As String = ""
        Dim pDay As String = "01"

        pYear = Year(dtPeriodFrom.Value)
        pMonth = Month(dtPeriodFrom.Value)

        pFullDate = pYear & "-" & pMonth & "-" & pDay

        If cboPartNo.Text <> clsGlobal.gs_All Then
            pWhere = " where a.PartNo = '" & cboPartNo.Text & "'"
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                    " 	row_number() over (order by a.PartNo) as NoUrut,a.PartNo, a.PartName, " & vbCrLf & _
                    " 	SUM(case when Period = '" & pFullDate & "' then Qty else 0 end) Bln1, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,1,'" & pFullDate & "') then Qty else 0 end) Bln2, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,2,'" & pFullDate & "') then Qty else 0 end) Bln3, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,3,'" & pFullDate & "') then Qty else 0 end) Bln4, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,4,'" & pFullDate & "') then Qty else 0 end) Bln5, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,5,'" & pFullDate & "') then Qty else 0 end) Bln6, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,6,'" & pFullDate & "') then Qty else 0 end) Bln7, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,7,'" & pFullDate & "') then Qty else 0 end) Bln8, " & vbCrLf & _
                    " 	SUM(case when Period = dateadd(Month,8,'" & pFullDate & "') then Qty else 0 end) Bln9, "

            ls_SQL = ls_SQL + " 	SUM(case when Period = dateadd(Month,9,'" & pFullDate & "') then Qty else 0 end) Bln10, " & vbCrLf & _
                              " 	SUM(case when Period = dateadd(Month,10,'" & pFullDate & "') then Qty else 0 end) Bln11, " & vbCrLf & _
                              " 	SUM(case when Period = dateadd(Month,11,'" & pFullDate & "') then Qty else 0 end) Bln12 " & vbCrLf & _
                              " from MS_Parts a " & vbCrLf & _
                              " inner join MS_PartMapping c on a.PartNo = c.PartNo and c.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " inner join MS_PartSetting d on a.PartNo = d.PartNo and d.AffiliateID = '" & Session("AffiliateID") & "' and ShowCls = '1'" & vbCrLf & _
                              " left join MS_Forecast b on a.PartNo = b.PartNo " & vbCrLf & _
                              " " & pWhere & "" & vbCrLf & _
                              " group by a.PartNo, a.PartName "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' PartNo, '' PartName, ''WK1, ''WK2, ''WK3, ''WK4, ''WK5, ''WK6, ''WK7, ''WK8, ''WK9, ''WK10, ''WK11, ''WK12"

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

        For iWeek = 1 To 12
            'grid.Columns("WK" & iWeek).Caption = Format(DateAdd(DateInterval.Month, iWeek - 1, dtPeriodFrom.Value), "MMM-yy") 'Format(FirstDay, "dd") & " - " & Format(WeekEnd, "dd MMM")
            grid.Columns("Bln" & iWeek).Caption = Format(DateAdd(DateInterval.Month, iWeek - 1, dtPeriodFrom.Value), "MMM-yy")
        Next
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all " & vbCrLf & _
                 "select distinct RTRIM(a.PartNo) PartCode, PartName from MS_Parts a" & vbCrLf & _
                 "inner join MS_PartMapping b on a.PartNo = b.PartNo and b.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                 "inner join MS_PartSetting c on a.PartNo = c.PartNo and b.AffiliateID = '" & Session("AffiliateID") & "' and ShowCls = '1'" & vbCrLf & _
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
    End Sub

    Private Function getInsertOrUpdate(ByVal pPartNo As String, ByVal pPeriod As String) As Boolean
        Dim sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            sql = " select * " & vbCrLf & _
                    " from MS_Forecast " & vbCrLf & _
                    " where PartNo='" & pPartNo.Trim & "' " & vbCrLf & _
                    " and Period='" & pPeriod.Trim & "' and  AffiliateID = '" & Session("AffiliateID") & "'"

            Dim sqlDA As New SqlDataAdapter(sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If

            sqlConn.Close()

        End Using
    End Function
#End Region
End Class