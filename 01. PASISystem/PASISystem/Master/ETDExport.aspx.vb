Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO

Public Class ETDExport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim menuID As String = "A25"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        up_FillCombo2()
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), menuID)
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        'up_FillCombo()

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "TIME CHART MASTER"
            up_FillCombo()
            DeleteHistory()
            DtPeriod.Focus()
            DtPeriod.Value = Now
            Format(DtPeriod.Value.Now, ("MMM yyyy"))
            dtETDVendor.Value = Now
            dtETAForwarder.Value = Now
            dtETDPort.Value = Now
            dtETAPort.Value = Now
            dtETAFactory.Value = Now
            dtcutoff.Value = Now
            If Session("M01Url") <> "" Then
                Call up_GridLoad()
                Session.Remove("M01Url")
            End If

            lblInfo.Text = ""

            grid.FocusedRowIndex = -1
            'ElseIf IsCallback = False Then
            '    cboAffiliate.Text = Session("A25AffiliateID")
            '    txtAffiliate.Text = Session("A25AffiliateName")
            '    cboSupplier.Text = Session("A25SupplierID")
            '    txtSupplier.Text = Session("A25SupplierName")
        End If

        'Session.Remove("A25AffiliateID")
        'Session.Remove("A25AffiliateName")
        'Session.Remove("A25SupplierID")
        'Session.Remove("A25SupplierName")

        If ls_AllowDownload = False Then btnDownload.Enabled = False
        If ls_AllowUpdate = False Then btnUpload.Enabled = False
        If ls_AllowUpdate = False Then btnSubmit.Enabled = False
        If ls_AllowDelete = False Then btnDelete.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
        ScriptManager.RegisterStartupScript(grid, grid.GetType(), "scriptKey", "txtForwarderCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;'); grid.SetFocusedRowIndex(-1); grid.SetFocusedRowIndex(-1);", True)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared        
        If grid.VisibleRowCount > 0 Then
            If e.GetValue("DeleteCls") = "1" Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        grid.JSProperties("cpMessage") = ""
        Try
            Select Case pAction
                Case "save"
                    'If Format(DtPeriod.Value, "MMM yyyy") <> Format(dtcutoff.Value, "MMM yyyy") Then
                    '    Call clsMsg.DisplayMessage(lblInfo, "6030", clsMessage.MsgType.ErrorMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    '    Exit Sub
                    'End If

                    If Split(e.Parameters, "|")(5) = "null" Then
                        Call up_SaveDataDate(Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(4))
                        Exit Sub
                    End If

                    If validasiETD(Split(e.Parameters, "|")(6), _
                                   Split(e.Parameters, "|")(10), _
                                   Split(e.Parameters, "|")(7), _
                                   Split(e.Parameters, "|")(8), _
                                   Split(e.Parameters, "|")(9)) = False Then Exit Select


                    Dim lb_IsUpdate As Boolean = validasiInput(Split(e.Parameters, "|")(5))
                    Call up_SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(4), _
                                     Split(e.Parameters, "|")(5), _
                                     Split(e.Parameters, "|")(6), _
                                     Split(e.Parameters, "|")(7), _
                                     Split(e.Parameters, "|")(8), _
                                     Split(e.Parameters, "|")(9))

                    Call up_GridLoad()
                Case "delete"

                    'Call up_DeleteData(Format(DtPeriod.Value, "yyyy-MM-01"), Trim(cboAffiliate.Text), Trim(cboSupplier.Text), Split(e.Parameters, "|")(1))

                    If HF.Get("DeleteCls") = "0" Then
                        up_DeleteData(Format(DtPeriod.Value, "yyyy-MM-01"), Trim(cboAffiliate.Text), Trim(cboSupplier.Text), Split(e.Parameters, "|")(1))
                    Else
                        up_DeleteDataRec(Format(DtPeriod.Value, "yyyy-MM-01"), Trim(cboAffiliate.Text), Trim(cboSupplier.Text), Split(e.Parameters, "|")(1))
                    End If

                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Call up_GridLoad()

                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    ElseIf grid.VisibleRowCount > 1 Then
                        grid.JSProperties("cpMessage") = ""
                        lblInfo.Text = ""
                    End If

                    grid.FocusedRowIndex = -1


                Case "loadeventchange"
                    Call up_GridLoadWhenEventChange()

                Case "loadaftersubmit"
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Call up_GridLoad()
                    grid.FocusedRowIndex = -1

                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = clsMaster.GetTableETDExport(Format(DtPeriod.Value, "yyyy-MM-01"), Trim(cboAffiliate.Text), Trim(cboSupplier.Text))
                    FileName = "TemplateMSTimeChart.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:3", psERR)
                    Else
                        Call clsMsg.DisplayMessage(lblInfo, "2007", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            grid.JSProperties("cpError") = lblInfo.Text
            grid.FocusedRowIndex = -1
        End Try
    End Sub

#End Region

#Region "PROCEDURE"
    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_where As String = ""
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If cboAffiliate.Text <> "" Then
                ls_where = ls_where + " AND AffiliateID = '" & Trim(cboAffiliate.Text) & "'"
            End If

            If cboSupplier.Text <> "" Then
                ls_where = ls_where + " AND SupplierID = '" & Trim(cboSupplier.Text) & "'"
            End If

            ls_SQL = " SELECT ROW_NUMBER() OVER (ORDER BY SupplierID) AS RowNumber, * FROM (SELECT Period," & vbCrLf & _
                     " MSE.Week, MSE.SupplierID, SupplierName, MSA.AffiliateID, AffiliateName," & vbCrLf & _
                     " CONVERT(DATETIME,ETDVendor,116)ETDVendor, ETAForwarder, ETDPort, ETAPort, ETAFactory, CutOfDate, 0 DeleteCls, MSE.EntryDate, MSE.EntryUser, MSE.UpdateDate, MSE.UpdateUser" & vbCrLf & _
                     " FROM MS_ETD_Export MSE" & vbCrLf & _
                     " LEFT JOIN MS_AFFILIATE MSA" & vbCrLf & _
                     " ON MSE.AffiliateID = MSA.AffiliateID" & vbCrLf & _
                     " LEFT JOIN MS_Supplier MSS" & vbCrLf & _
                     " ON MSE.SupplierID = MSS.SupplierID" & vbCrLf 
            ls_SQL = ls_SQL + " UNION ALL SELECT Period," & vbCrLf & _
                     " MSE.Week, MSE.SupplierID, SupplierName, MSA.AffiliateID, AffiliateName," & vbCrLf & _
                     " CONVERT(DATETIME,ETDVendor,116)ETDVendor, ETAForwarder, ETDPort, ETAPort, ETAFactory, CutOfDate, 1 DeleteCls, MSE.EntryDate, MSE.EntryUser, MSE.UpdateDate, MSE.UpdateUser" & vbCrLf & _
                     " FROM MS_ETD_Export_History MSE" & vbCrLf & _
                     " LEFT JOIN MS_AFFILIATE MSA" & vbCrLf & _
                     " ON MSE.AffiliateID = MSA.AffiliateID" & vbCrLf & _
                     " LEFT JOIN MS_Supplier MSS" & vbCrLf & _
                     " ON MSE.SupplierID = MSS.SupplierID" & vbCrLf & _
                     " ) XYZ WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "'" & vbCrLf & _
                     " " & ls_where & "" & vbCrLf & _
                     " ORDER BY Period, SupplierID, AffiliateID"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            'Session("A25AffiliateID") = cboAffiliate.Text.Trim
            'Session("A25AffiliateName") = txtAffiliate.Text.Trim
            'Session("A25SupplierID") = cboSupplier.Text.Trim
            'Session("A25SupplierName") = txtSupplier.Text.Trim

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                '.VisibleColumns(0).Width = 35 'RowNumber
                '.VisibleColumns(1).Width = 110 'AffiliateID
                '.VisibleColumns(2).Width = 110 'AffiliateID
                '.VisibleColumns(3).Width = 110 'AffiliateID
                '.VisibleColumns(3).Width = 110 'ETDVendor
                '.VisibleColumns(4).Width = 110 'ETDPort
                '.VisibleColumns(5).Width = 110 'ETAPort
                '.VisibleColumns(6).Width = 110 'ETAFactory
            End With

            If ds.Tables(0).Rows.Count > 0 Then
                'grid.JSProperties("cpDate") = ds.Tables(0).Rows(0)("CutOfDate").ToString.Trim
                dtcutoff.Value = ds.Tables(0).Rows(0)("CutOfDate")
            Else
                dtcutoff.Value = Now
                'grid.JSProperties("cpDate") = ""
            End If


            sqlConn.Close()
        End Using
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT TOP 0 ROW_NUMBER() OVER (ORDER BY MSE.SupplierID) AS RowNumber, Period," & vbCrLf & _
                     " MSE.Week, MSE.SupplierID, SupplierName, MSA.AffiliateID, AffiliateName," & vbCrLf & _
                     " CONVERT(DATETIME,ETDVendor,116)ETDVendor, ETDPort, ETAPort, ETAFactory" & vbCrLf & _
                     " FROM MS_ETD_Export MSE" & vbCrLf & _
                     " LEFT JOIN MS_AFFILIATE MSA" & vbCrLf & _
                     " ON MSE.AffiliateID = MSA.AffiliateID" & vbCrLf & _
                     " LEFT JOIN MS_Supplier MSS" & vbCrLf & _
                     " ON MSE.SupplierID = MSS.SupplierID" & vbCrLf & _
                     " WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "'" & vbCrLf & _
                     " AND MSE.AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND MSE.SupplierID = '" & Trim(cboSupplier.Text) & "' " & vbCrLf & _
                     " ORDER BY Period, MSE.SupplierID, MSE.AffiliateID"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                '.VisibleColumns(0).Width = 35 'RowNumber
                '.VisibleColumns(1).Width = 110 'AffiliateID
                '.VisibleColumns(2).Width = 110 'AffiliateID
                '.VisibleColumns(3).Width = 110 'ETDVendor
                '.VisibleColumns(4).Width = 110 'ETDPort
                '.VisibleColumns(5).Width = 110 'ETAPort
                '.VisibleColumns(6).Width = 110 'ETAFactory
            End With

            sqlConn.Close()
        End Using
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub TabIndex()
        DtPeriod.TabIndex = 1
        cboSupplier.TabIndex = 2
        cboAffiliate.TabIndex = 3
        cboAffiliate.TabIndex = 4
        cboSupplier.TabIndex = 5
        dtETDVendor.TabIndex = 6
        dtETDPort.TabIndex = 7
        dtETAPort.TabIndex = 8
        dtETAFactory.TabIndex = 9
        btnSubmit.TabIndex = 10
        btnDelete.TabIndex = 11
        btnClear.TabIndex = 12
        btnSubMenu.TabIndex = 13
    End Sub

    Private Sub up_SaveData(ByVal pIsUpdate As Boolean, _
                            Optional ByVal pPeriod As String = "", _
                            Optional ByVal pAffiliateID As String = "", _
                            Optional ByVal pSupplierID As String = "", _
                            Optional ByVal pWeek As String = "", _
                            Optional ByVal pETDVendor As String = "", _
                            Optional ByVal pETDPort As String = "", _
                            Optional ByVal pETAPort As String = "", _
                            Optional ByVal pETAFactory As String = "", _
                            Optional ByVal pETAForwarder As String = "")
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim shostname As String = System.Net.Dns.GetHostName
        Dim ls_week As String = ""
        Dim ls_ETDVendor As String = ""
        Dim ls_ETAFWD As String = ""
        Dim ls_ETDPORT As String = ""
        Dim ls_ETAPort As String = ""
        Dim ls_ETAFactory As String = ""
        Dim ls_CutOffDatea As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT * FROM dbo.MS_ETD_Export WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND SupplierID = '" & Trim(cboSupplier.Text) & "' AND WEEK = '" & pWeek & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                pIsUpdate = False

                ls_week = ds.Tables(0).Rows(0)("Week")
                ls_ETDVendor = Format(ds.Tables(0).Rows(0)("ETDVendor"), "yyyy-MM-dd")
                ls_ETAFWD = Format(ds.Tables(0).Rows(0)("ETAForwarder"), "yyyy-MM-dd")
                ls_ETDPORT = Format(ds.Tables(0).Rows(0)("ETDPort"), "yyyy-MM-dd")
                ls_ETAPort = Format(ds.Tables(0).Rows(0)("ETAPort"), "yyyy-MM-dd")
                ls_ETAFactory = Format(ds.Tables(0).Rows(0)("ETAFactory"), "yyyy-MM-dd")
                ls_CutOffDatea = Format(ds.Tables(0).Rows(0)("CutOfDate"), "yyyy-MM-dd")
            Else
                pIsUpdate = True
            End If
            sqlConn.Close()
        End Using

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ForwarderID")
                Dim sqlComm As New SqlCommand()

                If pIsUpdate = True Then
                    'INSERT DATA
                    ls_SQL = " INSERT INTO dbo.MS_ETD_Export(Period, AffiliateID, SupplierID, Week, ETDVendor,ETAForwarder,ETDPort,ETAPort,ETAFactory,CutOfDate)" & vbCrLf & _
                             " VALUES ('" & Format(DtPeriod.Value, "MMM yyyy") & "'," & _
                             " '" & Trim(cboAffiliate.Text) & "'," & _
                             " '" & Trim(cboSupplier.Text) & "'," & _
                             " '" & pWeek & "'," & _
                             " '" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "'," & _
                             " '" & Convert.ToDateTime(dtETAForwarder.Value).ToString("yyyy-MM-dd") & "'," & _
                             " '" & Convert.ToDateTime(dtETDPort.Value).ToString("yyyy-MM-dd") & "'," & _
                             " '" & Convert.ToDateTime(dtETAPort.Value).ToString("yyyy-MM-dd") & "'," & _
                             " '" & Convert.ToDateTime(dtETAFactory.Value).ToString("yyyy-MM-dd") & "'," & _
                             " '" & Convert.ToDateTime(dtcutoff.Value).ToString("yyyy-MM-dd") & "')"
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    'ElseIf pIsUpdate = False And flag = True Then
                    '    ls_MsgID = "6018"
                    '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    '    lblInfo.Visible = True
                    '    Exit Sub

                ElseIf pIsUpdate = False Then
                    ls_SQL = " UPDATE dbo.MS_ETD_Export " & vbCrLf & _
                             " SET ETDVendor = '" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                             " ETAForwarder = '" & Convert.ToDateTime(dtETAForwarder.Value).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                             " ETDPort = '" & Convert.ToDateTime(dtETDPort.Value).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                             " ETAPort = '" & Convert.ToDateTime(dtETAPort.Value).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                             " ETAFactory = '" & Convert.ToDateTime(dtETAFactory.Value).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                             " CutOfDate = '" & Convert.ToDateTime(dtcutoff.Value).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                             " WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & Trim(pWeek) & "'"
                    ls_MsgID = "1002"
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    Dim ls_Remarks As String = ""

                    If CDate(dtETDVendor.Value) <> CDate(ls_ETDVendor) Then
                        ls_Remarks = ls_Remarks + "Effective Date " + ls_ETDVendor & "->" & dtETDVendor.Value & " "
                    End If
                    If CDate(dtETAForwarder.Value) <> CDate(ls_ETAFWD) Then
                        ls_Remarks = ls_Remarks + "End Date " + ls_ETAFWD & "->" & dtETAForwarder.Value & " "
                    End If
                    If CDate(dtETDPort.Value) <> CDate(ls_ETDPORT) Then
                        ls_Remarks = ls_Remarks + "End Date " + ls_ETDPORT & "->" & dtETDPort.Value & " "
                    End If
                    If CDate(dtETAPort.Value) <> CDate(ls_ETAPort) Then
                        ls_Remarks = ls_Remarks + "End Date " + ls_ETAPort & "->" & dtETAPort.Value & " "
                    End If
                    If CDate(dtETAFactory.Value) <> CDate(ls_ETAFactory) Then
                        ls_Remarks = ls_Remarks + "End Date " + ls_ETAFactory & "->" & dtETAFactory.Value & " "
                    End If
                    If CDate(dtcutoff.Value) <> CDate(ls_CutOffDatea) Then
                        ls_Remarks = ls_Remarks + "End Date " + ls_CutOffDatea & "->" & dtcutoff.Value & " "
                    End If
                    

                    Dim ls_Remarks2 As String = "Week " & pWeek & " Affiliate " & pAffiliateID & " Supplier " & pSupplierID & ", "

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U','','Update [" & ls_Remarks2 & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                    End If
                End If

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        lblInfo.Visible = True
    End Sub

    Private Sub up_SaveDataDate(Optional ByVal pPeriod As String = "", _
                            Optional ByVal pAffiliateID As String = "", _
                            Optional ByVal pSupplierID As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ForwarderID")

                ls_SQL = " UPDATE dbo.MS_ETD_Export " & vbCrLf & _
                         " SET CutOfDate = '" & Convert.ToDateTime(dtcutoff.Value).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         " WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' "
                ls_MsgID = "1002"

                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        lblInfo.Visible = True
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
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Time Chart Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                    For icol = 1 To pData.Columns.Count
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                        If icol = 1 Then
                            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "mmm-yy"
                        End If
                        If icol > 4 Then
                            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "dd-mmm-yy"
                        End If
                    Next
                Next

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

    Private Sub up_DeleteData(ByVal pPeriod As String, ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pWeek As String)
        Dim ls_SQL As String = ""
        Dim x As Integer
        Dim shostname As String = System.Net.Dns.GetHostName

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ForwarderID")

                'ls_SQL = "DELETE from dbo.MS_ETD_Export WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & pWeek & "'"
                'Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                'sqlComm.ExecuteNonQuery()

                ls_SQL = " INSERT INTO MS_ETD_Export_HISTORY" & vbCrLf & _
                         " SELECT * FROM MS_ETD_Export WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & pWeek & "'"
                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                ls_SQL = "DELETE from dbo.MS_ETD_Export WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & pWeek & "'"
                SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                'insert into history
                ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                         "VALUES ('" & shostname & "','" & menuID & "','D','','Delete Affiliate " & pAffiliateID & ", Supplier " & pSupplierID & ", Week " & pWeek & ", Period " & pPeriod & "', GETDATE(),'" & Session("UserID") & "')  "
                SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        Call up_GridLoad()
        lblInfo.Visible = True
    End Sub

    Private Sub up_DeleteDataRec(ByVal pPeriod As String, ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pWeek As String)
        Dim ls_SQL As String = ""
        Dim x As Integer
        Dim shostname As String = System.Net.Dns.GetHostName

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ForwarderID")

                'ls_SQL = "DELETE from dbo.MS_ETD_Export WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & pWeek & "'"
                'Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                'sqlComm.ExecuteNonQuery()

                ls_SQL = " INSERT INTO MS_ETD_Export" & vbCrLf & _
                         " SELECT * FROM MS_ETD_Export_HISTORY WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & pWeek & "'"
                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                ls_SQL = "DELETE from dbo.MS_ETD_Export_HISTORY WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & Trim(pAffiliateID) & "' AND SupplierID = '" & Trim(pSupplierID) & "' AND Week = '" & pWeek & "'"
                SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                'insert into history
                ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                         "VALUES ('" & shostname & "','" & menuID & "','R','','Recovery Affiliate " & pAffiliateID & ", Supplier " & pSupplierID & ", Week " & pWeek & ", Period " & pPeriod & "', GETDATE(),'" & Session("UserID") & "')  "
                SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblInfo, "1016", clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        Call up_GridLoad()
        lblInfo.Visible = True
    End Sub
#End Region

#Region "FUNCTION"
    Private Sub DeleteHistory()
        Dim ls_sql As String

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                    Dim sqlComm As New SqlCommand()

                    ls_sql = " delete MS_ETD_Export_History " & vbCrLf & _
                              " where exists " & vbCrLf & _
                              " (select * from MS_ETD_Export a where MS_ETD_Export_History.Period = a.Period and MS_ETD_Export_History.SupplierID = a.SupplierID " & vbCrLf & _
                              "   and MS_ETD_Export_History.AffiliateID = a.AffiliateID and a.Week = MS_ETD_Export_History.Week) "

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

    Private Function validasiInput(ByVal pWeek As String) As Boolean
        Try
            Dim sqlstring As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                sqlstring = "SELECT * FROM dbo.MS_ETD_Export WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND SupplierID = '" & Trim(cboSupplier.Text) & "' AND WEEK = '" & pWeek & "'"
                Dim sqlDA As New SqlDataAdapter(sqlstring, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    lblInfo.Visible = True
                    lblInfo.Text = "Data with Affiliate Code " & Trim(cboAffiliate.Text) & " and Supplier Code " & Trim(cboSupplier.Text) & " already exists in the database. Data updated"
                    cboAffiliate.Focus()
                    Return False
                End If
                Return True
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Me.lblInfo.Visible = True
            Me.lblInfo.Text = ex.Message.ToString
        End Try
    End Function

    Private Function validasiETD(ByVal pETDVendor As String, ByVal pETAForwarder As String, ByVal pETDPort As String, ByVal pETAPort As String, ByVal pETAFactory As String) As Boolean
        Try
            Dim ls_ETDvendor As Date = Mid(pETDVendor, 5, 11)
            Dim ls_ETAForwarder As Date = Mid(pETAForwarder, 5, 11)
            Dim ls_ETDPort As Date = Mid(pETDPort, 5, 11)
            Dim ls_ETAPort As Date = Mid(pETAPort, 5, 11)
            Dim ls_ETAFactory As Date = Mid(pETAFactory, 5, 11)

            Dim sqlstring As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                sqlstring = "Select * From [dbo].[MS_ETD_Export] Where Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "' And SupplierID = '" & Trim(cboSupplier.Text) & "' And AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                            "And [ETDVendor] = '" & Format(ls_ETDvendor, "yyyy-MM-dd") & "' And [ETAForwarder] = '" & Format(ls_ETAForwarder, "yyyy-MM-dd") & "' And [ETDPort] = '" & Format(ls_ETDPort, "yyyy-MM-dd") & "' And [ETAPort] = '" & Format(ls_ETAPort, "yyyy-MM-dd") & "' And [ETAFactory] = '" & Format(ls_ETAFactory, "yyyy-MM-dd") & "'"

                'sqlstring = "SELECT * FROM dbo.MS_ETD_Export WHERE Period = '" & Format(DtPeriod.Value, "yyyy-MM-01") & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND SupplierID = '" & Trim(cboSupplier.Text) & "' AND WEEK = '" & pWeek & "'"
                Dim sqlDA As New SqlDataAdapter(sqlstring, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    lblInfo.Visible = True
                    lblInfo.Text = "Data already exists in Week " & Trim(ds.Tables(0).Rows(0)("Week")) & ""
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Return False
                Else
                    Return True
                End If
                'Return True
                grid.JSProperties("cpMessage") = lblInfo.Text
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Me.lblInfo.Visible = True
            Me.lblInfo.Text = ex.Message.ToString
        End Try
    End Function

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "select RTRIM(AffiliateID) AffiliateCode, RTRIM(AffiliateName) AffiliateName from MS_Affiliate order by AffiliateCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateCode")
                .Columns(0).Width = 75
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 400

                .TextField = "AffiliateCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplier
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 75
                .Columns.Add("SupplierName")
                .Columns(1).Width = 400

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select '1' WEEK UNION ALL select '2' WEEK UNION ALL select '3' WEEK UNION ALL select '4' WEEK UNION ALL select '5' WEEK" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboWeek
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("WEEK")
                .Columns(0).Width = 60

                .TextField = "WEEK"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub up_FillCombo2()
        Dim ls_SQL As String = ""

        ls_SQL = "select '1' WEEK UNION ALL select '2' WEEK UNION ALL select '3' WEEK UNION ALL select '4' WEEK UNION ALL select '5' WEEK" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboWeek
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("WEEK")
                .Columns(0).Width = 60

                .TextField = "WEEK"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadETDExport.aspx")
    End Sub

#End Region

    Private Sub btnRefresh_Click(sender As Object, e As System.EventArgs) Handles btnRefresh.Click
        Call up_GridLoad()

        If grid.VisibleRowCount = 0 Then
            Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblInfo.Text
        ElseIf grid.VisibleRowCount > 1 Then
            grid.JSProperties("cpMessage") = ""
            lblInfo.Text = ""
        End If

        grid.FocusedRowIndex = -1
    End Sub
End Class