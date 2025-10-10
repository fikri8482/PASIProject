Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO
Imports System.Web.Services

Public Class PartMapping
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

    Dim menuID As String = "A07"

    Public PartNos As String = ""
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), menuID)
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "PART MAPPING MASTER"
            up_FillCombo()
            DeleteHistory()
            If Session("M01Url") <> "" Then
                Call bindData()
                Session.Remove("M01Url")
            End If

            lblInfo.Text = ""
        End If

        If ls_AllowDownload = False Then btnDownload.Enabled = False
        If ls_AllowUpdate = False Then btnUpload.Enabled = False
        If ls_AllowUpdate = False Then btnSubmit.Enabled = False
        If ls_AllowDelete = False Then btnDelete.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" _
            Or e.Column.FieldName = "AffiliateID" Or e.Column.FieldName = "AffiliateName" Or e.Column.FieldName = "SupplierID" _
            Or e.Column.FieldName = "SupplierName" Or e.Column.FieldName = "Quota" Or e.Column.FieldName = "LocationID") _
            And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If grid.VisibleRowCount > 0 Then
            If e.GetValue("DeleteCls") = "1" Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadPartMapping.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(4), _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(5), _
                                     Split(e.Parameters, "|")(6), _
                                     Split(e.Parameters, "|")(7), _
                                     Split(e.Parameters, "|")(8), _
                                     Split(e.Parameters, "|")(9), _
                                     Split(e.Parameters, "|")(10), _
                                     Split(e.Parameters, "|")(11), _
                                     Split(e.Parameters, "|")(12), _
                                     Split(e.Parameters, "|")(13), _
                                     Split(e.Parameters, "|")(14), _
                                     Split(e.Parameters, "|")(15))
                    bindData()
                Case "delete"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim pSupplierID As String = Split(e.Parameters, "|")(3)
                    Dim pPartNo As String = Split(e.Parameters, "|")(1)
                    If AlreadyUsed(pAffiliateID) = False Then
                        If HF.Get("DeleteCls") = "0" Then
                            Call DeleteData(pAffiliateID, pSupplierID, pPartNo)
                        Else
                            Call DeleteDataRec(pAffiliateID, pSupplierID, pPartNo)
                        End If
                    End If
                    'bindData()
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

                    Dim dtProd As DataTable = clsMaster.GetTablePartMapping(txtPartNo.Text, cboAffiliate.Text, cboSupplier.Text)
                    FileName = "TemplateMSPartMapping.xlsx"
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

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtPartnoDetail.ReadOnly = False
        txtPartnoDetail.BackColor = Color.FromName("#FFFFFF")
        cboAffiliate2.ReadOnly = False
        cboAffiliate2.BackColor = Color.FromName("#FFFFFF")
        cboPacking.ReadOnly = False
        cboPacking.BackColor = Color.FromName("#FFFFFF")

        txtPartNo2.Text = ""
        txtAffiliate2.Text = ""
        txtSupplier2.Text = ""
        txtQuota.Text = "0"
        txtLocation.Text = ""
        txtPacking.Text = ""
        txtMOQ.Text = "0"
        txtQtyBox.Text = "0"
        txtBoxPallet.Text = "0"
        txtNetWeight.Text = "0"
        txtGrossWeight.Text = "0"
        txtLength.Text = "0"
        txtWidth.Text = "0"
        txtHeight.Text = "0"

        up_FillCombo()

        up_GridLoadWhenEventChange()

        lblInfo.Text = ""

    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboAffiliate.Text.Trim <> "" Then
            If cboAffiliate.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + " and AffiliateID = '" & cboAffiliate.Text.Trim & "' "
            End If
        End If

        If cboSupplier.Text.Trim <> "" Then
            If cboSupplier.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + "and SupplierID = '" & cboSupplier.Text.Trim & "' "
            End If
        End If

        If txtPartNo.Text <> "" Then
            If txtPartNo.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + "and PartNo like '%" & txtPartNo.Text & "%' "
            End If
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select " & vbCrLf & _
                    " 	row_number() over (order by PartNo, AffiliateID, SupplierID) NoUrut, * FROM (" & vbCrLf & _
                    " 	SELECT RTRIM(a.PartNo)PartNo, " & vbCrLf & _
                    " 	RTRIM(d.PartName)PartName, " & vbCrLf & _
                    " 	RTRIM(a.SupplierID)SupplierID, " & vbCrLf & _
                    " 	RTRIM(c.SupplierName)SupplierName, " & vbCrLf & _
                    " 	RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                    " 	RTRIM(b.AffiliateName)AffiliateName, " & vbCrLf & _
                    " 	RTRIM(a.LocationID)LocationID, " & vbCrLf & _
                    " 	ISNULL(a.Quota, 0)Quota, " & vbCrLf & _
                    " 	RTRIM(a.PackingCls) PackingCls, " & vbCrLf & _
                    " 	RTRIM(e.[Description])PackingDesc, " & vbCrLf & _
                    " 	ISNULL(a.MOQ, 0)MOQ, ISNULL(a.QtyBox, 0)QtyBox, ISNULL(a.BoxPallet, 0)BoxPallet, " & vbCrLf & _
                    " 	ISNULL(a.NetWeight, 0)NetWeight, ISNULL(a.GrossWeight, 0)GrossWeight, ISNULL(a.[Length], 0)[Length], ISNULL(a.Width, 0)Width, ISNULL(a.Height, 0)Height, 0 DeleteCls, a.EntryDate, a.EntryUser, a.UpdateDate, a.UpdateUser " & vbCrLf & _
                    " from MS_PartMapping a " & vbCrLf & _
                    " left join MS_Affiliate b on a.AffiliateID = b.AffiliateID " & vbCrLf & _
                    " left join MS_Supplier c on a.SupplierID = c.SupplierID " & vbCrLf & _
                    " left join MS_Parts d on a.PartNo = d.PartNo " & vbCrLf & _
                    " left join MS_PackingCls e on a.PackingCls = e.PackingCls " & vbCrLf

            ls_SQL = ls_SQL + "UNION ALL select " & vbCrLf & _
                    " 	RTRIM(a.PartNo)PartNo, " & vbCrLf & _
                    " 	RTRIM(d.PartName)PartName, " & vbCrLf & _
                    " 	RTRIM(a.SupplierID)SupplierID, " & vbCrLf & _
                    " 	RTRIM(c.SupplierName)SupplierName, " & vbCrLf & _
                    " 	RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                    " 	RTRIM(b.AffiliateName)AffiliateName, " & vbCrLf & _
                    " 	RTRIM(a.LocationID)LocationID, " & vbCrLf & _
                    " 	ISNULL(a.Quota, 0)Quota, " & vbCrLf & _
                    " 	RTRIM(a.PackingCls) PackingCls, " & vbCrLf & _
                    " 	RTRIM(e.[Description])PackingDesc, " & vbCrLf & _
                    " 	ISNULL(a.MOQ, 0)MOQ, ISNULL(a.QtyBox, 0)QtyBox, ISNULL(a.BoxPallet, 0)BoxPallet, " & vbCrLf & _
                    " 	ISNULL(a.NetWeight, 0)NetWeight, ISNULL(a.GrossWeight, 0)GrossWeight, ISNULL(a.[Length], 0)[Length], ISNULL(a.Width, 0)Width, ISNULL(a.Height, 0)Height, 1 DeleteCls, a.EntryDate, a.EntryUser, a.UpdateDate, a.UpdateUser " & vbCrLf & _
                    " from MS_PartMapping_History a " & vbCrLf & _
                    " inner join MS_Affiliate b on a.AffiliateID = b.AffiliateID " & vbCrLf & _
                    " inner join MS_Supplier c on a.SupplierID = c.SupplierID " & vbCrLf & _
                    " inner join MS_Parts d on a.PartNo = d.PartNo " & vbCrLf & _
                    " left join MS_PackingCls e on a.PackingCls = e.PackingCls " & _
                    " )xyz where 'A' = 'A' " & pWhere & ""

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

            ls_SQL = " select top 0  '' NoUrut, '' as PartNo, '' as LocationID, '' PartName, ''AffiliateID, '' AffiliateName, '' SupplierID, '' SupplierName, '' Quota"

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

        'Affiliate
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' AffiliateCode, '" & clsGlobal.gs_All & "' AffiliateName union all select RTRIM(AffiliateID) AffiliateCode, AffiliateName from MS_Affiliate order by AffiliateCode " & vbCrLf
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
                .Columns(0).Width = 140
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 300

                .TextField = "AffiliateCode"
                .DataBind()
                .SelectedIndex = 0
                txtAffiliate.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        'Supplier
        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierCode, '" & clsGlobal.gs_All & "' SupplierName union all select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
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
                .Columns(0).Width = 140
                .Columns.Add("SupplierName")
                .Columns(1).Width = 300

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = 0
                txtSupplier.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        ''Part
        'ls_SQL = "select RTRIM(PartNo) PartCode, PartName from MS_Parts order by PartCode " & vbCrLf
        'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
        '    sqlConn.Open()

        '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
        '    Dim ds As New DataSet
        '    sqlDA.Fill(ds)

        '    With cboPartNo2
        '        .Items.Clear()
        '        .Columns.Clear()
        '        .DataSource = ds.Tables(0)
        '        .Columns.Add("PartCode")
        '        .Columns(0).Width = 100
        '        .Columns.Add("PartName")
        '        .Columns(1).Width = 200

        '        .TextField = "PartCode"
        '        .DataBind()
        '        .SelectedIndex = -1
        '    End With

        '    sqlConn.Close()
        'End Using

        'Affiliate
        ls_SQL = "select RTRIM(AffiliateID) AffiliateCode, AffiliateName from MS_Affiliate order by AffiliateCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateCode")
                .Columns(0).Width = 100
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 200

                .TextField = "AffiliateCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        'Supplier
        ls_SQL = "select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplier2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 100
                .Columns.Add("SupplierName")
                .Columns(1).Width = 200

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(PackingCls) [Packing Group], RTRIM(Description) [Packing Description] from MS_PackingCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPacking
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Packing Group")
                .Columns(0).Width = 100
                .Columns.Add("Packing Description")
                .Columns(1).Width = 200

                .TextField = "Packing Group"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Function AlreadyUsed(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo FROM PO_Detail WHERE PartNo = '" & Trim(pAffiliate) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    Call clsMsg.DisplayMessage(lblInfo, "5004", clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    grid.JSProperties("cpType") = "error"
                    Return True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Return False
    End Function

    Private Sub DeleteHistory()
        Dim ls_sql As String

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("Price")

                    Dim sqlComm As New SqlCommand()

                    ls_sql = " delete MS_PartMapping_History " & vbCrLf & _
                              " where exists " & vbCrLf & _
                              " (select * from MS_PartMapping a where MS_PartMapping_History.AffiliateID = a.AffiliateID and MS_PartMapping_History.PartNo = a.PartNo " & vbCrLf & _
                              "   and MS_PartMapping_History.SupplierID = a.SupplierID ) "

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

    Private Sub DeleteData(ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pPartNo As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    'ls_SQL = " DELETE MS_PartMapping " & vbCrLf & _
                    '            " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'" & vbCrLf

                    'Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    'x = SqlComm.ExecuteNonQuery()
                    ls_SQL = " INSERT INTO MS_PartMapping_HISTORY" & vbCrLf & _
                             " SELECT * FROM MS_PartMapping " & vbCrLf & _
                             " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'" & vbCrLf
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " DELETE MS_PartMapping " & vbCrLf & _
                             " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'" & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    'insert into history
                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','D','" & pPartNo & "','Delete PartNo " & pPartNo & ", Affiliate " & pAffiliateID & ", Supplier " & pSupplierID & "', GETDATE(),'" & Session("UserID") & "')  "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            If x > 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                grid.JSProperties("cpType") = "info"
                grid.JSProperties("cpFunction") = "delete"
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub DeleteDataRec(ByVal pAffiliateID As String, ByVal pSupplierID As String, ByVal pPartNo As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    'ls_SQL = " DELETE MS_PartMapping " & vbCrLf & _
                    '            " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'" & vbCrLf

                    'Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    'x = SqlComm.ExecuteNonQuery()
                    ls_SQL = " INSERT INTO MS_PartMapping" & vbCrLf & _
                             " SELECT * FROM MS_PartMapping_HISTORY " & vbCrLf & _
                             " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'" & vbCrLf
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " DELETE MS_PartMapping_HISTORY " & vbCrLf & _
                             " WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'" & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    'insert into history
                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','R','" & pPartNo & "','Recovery PartNo " & pPartNo & ", Affiliate " & pAffiliateID & ", Supplier " & pSupplierID & "', GETDATE(),'" & Session("UserID") & "')  "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            If x > 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "1016", clsMessage.MsgType.InformationMessage)
                grid.JSProperties("cpMessage") = lblInfo.Text
                grid.JSProperties("cpType") = "info"
                grid.JSProperties("cpFunction") = "delete"
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Return False
    End Function

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                        Optional ByVal pAffiliateID As String = "", _
                        Optional ByVal pSupplierID As String = "", _
                        Optional ByVal pPartNo As String = "", _
                        Optional ByVal pQuota As String = "", _
                        Optional ByVal pLocation As String = "", _
                        Optional ByVal pPackingID As String = "", _
                        Optional ByVal pMOQ As String = "", _
                        Optional ByVal pQtyBox As String = "", _
                        Optional ByVal pBoxPallet As String = "", _
                        Optional ByVal pNetWeight As String = "", _
                        Optional ByVal pGrossWeight As String = "", _
                        Optional ByVal pLength As String = "", _
                        Optional ByVal pWidth As String = "", _
                        Optional ByVal pHeight As String = "")


        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName
        Dim pub_SupplierID As String
        Dim pub_PartNo As String
        Dim pub_Quota As String
        Dim pub_Location As String
        Dim pub_PackingID As String
        Dim pub_MOQ As String
        Dim pub_QtyBox As String
        Dim pub_BoxPallet As String
        Dim pub_NetWeight As String
        Dim pub_GrossWeight As String
        Dim pub_Length As String
        Dim pub_Width As String
        Dim pub_Height As String

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo, SupplierID, Quota, LocationID, PackingCls, MOQ, QtyBox, BoxPallet, NetWeight, GrossWeight, Length, Width, Height FROM MS_PartMapping WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False

                    'pSupplierID = ds.Tables(0).Rows(0)("SupplierID")
                    'pPartNo = ds.Tables(0).Rows(0)("PartNo")
                    pub_Quota = ds.Tables(0).Rows(0)("Quota")
                    pub_Location = ds.Tables(0).Rows(0)("LocationID")
                    pub_PackingID = ds.Tables(0).Rows(0)("PackingCls")
                    pub_MOQ = ds.Tables(0).Rows(0)("MOQ")
                    pub_QtyBox = ds.Tables(0).Rows(0)("QtyBox")
                    pub_BoxPallet = ds.Tables(0).Rows(0)("BoxPallet")
                    pub_NetWeight = ds.Tables(0).Rows(0)("NetWeight")
                    pub_GrossWeight = ds.Tables(0).Rows(0)("GrossWeight")
                    pub_Length = ds.Tables(0).Rows(0)("Length")
                    pub_Width = ds.Tables(0).Rows(0)("Width")
                    pub_Height = ds.Tables(0).Rows(0)("Height")
                Else
                    pIsNewData = True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CostCenter")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_PartMapping " & _
                                "(PartNo, AffiliateID, SupplierID, Quota, LocationID, " & _
                                "PackingCls, MOQ, QtyBox, BoxPallet, NetWeight, GrossWeight, Length, Width, Height, EntryDate, EntryUser) " & _
                                "VALUES ('" & pPartNo & "','" & pAffiliateID & "','" & pSupplierID & "','" & pQuota & "','" & pLocation & "', " & _
                                "'" & pPackingID & "','" & pMOQ & "','" & pQtyBox & "','" & pBoxPallet & "','" & pNetWeight & "', '" & pGrossWeight & "', '" & pLength & "', '" & pWidth & "', '" & pHeight & "', " & _
                                "getdate(),'" & admin & "')" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"
                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    ls_SQL = "UPDATE MS_PartMapping SET " & _
                            "SupplierID ='" & pSupplierID & "'," & _
                            "Quota ='" & pQuota & "'," & _
                            "LocationID ='" & pLocation & "'," & _
                            "PackingCls ='" & pPackingID & "'," & _
                            "MOQ ='" & pMOQ & "'," & _
                            "QtyBox ='" & pQtyBox & "'," & _
                            "BoxPallet ='" & pBoxPallet & "'," & _
                            "NetWeight ='" & pNetWeight & "'," & _
                            "GrossWeight ='" & pGrossWeight & "'," & _
                            "Length ='" & pLength & "'," & _
                            "Width ='" & pWidth & "'," & _
                            "Height ='" & pHeight & "'," & _
                            "UpdateDate = getdate()," & _
                            "UpdateUser ='" & admin & "'" & _
                            "WHERE PartNo='" & pPartNo & "' and AffiliateID ='" & pAffiliateID & "' and SupplierID = '" & pSupplierID & "'"
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "update"


                    'pSupplierID = ds.Tables(0).Rows(0)("SupplierID")
                    'pPartNo = ds.Tables(0).Rows(0)("PartNo")
                    'pQuota = ds.Tables(0).Rows(0)("Quota")
                    'pLocation = ds.Tables(0).Rows(0)("LocationID")
                    'pPackingID = ds.Tables(0).Rows(0)("PackingCls")
                    ' pMOQ = ds.Tables(0).Rows(0)("MOQ")
                    'pQtyBox = ds.Tables(0).Rows(0)("QtyBox")
                    'pBoxPallet = ds.Tables(0).Rows(0)("BoxPallet")
                    'pNetWeight = ds.Tables(0).Rows(0)("NetWeight")
                    'pGrossWeight = ds.Tables(0).Rows(0)("GrossWeight")
                    'pLength = ds.Tables(0).Rows(0)("Length")
                    'pWidth = ds.Tables(0).Rows(0)("Width")
                    'pHeight = ds.Tables(0).Rows(0)("Height")

                    Dim ls_Remarks As String = ""

                    If CDbl(pQuota) <> CDbl(pub_Quota) Then
                        ls_Remarks = ls_Remarks + "Quota " + pub_Quota & "->" & pQuota & " "
                    End If

                    If CDbl(pMOQ) <> CDbl(pub_MOQ) Then
                        ls_Remarks = ls_Remarks + "MOQ " + pub_MOQ & "->" & pMOQ & " "
                    End If

                    If pLocation <> pub_Location Then
                        ls_Remarks = ls_Remarks + "Location " + pub_Location & "->" & pLocation & " "
                    End If

                    If pPackingID <> pub_PackingID Then
                        ls_Remarks = ls_Remarks + "Packing " + pub_PackingID & "->" & pPackingID & " "
                    End If

                    If CDbl(pQtyBox) <> CDbl(pub_QtyBox) Then
                        ls_Remarks = ls_Remarks + "Qty Box " + pub_QtyBox & "->" & pQtyBox & " "
                    End If

                    If CDbl(pBoxPallet) <> CDbl(pub_BoxPallet) Then
                        ls_Remarks = ls_Remarks + "Box Pallet " + pub_BoxPallet & "->" & pBoxPallet & " "
                    End If
                    If CDbl(pNetWeight) <> CDbl(pub_NetWeight) Then
                        ls_Remarks = ls_Remarks + "Net Weight " + pub_NetWeight & "->" & pNetWeight & " "
                    End If
                    If CDbl(pGrossWeight) <> CDbl(pub_GrossWeight) Then
                        ls_Remarks = ls_Remarks + "Gross Weight " + pub_GrossWeight & "->" & pGrossWeight & " "
                    End If
                    If CDbl(pLength) <> CDbl(pub_Length) Then
                        ls_Remarks = ls_Remarks + "Length " + pub_Length & "->" & pLength & " "
                    End If
                    If CDbl(pWidth) <> CDbl(pub_Width) Then
                        ls_Remarks = ls_Remarks + "Width " + pub_Width & "->" & pWidth & " "
                    End If
                    If CDbl(pHeight) <> CDbl(pub_Height) Then
                        ls_Remarks = ls_Remarks + "Height " + pub_Quota & "->" & pQuota & " "
                    End If
                    Dim ls_Remarks2 As String = "Affiliate " & pAffiliateID & " Supplier " & pSupplierID & ", "

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U','" & pPartNo & "','Update [" & ls_Remarks2 & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                    End If
                End If

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        grid.JSProperties("cpType") = "info"

    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Part Mapping Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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

                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 14)
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

    Private Sub PartNoCallBack_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles PartNoCallBack.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Select Case pAction
            Case "Load"
                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()
                    Dim partno = Split(e.Parameter, "|")(1).Trim
                    Dim sqlDA As New SqlDataAdapter("EXEC sp_PartNoList '" & partno & "'", sqlConn)
                    Dim ds As New DataSet
                    sqlDA.Fill(ds)
                    PartNoCallBack.JSProperties("cpPartno") = partno
                    If ds.Tables(0).Rows.Count > 0 Then
                        PartNoCallBack.JSProperties("cpPartnosNames") = ds.Tables(0).Rows(0)("PartName").ToString()
                    Else
                        PartNoCallBack.JSProperties("cpPartnosNames") = ""
                    End If
                    sqlConn.Close()
                End Using
        End Select
    End Sub
End Class