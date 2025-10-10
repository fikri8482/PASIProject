Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports OfficeOpenXml
Imports System.IO

Public Class SupplierCapacity
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
    Dim menuID As String = "A09"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowDownload As String = clsGlobal.Auth_UserConfirm(Session("UserID"), menuID)
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "SUPPLIER CAPACITY MASTER"
            up_FillCombo()
            txtDailyCapacity.Text = 0
            txtMontlyCapacity.Text = 0
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

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "SupplierID" Or e.Column.FieldName = "SupplierName" _
            Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" Or e.Column.FieldName = "DailyCapacity" _
            Or e.Column.FieldName = "MontlyCapacity") _
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

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        bindData()
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared      
        If grid.VisibleRowCount > 0 Then
            If e.GetValue("DeleteCls") = "1" Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Response.Redirect("~/Upload/UploadSuppCapacity.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            'grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(4), _
                                     Split(e.Parameters, "|")(5))
                    'bindData()
                Case "delete"
                    Dim pSupplierID As String = Split(e.Parameters, "|")(1)
                    Dim pPartNo As String = Split(e.Parameters, "|")(2)
                    If AlreadyUsed(pPartNo) = False Then
                        If HF.Get("DeleteCls") = "0" Then
                            Call DeleteData(pSupplierID, pPartNo)
                        Else
                            Call DeleteDataRec(pSupplierID, pPartNo)
                        End If
                    End If
                    'bindData()
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    ElseIf grid.VisibleRowCount > 1 Then
                        grid.JSProperties("cpMessage") = ""
                        lblInfo.Text = ""
                    End If

                    grid.FocusedRowIndex = -1
                Case "loadaftersubmit"
                    Call bindData()
                    grid.FocusedRowIndex = -1
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
                Case "downloadSummary"
                    Dim psERR As String = ""
                    Dim pSuppType As String = ""
                    'If cbSupplierType.Text <> clsGlobal.gs_All Then
                    '    If cbSupplierType.Text = "PASI SUPPLIER" Then
                    '        pSuppType = "1"
                    '    Else
                    '        pSuppType = "0"
                    '    End If
                    'Else
                    '    pSuppType = clsGlobal.gs_All
                    'End If
                    Dim dtProd As DataTable = clsMaster.GetTableSuppCapacity(cboSupplierCode.Text, cboPartNo.Text)
                    FileName = "TemplateMSSupCapacity.xlsx"
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
        cboPartNo2.ReadOnly = False
        cboPartNo2.BackColor = Color.FromName("#FFFFFF")
        cboSupplierCode2.ReadOnly = False
        cboSupplierCode2.BackColor = Color.FromName("#FFFFFF")
        txtMode.Text = "new"
        txtDailyCapacity.Text = 0
        txtMontlyCapacity.Text = 0
        txtSupplierCode2.Text = ""
        txtPartNo2.Text = ""
        up_FillCombo()

        up_GridLoadWhenEventChange()

        lblInfo.Text = ""

    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboSupplierCode.Text.Trim <> "" Then
            If cboSupplierCode.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + "and SupplierID like '%" & cboSupplierCode.Text.Trim & "%' "
            End If
        End If

        If cboPartNo.Text <> "" Then
            If cboPartNo.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + "and PartNo like '%" & cboPartNo.Text & "%' "
            End If
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT row_number() over (order by SupplierID, PartNo) NoUrut, * FROM (select " & vbCrLf & _
                  " RTRIM(a.SupplierID) SupplierID, b.SupplierName, RTRIM(a.PartNo) PartNo, c.PartName, a.DailyDeliveryCapacity DailyCapacity, a.MonthlyInjectionCapacity MontlyCapacity, 0 DeleteCls, a.EntryDate, a.EntryUser, a.UpdateDate, a.UpdateUser " & vbCrLf & _
                  " from MS_SupplierCapacity a " & vbCrLf & _
                  " left join MS_Supplier b on a.SupplierID = b.SupplierID " & vbCrLf & _
                  " left join MS_Parts c on a.PartNo = c.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " UNION ALL select " & vbCrLf & _
                  " RTRIM(a.SupplierID) SupplierID, b.SupplierName, RTRIM(a.PartNo) PartNo, c.PartName, a.DailyDeliveryCapacity DailyCapacity, a.MonthlyInjectionCapacity MontlyCapacity, 1 DeleteCls, a.EntryDate, a.EntryUser, a.UpdateDate, a.UpdateUser " & vbCrLf & _
                  " from MS_SupplierCapacity_History a " & vbCrLf & _
                  " left join MS_Supplier b on a.SupplierID = b.SupplierID " & vbCrLf & _
                  " left join MS_Parts c on a.PartNo = c.PartNo " & vbCrLf & _
                  " )XYZ where 'a' = 'a' " & pWhere & ""

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' as SupplierID, '' SupplierName, '' PartNo, '' PartName, '' DailyCapacity, '' MontlyCapacity"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()

            End With

            sqlConn.Close()
        End Using

        grid.FocusedRowIndex = -1
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all select RTRIM(PartNo) PartCode, PartName from MS_Parts order by PartCode " & vbCrLf
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


        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierCode, '" & clsGlobal.gs_All & "' SupplierName union all select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 85
                .Columns.Add("SupplierName")
                .Columns(1).Width = 180

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = 0
                txtSupplierCode.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(PartNo) PartCode, PartName, b.Description UnitDesc from MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls order by PartCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartCode")
                .Columns(0).Width = 85
                .Columns.Add("PartName")
                .Columns(1).Width = 180
                .Columns.Add("UnitDesc")
                .Columns(2).Width = 25

                .TextField = "PartCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(SupplierID) SupplierCode, SupplierName from MS_Supplier order by SupplierCode" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 85
                .Columns.Add("SupplierName")
                .Columns(1).Width = 180

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = -1
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Function AlreadyUsed(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo FROM PO_Detail WHERE PartNo= '" & Trim(pAffiliate) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    'lblInfo.Text = "Affiliate ID already used in other screen"
                    Call clsMsg.DisplayMessage(lblInfo, "5004", clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    grid.JSProperties("cpType") = "error"
                    Return True
                Else
                    Return False
                End If
                Return True
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Function

    Private Sub DeleteData(ByVal pSupplierID As String, ByVal pPartNo As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    'ls_SQL =    " DELETE MS_SupplierCapacity " & vbCrLf & _
                    '            " WHERE SupplierID='" & pSupplierID & "' and PartNo ='" & pPartNo & "'" & vbCrLf

                    'Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    'x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " INSERT INTO MS_SupplierCapacity_HISTORY " & vbCrLf & _
                             " SELECT * FROM MS_SupplierCapacity " & vbCrLf & _
                              " WHERE SupplierID='" & pSupplierID & "' and PartNo ='" & pPartNo & "'" & vbCrLf
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " DELETE MS_SupplierCapacity " & vbCrLf & _
                             " WHERE SupplierID='" & pSupplierID & "' and PartNo ='" & pPartNo & "'" & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    'insert into history
                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','D','" & pPartNo & "','Delete PartNo " & pPartNo & ", SupplierID " & pSupplierID & "', GETDATE(),'" & Session("UserID") & "')  "
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

    Private Sub DeleteDataRec(ByVal pSupplierID As String, ByVal pPartNo As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    'ls_SQL = " DELETE MS_SupplierCapacity " & vbCrLf & _
                    '            " WHERE SupplierID='" & pSupplierID & "' and PartNo ='" & pPartNo & "'" & vbCrLf

                    'Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    'x = SqlComm.ExecuteNonQuery()
                    ls_SQL = " INSERT INTO MS_SupplierCapacity" & vbCrLf & _
                             " SELECT * FROM MS_SupplierCapacity_HISTORY " & vbCrLf & _
                             " WHERE SupplierID='" & pSupplierID & "' and PartNo ='" & pPartNo & "'" & vbCrLf
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    ls_SQL = " DELETE MS_SupplierCapacity_HISTORY " & vbCrLf & _
                             " WHERE SupplierID='" & pSupplierID & "' and PartNo ='" & pPartNo & "'" & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    'insert into history
                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','R','" & pPartNo & "','Recovery PartNo " & pPartNo & ", Supplier " & pSupplierID & "', GETDATE(),'" & Session("UserID") & "')  "
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

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try
            'Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""

            'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            '    sqlConn.Open()

            '    ls_SQL = "SELECT AffiliateID" & vbCrLf & _
            '                " FROM MS_Affiliate " & _
            '                " WHERE AffiliateID= '" & Trim(pAffiliate) & "'"

            '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            '    Dim ds As New DataSet
            '    sqlDA.Fill(ds)

            '    If ds.Tables(0).Rows.Count > 0 And grid.FocusedRowIndex = -1 Then
            '        ls_MsgID = "6018"
            '        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            '        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            '        flag = False
            '        Return False
            '    ElseIf ds.Tables(0).Rows.Count > 0 Then
            '        lblInfo.Text = "Affiliate ID with ID " & txtPartID.Text & " already exists in the database."
            '        Return False
            '    End If
            '    Return True
            '    sqlConn.Close()
            'End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Function

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pSupplierID As String = "", _
                         Optional ByVal pPartNo As String = "", _
                         Optional ByVal pDailyCapacity As String = "", _
                         Optional ByVal pMontlyCapacity As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim pub_1 As String
        Dim pub_2 As String
        Dim shostname As String = System.Net.Dns.GetHostName

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo, DailyDeliveryCapacity, MonthlyInjectionCapacity FROM MS_SupplierCapacity WHERE PartNo='" & pPartNo & "' and SupplierID ='" & pSupplierID & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False
                    pub_1 = ds.Tables(0).Rows(0)("DailyDeliveryCapacity")
                    pub_2 = ds.Tables(0).Rows(0)("MonthlyInjectionCapacity")
                Else
                    pIsNewData = True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        'If txtMode.Text = "update" Then
        '    flag = False
        'Else
        '    flag = True
        'End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CostCenter")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_SupplierCapacity " & _
                                "(SupplierID, PartNo, DailyDeliveryCapacity, MonthlyInjectionCapacity, EntryDate, EntryUser)" & _
                                " VALUES ('" & cboSupplierCode2.Text & "','" & cboPartNo2.Text & "','" & txtDailyCapacity.Text & "','" & txtMontlyCapacity.Text & "', getdate(),'" & admin & "')" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                    'ElseIf pIsNewData = False And flag = True Then
                    '    ls_MsgID = "6018"
                    '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    '    lblInfo.Visible = True
                    '    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    ls_SQL = "UPDATE MS_SupplierCapacity SET " & _
                            "DailyDeliveryCapacity='" & pDailyCapacity & "'," & _
                            "MonthlyInjectionCapacity='" & pMontlyCapacity & "'," & _
                            "UpdateDate = getdate()," & _
                            "UpdateUser ='" & admin & "'" & _
                            "WHERE PartNo='" & pPartNo & "' and SupplierID ='" & pSupplierID & "'"
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "update"

                    Dim ls_Remarks As String = ""

                    If CDbl(pDailyCapacity) <> CDbl(pub_1) Then
                        ls_Remarks = ls_Remarks + "Daily " + pub_1 & "->" & pDailyCapacity & " "
                    End If

                    If CDbl(pMontlyCapacity) <> CDbl(pub_2) Then
                        ls_Remarks = ls_Remarks + "Daily " + pub_2 & "->" & pMontlyCapacity & " "
                    End If

                    Dim ls_Remarks2 As String = "SupplierID " & pSupplierID & ", "

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
        lblInfo.Visible = True

    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Supplier Capacity Master " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                Dim rgAll As ExcelRange = .Cells(2, 1, irow + 2, 4)
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