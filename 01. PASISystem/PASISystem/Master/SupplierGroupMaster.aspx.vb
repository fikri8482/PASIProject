Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing

Partial Class SupplierGroupMaster
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
    Dim menuID As String = "A16"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
        Dim ls_AllowDelete As String = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Call bindData()
            ScriptManager.RegisterStartupScript(grid, grid.GetType(), "init", "grid.SetFocusedRowIndex(-1);", True)
            lblInfo.Text = ""
        End If

        If ls_AllowUpdate = False Then btnSubmit.Enabled = False
        If ls_AllowDelete = False Then btnDelete.Enabled = False

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "SupplierGroupCode" Or e.Column.FieldName = "Description") _
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

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            'grid.JSProperties("cpMessage") = Session("AA220Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pSupplierGroupCode As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pSupplierGroupCode)
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(3))

                Case "delete"
                    Dim pSupplierGroupCode As String = Split(e.Parameters, "|")(1)

                    'If AlreadyUsed(pSupplierGroupCode) = True Then
                    'pSupplierID,
                    Call DeleteData(pSupplierGroupCode)
                    Call bindData()
                    'End If
                    'txtMode.Text = "new"
                Case "load"
                    Call bindData()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If
                    'txtMode.Text = "new"
                Case "kosong"
                    Call up_GridLoadWhenEventChange()

            End Select

EndProcedure:
            Session("AA220Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub


#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        'If cboAffiliate.Text.Trim <> "" Then
        '    If cboAffiliate.Text <> clsGlobal.gs_All Then
        '        pWhere = pWhere + " and a.AffiliateID = '" & cboAffiliate.Text.Trim & "' "
        '    End If
        'End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'a.AffiliateID,
            '" 	RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
            '" 	RTRIM(b.AffiliateName)AffiliateName, " & vbCrLf & _
            ls_SQL = " select " & vbCrLf & _
                  " 	row_number() over (order by SupplierGroupCode) NoUrut, " & vbCrLf & _
                  " 	RTRIM(SupplierGroupCode)SupplierGroupCode, " & vbCrLf & _
                  " 	RTRIM(Description)Description " & vbCrLf & _
                  " from MS_SupplierGroup " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            End With
            sqlConn.Close()

            grid.FocusedRowIndex = -1

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        grid.FocusedRowIndex = -1

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ''AffiliateID, '' AffiliateName, 
            ls_SQL = " select top 0  '' NoUrut, '' SupplierGroupCode, '' Description"

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

    'Private Sub up_FillCombo()
    '    Dim ls_SQL As String = ""

    '    'Person In Charge

    '    ls_SQL = "select RTRIM(SupplierGroupCode) SupplierGroupCode, Description from MS_SupplierGroup order by SupplierGroupCode " & vbCrLf
    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        With cboSupplier2
    '            .Items.Clear()
    '            .Columns.Clear()
    '            .DataSource = ds.Tables(0)
    '            .Columns.Add("SupplierGroupCode")
    '            .Columns(0).Width = 75
    '            .Columns.Add("SupplierGroupName")
    '            .Columns(1).Width = 400

    '            .TextField = "SupplierGroupCode"
    '            .DataBind()
    '            .SelectedIndex = -1
    '        End With

    '        sqlConn.Close()
    '    End Using


    'End Sub

    Private Function AlreadyUsed(ByVal pSupplierGroupCode As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT SupplierGroupCode FROM MS_SupplierGroup WHERE SupplierGroupCode= '" & Trim(pSupplierGroupCode) & "'"

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

    Private Sub DeleteData(ByVal pSupplierGroupCode As String)
        'ByVal pSupplierID As String,
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                'and SupplierID = '" & pSupplierID & "'
                'and AffiliateID ='" & pAffiliateID & "'
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("SupplierGroup")
                    ls_SQL = " DELETE MS_SupplierGroup " & vbCrLf & _
                                " WHERE SupplierGroupCode='" & pSupplierGroupCode & "' " & vbCrLf

                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

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
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
        End Try
    End Sub

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            Dim ls_MsgID As String = ""


            'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            '    sqlConn.Open()

            '    ls_SQL = "SELECT AffiliateID, PartNo" & vbCrLf & _
            '                " FROM MS_Price " & _
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
            '        lblInfo.Text = "Affiliate ID with ID " & txtPartNo.Text & " already exists in the database."
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
                         Optional ByVal pSupplierGroupCode As String = "", _
                         Optional ByVal pSupplierGroupName As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        'and AffiliateID ='" & pAffiliateID & "'
        grid.FocusedRowIndex = -1

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT SupplierGroupCode FROM MS_SupplierGroup WHERE SupplierGroupCode ='" & pSupplierGroupCode & "' "
                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsNewData = False
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

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("SupplierGroup")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_SupplierGroup " & _
                                "(SupplierGroupCode, Description)" & _
                                " VALUES ('" & txtSupplier.Text & "','" & txtSupplier2.Text & "')" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                ElseIf pIsNewData = True And flag = False Then
                    ls_MsgID = "6018"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    grid.JSProperties("cpType") = "error"
                    grid.FocusedRowIndex = -1
                    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    'and SupplierID = '" & pSupplierID & "'
                    ls_SQL = "UPDATE MS_SupplierGroup SET " & _
                            "Description='" & txtSupplier2.Text & "' " & vbCrLf & _
                    " WHERE SupplierGroupCode='" & pSupplierGroupCode & "'" & vbCrLf
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "update"

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
#End Region


End Class