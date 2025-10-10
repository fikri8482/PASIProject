Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing

Public Class PartConversion
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "PART CONVERSION MASTER"
            up_FillCombo()
            txtFGQty.Text = 0
            txtPartQty.Text = 0
            If Session("M01Url") <> "" Then
                Call bindData()
                Session.Remove("M01Url")
            End If

            lblInfo.Text = ""
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "FinishGoodNo" Or e.Column.FieldName = "FinishGoodName" _
            Or e.Column.FieldName = "FGUnitCls" Or e.Column.FieldName = "FGQty" Or e.Column.FieldName = "PartNo" _
            Or e.Column.FieldName = "PartName" Or e.Column.FieldName = "PartUnitCls" Or e.Column.FieldName = "PartQty") _
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
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameters, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliateID)
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(4), _
                                     Split(e.Parameters, "|")(5))

                    grid.FocusedRowIndex = -1
                    'bindData()
                Case "delete"
                    Dim pFGNo As String = Split(e.Parameters, "|")(1)
                    Dim pPartNo As String = Split(e.Parameters, "|")(2)
                    If AlreadyUsed(pFGNo) = False Then
                        Call DeleteData(pFGNo, pPartNo)
                    End If

                    grid.FocusedRowIndex = -1
                    'bindData()
                Case "load"
                    lblInfo.Text = ""
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

            End Select

            'EndProcedure:
            '            Session("AA220Msg") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        cboPartNo2.ReadOnly = False
        cboPartNo2.BackColor = Color.FromName("#FFFFFF")
        cboFGNo2.ReadOnly = False
        cboFGNo2.BackColor = Color.FromName("#FFFFFF")
        txtModee.Text = "new"
        txtFGQty.Text = 0
        txtPartQty.Text = 0
        txtFGUnit.Text = ""
        txtPartUnit.Text = ""
        up_FillCombo()

        up_GridLoadWhenEventChange()

        lblInfo.Text = ""

    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        If cboFGNo.Text.Trim <> "" Then
            If cboFGNo.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + "and a.FinishGoodNo like '%" & cboFGNo.Text.Trim & "%' "
            End If
        End If

        If cboPartNo.Text <> "" Then
            If cboPartNo.Text <> clsGlobal.gs_All Then
                pWhere = pWhere + "and a.PartNo like '%" & cboPartNo.Text & "%' "
            End If
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  "     row_number() over (order by a.FinishGoodNo, a.PartNo) NoUrut," & vbCrLf & _
                  " 	RTRIM(a.FinishGoodNo) FinishGoodNo, b.PartName FinishGoodName, d.Description FGUnitCls, Convert(numeric(18,0),a.FinishGoodQty) FGQty, " & vbCrLf & _
                  " 	RTRIM(a.PartNo) PartNo, c.PartName, e.Description PartUnitCls, Convert(numeric(18,0),a.PartQty)PartQty " & vbCrLf & _
                  " from MS_PartConversion a " & vbCrLf & _
                  " left join MS_Parts b on a.FinishGoodNo = b.PartNo " & vbCrLf & _
                  " left join MS_Parts c on a.PartNo = c.PartNo " & vbCrLf & _
                  " left join MS_UnitCls d on b.UnitCls = d.UnitCls " & vbCrLf & _
                  " left join MS_UnitCls e on c.UnitCls = e.UnitCls " & vbCrLf & _
                  " where 'A' = 'A' " & pWhere & ""

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0  '' NoUrut, '' as FinishGoodNo, '' FinishGoodName, ''FGUnitCls, '' FGQty, '' PartNo, '' PartName, '' PartUnitCls, '' PartQty"

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

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all select RTRIM(PartNo) PartCode, PartName from MS_Parts where FinishGoodCls = '2' order by PartCode " & vbCrLf
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


        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' PartCode, '" & clsGlobal.gs_All & "' PartName union all select RTRIM(PartNo) PartCode, PartName from MS_Parts where FinishGoodCls = '1' order by PartCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboFGNo
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
                txtFGNo.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "select RTRIM(PartNo) PartCode, PartName, b.Description UnitDesc from MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls where FinishGoodCls = '2'  order by PartCode " & vbCrLf
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

        ls_SQL = "select RTRIM(PartNo) PartCode, PartName, b.Description UnitDesc from MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls where FinishGoodCls = '1'  order by PartCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboFGNo2
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
    End Sub

    Private Function AlreadyUsed(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo FROM PO_Detail WHERE PartNo= '" & Trim(pAffiliate) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    'lblInfo.Text = "Affiliate ID already used in other screen"
                    ls_MsgID = "5004"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    lblInfo.Visible = True
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

    Private Sub DeleteData(ByVal pFGNo As String, ByVal pPartNo As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    ls_SQL = " DELETE MS_PartConversion " & vbCrLf & _
                                " WHERE FinishGoodNo='" & pFGNo & "' and PartNo ='" & pPartNo & "'" & vbCrLf

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
                         Optional ByVal pFGNo As String = "", _
                         Optional ByVal pPartNo As String = "", _
                         Optional ByVal pFGQty As String = "", _
                         Optional ByVal pPartQty As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT * FROM MS_PartConversion WHERE PartNo='" & pPartNo & "' and FinishGoodNo ='" & pFGNo & "'"

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

        If txtModee.Text = "update" Then
            flag = False
        Else
            flag = True
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CostCenter")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_PartConversion " & _
                                "(FinishGoodNo, FinishGoodQty, PartNo, PartQty, EntryDate, EntryUser)" & _
                                " VALUES ('" & cboFGNo2.Text & "','" & txtFGQty.Text & "','" & cboPartNo2.Text & "','" & txtPartQty.Text & "', getdate(),'" & admin & "')" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    grid.JSProperties("cpFunction") = "insert"

                ElseIf pIsNewData = False And flag = True Then
                    ls_MsgID = "6018"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    lblInfo.Visible = True
                    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    ls_SQL = "UPDATE MS_PartConversion SET " & _
                            "FinishGoodQty = '" & pFGQty & "'," & _
                            "PartQty = '" & pPartQty & "'," & _
                            "UpdateDate = getdate()," & _
                            "UpdateUser = '" & admin & "'" & _
                            "WHERE PartNo = '" & pPartNo & "' and FinishGoodNo = '" & pFGNo & "'"
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
        lblInfo.Visible = True

    End Sub
#End Region
End Class