Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress

Public Class Forwarder
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "FORWARDER MASTER"
            lblInfo.Visible = True
            lblInfo.Text = ""
            Call TabIndex()
            txtForwarderCode.Focus()
            grid.FocusedRowIndex = -1
            Call up_GridLoad()
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowAllRecord)
        ScriptManager.RegisterStartupScript(grid, grid.GetType(), "scriptKey", "txtForwarderCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;'); grid.SetFocusedRowIndex(-1); grid.SetFocusedRowIndex(-1);", True)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        grid.JSProperties("cpMessage") = ""
        Try
            Select Case pAction
                Case "save"
                    Dim pIsUpdate As String = Split(e.Parameters, "|")(1)
                    Dim lb_IsUpdate As Boolean = validasiInput()
                    Call up_SaveData(lb_IsUpdate, _
                                     Split(e.Parameters, "|")(2), _
                                     Split(e.Parameters, "|")(3), _
                                     Split(e.Parameters, "|")(4), _
                                     Split(e.Parameters, "|")(5), _
                                     Split(e.Parameters, "|")(6), _
                                     Split(e.Parameters, "|")(7), _
                                     Split(e.Parameters, "|")(8), _
                                     Split(e.Parameters, "|")(9), _
                                     Split(e.Parameters, "|")(10), _
                                     Split(e.Parameters, "|")(11))

                Case "delete"
                    Dim pForwarderCode As String = Split(e.Parameters, "|")(1)

                    If AlreadyUsed(pForwarderCode) = False Then
                        Call up_DeleteData(pForwarderCode)
                    End If

                    grid.FocusedRowIndex = -1
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

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT ROW_NUMBER() OVER (ORDER BY ForwarderID) AS RowNumber," & vbCrLf & _
                     " ForwarderID = RTRIM(ForwarderID), ForwarderName = RTRIM(ForwarderName), Address = RTRIM(Address), City = RTRIM(City), PostalCode = RTRIM(PostalCode)," & vbCrLf & _
                     " Phone1 = RTRIM(Phone1), Phone2 = RTRIM(Phone2), Fax = RTRIM(Fax), NPWP = RTRIM(NPWP), Case DefaultCls WHEN '0' THEN 'NO' WHEN '1' THEN 'YES' ELSE '' END DefaultCls, PORT" & vbCrLf & _
                     " FROM MS_Forwarder" & vbCrLf & _
                     " ORDER BY ForwarderID"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                .VisibleColumns(0).Width = 35 'RowNumber
                .VisibleColumns(1).Width = 120 'ForwarderID
                .VisibleColumns(2).Width = 250 'ForwarderName
                .VisibleColumns(3).Width = 300 'Address
                .VisibleColumns(4).Width = 130 'City
                .VisibleColumns(4).Width = 100 'PostalCode
                .VisibleColumns(4).Width = 150 'Phone1
                .VisibleColumns(4).Width = 150 'Phone2
                .VisibleColumns(4).Width = 150 'Fax
                .VisibleColumns(4).Width = 150 'NPWP
                .VisibleColumns(4).Width = 100 'DefaultCls
            End With

            sqlConn.Close()
        End Using
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowAllRecord)
    End Sub

    Private Sub TabIndex()
        txtForwarderCode.TabIndex = 1
        txtForwarderName.TabIndex = 2
        txtAddress.TabIndex = 3
        txtCity.TabIndex = 4
        txtPostalCode.TabIndex = 5
        txtPhone1.TabIndex = 6
        txtPhone2.TabIndex = 7
        txtFax.TabIndex = 8
        txtNPWP.TabIndex = 9
        cboDefault.TabIndex = 10
        btnSubmit.TabIndex = 11
        btnDelete.TabIndex = 12
        btnClear.TabIndex = 13
        btnSubMenu.TabIndex = 14
    End Sub

    Private Sub clear()
        txtForwarderCode.Text = ""
        txtForwarderName.Text = ""
        txtAddress.Text = ""
        txtCity.Text = ""
        txtPostalCode.Text = ""
        txtPhone1.Text = ""
        txtPhone2.Text = ""
        txtFax.Text = ""
        txtNPWP.Text = ""
        txtPort.Text = ""
        cboDefault.Text = ""
        grid.FocusedRowIndex = -1
    End Sub

    Private Sub up_SaveData(ByVal pIsUpdate As Boolean, _
                            Optional ByVal pForwarderCode As String = "", _
                            Optional ByVal pForwarderName As String = "", _
                            Optional ByVal pAddress As String = "", _
                            Optional ByVal pCity As String = "", _
                            Optional ByVal pPostalCode As String = "", _
                            Optional ByVal pPhone1 As String = "", _
                            Optional ByVal pPhone2 As String = "", _
                            Optional ByVal pFax As String = "", _
                            Optional ByVal pNPWP As String = "", _
                            Optional ByVal pDefaultCls As String = "", _
                            Optional ByVal pPort As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""

        'If txtMode.Text = "update" Then
        '    flag = False
        'Else
        '    flag = True
        'End If
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT ForwarderID FROM MS_Forwarder WHERE ForwarderID ='" & txtForwarderCode.Text & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    pIsUpdate = False
                Else
                    pIsUpdate = True
                End If
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ForwarderID")

                If pIsUpdate = True Then
                    'INSERT DATA
                    ls_SQL = " INSERT INTO dbo.MS_Forwarder(ForwarderID,ForwarderName,Address,City,PostalCode,Phone1,Phone2,Fax,NPWP,DefaultCls,EntryDate,EntryUser,Port)" & vbCrLf & _
                             " VALUES ('" & Trim(txtForwarderCode.Text) & "'," & _
                             " '" & Trim(txtForwarderName.Text) & "'," & _
                             " '" & Trim(txtAddress.Text) & "'," & _
                             " '" & Trim(txtCity.Text) & "'," & _
                             " '" & Trim(txtPostalCode.Text) & "'," & _
                             " '" & Trim(txtPhone1.Text) & "'," & _
                             " '" & Trim(txtPhone2.Text) & "'," & _
                             " '" & Trim(txtFax.Text) & "'," & _
                             " '" & Trim(txtNPWP.Text) & "'," & _
                             " '" & cboDefault.Value & "'," & _
                             " GETDATE()," & _
                             " '" & Session("UserID").ToString & "'," & _
                             " '" & txtPort.Text.Trim & "')"
                    ls_MsgID = "1001"
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()

                    If cboDefault.Text = "YES" Then
                        ls_SQL = " UPDATE dbo.MS_Forwarder " & vbCrLf & _
                            " SET DefaultCls = '0'," & vbCrLf & _
                            " UpdateDate = GETDATE()," & vbCrLf & _
                            " UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                            " WHERE ForwarderID <> '" & Trim(pForwarderCode) & "'"

                        Dim SqlComm2 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        SqlComm2.ExecuteNonQuery()

                        SqlComm2.Dispose()

                    End If

                    'ElseIf pIsUpdate = False And flag = True Then
                    '    ls_MsgID = "6018"
                    '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    '    lblInfo.Visible = True
                    '    Exit Sub

                ElseIf pIsUpdate = False Then
                    ls_SQL = " UPDATE dbo.MS_Forwarder " & vbCrLf & _
                             " SET ForwarderName = '" & Trim(pForwarderName) & "'," & vbCrLf & _
                             " Address = '" & Trim(pAddress) & "'," & vbCrLf & _
                             " City = '" & Trim(pCity) & "'," & vbCrLf & _
                             " PostalCode = '" & Trim(pPostalCode) & "'," & vbCrLf & _
                             " Phone1 = '" & Trim(pPhone1) & "'," & vbCrLf & _
                             " Phone2 = '" & Trim(pPhone2) & "'," & vbCrLf & _
                             " Fax = '" & Trim(pFax) & "'," & vbCrLf & _
                             " NPWP = '" & Trim(pNPWP) & "'," & vbCrLf & _
                             " DefaultCls = '" & cboDefault.Value & "'," & vbCrLf & _
                             " Port = '" & txtPort.Text.Trim & "'," & vbCrLf & _
                             " UpdateDate = GETDATE()," & vbCrLf & _
                             " UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                             " WHERE ForwarderID = '" & Trim(pForwarderCode) & "'"
                    ls_MsgID = "1002"

                    Dim SqlComm3 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm3.ExecuteNonQuery()

                    SqlComm3.Dispose()

                    If cboDefault.Text = "YES" Then
                        ls_SQL = " UPDATE dbo.MS_Forwarder " & vbCrLf & _
                            " SET DefaultCls = '0'," & vbCrLf & _
                            " UpdateDate = GETDATE()," & vbCrLf & _
                            " UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                            " WHERE ForwarderID <> '" & Trim(pForwarderCode) & "'"

                        Dim SqlComm4 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        SqlComm4.ExecuteNonQuery()

                        SqlComm4.Dispose()
                    End If

                End If
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call up_GridLoad()
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowAllRecord)
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        lblInfo.Visible = True
        clear()
    End Sub

    Private Sub up_DeleteData(ByVal pForwarderCode As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ForwarderID")

                ls_SQL = "DELETE from dbo.MS_Forwarder where ForwarderID = '" & Trim(pForwarderCode) & "'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
        Call up_GridLoad()
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowAllRecord)
        lblInfo.Visible = True
        clear()
    End Sub
#End Region

#Region "FUNCTION"
    Private Function validasiInput() As Boolean
        Try
            Dim sqlstring As String = ""

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                sqlstring = "SELECT ForwarderID FROM dbo.MS_Forwarder WHERE ForwarderID= '" & Trim(txtForwarderCode.Text) & "'"
                Dim sqlDA As New SqlDataAdapter(sqlstring, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    lblInfo.Visible = True
                    lblInfo.Text = "Forwarder Code with Forwarder Code " & txtForwarderCode.Text & " already exists in the database. Data updated"
                    txtForwarderCode.Focus()
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

    Private Function AlreadyUsed(ByVal pForwarderCode) As Boolean
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT DISTINCT ForwarderID FROM dbo.PO_Master_Export WHERE ForwarderID = '" & Trim(pForwarderCode) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    'lblInfo.Text = "ForwarderID ID already used in other screen"
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

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region
End Class