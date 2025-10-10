Imports System.Drawing
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxEditors

Public Class UserSetup
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim dsUser As New DataSet
    Dim clsDESEncryption As New clsDESEncryption("TOS")
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not Page.IsPostBack) AndAlso (Not Page.IsCallback) Then
            cboAffiliateID.Focus()
            lblErrMsg.Visible = True
            lblErrMsg.Text = ""
            gridUser.FocusedRowIndex = -1
            gridMenu.FocusedRowIndex = -1
            Call TabIndex()
            Call clear()
            Call up_GridLoadUser()
            Call up_GridLoadPrivilege(False)
        End If
        
        txtCCTemp.ForeColor = Color.FromName("White")
        txtUserIDTemp.ForeColor = Color.FromName("White")
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub gridUser_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles gridUser.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        gridUser.JSProperties("cpMessage") = ""
        Try
            Select Case pAction
                Case "load"
                    Call up_GridLoadUser()
                    lblErrMsg.Text = ""

                Case "save"
                    Dim lb_IsUpdate As Boolean = validasiInput()
                    Dim ls_userPr As String = Split(e.Parameters, "|")(8)

                    Call up_SaveData(lb_IsUpdate, _
                                 Split(e.Parameters, "|")(1), _
                                 Split(e.Parameters, "|")(2), _
                                 Split(e.Parameters, "|")(3), _
                                 Split(e.Parameters, "|")(4), _
                                 Split(e.Parameters, "|")(5), _
                                 Split(e.Parameters, "|")(6), _
                                 Split(e.Parameters, "|")(7))
                    If ls_userPr <> "" Then
                        Call up_SaveDataByUP(ls_userPr)
                    End If
                    Call up_GridLoadUser()
                    Call up_GridLoadPrivilege(False)

                Case "delete"
                    Call up_DeleteDataPP(Split(e.Parameters, "|")(2))
                    Call up_DeleteData(Split(e.Parameters, "|")(1), _
                                       Split(e.Parameters, "|")(2))
                    txtUserIDTemp.Text = ""
                    txtFullName.Text = ""
                    txtPasswordUS.Text = ""
                    txtConfPassword.Text = ""
                    Call up_GridLoadUser()
                    Call up_GridLoadPrivilege(False)

                Case "loadPrevilege"
                    Call up_GridLoadPrivilege(True)

            End Select
            txtCCTemp.ForeColor = Color.FromName("#96C8FF")
            txtUserIDTemp.ForeColor = Color.FromName("#96C8FF")
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            gridUser.JSProperties("cpError") = lblErrMsg.Text
            gridUser.FocusedRowIndex = -1
        End Try
    End Sub

    Private Sub gridMenu_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles gridMenu.BatchUpdate
        Dim ls_SQL As String = "", ls_MenuID As String = "", ls_MsgID As String = ""
        Dim iLoop As Long = 0, jLoop As Long = 0
        Dim ls_AllowAccess As String = "", ls_AllowUpdate As String = "", ls_AllowConfirm As String = "", ls_Active As String = "", ls_AllowDelete As String = ""
        Dim ls_UserID As String = ""
        If txtUserIDTemp.Text = "" Then
            ls_UserID = Trim(txtUserId.Text)
        Else
            ls_UserID = Trim(txtUserIDTemp.Text)
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UserMenu")
                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("ZZ010Msg") = lblErrMsg.Text
                    Exit Sub
                End If

                Dim a As Integer
                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1

                    ls_AllowAccess = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())
                    ls_AllowUpdate = (e.UpdateValues(iLoop).NewValues("AllowUpdate").ToString())
                    ls_AllowConfirm = (e.UpdateValues(iLoop).NewValues("AllowConfirm").ToString())
                    ls_AllowDelete = (e.UpdateValues(iLoop).NewValues("AllowDelete").ToString())

                    If ls_AllowAccess = True Then ls_AllowAccess = "1" Else ls_AllowAccess = "0"
                    If ls_AllowUpdate = True Then ls_AllowUpdate = "1" Else ls_AllowUpdate = "0"
                    If ls_AllowConfirm = True Then ls_AllowConfirm = "1" Else ls_AllowConfirm = "0"
                    If ls_AllowDelete = True Then ls_AllowDelete = "1" Else ls_AllowDelete = "0"

                    ls_MenuID = Trim(e.UpdateValues(iLoop).NewValues("MenuID").ToString())

                    ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.SC_UserPrivilege WHERE UserID ='" & ls_UserID & "' AND MenuID='" & ls_MenuID & "' AND UserCls = '0')" & vbCrLf & _
                              " BEGIN " & vbCrLf & _
                            " INSERT INTO dbo.SC_UserPrivilege( AppID, UserCls ,UserID ,MenuID ,AllowAccess ,AllowUpdate, AllowDelete,AllowConfirm) " & vbCrLf & _
                            " VALUES( 'P01', '0' ," & vbCrLf & _
                            " '" & ls_UserID & "' ," & vbCrLf & _
                            " '" & ls_MenuID & "' ," & vbCrLf & _
                            " '" & ls_AllowAccess & "' ," & vbCrLf & _
                            " '" & ls_AllowUpdate & "' ," & vbCrLf & _
                            " '" & ls_AllowDelete & "' ," & vbCrLf & _
                            " '" & ls_AllowConfirm & "')" & vbCrLf & _
                            " END " & vbCrLf & _
                            " ELSE " & vbCrLf & _
                            " BEGIN "
                    ls_SQL = ls_SQL + " UPDATE dbo.SC_UserPrivilege " & vbCrLf & _
                          " SET AllowAccess='" & ls_AllowAccess & "', " & vbCrLf & _
                          " AllowUpdate='" & ls_AllowUpdate & "', " & vbCrLf & _
                          " AllowDelete='" & ls_AllowDelete & "', " & vbCrLf & _
                          " AllowConfirm='" & ls_AllowConfirm & "' " & vbCrLf & _
                          " WHERE AppID='P01' AND UserID='" & ls_UserID & "' AND MenuID='" & ls_MenuID & "' AND UserCls = '0'" & vbCrLf & _
                          " END "
                    ls_MsgID = "1002"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                Next iLoop

                sqlTran.Commit()
                Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
                If lblErrMsg.Text = "[] " Then lblErrMsg.Text = ""
                Session("ZZ010Msg") = lblErrMsg.Text
            End Using

            sqlConn.Close()
        End Using
    End Sub

    Private Sub gridMenu_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles gridMenu.CellEditorInitialize
        If (e.Column.FieldName = "GroupID" Or e.Column.FieldName = "MenuID" Or e.Column.FieldName = "MenuDesc") And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub gridMenu_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles gridMenu.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)

        Try
            Select Case pAction
                Case "load"
                    Call up_GridLoadPrivilege(False)
                    gridMenu.PageIndex = 0
                Case "loadPrevilege"
                    If cboUserGroup.Text <> "" Then
                        Call up_GridLoadPrivilege(True)
                    Else
                        Call up_GridLoadPrivilege(False)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            gridMenu.JSProperties("cpError") = lblErrMsg.Text
            gridMenu.FocusedRowIndex = -1
        End Try
    End Sub

    Private Sub gridUser_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridUser.PageIndexChanged
        Call up_GridLoadUser()
    End Sub

    Private Sub gridMenu_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridMenu.PageIndexChanged
        If cboUserGroup.Text <> "" Then
            up_GridLoadMenuCombo()
        Else
            up_GridLoadMenu()
        End If
    End Sub

    Private Sub ASPxCallback1_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
        Dim ls_SQL As String = ""
        Dim pwd As String = ""
        Dim pAction As String = Split(e.Parameter, "|")(0)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            If txtUserIDTemp.Text <> "" Then
                ls_SQL = " SELECT RTRIM(AffiliateID) AffiliateID,RTRIM(UserID) UserID,RTRIM(FullName)FullName,Password,InvalidLogin,Locked,StatusAdmin,RTRIM(Description)Description FROM dbo.SC_UserSetup " & vbCrLf & _
                    " WHERE	AppID='P01' AND AffiliateID='" & Trim(cboAffiliateID.Text) & "'  " & vbCrLf & _
                    " AND UserID='" & Trim(txtUserId.Text) & "' AND UserCls = '0' " & vbCrLf & _
                    " ORDER BY UserID  "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                dsUser = New DataSet
                sqlDA.Fill(dsUser)
                If dsUser.Tables(0).Rows.Count > 0 Then
                    pwd = dsUser.Tables(0).Rows(0)("Password")
                    pwd = clsDESEncryption.DecryptData(pwd)
                End If
                sqlConn.Close()
            End If
            If pAction = "search" Then

                ls_SQL = " SELECT RTRIM(AffiliateID) AffiliateID,RTRIM(UserID) UserID,RTRIM(FullName)FullName,Password,InvalidLogin,Locked,StatusAdmin,RTRIM(Description)Description FROM dbo.SC_UserSetup " & vbCrLf & _
                        " WHERE	AppID='P01' " & vbCrLf & _
                        " AND UserID='" & Trim(txtsearch.Text) & "' AND UserCls = '0' " & vbCrLf & _
                        " ORDER BY UserID  "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                dsUser = New DataSet
                sqlDA.Fill(dsUser)
                If dsUser.Tables(0).Rows.Count > 0 Then
                    pwd = dsUser.Tables(0).Rows(0)("Password")
                    pwd = clsDESEncryption.DecryptData(pwd)
                End If
                sqlConn.Close()
            End If

        End Using
        e.Result = pwd
    End Sub

    Private Sub ASPxCallback2_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback2.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Select Case pAction
            Case "search"
                Call up_GridLoadSearchUser()
        End Select

    End Sub

    Protected Sub ASPxTextBox_CustomJSProperties(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CustomJSPropertiesEventArgs)
        Dim ls_SQL As String = ""
        Dim pwd As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            If txtUserIDTemp.Text <> "" Then
                ls_SQL = " SELECT RTRIM(AffiliateID) AffiliateID,RTRIM(UserID) UserID,RTRIM(FullName)FullName,Password,InvalidLogin,Locked,StatusAdmin,RTRIM(Description)Description FROM dbo.SC_UserSetup " & vbCrLf & _
                    " WHERE	AppID='P01' AND AffiliateID='" & Trim(cboAffiliateID.Text) & "'  " & vbCrLf & _
                    " AND UserID='" & Trim(txtUserId.Text) & "' AND UserCls = '0' " & vbCrLf & _
                    " ORDER BY UserID  "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                dsUser = New DataSet
                sqlDA.Fill(dsUser)
                If dsUser.Tables(0).Rows.Count > 0 Then
                    pwd = dsUser.Tables(0).Rows(0)("Password")
                    pwd = clsDESEncryption.DecryptData(pwd)
                End If
                sqlConn.Close()
            End If
        End Using
        e.Properties("cp_myPassword") = pwd
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub up_GridLoadUser()
        Dim ls_SQL As String = ""
        Dim pwd As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT RTRIM(AffiliateID) AffiliateID,RTRIM(UserID) UserID,RTRIM(FullName)FullName,Password,InvalidLogin,Locked,StatusAdmin,RTRIM(Description)Description FROM dbo.SC_UserSetup " & vbCrLf & _
                    " WHERE	AppID='P01' AND UserCls = '0'" & vbCrLf & _
                    " ORDER BY UserID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            dsUser = New DataSet
            sqlDA.Fill(dsUser)

            With gridUser
                .DataSource = dsUser.Tables(0)
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoadUserSearch()
        Dim ls_SQL As String = ""
        Dim pwd As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT RTRIM(AffiliateID) AffiliateID,RTRIM(UserID) UserID,RTRIM(FullName)FullName,Password,InvalidLogin,Locked,StatusAdmin,RTRIM(Description)Description FROM dbo.SC_UserSetup " & vbCrLf & _
                    " WHERE	AppID='P01' AND AffiliateID='" & Trim(txtCCTemp.Text) & "' AND UserID='" & Trim(txtsearch.Text) & "' AND UserCls = '0' " & vbCrLf & _
                    " ORDER BY UserID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            dsUser = New DataSet
            sqlDA.Fill(dsUser)

            With gridUser
                .DataSource = dsUser.Tables(0)
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoadSearchUser()
        Dim ls_SQL As String = ""
        Dim pwd As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT RTRIM(AffiliateID) AffiliateID,RTRIM(UserID) UserID,RTRIM(FullName)FullName,Password,InvalidLogin,Locked,StatusAdmin,RTRIM(Description)Description FROM dbo.SC_UserSetup " & vbCrLf & _
                    " WHERE	AppID='P01' AND UserID='" & Trim(txtsearch.Text) & "' AND UserCls = '0' " & vbCrLf & _
                    " ORDER BY UserID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            dsUser = New DataSet
            sqlDA.Fill(dsUser)
            With ASPxCallback2
                If dsUser.Tables(0).Rows.Count > 0 Then
                    .JSProperties("cpCCCode") = Trim(dsUser.Tables(0).Rows(0)("AffiliateID"))
                    .JSProperties("cpUserId") = Trim(dsUser.Tables(0).Rows(0)("UserID"))
                    .JSProperties("cpFullName") = Trim(dsUser.Tables(0).Rows(0)("FullName"))
                    .JSProperties("cpLocked") = dsUser.Tables(0).Rows(0)("Locked")
                    .JSProperties("cpStatusAdmin") = dsUser.Tables(0).Rows(0)("StatusAdmin")
                    .JSProperties("cpDescription") = Trim(dsUser.Tables(0).Rows(0)("Description"))
                End If
            End With
            sqlConn.Close()
        End Using
    End Sub

    Public Sub up_GridLoadPrivilege(ByVal combo As Boolean)
        If combo = True Then
            Call up_GridLoadMenuCombo()
        Else
            Call up_GridLoadMenu()        
        End If
        txtCCTemp.ForeColor = Color.FromName("#96C8FF")
        txtUserIDTemp.ForeColor = Color.FromName("#96C8FF")
    End Sub

    Public Sub up_GridLoadMenu()
        Dim ls_SQL As String = ""
        Dim ls_UserID As String = ""
        If txtUserIDTemp.Text = "" Then
            ls_UserID = Trim(txtUserId.Text)
        Else
            ls_UserID = Trim(txtUserIDTemp.Text)
        End If

        'GridMenuP
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  SELECT GroupID, USM.MenuID, MenuDesc,   " & vbCrLf & _
                  "  ISNULL(AllowAccess,'0') as AllowAccess,  " & vbCrLf & _
                  "  ISNULL(AllowUpdate,'0') as AllowUpdate,  " & vbCrLf & _
                  "  ISNULL(AllowDelete,'0') as AllowDelete,  " & vbCrLf & _
                  "  ISNULL(AllowConfirm,'0') as AllowConfirm   " & vbCrLf & _
                  "  FROM SC_UserMenu USM " & vbCrLf & _
                  "  LEFT JOIN (Select * from SC_UserPrivilege where UserID='" & ls_UserID & "' AND UserCls = '0' ) UP   " & vbCrLf & _
                  "  ON USM.AppID = UP.AppID and USM.MenuID=UP.MenuID    " & vbCrLf & _
                  "  WHERE USM.AppID='P01' and GroupIndex is not null and PASIMenu IN ('0','2') " & vbCrLf & _
                  "  ORDER by USM.MenuID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With gridMenu
                .DataSource = ds.Tables(0)
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Public Sub up_GridLoadMenuCombo()
        Dim ls_SQL As String = ""
        Dim ls_UserID As String = ""
        If cboUserGroup.Text <> "" Then
            ls_UserID = Trim(cboUserGroup.Text)
        End If

        'GridMenuP
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  SELECT GroupID, USM.MenuID, MenuDesc,   " & vbCrLf & _
                  "  ISNULL(AllowAccess,'0') as AllowAccess,  " & vbCrLf & _
                  "  ISNULL(AllowUpdate,'0') as AllowUpdate,  " & vbCrLf & _
                  "  ISNULL(AllowDelete,'0') as AllowDelete,  " & vbCrLf & _
                  "  ISNULL(AllowConfirm,'0') as AllowConfirm   " & vbCrLf & _
                  "  FROM SC_UserMenu USM " & vbCrLf & _
                  "  LEFT JOIN (Select * from SC_UserPrivilege where UserID='" & ls_UserID & "' AND UserCls = '0' ) UP   " & vbCrLf & _
                  "  ON USM.AppID = UP.AppID and USM.MenuID=UP.MenuID    " & vbCrLf & _
                  "  WHERE USM.AppID='P01' and GroupIndex is not null and PASIMenu IN ('0','2') " & vbCrLf & _
                  "  ORDER by USM.MenuID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With gridMenu
                .DataSource = ds.Tables(0)
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Public Sub up_SaveData(ByVal pIsUpdate As Boolean, _
                            Optional ByVal pCCCode As String = "", _
                            Optional ByVal pUserID As String = "", _
                            Optional ByVal pFullName As String = "", _
                            Optional ByVal pPwd As String = "", _
                            Optional ByVal pLocked As String = "", _
                            Optional ByVal pStatusAdmin As String = "", _
                            Optional ByVal pDesc As String = "")
        Dim ls_SQL As String = "", ls_MsgID As String = ""

        Dim a As String
        a = clsDESEncryption.EncryptData(Trim(pPwd))

        If pStatusAdmin = "null" Then pStatusAdmin = "0"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UserSetup")
                If pIsUpdate = True Then
                    'INSERT DATA
                    ls_SQL = " INSERT INTO dbo.SC_UserSetup " & vbCrLf & _
                             " (AppID ,AffiliateID, UserCls ,UserID ,FullName ,Password ,  " & vbCrLf & _
                             " InvalidLogin ,Locked ,StatusAdmin ,Description)" & vbCrLf & _
                             " VALUES ('P01','" & Trim(pCCCode) & "', '0' ,'" & Trim(pUserID) & "','" & Trim(pFullName) & "','" & a & "'," & vbCrLf & _
                             " 0 ,'" & Trim(pLocked) & "','" & Trim(pStatusAdmin) & "', '" & Trim(pDesc) & "')"
                    ls_MsgID = "1001"
                Else
                    ls_SQL = " UPDATE dbo.SC_UserSetup " & vbCrLf & _
                             " SET FullName='" & Trim(pFullName) & "', " & vbCrLf & _
                             " Password='" & Trim(a) & "', " & vbCrLf & _
                             " InvalidLogin= 0 , " & vbCrLf & _
                             " Locked=" & Trim(pLocked) & ", " & vbCrLf & _
                             " StatusAdmin=" & Trim(pStatusAdmin) & ", " & vbCrLf & _
                             " Description='" & Trim(pDesc) & "' " & vbCrLf & _
                             " WHERE UserID='" & Trim(pUserID) & "' AND " & vbCrLf & _
                             " AffiliateID='" & Trim(pCCCode) & "' AND UserCls = '0'"

                    ls_MsgID = "1002"
                End If

                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
        gridUser.JSProperties("cpMessage") = lblErrMsg.Text
        lblErrMsg.Visible = True
        clear()
    End Sub

    Public Sub up_SaveDataByUP(ByVal pbyUserID As String)
        Dim ls_SQL As String = "", ls_MenuID As String = "", ls_MsgID As String = ""
        Dim iLoop As Long = 0, jLoop As Long = 0
        Dim ls_AllowAccess As String = "", ls_AllowUpdate As String = "", ls_AllowConfirm As String = "", ls_AllowDelete As String = ""
        Dim ls_UserID As String = ""
        Dim ds As New DataSet
        If txtUserIDTemp.Text = "" Then
            ls_UserID = Left(Trim(txtUserId.Text) & Space(15), 15)
        Else
            ls_UserID = Left(Trim(txtUserIDTemp.Text) & Space(15), 15)
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  SELECT GroupID, USM.MenuID, MenuDesc,   " & vbCrLf & _
                  "  ISNULL(AllowAccess,'0') as AllowAccess,  " & vbCrLf & _
                  "  ISNULL(AllowUpdate,'0') as AllowUpdate,  " & vbCrLf & _
                  "  ISNULL(AllowDelete,'0') as AllowDelete,  " & vbCrLf & _
                  "  ISNULL(AllowConfirm,'0') as AllowConfirm   " & vbCrLf & _
                  "  FROM SC_UserMenu USM " & vbCrLf & _
                  "  LEFT JOIN (Select * from SC_UserPrivilege where UserID='" & pbyUserID & "' AND UserCls = '0' ) UP   " & vbCrLf & _
                  "  ON USM.AppID = UP.AppID and USM.MenuID=UP.MenuID    " & vbCrLf & _
                  "  WHERE USM.AppID='P01' and GroupIndex is not null and PASIMenu IN ('0','2') " & vbCrLf & _
                  "  ORDER by USM.MenuID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            ds = New DataSet
            sqlDA.Fill(ds)

            sqlConn.Close()
        End Using
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            If ds.Tables(0).Rows.Count > 0 Then
                With ds.Tables(0)
                    For iLoop = 0 To .Rows.Count - 1
                        ls_MenuID = ds.Tables(0).Rows(iLoop)("MenuID")
                        ls_AllowAccess = ds.Tables(0).Rows(iLoop)("AllowAccess")
                        ls_AllowUpdate = ds.Tables(0).Rows(iLoop)("AllowUpdate")
                        ls_AllowConfirm = ds.Tables(0).Rows(iLoop)("AllowConfirm")
                        ls_AllowDelete = ds.Tables(0).Rows(iLoop)("AllowDelete")

                        'TRANSACTION PROCESS
                        ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.SC_UserPrivilege WHERE UserID ='" & ls_UserID & "' AND MenuID='" & ls_MenuID & "' AND UserCls = '0')" & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                        " INSERT INTO dbo.SC_UserPrivilege( AppID, UserCls ,UserID ,MenuID ,AllowAccess ,AllowUpdate, AllowDelete ,AllowConfirm) " & vbCrLf & _
                        " VALUES( 'P01', '0' ," & vbCrLf & _
                        " '" & ls_UserID & "' ," & vbCrLf & _
                        " '" & ls_MenuID & "' ," & vbCrLf & _
                        " '" & ls_AllowAccess & "' ," & vbCrLf & _
                        " '" & ls_AllowUpdate & "' ," & vbCrLf & _
                        " '" & ls_AllowDelete & "' ," & vbCrLf & _
                        " '" & ls_AllowConfirm & "')" & vbCrLf & _
                        " END " & vbCrLf & _
                        " ELSE " & vbCrLf & _
                        " BEGIN "
                        ls_SQL = ls_SQL + " UPDATE dbo.SC_UserPrivilege " & vbCrLf & _
                              " SET AllowAccess='" & ls_AllowAccess & "', " & vbCrLf & _
                              " AllowUpdate='" & ls_AllowUpdate & "', " & vbCrLf & _
                              " AllowDelete='" & ls_AllowDelete & "', " & vbCrLf & _
                              " AllowConfirm='" & ls_AllowConfirm & "' " & vbCrLf & _
                              " WHERE AppID='P01' AND UserID='" & ls_UserID & "' AND MenuID='" & ls_MenuID & "' AND UserCls = '0'" & vbCrLf & _
                              " END "

                        If ls_SQL <> "" Then
                            Dim SqlComm As New SqlCommand(ls_SQL, sqlConn)
                            SqlComm.ExecuteNonQuery()
                            ls_MsgID = "1002"
                            SqlComm.Dispose()

                        End If
                    Next
                End With
            End If
            sqlConn.Close()
        End Using
    End Sub

    Public Sub up_SavePrivilegeMenu()
        Dim ls_SQL As String = "", ls_MenuID As String = "", ls_MsgID As String = ""
        Dim iLoop As Long = 0, jLoop As Long = 0
        Dim ls_AllowAccess As String = "", ls_AllowUpdate As String = "", ls_AllowConfirm As String = "", ls_Active As String = "", ls_AllowDelete As String = ""
        Dim ls_UserID As String = ""
        If txtUserIDTemp.Text = "" Then
            ls_UserID = Left(Trim(txtUserId.Text) & Space(15), 15)
        Else
            ls_UserID = Left(Trim(txtUserIDTemp.Text) & Space(15), 15)
        End If

        'Menu Privileges
        With gridMenu
            For iLoop = 0 To .VisibleRowCount - 1

                Dim chkAllowAccess As ASPxCheckBox = CType(.FindRowCellTemplateControl(iLoop, Nothing, "chkAllowAccess"), ASPxCheckBox)
                Dim chkAllowUpdate As ASPxCheckBox = CType(.FindRowCellTemplateControl(iLoop, Nothing, "chkAllowUpdate"), ASPxCheckBox)
                Dim chkAllowConfirm As ASPxCheckBox = CType(.FindRowCellTemplateControl(iLoop, Nothing, "chkAllowConfirm"), ASPxCheckBox)
                Dim chkAllowDelete As ASPxCheckBox = CType(.FindRowCellTemplateControl(iLoop, Nothing, "chkAllowDelete"), ASPxCheckBox)

                If chkAllowAccess.Checked = True Then ls_AllowAccess = "1" Else ls_AllowAccess = "0"
                If chkAllowUpdate.Checked = True Then ls_AllowUpdate = "1" Else ls_AllowUpdate = "0"
                If chkAllowConfirm.Checked = True Then ls_AllowConfirm = "1" Else ls_AllowConfirm = "0"
                If chkAllowDelete.Checked = True Then ls_AllowDelete = "1" Else ls_AllowDelete = "0"

                ls_MenuID = Trim(.GetRowValues(iLoop, "MenuID").ToString)
                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UserPrivilege")
                        ls_SQL = " IF NOT EXISTS (SELECT * FROM dbo.SC_UserPrivilege WHERE UserID ='" & ls_UserID & "' AND MenuID='" & ls_MenuID & "' AND UserCls = '0')" & vbCrLf & _
                              " BEGIN " & vbCrLf & _
                            " INSERT INTO dbo.SC_UserPrivilege( AppID, UserCls ,UserID ,MenuID ,AllowAccess ,AllowUpdate, AllowDelete ,AllowSpecial) " & vbCrLf & _
                            " VALUES( 'P01', '0' ," & vbCrLf & _
                            " '" & ls_UserID & "' ," & vbCrLf & _
                            " '" & ls_MenuID & "' ," & vbCrLf & _
                            " '" & ls_AllowAccess & "' ," & vbCrLf & _
                            " '" & ls_AllowUpdate & "' ," & vbCrLf & _
                            " '" & ls_AllowDelete & "' ," & vbCrLf & _
                            " '" & ls_AllowConfirm & "')" & vbCrLf & _
                            " END " & vbCrLf & _
                            " ELSE " & vbCrLf & _
                            " BEGIN "
                        ls_SQL = ls_SQL + " UPDATE dbo.SC_UserPrivilege " & vbCrLf & _
                              " SET AllowAccess='" & ls_AllowAccess & "', " & vbCrLf & _
                              " AllowUpdate='" & ls_AllowUpdate & "', " & vbCrLf & _
                              " AllowDelete='" & ls_AllowDelete & "', " & vbCrLf & _
                              " AllowSpecial='" & ls_AllowConfirm & "' " & vbCrLf & _
                              " WHERE AppID='P01' AND UserID='" & ls_UserID & "' AND MenuID='" & ls_MenuID & "' AND UserCls = '0'" & vbCrLf & _
                              " END "

                        ls_MsgID = "1002"

                        Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        SqlComm.ExecuteNonQuery()

                        SqlComm.Dispose()
                        sqlTran.Commit()
                    End Using
                End Using
            Next iLoop
        End With


    End Sub

    Public Sub up_Save()
        txtUserIDTemp.Text = txtUserId.Text
        Dim lb_IsUpdate As Boolean = validasiInput()
        Dim a As String
        a = clsDESEncryption.EncryptData(txtPasswordUS.Text)

        If validation() Then
            up_SaveData(lb_IsUpdate, Trim(cboAffiliateID.Text), Trim(txtUserId.Text), Trim(txtFullName.Text), a, cbAccount.Value, rblAdminStatus.Value, Trim(txtDesc.Text))
            up_SavePrivilegeMenu()            
        End If
        txtUserIDTemp.Text = ""
        Call up_GridLoadPrivilege(False)
        txtCCTemp.ForeColor = Color.FromName("#96C8FF")
        txtUserIDTemp.ForeColor = Color.FromName("#96C8FF")
    End Sub

    Private Sub up_DeleteData(ByVal pCCCode As String, ByVal pUserID As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("FAType")

                ls_SQL = "DELETE dbo.SC_UserSetup WHERE AffiliateID ='" & pCCCode & "' AND UserID='" & pUserID & "' AND UserCls = '0'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsMsg.DisplayMessage(lblErrMsg, "1003", clsMessage.MsgType.InformationMessage)
        gridUser.JSProperties("cpMessage") = lblErrMsg.Text
        lblErrMsg.Visible = True
        clear()
    End Sub

    Private Sub up_DeleteDataPP(ByVal pUserID As String)
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("User")

                ls_SQL = "DELETE FROM dbo.SC_UserPrivilege WHERE UserID='" & pUserID & "' AND UserCls = '0'"
                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsMsg.DisplayMessage(lblErrMsg, "1003", clsMessage.MsgType.InformationMessage)
        gridUser.JSProperties("cpMessage") = lblErrMsg.Text
        lblErrMsg.Visible = True
        clear()
    End Sub

    Private Sub clear()
        cboAffiliateID.SelectedIndex = -1
        txtUserId.Text = ""
        txtFullName.Text = ""
        txtPasswordUS.Text = ""
        txtConfPassword.Text = ""
        cboUserGroup.SelectedIndex = -1
        cbAccount.Checked = False
        txtDesc.Text = ""
    End Sub

    Private Sub TabIndex()
        cboAffiliateID.TabIndex = 1
        txtUserId.TabIndex = 2
        txtFullName.TabIndex = 3
        txtPasswordUS.TabIndex = 4
        txtConfPassword.TabIndex = 5
        cboUserGroup.TabIndex = 6
        rblAdminStatus.TabIndex = 7
        cbAccount.TabIndex = 8
        txtDesc.TabIndex = 9
        gridMenu.TabIndex = 10
        btnSubmit.TabIndex = 11
        btnDelete.TabIndex = 12
        btnClear.TabIndex = 13
        btnSubMenu.TabIndex = 14
    End Sub

    Private Function validation() As Boolean
        If cboAffiliateID.Text = "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7003", clsMessage.MsgType.ErrorMessage)
            gridUser.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        ElseIf txtUserId.Text = "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7004", clsMessage.MsgType.ErrorMessage)
            gridUser.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        ElseIf txtFullName.Text = "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7005", clsMessage.MsgType.ErrorMessage)
            gridUser.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        ElseIf txtPasswordUS.Text = "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7006", clsMessage.MsgType.ErrorMessage)
            gridUser.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        ElseIf txtConfPassword.Text = "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7007", clsMessage.MsgType.ErrorMessage)
            gridUser.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        ElseIf Trim(txtPasswordUS.Text) <> Trim(txtConfPassword.Text) Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7008", clsMessage.MsgType.ErrorMessage)
            gridUser.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        Else
            Return True
        End If

    End Function

    Private Function validasiInput() As Boolean
        validasiInput = True
        Try
            Dim sqlstring As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                sqlstring = "SELECT * FROM dbo.SC_UserSetup WHERE AppID='P01' AND UserID='" & Trim(txtUserId.Text) & "' AND AffiliateID='" & Trim(cboAffiliateID.Text) & "' AND UserCls = '0'"
                Dim sqlDA As New SqlDataAdapter(sqlstring, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    lblErrMsg.Visible = True
                    lblErrMsg.Text = "User ID with ID " & Trim(txtUserId.Text) & " and Affiliate ID " & Trim(cboAffiliateID.Text) & " already exists in system"
                    txtUserId.Focus()
                    Return False
                End If
                Return True
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Function
#End Region
End Class