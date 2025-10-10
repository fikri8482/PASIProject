Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI
Imports DevExpress

Public Class ChangePassword
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim clsDESEncryption As New clsDESEncryption("TOS")
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not Page.IsPostBack) AndAlso (Not Page.IsCallback) Then
            txtCurrentPassword.Focus()
            lblErrMsg.Visible = True
            lblErrMsg.Text = ""
            Call TabIndex()
            Call clear()
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub cbProgress_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbProgress.Callback        
        Try
            Dim lb_IsUpdate As Boolean = validasiInput()
            If lb_IsUpdate = True Then
                Call up_SaveData(lb_IsUpdate, _
                         Trim(txtCurrentPassword.Text), _
                         Trim(txtNewPassword.Text), _
                         Trim(txtConfirmPassword.Text))
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)

            End If
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            cbProgress.JSProperties("cpError") = lblErrMsg.Text
        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Function validasiInput() As Boolean
        If RTrim(clsDESEncryption.EncryptData(txtCurrentPassword.Text)) <> RTrim(RTrim(isExist(Session("UserID")))) Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7002", clsMessage.MsgType.ErrorMessage)
            cbProgress.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        ElseIf Trim(txtNewPassword.Text) <> Trim(txtConfirmPassword.Text) Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7003", clsMessage.MsgType.ErrorMessage)
            cbProgress.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        Else
            Return True
        End If
    End Function

    Private Function isExist(ByVal pUserID As String) As String
        isExist = ""

        Try
            Dim sqlstring As String = ""
            Dim dt As New DataTable

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                sqlstring = "SELECT Password FROM dbo.SC_UserSetup WHERE UserID = '" & pUserID & "' and UserCls = '0'"

                Dim sqlDA As New SqlDataAdapter(sqlstring, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)
                dt = ds.Tables(0)
                If ds.Tables(0).Rows.Count > 0 Then
                    Return dt.Rows(0)("Password")
                    txtCurrentPassword.Focus()
                End If
                Return True
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Function

    Private Sub clear()
        txtCurrentPassword.Text = ""
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
    End Sub

    Private Sub up_SaveData(ByVal pIsUpdate As Boolean, _
                            Optional ByVal pCurrentPass As String = "", _
                            Optional ByVal pNewPass As String = "", _
                            Optional ByVal pConfirmPass As String = "")
        Dim ls_SQL As String = "", ls_MsgID As String = ""


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ChangePassword")
                If pIsUpdate = True Then
                    ls_SQL = "UPDATE dbo.SC_UserSetup SET Password='" & clsDESEncryption.EncryptData(txtNewPassword.Text) & "' WHERE UserID='" & Session("UserID") & "' and UserCls = '0'"
                    ls_MsgID = "1002"
                End If
                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
        cbProgress.JSProperties("cpMessage") = lblErrMsg.Text
        lblErrMsg.Visible = True
        clear()
    End Sub

    Private Sub TabIndex()
        txtCurrentPassword.TabIndex = 1
        txtNewPassword.TabIndex = 2
        txtConfirmPassword.TabIndex = 3
        btnSubmit.TabIndex = 4
        btnClear.TabIndex = 5
        btnSubMenu.TabIndex = 6
    End Sub

#End Region
End Class