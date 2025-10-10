Imports System.Data
Imports System.Data.SqlClient

Public Class Login
    Inherits System.Web.UI.Page

#Region "DECLARATION"    
    Dim clsMsg As New clsMessage
    Dim clsDESEncryption As New clsDESEncryption("TOS")
    Dim clsGlobal As New clsGlobal
#End Region

#Region "Function"
    Public Function LoginValidate(ByVal pUserID As String, ByVal pPassword As String) As Boolean
        Using SqlConn As New SqlConnection(clsGlobal.ConnectionString)
            SqlConn.Open()

            'sql login
            Dim ls_SQL As String = ""
            ls_SQL = "SELECT * FROM SC_UserSetup WHERE UserID = '" & Trim(pUserID) & "' AND Password = '" & Trim(pPassword) & "' and UserCls = '0'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, clsGlobal.ConnectionString)
            Dim ds As New DataSet

            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Session("FullName") = ds.Tables(0).Rows(0)("FullName").ToString.Trim
                Session("AffiliateID") = ds.Tables(0).Rows(0)("AffiliateID").ToString.Trim
                Session("UserCls") = ds.Tables(0).Rows(0)("UserCls").ToString.Trim
                LoginValidate = True
            Else
                LoginValidate = False
            End If
        End Using
    End Function
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Page.Title = "LOGIN - AFFILIATE SYSTEM"

        txtUserID.Focus()

        If Session("UserID") <> "" Then
            Response.Redirect("~/MainMenu.aspx")
            Exit Sub
        End If

        If Session("Msg") <> "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "0000", 2, Session("Msg").ToString)
            Session.Remove("Msg")
        End If
    End Sub

    Private Sub btnLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Try
            lblErrMsg.Text = ""

            If LoginValidate(txtUserID.Text, clsDESEncryption.EncryptData(txtPassword.Text)) = True Then
                Session("UserID") = txtUserID.Text
                Session.Timeout = 600
                If Session("GlobalURL") = "" Then
                    Response.Redirect("~/MainMenu.aspx")
                Else
                    Dim tempUrl As String = Session("GlobalURL")
                    Session.Remove("GlobalURL")
                    Response.Redirect(tempUrl)
                End If
            Else
                Call clsMsg.DisplayMessage(lblErrMsg, "6001", 1)
                txtUserID.Focus()
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, 2, Err.Description.ToString)

        End Try
    End Sub
#End Region
    
End Class