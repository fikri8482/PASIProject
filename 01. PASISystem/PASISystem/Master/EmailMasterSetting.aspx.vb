Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports System
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO
Imports System.Object

Public Class EmailMasterSetting
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_EmailAddress As String
    'Dim ls_AllowUpdate As Boolean = False
    'Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "A18"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            'ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
                '    flag = False
                'Else
                '    flag = True
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Session("M01Url") <> "" Then
                    flag = False
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "GET MAIL"
                        pub_EmailAddress = Request.QueryString("id")
                        tabIndex()
                        'bindData()
                        lblInfo.Text = ""
                        'btnSubMenu.Text = "Back"
                        'txtEmailAddress.ReadOnly = True
                        'txtEmailAddress.BackColor = Color.FromName("#CCCCCC")
                    Else
                        flag = True
                        Ext = Server.MapPath("")
                        'btnClear.Visible = True
                        txtEmailAddress.Focus()
                        tabIndex()
                        'clear()
                    End If
                End If
            End If


            'If ls_AllowUpdate = False Then btnSubmit.Enabled = False

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowPager)

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub AffiliateSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles AffiliateSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"
                    Dim pEmailAddress As String = Split(e.Parameter, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pEmailAddress)
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameter, "|")(2), _
                                     Split(e.Parameter, "|")(3), _
                                     Split(e.Parameter, "|")(4), _
                                     Split(e.Parameter, "|")(5), _
                                     Split(e.Parameter, "|")(6), _
                                     Split(e.Parameter, "|")(7), _
                                     Split(e.Parameter, "|")(8), _
                                     Split(e.Parameter, "|")(9))
                    'bindData()

                Case "delete"
                    Dim pEmailAddress As String = Split(e.Parameter, "|")(1)
                    Dim pUserName As String = Split(e.Parameter, "|")(2)
                    Dim pPassword As String = Split(e.Parameter, "|")(3)
                    Dim pPOP3 As String = Split(e.Parameter, "|")(4)
                    Dim pPort As String = Split(e.Parameter, "|")(5)
                    Dim pAttachmentSave As String = Split(e.Parameter, "|")(6)
                    Dim pAttachmentBackup As String = Split(e.Parameter, "|")(7)
                    Dim pInterval As String = Split(e.Parameter, "|")(8)

                    'If AlreadyUsed(pEmailAddress) = False Then
                    Call DeleteData(Split(e.Parameter, "|")(1), _
                                    Split(e.Parameter, "|")(2), _
                                     Split(e.Parameter, "|")(3), _
                                     Split(e.Parameter, "|")(4), _
                                     Split(e.Parameter, "|")(5), _
                                     Split(e.Parameter, "|")(6), _
                                     Split(e.Parameter, "|")(7), _
                                     Split(e.Parameter, "|")(8))
                    'End If
                    '    clear()
                Case Else
                    'If grid.FindVisibleIndexByKeyValue(txtsearch.Text) >= 0 Then
                    '    grid.FocusedRowIndex = grid.FindVisibleIndexByKeyValue(txtsearch.Text)
                    'Else
                    '    Call bindData()
                    '    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowPager)

                    '    If txtsearch.Text <> "" Then
                    '        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                    '        grid.JSProperties("cpMessage") = lblInfo.Text
                    '    End If
                    'End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'grid.JSProperties("cpMessage") = lblInfo.Text
        End Try
    End Sub

    Protected Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubMenu.Click
            Response.Redirect("~/MainMenu.aspx")

    End Sub

    'Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
    '    clear()
    '    flag = True
    'End Sub
#End Region

#Region "PROCEDURE"
    'Private Sub bindData()
    '    Dim ls_SQL As String = ""

    '    Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '        sqlConn.Open()

    '        ls_SQL = " SELECT RTRIM(AffiliateID)AffiliateID," & vbCrLf & _
    '                    "RTRIM(AffiliateName)AffiliateName," & vbCrLf & _
    '                    "RTRIM(Address)Address," & vbCrLf & _
    '                    "RTRIM(City)City," & vbCrLf & _
    '                    "RTRIM(PostalCode)PostalCode," & vbCrLf & _
    '                    "RTRIM(Phone1)Phone1," & vbCrLf & _
    '                    "RTRIM(Phone2)Phone2," & vbCrLf & _
    '                    "RTRIM(Fax)Fax," & vbCrLf & _
    '                    "RTRIM(NPWP)NPWP," & vbCrLf & _
    '                    "RTRIM(PODeliveryBy)PODeliveryBy" & vbCrLf & _
    '                    "FROM MS_Affiliate where AffiliateID = '" & pub_AffiliateID & "' " & vbCrLf
    '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '        Dim ds As New DataSet
    '        sqlDA.Fill(ds)

    '        If ds.Tables(0).Rows.Count > 0 Then
    '            txtEmailAddress.Text = ds.Tables(0).Rows(0)("EmailAddress")
    '            txtUserName.Text = ds.Tables(0).Rows(0)("Username")
    '            txtPassword.Text = ds.Tables(0).Rows(0)("Password")
    '            txtPop3.Text = ds.Tables(0).Rows(0)("POP3")
    '            txtPort.Text = ds.Tables(0).Rows(0)("Port")
    '            txtAttachmentSave.Text = ds.Tables(0).Rows(0)("AttachmentFolder")
    '            txtAttachmentBackup.Text = ds.Tables(0).Rows(0)("AttachmentBackupFolder")
    '            txtSchedule.Text = ds.Tables(0).Rows(0)("Interval")


    '        End If
    '        sqlConn.Close()
    '    End Using
    'End Sub

    'Private Sub clear()
    '    txtEmailAddress.Text = ""
    '    txtUserName.Text = ""
    '    txtPassword.Text = ""
    '    txtPop3.Text = ""
    '    txtPort.Text = ""
    '    txtAttachmentSave.Text = ""
    '    txtSchedule.Text = ""
    '    txtAttachmentBackup.Text = ""

    '    'txtEmailAddress.ReadOnly = False
    '    'txtEmailAddress.BackColor = Color.FromName("#FFFFFF")
    '    'lblInfo.Text = ""
    'End Sub

    Private Sub tabIndex()
        txtEmailAddress.TabIndex = 1
        txtUserName.TabIndex = 2
        txtPassword.TabIndex = 3
        txtPort.TabIndex = 4
        txtPop3.TabIndex = 5
        txtAttachmentSave.TabIndex = 6
        txtAttachmentBackup.TabIndex = 7
        txtSchedule.TabIndex = 8
        btnSubmit.TabIndex = 9
        btnDelete.TabIndex = 10
        btnSubMenu.TabIndex = 11
    End Sub

    Private Function AlreadyUsed(ByVal pEmailAddress As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT EmailAddress From MS_EmailSetting WHERE EmailAddress = '" & Trim(pEmailAddress) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    'lblInfo.Text = "Affiliate ID already used in other screen"
                    Call clsMsg.DisplayMessage(lblInfo, "5001", clsMessage.MsgType.ErrorMessage)
                    AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                    AffiliateSubmit.JSProperties("cpType") = "error"
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

    Private Sub DeleteData(ByVal pEmailAddress As String, ByVal pUserName As String, ByVal pPassword As String, ByVal pPOP3 As String, ByVal pPort As String, ByVal pAttachmentSave As String, ByVal pAttachmentBackup As String, ByVal pInterval As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("EmailSetting")
                    ls_SQL = " DELETE MS_EmailSetting " & _
                        " WHERE EmailAddress='" & pEmailAddress & "' and " & _
                            "Username='" & pUserName & "' and " & _
                            "Password='" & pPassword & "' and " & _
                            "Port='" & pPort & "' and " & _
                            "POP3='" & pPOP3 & "' and " & _
                            "AttachmentFolder='" & pAttachmentSave & "' and " & _
                            "AttachmentBackupFolder='" & pAttachmentBackup & "' and " & _
                            "Interval='" & pInterval & "'" & vbCrLf

                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            If x > 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
                AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                AffiliateSubmit.JSProperties("cpType") = "info"
                AffiliateSubmit.JSProperties("cpFunction") = "delete"
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Function ValidasiInput(ByVal pEmailAddress As String) As Boolean
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
            '        lblInfo.Text = "Affiliate ID with ID " & txtAffiliateID.Text & " already exists in the database."
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
                         Optional ByVal pEmailAddress As String = "", _
                         Optional ByVal pUserName As String = "", _
                         Optional ByVal pPassword As String = "", _
                         Optional ByVal pPOP3 As String = "", _
                         Optional ByVal pPort As String = "", _
                         Optional ByVal pAttachmentSave As String = "", _
                         Optional ByVal pAttachmentBackup As String = "", _
                         Optional ByVal pInterval As String = "", _
                         Optional ByVal pusernameSMTP As String = "", _
                         Optional ByVal pPasswordSMTP As String = "", _
                         Optional ByVal pSMTP As String = "", _
                         Optional ByVal pPortSMTP As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT * FROM MS_EmailSetting "

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

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("EmailSetting")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_EmailSetting " & _
                                "(EmailAddress, Username, Password, Port, POP3, AttachmentFolder, AttachmentBackupFolder, Interval, usernameSMTP, passwordSMTP, SMTP, PORTSMTP)" & _
                                " VALUES ('" & txtEmailAddress.Text & "','" & txtUserName.Text & "','" & txtPassword.Text & "','" & txtPort.Text & "'," & _
                                "'" & txtPop3.Text & "','" & txtAttachmentSave.Text & "','" & txtAttachmentBackup.Text & "','" & txtSchedule.Text & "')" & vbCrLf
                    ls_MsgID = "1001"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    AffiliateSubmit.JSProperties("cpFunction") = "insert"

                ElseIf pIsNewData = True And flag = False Then
                    ls_MsgID = "6018"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                    AffiliateSubmit.JSProperties("cpType") = "error"
                    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    ls_SQL = "UPDATE MS_Emailsetting SET " & _
                        "EmailAddress='" & txtEmailAddress.Text & "'," & _
                            "Username='" & txtUserName.Text & "'," & _
                            "Password='" & txtPassword.Text & "'," & _
                            "Port='" & txtPort.Text & "'," & _
                            "POP3='" & txtPop3.Text & "'," & _
                            "AttachmentFolder='" & txtAttachmentSave.Text & "'," & _
                            "AttachmentBackupFolder='" & txtAttachmentBackup.Text & "'," & _
                            "Interval='" & txtSchedule.Text & "'," & _
                    "WHERE EmailAddress='" & pEmailAddress & "' and Username ='" & pUserName & "'" & vbCrLf & _
                            " and Password ='" & pPassword & "' and Port='" & pPort & "' " & vbCrLf & _
                            " and POP3 ='" & pPOP3 & "' and AttachmentFolder='" & pAttachmentSave & "' " & vbCrLf & _
                            " and AttachmentBackupFolder='" & pAttachmentBackup & "' and Interval ='" & pInterval & "'  " & vbCrLf
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    AffiliateSubmit.JSProperties("cpFunction") = "update"
                    flag = False
                End If

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
        AffiliateSubmit.JSProperties("cpType") = "info"

    End Sub
#End Region

    Private Sub LoadingPath()
        Dim serverpath As String = ""
        serverpath = Path.Combine(MapPath(""))
        'Path.GetFullPath(FileUpload1.FileName)

        'Server.MapPath(FileUpload1.FileName);
    End Sub

    Private Sub cbBrowse_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbBrowse.Callback
        Call LoadingPath()


    End Sub

    Private Sub cbSetData_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbSetData.Callback
        Dim ls_SQL As String = ""

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " Select  " & _
                                " Username, Password, Port, POP3, AttachmentFolder, AttachmentBackupFolder, Interval, usernameSMTP, passwordSMTP, SMTP, PORTSMTP " & _
                " from MS_EmailSetting " & vbCrLf & _
                " WHERE EmailAddress = '" & txtEmailAddress.Text & "'" & vbCrLf
                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then

                    cbSetData.JSProperties("cpUsername") = ds.Tables(0).Rows(0).Item("Username").ToString()
                    cbSetData.JSProperties("cpPassword") = ds.Tables(0).Rows(0).Item("Password").ToString()
                    cbSetData.JSProperties("cpPort") = ds.Tables(0).Rows(0).Item("Port").ToString()
                    cbSetData.JSProperties("cpPOP3") = ds.Tables(0).Rows(0).Item("POP3").ToString()
                    cbSetData.JSProperties("cpAttachmentFolder") = ds.Tables(0).Rows(0).Item("AttachmentFolder").ToString()
                    cbSetData.JSProperties("cpAttachmentBackupFolder") = ds.Tables(0).Rows(0).Item("AttachmentBackupFolder").ToString()
                    cbSetData.JSProperties("cpInterval") = ds.Tables(0).Rows(0).Item("Interval").ToString()

                Else

                    cbSetData.JSProperties("cpUsername") = ""
                    cbSetData.JSProperties("cpPassword") = ""
                    cbSetData.JSProperties("cpPort") = ""
                    cbSetData.JSProperties("cpPOP3") = ""
                    cbSetData.JSProperties("cpAttachmentFolder") = ""
                    cbSetData.JSProperties("cpAttachmentBackupFolder") = ""
                    cbSetData.JSProperties("cpInterval") = ""
                End If

            End Using

        Catch ex As Exception

        End Try
    End Sub
End Class