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

Public Class EmailSetting
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
    Dim menuID As String = "A19"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim ls_AllowUpdate As String = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Session("MenuDesc") = "EMAIL SETTING"

                pub_EmailAddress = Request.QueryString("id")
                tabIndex()

                lblInfo.Text = ""
                flag = True
                Ext = Server.MapPath("")
                txtEmailAddress.Focus()
                tabIndex()
                cbotype.Items.Clear()
                cbotype.Items.Add("DOMESTIC")
                cbotype.Items.Add("EXPORT")
                cbotype.Text = "DOMESTIC"

                bindData("DOMESTIC")

            End If


            If ls_AllowUpdate = False Then btnSubmit.Enabled = False

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
                                     Split(e.Parameter, "|")(9), _
                                     Split(e.Parameter, "|")(10), _
                                     Split(e.Parameter, "|")(11), _
                                     Split(e.Parameter, "|")(12), _
                                     Split(e.Parameter, "|")(13), _
                                     Split(e.Parameter, "|")(14), _
                                     Split(e.Parameter, "|")(15), _
                                     Split(e.Parameter, "|")(16), _
                                     Split(e.Parameter, "|")(17), _
                                     Split(e.Parameter, "|")(18), _
                                     Split(e.Parameter, "|")(19), _
                                     Split(e.Parameter, "|")(20), _
                                     Split(e.Parameter, "|")(21))
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
                    Dim pusernameSMTP As String = Split(e.Parameter, "|")(9)
                    Dim pPasswordSMTP As String = Split(e.Parameter, "|")(10)
                    Dim pSMTP As String = Split(e.Parameter, "|")(11)
                    Dim pPortSMTP As String = Split(e.Parameter, "|")(12)
                    Dim pTemplate As String = Split(e.Parameter, "|")(13)
                    Dim pResult As String = Split(e.Parameter, "|")(14)
                    Dim pSendExcel As String = Split(e.Parameter, "|")(15)
                    Dim pIntervalPOApproval As String = Split(e.Parameter, "|")(16)
                    Dim pPOApprovalDate As String = Split(e.Parameter, "|")(17)
                    Dim pPORevisionApprovalDate As String = Split(e.Parameter, "|")(18)
                    Dim pIntervalPORevisionApproval As String = Split(e.Parameter, "|")(19)
                    Dim pIntervalKanbanApproval As String = Split(e.Parameter, "|")(20)
                    Dim pKanbanApprovalHour As String = Split(e.Parameter, "|")(21)
                    Dim ptype As String = Split(e.Parameter, "|")(22)

                    'If AlreadyUsed(pEmailAddress) = False Then
                    Call DeleteData(Split(e.Parameter, "|")(1), _
                                    Split(e.Parameter, "|")(2), _
                                     Split(e.Parameter, "|")(3), _
                                     Split(e.Parameter, "|")(4), _
                                     Split(e.Parameter, "|")(5), _
                                     Split(e.Parameter, "|")(6), _
                                     Split(e.Parameter, "|")(7), _
                                     Split(e.Parameter, "|")(8), _
                                     Split(e.Parameter, "|")(9), _
                                     Split(e.Parameter, "|")(10), _
                                     Split(e.Parameter, "|")(11), _
                                     Split(e.Parameter, "|")(12), _
                                     Split(e.Parameter, "|")(13), _
                                     Split(e.Parameter, "|")(14), _
                                     Split(e.Parameter, "|")(15), _
                                     Split(e.Parameter, "|")(16), _
                                     Split(e.Parameter, "|")(17), _
                                     Split(e.Parameter, "|")(18), _
                                     Split(e.Parameter, "|")(19), _
                                     Split(e.Parameter, "|")(20), _
                                     Split(e.Parameter, "|")(21), _
                                     Split(e.Parameter, "|")(22))
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
    Private Sub bindData(ByVal ls_type As String)

        Dim ls_SQL As String = ""


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " Select  " & _
                                " Rtrim(EmailAddress) EmailAddress, Rtrim(Username) Username, Rtrim(Password) Password, " & vbCrLf & _
                                " Rtrim(Port) Port, Rtrim(POP3) POP3, Rtrim(AttachmentFolder)AttachmentFolder, " & vbCrLf & _
                                " Rtrim(AttachmentBackupFolder) AttachmentBackupFolder, Rtrim(Interval) Interval, " & vbCrLf & _
                                " Rtrim(UsernameSMTP) usernameSMTP, Rtrim(PasswordSMTP) passwordSMTP, Rtrim(SMTP) SMTP, Rtrim(PortSMTP )PORTSMTP, " & vbCrLf & _
                                " Rtrim(OriginalTemplateFolder) OriginalTemplateFolder, Rtrim(SaveAsTemplateFolder) SaveAsTemplateFolder, " & vbCrLf & _
                                " IntervalSendExcel, Rtrim(IntervalPOApproval) IntervalPOApproval, Rtrim(POApprovalDate) POApprovalDate, Rtrim(IntervalPORevisionApproval) IntervalPORevisionApproval, Rtrim(PORevisionApprovalDate) PORevisionApprovalDate, Rtrim(IntervalKanbanApproval) IntervalKanbanApproval, Rtrim(KanbanApprovalHour) KanbanApprovalHour "

            If ls_type = "DOMESTIC" Then
                ls_SQL = ls_SQL + " from MS_EmailSetting " & vbCrLf
            Else
                ls_SQL = ls_SQL + " from MS_EmailSetting_export " & vbCrLf
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtEmailAddress.Text = ds.Tables(0).Rows(0)("EmailAddress") & ""
                txtUserName.Text = ds.Tables(0).Rows(0)("Username") & ""
                txtPassword.JSProperties("cp_myPassword") = ds.Tables(0).Rows(0)("Password") & ""
                'txtPassword.Text = ds.Tables(0).Rows(0)("Password") & ""

                txtPort.Text = ds.Tables(0).Rows(0)("Port") & ""
                txtPop3.Text = ds.Tables(0).Rows(0)("Pop3") & ""
                txtAttachmentSave.Text = ds.Tables(0).Rows(0)("AttachmentFolder") & ""
                txtAttachmentBackup.Text = ds.Tables(0).Rows(0)("AttachmentBackupFolder") & ""
                txtSchedule.Text = ds.Tables(0).Rows(0)("Interval") & ""
                txtUserNameSMTP.Text = ds.Tables(0).Rows(0)("usernameSMTP") & ""
                txtPasswordSMTP.JSProperties("cp_myPasswords") = ds.Tables(0).Rows(0)("passwordSMTP") & ""
                'txtPasswordSMTP.Text = ds.Tables(0).Rows(0)("passwordSMTP") & ""

                txtSMTP.Text = ds.Tables(0).Rows(0)("SMTP") & ""
                txtPortSMTP.Text = ds.Tables(0).Rows(0)("PortSMTP") & ""
                txtTemplate.Text = ds.Tables(0).Rows(0)("OriginalTemplateFolder") & ""
                txtResult.Text = ds.Tables(0).Rows(0)("SaveAsTemplateFolder") & ""
                txtSendExcel.Text = ds.Tables(0).Rows(0)("IntervalSendExcel") & ""
                txtPO.Text = ds.Tables(0).Rows(0)("POApprovalDate") & ""
                txtPOInterval.Text = ds.Tables(0).Rows(0)("IntervalPOApproval") & ""
                txtPORevision.Text = ds.Tables(0).Rows(0)("PORevisionApprovalDate") & ""
                txtPORevisionInterval.Text = ds.Tables(0).Rows(0)("IntervalPORevisionApproval") & ""
                txtKanban.Text = ds.Tables(0).Rows(0)("KanbanApprovalHour") & ""
                txtKanbanInterval.Text = ds.Tables(0).Rows(0)("IntervalKanbanApproval") & ""

                cbBind.JSProperties("cptxtEmailAddress") = ds.Tables(0).Rows(0)("EmailAddress") & ""
                cbBind.JSProperties("cptxtUserName") = ds.Tables(0).Rows(0)("Username") & ""
                cbBind.JSProperties("cptxtPassword") = ds.Tables(0).Rows(0)("Password") & ""
                'txtPassword.JSProperties("cp_myPassword") = ds.Tables(0).Rows(0)("Password") & ""
                cbBind.JSProperties("cptxtPort") = ds.Tables(0).Rows(0)("Port") & ""
                cbBind.JSProperties("cptxtPop3") = ds.Tables(0).Rows(0)("Pop3") & ""
                cbBind.JSProperties("cptxtAttachmentSave") = ds.Tables(0).Rows(0)("AttachmentFolder") & ""
                cbBind.JSProperties("cptxtAttachmentBackup") = ds.Tables(0).Rows(0)("AttachmentBackupFolder") & ""
                cbBind.JSProperties("cptxtSchedule") = ds.Tables(0).Rows(0)("Interval") & ""
                cbBind.JSProperties("cptxtUserNameSMTP") = ds.Tables(0).Rows(0)("usernameSMTP") & ""
                cbBind.JSProperties("cptxtPasswordSMTP") = ds.Tables(0).Rows(0)("passwordSMTP") & ""
                'txtPasswordSMTP.JSProperties("cp_myPasswords") = ds.Tables(0).Rows(0)("passwordSMTP") & ""
                cbBind.JSProperties("cptxtSMTP") = ds.Tables(0).Rows(0)("SMTP") & ""
                cbBind.JSProperties("cptxtPortSMTP") = ds.Tables(0).Rows(0)("PortSMTP") & ""
                cbBind.JSProperties("cptxtTemplate") = ds.Tables(0).Rows(0)("OriginalTemplateFolder") & ""
                cbBind.JSProperties("cptxtResult") = ds.Tables(0).Rows(0)("SaveAsTemplateFolder") & ""
                cbBind.JSProperties("cptxtSendExcel") = ds.Tables(0).Rows(0)("IntervalSendExcel") & ""
                cbBind.JSProperties("cptxtPO") = ds.Tables(0).Rows(0)("POApprovalDate") & ""
                cbBind.JSProperties("cptxtPOInterval") = ds.Tables(0).Rows(0)("IntervalPOApproval") & ""
                cbBind.JSProperties("cptxtPORevision") = ds.Tables(0).Rows(0)("PORevisionApprovalDate") & ""
                cbBind.JSProperties("cptxtPORevisionInterval") = ds.Tables(0).Rows(0)("IntervalPORevisionApproval") & ""
                cbBind.JSProperties("cptxtKanban") = ds.Tables(0).Rows(0)("KanbanApprovalHour") & ""
                cbBind.JSProperties("cptxtKanbanInterval") = ds.Tables(0).Rows(0)("IntervalKanbanApproval") & ""
            Else
                txtEmailAddress.Text = ""
                txtUserName.Text = ""
                txtPassword.JSProperties("cp_myPassword") = ""
                txtPort.Text = ""
                txtPop3.Text = ""
                txtAttachmentSave.Text = ""
                txtAttachmentBackup.Text = ""
                txtSchedule.Text = ""
                txtUserNameSMTP.Text = ""
                txtPasswordSMTP.JSProperties("cp_myPasswords") = ""
                txtSMTP.Text = ""
                txtPortSMTP.Text = ""
                txtTemplate.Text = ""
                txtResult.Text = ""
                txtSendExcel.Text = ""
                txtPO.Text = ""
                txtPOInterval.Text = ""
                txtPORevision.Text = ""
                txtPORevisionInterval.Text = ""
                txtKanban.Text = ""
                txtKanbanInterval.Text = ""

                cbBind.JSProperties("cptxtEmailAddress") = ""
                cbBind.JSProperties("cptxtUserName") = ""
                cbBind.JSProperties("cptxtPassword") = ""                
                cbBind.JSProperties("cptxtPort") = ""
                cbBind.JSProperties("cptxtPop3") = ""
                cbBind.JSProperties("cptxtAttachmentSave") = ""
                cbBind.JSProperties("cptxtAttachmentBackup") = ""
                cbBind.JSProperties("cptxtSchedule") = ""
                cbBind.JSProperties("cptxtUserNameSMTP") = ""
                cbBind.JSProperties("cptxtPasswordSMTP") = ""                
                cbBind.JSProperties("cptxtSMTP") = ""
                cbBind.JSProperties("cptxtPortSMTP") = ""
                cbBind.JSProperties("cptxtTemplate") = ""
                cbBind.JSProperties("cptxtResult") = ""
                cbBind.JSProperties("cptxtSendExcel") = ""
                cbBind.JSProperties("cptxtPO") = ""
                cbBind.JSProperties("cptxtPOInterval") = ""
                cbBind.JSProperties("cptxtPORevision") = ""
                cbBind.JSProperties("cptxtPORevisionInterval") = ""
                cbBind.JSProperties("cptxtKanban") = ""
                cbBind.JSProperties("cptxtKanbanInterval") = ""
            End If
            sqlConn.Close()
        End Using
    End Sub

    
    Private Sub tabIndex()
        txtEmailAddress.TabIndex = 1
        txtUserName.TabIndex = 2
        txtPassword.TabIndex = 3
        txtPort.TabIndex = 4
        txtPop3.TabIndex = 5
        txtAttachmentSave.TabIndex = 6
        txtAttachmentBackup.TabIndex = 7
        txtSchedule.TabIndex = 8
        txtUserNameSMTP.TabIndex = 9
        txtPasswordSMTP.TabIndex = 10
        txtSMTP.TabIndex = 11
        txtPortSMTP.TabIndex = 12
        txtTemplate.TabIndex = 13
        txtResult.TabIndex = 14
        txtSendExcel.TabIndex = 15
        txtPOInterval.TabIndex = 16
        txtPO.TabIndex = 17
        txtPORevision.TabIndex = 18
        txtPORevisionInterval.TabIndex = 19
        txtKanban.TabIndex = 20
        txtKanbanInterval.TabIndex = 21
        btnSubmit.TabIndex = 22
        btnDelete.TabIndex = 23
        btnSubMenu.TabIndex = 24
    End Sub

    Private Function AlreadyUsed(ByVal pEmailAddress As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                If cbotype.Text = "DOMESTIC" Then
                    ls_SQL = " SELECT EmailAddress From MS_EmailSetting WHERE EmailAddress = '" & Trim(pEmailAddress) & "'"
                Else
                    ls_SQL = " SELECT EmailAddress From MS_EmailSetting_Export WHERE EmailAddress = '" & Trim(pEmailAddress) & "'"
                End If



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

    Private Sub DeleteData(ByVal pEmailAddress As String, ByVal pUserName As String, ByVal pPassword As String, ByVal pPOP3 As String, ByVal pPort As String, ByVal pAttachmentSave As String, ByVal pAttachmentBackup As String, ByVal pInterval As String, ByVal pusernameSMTP As String, ByVal pPasswordSMTP As String, ByVal pSMTP As String, ByVal pPortSMTP As String, ByVal pTemplate As String, ByVal pResult As String, ByVal pSendExcel As String, ByVal pIntervalPOApproval As String, ByVal pPOApprovalDate As String, ByVal pPORevisionApprovalDate As String, ByVal pIntervalPORevisionApproval As String, ByVal pIntervalKanbanApproval As String, ByVal pKanbanApprovalHour As String, ByVal pType As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("EmailSetting")
                    If Trim(pType) = "DOMESTIC" Then
                        ls_SQL = " DELETE MS_EmailSetting "
                    Else
                        ls_SQL = " DELETE MS_EmailSetting_Export "
                    End If

                    ls_SQL = ls_SQL + " WHERE EmailAddress='" & pEmailAddress & "' and " & _
                            "Username='" & pUserName & "' and " & _
                            "Password='" & pPassword & "' and " & _
                            "Port='" & pPort & "' and " & _
                            "POP3='" & pPOP3 & "' and " & _
                            "AttachmentFolder='" & pAttachmentSave & "' and " & _
                            "AttachmentBackupFolder='" & pAttachmentBackup & "' and " & _
                            "Interval='" & pInterval & "' and " & _
                            "usernameSMTP='" & pusernameSMTP & "' and " & _
                            "passwordSMTP='" & pPasswordSMTP & "' and " & _
                            "SMTP='" & pSMTP & "' and " & _
                            "PortSMTP='" & pPortSMTP & "' and " & _
                            "OriginalTemplateFolder='" & pTemplate & "' and " & _
                            "SaveasTemplateFolder='" & pResult & "' and " & _
                            "IntervalSendExcel='" & pSendExcel & "' and " & _
                            "IntervalPOApproval='" & pIntervalPOApproval & "' and " & _
                            "POApprovalDate='" & pPOApprovalDate & "' and " & _
                            "IntervalPORevisionApproval='" & pIntervalPORevisionApproval & "' and " & _
                            "PORevisionApprovalDate='" & pPORevisionApprovalDate & "' and " & _
                            "IntervalKanbanApproval ='" & pIntervalKanbanApproval & "' and " & _
                            " KanbanApprovalHour='" & pKanbanApprovalHour & "' " & vbCrLf

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
                         Optional ByVal pPortSMTP As String = "", _
                         Optional ByVal pTemplate As String = "", _
                         Optional ByVal pResult As String = "", _
                         Optional ByVal pSendExcel As String = "", _
                         Optional ByVal pPOApprovalDate As String = "", _
                         Optional ByVal pIntervalPOApproval As String = "", _
                         Optional ByVal pPORevisionApprovalDate As String = "", _
                         Optional ByVal pIntervalPORevisionApproval As String = "", _
                         Optional ByVal pKanbanApprovalHour As String = "", _
                         Optional ByVal pIntervalKanbanApproval As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                If cbotype.Text = "DOMESTIC" Then
                    ls_SQL = " SELECT * FROM MS_EmailSetting "
                Else
                    ls_SQL = " SELECT * FROM MS_EmailSetting_export "

                End If


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
                    If cbotype.Text = "DOMESTIC" Then
                        ls_SQL = " INSERT INTO MS_EmailSetting " & vbCrLf
                    Else
                        ls_SQL = " INSERT INTO MS_EmailSetting_export " & vbCrLf
                    End If

                    ls_SQL = ls_SQL + "(EmailAddress, Username, Password, Port, POP3, AttachmentFolder, AttachmentBackupFolder, Interval, usernameSMTP, passwordSMTP, SMTP, PORTSMTP, OriginalTemplateFolder, SaveasTemplateFolder, IntervalSendExcel, IntervalPOApproval, POApprovalDate, IntervalPORevisionApproval, PORevisionApprovalDate, IntervalKanbanApproval, KanbanApprovalHour)" & _
                                      " VALUES ('" & txtEmailAddress.Text & "','" & txtUserName.Text & "','" & txtPassword.Text & "','" & txtPort.Text & "'," & _
                                      "'" & txtPop3.Text & "','" & txtAttachmentSave.Text & "','" & txtAttachmentBackup.Text & "','" & txtSchedule.Text & "'," & _
                                      "'" & txtUserNameSMTP.Text & "','" & txtPasswordSMTP.Text & "','" & txtSMTP.Text & "','" & txtPortSMTP.Text & "'," & _
                                      "'" & txtTemplate.Text & "','" & txtResult.Text & "','" & txtSendExcel.Text & "'," & _
                                      "'" & IIf(txtPOInterval.Text = "", 0, txtPOInterval.Text) & "','" & IIf(txtPO.Text = "", 0, txtPO.Text) & "','" & IIf(txtPORevisionInterval.Text = "", 0, txtPORevisionInterval.Text) & "','" & IIf(txtPORevision.Text = "", 0, txtPORevision.Text) & "','" & IIf(txtKanbanInterval.Text = "", 0, txtKanbanInterval.Text) & "','" & IIf(txtKanban.Text = "", 0, txtKanban.Text) & "')" & vbCrLf
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
                    If cbotype.Text = "DOMESTIC" Then
                        ls_SQL = " UPDATE MS_Emailsetting SET " & vbCrLf
                    Else
                        ls_SQL = " UPDATE MS_Emailsetting SET_export " & vbCrLf
                    End If

                    ls_SQL = ls_SQL + "EmailAddress='" & txtEmailAddress.Text & "'," & _
                                      "Username='" & txtUserName.Text & "'," & _
                                      "Password='" & txtPassword.Text & "'," & _
                                      "Port='" & txtPort.Text & "'," & _
                                      "POP3='" & txtPop3.Text & "'," & _
                                      "AttachmentFolder='" & txtAttachmentSave.Text & "'," & _
                                      "AttachmentBackupFolder='" & txtAttachmentBackup.Text & "'," & _
                                      "Interval='" & txtSchedule.Text & "'," & _
                                      "usernameSMTP='" & txtUserNameSMTP.Text & "'," & _
                                      "passwordSMTP='" & txtPasswordSMTP.Text & "'," & _
                                      "SMTP='" & txtSMTP.Text & "'," & _
                                      "PortSMTP='" & txtPortSMTP.Text & "'," & vbCrLf & _
                                      "OriginalTemplateFolder = '" & txtTemplate.Text & "', " & vbCrLf & _
                                      " SaveasTemplateFolder= '" & txtResult.Text & "', " & vbCrLf & _
                                      " IntervalSendExcel= '" & txtSendExcel.Text & "', " & vbCrLf & _
                                      "IntervalPOApproval='" & txtPOInterval.Text & "'," & _
                                      "POApprovalDate='" & txtPO.Text & "'," & _
                                      "IntervalPORevisionApproval='" & IIf(txtPORevisionInterval.Text, 0, txtPORevisionInterval.Text) & "'," & _
                                      "PORevisionApprovalDate='" & IIf(txtPO.Text = "", 0, txtPO.Text) & "'," & _
                                      "IntervalKanbanApproval='" & IIf(txtKanbanInterval.Text = "", 0, txtKanbanInterval.Text) & "', " & _
                                      "KanbanApprovalHour='" & IIf(txtKanban.Text, 0, txtKanban.Text) & "' " & vbCrLf

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

    Private Sub cbBind_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbBind.Callback
        Call bindData(cbotype.Value)
    End Sub
End Class