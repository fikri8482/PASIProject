Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing

Public Class ForwarderEmailMaster
    Inherits System.Web.UI.Page

#Region "DECLARATION"

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_AffiliateID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim menuID As String = "A26"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("A26Url") = Request.QueryString("Session")
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                up_FillCombo()
                If Session("A26Url") <> "" Then
                    flag = False
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "FORWARDER EMAIL MASTER"
                        pub_AffiliateID = Request.QueryString("id")
                        tabIndex()
                        lblInfo.Text = ""
                    Else
                        Session("MenuDesc") = "FORWARDER EMAIL MASTER"
                        flag = True
                        btnClear.Visible = True
                        cboForwarderID.Focus()
                        tabIndex()
                        clear()
                    End If
                End If
            End If

            If ls_AllowUpdate = False Then btnSubmit.Enabled = False

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowPager)

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pForwarderID As String = "", _
                         Optional ByVal pForwarderExportTO As String = "", _
                         Optional ByVal pForwarderExportCC As String = "", _
                         Optional ByVal pForwarderRevisionTO As String = "", _
                         Optional ByVal pForwarderRevisionCC As String = "", _
                         Optional ByVal pSupplierDeliveryTO As String = "", _
                         Optional ByVal pSupplierDeliveryCC As String = "", _
                         Optional ByVal pForwarderReceivingTO As String = "", _
                         Optional ByVal pForwarderReceivingCC As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT ForwarderID FROM dbo.MS_EmailForwarder WHERE ForwarderID= '" & Trim(pForwarderID) & "'"

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

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("EmailForwarder")

                    Dim sqlComm As New SqlCommand()

                    If pIsNewData = True Then
                        '#INSERT NEW DATA
                        ls_SQL = " INSERT dbo.MS_EmailForwarder " & _
                                    "(ForwarderID, ForwarderReceivingTO, ForwarderReceivingCC, POExportTO, POExportCC, PORevisionTO, PORevisionCC, SupplierDeliveryTO, SupplierDeliveryCC)" & _
                                    " VALUES ('" & Trim(cboForwarderID.Text) & "','" & Trim(txtForwarderReceivingTO.Text) & "','" & Trim(txtForwarderReceivingCC.Text) & "'," & _
                                    "'" & Trim(txtForwarderPOExportTO.Text) & "','" & Trim(txtForwarderExportCC.Text) & "','" & Trim(txtForwarderRevisionTO.Text) & "','" & Trim(txtForwarderRevisionCC.Text) & "'," & _
                                    "'" & Trim(txtSupplierDeliveryTO.Text) & "','" & Trim(txtSupplierDeliveryCC.Text) & "')" & vbCrLf
                        ls_MsgID = "1001"

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()

                        AffiliateSubmit.JSProperties("cpFunction") = "insert"
                        flag = False
                    ElseIf pIsNewData = False And flag = True Then
                        ls_MsgID = "6018"
                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                        AffiliateSubmit.JSProperties("cpType") = "error"
                        Exit Sub

                    ElseIf pIsNewData = False Then
                        '#UPDATE DATA
                        ls_SQL = "UPDATE dbo.MS_EmailForwarder SET " & _
                                "ForwarderReceivingTO='" & Trim(txtForwarderReceivingTO.Text) & "'," & _
                                "ForwarderReceivingCC='" & Trim(txtForwarderReceivingCC.Text) & "'," & _
                                "POExportTO='" & Trim(txtForwarderPOExportTO.Text) & "'," & _
                                "POExportCC='" & Trim(txtForwarderExportCC.Text) & "'," & _
                                "PORevisionTO='" & Trim(txtForwarderRevisionTO.Text) & "'," & _
                                "PORevisionCC='" & Trim(txtForwarderRevisionCC.Text) & "'," & _
                                "SupplierDeliveryTO='" & Trim(txtSupplierDeliveryTO.Text) & "'," & _
                                "SupplierDeliveryCC='" & Trim(txtSupplierDeliveryCC.Text) & "'" & _
                                "WHERE ForwarderID='" & Trim(pForwarderID) & "'"
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

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
        AffiliateSubmit.JSProperties("cpType") = "info"

    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT RTRIM(ForwarderID) ForwarderID, RTRIM(ForwarderName) ForwarderName from MS_Forwarder ORDER by ForwarderID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboForwarderID
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("ForwarderID")
                .Columns(0).Width = 130
                .Columns.Add("ForwarderName")
                .Columns(1).Width = 320

                .TextField = "ForwarderID"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Function ValidasiInput(ByVal pForwarderID As String) As Boolean
        Dim ls_MsgID As String = ""

        If cboForwarderID.Text = "" Then
            ls_MsgID = "6010"
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            Return False
        ElseIf txtForwarderName.Text = "" Then
            ls_MsgID = "6012"
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            Return False

        End If

        Return True

    End Function

    Private Sub clear()
        cboForwarderID.Text = ""
        txtForwarderName.Text = ""
        txtForwarderPOExportTO.Text = ""
        txtForwarderReceivingCC.Text = ""
        txtForwarderReceivingTO.Text = ""
        txtForwarderRevisionCC.Text = ""
        txtSupplierDeliveryTO.Text = ""
        txtSupplierDeliveryCC.Text = ""
        txtForwarderReceivingTO.Text = ""
        txtForwarderReceivingCC.Text = ""
        lblInfo.Text = ""
    End Sub

    Private Sub tabIndex()
        cboForwarderID.TabIndex = 1
        txtForwarderPOExportTO.TabIndex = 2
        txtForwarderReceivingCC.TabIndex = 3
        txtForwarderReceivingTO.TabIndex = 4
        txtForwarderRevisionCC.TabIndex = 5
        txtSupplierDeliveryTO.TabIndex = 6
        txtSupplierDeliveryCC.TabIndex = 7
        txtForwarderReceivingTO.TabIndex = 8
        txtForwarderReceivingCC.TabIndex = 9
        btnSubmit.TabIndex = 10
        btnClear.TabIndex = 11
        btnSubMenu.TabIndex = 12
    End Sub

#End Region

    Private Sub cbSetData_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbSetData.Callback
        Dim ls_SQL As String = ""

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = "SELECT [ForwarderID],[ForwarderReceivingTO],[ForwarderReceivingCC], " & vbCrLf & _
                         "       [POExportTO], [POExportCC], [PORevisionTO], [PORevisionCC]," & vbCrLf & _
                         "       [SupplierDeliveryTO], [SupplierDeliveryCC]" & vbCrLf & _
                         "  FROM dbo.MS_EmailForwarder  " & vbCrLf & _
                         " WHERE ForwarderID = '" & e.Parameter & "' " & vbCrLf
                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    cbSetData.JSProperties("cpForwarderReceivingTO") = ds.Tables(0).Rows(0).Item("ForwarderReceivingTO").ToString()
                    cbSetData.JSProperties("cpForwarderReceivingCC") = ds.Tables(0).Rows(0).Item("ForwarderReceivingCC").ToString()
                    cbSetData.JSProperties("cpPOExportTO") = ds.Tables(0).Rows(0).Item("POExportTO").ToString()
                    cbSetData.JSProperties("cpPOExportCC") = ds.Tables(0).Rows(0).Item("POExportCC").ToString()
                    cbSetData.JSProperties("cpPORevisionTO") = ds.Tables(0).Rows(0).Item("PORevisionTO").ToString()
                    cbSetData.JSProperties("cpPORevisionCC") = ds.Tables(0).Rows(0).Item("PORevisionCC").ToString()
                    cbSetData.JSProperties("cpSupplierDeliveryTO") = ds.Tables(0).Rows(0).Item("SupplierDeliveryTO").ToString()
                    cbSetData.JSProperties("cpSupplierDeliveryCC") = ds.Tables(0).Rows(0).Item("SupplierDeliveryCC").ToString()
                Else
                    cbSetData.JSProperties("cpForwarderReceivingTO") = ""
                    cbSetData.JSProperties("cpForwarderReceivingCC") = ""
                    cbSetData.JSProperties("cpPOExportTO") = ""
                    cbSetData.JSProperties("cpPOExportCC") = ""
                    cbSetData.JSProperties("cpPORevisionTO") = ""
                    cbSetData.JSProperties("cpPORevisionCC") = ""
                    cbSetData.JSProperties("cpSupplierDeliveryTO") = ""
                    cbSetData.JSProperties("cpSupplierDeliveryCC") = ""
                End If

            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("A26Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub AffiliateSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles AffiliateSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameter, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(Trim(pAffiliateID))
                    Call SaveData(lb_IsUpdate, _
                                     Split(e.Parameter, "|")(2), _
                                     Split(e.Parameter, "|")(3), _
                                     Split(e.Parameter, "|")(4), _
                                     Split(e.Parameter, "|")(5), _
                                     Split(e.Parameter, "|")(6), _
                                     Split(e.Parameter, "|")(7), _
                                     Split(e.Parameter, "|")(8), _
                                     Split(e.Parameter, "|")(9), _
                                     Split(e.Parameter, "|")(10))

                Case Else

                    '    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowPager)
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'grid.JSProperties("cpMessage") = lblInfo.Text
        End Try
    End Sub
End Class