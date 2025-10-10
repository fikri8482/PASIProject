Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing

Public Class NotificationMaster
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_NotificationID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim menuID As String = "A15"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
                '    flag = False
                'Else
                '    flag = True
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                up_FillCombo()
                If Session("M01Url") <> "" Then
                    flag = False
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "NOTIFICATION MASTER"
                        pub_NotificationID = Request.QueryString("id")
                        tabIndex()
                        'bindData()
                        lblInfo.Text = ""

                    Else
                        Session("MenuDesc") = "NOTIFICATION MASTER"
                        flag = True
                        btnClear.Visible = True
                        cboNotificationCode.Focus()
                        tabIndex()

                    End If
                End If
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
                    Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pub_NotificationID)
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
                                     Split(e.Parameter, "|")(18))
                    'bindData()

                Case Else

                    '    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowPager)
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'grid.JSProperties("cpMessage") = lblInfo.Text
        End Try
    End Sub

    Protected Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    'Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
    '    clear()

    '    up_FillCombo()

    '    flag = True
    'End Sub
#End Region

#Region "PROCEDURE"

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT RTRIM(NotificationCode) NotificationCode, RTRIM(Description) Description from MS_NotificationCls order by NotificationCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboNotificationCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("NotificationCode")
                .Columns(0).Width = 85
                .Columns.Add("Description")
                .Columns(1).Width = 400

                .TextField = "NotificationCode"
                .DataBind()
                '.SelectedIndex = 4
                'txtSupplierCode.Text = clsGlobal.gs_empty
            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine1
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine2
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine3
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine4
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine5
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine6
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine7
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using

        ls_SQL = "SELECT RTRIM(IncludeCls) IncludeCls, RTRIM(Description) Description from MS_IncludeCls order by IncludeCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboLine8
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("IncludeCls")
                .Columns(0).Width = 70
                .Columns.Add("Description")
                .Columns(1).Width = 80

                .TextField = "Description"
                .ValueField = "IncludeCls"
                .DataBind()
                '.SelectedIndex = 4

            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub tabIndex()
        cboNotificationCode.TabIndex = 1
        txtLine1.TabIndex = 2
        cboLine1.TabIndex = 3
        txtLine2.TabIndex = 4
        cboLine2.TabIndex = 5
        txtLine3.TabIndex = 6
        cboLine3.TabIndex = 7
        txtLine4.TabIndex = 8
        cboLine4.TabIndex = 9
        txtLine5.TabIndex = 10
        cboLine5.TabIndex = 11
        txtLine6.TabIndex = 12
        cboLine6.TabIndex = 13
        txtLine7.TabIndex = 14
        cboLine7.TabIndex = 15
        txtLine8.TabIndex = 16
        cboLine8.TabIndex = 17
        btnSubmit.TabIndex = 18
        btnClear.TabIndex = 19
        btnSubMenu.TabIndex = 20
    End Sub

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try

            Dim ls_MsgID As String = ""

            If cboNotificationCode.Text = "" Then
                ls_MsgID = "6010"
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                Return False
            ElseIf txtNotificationCode.Text = "" Then
                ls_MsgID = "6012"
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                Return False

            End If

            Return True

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Function

    'Try
    '    'Dim ls_SQL As String = ""
    '    'Dim ls_MsgID As String = ""

    '    'Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '    '    sqlConn.Open()

    '    '    ls_SQL = "SELECT SupplierID" & vbCrLf & _
    '    '                " FROM MS_EmailSupplier " & _
    '    '                " WHERE SupplierID= '" & Trim(pAffiliate) & "'"

    '    '    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
    '    '    Dim ds As New DataSet
    '    '    sqlDA.Fill(ds)

    '    'If cboSupplierCode.Text Then
    '    '    ls_MsgID = "6018"
    '    '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
    '    '    AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
    '    '    flag = False
    '    '    Return False
    '    'ElseIf cboSupplierCode.Text Then
    '    '    lblInfo.Text = "Supplier ID with ID " & txtSupplierCode.Text & " already exists in the database."
    '    '    Return False
    '    'End If
    '    'Return True
    '    'sqlConn.Close()
    '    'End Using
    'Catch ex As Exception
    '    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    'End Try

    'End Function

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pNotificationCode As String = "", _
                         Optional ByVal pLine1 As String = "", _
                         Optional ByVal pLineCls1 As String = "", _
                         Optional ByVal pLine2 As String = "", _
                         Optional ByVal pLineCls2 As String = "", _
                         Optional ByVal pLine3 As String = "", _
                         Optional ByVal pLineCls3 As String = "", _
                         Optional ByVal pLine4 As String = "", _
                         Optional ByVal pLineCls4 As String = "", _
                         Optional ByVal pLine5 As String = "", _
                         Optional ByVal pLineCls5 As String = "", _
                         Optional ByVal pLine6 As String = "", _
                         Optional ByVal pLineCls6 As String = "", _
                         Optional ByVal pLine7 As String = "", _
                         Optional ByVal pLineCls7 As String = "", _
                         Optional ByVal pLine8 As String = "", _
                         Optional ByVal pLineCls8 As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT NotificationCode FROM MS_Notification WHERE NotificationCode= '" & Trim(pNotificationCode) & "'"

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

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("")

                    Dim sqlComm As New SqlCommand()

                    If pIsNewData = True Then
                        '#INSERT NEW DATA
                        ls_SQL = " INSERT INTO MS_Notification " & _
                                    "(NotificationCode, Line1, Line1Cls, Line2, Line2Cls, Line3, Line3Cls, Line4, Line4Cls, Line5, Line5Cls, Line6, Line6Cls, Line7, Line7Cls, Line8, Line8Cls)" & _
                                    " VALUES ('" & cboNotificationCode.Text & "','" & txtLine1.Text & "','" & cboLine1.Value & "'," & _
                                    "'" & txtLine2.Text & "','" & cboLine2.Value & "','" & txtLine3.Text & "', '" & cboLine3.Value & "'," & _
                                    "'" & txtLine4.Text & "','" & cboLine4.Value & "','" & txtLine5.Text & "', '" & cboLine5.Value & "','" & txtLine6.Text & "','" & cboLine6.Value & "','" & txtLine7.Text & "','" & cboLine7.Value & "','" & txtLine8.Text & "' ,'" & cboLine8.Value & "')" & vbCrLf
                        ls_MsgID = "1001"

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()

                        AffiliateSubmit.JSProperties("cpFunction") = "insert"
                        flag = False
                    ElseIf pIsNewData = True And flag = False Then
                        ls_MsgID = "6018"
                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                        AffiliateSubmit.JSProperties("cpType") = "error"
                        Exit Sub

                    ElseIf pIsNewData = False Then
                        '#UPDATE DATA
                        ls_SQL = "UPDATE MS_Notification SET " & _
                                "NotificationCode='" & cboNotificationCode.Text & "'," & _
                                "Line1='" & txtLine1.Text & "'," & _
                                "Line1Cls='" & cboLine1.Value & "'," & _
                                "Line2='" & txtLine2.Text & "'," & _
                                "Line2Cls='" & cboLine2.Value & "'," & _
                                "Line3='" & txtLine3.Text & "'," & _
                                "Line3Cls='" & cboLine3.Value & "'," & _
                                "Line4='" & txtLine4.Text & "'," & _
                                "Line4Cls='" & cboLine4.Value & "'," & _
                                "Line5='" & txtLine5.Text & "'," & _
                                "Line5Cls='" & cboLine5.Value & "'," & _
                                "Line6='" & txtLine6.Text & "'," & _
                                "Line6Cls='" & cboLine6.Value & "'," & _
                                "Line7='" & txtLine7.Text & "'," & _
                                "Line7Cls = '" & cboLine7.Value & "'," & _
                                "Line8 = '" & txtLine8.Text & "'," & _
                                "Line8Cls ='" & cboLine8.Value & "' " & _
                                "WHERE NotificationCode='" & pNotificationCode & "'"
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


    Private Sub cbSetData_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbSetData.Callback
        Dim ls_SQL As String = ""

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT [Line1], Line1Cls = MI1.Description, [Line2], Line2Cls = MI2.Description,  " & vbCrLf & _
                  "   [Line3], Line3Cls = MI3.Description, [Line4], Line4Cls = MI4.Description, [Line5], Line5Cls = MI5.Description,   " & vbCrLf & _
                  "   [Line6], Line6Cls = MI6.Description, [Line7], Line7Cls = MI7.Description, [Line8], Line8Cls = MI8.Description   " & vbCrLf & _
                  "   FROM MS_Notification MN  " & vbCrLf & _
                  "   left join MS_IncludeCls MI1 on MN.Line1Cls = MI1.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI2 on MN.Line2Cls = MI2.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI3 on MN.Line3Cls = MI3.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI4 on MN.Line4Cls = MI4.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI5 on MN.Line5Cls = MI5.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI6 on MN.Line6Cls = MI6.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI7 on MN.Line7Cls = MI7.IncludeCls   " & vbCrLf & _
                  "  left join MS_IncludeCls MI8 on MN.Line8Cls = MI8.IncludeCls   " & vbCrLf & _
                                  "   WHERE NotificationCode = '" & e.Parameter & "'  " & vbCrLf


                'ls_SQL = "SELECT [Line1], [Line1Cls], [Line2], [Line2Cls], " & vbCrLf & _
                '        " [Line3], [Line3Cls], [Line4], [Line4Cls], [Line5], [Line5Cls], " & vbCrLf & _
                '        " [Line6], [Line6Cls], [Line7], [Line7Cls], [Line8], [Line8Cls] " & vbCrLf & _
                '         " FROM MS_Notification " & vbCrLf & _
                '         " WHERE NotificationCode = '" & e.Parameter & "' " & vbCrLf


                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then

                    cbSetData.JSProperties("cpLine1") = Trim(ds.Tables(0).Rows(0).Item("Line1").ToString())
                    cbSetData.JSProperties("cpLine1Cls") = Trim(ds.Tables(0).Rows(0).Item("Line1Cls").ToString())
                    cbSetData.JSProperties("cpLine2") = Trim(ds.Tables(0).Rows(0).Item("Line2").ToString())
                    cbSetData.JSProperties("cpLine2Cls") = Trim(ds.Tables(0).Rows(0).Item("Line2Cls").ToString())
                    cbSetData.JSProperties("cpLine3") = Trim(ds.Tables(0).Rows(0).Item("Line3").ToString())
                    cbSetData.JSProperties("cpLine3Cls") = Trim(ds.Tables(0).Rows(0).Item("Line3Cls").ToString())
                    cbSetData.JSProperties("cpLine4") = Trim(ds.Tables(0).Rows(0).Item("Line4").ToString())
                    cbSetData.JSProperties("cpLine4Cls") = Trim(ds.Tables(0).Rows(0).Item("Line4Cls").ToString())
                    cbSetData.JSProperties("cpLine5") = Trim(ds.Tables(0).Rows(0).Item("Line5").ToString())
                    cbSetData.JSProperties("cpLine5Cls") = Trim(ds.Tables(0).Rows(0).Item("Line5Cls").ToString())
                    cbSetData.JSProperties("cpLine6") = Trim(ds.Tables(0).Rows(0).Item("Line6").ToString())
                    cbSetData.JSProperties("cpLine6Cls") = Trim(ds.Tables(0).Rows(0).Item("Line6Cls").ToString())
                    cbSetData.JSProperties("cpLine7") = Trim(ds.Tables(0).Rows(0).Item("Line7").ToString())
                    cbSetData.JSProperties("cpLine7Cls") = Trim(ds.Tables(0).Rows(0).Item("Line7Cls").ToString())
                    cbSetData.JSProperties("cpLine8") = Trim(ds.Tables(0).Rows(0).Item("Line8").ToString())
                    cbSetData.JSProperties("cpLine8Cls") = Trim(ds.Tables(0).Rows(0).Item("Line8Cls").ToString())

                Else

                    cbSetData.JSProperties("cpLine1") = ""
                    cbSetData.JSProperties("cpLine1Cls") = ""
                    cbSetData.JSProperties("cpLine2") = ""
                    cbSetData.JSProperties("cpLine2Cls") = ""
                    cbSetData.JSProperties("cpLine3") = ""
                    cbSetData.JSProperties("cpLine3Cls") = ""
                    cbSetData.JSProperties("cpLine4") = ""
                    cbSetData.JSProperties("cpLine4Cls") = ""
                    cbSetData.JSProperties("cpLine5") = ""
                    cbSetData.JSProperties("cpLine5Cls") = ""
                    cbSetData.JSProperties("cpLine6") = ""
                    cbSetData.JSProperties("cpLine6Cls") = ""
                    cbSetData.JSProperties("cpLine7") = ""
                    cbSetData.JSProperties("cpLine7Cls") = ""
                    cbSetData.JSProperties("cpLine8") = ""
                    cbSetData.JSProperties("cpLine8Cls") = ""
                End If

            End Using
        Catch ex As Exception

        End Try
    End Sub


End Class