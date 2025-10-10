'Update By Robby
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing

Public Class SupplierMasterDetail
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_AffiliateID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "A04"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

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
                        Session("MenuDesc") = "SUPPLIER MASTER ENTRY"
                        pub_AffiliateID = Request.QueryString("id")
                        tabIndex()
                        bindData()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        txtAffiliateID.ReadOnly = True
                        txtAffiliateID.BackColor = Color.FromName("#CCCCCC")                        
                    Else
                        Session("MenuDesc") = "SUPPLIER MASTER ENTRY"
                        tabIndex()
                        clear()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "Back"
                        btnClear.Visible = True
                    End If
                Else
                    flag = True
                    btnClear.Visible = True
                    txtAffiliateID.Focus()
                    tabIndex()
                    clear()
                End If
            End If

            If ls_AllowDelete = False Then btnDelete.Enabled = False
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
                    Dim pAffiliateID As String = Split(e.Parameter, "|")(2)
                    Dim lb_IsUpdate As Boolean = False 'ValidasiInput(pAffiliateID)
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
                                     Split(e.Parameter, "|")(13))
                    'bindData()

                Case "delete"
                    Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
                    If AlreadyUsed(pAffiliateID) = False Then
                        Call DeleteData(pAffiliateID)
                    End If

                Case "load"
                    pub_AffiliateID = txtAffiliateID.Text
                    AffiliateSubmit.JSProperties("cpKeyPress") = "ON"
                    bindData()
                    lblInfo.Text = ""

            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'grid.JSProperties("cpMessage") = lblInfo.Text
        End Try
    End Sub

    Protected Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubMenu.Click
        Session.Remove("1SupplierID")
        Session.Remove("1SupplierName")
        Session.Remove("1Address")
        Session.Remove("1City")
        Session.Remove("1PostalCode")
        Session.Remove("1Phone1")
        Session.Remove("1Phone2")
        Session.Remove("1Fax")
        Session.Remove("1NPWP")
        Session.Remove("1LabelCode")
        Session.Remove("1SupplierType")
        Session.Remove("1Overseas")

        If Session("M01Url") <> "" Then
            'Session.Remove("M01Url")
            Response.Redirect("~/Master/SupplierMaster.aspx")
        Else
            'Session.Remove("M01Url")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        clear()
        flag = True
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " SELECT RTRIM(SupplierID)SupplierID," & vbCrLf & _
                        "RTRIM(SupplierName)SupplierName," & vbCrLf & _
                        "RTRIM(Address)Address," & vbCrLf & _
                        "RTRIM(City)City," & vbCrLf & _
                        "RTRIM(PostalCode)PostalCode," & vbCrLf & _
                        "RTRIM(Phone1)Phone1," & vbCrLf & _
                        "RTRIM(Phone2)Phone2," & vbCrLf & _
                        "RTRIM(Fax)Fax," & vbCrLf & _
                        "RTRIM(NPWP)NPWP," & vbCrLf & _
                        "RTRIM(SupplierCode)Overseas," & vbCrLf & _
                        "RTRIM(LabelCode)LabelCode," & vbCrLf & _
                        "RTRIM(SupplierType)SupplierType" & vbCrLf & _
                        "FROM MS_Supplier where SupplierID = '" & pub_AffiliateID & "' " & vbCrLf
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtAffiliateID.Text = ds.Tables(0).Rows(0)("SupplierID")
                AffiliateSubmit.JSProperties("cpAffiliateID") = txtAffiliateID.Text
                txtAffiliateName.Text = ds.Tables(0).Rows(0)("SupplierName")
                AffiliateSubmit.JSProperties("cpAffiliateName") = txtAffiliateName.Text
                txtAddress.Text = ds.Tables(0).Rows(0)("Address")
                AffiliateSubmit.JSProperties("cpAddress") = txtAddress.Text
                txtCity.Text = ds.Tables(0).Rows(0)("City")
                AffiliateSubmit.JSProperties("cpCity") = txtCity.Text
                txtPostalCode.Text = ds.Tables(0).Rows(0)("PostalCode")
                AffiliateSubmit.JSProperties("cpPostalCode") = txtPostalCode.Text
                txtPhone1.Text = ds.Tables(0).Rows(0)("Phone1")
                AffiliateSubmit.JSProperties("cpPhone1") = txtPhone1.Text
                txtPhone2.Text = ds.Tables(0).Rows(0)("Phone2")
                AffiliateSubmit.JSProperties("cpPhone2") = txtPhone2.Text
                txtFax.Text = ds.Tables(0).Rows(0)("Fax")
                AffiliateSubmit.JSProperties("cpFax") = txtFax.Text
                txtNPWP.Text = ds.Tables(0).Rows(0)("NPWP")
                AffiliateSubmit.JSProperties("cpNPWP") = txtNPWP.Text

                txtPrefix.Text = ds.Tables(0).Rows(0)("LabelCode")
                AffiliateSubmit.JSProperties("cpPrefix") = txtPrefix.Text

                If ds.Tables(0).Rows(0)("SupplierType") = "1" Then
                    rdrPASI.Checked = True
                    AffiliateSubmit.JSProperties("cpSupplierType") = 1
                Else
                    rdrPOTENTIAL.Checked = True
                    AffiliateSubmit.JSProperties("cpSupplierType") = 0
                End If
                'If ds.Tables(0).Rows(0)("Overseas") = "1" Then
                '    rdrOverseas.Checked = True
                '    AffiliateSubmit.JSProperties("cpOverseas") = 1
                'Else
                '    rdrDomestic.Checked = True
                '    AffiliateSubmit.JSProperties("cpOverseas") = 0
                'End If
                txtSupplierCode.Text = ds.Tables(0).Rows(0)("Overseas")
                AffiliateSubmit.JSProperties("cpOverseas") = txtSupplierCode.Text

                Session("1SupplierID") = txtAffiliateID.Text
                Session("1SupplierName") = txtAffiliateName.Text
                Session("1Address") = txtAddress.Text
                Session("1City") = txtCity.Text
                Session("1PostalCode") = txtPostalCode.Text
                Session("1Phone1") = txtPhone1.Text
                Session("1Phone2") = txtPhone2.Text
                Session("1Fax") = txtFax.Text
                Session("1NPWP") = txtNPWP.Text
                Session("1LabelCode") = txtPrefix.Text
                Session("1SupplierType") = ds.Tables(0).Rows(0)("SupplierType")
                Session("1Overseas") = txtSupplierCode.Text

            Else
                AffiliateSubmit.JSProperties("cpKeyPress") = "OFF"

                Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                AffiliateSubmit.JSProperties("cpType") = "info"
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub clear()
        txtAffiliateID.Text = ""
        txtAffiliateName.Text = ""
        txtAddress.Text = ""
        txtCity.Text = ""
        txtPostalCode.Text = ""
        txtPhone1.Text = ""
        txtPhone2.Text = ""
        txtFax.Text = ""
        txtNPWP.Text = ""
        txtPrefix.Text = ""
        rdrPASI.Checked = True
        txtSupplierCode.Text = ""
        txtAffiliateID.ReadOnly = False
        txtAffiliateID.BackColor = Color.FromName("#FFFFFF")
        lblInfo.Text = ""
    End Sub

    Private Sub tabIndex()
        txtAffiliateID.TabIndex = 1
        txtAffiliateName.TabIndex = 2
        rdrPASI.TabIndex = 3
        rdrPOTENTIAL.TabIndex = 4
        'rdrOverseas.TabIndex = 5
        'rdrDomestic.TabIndex = 6
        txtSupplierCode.TabIndex = 5
        txtAddress.TabIndex = 7
        txtCity.TabIndex = 8
        txtPostalCode.TabIndex = 9
        txtPhone1.TabIndex = 10
        txtPhone2.TabIndex = 11
        txtFax.TabIndex = 12
        txtNPWP.TabIndex = 13

        btnSubmit.TabIndex = 14
        btnDelete.TabIndex = 15
        btnClear.TabIndex = 16
        btnSubMenu.TabIndex = 17
    End Sub

    Private Function AlreadyUsed(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID FROM SC_UserSetup WHERE AffiliateID= '" & Trim(pAffiliate) & "'" & vbCrLf & _
                         " Union ALL" & vbCrLf & _
                         " SELECT SupplierID From MS_PartMapping WHERE SupplierID= '" & Trim(pAffiliate) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    'lblInfo.Text = "Affiliate ID already used in other screen"
                    Call clsMsg.DisplayMessage(lblInfo, "5002", clsMessage.MsgType.ErrorMessage)
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

    Private Sub DeleteData(ByVal pAffiliateID As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    ls_SQL = " DELETE MS_Supplier " & vbCrLf & _
                                " WHERE SupplierID = '" & pAffiliateID & "' " & vbCrLf

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
#End Region

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffiliateID As String = "", _
                         Optional ByVal pAffiliateName As String = "", _
                         Optional ByVal pAddress As String = "", _
                         Optional ByVal pCity As String = "", _
                         Optional ByVal pPostalCode As String = "", _
                         Optional ByVal pPhone1 As String = "", _
                         Optional ByVal pPhone2 As String = "", _
                         Optional ByVal pFax As String = "", _
                         Optional ByVal pNPWP As String = "", _
                         Optional ByVal pPODel As String = "", _
                         Optional ByVal pOverseas As String = "", _
                         Optional ByVal pLabelCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_PODel As String = "", ls_xPODel As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT SupplierID FROM MS_Supplier WHERE SupplierID= '" & Trim(pAffiliateID) & "'"

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

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CostCenter")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_Supplier " & _
                                "(SupplierID, SupplierName, Address, City, PostalCode, Phone1, Phone2, Fax, NPWP, SupplierType, SupplierCode, LabelCode, EntryDate, EntryUser)" & _
                                " VALUES ('" & txtAffiliateID.Text & "','" & txtAffiliateName.Text & "','" & txtAddress.Text & "','" & txtCity.Text & "'," & _
                                "'" & txtPostalCode.Text & "','" & txtPhone1.Text & "','" & txtPhone2.Text & "','" & txtFax.Text & "','" & txtNPWP.Text & "','" & pPODel & "', '" & pOverseas & "','" & pLabelCode & "',getdate(),'" & admin & "')" & vbCrLf
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
                    ls_SQL = "UPDATE MS_Supplier SET " & _
                            "SupplierName='" & txtAffiliateName.Text & "'," & _
                            "Address='" & txtAddress.Text & "'," & _
                            "City='" & txtCity.Text & "'," & _
                            "PostalCode='" & txtPostalCode.Text & "'," & _
                            "Phone1='" & txtPhone1.Text & "'," & _
                            "Phone2='" & txtPhone2.Text & "'," & _
                            "Fax='" & txtFax.Text & "'," & _
                            "NPWP='" & txtNPWP.Text & "'," & _
                            "SupplierType='" & pPODel & "'," & _
                            "SupplierCode='" & pOverseas & "'," & _
                            "LabelCode='" & pLabelCode & "'," & _
                            "UpdateDate = getdate()," & _
                            "UpdateUser ='" & admin & "'" & _
                            "WHERE SupplierID='" & pAffiliateID & "'"
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    Dim ls_Remarks As String = ""

                    If Session("1SupplierName").ToString.Trim <> txtAffiliateName.Text.Trim And Session("1SupplierName").ToString.Trim <> "" And txtAffiliateName.Text.Trim <> "" Then
                        ls_Remarks = ls_Remarks + "SupplierName " + Session("1SupplierName").ToString.Trim & "->" & txtAffiliateName.Text.Trim & ""
                        Session("1SupplierName") = txtAffiliateName.Text
                    End If

                    If Session("1SupplierType").ToString.Trim <> pPODel.Trim And Session("1SupplierType").ToString.Trim <> "" And pPODel.Trim <> "" Then
                        If Session("1SupplierType").ToString.Trim = "1" Then ls_xPODel = "PASI Supplier" Else ls_xPODel = "Potential Supplier"
                        If pPODel.Trim = "1" Then ls_PODel = "PASI Supplier" Else ls_PODel = "Potential Supplier"
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "SupplierType " + ls_xPODel & "->" & ls_PODel & ""
                        Session("1SupplierType") = pPODel
                    End If

                    If Session("1Overseas").ToString.Trim <> pOverseas.Trim And Session("1Overseas").ToString.Trim <> "" And pOverseas.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "SupplierCode " + Session("1Overseas").ToString.Trim & "->" & pOverseas.Trim & ""
                        Session("1Overseas") = pOverseas
                    End If

                    If Session("1LabelCode").ToString.Trim <> pLabelCode.Trim And Session("1LabelCode").ToString.Trim <> "" And pLabelCode.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "PrefixLabel " + Session("1LabelCode").ToString.Trim & "->" & pLabelCode.Trim & ""
                        Session("1LabelCode") = pLabelCode
                    End If

                    If Session("1Address").ToString.Trim <> txtAddress.Text.Trim And Session("1Address").ToString.Trim <> "" And txtAddress.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "Address " + Session("1Address").ToString.Trim & "->" & txtAddress.Text.Trim & ""
                        Session("1Address") = txtAddress.Text
                    End If

                    If Session("1City").ToString.Trim <> txtCity.Text.Trim And Session("1City").ToString.Trim <> "" And txtCity.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "City " + Session("1City").ToString.Trim & "->" & txtCity.Text.Trim & ""
                        Session("1City") = txtCity.Text
                    End If

                    If Session("1PostalCode").ToString.Trim <> txtPostalCode.Text.Trim And Session("1PostalCode").ToString.Trim <> "" And txtPostalCode.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "PostalCode " + Session("1PostalCode").ToString.Trim & "->" & txtPostalCode.Text.Trim & ""
                        Session("1PostalCode") = txtPostalCode.Text
                    End If

                    If Session("1Phone1").ToString.Trim <> txtPhone1.Text.Trim And Session("1Phone1").ToString.Trim <> "" And txtPhone1.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "Phone1 " + Session("1Phone1").ToString.Trim & "->" & txtPhone1.Text.Trim & ""
                        Session("1Phone1") = txtPhone1.Text
                    End If

                    If Session("1Phone2").ToString.Trim <> txtPhone2.Text.Trim And Session("1Phone2").ToString.Trim <> "" And txtPhone2.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "Phone2 " + Session("1Phone2").ToString.Trim & "->" & txtPhone2.Text.Trim & ""
                        Session("1Phone2") = txtPhone2.Text
                    End If

                    If Session("1Fax").ToString.Trim <> txtFax.Text.Trim And Session("1Fax").ToString.Trim <> "" And txtFax.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "Fax " + Session("1Fax").ToString.Trim & "->" & txtFax.Text.Trim & ""
                        Session("1Fax") = txtFax.Text
                    End If

                    If Session("1NPWP").ToString.Trim <> txtNPWP.Text.Trim And Session("1NPWP").ToString.Trim <> "" And txtNPWP.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "NPWP " + Session("1NPWP").ToString.Trim & "->" & txtNPWP.Text.Trim & ""
                        Session("1NPWP") = txtNPWP.Text
                    End If

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U', 'Update [" & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                    End If

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

End Class