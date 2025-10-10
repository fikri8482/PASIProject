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

Public Class AffiliateMasterDetail
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_AffiliateID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "A02"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Session("M01Url") <> "" Then
                    flag = False
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "AFFILIATE MASTER ENTRY"
                        pub_AffiliateID = Request.QueryString("id")
                        tabIndex()
                        bindData()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        txtAffiliateID.ReadOnly = True
                        txtAffiliateID.BackColor = Color.FromName("#CCCCCC")
                    Else
                        Session("MenuDesc") = "AFFILIATE MASTER ENTRY"
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
                    Dim lb_IsUpdate As Boolean = True
                    Call SaveData(lb_IsUpdate,
                                     Split(e.Parameter, "|")(2),
                                     Split(e.Parameter, "|")(3),
                                     Split(e.Parameter, "|")(4),
                                     Split(e.Parameter, "|")(5),
                                     Split(e.Parameter, "|")(6),
                                     Split(e.Parameter, "|")(7),
                                     Split(e.Parameter, "|")(8),
                                     Split(e.Parameter, "|")(9),
                                     Split(e.Parameter, "|")(10),
                                     Split(e.Parameter, "|")(11),
                                     Split(e.Parameter, "|")(12),
                                     Split(e.Parameter, "|")(13),
                                     Split(e.Parameter, "|")(14),
                                     Split(e.Parameter, "|")(15),
                                     Split(e.Parameter, "|")(16))
                    'bindData()

                Case "delete"
                    Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
                    If AlreadyUsed(pAffiliateID) = False Then
                        Call DeleteData(pAffiliateID)
                    End If
                    clear()
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
        Session.Remove("1AffiliateID")
        Session.Remove("1AffiliateName")
        Session.Remove("1ConsigneeCode")
        Session.Remove("1ConsigneeAddress")
        Session.Remove("1ConsigneeName")
        Session.Remove("1BuyerCode")
        Session.Remove("1BuyerAddress")
        Session.Remove("1BuyerName")
        Session.Remove("1Address")
        Session.Remove("1City")
        Session.Remove("1PostalCode")
        Session.Remove("1Phone1")
        Session.Remove("1Phone2")
        Session.Remove("1Fax")
        Session.Remove("1NPWP")
        Session.Remove("1KantorPabean")
        Session.Remove("1IzinTPB")
        Session.Remove("1BCPerson")
        Session.Remove("1PODeliveryBy")
        Session.Remove("1OverseasCls")
        Session.Remove("1DestinationPort")
        Session.Remove("1FolderOES")

        If Session("M01Url") <> "" Then
            'Session.Remove("M01Url")
            Response.Redirect("~/Master/AffiliateMaster.aspx")
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

            ls_SQL = " SELECT RTRIM(AffiliateID)AffiliateID," & vbCrLf &
                        "RTRIM(ConsigneeCode)ConsigneeCode," & vbCrLf &
                        "RTRIM(BuyerCode)BuyerCode," & vbCrLf &
                        "RTRIM(AffiliateName)AffiliateName," & vbCrLf &
                        "RTRIM(Address)Address," & vbCrLf &
                        "RTRIM(ConsigneeName)ConsigneeName," & vbCrLf &
                        "RTRIM(ConsigneeAddress)ConsigneeAddress," & vbCrLf &
                        "RTRIM(BuyerName)BuyerName," & vbCrLf &
                        "RTRIM(BuyerAddress)BuyerAddress," & vbCrLf &
                        "RTRIM(City)City," & vbCrLf &
                        "RTRIM(PostalCode)PostalCode," & vbCrLf &
                        "RTRIM(Phone1)Phone1," & vbCrLf &
                        "RTRIM(Phone2)Phone2," & vbCrLf &
                        "RTRIM(Fax)Fax," & vbCrLf &
                        "RTRIM(NPWP)NPWP," & vbCrLf &
                        "ISNULL(RTRIM(KantorPabean),'')KantorPabean," & vbCrLf &
                        "ISNULL(RTRIM(IzinTPB),'')IzinTPB," & vbCrLf &
                        "ISNULL(RTRIM(BCPerson),'')BCPerson," & vbCrLf &
                        "ISNULL(RTRIM(DestinationPort),'')DestinationPort," & vbCrLf &
                        "ISNULL(RTRIM(DestinationPortAir),'')DestinationPortAir," & vbCrLf &
                        "ISNULL(RTRIM(OverseasCls),'0')OverseasCls," & vbCrLf &
                        "ISNULL(RTRIM(PaymentTerm),'')PaymentTerm," & vbCrLf &
                        "ISNULL(RTRIM(POCode),'0')POCode," & vbCrLf &
                        "ISNULL(RTRIM(AffiliateCls),'A')AffiliateCls," & vbCrLf &
                        "ISNULL(TRIM(AffiliateCode),'')AffiliateCode," & vbCrLf &
                        "RTRIM(PODeliveryBy)PODeliveryBy, RTRIM(FolderOES)FolderOES, RTRIM(ISNULL(Att,''))Att" & vbCrLf &
                        "FROM MS_Affiliate where AffiliateID = '" & pub_AffiliateID & "' " & vbCrLf
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtAffiliateID.Text = ds.Tables(0).Rows(0)("AffiliateID")
                AffiliateSubmit.JSProperties("cpAffiliateID") = txtAffiliateID.Text

                txtConsigneeCode.Text = ds.Tables(0).Rows(0)("ConsigneeCode") & ""
                AffiliateSubmit.JSProperties("cpConsigneeCode") = txtConsigneeCode.Text
                txtBuyerCode.Text = ds.Tables(0).Rows(0)("BuyerCode") & ""
                AffiliateSubmit.JSProperties("cpBuyerCode") = txtBuyerCode.Text

                txtAffiliateName.Text = ds.Tables(0).Rows(0)("AffiliateName") & ""
                AffiliateSubmit.JSProperties("cpAffiliateName") = txtAffiliateName.Text
                txtAddress.Text = ds.Tables(0).Rows(0)("Address") & ""
                AffiliateSubmit.JSProperties("cpAddress") = txtAddress.Text

                txtConsigneeAddress.Text = ds.Tables(0).Rows(0)("ConsigneeAddress") & ""
                AffiliateSubmit.JSProperties("cpConsigneeAddress") = txtConsigneeAddress.Text
                txtConsigneeName.Text = ds.Tables(0).Rows(0)("ConsigneeName") & ""
                AffiliateSubmit.JSProperties("cpConsigneeName") = txtConsigneeName.Text

                txtBuyerAddress.Text = ds.Tables(0).Rows(0)("BuyerAddress") & ""
                AffiliateSubmit.JSProperties("cpBuyerAddress") = txtBuyerAddress.Text
                txtBuyerName.Text = ds.Tables(0).Rows(0)("BuyerName") & ""
                AffiliateSubmit.JSProperties("cpBuyerName") = txtBuyerName.Text

                txtCity.Text = ds.Tables(0).Rows(0)("City") & ""
                AffiliateSubmit.JSProperties("cpCity") = txtCity.Text
                txtPostalCode.Text = ds.Tables(0).Rows(0)("PostalCode") & ""
                AffiliateSubmit.JSProperties("cpPostalCode") = txtPostalCode.Text
                txtPhone1.Text = ds.Tables(0).Rows(0)("Phone1") & ""
                AffiliateSubmit.JSProperties("cpPhone1") = txtPhone1.Text
                txtPhone2.Text = ds.Tables(0).Rows(0)("Phone2") & ""
                AffiliateSubmit.JSProperties("cpPhone2") = txtPhone2.Text
                txtFax.Text = ds.Tables(0).Rows(0)("Fax") & ""
                AffiliateSubmit.JSProperties("cpFax") = txtFax.Text
                txtNPWP.Text = ds.Tables(0).Rows(0)("NPWP") & ""
                AffiliateSubmit.JSProperties("cpNPWP") = txtNPWP.Text

                txtKantorPabean.Text = ds.Tables(0).Rows(0)("KantorPabean") & ""
                AffiliateSubmit.JSProperties("cpKantorPabean") = txtKantorPabean.Text
                txtIzinTPB.Text = ds.Tables(0).Rows(0)("IzinTPB") & ""
                AffiliateSubmit.JSProperties("cpIzinTPB") = txtIzinTPB.Text
                txtBCPerson.Text = ds.Tables(0).Rows(0)("BCPerson") & ""
                AffiliateSubmit.JSProperties("cpBCPerson") = txtBCPerson.Text

                If ds.Tables(0).Rows(0)("PODeliveryBy") = "1" Then
                    rdrPASI.Checked = True
                    AffiliateSubmit.JSProperties("cpPODeliveryBy") = 1
                Else
                    rdrSupplier.Checked = True
                    AffiliateSubmit.JSProperties("cpPODeliveryBy") = 0
                End If

                If ds.Tables(0).Rows(0)("OverseasCls") = "1" Then
                    rdrYes.Checked = True
                    AffiliateSubmit.JSProperties("cpOverseasCls") = 1
                Else
                    rdrNo.Checked = True
                    AffiliateSubmit.JSProperties("cpOverseasCls") = 0
                End If

                If ds.Tables(0).Rows(0)("AffiliateCls") = "A" Then
                    rdrAff1.Checked = True
                    AffiliateSubmit.JSProperties("cpAffiliateCls") = 1
                Else
                    rdrAff2.Checked = True
                    AffiliateSubmit.JSProperties("cpAffiliateCls") = 0
                End If

                txtPaymentTerm.Text = ds.Tables(0).Rows(0)("PaymentTerm") & ""
                AffiliateSubmit.JSProperties("cpPaymentTerm") = txtPaymentTerm.Text

                txtPOCode.Text = ds.Tables(0).Rows(0)("POCode") & ""
                AffiliateSubmit.JSProperties("cpPOCode") = txtPOCode.Text

                txtPort.Text = ds.Tables(0).Rows(0)("DestinationPort") & ""
                AffiliateSubmit.JSProperties("cpPort") = txtPort.Text

                txtPortAir.Text = ds.Tables(0).Rows(0)("DestinationPortAir") & ""
                AffiliateSubmit.JSProperties("cpPortAir") = txtPortAir.Text

                txtPath.Text = ds.Tables(0).Rows(0)("FolderOES") & ""
                AffiliateSubmit.JSProperties("cpFolderOES") = txtPath.Text

                txtAtt.Text = ds.Tables(0).Rows(0)("Att") & ""
                AffiliateSubmit.JSProperties("cpAtt") = txtAtt.Text

                txtAffCode.Text = ds.Tables(0).Rows(0)("AffiliateCode") & ""
                AffiliateSubmit.JSProperties("cpAffCode") = txtAffCode.Text

                Session("1AffiliateID") = txtAffiliateID.Text
                Session("1AffiliateName") = txtAffiliateName.Text
                Session("1ConsigneeCode") = txtConsigneeCode.Text
                Session("1ConsigneeAddress") = txtConsigneeAddress.Text
                Session("1ConsigneeName") = txtConsigneeName.Text
                Session("1BuyerCode") = txtBuyerCode.Text
                Session("1BuyerAddress") = txtBuyerAddress.Text
                Session("1BuyerName") = txtBuyerName.Text
                Session("1Address") = txtAddress.Text
                Session("1City") = txtCity.Text
                Session("1PostalCode") = txtPostalCode.Text
                Session("1Phone1") = txtPhone1.Text
                Session("1Phone2") = txtPhone2.Text
                Session("1Fax") = txtFax.Text
                Session("1NPWP") = txtNPWP.Text
                Session("1KantorPabean") = txtKantorPabean.Text
                Session("1IzinTPB") = txtIzinTPB.Text
                Session("1BCPerson") = txtBCPerson.Text
                Session("1PODeliveryBy") = ds.Tables(0).Rows(0)("PODeliveryBy")
                Session("1OverseasCls") = ds.Tables(0).Rows(0)("OverseasCls")
                Session("1DestinationPort") = txtPort.Text
                Session("1DestinationPortAir") = txtPortAir.Text
                Session("1FolderOES") = txtPath.Text

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
        txtConsigneeCode.Text = ""
        txtBuyerCode.Text = ""
        txtAffiliateName.Text = ""
        txtAddress.Text = ""
        txtConsigneeAddress.Text = ""
        txtConsigneeName.Text = ""
        txtBuyerAddress.Text = ""
        txtBuyerName.Text = ""
        txtCity.Text = ""
        txtPostalCode.Text = ""
        txtPhone1.Text = ""
        txtPhone2.Text = ""
        txtFax.Text = ""
        txtNPWP.Text = ""
        txtPath.Text = ""
        txtKantorPabean.Text = ""
        txtIzinTPB.Text = ""
        txtBCPerson.Text = ""
        txtPaymentTerm.Text = ""
        txtPOCode.Text = ""
        rdrPASI.Checked = True
        rdrYes.Checked = True
        rdrAff1.Checked = True
        txtAffiliateID.ReadOnly = False
        txtAffiliateID.BackColor = Color.FromName("#FFFFFF")
        lblInfo.Text = ""
    End Sub

    Private Sub tabIndex()
        txtAffiliateID.TabIndex = 1
        txtAffiliateName.TabIndex = 2
        txtAddress.TabIndex = 3
        txtCity.TabIndex = 4
        txtPostalCode.TabIndex = 5
        txtPhone1.TabIndex = 6
        txtPhone2.TabIndex = 7
        txtFax.TabIndex = 8
        txtNPWP.TabIndex = 9
        txtKantorPabean.TabIndex = 10
        txtIzinTPB.TabIndex = 11
        txtBCPerson.TabIndex = 12
        rdrPASI.TabIndex = 13
        rdrSupplier.TabIndex = 14
        btnSubmit.TabIndex = 15
        btnDelete.TabIndex = 16
        btnClear.TabIndex = 17
        btnSubMenu.TabIndex = 18
    End Sub

    Private Function AlreadyUsed(ByVal pAffiliate As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID FROM SC_UserSetup WHERE AffiliateID= '" & Trim(pAffiliate) & "'" & vbCrLf & _
                         " Union ALL" & vbCrLf & _
                         " SELECT AffiliateID From MS_PartMapping WHERE AffiliateID= '" & Trim(pAffiliate) & "'"

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

    Private Sub DeleteData(ByVal pAffiliateID As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    ls_SQL = " DELETE MS_Affiliate " & vbCrLf & _
                                " WHERE AffiliateID = '" & pAffiliateID & "' " & vbCrLf

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

    Private Sub SaveData(ByVal pIsNewData As Boolean,
                         Optional ByVal pAffiliateID As String = "",
                         Optional ByVal pAffiliateName As String = "",
                         Optional ByVal pAddress As String = "",
                         Optional ByVal pCity As String = "",
                         Optional ByVal pPostalCode As String = "",
                         Optional ByVal pPhone1 As String = "",
                         Optional ByVal pPhone2 As String = "",
                         Optional ByVal pFax As String = "",
                         Optional ByVal pNPWP As String = "",
                         Optional ByVal pPODel As String = "",
                         Optional ByVal pOverseasCls As String = "",
                         Optional ByVal pPort As String = "",
                         Optional ByVal pPortAir As String = "",
                         Optional ByVal pAffCls As String = "",
                         Optional ByVal pAffCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_OverseasCls As String = "", ls_PODel As String = "", ls_xOverseasCls As String = "", ls_xPODel As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName
        Dim ls_Consignee As Boolean = False

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID FROM MS_Affiliate WHERE AffiliateID= '" & Trim(pAffiliateID) & "'"

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

        If Trim(txtConsigneeCode.Text) <> "" Then
            Try
                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    ls_SQL = " SELECT AffiliateID FROM MS_Affiliate WHERE ConsigneeCode = '" & Trim(txtConsigneeCode.Text) & "' and AffiliateID <> '" & Trim(txtAffiliateID.Text) & "'"

                    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                    Dim ds As New DataSet
                    sqlDA.Fill(ds)

                    If ds.Tables(0).Rows.Count > 0 Then
                        ls_Consignee = True
                    Else
                        ls_Consignee = False
                    End If
                    sqlConn.Close()
                End Using
            Catch ex As Exception
                Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            End Try

            If ls_Consignee = True Then
                Call clsMsg.DisplayMessage(lblInfo, "6105", clsMessage.MsgType.ErrorMessage)
                AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                AffiliateSubmit.JSProperties("cpType") = "error"
                Exit Sub
            End If
        End If

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("CostCenter")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_Affiliate " &
                                "(AffiliateID, ConsigneeCode, ConsigneeName, BuyerCode, BuyerName, AffiliateName, Address, ConsigneeAddress, BuyerAddress, City, PostalCode, Phone1, Phone2, Fax, NPWP, KantorPabean, IzinTPB, BCPerson, PODeliveryBy, EntryDate, EntryUser, FolderOES, DestinationPort, DestinationPortAir, OverseasCls, SeqNo, Att, PaymentTerm, POCode, AffiliateCls, AffiliateCode)" &
                                " VALUES ('" & txtAffiliateID.Text & "','" & txtConsigneeCode.Text & "','" & txtConsigneeName.Text & "','" & txtBuyerCode.Text & "','" & txtBuyerName.Text & "','" & txtAffiliateName.Text & "','" & txtAddress.Text & "','" & txtConsigneeAddress.Text & "','" & txtBuyerAddress.Text & "','" & txtCity.Text & "'," &
                                "'" & txtPostalCode.Text & "','" & txtPhone1.Text & "','" & txtPhone2.Text & "','" & txtFax.Text & "','" & txtNPWP.Text & "','" & txtKantorPabean.Text & "','" & txtIzinTPB.Text & "','" & txtBCPerson.Text & "','" & pPODel & "',getdate(),'" & admin & "', '" & txtPath.Text & "','" & txtPort.Text & "','" & txtPortAir.Text & "','" & pOverseasCls & "','0000','" & txtAtt.Text & "','" & txtPaymentTerm.Text & "','" & txtPOCode.Text & "','" & pAffCls & "','" & pAffCode & "')" & vbCrLf
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
                    ls_SQL = "UPDATE MS_Affiliate SET " &
                            "AffiliateName='" & txtAffiliateName.Text & "'," &
                            "ConsigneeCode='" & txtConsigneeCode.Text & "'," &
                            "BuyerCode='" & txtBuyerCode.Text & "'," &
                            "Address='" & txtAddress.Text & "'," &
                            "ConsigneeName='" & txtConsigneeName.Text & "'," &
                            "ConsigneeAddress='" & txtConsigneeAddress.Text & "'," &
                            "BuyerName='" & txtBuyerName.Text & "'," &
                            "BuyerAddress='" & txtBuyerAddress.Text & "'," &
                            "City='" & txtCity.Text & "'," &
                            "PostalCode='" & txtPostalCode.Text & "'," &
                            "Phone1='" & txtPhone1.Text & "'," &
                            "Phone2='" & txtPhone2.Text & "'," &
                            "Fax='" & txtFax.Text & "'," &
                            "NPWP='" & txtNPWP.Text & "'," &
                            "Att='" & txtAtt.Text & "'," &
                            "KantorPabean='" & txtKantorPabean.Text & "'," &
                            "IzinTPB='" & txtIzinTPB.Text & "'," &
                            "BCPerson='" & txtBCPerson.Text & "'," &
                            "PODeliveryBy='" & pPODel & "'," &
                            "FolderOES ='" & txtPath.Text & "'," &
                            "POCode ='" & txtPOCode.Text & "'," &
                            "PaymentTerm ='" & txtPaymentTerm.Text & "'," &
                            "OverseasCls ='" & pOverseasCls & "'," &
                            "DestinationPort ='" & pPort & "'," &
                            "DestinationPortAir ='" & pPortAir & "'," &
                            "AffiliateCls ='" & pAffCls & "'," &
                            "AffiliateCode ='" & pAffCode & "'," &
                            "UpdateDate = getdate()," &
                            "UpdateUser ='" & admin & "'" &
                            "WHERE AffiliateID='" & pAffiliateID & "'"
                    ls_MsgID = "1002"

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    Dim ls_Remarks As String = ""

                    If Session("1AffiliateName").ToString.Trim <> txtAffiliateName.Text.Trim And Session("1AffiliateName").ToString.Trim <> "" And txtAffiliateName.Text.Trim <> "" Then
                        ls_Remarks = ls_Remarks + "AffiliateName " + Session("1AffiliateName").ToString.Trim & "->" & txtAffiliateName.Text.Trim & ""
                        Session("1AffiliateName") = txtAffiliateName.Text
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

                    If Session("1KantorPabean").ToString.Trim <> txtKantorPabean.Text.Trim And Session("1KantorPabean").ToString.Trim <> "" And txtKantorPabean.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "KantorPabean " + Session("1KantorPabean").ToString.Trim & "->" & txtKantorPabean.Text.Trim & ""
                        Session("1KantorPabean") = txtKantorPabean.Text
                    End If

                    If Session("1IzinTPB").ToString.Trim <> txtIzinTPB.Text.Trim And Session("1IzinTPB").ToString.Trim <> "" And txtIzinTPB.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "IzinTPB " + Session("1IzinTPB").ToString.Trim & "->" & txtIzinTPB.Text.Trim & ""
                        Session("1IzinTPB") = txtIzinTPB.Text
                    End If

                    If Session("1BCPerson").ToString.Trim <> txtBCPerson.Text.Trim And Session("1BCPerson").ToString.Trim <> "" And txtBCPerson.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "BCPerson " + Session("1BCPerson").ToString.Trim & "->" & txtBCPerson.Text.Trim & ""
                        Session("1BCPerson") = txtBCPerson.Text
                    End If

                    If Session("1PODeliveryBy").ToString.Trim <> pPODel.Trim And Session("1PODeliveryBy").ToString.Trim <> "" And pPODel.Trim <> "" Then
                        If Session("1PODeliveryBy").ToString.Trim = "1" Then ls_xPODel = "PASI" Else ls_xPODel = "Supplier"
                        If pPODel.Trim = "1" Then ls_PODel = "PASI" Else ls_PODel = "Supplier"
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "PODeliveryBy " + ls_xPODel & "->" & ls_PODel & ""
                        Session("1PODeliveryBy") = pPODel
                    End If

                    If Session("1OverseasCls").ToString.Trim <> pOverseasCls.Trim And Session("1OverseasCls").ToString.Trim <> "" And pOverseasCls.Trim <> "" Then
                        If Session("1OverseasCls").ToString.Trim = "1" Then ls_xOverseasCls = "YES" Else ls_xOverseasCls = "NO"
                        If pOverseasCls.Trim = "1" Then ls_OverseasCls = "YES" Else ls_OverseasCls = "NO"
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "OverseasCls " + ls_xOverseasCls & "->" & ls_OverseasCls & ""
                        Session("1OverseasCls") = pOverseasCls
                    End If

                    If Session("1ConsigneeCode").ToString.Trim <> txtConsigneeCode.Text.Trim And Session("1ConsigneeCode").ToString.Trim <> "" And txtConsigneeCode.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "ConsigneeCode " + Session("1ConsigneeCode").ToString.Trim & "->" & txtConsigneeCode.Text.Trim & ""
                        Session("1ConsigneeCode") = txtConsigneeCode.Text
                    End If

                    If Session("1ConsigneeName").ToString.Trim <> txtConsigneeName.Text.Trim And Session("1ConsigneeName").ToString.Trim <> "" And txtConsigneeName.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "ConsigneeName " + Session("1ConsigneeName").ToString.Trim & "->" & txtConsigneeName.Text.Trim & ""
                        Session("1ConsigneeName") = txtConsigneeName.Text
                    End If

                    If Session("1ConsigneeAddress").ToString.Trim <> txtConsigneeAddress.Text.Trim And Session("1ConsigneeAddress").ToString.Trim <> "" And txtConsigneeAddress.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "ConsigneeAddress " + Session("1ConsigneeAddress").ToString.Trim & "->" & txtConsigneeAddress.Text.Trim & ""
                        Session("1ConsigneeAddress") = txtConsigneeAddress.Text
                    End If

                    If Session("1BuyerCode").ToString.Trim <> txtBuyerCode.Text.Trim And Session("1BuyerCode").ToString.Trim <> "" And txtBuyerCode.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "BuyerCode " + Session("1BuyerCode").ToString.Trim & "->" & txtBuyerCode.Text.Trim & ""
                        Session("1BuyerCode") = txtBuyerCode.Text
                    End If

                    If Session("1BuyerName").ToString.Trim <> txtBuyerName.Text.Trim And Session("1BuyerName").ToString.Trim <> "" And txtBuyerName.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "BuyerName " + Session("1BuyerName").ToString.Trim & "->" & txtBuyerName.Text.Trim & ""
                        Session("1BuyerName") = txtBuyerName.Text
                    End If

                    If Session("1BuyerAddress").ToString.Trim <> txtBuyerAddress.Text.Trim And Session("1BuyerAddress").ToString.Trim <> "" And txtBuyerAddress.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "BuyerAddress " + Session("1BuyerAddress").ToString.Trim & "->" & txtBuyerAddress.Text.Trim & ""
                        Session("1BuyerAddress") = txtBuyerAddress.Text
                    End If

                    If Session("1DestinationPort").ToString.Trim <> pPort.Trim And Session("1DestinationPort").ToString.Trim <> "" And pPort.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "DestinationPort " + Session("1DestinationPort").ToString.Trim & "->" & pPort.Trim & ""
                        Session("1DestinationPort") = pPort
                    End If

                    If Session("1FolderOES").ToString.Trim <> txtPath.Text.Trim And Session("1FolderOES").ToString.Trim <> "" And txtPath.Text.Trim <> "" Then
                        If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                        ls_Remarks = ls_Remarks + "FolderOES " + Session("1FolderOES").ToString.Trim & "->" & txtPath.Text.Trim & ""
                        Session("1FolderOES") = txtPath.Text
                    End If

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, Remarks, RegisterDate, RegisterUserID) " & vbCrLf &
                                 "VALUES ('" & shostname & "','" & menuID & "','U', 'Update [" & ls_Remarks & "]', " & vbCrLf &
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
#End Region
End Class