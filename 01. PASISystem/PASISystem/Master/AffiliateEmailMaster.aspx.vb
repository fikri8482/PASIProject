Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing

Public Class AffiliateEmailMaster
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim ls_AllowUpdate As Boolean = False    
    Dim menuID As String = "A11"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Session("MenuDesc") = "AFFILIATE EMAIL MASTER"
                up_FillCombo()
                cbotype.Items.Clear()
                cbotype.Items.Add("DOMESTIC")
                cbotype.Items.Add("EXPORT")
                cbotype.Text = "DOMESTIC"

                If ls_AllowUpdate = False Then
                    btnClear.Enabled = False
                    btnSubmit.Enabled = False
                Else
                    btnClear.Enabled = True
                    btnSubmit.Enabled = True
                End If
            End If
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub AffiliateSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles AffiliateSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"
                    Dim lb_IsUpdate As Boolean = ValidasiInput(Split(e.Parameter, "|")(2))
                    If lb_IsUpdate = True Then
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
                                     Split(e.Parameter, "|")(20))
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        clear()
        up_FillCombo()
    End Sub

    Private Sub cbSetData_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles cbSetData.Callback
        Dim ls_SQL As String = ""

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = "SELECT [AffiliatePOTO],[AffiliatePOCC],[AffiliatePORevisionTO],[AffiliatePORevisionCC],[KanbanTO], " & vbCrLf & _
                         "       [KanbanCC], [SupplierDeliveryTO], [SupplierDeliveryCC], [PASIReceivingTO], [PASIReceivingCC], [AffiliateReceivingTO], " & vbCrLf & _
                         "       [AffiliateReceivingCC], [GoodReceiveTO], [GoodReceiveCC], [InvoiceTO], [InvoiceCC], [SummaryOutstandingTO], [SummaryOutstandingCC] " & vbCrLf

                If cbotype.Text = "DOMESTIC" Then
                    ls_SQL = ls_SQL + "  FROM MS_EmailAffiliate  " & vbCrLf
                Else
                    ls_SQL = ls_SQL + "  FROM MS_EmailAffiliate_EXPORT  " & vbCrLf
                End If

                ls_SQL = ls_SQL + " WHERE AffiliateID = '" & e.Parameter & "' " & vbCrLf
                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    cbSetData.JSProperties("cpAffiliatePOTO") = ds.Tables(0).Rows(0).Item("AffiliatePOTO").ToString()
                    cbSetData.JSProperties("cpAffiliatePOCC") = ds.Tables(0).Rows(0).Item("AffiliatePOCC").ToString()
                    cbSetData.JSProperties("cpAffiliatePORevisionTO") = ds.Tables(0).Rows(0).Item("AffiliatePORevisionTO").ToString()
                    cbSetData.JSProperties("cpAffiliatePORevisionCC") = ds.Tables(0).Rows(0).Item("AffiliatePORevisionCC").ToString()
                    cbSetData.JSProperties("cpKanbanTO") = ds.Tables(0).Rows(0).Item("KanbanTO").ToString()
                    cbSetData.JSProperties("cpKanbanCC") = ds.Tables(0).Rows(0).Item("KanbanCC").ToString()
                    cbSetData.JSProperties("cpSupplierDeliveryTO") = ds.Tables(0).Rows(0).Item("SupplierDeliveryTO").ToString()
                    cbSetData.JSProperties("cpSupplierDeliveryCC") = ds.Tables(0).Rows(0).Item("SupplierDeliveryCC").ToString()
                    cbSetData.JSProperties("cpPASIReceivingTO") = ds.Tables(0).Rows(0).Item("PASIReceivingTO").ToString()
                    cbSetData.JSProperties("cpPASIReceivingCC") = ds.Tables(0).Rows(0).Item("PASIReceivingCC").ToString()
                    cbSetData.JSProperties("cpAffiliateReceivingTO") = ds.Tables(0).Rows(0).Item("AffiliateReceivingTO").ToString()
                    cbSetData.JSProperties("cpAffiliateReceivingCC") = ds.Tables(0).Rows(0).Item("AffiliateReceivingCC").ToString()
                    cbSetData.JSProperties("cpGoodReceiveTO") = ds.Tables(0).Rows(0).Item("GoodReceiveTO").ToString()
                    cbSetData.JSProperties("cpGoodReceiveCC") = ds.Tables(0).Rows(0).Item("GoodReceiveCC").ToString()
                    cbSetData.JSProperties("cpInvoiceTO") = ds.Tables(0).Rows(0).Item("InvoiceTO").ToString()
                    cbSetData.JSProperties("cpInvoiceCC") = ds.Tables(0).Rows(0).Item("InvoiceCC").ToString()
                    cbSetData.JSProperties("cpSummaryOutstandingTO") = ds.Tables(0).Rows(0).Item("SummaryOutstandingTO").ToString()
                    cbSetData.JSProperties("cpSummaryOutstandingCC") = ds.Tables(0).Rows(0).Item("SummaryOutstandingCC").ToString()
                Else
                    cbSetData.JSProperties("cpAffiliatePOTO") = ""
                    cbSetData.JSProperties("cpAffiliatePOCC") = ""
                    cbSetData.JSProperties("cpAffiliatePORevisionTO") = ""
                    cbSetData.JSProperties("cpAffiliatePORevisionCC") = ""
                    cbSetData.JSProperties("cpKanbanTO") = ""
                    cbSetData.JSProperties("cpKanbanCC") = ""
                    cbSetData.JSProperties("cpSupplierDeliveryTO") = ""
                    cbSetData.JSProperties("cpSupplierDeliveryCC") = ""
                    cbSetData.JSProperties("cpPASIReceivingTO") = ""
                    cbSetData.JSProperties("cpPASIReceivingCC") = ""
                    cbSetData.JSProperties("cpAffiliateReceivingTO") = ""
                    cbSetData.JSProperties("cpAffiliateReceivingCC") = ""
                    cbSetData.JSProperties("cpGoodReceiveTO") = ""
                    cbSetData.JSProperties("cpGoodReceiveCC") = ""
                    cbSetData.JSProperties("cpInvoiceTO") = ""
                    cbSetData.JSProperties("cpInvoiceCC") = ""
                    cbSetData.JSProperties("cpSummaryOutstandingTO") = ""
                    cbSetData.JSProperties("cpSummaryOutstandingCC") = ""
                End If

            End Using
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT RTRIM(AffiliateID) AffiliateID, RTRIM(AffiliateName) AffiliateName from MS_Affiliate  order by AffiliateID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 85
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 400

                .TextField = "AffiliateID"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub clear()
        cboAffiliateCode.Text = ""
        txtAffiliateName.Text = ""
        txtEmailToAffiliatePO.Text = ""
        txtEmailCCAffiliatePO.Text = ""
        txtEmailToAffiliatePORevision.Text = ""
        txtEmailCCAffiliatePORevision.Text = ""
        txtEmailToKanban.Text = ""
        txtEmailCCKanban.Text = ""
        txtEmailToSupplierDelivery.Text = ""
        txtEmailCCSupplierDelivery.Text = ""
        txtEmailToPASIReceiving.Text = ""
        txtEmailCCPASIReceiving.Text = ""
        txtEmailToAffiliateReceiving.Text = ""
        txtEmailCCAffiliateReceiving.Text = ""
        txtEmailToGoodReceive.Text = ""
        txtEmailCCGoodReceive.Text = ""
        txtEmailToInvoice.Text = ""
        txtEmailCCInvoice.Text = ""
        txtEmailToSummaryOutstanding.Text = ""
        txtEmailCCSummaryOutstanding.Text = ""
        txtEmailToAffiliatePO.ReadOnly = False
        txtEmailToAffiliatePO.BackColor = Color.FromName("#FFFFFF")
        lblInfo.Text = ""
    End Sub

    Private Sub tabIndex()
        cboAffiliateCode.TabIndex = 1
        txtEmailToAffiliatePO.TabIndex = 2
        txtEmailCCAffiliatePO.TabIndex = 3
        txtEmailToAffiliatePORevision.TabIndex = 4
        txtEmailCCAffiliatePORevision.TabIndex = 5
        txtEmailToKanban.TabIndex = 6
        txtEmailCCKanban.TabIndex = 7
        txtEmailToSupplierDelivery.TabIndex = 8
        txtEmailCCSupplierDelivery.TabIndex = 9
        txtEmailToPASIReceiving.TabIndex = 10
        txtEmailCCPASIReceiving.TabIndex = 11
        txtEmailToAffiliateReceiving.TabIndex = 12
        txtEmailCCAffiliateReceiving.TabIndex = 13
        txtEmailToGoodReceive.TabIndex = 14
        txtEmailCCGoodReceive.TabIndex = 15
        txtEmailToInvoice.TabIndex = 16
        txtEmailCCInvoice.TabIndex = 17
        txtEmailToSummaryOutstanding.TabIndex = 18
        txtEmailCCSummaryOutstanding.TabIndex = 19
        btnSubmit.TabIndex = 20
        btnClear.TabIndex = 21
        btnSubMenu.TabIndex = 22
    End Sub

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Dim ls_MsgID As String = ""

        If cboAffiliateCode.Text = "" Then
            ls_MsgID = "6010"
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            Return False
        ElseIf txtAffiliateName.Text = "" Then
            ls_MsgID = "6012"
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
            AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
            Return False
        End If

        Return True

    End Function

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pAffiliateID As String = "", _
                         Optional ByVal pAffiliatePOTO As String = "", _
                         Optional ByVal pAffiliatePOCC As String = "", _
                         Optional ByVal pAffiliatePORevisionTO As String = "", _
                         Optional ByVal pAffiliatePORevisionCC As String = "", _
                         Optional ByVal pKanbanTO As String = "", _
                         Optional ByVal pKanbanCC As String = "", _
                         Optional ByVal pSupplierDeliveryTO As String = "", _
                         Optional ByVal pSupplierDeliveryCC As String = "", _
                         Optional ByVal pPASIReceivingTO As String = "", _
                         Optional ByVal pPASIReceivingCC As String = "", _
                         Optional ByVal pAffiliateReceivingTO As String = "", _
                         Optional ByVal pAffiliateReceivingCC As String = "", _
                         Optional ByVal pGoodReceiveTO As String = "", _
                         Optional ByVal pGoodReceiveCC As String = "", _
                         Optional ByVal pInvoiceTO As String = "", _
                         Optional ByVal pInvoiceCC As String = "", _
                         Optional ByVal pSummaryOutstandingTO As String = "", _
                         Optional ByVal pSummaryOutstandingCC As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                If cbotype.Text = "DOMESTIC" Then
                    ls_SQL = " SELECT AffiliateID FROM MS_EmailAffiliate WHERE AffiliateID= '" & Trim(pAffiliateID) & "'"
                Else
                    ls_SQL = " SELECT AffiliateID FROM MS_EmailAffiliate_Export WHERE AffiliateID= '" & Trim(pAffiliateID) & "'"
                End If

                'ls_SQL = " SELECT AffiliateID FROM MS_EmailAffiliate WHERE AffiliateID= '" & Trim(pAffiliateID) & "'"

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

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("EmailAffiliate")

                    Dim sqlComm As New SqlCommand()

                    If pIsNewData = True Then
                        '#INSERT NEW DATA
                        If cbotype.Text = "DOMESTIC" Then
                            ls_SQL = " INSERT INTO MS_EmailAffiliate "
                        Else
                            ls_SQL = " INSERT INTO MS_EmailAffiliate_EXPORT "
                        End If

                        ls_SQL = ls_SQL + "(AffiliateID, AffiliatePOTO, AffiliatePOCC, AffiliatePORevisionTO, AffiliatePORevisionCC, KanbanTO, KanbanCC, SupplierDeliveryTO, SupplierDeliveryCC, PASIReceivingTo, PASIReceivingCC, AffiliateReceivingTo, AffiliateReceivingCC, GoodReceiveTo, GoodReceiveCC, InvoiceTo, InvoiceCC, SummaryOutstandingTo, SummaryOutstandingCC)" & _
                                    " VALUES ('" & cboAffiliateCode.Text & "','" & txtEmailToAffiliatePO.Text & "','" & txtEmailCCAffiliatePO.Text & "','" & txtEmailToAffiliatePORevision.Text & "'," & _
                                    "'" & txtEmailCCAffiliatePORevision.Text & "','" & txtEmailToKanban.Text & "','" & txtEmailCCKanban.Text & "','" & txtEmailToSupplierDelivery.Text & "','" & txtEmailCCSupplierDelivery.Text & "', '" & txtEmailToPASIReceiving.Text & "'," & _
                                    "'" & txtEmailCCPASIReceiving.Text & "','" & txtEmailToAffiliateReceiving.Text & "','" & txtEmailCCAffiliateReceiving.Text & "', '" & txtEmailToGoodReceive.Text & "','" & txtEmailCCGoodReceive.Text & "','" & txtEmailToInvoice.Text & "','" & txtEmailCCInvoice.Text & "','" & txtEmailToSummaryOutstanding.Text & "','" & txtEmailCCSummaryOutstanding.Text & "')" & vbCrLf
                        ls_MsgID = "1001"

                        ls_SQL = ls_SQL.Replace(vbCr, "")
                        ls_SQL = ls_SQL.Replace(vbLf, "")
                        ls_SQL = ls_SQL.Replace(vbCrLf, "")

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()

                        AffiliateSubmit.JSProperties("cpFunction") = "insert"

                    ElseIf pIsNewData = False Then
                        '#UPDATE DATA
                        If cbotype.Text = "DOMESTIC" Then
                            ls_SQL = "UPDATE MS_EmailAffiliate SET "
                        Else
                            ls_SQL = "UPDATE MS_EmailAffiliate_EXPORT SET "
                        End If

                        ls_SQL = ls_SQL + "AffiliatePOTO='" & txtEmailToAffiliatePO.Text & "'," & _
                                          "AffiliatePOCC='" & txtEmailCCAffiliatePO.Text & "'," & _
                                          "AffiliatePORevisionTo='" & txtEmailToAffiliatePORevision.Text & "'," & _
                                          "AffiliatePORevisionCC='" & txtEmailCCAffiliatePORevision.Text & "'," & _
                                          "KanbanTo='" & txtEmailToKanban.Text & "'," & _
                                          "KanbanCC='" & txtEmailCCKanban.Text & "'," & _
                                          "SupplierDeliveryTo='" & txtEmailToSupplierDelivery.Text & "'," & _
                                          "SupplierDeliveryCC='" & txtEmailCCSupplierDelivery.Text & "'," & _
                                          "PASIReceivingTo='" & txtEmailToPASIReceiving.Text & "'," & _
                                          "PASIReceivingCC='" & txtEmailCCPASIReceiving.Text & "'," & _
                                          "AffiliateReceivingTo='" & txtEmailToAffiliateReceiving.Text & "'," & _
                                          "AffiliateReceivingCC='" & txtEmailCCAffiliateReceiving.Text & "'," & _
                                          "GoodReceiveTo='" & txtEmailToGoodReceive.Text & "'," & _
                                          "GoodReceiveCC='" & txtEmailCCGoodReceive.Text & "'," & _
                                          "InvoiceTO = '" & txtEmailToInvoice.Text & "'," & _
                                          "InvoiceCC = '" & txtEmailCCInvoice.Text & "'," & _
                                          "SummaryOutstandingTo ='" & txtEmailToSummaryOutstanding.Text & "'," & _
                                          "SummaryOutstandingCC ='" & txtEmailCCSummaryOutstanding.Text & "'" & _
                                          "WHERE AffiliateID='" & pAffiliateID & "'"
                        ls_MsgID = "1002"

                        ls_SQL = ls_SQL.Replace(vbCr, "")
                        ls_SQL = ls_SQL.Replace(vbLf, "")
                        ls_SQL = ls_SQL.Replace(vbCrLf, "")

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()

                        AffiliateSubmit.JSProperties("cpFunction") = "update"
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

End Class