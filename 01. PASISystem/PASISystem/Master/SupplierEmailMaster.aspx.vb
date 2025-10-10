Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing

Public Class SupplierEmailMaster
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_SupplierID As String
    Dim ls_AllowUpdate As Boolean = False
    Dim menuID As String = "A12"
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
                cbotype.Items.Clear()
                cbotype.Items.Add("DOMESTIC")
                cbotype.Items.Add("EXPORT")
                cbotype.Text = "DOMESTIC"

                If Session("M01Url") <> "" Then
                    flag = False
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "SUPPLIER EMAIL MASTER"
                        pub_SupplierID = Request.QueryString("id")
                        tabIndex()
                        'bindData()
                        lblInfo.Text = ""
                        txtEmailToAffiliatePO.ReadOnly = True
                    Else
                        Session("MenuDesc") = "SUPPLIER EMAIL MASTER"
                        flag = True
                        btnClear.Visible = True
                        cboSupplierCode.Focus()
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

    Private Sub AffiliateSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles AffiliateSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"
                    Dim pAffiliateID As String = Split(e.Parameter, "|")(1)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pub_SupplierID)
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

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        clear()

        up_FillCombo()

        flag = True
    End Sub
#End Region

#Region "PROCEDURE"

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT RTRIM(SupplierID) SupplierID, RTRIM(SupplierName) SupplierName from MS_Supplier order by SupplierID " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 85
                .Columns.Add("SupplierName")
                .Columns(1).Width = 400

                .TextField = "SupplierID"
                .DataBind()
                '.SelectedIndex = 0
                'txtSupplierCode.Text = clsGlobal.gs_empty
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub up_FillComboSubmit()
        Dim ls_SQL As String = ""

        ls_SQL = "SELECT '" & clsGlobal.gs_All & "' SupplierID, '" & clsGlobal.gs_All & "' AffiliatePOTO, AffiliatePOCC, AffiliatePORevisionTo, AffiliatePORevisionCC, KanbanTo, KanbanCC, SupplierDeliveryTo, SupplierDeliveryCC, PASIReceivingTo, PASIReceivingCC, AffiliateReceivingTo, AffiliateReceivingCC, GoodReceiveTo, GoodReceiveCC, InvoiceTO, InvoiceCC, SummaryOutstandingTo, SummaryOutstandingCC select RTRIM(SupplierID) SupplierCode,  from MS_EmailSupplier order by SupplierCode " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierCode")
                .Columns(0).Width = 85
                .Columns.Add("AffiliatePOTO")
                .Columns(1).Width = 335
                .Columns.Add("AffiliatePOCC")
                .Columns(2).Width = 335
                .Columns.Add("AffiliatePORevisionTo")
                .Columns(3).Width = 335
                .Columns.Add("AffiliatePORevisionCC")
                .Columns(4).Width = 335
                .Columns.Add("KanbanTo")
                .Columns(5).Width = 335
                .Columns.Add("KanbanCC")
                .Columns(6).Width = 335
                .Columns.Add("SupplierDeliveryTo")
                .Columns(7).Width = 335
                .Columns.Add("SupplierDeliveryCC")
                .Columns(8).Width = 335
                .Columns.Add("PASIReceivingTo")
                .Columns(9).Width = 335
                .Columns.Add("PASIReceivingCC")
                .Columns(10).Width = 335
                .Columns.Add("AffiliateReceivingTo")
                .Columns(11).Width = 335
                .Columns.Add("AffiliateReceivingCC")
                .Columns(12).Width = 335
                .Columns.Add("GoodReceiveTo")
                .Columns(13).Width = 335
                .Columns.Add("GoodReceiveCC")
                .Columns(14).Width = 335
                .Columns.Add("InvoiceTO")
                .Columns(15).Width = 335
                .Columns.Add("InvoiceCC")
                .Columns(16).Width = 335
                .Columns.Add("SummaryOutstandingTo")
                .Columns(17).Width = 335
                .Columns.Add("SummaryOutstandingCC")
                .Columns(18).Width = 335

                .TextField = "SupplierCode"
                .DataBind()
                .SelectedIndex = 0
                txtSupplierCode.Text = clsGlobal.gs_All
            End With

            sqlConn.Close()
        End Using

    End Sub

    Private Sub clear()
        cboSupplierCode.Text = ""
        txtSupplierCode.Text = ""
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
        cboSupplierCode.TabIndex = 1
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

    'Private Function AlreadyUsed(ByVal pAffiliate As String) As Boolean
    '    Try
    '        Dim ls_SQL As String = ""
    '        'Dim ls_MsgID As String = ""
    '        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '            sqlConn.Open()

    '            ls_SQL = " SELECT SupplierID FROM MS_EmailSupplier WHERE SupplierID= '" & Trim(pAffiliate) & "'"

    '            sqlConn.Close()
    '        End Using
    '    Catch ex As Exception
    '        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '    End Try
    'End Function

    Private Function ValidasiInput(ByVal pAffiliate As String) As Boolean
        Try

            Dim ls_MsgID As String = ""

            If cboSupplierCode.Text = "" Then
                ls_MsgID = "6010"
                Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                Return False
            ElseIf txtSupplierCode.Text = "" Then
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
                         Optional ByVal pSupplierID As String = "", _
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
                    ls_SQL = " SELECT SupplierID FROM MS_EmailSUpplier WHERE SupplierID= '" & Trim(pSupplierID) & "'"
                Else
                    ls_SQL = " SELECT SupplierID FROM MS_EmailSUpplier_EXPORT WHERE SupplierID= '" & Trim(pSupplierID) & "'"
                End If

                'ls_SQL = " SELECT SupplierID FROM MS_EmailSUpplier WHERE SupplierID= '" & Trim(pSupplierID) & "'"

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

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("EmailSupplier")

                    Dim sqlComm As New SqlCommand()

                    If pIsNewData = True Then
                        '#INSERT NEW DATA
                        If cbotype.Text = "DOMESTIC" Then
                            ls_SQL = " INSERT INTO MS_EmailSupplier "
                        Else
                            ls_SQL = " INSERT INTO MS_EmailSupplier_EXPORT "
                        End If

                        ls_SQL = ls_SQL + "(SupplierID, AffiliatePOTO, AffiliatePOCC, AffiliatePORevisionTO, AffiliatePORevisionCC, KanbanTO, KanbanCC, SupplierDeliveryTO, SupplierDeliveryCC, PASIReceivingTo, PASIReceivingCC, AffiliateReceivingTo, AffiliateReceivingCC, GoodReceiveTo, GoodReceiveCC, InvoiceTo, InvoiceCC, SummaryOutstandingTo, SummaryOutstandingCC)" & _
                                    " VALUES ('" & cboSupplierCode.Text & "','" & txtEmailToAffiliatePO.Text & "','" & txtEmailCCAffiliatePO.Text & "','" & txtEmailToAffiliatePORevision.Text & "'," & _
                                    "'" & txtEmailCCAffiliatePORevision.Text & "','" & txtEmailToKanban.Text & "','" & txtEmailCCKanban.Text & "','" & txtEmailToSupplierDelivery.Text & "','" & txtEmailCCSupplierDelivery.Text & "', '" & txtEmailToPASIReceiving.Text & "'," & _
                                    "'" & txtEmailCCPASIReceiving.Text & "','" & txtEmailToAffiliateReceiving.Text & "','" & txtEmailCCAffiliateReceiving.Text & "', '" & txtEmailToGoodReceive.Text & "','" & txtEmailCCGoodReceive.Text & "','" & txtEmailToInvoice.Text & "','" & txtEmailCCInvoice.Text & "','" & txtEmailToSummaryOutstanding.Text & "','" & txtEmailCCSummaryOutstanding.Text & "')" & vbCrLf
                        ls_MsgID = "1001"

                        ls_SQL = ls_SQL.Replace(vbCr, "")
                        ls_SQL = ls_SQL.Replace(vbLf, "")
                        ls_SQL = ls_SQL.Replace(vbCrLf, "")

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
                        If cbotype.Text = "DOMESTIC" Then
                            ls_SQL = " UPDATE MS_EmailSupplier SET "
                        Else
                            ls_SQL = " UPDATE MS_EmailSupplier_Export SET "
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
                                          "WHERE SupplierID='" & pSupplierID & "'"
                        ls_MsgID = "1002"

                        ls_SQL = ls_SQL.Replace(vbCr, "")
                        ls_SQL = ls_SQL.Replace(vbLf, "")
                        ls_SQL = ls_SQL.Replace(vbCrLf, "")

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

                ls_SQL = "SELECT [AffiliatePOTO],[AffiliatePOCC],[AffiliatePORevisionTO],[AffiliatePORevisionCC],[KanbanTO], " & vbCrLf & _
                         "       [KanbanCC], [SupplierDeliveryTO], [SupplierDeliveryCC], [PASIReceivingTO], [PASIReceivingCC], [AffiliateReceivingTO], " & vbCrLf & _
                         "       [AffiliateReceivingCC], [GoodReceiveTO], [GoodReceiveCC], [InvoiceTO], [InvoiceCC], [SummaryOutstandingTO], [SummaryOutstandingCC] " & vbCrLf
                If cbotype.Text = "DOMESTIC" Then
                    ls_SQL = ls_SQL + "  FROM MS_EmailSupplier  " & vbCrLf
                Else
                    ls_SQL = ls_SQL + "  FROM MS_EmailSupplier_EXPORT  " & vbCrLf
                End If

                '"  FROM MS_EmailSupplier  " & vbCrLf & _

                ls_SQL = ls_SQL + " WHERE SupplierID = '" & e.Parameter & "' " & vbCrLf

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

End Class