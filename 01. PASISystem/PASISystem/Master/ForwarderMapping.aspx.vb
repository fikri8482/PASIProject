Imports System.Drawing
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxEditors

Public Class ForwarderMapping
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim dsUser As New DataSet
    Dim clsDESEncryption As New clsDESEncryption("TOS")
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not Page.IsPostBack) AndAlso (Not Page.IsCallback) Then
            Session("MenuDesc") = "FORWARDER MAPPING MASTER"
            lblErrMsg.Visible = True
            lblErrMsg.Text = ""
            gridMenu.FocusedRowIndex = -1
            Call TabIndex()
            Call up_GridLoadMenu()
        End If

        txtCCTemp.ForeColor = Color.FromName("White")
        txtUserIDTemp.ForeColor = Color.FromName("White")
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/MainMenu.aspx")
    End Sub


    Private Sub gridMenu_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles gridMenu.BatchUpdate
        Dim ls_SQL As String = "", ls_MenuID As String = "", ls_MsgID As String = ""
        Dim iLoop As Long = 0, jLoop As Long = 0
        Dim ls_air As String = "", ls_boat As String = "", ls_ForwarderID As String = ""
        Dim ls_affiliateID As String = ""

        Dim ls_UserCls As String = IIf(cboAffiliateID.Text = "PASI", 1, 0)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UserMenu")
                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("ZZ010Msg") = lblErrMsg.Text
                    Exit Sub
                End If

                If cboAffiliateID.Text <> "" Then ls_affiliateID = Trim(cboAffiliateID.Text)

                Dim a As Integer
                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1

                    ls_air = IIf(e.UpdateValues(iLoop).NewValues("AirCls").ToString() = "", "0", e.UpdateValues(iLoop).NewValues("AirCls").ToString())
                    ls_boat = IIf(e.UpdateValues(iLoop).NewValues("BoatCls").ToString() = "", "0", e.UpdateValues(iLoop).NewValues("BoatCls").ToString())
                    ls_ForwarderID = e.UpdateValues(iLoop).NewValues("ForwarderID").ToString()

                    If ls_air = True Then ls_air = "1" Else ls_air = "0"
                    If ls_boat = True Then ls_boat = "1" Else ls_boat = "0"

                    ls_SQL = ""
                    If ls_air = 0 Then
                        ls_SQL = ls_SQL + " delete MS_ForwarderMapping WHERE AffiliateID ='" & ls_affiliateID & "' AND ForwarderID='" & ls_ForwarderID & "' AND ShipCls = 'A'" & vbCrLf
                    End If

                    If ls_boat = 0 Then
                        ls_SQL = ls_SQL + " delete MS_ForwarderMapping WHERE AffiliateID ='" & ls_affiliateID & "' AND ForwarderID='" & ls_ForwarderID & "' AND ShipCls = 'B'" & vbCrLf
                    End If

                    If ls_air = 1 Then
                        ls_SQL = ls_SQL + " INSERT INTO dbo.MS_ForwarderMapping" & vbCrLf & _
                                    " VALUES( '" & ls_affiliateID & "', 'A','" & ls_ForwarderID & "')" & vbCrLf
                    End If

                    If ls_boat = 1 Then
                        ls_SQL = ls_SQL + " INSERT INTO dbo.MS_ForwarderMapping" & vbCrLf & _
                                    " VALUES( '" & ls_affiliateID & "', 'B','" & ls_ForwarderID & "')" & vbCrLf
                    End If
                    
                    ls_MsgID = "1002"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                Next iLoop

                sqlTran.Commit()
                Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
                If lblErrMsg.Text = "[] " Then lblErrMsg.Text = ""
                Session("ZZ010Msg") = lblErrMsg.Text
            End Using

            sqlConn.Close()
        End Using
    End Sub

    Private Sub gridMenu_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles gridMenu.CellEditorInitialize
        If (e.Column.FieldName = "fwdid") Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If
    End Sub

    Private Sub gridMenu_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles gridMenu.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Dim pAffiliate As String = Split(e.Parameters, "|")(1)

        Try
            Select Case pAction
                Case "load"
                    Session("AFF") = pAffiliate
                    Call up_GridLoadMenu()
                    gridMenu.PageIndex = 0
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            gridMenu.JSProperties("cpError") = lblErrMsg.Text
            gridMenu.FocusedRowIndex = -1
        End Try
    End Sub

    Private Sub gridMenu_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridMenu.PageIndexChanged
        Dim ls_UserCls As String = IIf(cboAffiliateID.Text = "PASI", 1, 0)
        If cboAffiliateID.Text <> "" Then
            up_GridLoadMenu()
        Else
            up_GridLoadMenu()
        End If
    End Sub
#End Region

#Region "PROCEDURE"

    Public Sub up_GridLoadMenu()
        Dim ls_SQL As String = ""
        Dim ls_affiliateID As String = ""
        Dim ls_Where As String = ""

        If cboAffiliateID.Text = "" Then
            ls_affiliateID = Trim(cboAffiliateID.Text)
        Else
            ls_affiliateID = Trim(cboAffiliateID.Text)
        End If

        ls_Where = " Where AffiliateID = '" & ls_affiliateID & "'"

        'GridMenuP
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            'ls_SQL = " select fwdid = B.ForwarderID, air = isnull(airCls,0),boat = isnull(boatcls,0)  " & vbCrLf & _
            '      " From MS_Forwarder B left join MS_ForwarderMapping A ON A.ForwarderID = B.ForwarderID " & vbCrLf & _
            '      " Where AffiliateID = '" & ls_affiliateID & "'  " & vbCrLf & _
            '      " UNION ALL " & vbCrLf & _
            '      " select fwdid = ForwarderID, air =0,boat = 0 from MS_Forwarder  " & vbCrLf & _
            '      " where forwarderID not IN(select forwarderID from MS_ForwarderMapping where AffiliateID = '" & ls_affiliateID & "') "

            ls_SQL = " SELECT a.ForwarderID, CASE WHEN ShipCls = 'A' then '1' else '0' END AirCls, CASE WHEN ShipCls = 'B' then '1' else '0' END BoatCls " & vbCrLf & _
                  "  FROM MS_Forwarder a LEFT JOIN " & vbCrLf & _
                  "  MS_ForwarderMapping b ON a.ForwarderID = b.ForwarderID " & vbCrLf & _
                  "  WHERE AffiliateID = '" & ls_affiliateID & "' " & vbCrLf & _
                  "  UNION ALL  " & vbCrLf & _
                  "  SELECT a.ForwarderID, air=0,boat = 0  " & vbCrLf & _
                  "  from MS_Forwarder a " & vbCrLf & _
                  "  where forwarderID not IN(select forwarderID from MS_ForwarderMapping where AffiliateID = '" & ls_affiliateID & "')  " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With gridMenu
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
            
            sqlConn.Close()
        End Using
    End Sub

    Public Sub up_SaveData(ByVal pIsUpdate As Boolean, _
                            Optional ByVal pCCCode As String = "", _
                            Optional ByVal pAir As String = "", _
                            Optional ByVal pBoat As String = "", _
                            Optional ByVal pFwd As String = "")
        Dim ls_SQL As String = "", ls_MsgID As String = ""


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("UserSetup")
                If pIsUpdate = True Then
                    'INSERT DATA
                    ls_SQL = " INSERT INTO dbo.MS_ForwarderMapping " & vbCrLf & _
                             " VALUES ('" & pCCCode & "','" & Trim(pFwd) & "', '" & pAir & "' ,'" & Trim(pBoat) & "')"
                    ls_MsgID = "1001"
                Else
                    ls_SQL = " UPDATE dbo.MS_ForwarderMapping " & vbCrLf & _
                             " SET AirCls='" & Trim(pAir) & "', " & vbCrLf & _
                             " BoatCls='" & Trim(pBoat) & "' " & vbCrLf & _
                             " WHERE AffiliateID='" & Trim(pCCCode) & "' AND " & vbCrLf & _
                             " ForwarderID='" & Trim(pFwd) & "'"

                    ls_MsgID = "1002"
                End If

                Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 0, False, clsAppearance.PagerMode.ShowPager)
        Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
        lblErrMsg.Visible = True
    End Sub

    Private Sub TabIndex()
        cboAffiliateID.TabIndex = 1
        gridMenu.TabIndex = 2
        btnSubmit.TabIndex = 3
        btnSubMenu.TabIndex = 4
    End Sub

    Private Function validation() As Boolean
        If cboAffiliateID.Text = "" Then
            Call clsMsg.DisplayMessage(lblErrMsg, "7003", clsMessage.MsgType.ErrorMessage)
            gridMenu.JSProperties("cpMessage") = lblErrMsg.Text
            lblErrMsg.Visible = True
            Return False
        Else
            Return True
        End If
    End Function

#End Region

End Class