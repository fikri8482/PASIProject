Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing

Public Class PartMasterDetail
#Region "DECLARATION"
    Inherits System.Web.UI.Page

    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean
    Dim pub_PartNo As String
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "A06"
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)
            ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

            up_FillCombo()

            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")

                'Else

            End If

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Session("M01Url") <> "" Then
                    flag = False
                    If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                        Session("MenuDesc") = "PART MASTER ENTRY"
                        pub_PartNo = Request.QueryString("id")
                        tabIndex()
                        bindData()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "BACK"
                        txtPartID.ReadOnly = True
                        txtPartID.BackColor = Color.FromName("#CCCCCC")
                    Else
                        Session("MenuDesc") = "PART MASTER ENTRY"
                        tabIndex()
                        clear()
                        lblInfo.Text = ""
                        btnSubMenu.Text = "Back"
                        btnClear.Visible = True
                    End If
                Else
                    flag = True
                    btnClear.Visible = True
                    txtPartID.Focus()
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

    Private Sub PartMasterSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles PartMasterSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"
                    Dim pPartNo As String = Split(e.Parameter, "|")(2)
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pPartNo)
                    If btnSubmit.Text = "RECOVERY" Then
                        Call SaveDataRec(lb_IsUpdate, _
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
                                     Split(e.Parameter, "|")(12))
                    Else
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
                                     Split(e.Parameter, "|")(12))
                    End If
                    
                    'bindData()

                Case "delete"
                    Dim pPartNo As String = Split(e.Parameter, "|")(1)
                    If AlreadyUsed(pPartNo) = False Then
                        Call DeleteData(pPartNo)
                    End If

                Case "load"
                    pub_PartNo = txtPartID.Text
                    PartMasterSubmit.JSProperties("cpKeyPress") = "ON"
                    bindData()
                    lblInfo.Text = ""
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'grid.JSProperties("cpMessage") = lblInfo.Text
        End Try
    End Sub

    Protected Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubMenu.Click
        Session.Remove("MsgID")
        Session.Remove("1PartNo")
        Session.Remove("1PartName")
        Session.Remove("1PartCarMaker")
        Session.Remove("1PartCarName")
        Session.Remove("1PartNameGroup")
        Session.Remove("1HSCode")
        Session.Remove("1Maker")
        Session.Remove("1Project")
        Session.Remove("1UOM")
        Session.Remove("1KanbanCls")

        If Session("M01Url") <> "" Then
            'Session.Remove("M01Url")
            Response.Redirect("~/Master/PartMaster.aspx")
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

            ls_SQL = "  select * from ( select  " & vbCrLf & _
                      "  	RTRIM(PartNo)PartNo,  " & vbCrLf & _
                      "  	RTRIM(PartName)PartName,  " & vbCrLf & _
                      "  	RTRIM(PartCarMaker)PartCarMaker,  " & vbCrLf & _
                      "  	RTRIM(PartCarName)PartCarName,  " & vbCrLf & _
                      "  	RTRIM(PartGroupName)PartNameGroup,  " & vbCrLf & _
                      "  	RTRIM(HSCode)HSCode,  " & vbCrLf & _
                      "  	FinishGoodCls,  " & vbCrLf & _
                      "     RTRIM(b.Description) UOM,  " & vbCrLf & _
                      "     RTRIM(a.Maker) Maker,  " & vbCrLf & _
                      "     RTRIM(a.Project) Project,  " & vbCrLf & _
                      "  	KanbanCls,  " & vbCrLf & _
                      "  'DETAIL' DetailPage, 0 DeleteCls  " & vbCrLf & _
                      "  from MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls " & vbCrLf

            ls_SQL = ls_SQL + "UNION ALL" & vbCrLf & _
                      "  select  " & vbCrLf & _
                      "  	RTRIM(PartNo)PartNo,  " & vbCrLf & _
                      "  	RTRIM(PartName)PartName,  " & vbCrLf & _
                      "  	RTRIM(PartCarMaker)PartCarMaker,  " & vbCrLf & _
                      "  	RTRIM(PartCarName)PartCarName,  " & vbCrLf & _
                      "  	RTRIM(PartGroupName)PartNameGroup,  " & vbCrLf & _
                      "  	RTRIM(HSCode)HSCode,  " & vbCrLf & _
                      "  	FinishGoodCls,  " & vbCrLf & _
                      "     RTRIM(b.Description) UOM,  " & vbCrLf & _
                      "     RTRIM(a.Maker) Maker,  " & vbCrLf & _
                      "     RTRIM(a.Project) Project,  " & vbCrLf & _
                      "  	KanbanCls,  " & vbCrLf & _
                      "  'DETAIL' DetailPage, 1 DeleteCls  " & vbCrLf & _
                      "  from MS_Parts_History a left join MS_UnitCls b on a.UnitCls = b.UnitCls )xyz" & vbCrLf & _
                      "  where PartNo = '" & pub_PartNo & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtPartID.Text = ds.Tables(0).Rows(0)("PartNo") & ""
                PartMasterSubmit.JSProperties("cpPartNo") = txtPartID.Text

                txtPartName.Text = ds.Tables(0).Rows(0)("PartName") & ""
                PartMasterSubmit.JSProperties("cpPartName") = txtPartName.Text

                txtCarMakerCode.Text = ds.Tables(0).Rows(0)("PartCarMaker") & ""
                PartMasterSubmit.JSProperties("cpCarMakerCode") = txtPartID.Text

                txtCarMakerName.Text = ds.Tables(0).Rows(0)("PartCarName") & ""
                PartMasterSubmit.JSProperties("cpCarMakerName") = txtPartName.Text

                txtPartNameGroup.Text = ds.Tables(0).Rows(0)("PartNameGroup") & ""
                PartMasterSubmit.JSProperties("cpPartNameGroup") = txtPartNameGroup.Text

                txtHSCode.Text = ds.Tables(0).Rows(0)("HSCode") & ""
                PartMasterSubmit.JSProperties("cpHSCode") = txtHSCode.Text

                txtMaker.Text = ds.Tables(0).Rows(0)("Maker") & ""
                PartMasterSubmit.JSProperties("cpMaker") = txtMaker.Text

                txtProject.Text = ds.Tables(0).Rows(0)("Project") & ""
                PartMasterSubmit.JSProperties("cpProject") = txtProject.Text

                cboUnit.Text = ds.Tables(0).Rows(0)("UOM") & ""
                PartMasterSubmit.JSProperties("cpUOM") = cboUnit.Text

                'If ds.Tables(0).Rows(0)("FinishGoodCls") = "1" Then
                '    rdrFG.Checked = True
                '    PartMasterSubmit.JSProperties("cpFinishGoodCls") = 1
                'Else
                '    rdrPart.Checked = True
                '    PartMasterSubmit.JSProperties("cpFinishGoodCls") = 0
                'End If

                If ds.Tables(0).Rows(0)("KanbanCls") = "1" Then
                    rdrYes.Checked = True
                    PartMasterSubmit.JSProperties("cpKanbanCls") = 1
                Else
                    rdrNo.Checked = True
                    PartMasterSubmit.JSProperties("cpKanbanCls") = 0
                End If

                Session("1PartNo") = txtPartID.Text
                Session("1PartName") = txtPartName.Text
                Session("1PartCarMaker") = txtCarMakerCode.Text
                Session("1PartCarName") = txtCarMakerName.Text
                Session("1PartNameGroup") = txtPartNameGroup.Text
                Session("1HSCode") = txtHSCode.Text
                Session("1Maker") = txtMaker.Text
                Session("1Project") = txtProject.Text
                Session("1UOM") = cboUnit.Text
                Session("1KanbanCls") = ds.Tables(0).Rows(0)("KanbanCls")

                If ds.Tables(0).Rows(0)("DeleteCls") = "1" Then
                    btnDelete.Visible = False
                    btnClear.Visible = False
                    btnSubmit.Text = "RECOVERY"
                Else
                    btnDelete.Visible = True
                    btnClear.Visible = True
                    btnSubmit.Text = "SAVE"
                End If
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub clear()
        txtPartID.Text = ""
        txtPartName.Text = ""
        txtCarMakerCode.Text = ""
        txtCarMakerName.Text = ""
        txtPartNameGroup.Text = ""
        txtHSCode.Text = ""
        txtMaker.Text = ""
        txtProject.Text = ""
        cboUnit.SelectedIndex = 0
        'rdrPart.Checked = True
        'rdrFG.Checked = False
        rdrYes.Checked = True
        rdrNo.Checked = False

        txtPartID.ReadOnly = False
        txtPartID.BackColor = Color.FromName("#FFFFFF")
        lblInfo.Text = ""
    End Sub

    Private Sub tabIndex()
        txtPartID.TabIndex = 1
        txtPartName.TabIndex = 2
        txtCarMakerCode.TabIndex = 3
        txtCarMakerName.TabIndex = 4
        txtPartNameGroup.TabIndex = 5
        txtHSCode.TabIndex = 6

        'rdrFG.TabIndex = 7
        'rdrPart.TabIndex = 8
        cboUnit.TabIndex = 9

        txtMaker.TabIndex = 10
        txtProject.TabIndex = 11

        rdrYes.TabIndex = 12
        rdrNo.TabIndex = 13

        btnSubmit.TabIndex = 14
        btnDelete.TabIndex = 15
        btnClear.TabIndex = 16
        btnSubMenu.TabIndex = 17
    End Sub

    Private Function AlreadyUsed(ByVal pPartNo As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo From MS_PartMapping WHERE PartNo= '" & Trim(pPartNo) & "'" & vbCrLf & _
                         " Union ALL" & vbCrLf & _
                         " SELECT PartNo From MS_PartConversion WHERE PartNo= '" & Trim(pPartNo) & "'"

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    Call clsMsg.DisplayMessage(lblInfo, "5003", clsMessage.MsgType.ErrorMessage)
                    PartMasterSubmit.JSProperties("cpMessage") = lblInfo.Text
                    PartMasterSubmit.JSProperties("cpType") = "error"
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

    Private Sub DeleteData(ByVal pPartNo As String)
        Try
            Dim ls_SQL As String = ""
            Dim x As Integer
            Dim shostname As String = System.Net.Dns.GetHostName

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("MS_Parts_Delete")
                    ls_SQL = " INSERT INTO MS_PARTS_HISTORY" & vbCrLf & _
                              "SELECT * FROM MS_Parts WHERE PARTNO ='" & pPartNo & "'"                    
                    Dim SqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    'delete from history
                    ls_SQL = " DELETE MS_Parts " & vbCrLf & _
                                " WHERE PartNo = '" & pPartNo & "' " & vbCrLf
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()


                    'insert into history
                    ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                             "VALUES ('" & shostname & "','" & menuID & "','D','" & pPartNo & "','Delete PartNo " & pPartNo & "', GETDATE(),'" & Session("UserID") & "')  "
                    SqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            If x > 0 Then
                Call clsMsg.DisplayMessage(lblInfo, "1003", clsMessage.MsgType.InformationMessage)
                PartMasterSubmit.JSProperties("cpMessage") = lblInfo.Text
                PartMasterSubmit.JSProperties("cpType") = "info"
                PartMasterSubmit.JSProperties("cpFunction") = "delete"
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub up_FillCombo()
        Dim ls_SQL As String = ""

        'Unit
        ls_SQL = "select RTRIM(UnitCls)UnitCls, rtrim(Description) UnitDesc from MS_UnitCls" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboUnit
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("UnitCls")
                .Columns(0).Width = 50
                .Columns.Add("UnitDesc")
                .Columns(1).Width = 120

                .TextField = "UnitCls"
                .DataBind()
                .SelectedIndex = 0
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Function ValidasiInput(ByVal pPartNo As String) As Boolean
        Try

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

    End Function

    Private Sub SaveData(ByVal pIsNewData As Boolean, _
                         Optional ByVal pPartNo As String = "", _
                         Optional ByVal pPartName As String = "", _
                         Optional ByVal pCarMakerCode As String = "", _
                         Optional ByVal pCarMakerName As String = "", _
                         Optional ByVal pPartNameGroup As String = "", _
                         Optional ByVal pHSCode As String = "", _
                         Optional ByVal pFGCls As String = "", _
                         Optional ByVal pUnit As String = "", _
                         Optional ByVal pMaker As String = "", _
                         Optional ByVal pProject As String = "", _
                         Optional ByVal pKanbanCls As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT PartNo FROM MS_Parts WHERE PartNo = '" & Trim(pPartNo) & "'"

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

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("MS_Parts_Update")

                Dim sqlComm As New SqlCommand()

                If pIsNewData = True Then
                    '#INSERT NEW DATA
                    ls_SQL = " INSERT INTO MS_Parts " & _
                                "(PartNo, PartName, PartCarMaker, PartCarName, PartGroupName, HSCode, UnitCls, Maker, Project, KanbanCls, EntryDate, EntryUser)" & _
                                " VALUES ('" & pPartNo & "','" & pPartName & "','" & pCarMakerCode & "','" & pCarMakerName & "','" & pPartNameGroup & "','" & pHSCode & "','" & pUnit & "','" & pMaker & "','" & pProject & "', '" & pKanbanCls & "', getdate(), '" & admin & "')"
                    ls_MsgID = "1001"
                    Session("MsgID") = ls_MsgID

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    PartMasterSubmit.JSProperties("cpFunction") = "insert"
                    flag = False
                ElseIf pIsNewData = False And flag = True Then
                    ls_MsgID = "6018"
                    Session("MsgID") = ls_MsgID

                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    PartMasterSubmit.JSProperties("cpMessage") = lblInfo.Text
                    PartMasterSubmit.JSProperties("cpType") = "error"
                    Exit Sub

                ElseIf pIsNewData = False Then
                    '#UPDATE DATA
                    ls_SQL = "UPDATE MS_Parts SET " & _
                            "PartName='" & pPartName & "'," & _
                            "PartCarMaker='" & pCarMakerCode & "'," & _
                            "PartCarName='" & pCarMakerName & "'," & _
                            "PartGroupName='" & pPartNameGroup & "'," & _
                            "HSCode='" & pHSCode & "'," & _                           
                            "UnitCls='" & pUnit & "'," & _
                            "Maker = '" & pMaker & "'," & _
                            "Project = '" & pProject & "'," & _
                            "KanbanCls='" & pKanbanCls & "'," & _
                            "UpdateDate = getdate()," & _
                            "UpdateUser ='" & admin & "'" & _
                            "WHERE PartNo='" & pPartNo & "'"
                    ls_MsgID = "1002"
                    Session("MsgID") = ls_MsgID

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()

                    Dim ls_Remarks As String = ""

                    If Session("1PartName").ToString.Trim <> pPartName.Trim Then
                        ls_Remarks = ls_Remarks + "PartName " + Session("1PartName").ToString.Trim & "->" & pPartName.Trim & " "
                    End If
                    If Session("1PartCarMaker").ToString.Trim <> pCarMakerCode.Trim Then
                        ls_Remarks = ls_Remarks + "PartCarCode " + Session("1PartCarMaker").ToString.Trim & "->" & pCarMakerCode.Trim & " "
                    End If
                    If Session("1PartCarName").ToString.Trim <> pCarMakerName.Trim Then
                        ls_Remarks = ls_Remarks + "PartCarName " + Session("1PartCarName").ToString.Trim & "->" & pCarMakerName.Trim & " "
                    End If
                    If Session("1PartNameGroup").ToString.Trim <> pPartNameGroup.Trim Then
                        ls_Remarks = ls_Remarks + "PartNameGroup " + Session("1PartNameGroup").ToString.Trim & "->" & pPartNameGroup.Trim & " "
                    End If
                    If Session("1HSCode").ToString.Trim <> pHSCode.Trim Then
                        ls_Remarks = ls_Remarks + "HSCode " + Session("1HSCode").ToString.Trim & "->" & pHSCode.Trim & " "
                    End If
                    If Session("1Maker").ToString.Trim <> pMaker.Trim Then
                        ls_Remarks = ls_Remarks + "Maker " + Session("1Maker").ToString.Trim & "->" & pMaker.Trim & " "
                    End If
                    If Session("1Project").ToString.Trim <> pProject.Trim Then
                        ls_Remarks = ls_Remarks + "Project " + Session("1Project").ToString.Trim & "->" & pProject.Trim & " "
                    End If

                    If pUnit = "01" Then
                        pUnit = "PC"
                    ElseIf pUnit = "" Then
                        pUnit = "KG"
                    ElseIf pUnit = "" Then
                        pUnit = "BOX"
                    ElseIf pUnit = "04" Then
                        pUnit = "PALLET"
                    ElseIf pUnit = "05" Then
                        pUnit = "MM"
                    ElseIf pUnit = "06" Then
                        pUnit = "M"
                    ElseIf pUnit = "07" Then
                        pUnit = "GR"
                    Else
                        pUnit = "PC"
                    End If

                    If Session("1UOM").ToString.Trim <> pUnit Then
                        ls_Remarks = ls_Remarks + "UOM " + Session("1UOM").ToString.Trim & "->" & pUnit & " "
                    End If
                    If Session("1KanbanCls").ToString.Trim <> pKanbanCls.Trim Then
                        ls_Remarks = ls_Remarks + "KanbanCls " + Session("1KanbanCls").ToString.Trim & "->" & pKanbanCls.Trim & " "
                    End If

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U','" & pPartNo & "','Update [" & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                    End If
                    

                    PartMasterSubmit.JSProperties("cpFunction") = "update"
                    flag = False
                End If

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            sqlConn.Close()
        End Using
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        PartMasterSubmit.JSProperties("cpMessage") = lblInfo.Text
        PartMasterSubmit.JSProperties("cpType") = "info"

    End Sub

    Private Sub SaveDataRec(ByVal pIsNewData As Boolean, _
                         Optional ByVal pPartNo As String = "", _
                         Optional ByVal pPartName As String = "", _
                         Optional ByVal pCarMakerCode As String = "", _
                         Optional ByVal pCarMakerName As String = "", _
                         Optional ByVal pPartNameGroup As String = "", _
                         Optional ByVal pHSCode As String = "", _
                         Optional ByVal pFGCls As String = "", _
                         Optional ByVal pUnit As String = "", _
                         Optional ByVal pMaker As String = "", _
                         Optional ByVal pProject As String = "", _
                         Optional ByVal pKanbanCls As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString
        Dim shostname As String = System.Net.Dns.GetHostName

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("MS_Parts_Update")

                Dim sqlComm As New SqlCommand()

                'insert into master
                ls_SQL = " INSERT INTO MS_Parts " & vbCrLf & _
                            "SELECT * FROM MS_PARTS_HISTORY" & vbCrLf & _
                            "WHERE PARTNO ='" & pPartNo & "'"
                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()


                'delete from history
                ls_SQL = " DELETE FROM MS_PARTS_HISTORY" & vbCrLf & _
                          "WHERE PARTNO ='" & pPartNo & "'"
                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()


                'insert into history
                ls_SQL = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                         "VALUES ('" & shostname & "','" & menuID & "','R','" & pPartNo & "','Recovery PartNo " & pPartNo & "',GETDATE(),'" & admin & "')  "
                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()

                sqlComm.Dispose()
                sqlTran.Commit()
                ls_MsgID = "1016"
                Session("MsgID") = "1016"
            End Using

            sqlConn.Close()
        End Using
        '{http://localhost:16054/Master/PartMasterDetail.aspx?id=7009-1343-02&t1=&t2=&Session=~/Master/PartMaster.aspx}
        'Response.Redirect("~/Master/PartMaster.aspx")
        'Dim rsUrl As String = Request.Url.ToString
        'Response.Redirect(Request.Url.ToString)
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        PartMasterSubmit.JSProperties("cpMessage") = lblInfo.Text
        PartMasterSubmit.JSProperties("cpType") = "info"

    End Sub
#End Region

    'Private Sub btnSubmit_Click(sender As Object, e As System.EventArgs) Handles btnSubmit.Click
    '    If btnSubmit.Text = "RECOVERY" Then
    '        btnDelete.Visible = True
    '        btnClear.Visible = True
    '        btnSubmit.Text = "SAVE"
    '        Session("MenuDesc") = "PART MASTER ENTRY"
    '        pub_PartNo = Request.QueryString("id")
    '        tabIndex()
    '        bindData()
    '        Call clsMsg.DisplayMessage(lblInfo, Session("MsgID"), clsMessage.MsgType.InformationMessage)
    '        btnSubMenu.Text = "BACK"
    '        txtPartID.ReadOnly = True
    '        txtPartID.BackColor = Color.FromName("#CCCCCC")
    '    Else
    '        btnDelete.Visible = True
    '        btnClear.Visible = True
    '        btnSubmit.Text = "SAVE"
    '        Session("MenuDesc") = "PART MASTER ENTRY"
    '        pub_PartNo = Request.QueryString("id")
    '        tabIndex()
    '        bindData()
    '        Call clsMsg.DisplayMessage(lblInfo, Session("MsgID"), clsMessage.MsgType.InformationMessage)
    '        btnSubMenu.Text = "BACK"
    '        txtPartID.ReadOnly = True
    '        txtPartID.BackColor = Color.FromName("#CCCCCC")
    '    End If
    'End Sub
End Class