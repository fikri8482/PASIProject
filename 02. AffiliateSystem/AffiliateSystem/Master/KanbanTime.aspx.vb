Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports System.Drawing
Imports DevExpress.Utils

Public Class KanbanTime
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim flag As Boolean = True
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            Session("MenuDesc") = "KANBAN TIME CYCLE"
            bindData()
            Cycle1.Focus()
            tabIndex()
            lblInfo.Text = ""
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, True, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub AffiliateSubmit_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles AffiliateSubmit.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "save"

                    Dim pAffiliate As String
                    Dim pCycle As String
                    Dim pTime As String
                    Dim lb_IsUpdate As Boolean = ValidasiInput(pAffiliate)

                    pAffiliate = Session("AffiliateID").ToString
                    pCycle = "1"
                    pTime = Cycle1.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)

                    pCycle = "2"
                    pTime = Cycle2.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)

                    pCycle = "3"
                    pTime = Cycle3.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)

                    pCycle = "4"
                    pTime = Cycle4.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "5"
                    pTime = Cycle5.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "6"
                    pTime = Cycle6.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "7"
                    pTime = Cycle7.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "8"
                    pTime = Cycle8.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "9"
                    pTime = Cycle9.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "10"
                    pTime = Cycle10.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "11"
                    pTime = Cycle11.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "12"
                    pTime = Cycle12.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "13"
                    pTime = Cycle13.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "14"
                    pTime = Cycle14.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "15"
                    pTime = Cycle15.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "16"
                    pTime = Cycle16.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "17"
                    pTime = Cycle17.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "18"
                    pTime = Cycle18.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "19"
                    pTime = Cycle19.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                    pCycle = "20"
                    pTime = Cycle20.Text
                    Call SaveData(lb_IsUpdate, _
                                     pAffiliate, _
                                     pCycle, _
                                     pTime)
                Case Else
                    
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())

        End Try
    End Sub

    Private Sub tabIndex()
        Cycle1.TabIndex = 1
        Cycle2.TabIndex = 2
        Cycle3.TabIndex = 3
        Cycle4.TabIndex = 4

        Cycle5.TabIndex = 5
        Cycle6.TabIndex = 6
        Cycle7.TabIndex = 7
        Cycle8.TabIndex = 8

        Cycle9.TabIndex = 9
        Cycle10.TabIndex = 10
        Cycle11.TabIndex = 11
        Cycle12.TabIndex = 12

        Cycle13.TabIndex = 13
        Cycle14.TabIndex = 14
        Cycle15.TabIndex = 15
        Cycle16.TabIndex = 16

        Cycle17.TabIndex = 17
        Cycle18.TabIndex = 18
        Cycle19.TabIndex = 19
        Cycle20.TabIndex = 20

        btnSubmit.TabIndex = 21
        btnSubMenu.TabIndex = 22
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub bindData()

        Dim ls_SQL As String = ""
        
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " Select   " & vbCrLf & _
                  " 	rtrim(AffiliateID) AffiliateID,	 " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 1 then KanbanTime end),'00:00') Kanban1, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 2 then KanbanTime end),'00:00') Kanban2, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 3 then KanbanTime end),'00:00') Kanban3, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 4 then KanbanTime end),'00:00') Kanban4, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 5 then KanbanTime end),'00:00') Kanban5, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 6 then KanbanTime end),'00:00') Kanban6, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 7 then KanbanTime end),'00:00') Kanban7, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 8 then KanbanTime end),'00:00') Kanban8, " & vbCrLf & _
                  " 	isnull(max(case when KanbanCycle = 9 then KanbanTime end),'00:00') Kanban9, "

            ls_SQL = ls_SQL + " 	isnull(max(case when KanbanCycle = 10 then KanbanTime end),'00:00') Kanban10, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 11 then KanbanTime end),'00:00') Kanban11, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 12 then KanbanTime end),'00:00') Kanban12, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 13 then KanbanTime end),'00:00') Kanban13, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 14 then KanbanTime end),'00:00') Kanban14, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 15 then KanbanTime end),'00:00') Kanban15, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 16 then KanbanTime end),'00:00') Kanban16, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 17 then KanbanTime end),'00:00') Kanban17, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 18 then KanbanTime end),'00:00') Kanban18, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 19 then KanbanTime end),'00:00') Kanban19, " & vbCrLf & _
                              " 	isnull(max(case when KanbanCycle = 20 then KanbanTime end),'00:00') Kanban20 "

            ls_SQL = ls_SQL + " from MS_KanbanTime " & vbCrLf & _
                              " Where AffiliateID= '" & Session("AffiliateID").ToString & "' " & vbCrLf & _
                              " GROUP BY AffiliateID " & vbCrLf & _
                              "  "



            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Cycle1.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban1"), 5))
                Cycle2.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban2"), 5))
                Cycle3.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban3"), 5))
                Cycle4.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban4"), 5))

                Cycle5.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban5"), 5))
                Cycle6.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban6"), 5))
                Cycle7.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban7"), 5))
                Cycle8.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban8"), 5))

                Cycle9.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban9"), 5))
                Cycle10.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban10"), 5))
                Cycle11.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban11"), 5))
                Cycle12.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban12"), 5))

                Cycle13.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban13"), 5))
                Cycle14.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban14"), 5))
                Cycle15.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban15"), 5))
                Cycle16.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban16"), 5))

                Cycle17.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban17"), 5))
                Cycle18.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban18"), 5))
                Cycle19.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban19"), 5))
                Cycle20.Text = RTrim(Left(ds.Tables(0).Rows(0)("Kanban20"), 5))
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Function AlreadyUsed(ByVal pAffiliateID As String, ByVal pCycle As String) As Boolean
        Try
            Dim ls_SQL As String = ""
            'Dim ls_MsgID As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID From MS_KanbanTime WHERE AffiliateID = '" & Session("AffiliateID").ToString & "' and KanbanCycle = '" & pCycle & "'"

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

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("KanbanTime")
                    ls_SQL = " DELETE MS_KanbanTime " & _
                        " WHERE AffiliateID='" & Session("AffiliateID").ToString & "' "

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

    Private Function ValidasiInput(ByVal pAffiliateID As String) As Boolean
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
            '        lblInfo.Text = "Affiliate ID with ID " & txtPartID.Text & " already exists in the database."
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
                         Optional ByVal pAffiliate As String = "", _
                         Optional ByVal pCycle As String = "", _
                         Optional ByVal pTime As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " SELECT AffiliateID FROM MS_KanbanTime WHERE AffiliateID= '" & Session("AffiliateID").ToString & "' and KanbanCycle = '" & pCycle & "'"

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

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("KanbanTime")

                    Dim sqlComm As New SqlCommand()

                    If pIsNewData = True Then
                        '#INSERT NEW DATA


                        ls_SQL = " INSERT INTO MS_KanbanTime " & _
                                    "(AffiliateID, KanbanCycle, KanbanTime)" & _
                                    " VALUES ('" & Session("AffiliateID").ToString & "','" & pCycle & "','" & pTime & "')" & vbCrLf
                        ls_MsgID = "1001"

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()

                        AffiliateSubmit.JSProperties("cpFunction") = "insert"
                        flag = False
                        'ElseIf pIsNewData = False And flag = True Then
                        '    ls_MsgID = "6018"
                        '    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        '    AffiliateSubmit.JSProperties("cpMessage") = lblInfo.Text
                        '    AffiliateSubmit.JSProperties("cpType") = "error"
                        '    Exit Sub

                    ElseIf pIsNewData = False Then
                        '#UPDATE DATA
                        ls_SQL = "UPDATE MS_KanbanTime SET " & _
                                "KanbanTime='" & pTime & "'" & _
                                "WHERE AffiliateID='" & Session("AffiliateID").ToString & "' and KanbanCycle='" & pCycle & "'"
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

                ls_SQL = "SELECT [KanbanCycle],[KanbanTime] " & vbCrLf & _
                         "  FROM MS_KanbanTime  " & vbCrLf & _
                         " WHERE AffiliateID = '" & Session("AffiliateID").ToString & "' " & vbCrLf
                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    cbSetData.JSProperties("cpCycle") = ds.Tables(0).Rows(0).Item("KanbanCycle").ToString()
                    cbSetData.JSProperties("cpTime") = ds.Tables(0).Rows(0).Item("KanbanTime").ToString()
                   
                    'Dim d As DateTime = DateTime.ParseExact(CDate(ds.Tables(0).Rows(0).Item("KanbanTime3").ToString()), "HHmm", Nothing)
                    'cbSetData.JSProperties("cpTime3") = d.ToString("hh:mm tt")

                Else
                    cbSetData.JSProperties("cpCycle") = ""
                    cbSetData.JSProperties("cpTime2") = ""
                    
                End If

            End Using
        Catch ex As Exception

        End Try
    End Sub
End Class