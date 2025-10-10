Imports System.Data
Imports System.Data.SqlClient

Public Class Site2
    Inherits System.Web.UI.MasterPage

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
#End Region

#Region "PROCEDURE"
    Protected Sub LoadMainMenu()
        Dim ls_SQL As String = "", ls_GroupID As String = "", ls_SubAppID As String = "", ls_SuppAppName As String = ""
        Dim iLoop As Long = 0, iLoop2 As Long = 0
        Dim iLoopParent As Long = 0, iLoopSub As Long = 0, iLoopChild As Long = 0, iLoopChild2 As Long = 0

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                With tvMenu
                    .Nodes.Clear()

                    'ROOT MENU
                    .Nodes.Add("MAIN MENU")

                    '+ PARENT MENU
                    'ls_SQL = " declare @StatAdmin char(1) " & vbCrLf & _
                    '          "  " & vbCrLf & _
                    '          " set @StatAdmin = (select StatusAdmin from SC_UserSetup where UserID = '" & Session("UserID").ToString() & "' and UserCls = '0') " & vbCrLf & _
                    '          "  " & vbCrLf & _
                    '          " if @StatAdmin = '0' " & vbCrLf & _
                    '          " 	BEGIN " & vbCrLf & _
                    '          " 		select distinct a.GroupIndex, a.GroupID from SC_UserMenu a " & vbCrLf & _
                    '          " 		inner join SC_UserPrivilege b on a.AppID = b.AppID and a.MenuID = b.MenuID " & vbCrLf & _
                    '          " 		where b.UserID = '" & Session("UserID").ToString() & "' and a.PASIMenu = '0'" & vbCrLf & _
                    '          " 		order by a.GroupIndex " & vbCrLf & _
                    '          " 	END "

                    'ls_SQL = ls_SQL + " else " & vbCrLf & _
                    '                  " 	BEGIN " & vbCrLf & _
                    '                  " 		select distinct x.GroupIndex, x.GroupID from ( " & vbCrLf & _
                    '                  " 		select a.GroupIndex, a.GroupID from SC_UserMenu a " & vbCrLf & _
                    '                  " 		inner join SC_UserPrivilege b on a.AppID = b.AppID and a.MenuID = b.MenuID " & vbCrLf & _
                    '                  " 		where b.UserID = '" & Session("UserID").ToString() & "' and a.PASIMenu = '0' and UserCls = '0'" & vbCrLf & _
                    '                  " 		union all " & vbCrLf & _
                    '                  " 		select a.GroupIndex, a.GroupID from SC_UserMenu a " & vbCrLf & _
                    '                  " 		where GroupID = 'Security System' and GroupIndex is not null)x " & vbCrLf & _
                    '                  " 		order by x.GroupIndex " & vbCrLf & _
                    '                  " 	END "

                    ls_SQL = " select distinct a.GroupIndex, a.GroupID from SC_UserMenu a " & vbCrLf & _
                              " 		inner join SC_UserPrivilege b on a.AppID = b.AppID and a.MenuID = b.MenuID " & vbCrLf & _
                              " 		where b.UserID = '" & Session("UserID").ToString() & "' and a.PASIMenu in ('0','2') and UserCls = '0' and AllowAccess = '1'" & vbCrLf & _
                              " 		order by a.GroupIndex " & vbCrLf

                    Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                    Dim dsParent As New DataSet
                    sqlDA.Fill(dsParent)

                    If dsParent.Tables(0).Rows.Count > 0 Then
                        For iLoopParent = 0 To dsParent.Tables(0).Rows.Count - 1
                            ls_SuppAppName = dsParent.Tables(0).Rows(iLoopParent).Item("GroupID").ToString()
                            ls_SubAppID = dsParent.Tables(0).Rows(iLoopParent).Item("GroupIndex").ToString()

                            .Nodes(0).Nodes.Add(ls_SuppAppName)

                            '++ SUB MENU

                            ls_SQL = " declare @StatAdmin char(1) " & vbCrLf & _
                                      " declare @GroupID char(50)" & vbCrLf & _
                                      "  " & vbCrLf & _
                                      " set @StatAdmin = (select StatusAdmin from SC_UserSetup where UserID = '" & Session("UserID").ToString() & "' and UserCls = '0') " & vbCrLf & _
                                      " set @GroupID = '" & ls_SuppAppName & "'" & vbCrLf & _
                                      "  " & vbCrLf & _
                                      " if @StatAdmin = '0' " & vbCrLf & _
                                      " 	BEGIN " & vbCrLf & _
                                      " 	    if @GroupID <> 'Security System' " & vbCrLf & _
                                      " 	        BEGIN " & vbCrLf & _
                                      " 		        SELECT a.MenuID, a.MenuDesc  " & vbCrLf & _
                                      " 		        FROM dbo.SC_UserMenu a " & vbCrLf & _
                                      " 		        INNER JOIN dbo.SC_UserPrivilege b ON a.AppID=b.AppID and a.MenuID = b.MenuID and a.PASIMenu = b.UserCls " & vbCrLf & _
                                      " 		        WHERE b.UserID = '" & Session("UserID").ToString() & "' and b.AllowAccess = '1' AND a.GroupID = @GroupID AND a.PASIMenu = '0' and UserCls = '0' and a.ShowCls = '1'" & vbCrLf & _
                                      " 		        ORDER BY a.MenuID " & vbCrLf & _
                                      " 	        END " & vbCrLf & _
                                      "  		else " & vbCrLf & _
                                      "  			begin " & vbCrLf & _
                                      "  				select distinct a.MenuID, a.MenuDesc from SC_UserMenu a  " & vbCrLf & _
                                      "                 INNER JOIN dbo.SC_UserPrivilege b ON a.AppID=b.AppID and a.MenuID = b.MenuID " & vbCrLf & _
                                      "  				where b.UserID = '" & Session("UserID").ToString() & "' and GroupID = 'Security System' and GroupIndex is not null and a.MenuID <> 'Z01' and b.UserCls = '0' and a.ShowCls = '1' order by a.MenuID" & vbCrLf & _
                                      "  			end " & vbCrLf

                            ls_SQL = ls_SQL + " 	END " & vbCrLf & _
                                              " else " & vbCrLf & _
                                              " BEGIN  " & vbCrLf & _
                                              "  		if @GroupID <> 'Security System' " & vbCrLf & _
                                              "  			begin " & vbCrLf & _
                                              "  				SELECT distinct a.MenuID, a.MenuDesc   " & vbCrLf & _
                                              "  				FROM dbo.SC_UserMenu a  " & vbCrLf & _
                                              "  				   INNER JOIN dbo.SC_UserPrivilege b ON a.AppID=b.AppID and a.MenuID = b.MenuID and a.PASIMenu = b.UserCls  " & vbCrLf & _
                                              "  				WHERE b.UserID = '" & Session("UserID").ToString() & "' and b.AllowAccess = '1' AND a.GroupID = @GroupID AND a.PASIMenu = '0' and UserCls = '0' and a.ShowCls = '1'" & vbCrLf & _
                                              "                 ORDER BY a.MenuID" & vbCrLf & _
                                              "  			end " & vbCrLf & _
                                              "  		else " & vbCrLf & _
                                              "  			begin " & vbCrLf & _
                                              "  				select distinct a.MenuID, a.MenuDesc from SC_UserMenu a  " & vbCrLf & _
                                              "                 INNER JOIN dbo.SC_UserPrivilege b ON a.AppID=b.AppID and a.MenuID = b.MenuID " & vbCrLf

                            ls_SQL = ls_SQL + "  				where GroupID = 'Security System' and GroupIndex is not null and a.MenuID <> 'Z01' and b.UserCls = '0' and a.ShowCls = '1' order by a.MenuID" & vbCrLf & _
                                              "  			end " & vbCrLf & _
                                              "  		END  "
                            sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                            Dim dsChild As New DataSet
                            sqlDA.Fill(dsChild)

                            If dsChild.Tables(0).Rows.Count > 0 Then
                                For iLoopChild = 0 To dsChild.Tables(0).Rows.Count - 1
                                    .Nodes(0).Nodes(iLoopParent).Nodes.Add("[" & dsChild.Tables(0).Rows(iLoopChild).Item("MenuID").ToString() & "] " & dsChild.Tables(0).Rows(iLoopChild).Item("MenuDesc").ToString())
                                Next iLoopChild
                            End If

                        Next iLoopParent
                    End If

                    .Nodes(0).Nodes.Add("LOGOUT")

                    tvMenu.ExpandToDepth(0)
                End With

            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Session("UserID") = "" Then Session("Msg") = "Please Login First!" : Session("GlobalURL") = Request.RawUrl : Response.Redirect("~/Login.aspx") : Exit Sub

            lblTitle.Text = "PT. AUTOCOMP SYSTEMS INDONESIA"
            lblUser.Text = "[USER ID: " & Session("UserID").ToString & "]"
            lblUserSystem.Text = "[AFFILIATE ID: " & Session("AffiliateID").ToString & "] "

            If Not Page.IsPostBack Then
                ASPxSplitter1.Panes("Menu").Visible = True
                ASPxSplitter1.Panes("Menu").Size = 370
                Call LoadMainMenu()

                ASPxSplitter1.Panes("Menu").Collapsed = False
                lblMenuID.Text = "AFFILIATE SYSTEM"
                'Page.Title = "AFFILIATE SYSTEM - MAIN MENU"
                Page.Title = "PURCHASING SYSTEM - AFFILIATE"

                Session.Remove("MenuDesc")
                'End If
            End If
        Catch ex As Exception
            If Session("Msg") = "Please Login First!" Then
                DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/Login.aspx")
                Exit Sub
            End If

        End Try
    End Sub

    Private Sub tvMenu_NodeClick(ByVal source As Object, ByVal e As DevExpress.Web.ASPxTreeView.TreeViewNodeEventArgs) Handles tvMenu.NodeClick
        Dim ls_Url As String = ""

        Try
            If e.Node.Text = "LOGOUT" Then
                Session("UserID") = ""
                Session("AffiliateID") = ""
                Session.RemoveAll()
                Response.Redirect("~/Login.aspx", False)
            Else
                Dim ls_node As String = Mid(Trim(e.Node.Text), 7, (e.Node.Text.Trim.Length) - 6)
                ls_Url = clsGlobal.GetUrl(ls_node)
                If ls_Url <> "" Then
                    Session("MenuDesc") = ls_node
                    Response.Redirect(ls_Url, False)

                    ASPxSplitter1.Panes("Menu").Collapsed = True
                    lblMenuID.Text = "[" & clsGlobal.GetMenuID(Session("MenuDesc")) & "] " & Session("MenuDesc")
                Else
                    Response.Redirect("~/MainMenu.aspx", False)
                    ASPxSplitter1.Panes("Menu").Collapsed = False
                    lblMenuID.Text = "[0000] MAIN MENU"
                End If
            End If

        Catch ex As Exception
            If Session("Msg") = "Please Login First!" Then
                DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~/Login.aspx")
                Exit Sub
            End If

            MsgBox(ex.Message.ToString)
        End Try
    End Sub
#End Region

    
End Class