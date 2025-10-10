Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Drawing
Imports System.Transactions
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO
Imports System.Data.OleDb

Public Class ForecastUploadDaily
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "E03"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim log As String = ""
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "FORM EVENTS"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_sql As String = ""
            ls_sql = "SELECT AffiliateID FROM SC_UserSetup Where UserID = '" & Session("UserID") & "' "
            Dim sqlCmd6 As New SqlCommand(ls_sql, sqlConn)
            Dim sqlDA6 As New SqlDataAdapter(sqlCmd6)
            Dim ds6 As New DataSet
            sqlDA6.Fill(ds6)

            If ds6.Tables(0).Rows.Count > 0 Then
                Session("AffiliateID") = ds6.Tables(0).Rows(0).Item("AffiliateID")
            End If
        End Using

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            If ls_AllowUpdate = False Then
                btnUpload.Enabled = False
                btnClear.Enabled = False
                btnSave.Enabled = False
                'btnDownload.Enabled = False
            End If
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 9, False, clsAppearance.PagerMode.ShowPager)
        'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("M01Url")
        Session.Remove("Period")
        Session.Remove("PONoUpload")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Uploader.NullText = "Click here to browse files..."

        lblInfo.Text = ""

        Uploader.Enabled = True
        btnSave.Enabled = True
        'btnDownload.Enabled = True
        btnUpload.Enabled = True

        up_GridLoadWhenEventChange()
        Session.Remove("Period")
        Session.Remove("PONoUpload")
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        up_Import()
    End Sub

    Private Sub ASPxCallback1_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
        Try
            Dim fi As New FileInfo(Server.MapPath("~\Kanban\Template Summary Kanban.xlsx"))
            If Not fi.Exists Then
                lblInfo.Text = "[9999] Excel Template Not Found !"
                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("Template Summary Kanban.xlsx")

            'lblInfo.Text = "[9998] Download template successful"
            'ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        End Try

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If Trim(e.GetValue("remarks")) = "" Then

        Else
            e.Cell.BackColor = Color.Red
        End If
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    Call up_Save()
                Case "kosong"
                    Call up_GridLoadWhenEventChange()
            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            'Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblInfo.Text
            Session("YA010IsSubmit") = lblInfo.Text
        End Try
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData(ByVal Part As String, ByVal Period As Date, ByVal Affiliate As String, ByVal Rev As Integer)
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            grid.VisibleColumns(5).Caption = "Forecast Quantity " & Format(Period, "MMM-yyyy")
            grid.VisibleColumns(6).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 1, Period), "MMM-yyyy")
            grid.VisibleColumns(7).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 2, Period), "MMM-yyyy")
            grid.VisibleColumns(8).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 3, Period), "MMM-yyyy")

            Dim PrevRev As Integer
            If Rev = 0 Then
                PrevRev = Rev
            Else
                PrevRev = Rev - 1
            End If
            'ls_SQL = " select  " & vbCrLf & _
            '      " 	row_number() over (order by a.SupplierID, a.PartNo asc) as no, " & vbCrLf & _
            '      " 	kanbanno, a.PartNo as partno, b.PartName as partname, a.Cycle1 as qty, a.remarks as remarks, a.supplierid as supplier, Convert(char(11),convert(date,deliverydate),120) as deliverydate, cycle = kanbancycle, direct, uom = unitcls, ISNULL(d.MOQ,0) MOQ " & vbCrLf

            'ls_SQL = ls_SQL + " from UploadKanban a  " & vbCrLf & _
            '                  " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
            '                  " left join MS_PartMapping d on a.AffiliateID = d.AffiliateID and a.SupplierID = d.SupplierID and a.PartNo = d.PartNo " & vbCrLf & _
            '                  " where a.kanbanno IN (" & kanban & ") and a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
            '                  " order by a.SupplierID, PartNo"

            ls_SQL = " Select row_number() over (order by UFD.Period, UFD.Rev, UFD.AffiliateID, UFD.PartNo asc) as no, UFD.[Period],UFD.[AffiliateID],UFD.[Rev],UFD.[PartNo],UFD.[ForecastQty1],UFD.[ForecastQty2],UFD.[ForecastQty3],UFD.[ForecastQty4]  " & vbCrLf & _
                     "       ,UFD.[F1],UFD.[F2],UFD.[F3],UFD.[F4],UFD.[F5],UFD.[F6],UFD.[F7],UFD.[F8],UFD.[F9],UFD.[F10] " & vbCrLf & _
                     "       ,UFD.[F11],UFD.[F12],UFD.[F13],UFD.[F14],UFD.[F15],UFD.[F16],UFD.[F17],UFD.[F18],UFD.[F19],UFD.[F20] " & vbCrLf & _
                     "       ,UFD.[F21],UFD.[F22],UFD.[F23],UFD.[F24],UFD.[F25],UFD.[F26],UFD.[F27],UFD.[F28],UFD.[F29],UFD.[F30],UFD.[F31],remarks " & vbCrLf & _
                     "       ,C1 = ISNULL(UFD.[F1]-FD.[F1],0),C2 = ISNULL(UFD.[F2]-FD.[F2],0),C3 = ISNULL(UFD.[F3]-FD.[F3],0),C4 = ISNULL(UFD.[F4]-FD.[F4],0),C5 = ISNULL(UFD.[F5]-FD.[F5],0),C6 = ISNULL(UFD.[F6]-FD.[F6],0),C7 = ISNULL(UFD.[F7]-FD.[F7],0),C8 = ISNULL(UFD.[F8]-FD.[F8],0),C9 = ISNULL(UFD.[F9]-FD.[F9],0),C10 = ISNULL(UFD.[F10]-FD.[F10],0) " & vbCrLf & _
                     "       ,C11 = ISNULL(UFD.[F11]-FD.[F11],0),C12 = ISNULL(UFD.[F12]-FD.[F12],0),C13 = ISNULL(UFD.[F13]-FD.[F13],0),C14 = ISNULL(UFD.[F14]-FD.[F14],0),C15 = ISNULL(UFD.[F15]-FD.[F15],0),C16 = ISNULL(UFD.[F16]-FD.[F16],0),C17 = ISNULL(UFD.[F17]-FD.[F17],0),C18 = ISNULL(UFD.[F18]-FD.[F18],0),C19 = ISNULL(UFD.[F19]-FD.[F19],0),C20 = ISNULL(UFD.[F20]-FD.[F20],0) " & vbCrLf & _
                     "       ,C21 = ISNULL(UFD.[F21]-FD.[F21],0),C22 = ISNULL(UFD.[F22]-FD.[F22],0),C23 = ISNULL(UFD.[F23]-FD.[F23],0),C24 = ISNULL(UFD.[F24]-FD.[F24],0),C25 = ISNULL(UFD.[F25]-FD.[F25],0),C26 = ISNULL(UFD.[F26]-FD.[F26],0),C27 = ISNULL(UFD.[F27]-FD.[F27],0),C28 = ISNULL(UFD.[F28]-FD.[F28],0),C29 = ISNULL(UFD.[F29]-FD.[F29],0),C30 = ISNULL(UFD.[F30]-FD.[F30],0),C31 = ISNULL(UFD.[F31]-FD.[F31],0) " & vbCrLf & _
                     " From UploadForecastDaily UFD" & vbCrLf & _
                     " Left Join ForecastDaily FD On UFD.Period = FD.Period And UFD.AffiliateID = FD.AffiliateID And UFD.PartNo = FD.PartNo And FD.Rev = '" & PrevRev & "'" & vbCrLf & _
                     " Where UFD.Period = '" & Period & "' And UFD.AffiliateID = '" & Affiliate & "' And UFD.Rev = '" & Rev & "' " & vbCrLf & _
                     " And UFD.PartNo IN (" & Part & ") "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 9, False, clsAppearance.PagerMode.ShowAllRecord)
            End With
            sqlConn.Close()

            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' no, '' kanbanno, '' PartNo, '' PartName, '' qty"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_Import()
        'Dim dt As New System.Data.DataTable
        Dim dtHeader As New System.Data.DataTable
        Dim dtDetail As New System.Data.DataTable
        Dim tempDate As Date
        Dim ls_MOQ As Double = 0
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = """"
        Dim ls_Part As String = ""
        Dim pDeleteSupplier As String = ""
        Dim pDeletePart As String = ""

        Dim connStr As String = ""

        Dim pPeriod As Date
        Dim pAffiliateID As String = Session("AffiliateID")
        Dim pRev As Integer

        'Try
        lblInfo.ForeColor = Color.Red
        If Uploader.HasFile Then
            FileName = Uploader.PostedFile.FileName
            FileExt = Path.GetExtension(Uploader.PostedFile.FileName)
            FilePath = Ext & "\Import\" & FileName
            Dim fi As New FileInfo(Server.MapPath("~\Import\" & FileName))
            If fi.Exists Then
                fi.Delete()
                fi = New FileInfo(Server.MapPath("~\Import\" & FileName))
            End If
            Uploader.SaveAs(FilePath)

            Select Case LCase(FileExt)
                Case ".xls"
                    'Excel 97-03
                    connStr = ConfigurationManager.ConnectionStrings("Excel03ConString").ConnectionString
                Case ".xlsx"
                    'Excel 07
                    connStr = ConfigurationManager.ConnectionStrings("Excel07ConString").ConnectionString
            End Select

            connStr = String.Format(connStr, FilePath, "No")

            Dim MyConnection As New OleDbConnection(connStr)
            Dim MyCommand As New OleDbCommand
            Dim MyAdapter As New OleDbDataAdapter
            MyCommand.Connection = MyConnection
            MyConnection.Open()

            Dim dtSheets As DataTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim listSheet As New List(Of String)
            Dim drSheet As DataRow

            'For Each drSheet In dtSheets.Rows
            '    listSheet.Add(drSheet("TABLE_NAME").ToString())
            'Next
            For Each drSheet In dtSheets.Rows
                If InStr("_xlnm#_FilterDatabase", drSheet("TABLE_NAME").ToString(), CompareMethod.Text) = 0 Then
                    If InStr("_xlnm#Print_Titles", drSheet("TABLE_NAME").ToString(), CompareMethod.Text) = 0 Then
                        listSheet.Add(drSheet("TABLE_NAME").ToString())
                    End If
                End If
            Next


            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ''==========Table EXCEL Master==========
                Dim pTableCode As String = listSheet(0)

                Try

                    ''DeliveryLocation
                    'ls_sql = "select DeliveryLocationCode from MS_DeliveryPlace where affiliateID = '" & Session("AffiliateID") & "' and Defaultcls = 1"

                    'Dim sqlCmd As New SqlCommand(ls_sql, sqlConn)
                    'Dim sqlDA As New SqlDataAdapter(sqlCmd)
                    'Dim ds As New DataSet
                    'sqlDA.Fill(ds)

                    'If ds.Tables(0).Rows.Count > 0 Then
                    '    Session("DeliveryLoc") = ds.Tables(0).Rows(0)("DeliveryLocationCode")
                    'Else
                    '    lblInfo.Text = "[9999] Invalid Delivery Location, please select default!"
                    '    grid.JSProperties("cpMessage") = lblInfo.Text
                    '    Exit Sub
                    'End If


                    'ls_sql = " select * from ms_kanbantime where affiliateID = '" & Session("AffiliateID") & "'"
                    'Dim sqlc As New SqlCommand(ls_sql, sqlConn)
                    'Dim sqlDC As New SqlDataAdapter(sqlc)
                    'Dim dsC As New DataSet
                    'sqlDC.Fill(dsC)

                    'If dsC.Tables(0).Rows.Count > 0 Then
                    '    If dsC.Tables(0).Rows(0)("kanbanCycle") = "1" Then
                    '        Session("KanbanTime1") = dsC.Tables(0).Rows(0)("kanbanTime")
                    '    End If

                    '    If dsC.Tables(0).Rows(1)("kanbanCycle") = "2" Then
                    '        Session("KanbanTime2") = dsC.Tables(0).Rows(1)("kanbanTime")
                    '    End If

                    '    If dsC.Tables(0).Rows(2)("kanbanCycle") = "3" Then
                    '        Session("KanbanTime3") = dsC.Tables(0).Rows(2)("kanbanTime")
                    '    End If

                    '    If dsC.Tables(0).Rows(3)("kanbanCycle") = "4" Then
                    '        Session("KanbanTime4") = dsC.Tables(0).Rows(3)("kanbanTime")
                    '    End If
                    'Else
                    '    Session("KanbanTime1") = "00:00"
                    '    Session("KanbanTime2") = "00:00"
                    '    Session("KanbanTime3") = "00:00"
                    '    Session("KanbanTime4") = "00:00"

                    'End If


                    'End If

                    'Get Header Data
                    MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A4:C6]")
                    MyAdapter.SelectCommand = MyCommand
                    MyAdapter.Fill(dtHeader)

                    If dtHeader.Rows.Count > 0 Then
                        'PERIOD
                        If IsDBNull(dtHeader.Rows(0).Item(2)) Then
                            lblInfo.Text = "[9999] Please input Period, please check the file again!"
                            grid.JSProperties("cpMessage") = lblInfo.Text
                            MyConnection.Close()
                            Exit Sub
                        Else
                            Try
                                pPeriod = "01-" & dtHeader.Rows(0).Item(2)
                            Catch ex As Exception
                                lblInfo.Text = "[9999] Invalid Format Period, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End Try
                        End If

                        ''AFFILIATE
                        'If IsDBNull(dtHeader.Rows(1).Item(2)) Then
                        '    lblInfo.Text = "[9999] Please input Affiliate Code!"
                        '    grid.JSProperties("cpMessage") = lblInfo.Text
                        '    MyConnection.Close()
                        '    Exit Sub
                        'Else
                        '    '03.1 Check Affiliate di PO MASTER
                        '    ls_sql = "SELECT * FROM MS_Affiliate WHERE AffiliateID = '" & dtHeader.Rows(1).Item(2) & "'"
                        '    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn)
                        '    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                        '    Dim ds5 As New DataSet
                        '    sqlDA5.Fill(ds5)

                        '    If ds5.Tables(0).Rows.Count = 0 Then
                        '        lblInfo.Text = "[9999] Affiliate Code not valid, please check the file again!"
                        '        grid.JSProperties("cpMessage") = lblInfo.Text
                        '        MyConnection.Close()
                        '        Exit Sub
                        '    End If

                        '    pAffiliateID = dtHeader.Rows(1).Item(2)
                        'End If

                        'REVISION
                        If IsDBNull(dtHeader.Rows(2).Item(2)) Then
                            lblInfo.Text = "[9999] Please input Revision, please check the file again!"
                            grid.JSProperties("cpMessage") = lblInfo.Text
                            MyConnection.Close()
                            Exit Sub
                        Else
                            Try
                                pRev = dtHeader.Rows(2).Item(2)
                                If pRev > "3" Then
                                    lblInfo.Text = "[9999] Revision > 3, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                lblInfo.Text = "[9999] Invalid Format Revision, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End Try
                        End If

                        'Check Revision yang sudah pernah diinput
                        ls_sql = "SELECT DISTINCT Rev FROM ForecastDaily WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & pAffiliateID & "' AND Rev = '" & pRev & "' " & vbCrLf & _
                                 "Order By Rev Desc"
                        Dim sqlCmd6 As New SqlCommand(ls_sql, sqlConn)
                        Dim sqlDA6 As New SqlDataAdapter(sqlCmd6)
                        Dim ds6 As New DataSet
                        sqlDA6.Fill(ds6)

                        If ds6.Tables(0).Rows.Count > 0 Then
                            If ds6.Tables(0).Rows(0).Item("Rev") = pRev Then
                                lblInfo.Text = "[9999] Rev " & pRev & " Already Upload, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End If
                        End If

                        'Check Urutan Revision
                        ls_sql = "SELECT DISTINCT Rev FROM ForecastDaily WHERE Period = '" & pPeriod & "' AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                 "Order By Rev Desc"
                        Dim sqlCmd7 As New SqlCommand(ls_sql, sqlConn)
                        Dim sqlDA7 As New SqlDataAdapter(sqlCmd7)
                        Dim ds7 As New DataSet
                        sqlDA7.Fill(ds7)

                        If ds7.Tables(0).Rows.Count > 0 Then
                            If ds7.Tables(0).Rows(0).Item("Rev") + 1 <> pRev Then
                                lblInfo.Text = "[9999] Please Upload Previous Rev First, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End If
                        Else
                            'Pertama Upload Harus Rev 0
                            If pRev > 0 Then
                                If ds7.Tables(0).Rows(0).Item("Rev") + 1 <> pRev Then
                                    lblInfo.Text = "[9999] Please Upload Rev 0 First, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If

                            End If
                        End If

                    End If

                    Dim dtUploadDetailList As New List(Of clsForecast)

                    'Get Detail Data
                    MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A11:AL65536]")
                    MyAdapter.SelectCommand = MyCommand
                    MyAdapter.Fill(dtDetail)

                    If dtDetail.Rows.Count > 0 Then
                        For i = 0 To dtDetail.Rows.Count - 1
                            If CStr(IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))) <> "" Then
                                Dim dtUploadDetail As New clsForecast
                                dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(1)
                                If pDeletePart = "" Then
                                    pDeletePart = "'" + Trim(dtUploadDetail.D_PartNo) + "'"
                                Else
                                    pDeletePart = pDeletePart + ",'" + Trim(dtUploadDetail.D_PartNo) + "'"
                                End If

                                dtUploadDetail.D_Period = pPeriod
                                dtUploadDetail.D_AffiliateID = pAffiliateID
                                dtUploadDetail.D_Rev = pRev

                                'dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(1)
                                'dtUploadDetail.D_ForecastQty1 = IIf(IsDBNull(dtDetail.Rows(i).Item(2)) = True, 0, dtDetail.Rows(i).Item(2))
                                dtUploadDetail.D_ForecastQty2 = IIf(IsDBNull(dtDetail.Rows(i).Item(3)) = True, 0, dtDetail.Rows(i).Item(3))
                                dtUploadDetail.D_ForecastQty3 = IIf(IsDBNull(dtDetail.Rows(i).Item(4)) = True, 0, dtDetail.Rows(i).Item(4))
                                dtUploadDetail.D_ForecastQty4 = IIf(IsDBNull(dtDetail.Rows(i).Item(5)) = True, 0, dtDetail.Rows(i).Item(5))
                                dtUploadDetail.D_F1 = IIf(IsDBNull(dtDetail.Rows(i).Item(6)) = True, 0, dtDetail.Rows(i).Item(6))
                                dtUploadDetail.D_F2 = IIf(IsDBNull(dtDetail.Rows(i).Item(7)) = True, 0, dtDetail.Rows(i).Item(7))
                                dtUploadDetail.D_F3 = IIf(IsDBNull(dtDetail.Rows(i).Item(8)) = True, 0, dtDetail.Rows(i).Item(8))
                                dtUploadDetail.D_F4 = IIf(IsDBNull(dtDetail.Rows(i).Item(9)) = True, 0, dtDetail.Rows(i).Item(9))
                                dtUploadDetail.D_F5 = IIf(IsDBNull(dtDetail.Rows(i).Item(10)) = True, 0, dtDetail.Rows(i).Item(10))
                                dtUploadDetail.D_F6 = IIf(IsDBNull(dtDetail.Rows(i).Item(11)) = True, 0, dtDetail.Rows(i).Item(11))
                                dtUploadDetail.D_F7 = IIf(IsDBNull(dtDetail.Rows(i).Item(12)) = True, 0, dtDetail.Rows(i).Item(12))
                                dtUploadDetail.D_F8 = IIf(IsDBNull(dtDetail.Rows(i).Item(13)) = True, 0, dtDetail.Rows(i).Item(13))
                                dtUploadDetail.D_F9 = IIf(IsDBNull(dtDetail.Rows(i).Item(14)) = True, 0, dtDetail.Rows(i).Item(14))
                                dtUploadDetail.D_F10 = IIf(IsDBNull(dtDetail.Rows(i).Item(15)) = True, 0, dtDetail.Rows(i).Item(15))
                                dtUploadDetail.D_F11 = IIf(IsDBNull(dtDetail.Rows(i).Item(16)) = True, 0, dtDetail.Rows(i).Item(16))
                                dtUploadDetail.D_F12 = IIf(IsDBNull(dtDetail.Rows(i).Item(17)) = True, 0, dtDetail.Rows(i).Item(17))
                                dtUploadDetail.D_F13 = IIf(IsDBNull(dtDetail.Rows(i).Item(18)) = True, 0, dtDetail.Rows(i).Item(18))
                                dtUploadDetail.D_F14 = IIf(IsDBNull(dtDetail.Rows(i).Item(19)) = True, 0, dtDetail.Rows(i).Item(19))
                                dtUploadDetail.D_F15 = IIf(IsDBNull(dtDetail.Rows(i).Item(20)) = True, 0, dtDetail.Rows(i).Item(20))
                                dtUploadDetail.D_F16 = IIf(IsDBNull(dtDetail.Rows(i).Item(21)) = True, 0, dtDetail.Rows(i).Item(21))
                                dtUploadDetail.D_F17 = IIf(IsDBNull(dtDetail.Rows(i).Item(22)) = True, 0, dtDetail.Rows(i).Item(22))
                                dtUploadDetail.D_F18 = IIf(IsDBNull(dtDetail.Rows(i).Item(23)) = True, 0, dtDetail.Rows(i).Item(23))
                                dtUploadDetail.D_F19 = IIf(IsDBNull(dtDetail.Rows(i).Item(24)) = True, 0, dtDetail.Rows(i).Item(24))
                                dtUploadDetail.D_F20 = IIf(IsDBNull(dtDetail.Rows(i).Item(25)) = True, 0, dtDetail.Rows(i).Item(25))
                                dtUploadDetail.D_F21 = IIf(IsDBNull(dtDetail.Rows(i).Item(26)) = True, 0, dtDetail.Rows(i).Item(26))
                                dtUploadDetail.D_F22 = IIf(IsDBNull(dtDetail.Rows(i).Item(27)) = True, 0, dtDetail.Rows(i).Item(27))
                                dtUploadDetail.D_F23 = IIf(IsDBNull(dtDetail.Rows(i).Item(28)) = True, 0, dtDetail.Rows(i).Item(28))
                                dtUploadDetail.D_F24 = IIf(IsDBNull(dtDetail.Rows(i).Item(29)) = True, 0, dtDetail.Rows(i).Item(29))
                                dtUploadDetail.D_F25 = IIf(IsDBNull(dtDetail.Rows(i).Item(30)) = True, 0, dtDetail.Rows(i).Item(30))
                                dtUploadDetail.D_F26 = IIf(IsDBNull(dtDetail.Rows(i).Item(31)) = True, 0, dtDetail.Rows(i).Item(31))
                                dtUploadDetail.D_F27 = IIf(IsDBNull(dtDetail.Rows(i).Item(32)) = True, 0, dtDetail.Rows(i).Item(32))
                                dtUploadDetail.D_F28 = IIf(IsDBNull(dtDetail.Rows(i).Item(33)) = True, 0, dtDetail.Rows(i).Item(33))
                                dtUploadDetail.D_F29 = IIf(IsDBNull(dtDetail.Rows(i).Item(34)) = True, 0, dtDetail.Rows(i).Item(34))
                                dtUploadDetail.D_F30 = IIf(IsDBNull(dtDetail.Rows(i).Item(35)) = True, 0, dtDetail.Rows(i).Item(35))
                                dtUploadDetail.D_F31 = IIf(IsDBNull(dtDetail.Rows(i).Item(36)) = True, 0, dtDetail.Rows(i).Item(36))

                                dtUploadDetail.D_ForecastQty1 = dtUploadDetail.D_F1 + dtUploadDetail.D_F2 + dtUploadDetail.D_F3 + dtUploadDetail.D_F4 + dtUploadDetail.D_F5 _
                                                              + dtUploadDetail.D_F6 + dtUploadDetail.D_F7 + dtUploadDetail.D_F8 + dtUploadDetail.D_F9 + dtUploadDetail.D_F10 _
                                                              + dtUploadDetail.D_F11 + dtUploadDetail.D_F12 + dtUploadDetail.D_F13 + dtUploadDetail.D_F14 + dtUploadDetail.D_F15 _
                                                              + dtUploadDetail.D_F16 + dtUploadDetail.D_F17 + dtUploadDetail.D_F18 + dtUploadDetail.D_F19 + dtUploadDetail.D_F20 _
                                                              + dtUploadDetail.D_F21 + dtUploadDetail.D_F22 + dtUploadDetail.D_F23 + dtUploadDetail.D_F24 + dtUploadDetail.D_F25 _
                                                              + dtUploadDetail.D_F26 + dtUploadDetail.D_F27 + dtUploadDetail.D_F28 + dtUploadDetail.D_F29 + dtUploadDetail.D_F30 _
                                                              + dtUploadDetail.D_F31


                                dtUploadDetailList.Add(dtUploadDetail)
                            Else
                                Exit For
                            End If
                        Next
                    End If

                    Dim ls_TempSupplierID As String = ""
                    Dim ls_DoubleSupplier As Boolean = False
                    Dim ls_supp As String = ""
                    Dim countSupplier As Integer = 0
                    Dim OLDSupp As String = ""
                    Dim OLDKanban As String = ""

                    Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")

                        '01.01 Delete TempoaryData

                        ls_sql = "delete UploadForecastDaily where AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                 " and Period = '" & pPeriod & "' " & vbCrLf & _
                                 " and Rev = '" & pRev & "' " & vbCrLf & _
                                 " and PartNo IN (" & pDeletePart & ") " & vbCrLf

                        Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        sqlComm9.ExecuteNonQuery()
                        sqlComm9.Dispose()


                        '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                        For i = 0 To dtUploadDetailList.Count - 1
                            Dim ls_error As String = ""
                            Dim Part As clsForecast = dtUploadDetailList(i)
                            'Dim ls_Qty As Integer


                            '02.1 Check PartNo di MS_Part
                            ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & Part.D_PartNo & "' "
                            Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                            Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                            Dim ds2 As New DataSet
                            sqlDA2.Fill(ds2)

                            If ds2.Tables(0).Rows.Count = 0 Then
                                ls_error = "PartNo not found in Part Master, please check again with PASI!"
                            End If

                            '02.2 Check PartNo di Ms_PartMapping
                            ls_sql = "select * from ms_partmapping WHERE PartNo = '" & Part.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                            Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                            Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                            Dim ds3 As New DataSet
                            sqlDA3.Fill(ds3)

                            If ds3.Tables(0).Rows.Count = 0 Then
                                If ls_error = "" Then
                                    ls_error = "PartNo not found in Part Mapping, please check again with PASI!"
                                End If
                            Else
                                'ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")

                                'ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                'ls_Qty = 0
                                'If CDbl(Kanban.D_c1) <> 0 Then ls_Qty = CDbl(Kanban.D_c1)

                                'If (ls_Qty Mod ls_MOQ) <> 0 And ls_Qty <> 0 Then
                                '    If ls_error = "" Then
                                '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                '    End If
                                'End If
                            End If

                            ls_sql = " INSERT INTO [dbo].[UploadForecastDaily] " & vbCrLf & _
                                      "            ([Period],[AffiliateID],[Rev],[PartNo],[ForecastQty1],[ForecastQty2],[ForecastQty3],[ForecastQty4] " & vbCrLf & _
                                      "             ,[F1],[F2],[F3],[F4],[F5],[F6],[F7],[F8],[F9],[F10] " & vbCrLf & _
                                      "             ,[F11],[F12],[F13],[F14],[F15],[F16],[F17],[F18],[F19],[F20] " & vbCrLf & _
                                      "             ,[F21],[F22],[F23],[F24],[F25],[F26],[F27],[F28],[F29],[F30],[F31],remarks) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & Part.D_Period & "' " & vbCrLf & _
                                      "            ,'" & Part.D_AffiliateID & "' " & vbCrLf & _
                                      "            ,'" & Part.D_Rev & "' " & vbCrLf & _
                                      "            ,'" & Part.D_PartNo & "' " & vbCrLf & _
                                      "            ,'" & Part.D_ForecastQty1 & "' " & vbCrLf & _
                                      "            ,'" & Part.D_ForecastQty2 & "' " & vbCrLf & _
                                      "            ,'" & Part.D_ForecastQty3 & "' " & vbCrLf & _
                                      "            ,'" & Part.D_ForecastQty4 & "' " & vbCrLf

                            ls_sql = ls_sql + "            ,'" & Part.D_F1 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F2 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F3 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F4 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F5 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F6 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F7 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F8 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F9 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F10 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F11 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F12 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F13 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F14 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F15 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F16 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F17 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F18 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F19 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F20 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F21 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F22 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F23 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F24 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F25 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F26 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F27 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F28 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F29 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F30 & "' " & vbCrLf & _
                                              "            ,'" & Part.D_F31 & "' " & vbCrLf & _
                                              "            ,'" & ls_error & "') " & vbCrLf
                            Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                            If i = 0 Then
                                ls_Part = "'" & Part.D_PartNo & "'"
                            End If

                            If Trim(ls_Part) <> Trim(Part.D_PartNo) Then
                                ls_Part = ls_Part + ",'" + Part.D_PartNo + "'"
                            End If

                        Next
                        sqlTran.Commit()

                        'Session("DeliveryDate") = dtUploadHeader.H_DeliveryDate
                        'Session("PONoUpload") = dtUploadHeader.H_kanbanNo
                        'Session("FilterKanbanNoNew") = ls_TempKanban

                        lblInfo.Text = "[7001] Data Checking Done!"
                        lblInfo.ForeColor = Color.Blue
                        grid.JSProperties("cpMessage") = lblInfo.Text


                        Call bindData(ls_Part, pPeriod, pAffiliateID, pRev)
                    End Using
                Catch ex As Exception
                    MyConnection.Close()
                    lblInfo.Text = ex.Message
                    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                    If InStr(ex.Message, "Cannot find column") = 1 Then
                        lblInfo.Text = "Format Template Tidak Sesuai"
                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, lblInfo.Text.ToString())
                    End If
                    'Exit Sub
                Finally
                    MyConnection.Close()
                    sqlConn.Close()
                End Try
                'dt.Reset()
                dtDetail.Reset()
                dtHeader.Reset()
            End Using
            MyConnection.Close()
        Else
            If FileName = "" Then
                lblInfo.Text = "[9999] Please choose the file!"
                up_GridLoadWhenEventChange()
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If
        End If
        'Catch ex As Exception
        '    MyConnection.Close()
        '    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        'End Try
    End Sub

    Protected Sub Uploader_FileUploadComplete(ByVal sender As Object, ByVal e As FileUploadCompleteEventArgs)
        Try
            e.CallbackData = SavePostedFiles(e.UploadedFile)
        Catch ex As Exception
            e.IsValid = False
            lblInfo.Text = ex.Message
        End Try
    End Sub

    Private Function SavePostedFiles(ByVal uploadedFile As UploadedFile) As String
        If (Not uploadedFile.IsValid) Then
            Return String.Empty
        End If

        Ext = Path.Combine(MapPath(""))
        FileName = Uploader.PostedFile.FileName
        FilePath = Ext & "\Import\" & FileName
        uploadedFile.SaveAs(FilePath)

        Return FilePath
    End Function

    Private Sub up_Save()
        Dim i As Integer, j As Integer, x As Integer
        Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        Dim ls_PONo As String = ""
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        'Dim ls_Period As Date
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim ls_DoubleSupplier As Boolean = False
        Dim ls_TempSupplierID As String = ""
        Dim ls_TempKanban As String = ""
        'Dim ls_KanbanNo As String
        Dim ls_cycle As Integer
        Dim ls_qty As Integer
        Dim ls_Date As String
        Dim ls_time As String
        Dim ls_seq As Integer
        'Dim ls_KanbanDate As Date

        Dim ls_Period As Date
        Dim ls_Rev As Integer
        Dim ls_Affiliate As String = ""
        Dim ls_TempPart As String = ""
        Dim ls_Part As String = ""

        Dim C1 As Integer
        Dim C2 As Integer
        Dim C3 As Integer
        Dim C4 As Integer
        Dim C5 As Integer
        Dim C6 As Integer
        Dim C7 As Integer
        Dim C8 As Integer
        Dim C9 As Integer
        Dim C10 As Integer
        Dim C11 As Integer
        Dim C12 As Integer
        Dim C13 As Integer
        Dim C14 As Integer
        Dim C15 As Integer
        Dim C16 As Integer
        Dim C17 As Integer
        Dim C18 As Integer
        Dim C19 As Integer
        Dim C20 As Integer
        Dim C21 As Integer
        Dim C22 As Integer
        Dim C23 As Integer
        Dim C24 As Integer
        Dim C25 As Integer
        Dim C26 As Integer
        Dim C27 As Integer
        Dim C28 As Integer
        Dim C29 As Integer
        Dim C30 As Integer
        Dim C31 As Integer


        Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
        Dim SqlTran As SqlTransaction

        SqlCon.Open()

        SqlTran = SqlCon.BeginTransaction

        Try
            If grid.VisibleRowCount = 0 Then Exit Sub
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "remarks").ToString.Trim <> "" Then
                    ls_Check = True
                    Exit For
                End If
            Next i

            Dim countSupplier As Integer = 0

            For i = 0 To grid.VisibleRowCount - 1
                ls_Period = grid.GetRowValues(0, "Period").ToString.Trim
                ls_Rev = grid.GetRowValues(0, "Rev").ToString.Trim
                ls_Affiliate = grid.GetRowValues(0, "AffiliateID").ToString.Trim

                If i = 0 Then
                    'ls_TempSupplierID = "'" & grid.GetRowValues(i, "supplier").ToString.Trim & "'"
                    'ls_TempKanban = "'" & grid.GetRowValues(i, "kanbanno").ToString.Trim & "'"
                    'countSupplier = 1
                    ls_TempPart = "'" & grid.GetRowValues(i, "PartNo").ToString.Trim & "'"
                End If

                'If ls_TempSupplierID <> grid.GetRowValues(i, "supplier").ToString.Trim Then
                '    ls_DoubleSupplier = True
                '    ls_TempSupplierID = ls_TempSupplierID + ",'" + grid.GetRowValues(i, "supplier").ToString.Trim + "'"
                '    countSupplier = countSupplier + 1
                'End If

                'If Trim(ls_TempKanban) <> Trim(grid.GetRowValues(i, "kanbanno").ToString.Trim) Then
                '    ls_TempKanban = ls_TempKanban + ",'" + grid.GetRowValues(i, "kanbanno").ToString.Trim + "'"
                'End If
                If Trim(ls_TempPart) <> Trim(grid.GetRowValues(i, "PartNo").ToString.Trim) Then
                    ls_TempPart = ls_TempPart + ",'" + grid.GetRowValues(i, "PartNo").ToString.Trim + "'"
                End If
            Next i

            If ls_Check = True Then
                lblInfo.Text = "[9999] Invalid data in this File Upload, please check the file again!"
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            'dian
            '1. Delete Data
            Dim SQLCom As SqlCommand = SqlCon.CreateCommand
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran
            'Dim ls_KanbanAsli As String = Trim(Session("PONoUpload"))
            'Dim ls_deliveryDate As String = Session("DeliveryDate")

            ls_Sql = "Delete ForecastDaily where Period = '" & ls_Period & "' and Rev = '" & ls_Rev & "' and AffiliateID = '" & ls_Affiliate & "' And PartNo IN (" & ls_TempPart & ")"
            'ls_Sql = "Delete Kanban_Detail where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            'ls_Sql = ls_Sql + "Delete Kanban_Master where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            'ls_Sql = ls_Sql + "Delete Kanban_Barcode where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            '2. Insert New Data
            'ls_KanbanNo = ""

            For i = 0 To grid.VisibleRowCount - 1
                ls_Period = Trim(grid.GetRowValues(i, "Period"))
                ls_Rev = Trim(grid.GetRowValues(i, "Rev"))
                ls_Affiliate = Trim(grid.GetRowValues(i, "AffiliateID"))
                ls_Part = Trim(grid.GetRowValues(i, "PartNo"))

                C1 = IIf(Trim(grid.GetRowValues(i, "C1")) = 0, 0, 1)
                C2 = IIf(Trim(grid.GetRowValues(i, "C2")) = 0, 0, 1)
                C3 = IIf(Trim(grid.GetRowValues(i, "C3")) = 0, 0, 1)
                C4 = IIf(Trim(grid.GetRowValues(i, "C4")) = 0, 0, 1)
                C5 = IIf(Trim(grid.GetRowValues(i, "C5")) = 0, 0, 1)
                C6 = IIf(Trim(grid.GetRowValues(i, "C6")) = 0, 0, 1)
                C7 = IIf(Trim(grid.GetRowValues(i, "C7")) = 0, 0, 1)
                C8 = IIf(Trim(grid.GetRowValues(i, "C8")) = 0, 0, 1)
                C9 = IIf(Trim(grid.GetRowValues(i, "C9")) = 0, 0, 1)
                C10 = IIf(Trim(grid.GetRowValues(i, "C10")) = 0, 0, 1)
                C11 = IIf(Trim(grid.GetRowValues(i, "C11")) = 0, 0, 1)
                C12 = IIf(Trim(grid.GetRowValues(i, "C12")) = 0, 0, 1)
                C13 = IIf(Trim(grid.GetRowValues(i, "C13")) = 0, 0, 1)
                C14 = IIf(Trim(grid.GetRowValues(i, "C14")) = 0, 0, 1)
                C15 = IIf(Trim(grid.GetRowValues(i, "C15")) = 0, 0, 1)
                C16 = IIf(Trim(grid.GetRowValues(i, "C16")) = 0, 0, 1)
                C17 = IIf(Trim(grid.GetRowValues(i, "C17")) = 0, 0, 1)
                C18 = IIf(Trim(grid.GetRowValues(i, "C18")) = 0, 0, 1)
                C19 = IIf(Trim(grid.GetRowValues(i, "C19")) = 0, 0, 1)
                C20 = IIf(Trim(grid.GetRowValues(i, "C20")) = 0, 0, 1)
                C21 = IIf(Trim(grid.GetRowValues(i, "C21")) = 0, 0, 1)
                C22 = IIf(Trim(grid.GetRowValues(i, "C22")) = 0, 0, 1)
                C23 = IIf(Trim(grid.GetRowValues(i, "C23")) = 0, 0, 1)
                C24 = IIf(Trim(grid.GetRowValues(i, "C24")) = 0, 0, 1)
                C25 = IIf(Trim(grid.GetRowValues(i, "C25")) = 0, 0, 1)
                C26 = IIf(Trim(grid.GetRowValues(i, "C26")) = 0, 0, 1)
                C27 = IIf(Trim(grid.GetRowValues(i, "C27")) = 0, 0, 1)
                C28 = IIf(Trim(grid.GetRowValues(i, "C28")) = 0, 0, 1)
                C29 = IIf(Trim(grid.GetRowValues(i, "C29")) = 0, 0, 1)
                C30 = IIf(Trim(grid.GetRowValues(i, "C30")) = 0, 0, 1)
                C31 = IIf(Trim(grid.GetRowValues(i, "C31")) = 0, 0, 1)

                ls_Sql = " INSERT INTO [dbo].[ForecastDaily] " & vbCrLf & _
                                      "            ([Period],[AffiliateID],[Rev],[PartNo],[ForecastQty1],[ForecastQty2],[ForecastQty3],[ForecastQty4] " & vbCrLf & _
                                      "             ,[F1],[F2],[F3],[F4],[F5],[F6],[F7],[F8],[F9],[F10] " & vbCrLf & _
                                      "             ,[F11],[F12],[F13],[F14],[F15],[F16],[F17],[F18],[F19],[F20] " & vbCrLf & _
                                      "             ,[F21],[F22],[F23],[F24],[F25],[F26],[F27],[F28],[F29],[F30],[F31],EntryDate,EntryUser " & vbCrLf & _
                                      "             ,[C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10] " & vbCrLf & _
                                      "             ,[C11],[C12],[C13],[C14],[C15],[C16],[C17],[C18],[C19],[C20] " & vbCrLf & _
                                      "             ,[C21],[C22],[C23],[C24],[C25],[C26],[C27],[C28],[C29],[C30],[C31]) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & ls_Period & "' " & vbCrLf & _
                                      "            ,'" & ls_Affiliate & "' " & vbCrLf & _
                                      "            ,'" & ls_Rev & "' " & vbCrLf & _
                                      "            ,'" & ls_Part & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "ForecastQty1") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "ForecastQty2") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "ForecastQty3") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "ForecastQty4") & "' " & vbCrLf

                ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "F1") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F2") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F3") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F4") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F5") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F6") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F7") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F8") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F9") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F10") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F11") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F12") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F13") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F14") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F15") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F16") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F17") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F18") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F19") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F20") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F21") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F22") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F23") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F24") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F25") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F26") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F27") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F28") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F29") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F30") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "F31") & "' " & vbCrLf & _
                                  "            ,GETDATE(),'" & Session("UserID").ToString & "' " & vbCrLf & _
                                  "            ,'" & C1 & "' " & vbCrLf & _
                                  "            ,'" & C2 & "' " & vbCrLf & _
                                  "            ,'" & C3 & "' " & vbCrLf & _
                                  "            ,'" & C4 & "' " & vbCrLf & _
                                  "            ,'" & C5 & "' " & vbCrLf & _
                                  "            ,'" & C6 & "' " & vbCrLf & _
                                  "            ,'" & C7 & "' " & vbCrLf & _
                                  "            ,'" & C8 & "' " & vbCrLf & _
                                  "            ,'" & C9 & "' " & vbCrLf & _
                                  "            ,'" & C10 & "' " & vbCrLf & _
                                  "            ,'" & C11 & "' " & vbCrLf & _
                                  "            ,'" & C12 & "' " & vbCrLf & _
                                  "            ,'" & C13 & "' " & vbCrLf & _
                                  "            ,'" & C14 & "' " & vbCrLf & _
                                  "            ,'" & C15 & "' " & vbCrLf & _
                                  "            ,'" & C16 & "' " & vbCrLf & _
                                  "            ,'" & C17 & "' " & vbCrLf & _
                                  "            ,'" & C18 & "' " & vbCrLf & _
                                  "            ,'" & C19 & "' " & vbCrLf & _
                                  "            ,'" & C20 & "' " & vbCrLf & _
                                  "            ,'" & C21 & "' " & vbCrLf & _
                                  "            ,'" & C22 & "' " & vbCrLf & _
                                  "            ,'" & C23 & "' " & vbCrLf & _
                                  "            ,'" & C24 & "' " & vbCrLf & _
                                  "            ,'" & C25 & "' " & vbCrLf & _
                                  "            ,'" & C26 & "' " & vbCrLf & _
                                  "            ,'" & C27 & "' " & vbCrLf & _
                                  "            ,'" & C28 & "' " & vbCrLf & _
                                  "            ,'" & C29 & "' " & vbCrLf & _
                                  "            ,'" & C30 & "' " & vbCrLf & _
                                  "            ,'" & C31 & "')" & vbCrLf


                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()


                ls_MsgID = "1001"
                ls_Detail = "ada"
            Next

            'delete tada di tempolary table
            'ls_Sql = "delete UploadKanban where AffiliateID = '" & Session("AffiliateID") & "' and KanbanNo IN (" & Trim(ls_TempKanban) & ")"
            ls_Sql = "Delete UploadForecastDaily where Period = '" & ls_Period & "' and Rev = '" & ls_Rev & "' and AffiliateID = '" & ls_Affiliate & "' And PartNo IN (" & ls_TempPart & ")"

            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()


            '2.3.3 Commit transaction
            SqlTran.Commit()
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblInfo.Text
            Session("YA010IsSubmit") = lblInfo.Text

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("YA010IsSubmit") = lblInfo.Text
            SqlTran.Rollback()
            SqlCon.Close()
        Finally
            SqlCon.Close()
        End Try
    End Sub
#End Region
End Class