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

Public Class KanbanUpload
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

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            If ls_AllowUpdate = False Then
                btnUpload.Enabled = False
                btnClear.Enabled = False
                btnSave.Enabled = False
                btnDownload.Enabled = False
            End If
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
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
        btnDownload.Enabled = True
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
        End Try
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	row_number() over (order by a.SupplierID, a.PartNo asc) as NoUrut, " & vbCrLf & _
                  " 	DeliveryDate as date, Shipby as shipby, a.PartNo as partno, b.PartName as partname, b.unitcls as uom, b.moq as moq, a.supplierID as supplier, a.cycle1 as c1, " & vbCrLf & _
                  " 	a.cycle2 as c2, a.cycle3 as c3, a.cycle4 as c4, KanbanNo as kanban, a.remarks as remarks " & vbCrLf

            ls_SQL = ls_SQL + " from UploadKanban a  " & vbCrLf & _
                              " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " --left join MS_PartMapping d on a.AffiliateID = d.AffiliateID and a.PartNo = d.PartNo " & vbCrLf & _
                              " where a.kanbanno = '" & Session("PONoUpload") & "' and a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " order by a.SupplierID, PartNo"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' Period, '' PONo, '' ShipBy, '' PartNo, '' PartName, '' UnitDesc, '' MOQ, '' Maker, " & vbCrLf & _
                  " '' Project, '' SupplierID, 0 POQty, 0 ForecastN1, 0 ForecastN2, 0 ForecastN3,   " & vbCrLf & _
                  " 0 DeliveryD1, 0 DeliveryD2, 0 DeliveryD3, 0 DeliveryD4, 0 DeliveryD5, " & vbCrLf & _
                  " 0 DeliveryD6, 0 DeliveryD7, 0 DeliveryD8, 0 DeliveryD9, 0 DeliveryD10, " & vbCrLf & _
                  " 0 DeliveryD11, 0 DeliveryD12, 0 DeliveryD13, 0 DeliveryD14, 0 DeliveryD15, " & vbCrLf & _
                  " 0 DeliveryD16, 0 DeliveryD17, 0 DeliveryD18, 0 DeliveryD19, 0 DeliveryD20, " & vbCrLf & _
                  " 0 DeliveryD21, 0 DeliveryD22, 0 DeliveryD23, 0 DeliveryD24, 0 DeliveryD25, " & vbCrLf & _
                  " 0 DeliveryD26, 0 DeliveryD27, 0 DeliveryD28, 0 DeliveryD29, 0 DeliveryD30, " & vbCrLf & _
                  " 0 DeliveryD31, '' ErrorCls"

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
        Dim dt As New System.Data.DataTable
        Dim dtHeader As New System.Data.DataTable
        Dim dtDetail As New System.Data.DataTable
        Dim tempDate As Date
        Dim ls_MOQ As Double = 0
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = """"

        Session.Remove("DeliveryDate")
        Session.Remove("PONoUpload")

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

            Dim connStr As String = ""
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

            Try
                Dim dtSheets As DataTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                Dim listSheet As New List(Of String)
                Dim drSheet As DataRow

                For Each drSheet In dtSheets.Rows
                    listSheet.Add(drSheet("TABLE_NAME").ToString())
                Next

                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    ''==========Table EXCEL Master==========
                    Dim pTableCode As String = listSheet(0)

                    Try

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A6:G13]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dt)

                        If dt.Rows.Count > 0 Then
                            'Vendor
                            If IsDBNull(dt.Rows(0).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid column ""Vendor"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'DeliveryDate
                            If IsDBNull(dt.Rows(2).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid column ""Delivery Req Date"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'ShipBy
                            If IsDBNull(dt.Rows(3).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid column ""Ship By."", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            If dt.Rows(3).Item(1).ToString.Trim.Length > 25 Then
                                lblInfo.Text = "[9999] Max 25 character in column ""Ship By."" , please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'KanbanNo
                            If IsDBNull(dt.Rows(5).Item(5)) Then
                                lblInfo.Text = "[9999] Invalid column ""PO No./Kanban No."", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            If dt.Rows(5).Item(5).ToString.Trim.Length > 25 Then
                                lblInfo.Text = "[9999] Max 25 character in column ""PO No./Kanban No."" , please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'Item
                            If IsDBNull(dt.Rows(7).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid colum ""PartNo."", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'DeliveryLocation
                            ls_sql = "select DeliveryLocationCode from MS_DeliveryPlace where affiliateID = '" & Session("AffiliateID") & "' and Defaultcls = 1"

                            Dim sqlCmd As New SqlCommand(ls_sql, sqlConn)
                            Dim sqlDA As New SqlDataAdapter(sqlCmd)
                            Dim ds As New DataSet
                            sqlDA.Fill(ds)

                            If ds.Tables(0).Rows.Count > 0 Then
                                Session("DeliveryLoc") = ds.Tables(0).Rows(0)("DeliveryLocationCode")
                            Else
                                lblInfo.Text = "[9999] Invalid Delivery Location, please select default!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If


                            Session.Remove("KanbanTime1")
                            Session.Remove("KanbanTime2")
                            Session.Remove("KanbanTime3")
                            Session.Remove("KanbanTime4")

                            ls_sql = " select * from ms_kanbantime where affiliateID = '" & Session("AffiliateID") & "'"
                            Dim sqlc As New SqlCommand(ls_sql, sqlConn)
                            Dim sqlDC As New SqlDataAdapter(sqlc)
                            Dim dsC As New DataSet
                            sqlDC.Fill(dsC)

                            If dsC.Tables(0).Rows.Count > 0 Then
                                If dsC.Tables(0).Rows(0)("kanbanCycle") = "1" Then
                                    Session("KanbanTime1") = dsC.Tables(0).Rows(0)("kanbanTime")
                                End If

                                If dsC.Tables(0).Rows(1)("kanbanCycle") = "2" Then
                                    Session("KanbanTime2") = dsC.Tables(0).Rows(1)("kanbanTime")
                                End If

                                If dsC.Tables(0).Rows(2)("kanbanCycle") = "3" Then
                                    Session("KanbanTime3") = dsC.Tables(0).Rows(2)("kanbanTime")
                                End If

                                If dsC.Tables(0).Rows(3)("kanbanCycle") = "4" Then
                                    Session("KanbanTime4") = dsC.Tables(0).Rows(3)("kanbanTime")
                                End If
                            Else
                                Session("KanbanTime1") = "00:00"
                                Session("KanbanTime2") = "00:00"
                                Session("KanbanTime3") = "00:00"
                                Session("KanbanTime4") = "00:00"

                            End If


                        End If

                        Dim dtUploadHeader As New clsKanbanHeader
                        Dim dtUploadHeaderList As New List(Of clsKanbanHeader)

                        'Dim dtUploadDetail As New clsPODetail
                        Dim dtUploadDetailList As New List(Of clsKanbanDetail)


                        'Get Header Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B6:B9]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtHeader)

                        If dtHeader.Rows.Count > 0 Then
                            dtUploadHeader.H_AffiliateID = Session("AffiliateID")
                            dtUploadHeader.H_Vendor = dtHeader.Rows(0).Item(0)
                            dtUploadHeader.H_DeliveryDate = Microsoft.VisualBasic.Left(dt.Rows(5).Item(5), 8) 'dtHeader.Rows(2).Item(0)
                            dtUploadHeader.H_ShipBy = dtHeader.Rows(3).Item(0)
                            dtUploadHeader.H_kanbanNo = dt.Rows(5).Item(5)
                            dtUploadHeader.H_Cycle = Microsoft.VisualBasic.Mid(dt.Rows(5).Item(5), 10, 2)
                        End If

                        'Get Detail Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B13:M65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IIf(IsDBNull(dtDetail.Rows(i).Item(0)), "", dtDetail.Rows(i).Item(0)) <> "" Then
                                    Dim dtUploadDetail As New clsKanbanDetail

                                    dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.D_MPQ = IIf(IsDBNull(dtDetail.Rows(i).Item(1)), 0, dtDetail.Rows(i).Item(1))
                                    dtUploadDetail.D_DailyUsage = IIf(IsDBNull(dtDetail.Rows(i).Item(2)), 0, dtDetail.Rows(i).Item(2))
                                    dtUploadDetail.D_uom = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), "", dtDetail.Rows(i).Item(3))
                                    dtUploadDetail.D_c1 = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), 0, dtDetail.Rows(i).Item(4))
                                    dtUploadDetail.D_b1 = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), 0, dtDetail.Rows(i).Item(5))
                                    'dtUploadDetail.D_c2 = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), 0, dtDetail.Rows(i).Item(6))
                                    'dtUploadDetail.D_b2 = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), 0, dtDetail.Rows(i).Item(7))
                                    'dtUploadDetail.D_c3 = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), 0, dtDetail.Rows(i).Item(8))
                                    'dtUploadDetail.D_b3 = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), 0, dtDetail.Rows(i).Item(9))
                                    'dtUploadDetail.D_c4 = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), 0, dtDetail.Rows(i).Item(10))
                                    'dtUploadDetail.D_b4 = IIf(IsDBNull(dtDetail.Rows(i).Item(11)), 0, dtDetail.Rows(i).Item(11))

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

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            '01. Check Kanban already Exists
                            ls_sql = "SELECT * FROM Kanban_Master WHERE KanbaNNo LIKE '%" & dtUploadHeader.H_kanbanNo & "%' " & vbCrLf & _
                                     " and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'" & vbCrLf & _
                                     " and SupplierID = '" & dtUploadHeader.H_Vendor & "' "

                            Dim sqlCmd As New SqlCommand(ls_sql, sqlConn, sqlTran)
                            Dim sqlDA As New SqlDataAdapter(sqlCmd)
                            Dim ds As New DataSet
                            sqlDA.Fill(ds)

                            If ds.Tables(0).Rows.Count > 0 Then
                                If Not IsDBNull(ds.Tables(0).Rows(0)("KanbanStatus")) Then
                                    Call clsMsg.DisplayMessage(lblInfo, "5012", clsMessage.MsgType.ErrorMessage)
                                    Exit Sub
                                End If
                            End If

                            '01.01 Delete TempoaryData
                            ls_sql = "delete UploadKanban where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'" & vbCrLf & _
                                     " and KanbanNO LIKE '%" & dtUploadHeader.H_kanbanNo & "%'" & vbCrLf & _
                                     " and SupplierID = '" & dtUploadHeader.H_Vendor & "' "

                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim Kanban As clsKanbanDetail = dtUploadDetailList(i)
                                Dim ls_Qty As Integer

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & Kanban.D_PartNo & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    ls_error = "PartNo not found in Part Master, please check again with PASI!"
                                Else
                                    ls_MOQ = IIf(IsDBNull(ds2.Tables(0).Rows(0)("MOQ")), 0, ds2.Tables(0).Rows(0)("MOQ"))
                                    ls_Qty = 0
                                    If CDbl(Kanban.D_c1) <> 0 Then ls_Qty = CDbl(Kanban.D_c1)
                                    'If CDbl(Kanban.D_c2) <> 0 Then ls_Qty = CDbl(Kanban.D_c2)
                                    'If CDbl(Kanban.D_c3) <> 0 Then ls_Qty = CDbl(Kanban.D_c3)
                                    'If CDbl(Kanban.D_c4) <> 0 Then ls_Qty = CDbl(Kanban.D_c4)

                                    If (ls_Qty Mod ls_MOQ) <> 0 And ls_Qty <> 0 Then
                                        If ls_error = "" Then
                                            ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                        End If
                                        'ElseIf ls_Qty = 0 Then
                                        'If ls_error = "" Then
                                        '    ls_error = "Please Input Qty !"
                                        'End If

                                    End If
                                End If

                                '02.2 Check PartNo di Ms_PartMapping
                                ls_sql = "select * from ms_partmapping WHERE PartNo = '" & Kanban.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'"
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)

                                If ds3.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "PartNo not found in Part Mapping, please check again with PASI!"
                                    End If
                                Else
                                    ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")
                                End If


                                '02.3 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.UploadKanban WHERE PartNo = '" & Kanban.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and KanbanNo Like '" & dtUploadHeader.H_kanbanNo & "'"
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                If ds4.Tables(0).Rows.Count > 0 Then
                                    ls_sql = "delete UploadKanban where PartNo = '" & Kanban.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and KanbanNo = '" & dtUploadHeader.H_kanbanNo & "'"
                                    Dim sqlComm1 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm1.ExecuteNonQuery()
                                    sqlComm1.Dispose()
                                End If

                                ls_sql = " INSERT INTO [dbo].[UploadKanban] " & vbCrLf & _
                                          "            ([AffiliateID], [ShipBy], [SupplierID], [Partno],[Cycle1],[Remarks],[KanbanNo],[DeliveryDate] )" & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & dtUploadHeader.H_AffiliateID & "' " & vbCrLf & _
                                          "            ,'" & dtUploadHeader.H_ShipBy & "' " & vbCrLf & _
                                          "            ,'" & dtUploadHeader.H_Vendor & "' " & vbCrLf

                                ls_sql = ls_sql + "            ,'" & Kanban.D_PartNo & "' " & vbCrLf & _
                                                  "            ,'" & Kanban.D_c1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_error & "' " & vbCrLf & _
                                                  "            , '" & dtUploadHeader.H_kanbanNo & "' " & vbCrLf & _
                                                  "            , '" & dtUploadHeader.H_DeliveryDate & "') " & vbCrLf
                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            Next
                            sqlTran.Commit()

                            Session("DeliveryDate") = dtUploadHeader.H_DeliveryDate
                            Session("PONoUpload") = dtUploadHeader.H_kanbanNo

                            lblInfo.Text = "[7001] Data Checking Done!"
                            lblInfo.ForeColor = Color.Blue
                            grid.JSProperties("cpMessage") = lblInfo.Text

                            Call bindData()
                        End Using
                    Catch ex As Exception
                        lblInfo.Text = ex.Message
                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                        Exit Sub
                    End Try
                    dt.Reset()
                    dtDetail.Reset()
                    dtHeader.Reset()
                End Using
                MyConnection.Close()

            Catch ex As Exception
                MyConnection.Close()
                Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            End Try
        Else
            If FileName = "" Then
                lblInfo.Text = "[9999] Please choose the file!"
                up_GridLoadWhenEventChange()
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If
        End If
        
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
        Dim ls_Period As Date
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim ls_DoubleSupplier As Boolean = False
        Dim ls_TempSupplierID As String = ""
        Dim ls_KanbanNo As String
        Dim ls_cycle As Integer
        Dim ls_Date As String
        Dim ls_time As String
        Dim ls_seq As Integer
        Dim pCycle As Integer

        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "remarks").ToString.Trim <> "" Then
                    ls_Check = True
                    Exit For
                End If
            Next i

            Dim countSupplier As Integer = 0

            For i = 0 To grid.VisibleRowCount - 1
                If i = 0 Then
                    ls_TempSupplierID = grid.GetRowValues(i, "supplier").ToString.Trim
                    countSupplier = 1
                End If

                If ls_TempSupplierID <> grid.GetRowValues(i, "supplier").ToString.Trim Then
                    ls_DoubleSupplier = True
                    ls_TempSupplierID = grid.GetRowValues(i, "supplier").ToString.Trim
                    countSupplier = countSupplier + 1
                End If
            Next i

            If ls_Check = True Then
                lblInfo.Text = "[9999] Invalid data in this File Upload, please check the file again!"
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If


            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()

            SqlTran = SqlCon.BeginTransaction

            Try
                'dian
                '1. Delete Data
                Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                SQLCom.Connection = SqlCon
                SQLCom.Transaction = SqlTran
                Dim ls_KanbanAsli As String = Trim(Session("PONoUpload"))
                Dim ls_deliveryDate As String = Session("DeliveryDate")

                ls_Sql = "Delete Kanban_Detail where KanbanNo like '%" & Trim(ls_KanbanAsli) & "%' and SupplierID = '" & ls_TempSupplierID & "' and AffiliateID = '" & Session("AffiliateID") & "'"
                ls_Sql = ls_Sql + "Delete Kanban_Master where KanbanNo like '" & Trim(ls_KanbanAsli) & "' and SupplierID = '" & ls_TempSupplierID & "' and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                '2. Insert New Data
                ls_KanbanNo = ""
                ls_Date = IIf(Mid(ls_deliveryDate, 7, 1) = 0, Right(ls_deliveryDate, 1), Right(ls_deliveryDate, 2))
                For i = 0 To grid.VisibleRowCount - 1
                    If grid.GetRowValues(i, "c1") <> 0 Then
                        'For x = 1 To 4
                        ls_KanbanNo = ls_KanbanAsli '& "-" & x
                        ls_cycle = grid.GetRowValues(i, "c1")
                        'If x = 2 Then ls_cycle = grid.GetRowValues(i, "c2")
                        'If x = 3 Then ls_cycle = grid.GetRowValues(i, "c3")
                        'If x = 4 Then ls_cycle = grid.GetRowValues(i, "c4")

                        ls_time = Session("KanbanTime1")
                        'If x = 2 Then ls_time = Session("KanbanTime2")
                        'If x = 3 Then ls_time = Session("KanbanTime3")
                        'If x = 4 Then ls_time = Session("KanbanTime4")
                        If (Microsoft.VisualBasic.Right(ls_KanbanNo, 1)) <> "E" Then
                            If CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 2)) <= 4 Then
                                ls_seq = "1"
                                pCycle = CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 2))
                            ElseIf CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 2)) Mod 4 >= 1 Then
                                ls_seq = (CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 2)) \ 4) + 1
                                pCycle = CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 2))
                            End If
                        Else
                            If CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 1)) <= 4 Then
                                ls_seq = "1"
                                pCycle = CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 1))
                            ElseIf CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 1)) Mod 4 >= 1 Then
                                ls_seq = (CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 1)) \ 4) + 1
                                pCycle = CInt(Microsoft.VisualBasic.Mid(ls_KanbanNo, 10, 1))
                            End If
                        End If

                        If ls_cycle <> 0 Then

                            ls_Sql = " IF NOT EXISTS (select * From Kanban_Master where KanbanNo = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                     "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                     "              and AffiliateID = '" & Session("AffiliateID") & "') BEGIN " & vbCrLf & _
                                     " Insert Into Kanban_Master Values( " & vbCrLf & _
                                     " '" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                     " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                     " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                     " '" & pCycle & "', " & vbCrLf & _
                                     " '" & Left(ls_deliveryDate, 4) + "-" + Mid(ls_deliveryDate, 5, 2) + "-" + Mid(ls_deliveryDate, 7, 2) & "', " & vbCrLf & _
                                     " '" & ls_time & "', " & vbCrLf & _
                                     " '0', " & vbCrLf & _
                                     " '', " & vbCrLf & _
                                     " NULL, " & vbCrLf & _
                                     " '' , " & vbCrLf & _
                                     " NULL, " & vbCrLf & _
                                     " GetDate(), " & vbCrLf & _
                                     " '" & Session("UserID").ToString & "', " & vbCrLf & _
                                     " NULL, NULL, " & vbCrLf & _
                                     " '" & Session("DeliveryLoc") & "', " & vbCrLf & _
                                     " '1', " & ls_seq & ") END" & vbCrLf
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()

                            ls_Sql = "Insert into Kanban_Detail values ( " & vbCrLf & _
                                     " '" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                     " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                     " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                     " '" & grid.GetRowValues(i, "partno") & "', " & vbCrLf & _
                                     " '" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                     " '" & Trim(Session("DeliveryLoc")) & "', " & vbCrLf & _
                                     " '" & Trim(grid.GetRowValues(i, "uom")) & "', " & vbCrLf & _
                                     "  " & CDbl(ls_cycle) & " )"

                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()

                            'insert PO
                            ls_Sql = "  IF NOT EXISTS (select * From PO_Master where PONO = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                     "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                     "              and AffiliateID = '" & Session("AffiliateID") & "') BEGIN " & vbCrLf & _
                                     "  INSERT INTO PO_Master ( PONO,AffiliateID, SupplierID, Period, CommercialCls,  " & vbCrLf & _
                                     " ShipCls, DeliveryBypasiCls, AffiliateApproveDate, " & vbCrLf & _
                                     " AffiliateApproveUser,PasiSendAffiliateDate, PasiSendAffiliateUSer, SupplierApproveDate, " & vbCrLf & _
                                     " SupplierApproveUser, PasiApproveDate, PasiApproveUser, FinalApproveDate, FinalApproveUser, " & vbCrLf & _
                                     " EntryDate, EntryUSer, UpdateDate, UpdateUser) " & vbCrLf & _
                                     " VALUES( " & vbCrLf & _
                                     " '" & ls_KanbanNo & "', " & vbCrLf & _
                                     " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                     " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                     " '" & Left(ls_deliveryDate, 4) + "-" + Mid(ls_deliveryDate, 5, 2) + "-01" & "', " & vbCrLf & _
                                     " '1', " & vbCrLf

                            ls_Sql = ls_Sql + " '" & grid.GetRowValues(i, "shipby") & "', " & vbCrLf & _
                                              " '1', " & vbCrLf & _
                                              " GETDATE(), " & vbCrLf & _
                                              " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                              " GETDATE(), " & vbCrLf & _
                                              " 'administrator', " & vbCrLf & _
                                              " GETDATE(), " & vbCrLf & _
                                              " '', " & vbCrLf & _
                                              " GETDATE(), " & vbCrLf & _
                                              " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                              " GETDATE(), "

                            ls_Sql = ls_Sql + " 'administrator', " & vbCrLf & _
                                              " GETDATE(), " & vbCrLf & _
                                              " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                              " GETDATE(), " & vbCrLf & _
                                              " '" & Session("AffiliateID") & "' " & vbCrLf & _
                                              " ) END"

                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()

                            ls_Sql = " INSERT INTO PO_Detail " & vbCrLf & _
                                     " (PONO, AffiliateID, SupplierID, PartNo, KanbanCls, POQty, DeliveryD" & ls_Date & ", EntryDate,EntryUser, UpdateDate, UpdateUser) " & vbCrLf & _
                                     " Values ('" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                     " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                     " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                     " '" & grid.GetRowValues(i, "partno") & "', " & vbCrLf & _
                                     " '1', " & vbCrLf & _
                                     " " & CDbl(ls_cycle) & ", " & vbCrLf & _
                                     " " & CDbl(ls_cycle) & ", " & vbCrLf & _
                                     " GETDATE(), " & vbCrLf & _
                                     " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                     " GETDATE(), "
                            ls_Sql = ls_Sql + " '" & Session("AffiliateID") & "') "
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()

                            ls_MsgID = "1001"
                            ls_Detail = "ada"
                        End If
                        'Next
                    End If
                Next

                ''Kanban Barcode
                'ls_Sql = "update kanban_master set AffiliateApproveUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                '                 " AffiliateApproveDate = getdate(), Excelcls = '1' " & vbCrLf & _
                '                 " where convert(char(19),convert(datetime,KanbanDate),112) ='" & ls_deliveryDate & "'" & vbCrLf & _
                '                 " and AffiliateID = '" & Session("AffiliateID") & " ' " & vbCrLf & _
                '                 " and SupplierID = '" & ls_TempSupplierID & "'" & vbCrLf & _
                '                 " AND deliveryLocationcode = '" & Session("DeliveryLoc") & "'" & vbCrLf

                'ls_Sql = ls_Sql + "  DECLARE @KanbanNo AS VARCHAR(10) ,  " & vbCrLf & _
                '                  "      @SupplierID AS VARCHAR(10) ,  " & vbCrLf & _
                '                  "      @SupplierName AS VARCHAR(100) ,  " & vbCrLf & _
                '                  "      @PartNo AS VARCHAR(50) ,  " & vbCrLf & _
                '                  "      @PartName AS VARCHAR(100) ,  " & vbCrLf & _
                '                  "      @Qty AS NUMERIC(10, 2) ,  " & vbCrLf & _
                '                  "      @Cust AS VARCHAR(50) ,  " & vbCrLf & _
                '                  "      @DeliveryDate AS VARCHAR(10) ,  " & vbCrLf & _
                '                  "      @TIME AS VARCHAR(10) ,  " & vbCrLf & _
                '                  "      @Location AS VARCHAR(50) ,  " & vbCrLf & _
                '                  "      @PONo AS VARCHAR(50) ,      @Barcode AS VARCHAR(1000) ,  " & vbCrLf

                'ls_Sql = ls_Sql + "      @QtyBox AS NUMERIC(10, 2) ,  " & vbCrLf & _
                '                  "      @Loop AS NUMERIC(10, 2) ,  " & vbCrLf & _
                '                  "      @StartNo AS NUMERIC(10, 2) ,  " & vbCrLf & _
                '                  "      @Total AS NUMERIC(10, 2) ,  " & vbCrLf & _
                '                  "      @PartNoSave AS CHAR(50)  " & vbCrLf & _
                '                  "  		  " & vbCrLf & _
                '                  "  SELECT TOP 1  " & vbCrLf & _
                '                  "          KanbanNo = CONVERT(CHAR(10), '') ,  " & vbCrLf & _
                '                  "          SupplierID = CONVERT(CHAR(10), '') ,  " & vbCrLf & _
                '                  "          SupplierName = CONVERT(CHAR(100), '') ,          PartNo = CONVERT(CHAR(50), '') ,  " & vbCrLf & _
                '                  "          PartName = CONVERT(CHAR(100), '') ,  " & vbCrLf

                'ls_Sql = ls_Sql + "          Qty = 0 ,  " & vbCrLf & _
                '                  "          Cust = CONVERT(CHAR(50), '') ,  " & vbCrLf & _
                '                  "          DeliveryDate = CONVERT(CHAR(10), '') ,  " & vbCrLf & _
                '                  "          TIME = CONVERT(CHAR(10), '') ,  " & vbCrLf & _
                '                  "          Location = CONVERT(CHAR(50), '') ,  " & vbCrLf & _
                '                  "          PONo = CONVERT(CHAR(50), '') ,  " & vbCrLf & _
                '                  "          Barcode = CONVERT(CHAR(1000), '') ,  " & vbCrLf & _
                '                  "          qtybox = 0 ,  " & vbCrLf & _
                '                  "          startno = 0 ,          total = 0  " & vbCrLf & _
                '                  "  INTO    #data  " & vbCrLf & _
                '                  "  FROM    dbo.Kanban_Master KM  " & vbCrLf

                'ls_Sql = ls_Sql + "    " & vbCrLf & _
                '                  "  DELETE  FROM #data  " & vbCrLf & _
                '                  "  WHERE   KanbanNo = ''  " & vbCrLf & _
                '                  "  		  " & vbCrLf & _
                '                  "  DECLARE cur_Print CURSOR FOR  " & vbCrLf & _
                '                  "  SELECT  KM.KanbanNo AS kanbanNo ,  " & vbCrLf & _
                '                  "  KM.SupplierID AS SupplierID ,  " & vbCrLf & _
                '                  "  MSS.SupplierName AS SupplierName ,  KD.PartNo AS PartNo ,  " & vbCrLf & _
                '                  "  MSP.PartName AS PartName ,  " & vbCrLf & _
                '                  "  KD.KanbanQty Qty ,  " & vbCrLf & _
                '                  "  KM.AffiliateID AS Cust ,  " & vbCrLf

                'ls_Sql = ls_Sql + "  DeliveryDate = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) ,  " & vbCrLf & _
                '                  "  CONVERT(CHAR(5), KM.KanbanTime) AS TIME ,  " & vbCrLf & _
                '                  "  ISNULL(KM.DeliveryLocationCode,'') Location ,  " & vbCrLf & _
                '                  "  KD.PONo ,  " & vbCrLf & _
                '                  "  Barcode = RTRIM(CONVERT(CHAR, KD.PONo)) " & vbCrLf & _
                '                  "  + RTRIM(CONVERT(CHAR, KM.KanbanNo)) " & vbCrLf & _
                '                  "  + RTRIM(CONVERT(CHAR, KM.AffiliateID)) + RTRIM(CONVERT(CHAR, KM.SupplierID))  " & vbCrLf & _
                '                  "  + RTRIM(CONVERT(CHAR, KD.PartNo))    " & vbCrLf & _
                '                  "  + RTRIM(CONVERT(CHAR, KD.KanbanQty)) ,  " & vbCrLf & _
                '                  "  QtyBox = MSP.QtyBox  " & vbCrLf & _
                '                  "  FROM    dbo.Kanban_Master KM  " & vbCrLf

                'ls_Sql = ls_Sql + "  LEFT JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID  " & vbCrLf & _
                '                  "  AND KM.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                '                  "  AND KM.SupplierID = KD.SupplierID  " & vbCrLf & _
                '                  "   LEFT JOIN dbo.MS_DeliveryPlace MDP ON MDP.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
                '                  "  LEFT JOIN dbo.MS_Supplier MSS ON MSS.SupplierID = KM.SupplierID  " & vbCrLf & _
                '                  "  LEFT JOIN dbo.MS_Parts MSP ON MSP.PartNo = KD.PartNo  " & vbCrLf & _
                '                  "  WHERE KanbanQty <> 0 and KD.SupplierID <> ''  " & vbCrLf & _
                '                  "  AND CONVERT(CHAR(8), CONVERT(DATETIME, KanbanDate),112) = '" & ls_deliveryDate & "'   AND KD.SupplierID  = '" & Trim(ls_TempSupplierID) & "'   " & vbCrLf & _
                '                  "  AND KD. AffiliateID = '" & Session("AffiliateID") & " '" & vbCrLf & _
                '                  "  AND KD.DeliveryLocationCode = '" & Session("DeliveryLoc") & "'" & vbCrLf & _
                '                  "  OPEN cur_Print  " & vbCrLf & _
                '                  "    "

                'ls_Sql = ls_Sql + "  FETCH NEXT FROM cur_Print  " & vbCrLf & _
                '                  "     INTO @KanbanNo, @SupplierID, @SupplierName, @PartNo, @PartName,  " & vbCrLf & _
                '                  "  		@Qty, @Cust, @DeliveryDate, @TIME, @Location, @PONo, @Barcode, @QtyBox  " & vbCrLf & _
                '                  "    " & vbCrLf & _
                '                  "  WHILE @@Fetch_Status = 0   " & vbCrLf & _
                '                  "      BEGIN  " & vbCrLf & _
                '                  "          SET @StartNo = 0          SET @total = 0  " & vbCrLf & _
                '                  "          WHILE @Total < @Qty   " & vbCrLf & _
                '                  "              BEGIN  " & vbCrLf & _
                '                  "                  BEGIN  " & vbCrLf & _
                '                  "                      SET @StartNo = @StartNo + 1  " & vbCrLf

                'ls_Sql = ls_Sql + "                      INSERT  INTO #Data  " & vbCrLf & _
                '                  "                      VALUES  ( @KanbanNo, @SupplierID, @SupplierName, @PartNo,  " & vbCrLf & _
                '                  "                                @PartName, @Qty, @Cust, @DeliveryDate, @Time,  " & vbCrLf & _
                '                  "                                @Location, @PONo,  " & vbCrLf & _
                '                  "                                (RTRIM(CONVERT(CHAR, @KanbanNo))  " & vbCrLf & _
                '                  "                                + RTRIM(CONVERT(CHAR, @SupplierID))  " & vbCrLf & _
                '                  "                                + RTRIM(CONVERT(NUMERIC, @StartNo)) " & vbCrLf & _
                '                  "                                + RTRIM(CONVERT(CHAR, @Cust)) " & vbCrLf & _
                '                  "                                + RTRIM(CONVERT(CHAR, @Partno))) " & vbCrLf & _
                '                  "                                , @QtyBox, @StartNo,  " & vbCrLf & _
                '                  "                                ( @Qty / @QtyBox ) )  " & vbCrLf & _
                '                  "                      SET @Total = @Total + @QtyBox                  END  " & vbCrLf

                'ls_Sql = ls_Sql + "  				     " & vbCrLf & _
                '                  "              END	  " & vbCrLf & _
                '                  "    " & vbCrLf & _
                '                  "          FETCH NEXT FROM cur_Print  " & vbCrLf & _
                '                  "  		 INTO @KanbanNo, @SupplierID, @SupplierName, @PartNo, @PartName,  " & vbCrLf & _
                '                  "  			@Qty, @Cust, @DeliveryDate, @TIME, @Location, @PONo, @Barcode, @QtyBox  " & vbCrLf & _
                '                  "    " & vbCrLf & _
                '                  "      END  " & vbCrLf & _
                '                  "  CLOSE cur_Print  " & vbCrLf & _
                '                  "  DEALLOCATE cur_Print   " & vbCrLf & _
                '                  "  INSERT INTO Kanban_Barcode  " & vbCrLf

                'ls_Sql = ls_Sql + "  SELECT  Barcode ,  " & vbCrLf & _
                '                  "          PONo ,  " & vbCrLf & _
                '                  "          KanbanNo ,  " & vbCrLf & _
                '                  "          startno ,  " & vbCrLf & _
                '                  "          cust ,  " & vbCrLf & _
                '                  "          SupplierID ,  " & vbCrLf & _
                '                  "          Location,  " & vbCrLf & _
                '                  "          PartNo ,                   " & vbCrLf & _
                '                  "          qtybox  " & vbCrLf & _
                '                  "  FROM    #data   "


                'SQLCom.CommandText = ls_Sql
                'SQLCom.ExecuteNonQuery()

                'insert affiliate Master
                ls_Sql = "Insert Into Affiliate_master (PONO, AffiliateID, SupplierID, Excelcls,Entrydate, EntryUser, UpdateDate, UpdateUSer) " & vbCrLf & _
                         " select PONO, AffiliateID, SupplierID, '2',Entrydate, EntryUser, UpdateDate, UpdateUSer " & vbCrLf & _
                         " From PO_Master where PONO like '%" & Trim(ls_KanbanAsli) & "%'" & vbCrLf & _
                         "              and SupplierID = '" & ls_TempSupplierID & "' " & vbCrLf & _
                         "              and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                'Insert Affiliate Detail
                ls_Sql = " Insert Into Affiliate_Detail " & vbCrLf & _
                         " SELECT PONO, AffiliateID, SupplierID, a.PartNo, '0','1',b.Maker,POQty,POQty, " & vbCrLf & _
                         " '',0,0,DeliveryD1, DeliveryD1, " & vbCrLf & _
                         " DeliveryD2, DeliveryD2, DeliveryD3, DeliveryD3, DeliveryD4, DeliveryD4, DeliveryD5, DeliveryD5," & vbCrLf & _
                         " DeliveryD6, DeliveryD6, DeliveryD7, DeliveryD7, DeliveryD8, DeliveryD8, DeliveryD9, DeliveryD9, " & vbCrLf & _
                         " DeliveryD10, DeliveryD10, DeliveryD11, DeliveryD11, DeliveryD12, DeliveryD12, DeliveryD13, DeliveryD13, " & vbCrLf & _
                         " DeliveryD14, DeliveryD14, DeliveryD15, DeliveryD15, DeliveryD16, DeliveryD16, DeliveryD17, DeliveryD17," & vbCrLf & _
                         " DeliveryD18, DeliveryD18, DeliveryD19, DeliveryD19, DeliveryD20, DeliveryD20, DeliveryD21, DeliveryD21," & vbCrLf & _
                         " DeliveryD22, DeliveryD22, DeliveryD23, DeliveryD23, DeliveryD24, DeliveryD24, DeliveryD25, DeliveryD25," & vbCrLf & _
                         " DeliveryD26, DeliveryD26, DeliveryD27, DeliveryD27, DeliveryD28, DeliveryD28, DeliveryD29, DeliveryD29, DeliveryD30," & vbCrLf & _
                         " DeliveryD30, DeliveryD31, DeliveryD31, a.Entrydate, a.entryuser, a.updatedate, a.updateUser" & vbCrLf & _
                         " FROM PO_Detail a LEFT JOIN MS_Parts b ON a.partno = b.partno" & vbCrLf & _
                         " where PONO like '%" & Trim(ls_KanbanAsli) & "%'" & vbCrLf & _
                         "              and SupplierID = '" & ls_TempSupplierID & "' " & vbCrLf & _
                         "              and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                'Insert PO_MasterUpload
                ls_Sql = "Insert Into PO_MasterUpload (PONO, AffiliateID, SupplierID, Remarks,Entrydate, EntryUser, UpdateDate, UpdateUSer)" & vbCrLf & _
                         " select PONO, AffiliateID, SupplierID, '',Entrydate, EntryUser, UpdateDate, UpdateUSer " & vbCrLf & _
                         " From PO_Master where PONO like '%" & Trim(ls_KanbanAsli) & "%'" & vbCrLf & _
                         "              and SupplierID = '" & ls_TempSupplierID & "' " & vbCrLf & _
                         "              and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                'Insert PO_DetailUpload
                ls_Sql = " Insert Into PO_DetailUpload " & vbCrLf & _
                         " SELECT PONO, AffiliateID, SupplierID, a.PartNo, '0','1',b.Maker,POQty,POQty, " & vbCrLf & _
                         " '',0,0,DeliveryD1, DeliveryD1, " & vbCrLf & _
                         " DeliveryD2, DeliveryD2, DeliveryD3, DeliveryD3, DeliveryD4, DeliveryD4, DeliveryD5, DeliveryD5," & vbCrLf & _
                         " DeliveryD6, DeliveryD6, DeliveryD7, DeliveryD7, DeliveryD8, DeliveryD8, DeliveryD9, DeliveryD9, " & vbCrLf & _
                         " DeliveryD10, DeliveryD10, DeliveryD11, DeliveryD11, DeliveryD12, DeliveryD12, DeliveryD13, DeliveryD13, " & vbCrLf & _
                         " DeliveryD14, DeliveryD14, DeliveryD15, DeliveryD15, DeliveryD16, DeliveryD16, DeliveryD17, DeliveryD17," & vbCrLf & _
                         " DeliveryD18, DeliveryD18, DeliveryD19, DeliveryD19, DeliveryD20, DeliveryD20, DeliveryD21, DeliveryD21," & vbCrLf & _
                         " DeliveryD22, DeliveryD22, DeliveryD23, DeliveryD23, DeliveryD24, DeliveryD24, DeliveryD25, DeliveryD25," & vbCrLf & _
                         " DeliveryD26, DeliveryD26, DeliveryD27, DeliveryD27, DeliveryD28, DeliveryD28, DeliveryD29, DeliveryD29, DeliveryD30," & vbCrLf & _
                         " DeliveryD30, DeliveryD31, DeliveryD31, a.Entrydate, a.entryuser, a.updatedate, a.updateUser" & vbCrLf & _
                         " FROM PO_Detail a LEFT JOIN MS_Parts b ON a.partno = b.partno" & vbCrLf & _
                         " where PONO like '%" & Trim(ls_KanbanAsli) & "%'" & vbCrLf & _
                         "              and SupplierID = '" & ls_TempSupplierID & "' " & vbCrLf & _
                         "              and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                'delete tada di tempolary table
                ls_Sql = "delete UploadKanban where AffiliateID = '" & Session("AffiliateID") & "' and KanbanNo = '" & Session("PONoUpload") & "'"

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
                Exit Sub
            End Try

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("YA010IsSubmit") = lblInfo.Text
        End Try
    End Sub

#End Region
End Class