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

Public Class POUpload
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "D05"
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
            Dim fi As New FileInfo(Server.MapPath("~\Template\TemplatePO.xlsx"))
            If Not fi.Exists Then
                lblInfo.Text = "[9999] Excel Template Not Found !"
                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Template\TemplatePO.xlsx")

            'lblInfo.Text = "[9998] Download template successful"
            'ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        End Try
        
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("ErrorCls") = "" Then
        Else
            e.Cell.BackColor = Color.Red
        End If
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
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
                  " 	row_number() over (order by d.SupplierID, a.PartNo asc) as NoUrut, " & vbCrLf & _
                  " 	Period, Period as Period2,PONo, ShipCls, a.PartNo, b.PartName, c.Description UnitDesc, b.Maker, b.Project, d.MOQ MinOrderQty, " & vbCrLf & _
                  " 	d.SupplierID, a.POQty, a.ForecastN1, a.ForecastN2, a.ForecastN3, " & vbCrLf & _
                  " 	DeliveryD1, DeliveryD2, DeliveryD3, DeliveryD4, DeliveryD5,     " & vbCrLf & _
                  "    	DeliveryD6, DeliveryD7, DeliveryD8, DeliveryD9, DeliveryD10,     " & vbCrLf & _
                  "    	DeliveryD11, DeliveryD12, DeliveryD13, DeliveryD14,  DeliveryD15,     " & vbCrLf & _
                  "    	DeliveryD16, DeliveryD17, DeliveryD18, DeliveryD19, DeliveryD20,     " & vbCrLf & _
                  "    	DeliveryD21, DeliveryD22, DeliveryD23, DeliveryD24, DeliveryD25,     " & vbCrLf & _
                  "    	DeliveryD26, DeliveryD27, DeliveryD28, DeliveryD29, DeliveryD30,     " & vbCrLf & _
                  "    	DeliveryD31, a.ErrorCls, a.KanbanCls "

            ls_SQL = ls_SQL + " from UploadPO a  " & vbCrLf & _
                              " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " left join MS_UnitCls c on b.UnitCls = c.UnitCls " & vbCrLf & _
                              " left join MS_PartMapping d on a.PartNo = d.PartNo and a.AffiliateID = d.AffiliateID and a.SupplierID = d.SupplierID " & vbCrLf & _
                              " where a.PONo = '" & Session("PONoUpload") & "' and a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " order by SupplierID, PartNo"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

            clsGlobal.HideColumTanggal1(Session("Period"), grid)
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
        Dim ls_SuppCode As String = ""

        Try
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

                If MyConnection.State = ConnectionState.Open Then MyConnection.Close()

                MyConnection.Open()

                Dim dtSheets As DataTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                Dim listSheet As New List(Of String)
                Dim drSheet As DataRow

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

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A3:C14]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dt)

                        If dt.Rows.Count > 0 Then
                            'Period
                            If IsDBNull(dt.Rows(0).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid column ""Period"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'KanbanCls
                            If IsDBNull(dt.Rows(2).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid column ""PO Kanban"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            If dt.Rows(2).Item(2).ToString.Trim.ToUpper <> "YES" And dt.Rows(2).Item(2).ToString.Trim.ToUpper <> "NO" Then
                                lblInfo.Text = "[9999] PO Kanban must be fill with ""Yes"" or ""No"" , please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'PONo
                            If IsDBNull(dt.Rows(3).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid column ""PONo."", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            If dt.Rows(3).Item(2).ToString.Trim.Length > 20 Then
                                lblInfo.Text = "[9999] Max 20 character in column ""PONo."" , please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'ShipBy
                            If IsDBNull(dt.Rows(4).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid column ""ShipBy"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            If dt.Rows(4).Item(2).ToString.Trim.Length > 25 Then
                                lblInfo.Text = "[9999] Max 25 character in column ""ShipBy"" , please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'Item
                            If IsDBNull(dt.Rows(11).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid colum ""PartNo."", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                        End If

                        Dim dtUploadHeader As New clsPOHeader
                        Dim dtUploadHeaderList As New List(Of clsPOHeader)

                        Dim dtUploadDetailList As New List(Of clsPODetail)


                        'Get Header Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "C3:C8]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtHeader)

                        If dtHeader.Rows.Count > 0 Then
                            dtUploadHeader.H_AffiliateID = Session("AffiliateID")
                            Try
                                tempDate = "01-" & dtHeader.Rows(0).Item(0)
                                Session("Period") = tempDate
                            Catch ex As Exception
                                lblInfo.Text = "[9999] Invalid Period, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End Try

                            dtUploadHeader.H_Period = tempDate
                            If dtHeader.Rows(2).Item(0).ToString.Trim.ToUpper = "YES" Then
                                dtUploadHeader.H_POKanban = 1
                            Else
                                dtUploadHeader.H_POKanban = 0
                            End If
                            dtUploadHeader.H_PONo = dtHeader.Rows(3).Item(0)
                            dtUploadHeader.H_ShipBy = dtHeader.Rows(4).Item(0)
                        End If


                        'Get Detail Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B14:AR65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsPODetail
                                    dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(0)
                                    'dtUploadDetail.H_SupplierID = dtDetail.Rows(i).Item(6)
                                    dtUploadDetail.Forecast1 = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), 0, dtDetail.Rows(i).Item(8))
                                    dtUploadDetail.Forecast2 = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), 0, dtDetail.Rows(i).Item(9))
                                    dtUploadDetail.Forecast3 = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), 0, dtDetail.Rows(i).Item(10))
                                    dtUploadDetail.D_D1 = IIf(IsDBNull(dtDetail.Rows(i).Item(12)), 0, dtDetail.Rows(i).Item(12))
                                    dtUploadDetail.D_D2 = IIf(IsDBNull(dtDetail.Rows(i).Item(13)), 0, dtDetail.Rows(i).Item(13))
                                    dtUploadDetail.D_D3 = IIf(IsDBNull(dtDetail.Rows(i).Item(14)), 0, dtDetail.Rows(i).Item(14))
                                    dtUploadDetail.D_D4 = IIf(IsDBNull(dtDetail.Rows(i).Item(15)), 0, dtDetail.Rows(i).Item(15))
                                    dtUploadDetail.D_D5 = IIf(IsDBNull(dtDetail.Rows(i).Item(16)), 0, dtDetail.Rows(i).Item(16))
                                    dtUploadDetail.D_D6 = IIf(IsDBNull(dtDetail.Rows(i).Item(17)), 0, dtDetail.Rows(i).Item(17))
                                    dtUploadDetail.D_D7 = IIf(IsDBNull(dtDetail.Rows(i).Item(18)), 0, dtDetail.Rows(i).Item(18))
                                    dtUploadDetail.D_D8 = IIf(IsDBNull(dtDetail.Rows(i).Item(19)), 0, dtDetail.Rows(i).Item(19))
                                    dtUploadDetail.D_D9 = IIf(IsDBNull(dtDetail.Rows(i).Item(20)), 0, dtDetail.Rows(i).Item(20))
                                    dtUploadDetail.D_D10 = IIf(IsDBNull(dtDetail.Rows(i).Item(21)), 0, dtDetail.Rows(i).Item(21))
                                    dtUploadDetail.D_D11 = IIf(IsDBNull(dtDetail.Rows(i).Item(22)), 0, dtDetail.Rows(i).Item(22))
                                    dtUploadDetail.D_D12 = IIf(IsDBNull(dtDetail.Rows(i).Item(23)), 0, dtDetail.Rows(i).Item(23))
                                    dtUploadDetail.D_D13 = IIf(IsDBNull(dtDetail.Rows(i).Item(24)), 0, dtDetail.Rows(i).Item(24))
                                    dtUploadDetail.D_D14 = IIf(IsDBNull(dtDetail.Rows(i).Item(25)), 0, dtDetail.Rows(i).Item(25))
                                    dtUploadDetail.D_D15 = IIf(IsDBNull(dtDetail.Rows(i).Item(26)), 0, dtDetail.Rows(i).Item(26))
                                    dtUploadDetail.D_D16 = IIf(IsDBNull(dtDetail.Rows(i).Item(27)), 0, dtDetail.Rows(i).Item(27))
                                    dtUploadDetail.D_D17 = IIf(IsDBNull(dtDetail.Rows(i).Item(28)), 0, dtDetail.Rows(i).Item(28))
                                    dtUploadDetail.D_D18 = IIf(IsDBNull(dtDetail.Rows(i).Item(29)), 0, dtDetail.Rows(i).Item(29))
                                    dtUploadDetail.D_D19 = IIf(IsDBNull(dtDetail.Rows(i).Item(30)), 0, dtDetail.Rows(i).Item(30))
                                    dtUploadDetail.D_D20 = IIf(IsDBNull(dtDetail.Rows(i).Item(31)), 0, dtDetail.Rows(i).Item(31))
                                    dtUploadDetail.D_D21 = IIf(IsDBNull(dtDetail.Rows(i).Item(32)), 0, dtDetail.Rows(i).Item(32))
                                    dtUploadDetail.D_D22 = IIf(IsDBNull(dtDetail.Rows(i).Item(33)), 0, dtDetail.Rows(i).Item(33))
                                    dtUploadDetail.D_D23 = IIf(IsDBNull(dtDetail.Rows(i).Item(34)), 0, dtDetail.Rows(i).Item(34))
                                    dtUploadDetail.D_D24 = IIf(IsDBNull(dtDetail.Rows(i).Item(35)), 0, dtDetail.Rows(i).Item(35))
                                    dtUploadDetail.D_D25 = IIf(IsDBNull(dtDetail.Rows(i).Item(36)), 0, dtDetail.Rows(i).Item(36))
                                    dtUploadDetail.D_D26 = IIf(IsDBNull(dtDetail.Rows(i).Item(37)), 0, dtDetail.Rows(i).Item(37))
                                    dtUploadDetail.D_D27 = IIf(IsDBNull(dtDetail.Rows(i).Item(38)), 0, dtDetail.Rows(i).Item(38))
                                    dtUploadDetail.D_D28 = IIf(IsDBNull(dtDetail.Rows(i).Item(39)), 0, dtDetail.Rows(i).Item(39))
                                    dtUploadDetail.D_D29 = IIf(IsDBNull(dtDetail.Rows(i).Item(40)), 0, dtDetail.Rows(i).Item(40))
                                    dtUploadDetail.D_D30 = IIf(IsDBNull(dtDetail.Rows(i).Item(41)), 0, dtDetail.Rows(i).Item(41))
                                    dtUploadDetail.D_D31 = IIf(IsDBNull(dtDetail.Rows(i).Item(42)), 0, dtDetail.Rows(i).Item(42))
                                    dtUploadDetail.D_POQty = dtUploadDetail.D_D1 + dtUploadDetail.D_D2 + dtUploadDetail.D_D3 + dtUploadDetail.D_D4 + dtUploadDetail.D_D5 + _
                                                             dtUploadDetail.D_D6 + dtUploadDetail.D_D7 + dtUploadDetail.D_D8 + dtUploadDetail.D_D9 + dtUploadDetail.D_D10 + _
                                                             dtUploadDetail.D_D11 + dtUploadDetail.D_D12 + dtUploadDetail.D_D13 + dtUploadDetail.D_D14 + dtUploadDetail.D_D15 + _
                                                             dtUploadDetail.D_D16 + dtUploadDetail.D_D17 + dtUploadDetail.D_D18 + dtUploadDetail.D_D19 + dtUploadDetail.D_D20 + _
                                                             dtUploadDetail.D_D21 + dtUploadDetail.D_D22 + dtUploadDetail.D_D23 + dtUploadDetail.D_D24 + dtUploadDetail.D_D25 + _
                                                             dtUploadDetail.D_D26 + dtUploadDetail.D_D27 + dtUploadDetail.D_D28 + dtUploadDetail.D_D29 + dtUploadDetail.D_D30 + dtUploadDetail.D_D31
                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            '01.01 Delete TempoaryData
                            ls_sql = "delete UploadPO where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and PONo = '" & dtUploadHeader.H_PONo & "'"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim PO As clsPODetail = dtUploadDetailList(i)

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & PO.D_PartNo & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "PartNo not found in Part Master, please check again with PASI!"
                                    Else
                                        ls_error = ls_error + ", PartNo not found in Part Master, please check again with PASI!"
                                    End If
                                End If

                                '02.1.1 Check PartNo di MS_Price
                                'ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & PO.D_PartNo & "' "
                                ls_sql = "select * from MS_Price where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and PartNo = '" & PO.D_PartNo & "' and ('" & dtUploadHeader.H_Period & "' between StartDate and EndDate)"
                                Dim sqlCmd8 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA8 As New SqlDataAdapter(sqlCmd8)
                                Dim ds8 As New DataSet
                                sqlDA8.Fill(ds8)

                                If ds8.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "This PartNo not found or expired in Price Master, please check again with PASI!"
                                    Else
                                        ls_error = ls_error + ", This PartNo not found or expired in Price Master, please check again with PASI!"
                                    End If
                                End If

                                '02.2 Check PartNo di Ms_PartMapping
                                'ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and SupplierID = '" & PO.H_SupplierID & "'"
                                ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' "
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)

                                If ds3.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "PartNo not found in Part Mapping, please check again with PASI!"
                                    Else
                                        ls_error = ls_error + ", PartNo not found in Part Mapping, please check again with PASI!"
                                    End If
                                Else
                                    ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                    ls_SuppCode = IIf(IsDBNull(ds3.Tables(0).Rows(0)("SupplierID")), 0, ds3.Tables(0).Rows(0)("SupplierID"))
                                    If PO.D_POQty > 0 Then
                                        If (PO.D_POQty Mod ls_MOQ) <> 0 Then
                                            If ls_error = "" Then
                                                ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                            Else
                                                ls_error = ls_error + ", Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                            End If
                                        End If
                                    End If
                                End If

                                '02.2 Check PONO di PO MASTER
                                ls_sql = "SELECT * FROM PO_Master WHERE PONo = '" & dtUploadHeader.H_PONo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and SupplierID = '" & PO.H_SupplierID & "'"
                                Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                Dim ds5 As New DataSet
                                sqlDA5.Fill(ds5)

                                If ds5.Tables(0).Rows.Count > 0 Then
                                    If Not IsDBNull(ds5.Tables(0).Rows(0)("AffiliateApproveDate")) Then
                                        If ls_error = "" Then
                                            ls_error = "PO No. Already exists!"
                                        Else
                                            ls_error = ls_error + ", PO No. Already exists!"
                                        End If
                                    End If
                                End If

                                '02.3 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.UploadPO WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and PONo = '" & dtUploadHeader.H_PONo & "' and SupplierID = '" & PO.H_SupplierID & "'"
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                If ds4.Tables(0).Rows.Count > 0 Then
                                    ls_sql = "delete UploadPO where PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and PONo = '" & dtUploadHeader.H_PONo & "' and SupplierID = '" & PO.H_SupplierID & "'"
                                    Dim sqlComm1 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm1.ExecuteNonQuery()
                                    sqlComm1.Dispose()
                                End If

                                ls_sql = " INSERT INTO [dbo].[UploadPO] " & vbCrLf & _
                                          "            ([AffiliateID],[PONo],[Period],[KanbanCls],[ShipCls],[PartNo],[SupplierID],[POQty] " & vbCrLf & _
                                          "            ,[ForecastN1],[ForecastN2],[ForecastN3] " & vbCrLf & _
                                          "            ,[DeliveryD1],[DeliveryD2],[DeliveryD3],[DeliveryD4],[DeliveryD5],[DeliveryD6],[DeliveryD7],[DeliveryD8],[DeliveryD9],[DeliveryD10] " & vbCrLf & _
                                          "            ,[DeliveryD11],[DeliveryD12],[DeliveryD13],[DeliveryD14],[DeliveryD15],[DeliveryD16],[DeliveryD17],[DeliveryD18],[DeliveryD19],[DeliveryD20] " & vbCrLf & _
                                          "            ,[DeliveryD21],[DeliveryD22],[DeliveryD23],[DeliveryD24],[DeliveryD25],[DeliveryD26],[DeliveryD27],[DeliveryD28],[DeliveryD29],[DeliveryD30] " & vbCrLf & _
                                          "            ,[DeliveryD31],[ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & dtUploadHeader.H_AffiliateID & "' " & vbCrLf & _
                                          "            ,'" & dtUploadHeader.H_PONo & "' " & vbCrLf & _
                                          "            ,'" & dtUploadHeader.H_Period & "' " & vbCrLf

                                ls_sql = ls_sql + "            ,'" & dtUploadHeader.H_POKanban & "' " & vbCrLf & _
                                                  "            ,'" & dtUploadHeader.H_ShipBy & "' " & vbCrLf & _
                                                  "            ,'" & PO.D_PartNo & "' " & vbCrLf & _
                                                  "            ,'" & ls_SuppCode & "' " & vbCrLf & _
                                                  "            ," & PO.D_POQty & " " & vbCrLf & _
                                                  "            ," & PO.Forecast1 & " " & vbCrLf & _
                                                  "            ," & PO.Forecast2 & " " & vbCrLf & _
                                                  "            ," & PO.Forecast3 & " " & vbCrLf & _
                                                  "            ," & PO.D_D1 & " " & vbCrLf & _
                                                  "            ," & PO.D_D2 & " " & vbCrLf & _
                                                  "            ," & PO.D_D3 & " " & vbCrLf & _
                                                  "            ," & PO.D_D4 & " " & vbCrLf & _
                                                  "            ," & PO.D_D5 & " " & vbCrLf

                                ls_sql = ls_sql + "            ," & PO.D_D6 & " " & vbCrLf & _
                                                  "            ," & PO.D_D7 & " " & vbCrLf & _
                                                  "            ," & PO.D_D8 & "" & vbCrLf & _
                                                  "            ," & PO.D_D9 & " " & vbCrLf & _
                                                  "            ," & PO.D_D10 & " " & vbCrLf & _
                                                  "            ," & PO.D_D11 & " " & vbCrLf & _
                                                  "            ," & PO.D_D12 & " " & vbCrLf & _
                                                  "            ," & PO.D_D13 & " " & vbCrLf & _
                                                  "            ," & PO.D_D14 & " " & vbCrLf & _
                                                  "            ," & PO.D_D15 & " " & vbCrLf & _
                                                  "            ," & PO.D_D16 & " " & vbCrLf

                                ls_sql = ls_sql + "            ," & PO.D_D17 & " " & vbCrLf & _
                                                  "            ," & PO.D_D18 & "" & vbCrLf & _
                                                  "            ," & PO.D_D19 & " " & vbCrLf & _
                                                  "            ," & PO.D_D20 & " " & vbCrLf & _
                                                  "            ," & PO.D_D21 & " " & vbCrLf & _
                                                  "            ," & PO.D_D22 & " " & vbCrLf & _
                                                  "            ," & PO.D_D23 & " " & vbCrLf & _
                                                  "            ," & PO.D_D24 & " " & vbCrLf & _
                                                  "            ," & PO.D_D25 & " " & vbCrLf & _
                                                  "            ," & PO.D_D26 & " " & vbCrLf & _
                                                  "            ," & PO.D_D27 & " " & vbCrLf

                                ls_sql = ls_sql + "            ," & PO.D_D28 & "" & vbCrLf & _
                                                  "            ," & PO.D_D29 & " " & vbCrLf & _
                                                  "            ," & PO.D_D30 & " " & vbCrLf & _
                                                  "            ," & PO.D_D31 & " " & vbCrLf & _
                                                  "            ,'" & ls_error & "') " & vbCrLf
                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            Next
                            sqlTran.Commit()

                            Session("PONoUpload") = dtUploadHeader.H_PONo

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
            Else
                If FileName = "" Then
                    lblInfo.Text = "[9999] Please choose the file!"
                    up_GridLoadWhenEventChange()
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())            
        End Try
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
        Dim i As Integer, j As Integer
        'Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        'Dim ls_PONo As String = ""
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        Dim ls_Period As Date
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        'Dim ls_DoubleSupplier As Boolean = False
        'Dim ls_TempSupplierID As String = ""
        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "ErrorCls").ToString.Trim <> "" Then
                    ls_Check = True
                    Exit For
                End If
            Next i

            'dinonaktifkan 2015-09-10
            'Dim countSupplier As Integer = 0

            'For i = 0 To grid.VisibleRowCount - 1
            '    If i = 0 Then
            '        ls_TempSupplierID = grid.GetRowValues(i, "SupplierID").ToString.Trim
            '        countSupplier = 1
            '    End If

            '    If ls_TempSupplierID <> grid.GetRowValues(i, "SupplierID").ToString.Trim Then
            '        ls_DoubleSupplier = True
            '        ls_TempSupplierID = grid.GetRowValues(i, "SupplierID").ToString.Trim
            '        countSupplier = countSupplier + 1
            '    End If
            'Next i

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
                '2.1 delete data 
                Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                SQLCom.Connection = SqlCon
                SQLCom.Transaction = SqlTran
                Dim ls_POAsli As String = Trim(Session("PONoUpload"))

                'If ls_DoubleSupplier = True Then
                '    If countSupplier < 10 Then
                '        ls_Sql = "delete PO_Detail where SUBSTRING(RTRIM(PONo),1,LEN(RTRIM(PONo))-2) = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "'"

                '        SQLCom.CommandText = ls_Sql
                '        SQLCom.ExecuteNonQuery()

                '        ls_Sql = "delete PO_Master where SUBSTRING(RTRIM(PONo),1,LEN(RTRIM(PONo))-2) = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "'"

                '        SQLCom.CommandText = ls_Sql
                '        SQLCom.ExecuteNonQuery()
                '    Else
                '        ls_Sql = "delete PO_Detail where SUBSTRING(RTRIM(PONo),1,LEN(RTRIM(PONo))-3) = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "'"

                '        SQLCom.CommandText = ls_Sql
                '        SQLCom.ExecuteNonQuery()

                '        ls_Sql = "delete PO_Master where SUBSTRING(RTRIM(PONo),1,LEN(RTRIM(PONo))-3) = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "'"

                '        SQLCom.CommandText = ls_Sql
                '        SQLCom.ExecuteNonQuery()
                '    End If

                'Else
                '    ls_Sql = "delete PO_Detail where PONo = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "'"

                '    SQLCom.CommandText = ls_Sql
                '    SQLCom.ExecuteNonQuery()

                '    ls_Sql = "delete PO_Master where PONo = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "'"

                '    SQLCom.CommandText = ls_Sql
                '    SQLCom.ExecuteNonQuery()
                'End If

                ls_Sql = "select * from PO_Master where PONo = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "' and AffiliateApproveDate is not null"

                SQLCom.CommandText = ls_Sql
                Dim da10 As New SqlDataAdapter(SQLCom)
                Dim ds10 As New DataSet
                da10.Fill(ds10)

                Dim ls_tempSupp As String = ""

                If ds10.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds10.Tables(0).Rows.Count - 1
                        If ls_tempSupp = "" Then
                            ls_tempSupp = "'" & ds10.Tables(0).Rows(i)("SupplierID").ToString.Trim & "'"
                        Else
                            ls_tempSupp = ls_tempSupp & ",'" & ds10.Tables(0).Rows(i)("SupplierID").ToString.Trim & "'"
                        End If
                    Next
                End If

                If ls_tempSupp = "" Then
                    ls_tempSupp = "''"
                End If

                ls_Sql = "delete PO_Detail where RTRIM(PONo) = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID not in (" & ls_tempSupp & ")"

                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                ls_Sql = "delete PO_Master where RTRIM(PONo) = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID not in (" & ls_tempSupp & ")"

                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                '2.2 Insert New Detail Data
                Dim ls_SuppCount As Integer = 1
                For i = 0 To grid.VisibleRowCount - 1
                    'If grid.GetRowValues(i, "POQty") <> 0 Then
                    'If ls_DoubleSupplier = True Then
                    '    If i = 0 Then
                    '        ls_TempSupplierID = grid.GetRowValues(i, "SupplierID").ToString.Trim
                    '        ls_POAsli = grid.GetRowValues(i, "PONo").ToString.Trim & "-" & ls_SuppCount
                    '    End If
                    '    If ls_TempSupplierID <> grid.GetRowValues(i, "SupplierID").ToString.Trim Then
                    '        ls_TempSupplierID = grid.GetRowValues(i, "SupplierID").ToString.Trim
                    '        ls_SuppCount = ls_SuppCount + 1
                    '        ls_POAsli = grid.GetRowValues(i, "PONo").ToString.Trim & "-" & ls_SuppCount
                    '    End If
                    'Else
                    ls_POAsli = grid.GetRowValues(i, "PONo").ToString.Trim
                    'End If
                    ls_Sql = " INSERT INTO [dbo].[PO_Detail] " & vbCrLf & _
                          "            ([PONo] " & vbCrLf & _
                          "            ,[AffiliateID] " & vbCrLf & _
                          "            ,[SupplierID] " & vbCrLf & _
                          "            ,[PartNo] " & vbCrLf & _
                          "            ,[KanbanCls] " & vbCrLf & _
                          "            ,[POQty] " & vbCrLf

                    ls_Sql = ls_Sql + "            ,[DeliveryD1] " & vbCrLf & _
                                      "            ,[DeliveryD2] " & vbCrLf & _
                                      "            ,[DeliveryD3] " & vbCrLf & _
                                      "            ,[DeliveryD4] " & vbCrLf & _
                                      "            ,[DeliveryD5] " & vbCrLf & _
                                      "            ,[DeliveryD6] " & vbCrLf & _
                                      "            ,[DeliveryD7] " & vbCrLf & _
                                      "            ,[DeliveryD8] "

                    ls_Sql = ls_Sql + "            ,[DeliveryD9] " & vbCrLf & _
                                      "            ,[DeliveryD10] " & vbCrLf & _
                                      "            ,[DeliveryD11] " & vbCrLf & _
                                      "            ,[DeliveryD12] " & vbCrLf & _
                                      "            ,[DeliveryD13] " & vbCrLf & _
                                      "            ,[DeliveryD14] " & vbCrLf & _
                                      "            ,[DeliveryD15] " & vbCrLf & _
                                      "            ,[DeliveryD16] " & vbCrLf & _
                                      "            ,[DeliveryD17] " & vbCrLf & _
                                      "            ,[DeliveryD18] " & vbCrLf & _
                                      "            ,[DeliveryD19] "

                    ls_Sql = ls_Sql + "            ,[DeliveryD20] " & vbCrLf & _
                                      "            ,[DeliveryD21] " & vbCrLf & _
                                      "            ,[DeliveryD22] " & vbCrLf & _
                                      "            ,[DeliveryD23] " & vbCrLf & _
                                      "            ,[DeliveryD24] " & vbCrLf & _
                                      "            ,[DeliveryD25] " & vbCrLf & _
                                      "            ,[DeliveryD26] " & vbCrLf & _
                                      "            ,[DeliveryD27] " & vbCrLf & _
                                      "            ,[DeliveryD28] " & vbCrLf & _
                                      "            ,[DeliveryD29] " & vbCrLf & _
                                      "            ,[DeliveryD30] "

                    ls_Sql = ls_Sql + "            ,[DeliveryD31] " & vbCrLf & _
                                      "            ,[ForecastN1] " & vbCrLf & _
                                      "            ,[ForecastN2] " & vbCrLf & _
                                      "            ,[ForecastN3] " & vbCrLf & _
                                      "            ,[EntryDate] " & vbCrLf & _
                                      "            ,[EntryUser]) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & ls_POAsli & "' " & vbCrLf & _
                                      "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "SupplierID").ToString & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "KanbanCls").ToString & "' "

                    ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "POQty").ToString & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD1").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD1").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD2").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD2").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD3").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD3").ToString) & "' "

                    ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD4").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD4").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD5").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD5").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD6").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD6").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD7").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD7").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD8").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD8").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD9").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD9").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD10").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD10").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD11").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD11").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD12").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD12").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD13").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD13").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD14").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD14").ToString) & "' "

                    ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD15").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD15").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD16").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD16").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD17").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD17").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD18").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD18").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD19").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD19").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD20").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD20").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD21").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD21").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD22").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD22").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD23").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD23").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD24").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD24").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD25").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD25").ToString) & "' "

                    ls_Sql = ls_Sql + "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD26").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD26").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD27").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD27").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD28").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD28").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD29").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD29").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD30").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD30").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "DeliveryD31").ToString = "", vbNull, grid.GetRowValues(i, "DeliveryD31").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "ForecastN1").ToString = "", vbNull, grid.GetRowValues(i, "ForecastN1").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "ForecastN2").ToString = "", vbNull, grid.GetRowValues(i, "ForecastN2").ToString) & "' " & vbCrLf & _
                                      "            ,'" & IIf(grid.GetRowValues(i, "ForecastN3").ToString = "", vbNull, grid.GetRowValues(i, "ForecastN3").ToString) & "' " & vbCrLf & _
                                      "            , getdate() " & vbCrLf & _
                                      "            ,'" & Session("UserID") & "' ) "

                    ls_Period = grid.GetRowValues(i, "Period2").ToString
                    ls_SupplierID = grid.GetRowValues(i, "SupplierID").ToString
                    ls_ShipBy = grid.GetRowValues(i, "ShipCls").ToString

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()
                    ls_MsgID = "1001"
                    ls_Detail = "ada"


                    ls_Sql = "select * from PO_Master where PONo = '" & ls_POAsli & "' and AffiliateID = '" & Session("AffiliateID") & "' and SupplierID = '" & ls_SupplierID & "'"

                    SQLCom.CommandText = ls_Sql
                    Dim da7 As New SqlDataAdapter(SQLCom)
                    Dim ds7 As New DataSet
                    da7.Fill(ds7)

                    If ds7.Tables(0).Rows.Count = 0 Then
                        ls_Sql = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                                  "            ([PONo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[Period] " & vbCrLf & _
                                  "            ,[CommercialCls] " & vbCrLf & _
                                  "            ,[ShipCls] " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & ls_POAsli & "' " & vbCrLf & _
                                  "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                                  "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                                  "            ,'" & ls_Period & "' " & vbCrLf & _
                                  "            ,'1' " & vbCrLf & _
                                  "            ,'" & ls_ShipBy & "' " & vbCrLf & _
                                  "            ,getdate() " & vbCrLf & _
                                  "            ,'" & Session("UserID") & "') "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    End If
                    'End If

                    ls_Period = grid.GetRowValues(i, "Period2").ToString

                    '2.2.2 Untuk Forecast
                    For j = 0 To 2
                        Dim ls_dateAdd As Date = DateAdd(DateInterval.Month, j + 1, ls_Period)
                        'ls_Sql = "delete MS_Forecast where PartNo = '" & grid.GetRowValues(i, "PartNo").ToString & "' and AffiliateID = '" & Session("AffiliateID") & "' and Period = '" & ls_dateAdd & "'"
                        'SQLCom.CommandText = ls_Sql
                        'SQLCom.ExecuteNonQuery()

                        ls_Sql = " IF NOT EXISTS (select * from MS_Forecast where PartNo = '" & grid.GetRowValues(i, "PartNo").ToString & "' and AffiliateID = '" & Session("AffiliateID") & "' and Period = '" & ls_dateAdd & "')" & vbCrLf & _
                                    " BEGIN" & vbCrLf & _
                                    " INSERT INTO [MS_Forecast] " & vbCrLf & _
                                    "            ([AffiliateID] " & vbCrLf & _
                                    "            ,[PartNo] " & vbCrLf & _
                                    "            ,[Period] " & vbCrLf & _
                                    "            ,[Qty] " & vbCrLf & _
                                    "            ,[EntryDate] " & vbCrLf & _
                                    "            ,[EntryUser]) " & vbCrLf & _
                                    "      VALUES " & vbCrLf & _
                                    "            ('" & Session("AffiliateID") & "','" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                    "            ,'" & ls_dateAdd & "' " & vbCrLf & _
                                    "            ,'" & grid.GetRowValues(i, "ForecastN" & j + 1).ToString & "' " & vbCrLf & _
                                    "            ,getdate() " & vbCrLf & _
                                    "            ,'" & Session("UserID") & "') " & vbCrLf & _
                                    " END " & vbCrLf & _
                                    " ELSE " & vbCrLf & _
                                    " BEGIN" & vbCrLf & _
                                    "      UPDATE [dbo].[MS_Forecast] SET " & vbCrLf & _
                                    "       [Qty] =  " & grid.GetRowValues(i, "ForecastN" & j + 1).ToString & " " & vbCrLf & _
                                    "      WHERE PartNo = '" & grid.GetRowValues(i, "PartNo").ToString & "' and AffiliateID = '" & Session("AffiliateID") & "' and Period = '" & ls_dateAdd & "'" & vbCrLf & _
                                    " END"
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                        ls_MsgID = "1001"
                    Next

                    ls_Sql = " select a.SupplierID, isnull(b.MonthlyInjectionCapacity,0)MonthlyProductionCapacity from [dbo].[MS_PartMapping] a " & vbCrLf & _
                              " left join [dbo].[MS_SupplierCapacity] b on a.SupplierID = b.SupplierID and a.PartNo = b.PartNo" & vbCrLf & _
                              " WHERE a.PartNo = '" & Trim(grid.GetRowValues(i, "PartNo").ToString) & "' AND a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                              " ORDER BY a.SupplierID  " & vbCrLf

                    SQLCom.CommandText = ls_Sql
                    Dim da8 As New SqlDataAdapter(SQLCom)
                    Dim ds8 As New DataSet
                    da8.Fill(ds8)

                    For j = 0 To ds8.Tables(0).Rows.Count - 1
                        Dim lk_SupplierID As String = ds8.Tables(0).Rows(0)("SupplierID").ToString.Trim

                        ls_Sql = "SELECT * FROM dbo.RemainingCapacity WHERE Period ='" & Format(ls_Period, "yyyyMM") & "' AND PartNo = '" & Trim(grid.GetRowValues(i, "PartNo").ToString) & "' AND SupplierID = '" & lk_SupplierID & "'"

                        SQLCom.CommandText = ls_Sql
                        Dim da9 As New SqlDataAdapter(SQLCom)
                        Dim ds9 As New DataSet
                        da9.Fill(ds9)

                        If ds9.Tables(0).Rows.Count = 0 Then 'If pIsUpdate = False Then
                            'INSERT DATA
                            ls_Sql = " INSERT INTO [dbo].[RemainingCapacity] " & vbCrLf & _
                                  "            ([Period] " & vbCrLf & _
                                  "            ,[PartNo] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[QtyRemaining]) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & Format(ls_Period, "yyyyMM") & "' " & vbCrLf & _
                                  "            ,'" & Trim(grid.GetRowValues(i, "PartNo").ToString) & "' " & vbCrLf & _
                                  "            ,'" & lk_SupplierID & "'" & vbCrLf & _
                                  "            ,'" & ds8.Tables(0).Rows(0)("MonthlyProductionCapacity").ToString.Trim & "' ) "
                            SQLCom.CommandText = ls_Sql
                            SQLCom.ExecuteNonQuery()
                        End If
                    Next
                Next i

                '2.3 Insert data to Master
                'If ls_Detail <> "" Then
                '    ls_Sql = " INSERT INTO [dbo].[PO_Master] " & vbCrLf & _
                '          "            ([PONo] " & vbCrLf & _
                '          "            ,[AffiliateID] " & vbCrLf & _
                '          "            ,[SupplierID] " & vbCrLf & _
                '          "            ,[Period] " & vbCrLf & _
                '          "            ,[CommercialCls] " & vbCrLf & _
                '          "            ,[ShipCls] " & vbCrLf & _
                '          "            ,[EntryDate] " & vbCrLf & _
                '          "            ,[EntryUser]) " & vbCrLf & _
                '          "      VALUES " & vbCrLf & _
                '          "            ('" & Session("PONoUpload") & "' " & vbCrLf & _
                '          "            ,'" & Session("AffiliateID") & "' " & vbCrLf & _
                '          "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                '          "            ,'" & ls_Period & "' " & vbCrLf & _
                '          "            ,'1' " & vbCrLf & _
                '          "            ,'" & ls_ShipBy & "' " & vbCrLf & _
                '          "            ,getdate() " & vbCrLf & _
                '          "            ,'" & Session("UserID") & "') "

                '    SQLCom.CommandText = ls_Sql
                '    SQLCom.ExecuteNonQuery()
                'End If

                '2.3.1 Habis save semua,.. delete tada di tempolary table
                ls_Sql = "delete UploadPO where AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & Session("PONoUpload") & "'"

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