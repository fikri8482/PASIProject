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

Public Class POUploadExport
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "I03"
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

        ls_AllowUpdate = True 'clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            If Session("buttonSubMenu") = "Direct" Then btnSubMenu.Text = "Back"
            If ls_AllowUpdate = False Then
                btnUpload.Enabled = False
                btnClear.Enabled = False
                btnSave.Enabled = False
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

        If btnSubMenu.Text = "Back" Then
            Session.Remove("buttonSubMenu")
            Response.Redirect("~/PurchaseOrderExport/POExportEntryMonthly.aspx")
        Else
            Session.Remove("buttonSubMenu")
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Uploader.NullText = "Click here to browse files..."

        lblInfo.Text = ""

        Uploader.Enabled = True
        btnSave.Enabled = True        
        btnUpload.Enabled = True

        up_GridLoadWhenEventChange()
        Session.Remove("Period")
        Session.Remove("PONoUpload")
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        up_Import()
    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("ErrorCls") > "0" Then
            e.Cell.BackColor = Color.Red
        End If

        If InStr(1, e.GetValue("ErrorDesc"), "PO Already Final Approval") = 1 Then
            e.Cell.BackColor = Color.Red
        End If

        If InStr(1, e.GetValue("ErrorDesc"), "PO Already Approve by Supplier") = 1 Then
            e.Cell.BackColor = Color.Yellow
        End If

        If InStr(1, e.GetValue("ErrorDesc"), "PO Already Send to Supplier") = 1 Then
            e.Cell.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub grid_CustomCallback(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Dim ls_Check As Boolean = False
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    'For i = 0 To grid.VisibleRowCount - 1
                    '    If InStr(1, grid.GetRowValues(i, "ErrorDesc"), "PO Already Approve by Supplier") = 1 Or _
                    '        InStr(1, grid.GetRowValues(i, "ErrorDesc"), "PO Already Send to Supplier") = 1 Then
                    '        ls_Check = True
                    '        Exit For
                    '    End If
                    '    'If
                    'Next i

                    'If ls_Check = True Then
                    '    popUp2.ShowOnPageLoad = True
                    'Else

                    'End If
                    up_Save()
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

            'ls_SQL = " select  " & vbCrLf & _
            '          " 	row_number() over (order by PONo, AffiliateID asc) as NoUrut, " & vbCrLf & _
            '          " 	AffiliateID, SupplierID, EmergencyCls, ShipCls, PONo, Period, " & vbCrLf & _
            '          " 	ETDVendor1, ETDPort1, ETAPort1, ETAFactory1, " & vbCrLf & _
            '          " 	sum(case when ErrorCls <> '' then 1 else 0 end) ErrorCls " & vbCrLf & _
            '          " from UploadPOExport a  " & vbCrLf & _
            '          " where PONo in (" & Session("PONoUpload") & ")" & vbCrLf & _
            '          " group by AffiliateID, SupplierID, EmergencyCls, ShipCls, PONo, Period, ETDVendor1, ETDPort1, ETAPort1, ETAFactory1 " & vbCrLf

            ls_SQL = "  select   " & vbCrLf & _
                  "  	row_number() over (order by a.PONo, a.AffiliateID asc) as NoUrut,  " & vbCrLf & _
                  "  	a.AffiliateID, a.SupplierID, a.EmergencyCls, a.ShipCls, a.PONo, a.Period,  " & vbCrLf & _
                  "  	a.ETDVendor1, a.ETDPort1, a.ETAPort1, a.ETAFactory1,  " & vbCrLf & _
                  "  	sum(case when ErrorCls <> '' then 1 else 0 end) ErrorCls, " & vbCrLf & _
                  " 	ISNULL(CASE WHEN b.PASIApproveDate IS NOT NULL THEN 'PO Already Final Approval'  " & vbCrLf & _
                  " 	ELSE CASE WHEN b.SupplierApproveUser IS NOT NULL THEN 'PO Already Approve by Supplier'  " & vbCrLf & _
                  " 		 ELSE CASE WHEN b.PASISendToSupplierDate IS NOT NULL THEN 'PO Already Send to Supplier'  " & vbCrLf & _
                  " 			  ELSE (select top 1 xc.ErrorCls from UploadPOExport xc  " & vbCrLf & _
                  " 					where xc.AffiliateID = a.AffiliateID and a.SupplierID = xc.SupplierID and xc.PONo = a.PONo and ErrorCls <>'') " & vbCrLf & _
                  " 			  END  "

            ls_SQL = ls_SQL + " 		 END " & vbCrLf & _
                              " 	END,'') ErrorDesc  " & vbCrLf & _
                              "  from UploadPOExport a  " & vbCrLf & _
                              "  left join PO_Master_Export b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PONo = b.PONo  " & vbCrLf & _
                              "  where a.PONo in (" & Session("PONoUpload") & ")" & vbCrLf & _
                              "  group by a.AffiliateID, a.SupplierID, a.EmergencyCls, a.ShipCls, a.PONo, a.Period,  " & vbCrLf & _
                              " 		  a.ETDVendor1, a.ETDPort1, a.ETAPort1, a.ETAFactory1, b.PASIApproveDate, b.SupplierApproveUser, b.PASISendToSupplierDate "


            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' AffiliateID, '' EmergencyCls, '' ShipCls, '' PONo, '' Period, '' ETDVendor1, " & vbCrLf & _
                     " '' ETDPort1, '' ETAPort1, '' ETAFactory1, '' ErrorCls, '' ErrorDesc, '' SupplierID"

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
        Dim ls_MOQ As Double = 0
        Dim ls_QtyBox As Double = 0
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = ""

        'Header 
        Dim pEmergencyCls As String = ""
        Dim pPeriod As Date
        Dim pAffiliateID As String = ""
        Dim pConsigneeCode As String = ""
        Dim pCommercial As String = ""
        Dim pShipBy As String = ""
        Dim pForwarderID As String = ""
        Dim pDestinationPort As String = ""

        Try
            Session.Remove("tempSupplierID")
            lblInfo.ForeColor = Color.Red
            If (Not Uploader.PostedFile Is Nothing) And (Uploader.PostedFile.ContentLength > 0) Then
                FileName = Path.GetFileName(Uploader.PostedFile.FileName)
                FileExt = Path.GetExtension(Uploader.PostedFile.FileName)
                FilePath = Ext & "\Import\" & FileName
                Dim fi As New FileInfo(Server.MapPath("~\Import\" & FileName))
                If fi.Exists Then
                    fi.Delete()
                    fi = New FileInfo(Server.MapPath("~\Import\" & FileName))
                End If
                Uploader.SaveAs(FilePath)

                Dim connStr As String = ""
                Select Case FileExt
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

                        '001. Check Header
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A1:C6]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dt)

                        If dt.Rows.Count > 0 Then
                            '01. Emergency Cls
                            If IsDBNull(dt.Rows(0).Item(2)) Then
                                lblInfo.Text = "[9999] Please input PO Type!, check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            Else
                                If dt.Rows(0).Item(2).ToString.Trim.ToUpper <> "M" And dt.Rows(0).Item(2).ToString.Trim.ToUpper <> "E" Then
                                    lblInfo.Text = "[9999] Invalid PO Type, must be fill with ""M"" or ""E"" , please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If
                                pEmergencyCls = dt.Rows(0).Item(2).ToString.ToUpper
                            End If

                            '02. Period
                            If dt.Rows(0).Item(2).ToString.Trim.ToUpper = "M" Then
                                If IsDBNull(dt.Rows(1).Item(2)) Then
                                    lblInfo.Text = "[9999] Please input Period, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                Else
                                    Try
                                        pPeriod = dt.Rows(1).Item(2) & "-01"
                                    Catch ex As Exception
                                        lblInfo.Text = "[9999] Invalid Format Period, please check the file again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End Try
                                End If
                            Else
                                Try
                                    pPeriod = dt.Rows(1).Item(2) & "-01"
                                Catch ex As Exception
                                    lblInfo.Text = "[9999] Invalid Format Period, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End Try
                            End If

                            '03. AffiliateID
                            If IsDBNull(dt.Rows(2).Item(2)) Then
                                lblInfo.Text = "[9999] Please input Affiliate Code!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            Else
                                '03.1 Check Affiliate di PO MASTER
                                ls_sql = "SELECT * FROM MS_Affiliate WHERE AffiliateID = '" & dt.Rows(2).Item(2) & "'"
                                Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn)
                                Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                Dim ds5 As New DataSet
                                sqlDA5.Fill(ds5)

                                If ds5.Tables(0).Rows.Count = 0 Then
                                    lblInfo.Text = "[9999] Affiliate Code not valid, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If

                                pAffiliateID = dt.Rows(2).Item(2)

                                If ds5.Tables(0).Rows.Count > 0 Then
                                    If Trim(dt.Rows(3).Item(2).ToString.Trim.ToUpper) <> Trim(ds5.Tables(0).Rows(0)("ConsigneeCode")) Then
                                        lblInfo.Text = "[9999] Consignee Code not valid, please check the file again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End If
                                    pConsigneeCode = dt.Rows(3).Item(2)
                                End If
                            End If

                            '02.1 Check Cutofdate
                            ls_sql = "select MAX(CutOfDate) CutOfDate from MS_ETD_Export where AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                     "and Period = '" & pPeriod & "'" & vbCrLf
                            Dim sqlCmd As New SqlCommand(ls_sql, sqlConn)
                            Dim sqlDA As New SqlDataAdapter(sqlCmd)
                            Dim ds As New DataSet
                            sqlDA.Fill(ds)

                            If ds.Tables(0).Rows.Count > 0 And Not IsDBNull(ds.Tables(0).Rows(0)("CutOfDate")) Then
                                If Format(ds.Tables(0).Rows(0)("CutOfDate"), "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                                    lblInfo.Text = "[9999] Can't upload this period, because period is bigger than Cut of PO Date!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If
                            Else
                                lblInfo.Text = "[9999] Can't upload this period, because Cut of PO Date not found"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End If

                            '04. Commercial
                            If IsDBNull(dt.Rows(4).Item(2)) Then
                                lblInfo.Text = "[9999] Please Input Commercial, check file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            Else
                                If dt.Rows(4).Item(2).ToString.Trim.ToUpper <> "YES" And dt.Rows(4).Item(2).ToString.Trim.ToUpper <> "NO" Then
                                    lblInfo.Text = "[9999] Invalid Commercial, must be fill with ""Yes"" or ""No"" , please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If

                                If dt.Rows(4).Item(2).ToString.Trim.ToUpper = "YES" Then
                                    pCommercial = "1"
                                Else
                                    pCommercial = "0"
                                End If

                            End If

                            '05. ShipBy
                            If IsDBNull(dt.Rows(5).Item(2)) Then
                                lblInfo.Text = "[9999] Please Input ShipBy, check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            Else
                                If dt.Rows(5).Item(2).ToString.Trim.ToUpper <> "A" And dt.Rows(5).Item(2).ToString.Trim.ToUpper <> "B" Then
                                    lblInfo.Text = "[9999] Invalid Ship by, must be fill with ""A"" or ""B"" , please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If
                                pShipBy = dt.Rows(5).Item(2)
                            End If

                            '06. Get ForwarderID
                            ls_sql = "select * from MS_ForwarderMapping where AffiliateID = '" & pAffiliateID & "' and ShipCls = '" & pShipBy & "'"
                            Dim sqlCmd11 As New SqlCommand(ls_sql, sqlConn)
                            Dim sqlDA11 As New SqlDataAdapter(sqlCmd11)
                            Dim ds11 As New DataSet
                            sqlDA11.Fill(ds11)

                            If ds11.Tables(0).Rows.Count > 0 Then
                                pForwarderID = ds11.Tables(0).Rows(0)("ForwarderID")
                            Else
                                '06. Get DEfault ForwarderID jika tidak ada mapping
                                ls_sql = "select * from MS_Forwarder where DefaultCls = '1'"
                                Dim sqlCmd76 As New SqlCommand(ls_sql, sqlConn)
                                Dim sqlDA76 As New SqlDataAdapter(sqlCmd76)
                                Dim ds76 As New DataSet
                                sqlDA76.Fill(ds76)
                                If ds76.Tables(0).Rows.Count > 0 Then
                                    pForwarderID = ds76.Tables(0).Rows(0)("ForwarderID")
                                Else
                                    pForwarderID = ""
                                End If
                            End If

                            ''07. Get Destination PORT
                            'ls_sql = "select DestinationPort from MS_DestinationExport where AffiliateID = '" & pAffiliateID & "' and ShipByCls = '" & pShipBy & "' and DefaultCls = '1'"
                            'Dim sqlCmd12 As New SqlCommand(ls_sql, sqlConn)
                            'Dim sqlDA12 As New SqlDataAdapter(sqlCmd12)
                            'Dim ds12 As New DataSet
                            'sqlDA12.Fill(ds12)

                            'If ds12.Tables(0).Rows.Count > 0 Then
                            '    pDestinationPort = ds12.Tables(0).Rows(0)("DestinationPort")
                            'Else
                            '    lblInfo.Text = "[9999] Destination Port not found in Master, please Check Destinaton Port Master!"
                            '    grid.JSProperties("cpMessage") = lblInfo.Text
                            '    MyConnection.Close()
                            '    Exit Sub
                            'End If
                        Else
                            lblInfo.Text = "[2023] Please enter week 1 amount!"
                            grid.JSProperties("cpMessage") = lblInfo.Text
                        End If

                        Dim dtUploadHeader As New clsPOExportHeader
                        Dim dtUploadHeaderList As New List(Of clsPOExportHeader)
                        Dim dtUploadDetailList As New List(Of clsPOExportDetail)

                        'Get Header Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "C1:I13]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtHeader)

                        If dtHeader.Rows.Count > 0 Then
                            '07. Get OrderNo1
                            If IsDBNull(dtHeader.Rows(8).Item(2)) = False Then
                                dtUploadHeader.H_OrderNo1 = dtHeader.Rows(8).Item(2)
                                Try
                                    If IsDBNull(dtHeader.Rows(9).Item(2)) Or IsDBNull(dtHeader.Rows(10).Item(2)) _
                                        Or IsDBNull(dtHeader.Rows(11).Item(2)) Or IsDBNull(dtHeader.Rows(12).Item(2)) Then
                                        lblInfo.Text = "[9999] Please Input ETD or ETA in Week1, please check file excel again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End If
                                    dtUploadHeader.H_ETDVendor1 = IIf(IsDBNull(dtHeader.Rows(9).Item(2)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(9).Item(2))
                                    dtUploadHeader.H_ETDPort1 = IIf(IsDBNull(dtHeader.Rows(10).Item(2)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(10).Item(2))
                                    dtUploadHeader.H_ETAPort1 = IIf(IsDBNull(dtHeader.Rows(11).Item(2)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(11).Item(2))
                                    dtUploadHeader.H_ETAFactory1 = IIf(IsDBNull(dtHeader.Rows(12).Item(2)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(12).Item(2))
                                Catch ex As Exception
                                    lblInfo.Text = "[9999] Invalid Format ETD or ETA in Week1, please check file excel again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End Try
                            End If

                            '08. Get OrderNo2
                            If IsDBNull(dtHeader.Rows(8).Item(3)) = False Then
                                dtUploadHeader.H_OrderNo2 = dtHeader.Rows(8).Item(3)
                                Try
                                    If IsDBNull(dtHeader.Rows(9).Item(3)) Or IsDBNull(dtHeader.Rows(10).Item(3)) _
                                        Or IsDBNull(dtHeader.Rows(11).Item(3)) Or IsDBNull(dtHeader.Rows(12).Item(3)) Then
                                        lblInfo.Text = "[9999] Please Input ETD or ETA in Week2, please check file excel again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End If
                                    dtUploadHeader.H_ETDVendor2 = IIf(IsDBNull(dtHeader.Rows(9).Item(3)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(9).Item(3))
                                    dtUploadHeader.H_ETDPort2 = IIf(IsDBNull(dtHeader.Rows(10).Item(3)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(10).Item(3))
                                    dtUploadHeader.H_ETAPort2 = IIf(IsDBNull(dtHeader.Rows(11).Item(3)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(11).Item(3))
                                    dtUploadHeader.H_ETAFactory2 = IIf(IsDBNull(dtHeader.Rows(12).Item(3)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(12).Item(3))
                                Catch ex As Exception
                                    lblInfo.Text = "[9999] Invalid Format ETD or ETA in Week2, please check file excel again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End Try
                            End If

                            '09. Get OrderNo3
                            If IsDBNull(dtHeader.Rows(8).Item(4)) = False Then
                                dtUploadHeader.H_OrderNo3 = dtHeader.Rows(8).Item(4)
                                Try
                                    If IsDBNull(dtHeader.Rows(9).Item(4)) Or IsDBNull(dtHeader.Rows(10).Item(4)) _
                                        Or IsDBNull(dtHeader.Rows(11).Item(4)) Or IsDBNull(dtHeader.Rows(12).Item(4)) Then
                                        lblInfo.Text = "[9999] Please Input ETD or ETA in Week3, please check file excel again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End If
                                    dtUploadHeader.H_ETDVendor3 = IIf(IsDBNull(dtHeader.Rows(9).Item(4)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(9).Item(4))
                                    dtUploadHeader.H_ETDPort3 = IIf(IsDBNull(dtHeader.Rows(10).Item(4)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(10).Item(4))
                                    dtUploadHeader.H_ETAPort3 = IIf(IsDBNull(dtHeader.Rows(11).Item(4)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(11).Item(4))
                                    dtUploadHeader.H_ETAFactory3 = IIf(IsDBNull(dtHeader.Rows(12).Item(4)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(12).Item(4))
                                Catch ex As Exception
                                    lblInfo.Text = "[9999] Invalid Format ETD or ETA in Week3, please check file excel again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End Try
                            End If

                            '10. Get OrderNo4
                            If IsDBNull(dtHeader.Rows(8).Item(5)) = False Then
                                dtUploadHeader.H_OrderNo4 = dtHeader.Rows(8).Item(5)
                                Try
                                    If IsDBNull(dtHeader.Rows(9).Item(5)) Or IsDBNull(dtHeader.Rows(10).Item(5)) _
                                        Or IsDBNull(dtHeader.Rows(11).Item(5)) Or IsDBNull(dtHeader.Rows(12).Item(5)) Then
                                        lblInfo.Text = "[9999] Please Input ETD or ETA in Week4, please check file excel again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End If
                                    dtUploadHeader.H_ETDVendor4 = IIf(IsDBNull(dtHeader.Rows(9).Item(5)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(9).Item(5))
                                    dtUploadHeader.H_ETDPort4 = IIf(IsDBNull(dtHeader.Rows(10).Item(5)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(10).Item(5))
                                    dtUploadHeader.H_ETAPort4 = IIf(IsDBNull(dtHeader.Rows(11).Item(5)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(11).Item(5))
                                    dtUploadHeader.H_ETAFactory4 = IIf(IsDBNull(dtHeader.Rows(12).Item(5)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(12).Item(5))
                                Catch ex As Exception
                                    lblInfo.Text = "[9999] Invalid Format ETD or ETA in Week4, please check file excel again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End Try
                            End If

                            '11. Get OrderNo5
                            If IsDBNull(dtHeader.Rows(8).Item(6)) = False Then
                                dtUploadHeader.H_OrderNo5 = dtHeader.Rows(8).Item(6)
                                Try
                                    If IsDBNull(dtHeader.Rows(9).Item(6)) Or IsDBNull(dtHeader.Rows(10).Item(6)) _
                                        Or IsDBNull(dtHeader.Rows(11).Item(6)) Or IsDBNull(dtHeader.Rows(12).Item(6)) Then
                                        lblInfo.Text = "[9999] Please Input ETD or ETA in Week5, please check file excel again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        MyConnection.Close()
                                        Exit Sub
                                    End If
                                    dtUploadHeader.H_ETDVendor5 = IIf(IsDBNull(dtHeader.Rows(9).Item(6)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(9).Item(6))
                                    dtUploadHeader.H_ETDPort5 = IIf(IsDBNull(dtHeader.Rows(10).Item(6)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(10).Item(6))
                                    dtUploadHeader.H_ETAPort5 = IIf(IsDBNull(dtHeader.Rows(11).Item(6)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(11).Item(6))
                                    dtUploadHeader.H_ETAFactory5 = IIf(IsDBNull(dtHeader.Rows(12).Item(6)), Format(Now, "yyyy-MM-dd"), dtHeader.Rows(12).Item(6))
                                Catch ex As Exception
                                    lblInfo.Text = "[9999] Invalid Format ETD or ETA in Week5, please check file excel again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End Try
                            End If
                        End If


                        'Get Detail Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B19:M65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsPOExportDetail
                                    dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.D_UOM = IIf(IsDBNull(dtDetail.Rows(i).Item(1)), 0, dtDetail.Rows(i).Item(1))
                                    dtUploadDetail.D_Week1 = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), 0, dtDetail.Rows(i).Item(3))
                                    dtUploadDetail.D_Week2 = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), 0, dtDetail.Rows(i).Item(4))
                                    dtUploadDetail.D_Week3 = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), 0, dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.D_Week4 = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), 0, dtDetail.Rows(i).Item(6))
                                    dtUploadDetail.D_Week5 = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), 0, dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.D_PreviousForecast = 0 'IIf(IsDBNull(dtDetail.Rows(i).Item(7)), 0, dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.D_Forecast1 = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), 0, dtDetail.Rows(i).Item(9))
                                    dtUploadDetail.D_Forecast2 = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), 0, dtDetail.Rows(i).Item(10))
                                    dtUploadDetail.D_Forecast3 = IIf(IsDBNull(dtDetail.Rows(i).Item(11)), 0, dtDetail.Rows(i).Item(11))
                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            '01.01 Delete TempoaryData
                            If dtUploadHeader.H_OrderNo1 <> "" Then
                                ls_sql = "delete UploadPOExport where AffiliateID = '" & pAffiliateID & "' and PONo = '" & dtUploadHeader.H_OrderNo1 & "'"
                                Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm9.ExecuteNonQuery()
                                sqlComm9.Dispose()
                            End If

                            If dtUploadHeader.H_OrderNo2 <> "" Then
                                ls_sql = "delete UploadPOExport where AffiliateID = '" & pAffiliateID & "' and PONo = '" & dtUploadHeader.H_OrderNo2 & "'"
                                Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm9.ExecuteNonQuery()
                                sqlComm9.Dispose()
                            End If

                            If dtUploadHeader.H_OrderNo3 <> "" Then
                                ls_sql = "delete UploadPOExport where AffiliateID = '" & pAffiliateID & "' and PONo = '" & dtUploadHeader.H_OrderNo3 & "'"
                                Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm9.ExecuteNonQuery()
                                sqlComm9.Dispose()
                            End If

                            If dtUploadHeader.H_OrderNo4 <> "" Then
                                ls_sql = "delete UploadPOExport where AffiliateID = '" & pAffiliateID & "' and PONo = '" & dtUploadHeader.H_OrderNo4 & "'"
                                Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm9.ExecuteNonQuery()
                                sqlComm9.Dispose()
                            End If

                            If dtUploadHeader.H_OrderNo5 <> "" Then
                                ls_sql = "delete UploadPOExport where AffiliateID = '" & pAffiliateID & "' and PONo = '" & dtUploadHeader.H_OrderNo5 & "'"
                                Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm9.ExecuteNonQuery()
                                sqlComm9.Dispose()
                            End If

                            Session.Remove("PONoUpload")

                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim ls_UOM As String = ""
                                Dim PO As clsPOExportDetail = dtUploadDetailList(i)

                                Dim ls_Tanggal As Date

                                If dtUploadHeader.H_OrderNo1 <> "" Then
                                    ls_MOQ = 0
                                    ls_QtyBox = 0
                                    ls_SupplierID = ""
                                    ls_error = ""

                                    If Format(dtUploadHeader.H_ETDPort1, "yyyy-MM-dd") <> "0001-01-01" Then
                                        ls_Tanggal = dtUploadHeader.H_ETDPort1
                                    Else
                                        ls_Tanggal = Format(Now, "yyyy-MM-dd")
                                    End If

                                    '99.0 Check PartNo di Ms_PartMapping
                                    ls_sql = "SELECT a.*, b.Description UOM FROM dbo.MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls WHERE a.PartNo = '" & PO.D_PartNo & "'"
                                    Dim sqlCmd20 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA20 As New SqlDataAdapter(sqlCmd20)
                                    Dim ds20 As New DataSet
                                    sqlDA20.Fill(ds20)

                                    If ds20.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        End If
                                    Else
                                        '99.0.1 Check UOM
                                        ls_UOM = ds20.Tables(0).Rows(0)("UOM")
                                        If ls_UOM.Trim.ToUpper <> PO.D_UOM.Trim.ToUpper Then
                                            If ls_error = "" Then
                                                ls_error = PO.D_PartNo.Trim & "UOM not found!, please check the file again!"
                                            Else
                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            End If
                                        End If

                                        '99.1 Check PartNo di Ms_PartMapping
                                        ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                                        Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                        Dim ds3 As New DataSet
                                        sqlDA3.Fill(ds3)

                                        If ds3.Tables(0).Rows.Count = 0 Then
                                            If ls_error = "" Then
                                                ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            Else
                                                ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            End If
                                        Else
                                            ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                            ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")
                                            ls_QtyBox = IIf(IsDBNull(ds3.Tables(0).Rows(0)("QtyBox")), 0, ds3.Tables(0).Rows(0)("QtyBox"))

                                            '99.2 Check ETA ETD                                    
                                            ls_sql = "select * from MS_ETD_Export where AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                                     "and Period = '" & pPeriod & "' and " & vbCrLf & _
                                                     "SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
                                                     "and ETAForwarder = '" & dtUploadHeader.H_ETDVendor1 & "'" & vbCrLf & _
                                                     "and ETDPort = '" & dtUploadHeader.H_ETDPort1 & "'" & vbCrLf & _
                                                     "and ETAPort = '" & dtUploadHeader.H_ETAPort1 & "'" & vbCrLf & _
                                                     "and ETAFactory = '" & dtUploadHeader.H_ETAFactory1 & "'"
                                            Dim sqlCmd71 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                            Dim sqlDA71 As New SqlDataAdapter(sqlCmd71)
                                            Dim ds71 As New DataSet
                                            sqlDA71.Fill(ds71)

                                            If ds71.Tables(0).Rows.Count = 0 Then
                                                If pEmergencyCls = "M" Then
                                                    lblInfo.Text = "[9999] Please registered ETA and ETD for this Period in Time Chart Master!"
                                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                                    MyConnection.Close()
                                                    Exit Sub
                                                End If

                                                '99.3 Check Price di MS_Price
                                                ls_sql = "SELECT Price FROM MS_Price WHERE ('" & Format(ls_Tanggal, "yyyy-MM-dd") & "' between StartDate and EndDate) and AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                                         "                  and PartNo = '" & PO.D_PartNo & "' "
                                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                                Dim ds2 As New DataSet
                                                sqlDA2.Fill(ds2)

                                                If ds2.Tables(0).Rows.Count = 0 Then
                                                    If ls_error = "" Then
                                                        ls_error = "Price " & PO.D_PartNo.Trim & " must be registered"
                                                    Else
                                                        ls_error = ls_error + ", Price " & PO.D_PartNo.Trim & " must be registered!"
                                                    End If
                                                Else
                                                    '99.4 Check data Qtybox
                                                    If PO.D_Week1 > 0 Then
                                                        If PO.D_Week1 >= ls_MOQ Then
                                                            If (PO.D_Week1 Mod ls_QtyBox) <> 0 Then
                                                                If ls_error = "" Then
                                                                    ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                Else
                                                                    ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                End If
                                                            End If
                                                        Else
                                                            If ls_error = "" Then
                                                                ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week1 is less than MOQ!, please check the file again!"
                                                            Else
                                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week1 is less than MOQ!, please check the file again!"
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                '99.3 Check Price di MS_Price
                                                ls_sql = "SELECT Price FROM MS_Price WHERE ('" & Format(ls_Tanggal, "yyyy-MM-dd") & "' between StartDate and EndDate) and AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                                         "                  and PartNo = '" & PO.D_PartNo & "' "
                                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                                Dim ds2 As New DataSet
                                                sqlDA2.Fill(ds2)

                                                If ds2.Tables(0).Rows.Count = 0 Then
                                                    If ls_error = "" Then
                                                        ls_error = "Price " & PO.D_PartNo.Trim & " must be registered"
                                                    Else
                                                        ls_error = ls_error + ", Price " & PO.D_PartNo.Trim & " must be registered!"
                                                    End If
                                                Else
                                                    '99.4 Check data Qtybox
                                                    If PO.D_Week1 > 0 Then
                                                        If PO.D_Week1 >= ls_MOQ Then
                                                            If (PO.D_Week1 Mod ls_QtyBox) <> 0 Then
                                                                If ls_error = "" Then
                                                                    ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                Else
                                                                    ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                End If
                                                            End If
                                                        Else
                                                            If ls_error = "" Then
                                                                ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week1 is less than MOQ!, please check the file again!"
                                                            Else
                                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week1 is less than MOQ!, please check the file again!"
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Check PONO di PO MASTER
                                    ls_sql = "SELECT * FROM PO_Master_Export WHERE PONo = '" & dtUploadHeader.H_OrderNo1 & "' " & vbCrLf & _
                                        "and AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                        "and SupplierID = '" & ls_SupplierID & "' " & vbCrLf

                                    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                    Dim ds5 As New DataSet
                                    sqlDA5.Fill(ds5)

                                    If ds5.Tables(0).Rows.Count > 0 Then
                                        If Not IsDBNull(ds5.Tables(0).Rows(0)("PASIApproveDate")) Then
                                            If ls_error = "" Then
                                                ls_error = "PO No. already Final Approve, can''t replace!"
                                            Else
                                                ls_error = ls_error + ", PO No. already Final Approve, can''t replace!"
                                            End If
                                        Else
                                            If Not IsDBNull(ds5.Tables(0).Rows(0)("SupplierApproveDate")) Then
                                                If ls_error = "" Then
                                                    ls_error = "PO No. already Approve by Supplier!"
                                                Else
                                                    ls_error = ls_error + ", PO No. already Approve by Supplier!"
                                                End If
                                            Else
                                                If Not IsDBNull(ds5.Tables(0).Rows(0)("PASISendToSupplierDate")) Then
                                                    If ls_error = "" Then
                                                        ls_error = "PO No. already send to Supplier!"
                                                    Else
                                                        ls_error = ls_error + ", PO No. already send to Supplier!"
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    If InStr(1, Session("PONoUpload"), dtUploadHeader.H_OrderNo1.Trim) = 0 Then
                                        If Session("PONoUpload") = "" Then
                                            Session("PONoUpload") = "'" & dtUploadHeader.H_OrderNo1.Trim & "'"
                                        Else
                                            Session("PONoUpload") = Session("PONoUpload") & ",'" & dtUploadHeader.H_OrderNo1.Trim & "'"
                                        End If
                                    End If

                                    ls_sql = " INSERT INTO [dbo].[UploadPOExport] " & vbCrLf & _
                                             "            ([PONo],[AffiliateID],[SupplierID],[ForwarderID],[Period],[CommercialCls],[EmergencyCls],[ShipCls] " & vbCrLf & _
                                             "            ,[PartNo],[OrderNo1],[ETDVendor1],[ETDPort1],[ETAPort1],[ETAFactory1],[Week1],[PreviousForecast],[ForecastN1] " & vbCrLf & _
                                             "            ,[ForecastN2],[ForecastN3],[ErrorCls],[UOM]) " & vbCrLf & _
                                             "      VALUES " & vbCrLf & _
                                             "            ('" & dtUploadHeader.H_OrderNo1 & "' " & vbCrLf & _
                                             "            ,'" & pAffiliateID & "' " & vbCrLf & _
                                             "            ,'" & ls_SupplierID & "' " & vbCrLf

                                    ls_sql = ls_sql + "            ,'" & pForwarderID & "' " & vbCrLf & _
                                                      "            ,'" & pPeriod & "' " & vbCrLf & _
                                                      "            ,'" & pCommercial & "' " & vbCrLf & _
                                                      "            ,'" & pEmergencyCls & "' " & vbCrLf & _
                                                      "            ,'" & pShipBy & "' " & vbCrLf & _
                                                      "            ,'" & PO.D_PartNo & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_OrderNo1 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDVendor1 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDPort1 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAPort1 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAFactory1 & "' " & vbCrLf & _
                                                      "            ," & PO.D_Week1 & " " & vbCrLf & _
                                                      "            ,0 " & vbCrLf

                                    ls_sql = ls_sql + "            ," & PO.D_Forecast1 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast2 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast3 & "" & vbCrLf & _
                                                      "            ,'" & ls_error & "'" & vbCrLf & _
                                                      "            ,'" & PO.D_UOM & "') " & vbCrLf

                                    Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                    sqlComm.Dispose()
                                End If

                                If dtUploadHeader.H_OrderNo2 <> "" Then
                                    ls_MOQ = 0
                                    ls_QtyBox = 0
                                    ls_SupplierID = ""
                                    ls_error = ""

                                    If Format(dtUploadHeader.H_ETDPort2, "yyyy-MM-dd") <> "0001-01-01" Then
                                        ls_Tanggal = dtUploadHeader.H_ETDPort2
                                    Else
                                        ls_Tanggal = Format(Now, "yyyy-MM-dd")
                                    End If

                                    '99.0 Check PartNo di Ms_PartMapping
                                    ls_sql = "SELECT a.*, b.Description UOM FROM dbo.MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls WHERE a.PartNo = '" & PO.D_PartNo & "'"
                                    Dim sqlCmd20 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA20 As New SqlDataAdapter(sqlCmd20)
                                    Dim ds20 As New DataSet
                                    sqlDA20.Fill(ds20)

                                    If ds20.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        End If
                                    Else
                                        '99.0.1 Check UOM
                                        ls_UOM = ds20.Tables(0).Rows(0)("UOM")
                                        If ls_UOM.Trim.ToUpper <> PO.D_UOM.Trim.ToUpper Then
                                            If ls_error = "" Then
                                                ls_error = PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            Else
                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            End If
                                        End If

                                        '99.1 Check PartNo di Ms_PartMapping
                                        ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                                        Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                        Dim ds3 As New DataSet
                                        sqlDA3.Fill(ds3)

                                        If ds3.Tables(0).Rows.Count = 0 Then
                                            If ls_error = "" Then
                                                ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            Else
                                                ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            End If
                                        Else
                                            ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                            ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")
                                            ls_QtyBox = IIf(IsDBNull(ds3.Tables(0).Rows(0)("QtyBox")), 0, ds3.Tables(0).Rows(0)("QtyBox"))

                                            '99.2 Check ETA ETD                                    
                                            ls_sql = "select * from MS_ETD_Export where AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                                     "and Period = '" & pPeriod & "' and " & vbCrLf & _
                                                     "SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
                                                     "and ETAForwarder = '" & dtUploadHeader.H_ETDVendor2 & "'" & vbCrLf & _
                                                     "and ETDPort = '" & dtUploadHeader.H_ETDPort2 & "'" & vbCrLf & _
                                                     "and ETAPort = '" & dtUploadHeader.H_ETAPort2 & "'" & vbCrLf & _
                                                     "and ETAFactory = '" & dtUploadHeader.H_ETAFactory2 & "'"
                                            Dim sqlCmd71 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                            Dim sqlDA71 As New SqlDataAdapter(sqlCmd71)
                                            Dim ds71 As New DataSet
                                            sqlDA71.Fill(ds71)

                                            If ds71.Tables(0).Rows.Count = 0 Then
                                                If pEmergencyCls = "M" Then
                                                    lblInfo.Text = "[9999] Please registered ETA and ETD for this Period in Time Chart Master!"
                                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                                    MyConnection.Close()
                                                    Exit Sub
                                                End If
                                            Else
                                                '99.3 Check Price di MS_Price
                                                ls_sql = "SELECT Price FROM MS_Price WHERE ('" & Format(ls_Tanggal, "yyyy-MM-dd") & "' between StartDate and EndDate) and AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                                         "                  and PartNo = '" & PO.D_PartNo & "' "
                                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                                Dim ds2 As New DataSet
                                                sqlDA2.Fill(ds2)

                                                If ds2.Tables(0).Rows.Count = 0 Then
                                                    If ls_error = "" Then
                                                        ls_error = "Price " & PO.D_PartNo.Trim & " must be registered"
                                                    Else
                                                        ls_error = ls_error + ", Price " & PO.D_PartNo.Trim & " must be registered!"
                                                    End If
                                                Else
                                                    '99.4 Check data Qtybox
                                                    If PO.D_Week2 > 0 Then
                                                        If PO.D_Week2 >= ls_MOQ Then
                                                            If (PO.D_Week2 Mod ls_QtyBox) <> 0 Then
                                                                If ls_error = "" Then
                                                                    ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                Else
                                                                    ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                End If
                                                            End If
                                                        Else
                                                            If ls_error = "" Then
                                                                ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week2 is less than MOQ!, please check the file again!"
                                                            Else
                                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week2 is less than MOQ!, please check the file again!"
                                                            End If
                                                        End If
                                                    End If

                                                    'If PO.D_Forecast1 > 0 Then
                                                    '    If PO.D_Forecast1 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast1 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast2 > 0 Then
                                                    '    If PO.D_Forecast2 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast2 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast3 > 0 Then
                                                    '    If PO.D_Forecast3 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast3 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Check PONO di PO MASTER
                                    ls_sql = "SELECT * FROM PO_Master_Export WHERE PONo = '" & dtUploadHeader.H_OrderNo2 & "' " & vbCrLf & _
                                        "and AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                        "and SupplierID = '" & ls_SupplierID & "' " & vbCrLf & _
                                        "--and PASISendToSupplierDate is not null "

                                    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                    Dim ds5 As New DataSet
                                    sqlDA5.Fill(ds5)

                                    If ds5.Tables(0).Rows.Count > 0 Then
                                        If Not IsDBNull(ds5.Tables(0).Rows(0)("PASISendToSupplierDate")) Then
                                            If ls_error = "" Then
                                                ls_error = "PO No. already send to supplier!"
                                            Else
                                                ls_error = ls_error + ", PO No. already send to supplier!"
                                            End If
                                        End If
                                    End If

                                    If InStr(1, Session("PONoUpload"), dtUploadHeader.H_OrderNo2.Trim) = 0 Then
                                        If Session("PONoUpload") = "" Then
                                            Session("PONoUpload") = "'" & dtUploadHeader.H_OrderNo2.Trim & "'"
                                        Else
                                            Session("PONoUpload") = Session("PONoUpload") & ",'" & dtUploadHeader.H_OrderNo2.Trim & "'"
                                        End If
                                    End If

                                    ls_sql = " INSERT INTO [dbo].[UploadPOExport] " & vbCrLf & _
                                             "            ([PONo],[AffiliateID],[SupplierID],[ForwarderID],[Period],[CommercialCls],[EmergencyCls],[ShipCls] " & vbCrLf & _
                                             "            ,[PartNo],[OrderNo1],[ETDVendor1],[ETDPort1],[ETAPort1],[ETAFactory1],[Week1],[PreviousForecast],[ForecastN1] " & vbCrLf & _
                                             "            ,[ForecastN2],[ForecastN3],[ErrorCls],[UOM]) " & vbCrLf & _
                                             "      VALUES " & vbCrLf & _
                                             "            ('" & dtUploadHeader.H_OrderNo2 & "' " & vbCrLf & _
                                             "            ,'" & pAffiliateID & "' " & vbCrLf & _
                                             "            ,'" & ls_SupplierID & "' " & vbCrLf

                                    ls_sql = ls_sql + "            ,'" & pForwarderID & "' " & vbCrLf & _
                                                      "            ,'" & pPeriod & "' " & vbCrLf & _
                                                      "            ,'" & pCommercial & "' " & vbCrLf & _
                                                      "            ,'" & pEmergencyCls & "' " & vbCrLf & _
                                                      "            ,'" & pShipBy & "' " & vbCrLf & _
                                                      "            ,'" & PO.D_PartNo & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_OrderNo2 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDVendor2 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDPort2 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAPort2 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAFactory2 & "' " & vbCrLf & _
                                                      "            ," & PO.D_Week2 & " " & vbCrLf & _
                                                      "            ,0 " & vbCrLf

                                    ls_sql = ls_sql + "            ," & PO.D_Forecast1 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast2 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast3 & "" & vbCrLf & _
                                                      "            ,'" & ls_error & "'" & vbCrLf & _
                                                      "            ,'" & PO.D_UOM & "') " & vbCrLf

                                    Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                    sqlComm.Dispose()
                                End If

                                If dtUploadHeader.H_OrderNo3 <> "" Then
                                    ls_MOQ = 0
                                    ls_QtyBox = 0
                                    ls_SupplierID = ""
                                    ls_error = ""

                                    If Format(dtUploadHeader.H_ETDPort3, "yyyy-MM-dd") <> "0001-01-01" Then
                                        ls_Tanggal = dtUploadHeader.H_ETDPort3
                                    Else
                                        ls_Tanggal = Format(Now, "yyyy-MM-dd")
                                    End If

                                    '99.0 Check PartNo di Ms_PartMapping
                                    ls_sql = "SELECT a.*, b.Description UOM FROM dbo.MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls WHERE a.PartNo = '" & PO.D_PartNo & "'"
                                    Dim sqlCmd20 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA20 As New SqlDataAdapter(sqlCmd20)
                                    Dim ds20 As New DataSet
                                    sqlDA20.Fill(ds20)

                                    If ds20.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        End If
                                    Else
                                        '99.0.1 Check UOM
                                        ls_UOM = ds20.Tables(0).Rows(0)("UOM")
                                        If ls_UOM.Trim.ToUpper <> PO.D_UOM.Trim.ToUpper Then
                                            If ls_error = "" Then
                                                ls_error = PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            Else
                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            End If
                                        End If

                                        '99.1 Check PartNo di Ms_PartMapping
                                        ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                                        Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                        Dim ds3 As New DataSet
                                        sqlDA3.Fill(ds3)

                                        If ds3.Tables(0).Rows.Count = 0 Then
                                            If ls_error = "" Then
                                                ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            Else
                                                ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            End If
                                        Else
                                            ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                            ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")
                                            ls_QtyBox = IIf(IsDBNull(ds3.Tables(0).Rows(0)("QtyBox")), 0, ds3.Tables(0).Rows(0)("QtyBox"))

                                            '99.2 Check ETA ETD                                    
                                            ls_sql = "select * from MS_ETD_Export where AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                                     "and Period = '" & pPeriod & "' and " & vbCrLf & _
                                                     "SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
                                                     "and ETAForwarder = '" & dtUploadHeader.H_ETDVendor3 & "'" & vbCrLf & _
                                                     "and ETDPort = '" & dtUploadHeader.H_ETDPort3 & "'" & vbCrLf & _
                                                     "and ETAPort = '" & dtUploadHeader.H_ETAPort3 & "'" & vbCrLf & _
                                                     "and ETAFactory = '" & dtUploadHeader.H_ETAFactory3 & "'"
                                            Dim sqlCmd71 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                            Dim sqlDA71 As New SqlDataAdapter(sqlCmd71)
                                            Dim ds71 As New DataSet
                                            sqlDA71.Fill(ds71)

                                            If ds71.Tables(0).Rows.Count = 0 Then
                                                If pEmergencyCls = "M" Then
                                                    lblInfo.Text = "[9999] Please registered ETA and ETD for this Period in Time Chart Master!"
                                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                                    MyConnection.Close()
                                                    Exit Sub
                                                End If
                                            Else
                                                '99.3 Check Price di MS_Price
                                                ls_sql = "SELECT Price FROM MS_Price WHERE ('" & Format(ls_Tanggal, "yyyy-MM-dd") & "' between StartDate and EndDate) and AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                                         "                  and PartNo = '" & PO.D_PartNo & "' "
                                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                                Dim ds2 As New DataSet
                                                sqlDA2.Fill(ds2)

                                                If ds2.Tables(0).Rows.Count = 0 Then
                                                    If ls_error = "" Then
                                                        ls_error = "Price " & PO.D_PartNo.Trim & " must be registered"
                                                    Else
                                                        ls_error = ls_error + ", Price " & PO.D_PartNo.Trim & " must be registered!"
                                                    End If
                                                Else
                                                    '99.4 Check data Qtybox
                                                    If PO.D_Week3 > 0 Then
                                                        If PO.D_Week3 >= ls_MOQ Then
                                                            If (PO.D_Week3 Mod ls_QtyBox) <> 0 Then
                                                                If ls_error = "" Then
                                                                    ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                Else
                                                                    ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                End If
                                                            End If
                                                        Else
                                                            If ls_error = "" Then
                                                                ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week3 is less than MOQ!, please check the file again!"
                                                            Else
                                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week3 is less than MOQ!, please check the file again!"
                                                            End If
                                                        End If
                                                    End If

                                                    'If PO.D_Forecast1 > 0 Then
                                                    '    If PO.D_Forecast1 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast1 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast2 > 0 Then
                                                    '    If PO.D_Forecast2 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast2 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast3 > 0 Then
                                                    '    If PO.D_Forecast3 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast3 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Check PONO di PO MASTER
                                    ls_sql = "SELECT * FROM PO_Master_Export WHERE PONo = '" & dtUploadHeader.H_OrderNo3 & "' " & vbCrLf & _
                                        "and AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                        "and SupplierID = '" & ls_SupplierID & "' " & vbCrLf & _
                                        "--and PASISendToSupplierDate is not null "

                                    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                    Dim ds5 As New DataSet
                                    sqlDA5.Fill(ds5)

                                    If ds5.Tables(0).Rows.Count > 0 Then
                                        If Not IsDBNull(ds5.Tables(0).Rows(0)("PASISendToSupplierDate")) Then
                                            If ls_error = "" Then
                                                ls_error = "PO No. already send to supplier!"
                                            Else
                                                ls_error = ls_error + ", PO No. already send to supplier!"
                                            End If
                                        End If
                                    End If

                                    If InStr(1, Session("PONoUpload"), dtUploadHeader.H_OrderNo3.Trim) = 0 Then
                                        If Session("PONoUpload") = "" Then
                                            Session("PONoUpload") = "'" & dtUploadHeader.H_OrderNo3.Trim & "'"
                                        Else
                                            Session("PONoUpload") = Session("PONoUpload") & ",'" & dtUploadHeader.H_OrderNo3.Trim & "'"
                                        End If
                                    End If

                                    ls_sql = " INSERT INTO [dbo].[UploadPOExport] " & vbCrLf & _
                                             "            ([PONo],[AffiliateID],[SupplierID],[ForwarderID],[Period],[CommercialCls],[EmergencyCls],[ShipCls] " & vbCrLf & _
                                             "            ,[PartNo],[OrderNo1],[ETDVendor1],[ETDPort1],[ETAPort1],[ETAFactory1],[Week1],[PreviousForecast],[ForecastN1] " & vbCrLf & _
                                             "            ,[ForecastN2],[ForecastN3],[ErrorCls],[UOM]) " & vbCrLf & _
                                             "      VALUES " & vbCrLf & _
                                             "            ('" & dtUploadHeader.H_OrderNo3 & "' " & vbCrLf & _
                                             "            ,'" & pAffiliateID & "' " & vbCrLf & _
                                             "            ,'" & ls_SupplierID & "' " & vbCrLf

                                    ls_sql = ls_sql + "            ,'" & pForwarderID & "' " & vbCrLf & _
                                                      "            ,'" & pPeriod & "' " & vbCrLf & _
                                                      "            ,'" & pCommercial & "' " & vbCrLf & _
                                                      "            ,'" & pEmergencyCls & "' " & vbCrLf & _
                                                      "            ,'" & pShipBy & "' " & vbCrLf & _
                                                      "            ,'" & PO.D_PartNo & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_OrderNo3 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDVendor3 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDPort3 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAPort3 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAFactory3 & "' " & vbCrLf & _
                                                      "            ," & PO.D_Week3 & " " & vbCrLf & _
                                                      "            ,0 " & vbCrLf

                                    ls_sql = ls_sql + "            ," & PO.D_Forecast1 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast2 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast3 & "" & vbCrLf & _
                                                      "            ,'" & ls_error & "'" & vbCrLf & _
                                                      "            ,'" & PO.D_UOM & "') " & vbCrLf

                                    Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                    sqlComm.Dispose()
                                End If

                                If dtUploadHeader.H_OrderNo4 <> "" Then
                                    ls_MOQ = 0
                                    ls_QtyBox = 0
                                    ls_SupplierID = ""
                                    ls_error = ""

                                    If Format(dtUploadHeader.H_ETDPort4, "yyyy-MM-dd") <> "0001-01-01" Then
                                        ls_Tanggal = dtUploadHeader.H_ETDPort4
                                    Else
                                        ls_Tanggal = Format(Now, "yyyy-MM-dd")
                                    End If

                                    '99.0 Check PartNo di Ms_PartMapping
                                    ls_sql = "SELECT a.*, b.Description UOM FROM dbo.MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls WHERE a.PartNo = '" & PO.D_PartNo & "'"
                                    Dim sqlCmd20 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA20 As New SqlDataAdapter(sqlCmd20)
                                    Dim ds20 As New DataSet
                                    sqlDA20.Fill(ds20)

                                    If ds20.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        End If
                                    Else
                                        '99.0.1 Check UOM
                                        ls_UOM = ds20.Tables(0).Rows(0)("UOM")
                                        If ls_UOM.Trim.ToUpper <> PO.D_UOM.Trim.ToUpper Then
                                            If ls_error = "" Then
                                                ls_error = PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            Else
                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            End If
                                        End If

                                        '99.1 Check PartNo di Ms_PartMapping
                                        ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                                        Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                        Dim ds3 As New DataSet
                                        sqlDA3.Fill(ds3)

                                        If ds3.Tables(0).Rows.Count = 0 Then
                                            If ls_error = "" Then
                                                ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            Else
                                                ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            End If
                                        Else
                                            ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                            ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")
                                            ls_QtyBox = IIf(IsDBNull(ds3.Tables(0).Rows(0)("QtyBox")), 0, ds3.Tables(0).Rows(0)("QtyBox"))

                                            '99.2 Check ETA ETD                                    
                                            ls_sql = "select * from MS_ETD_Export where AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                                     "and Period = '" & pPeriod & "' and " & vbCrLf & _
                                                     "SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
                                                     "and ETAForwarder = '" & dtUploadHeader.H_ETDVendor4 & "'" & vbCrLf & _
                                                     "and ETDPort = '" & dtUploadHeader.H_ETDPort4 & "'" & vbCrLf & _
                                                     "and ETAPort = '" & dtUploadHeader.H_ETAPort4 & "'" & vbCrLf & _
                                                     "and ETAFactory = '" & dtUploadHeader.H_ETAFactory4 & "'"
                                            Dim sqlCmd71 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                            Dim sqlDA71 As New SqlDataAdapter(sqlCmd71)
                                            Dim ds71 As New DataSet
                                            sqlDA71.Fill(ds71)

                                            If ds71.Tables(0).Rows.Count = 0 Then
                                                If pEmergencyCls = "M" Then
                                                    lblInfo.Text = "[9999] Please registered ETA and ETD for this Period in Time Chart Master!"
                                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                                    MyConnection.Close()
                                                    Exit Sub
                                                End If
                                            Else
                                                '99.3 Check Price di MS_Price
                                                ls_sql = "SELECT Price FROM MS_Price WHERE ('" & Format(ls_Tanggal, "yyyy-MM-dd") & "' between StartDate and EndDate) and AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                                         "                  and PartNo = '" & PO.D_PartNo & "' "
                                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                                Dim ds2 As New DataSet
                                                sqlDA2.Fill(ds2)

                                                If ds2.Tables(0).Rows.Count = 0 Then
                                                    If ls_error = "" Then
                                                        ls_error = "Price " & PO.D_PartNo.Trim & " must be registered"
                                                    Else
                                                        ls_error = ls_error + ", Price " & PO.D_PartNo.Trim & " must be registered!"
                                                    End If
                                                Else
                                                    '99.4 Check data Qtybox
                                                    If PO.D_Week4 > 0 Then
                                                        If PO.D_Week4 >= ls_MOQ Then
                                                            If (PO.D_Week4 Mod ls_QtyBox) <> 0 Then
                                                                If ls_error = "" Then
                                                                    ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week4 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                Else
                                                                    ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week4 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                End If
                                                            End If
                                                        Else
                                                            If ls_error = "" Then
                                                                ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week4 is less than MOQ!, please check the file again!"
                                                            Else
                                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week4 is less than MOQ!, please check the file again!"
                                                            End If
                                                        End If
                                                    End If

                                                    'If PO.D_Forecast1 > 0 Then
                                                    '    If PO.D_Forecast1 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast1 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast2 > 0 Then
                                                    '    If PO.D_Forecast2 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast2 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast3 > 0 Then
                                                    '    If PO.D_Forecast3 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast3 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Check PONO di PO MASTER
                                    ls_sql = "SELECT * FROM PO_Master_Export WHERE PONo = '" & dtUploadHeader.H_OrderNo4 & "' " & vbCrLf & _
                                        "and AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                        "and SupplierID = '" & ls_SupplierID & "' " & vbCrLf & _
                                        "--and PASISendToSupplierDate is not null "

                                    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                    Dim ds5 As New DataSet
                                    sqlDA5.Fill(ds5)

                                    If ds5.Tables(0).Rows.Count > 0 Then
                                        If Not IsDBNull(ds5.Tables(0).Rows(0)("PASISendToSupplierDate")) Then
                                            If ls_error = "" Then
                                                ls_error = "PO No. already send to supplier!"
                                            Else
                                                ls_error = ls_error + ", PO No. already send to supplier!"
                                            End If
                                        End If
                                    End If

                                    If InStr(1, Session("PONoUpload"), dtUploadHeader.H_OrderNo4.Trim) = 0 Then
                                        If Session("PONoUpload") = "" Then
                                            Session("PONoUpload") = "'" & dtUploadHeader.H_OrderNo4.Trim & "'"
                                        Else
                                            Session("PONoUpload") = Session("PONoUpload") & ",'" & dtUploadHeader.H_OrderNo4.Trim & "'"
                                        End If
                                    End If

                                    ls_sql = " INSERT INTO [dbo].[UploadPOExport] " & vbCrLf & _
                                             "            ([PONo],[AffiliateID],[SupplierID],[ForwarderID],[Period],[CommercialCls],[EmergencyCls],[ShipCls] " & vbCrLf & _
                                             "            ,[PartNo],[OrderNo1],[ETDVendor1],[ETDPort1],[ETAPort1],[ETAFactory1],[Week1],[PreviousForecast],[ForecastN1] " & vbCrLf & _
                                             "            ,[ForecastN2],[ForecastN3],[ErrorCls],[UOM]) " & vbCrLf & _
                                             "      VALUES " & vbCrLf & _
                                             "            ('" & dtUploadHeader.H_OrderNo4 & "' " & vbCrLf & _
                                             "            ,'" & pAffiliateID & "' " & vbCrLf & _
                                             "            ,'" & ls_SupplierID & "' " & vbCrLf

                                    ls_sql = ls_sql + "            ,'" & pForwarderID & "' " & vbCrLf & _
                                                      "            ,'" & pPeriod & "' " & vbCrLf & _
                                                      "            ,'" & pCommercial & "' " & vbCrLf & _
                                                      "            ,'" & pEmergencyCls & "' " & vbCrLf & _
                                                      "            ,'" & pShipBy & "' " & vbCrLf & _
                                                      "            ,'" & PO.D_PartNo & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_OrderNo4 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDVendor4 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDPort4 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAPort4 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAFactory4 & "' " & vbCrLf & _
                                                      "            ," & PO.D_Week4 & " " & vbCrLf & _
                                                      "            ,0 " & vbCrLf

                                    ls_sql = ls_sql + "            ," & PO.D_Forecast1 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast2 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast3 & "" & vbCrLf & _
                                                      "            ,'" & ls_error & "'" & vbCrLf & _
                                                      "            ,'" & PO.D_UOM & "') " & vbCrLf

                                    Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                    sqlComm.Dispose()
                                End If

                                If dtUploadHeader.H_OrderNo5 <> "" Then
                                    ls_MOQ = 0
                                    ls_QtyBox = 0
                                    ls_SupplierID = ""
                                    ls_error = ""

                                    If Format(dtUploadHeader.H_ETDPort5, "yyyy-MM-dd") <> "0001-01-01" Then
                                        ls_Tanggal = dtUploadHeader.H_ETDPort5
                                    Else
                                        ls_Tanggal = Format(Now, "yyyy-MM-dd")
                                    End If

                                    '99.0 Check PartNo di Ms_PartMapping
                                    ls_sql = "SELECT a.*, b.Description UOM FROM dbo.MS_Parts a left join MS_UnitCls b on a.UnitCls = b.UnitCls WHERE a.PartNo = '" & PO.D_PartNo & "'"
                                    Dim sqlCmd20 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA20 As New SqlDataAdapter(sqlCmd20)
                                    Dim ds20 As New DataSet
                                    sqlDA20.Fill(ds20)

                                    If ds20.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Master, please check again with PASI!"
                                        End If
                                    Else
                                        '99.0.1 Check UOM
                                        ls_UOM = ds20.Tables(0).Rows(0)("UOM")
                                        If ls_UOM.Trim.ToUpper <> PO.D_UOM.Trim.ToUpper Then
                                            If ls_error = "" Then
                                                ls_error = PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            Else
                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " UOM not found!, please check the file again!"
                                            End If
                                        End If

                                        '99.1 Check PartNo di Ms_PartMapping
                                        ls_sql = "SELECT * FROM dbo.MS_PartMapping WHERE PartNo = '" & PO.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                                        Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                        Dim ds3 As New DataSet
                                        sqlDA3.Fill(ds3)

                                        If ds3.Tables(0).Rows.Count = 0 Then
                                            If ls_error = "" Then
                                                ls_error = "PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            Else
                                                ls_error = ls_error + ", PartNo " & PO.D_PartNo.Trim & " not found in Part Mapping, please check again with PASI!"
                                            End If
                                        Else
                                            ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                            ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")
                                            ls_QtyBox = IIf(IsDBNull(ds3.Tables(0).Rows(0)("QtyBox")), 0, ds3.Tables(0).Rows(0)("QtyBox"))

                                            '99.2 Check ETA ETD                                    
                                            ls_sql = "select * from MS_ETD_Export where AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                                     "and Period = '" & pPeriod & "' and " & vbCrLf & _
                                                     "SupplierID = '" & ls_SupplierID & "'" & vbCrLf & _
                                                     "and ETAForwarder = '" & dtUploadHeader.H_ETDVendor5 & "'" & vbCrLf & _
                                                     "and ETDPort = '" & dtUploadHeader.H_ETDPort5 & "'" & vbCrLf & _
                                                     "and ETAPort = '" & dtUploadHeader.H_ETAPort5 & "'" & vbCrLf & _
                                                     "and ETAFactory = '" & dtUploadHeader.H_ETAFactory5 & "'"
                                            Dim sqlCmd71 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                            Dim sqlDA71 As New SqlDataAdapter(sqlCmd71)
                                            Dim ds71 As New DataSet
                                            sqlDA71.Fill(ds71)

                                            If ds71.Tables(0).Rows.Count = 0 Then
                                                If pEmergencyCls = "M" Then
                                                    lblInfo.Text = "[9999] Please registered ETA and ETD for this Period in Time Chart Master!"
                                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                                    MyConnection.Close()
                                                    Exit Sub
                                                End If
                                            Else
                                                '99.3 Check Price di MS_Price
                                                ls_sql = "SELECT Price FROM MS_Price WHERE ('" & Format(ls_Tanggal, "yyyy-MM-dd") & "' between StartDate and EndDate) and AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                                         "                  and PartNo = '" & PO.D_PartNo & "' "
                                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                                Dim ds2 As New DataSet
                                                sqlDA2.Fill(ds2)

                                                If ds2.Tables(0).Rows.Count = 0 Then
                                                    If ls_error = "" Then
                                                        ls_error = "Price " & PO.D_PartNo.Trim & " must be registered"
                                                    Else
                                                        ls_error = ls_error + ", Price " & PO.D_PartNo.Trim & " must be registered!"
                                                    End If
                                                Else
                                                    '99.4 Check data Qtybox
                                                    If PO.D_Week5 > 0 Then
                                                        If PO.D_Week5 >= ls_MOQ Then
                                                            If (PO.D_Week5 Mod ls_QtyBox) <> 0 Then
                                                                If ls_error = "" Then
                                                                    ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week5 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                Else
                                                                    ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week5 must be same or multiple of the Qty/Box!, please check the file again!"
                                                                End If
                                                            End If
                                                        Else
                                                            If ls_error = "" Then
                                                                ls_error = PO.D_PartNo.Trim & " Total Firm Qty Week5 is less than MOQ!, please check the file again!"
                                                            Else
                                                                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Total Firm Qty Week5 is less than MOQ!, please check the file again!"
                                                            End If
                                                        End If
                                                    End If

                                                    'If PO.D_Forecast1 > 0 Then
                                                    '    If PO.D_Forecast1 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast1 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N1 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast2 > 0 Then
                                                    '    If PO.D_Forecast2 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast2 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N2 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If

                                                    'If PO.D_Forecast3 > 0 Then
                                                    '    If PO.D_Forecast3 >= ls_MOQ Then
                                                    '        If (PO.D_Forecast3 Mod ls_QtyBox) <> 0 Then
                                                    '            If ls_error = "" Then
                                                    '                ls_error = PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            Else
                                                    '                ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 must be same or multiple of the Qty/Box!, please check the file again!"
                                                    '            End If
                                                    '        End If
                                                    '    Else
                                                    '        If ls_error = "" Then
                                                    '            ls_error = PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        Else
                                                    '            ls_error = ls_error + ", " & PO.D_PartNo.Trim & " Forecast N3 is less than MOQ!, please check the file again!"
                                                    '        End If
                                                    '    End If
                                                    'End If
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Check PONO di PO MASTER
                                    ls_sql = "SELECT * FROM PO_Master_Export WHERE PONo = '" & dtUploadHeader.H_OrderNo5 & "' " & vbCrLf & _
                                        "and AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                        "and SupplierID = '" & ls_SupplierID & "' " & vbCrLf & _
                                        "--and PASISendToSupplierDate is not null "

                                    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                                    Dim ds5 As New DataSet
                                    sqlDA5.Fill(ds5)

                                    If ds5.Tables(0).Rows.Count > 0 Then
                                        If Not IsDBNull(ds5.Tables(0).Rows(0)("PASISendToSupplierDate")) Then
                                            If ls_error = "" Then
                                                ls_error = "PO No. already send to supplier!"
                                            Else
                                                ls_error = ls_error + ", PO No. already send to supplier!"
                                            End If
                                        End If
                                    End If

                                    If InStr(1, Session("PONoUpload"), dtUploadHeader.H_OrderNo5.Trim) = 0 Then
                                        If Session("PONoUpload") = "" Then
                                            Session("PONoUpload") = "'" & dtUploadHeader.H_OrderNo5.Trim & "'"
                                        Else
                                            Session("PONoUpload") = Session("PONoUpload") & ",'" & dtUploadHeader.H_OrderNo5.Trim & "'"
                                        End If
                                    End If

                                    ls_sql = " INSERT INTO [dbo].[UploadPOExport] " & vbCrLf & _
                                             "            ([PONo],[AffiliateID],[SupplierID],[ForwarderID],[Period],[CommercialCls],[EmergencyCls],[ShipCls] " & vbCrLf & _
                                             "            ,[PartNo],[OrderNo1],[ETDVendor1],[ETDPort1],[ETAPort1],[ETAFactory1],[Week1],[PreviousForecast],[ForecastN1] " & vbCrLf & _
                                             "            ,[ForecastN2],[ForecastN3],[ErrorCls],[UOM]) " & vbCrLf & _
                                             "      VALUES " & vbCrLf & _
                                             "            ('" & dtUploadHeader.H_OrderNo5 & "' " & vbCrLf & _
                                             "            ,'" & pAffiliateID & "' " & vbCrLf & _
                                             "            ,'" & ls_SupplierID & "' " & vbCrLf

                                    ls_sql = ls_sql + "            ,'" & pForwarderID & "' " & vbCrLf & _
                                                      "            ,'" & pPeriod & "' " & vbCrLf & _
                                                      "            ,'" & pCommercial & "' " & vbCrLf & _
                                                      "            ,'" & pEmergencyCls & "' " & vbCrLf & _
                                                      "            ,'" & pShipBy & "' " & vbCrLf & _
                                                      "            ,'" & PO.D_PartNo & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_OrderNo5 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDVendor5 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETDPort5 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAPort5 & "' " & vbCrLf & _
                                                      "            ,'" & dtUploadHeader.H_ETAFactory5 & "' " & vbCrLf & _
                                                      "            ," & PO.D_Week5 & " " & vbCrLf & _
                                                      "            ,0 " & vbCrLf

                                    ls_sql = ls_sql + "            ," & PO.D_Forecast1 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast2 & " " & vbCrLf & _
                                                      "            ," & PO.D_Forecast3 & "" & vbCrLf & _
                                                      "            ,'" & ls_error & "'" & vbCrLf & _
                                                      "            ,'" & PO.D_UOM & "') " & vbCrLf

                                    Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                    sqlComm.Dispose()
                                End If
                            Next
                            sqlTran.Commit()

                            lblInfo.Text = "[7001] Data Checking Done!"
                            lblInfo.ForeColor = Color.Blue
                            grid.JSProperties("cpMessage") = lblInfo.Text

                            Call bindData()

                            'If pAda = True Then
                            '    'popUp2.ShowOnPageLoad = True
                            '    Call up_Save()
                            'Else
                            '    Call up_Save()
                            'End If

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

    Private Sub up_SaveOld()
        Dim k As Integer
        Dim ls_Check As Boolean = False
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim ls_OrderNoErr As String = ""
        Dim ls_Forecast1 As Integer = 0
        Dim ls_Forecast2 As Integer = 0
        Dim ls_forecast3 As Integer = 0
        Dim ls_PrevForecast As Integer = 0
        Dim ls_Variance As Integer = 0
        Dim ls_VariancePercentage As Double = 0
        Dim ls_Week1 As Integer = 0

        Try
            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()
            SqlTran = SqlCon.BeginTransaction

            Dim SQLCom As SqlCommand = SqlCon.CreateCommand
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran

            'delete data first
            ls_Sql = "delete PO_Detail_Export where PONo in (" & Session("PONoUpload") & ") and SupplierID in (" & Session("tempSupplierID") & ")" & vbCrLf & _
                "and exists( " & vbCrLf & _
                "    select * from PO_Master_Export a where a.PONo in (" & Session("PONoUpload") & ") " & vbCrLf & _
                "    and a.PASISendToSupplierDate is null " & vbCrLf & _
                "    and a.PONo = PO_Detail_Export.PONo and a.AffiliateID = PO_Detail_Export.AffiliateID and a.SupplierID = PO_Detail_Export.SupplierID " & vbCrLf & _
                ") "
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            ls_Sql = "delete PO_Master_Export where PONo in (" & Session("PONoUpload") & ") and SupplierID in (" & Session("tempSupplierID") & ")" & vbCrLf & _
                "    and PASISendToSupplierDate is null "
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            Try
                ls_Sql = "select * from UploadPOExport where PONo in (" & Session("PONoUpload") & ") order by PONo, AffiliateID, SupplierID "

                SQLCom.CommandText = ls_Sql
                Dim da11 As New SqlDataAdapter(SQLCom)
                Dim ds11 As New DataSet
                da11.Fill(ds11)

                If ds11.Tables(0).Rows.Count > 0 Then
                    For k = 0 To ds11.Tables(0).Rows.Count - 1
                        If ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim = "M" Then
                            ls_Week1 = ds11.Tables(0).Rows(k)("Week1").ToString.Trim
                            ls_PrevForecast = 0 'ds11.Tables(0).Rows(k)("PreviousForecast").ToString.Trim
                            ls_Forecast1 = ds11.Tables(0).Rows(k)("ForecastN1").ToString.Trim
                            ls_Forecast2 = ds11.Tables(0).Rows(k)("ForecastN2").ToString.Trim
                            ls_forecast3 = ds11.Tables(0).Rows(k)("ForecastN3").ToString.Trim
                            ls_Variance = 0 'ls_PrevForecast - ls_Week1
                            ls_VariancePercentage = 0 'IIf(ls_Week1 = 0, 0, (ls_Variance / ls_Week1) * 100)
                        Else
                            ls_Week1 = ds11.Tables(0).Rows(k)("Week1").ToString.Trim
                            ls_PrevForecast = 0
                            ls_Forecast1 = 0
                            ls_Forecast2 = 0
                            ls_forecast3 = 0
                            ls_Variance = 0
                            ls_VariancePercentage = 0
                        End If

                        If ds11.Tables(0).Rows(k)("ErrorCls").ToString.Trim = "" Then
                            If ls_Week1 > 0 Then
                                ls_Sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[ForwarderID] " & vbCrLf & _
                                      "            ,[OrderNo1] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[Week1] " & vbCrLf & _
                                      "            ,[TotalPOQty] " & vbCrLf

                                ls_Sql = ls_Sql + "            ,[PreviousForecast] " & vbCrLf & _
                                                  "            ,[Forecast1] " & vbCrLf & _
                                                  "            ,[Forecast2] " & vbCrLf & _
                                                  "            ,[Forecast3] " & vbCrLf & _
                                                  "            ,[Variance] " & vbCrLf & _
                                                  "            ,[VariancePercentage] " & vbCrLf & _
                                                  "            ,[EntryDate] " & vbCrLf & _
                                                  "            ,[EntryUser]) " & vbCrLf

                                ls_Sql = ls_Sql + "      VALUES " & vbCrLf & _
                                                  "            ('" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim.ToString & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("OrderNo1").ToString.Trim & "' "

                                ls_Sql = ls_Sql + "            ,'" & ds11.Tables(0).Rows(k)("PartNo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_PrevForecast & "' "

                                ls_Sql = ls_Sql + "            ,'" & ls_Forecast1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Forecast2 & "' " & vbCrLf & _
                                                  "            ,'" & ls_forecast3 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Variance & "' " & vbCrLf & _
                                                  "            ,'" & ls_VariancePercentage & "' " & vbCrLf & _
                                                  "            , getdate() " & vbCrLf & _
                                                  "            ,'" & Session("UserID") & "' ) "

                                SQLCom.CommandText = ls_Sql
                                SQLCom.ExecuteNonQuery()
                                ls_MsgID = "1001"
                                ls_Detail = "ada"

                                ls_Sql = "select * from PO_Master_Export where PONo = '" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' and AffiliateID = '" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' and SupplierID = '" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' and OrderNo1 ='" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "'"

                                SQLCom.CommandText = ls_Sql
                                Dim da7 As New SqlDataAdapter(SQLCom)
                                Dim ds7 As New DataSet
                                da7.Fill(ds7)

                                If ds7.Tables(0).Rows.Count = 0 Then
                                    ls_Sql = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                              "            ([PONo] " & vbCrLf & _
                                              "            ,[AffiliateID] " & vbCrLf & _
                                              "            ,[SupplierID] " & vbCrLf & _
                                              "            ,[ForwarderID] " & vbCrLf & _
                                              "            ,[Period] " & vbCrLf & _
                                              "            ,[CommercialCls] " & vbCrLf & _
                                              "            ,[EmergencyCls] " & vbCrLf & _
                                              "            ,[ShipCls] " & vbCrLf & _
                                              "            ,[ErrorStatus] " & vbCrLf & _
                                              "            ,[OrderNo1] " & vbCrLf & _
                                              "            ,[ETDVendor1] " & vbCrLf & _
                                              "            ,[ETDPort1] " & vbCrLf & _
                                              "            ,[ETAPort1] " & vbCrLf & _
                                              "            ,[ETAFactory1] " & vbCrLf & _
                                              "            ,[UploadDate] " & vbCrLf & _
                                              "            ,[UploadUser] " & vbCrLf & _
                                              "            ,[EntryDate] " & vbCrLf & _
                                              "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("Period").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("CommercialCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ShipCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'OK' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETDVendor1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETDPort1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETAPort1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETAFactory1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "') "

                                    SQLCom.CommandText = ls_Sql
                                    SQLCom.ExecuteNonQuery()
                                End If

                            End If
                        Else
                            'Split Error
                            ls_OrderNoErr = ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "-1"

                            If ls_Week1 > 0 Then
                                ls_Sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[ForwarderID] " & vbCrLf & _
                                      "            ,[OrderNo1] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[Week1] " & vbCrLf & _
                                      "            ,[TotalPOQty] " & vbCrLf

                                ls_Sql = ls_Sql + "            ,[PreviousForecast] " & vbCrLf & _
                                                  "            ,[Forecast1] " & vbCrLf & _
                                                  "            ,[Forecast2] " & vbCrLf & _
                                                  "            ,[Forecast3] " & vbCrLf & _
                                                  "            ,[Variance] " & vbCrLf & _
                                                  "            ,[VariancePercentage] " & vbCrLf & _
                                                  "            ,[EntryDate] " & vbCrLf & _
                                                  "            ,[EntryUser]) " & vbCrLf

                                ls_Sql = ls_Sql + "      VALUES " & vbCrLf & _
                                                  "            ('" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim.ToString & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ls_OrderNoErr & "' "

                                ls_Sql = ls_Sql + "            ,'" & ds11.Tables(0).Rows(k)("PartNo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_PrevForecast & "' "

                                ls_Sql = ls_Sql + "            ,'" & ls_Forecast1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Forecast2 & "' " & vbCrLf & _
                                                  "            ,'" & ls_forecast3 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Variance & "' " & vbCrLf & _
                                                  "            ,'" & ls_VariancePercentage & "' " & vbCrLf & _
                                                  "            , getdate() " & vbCrLf & _
                                                  "            ,'" & Session("UserID") & "' ) "

                                SQLCom.CommandText = ls_Sql
                                SQLCom.ExecuteNonQuery()
                                ls_MsgID = "1001"
                                ls_Detail = "ada"

                                ls_Sql = "select * from PO_Master_Export where PONo = '" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' and AffiliateID = '" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' and SupplierID = '" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' and OrderNo1 ='" & ls_OrderNoErr & "'"

                                SQLCom.CommandText = ls_Sql
                                Dim da7 As New SqlDataAdapter(SQLCom)
                                Dim ds7 As New DataSet
                                da7.Fill(ds7)

                                If ds7.Tables(0).Rows.Count = 0 Then
                                    ls_Sql = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                              "            ([PONo] " & vbCrLf & _
                                              "            ,[AffiliateID] " & vbCrLf & _
                                              "            ,[SupplierID] " & vbCrLf & _
                                              "            ,[ForwarderID] " & vbCrLf & _
                                              "            ,[Period] " & vbCrLf & _
                                              "            ,[CommercialCls] " & vbCrLf & _
                                              "            ,[EmergencyCls] " & vbCrLf & _
                                              "            ,[ShipCls] " & vbCrLf & _
                                              "            ,[ErrorStatus] " & vbCrLf & _
                                              "            ,[OrderNo1] " & vbCrLf & _
                                              "            ,[ETDVendor1] " & vbCrLf & _
                                              "            ,[ETDPort1] " & vbCrLf & _
                                              "            ,[ETAPort1] " & vbCrLf & _
                                              "            ,[ETAFactory1] " & vbCrLf & _
                                              "            ,[UploadDate] " & vbCrLf & _
                                              "            ,[UploadUser] " & vbCrLf & _
                                              "            ,[EntryDate] " & vbCrLf & _
                                              "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("Period").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("CommercialCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ShipCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'NG' " & vbCrLf & _
                                              "            ,'" & ls_OrderNoErr & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETDVendor1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETDPort1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETAPort1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETAFactory1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "') "

                                    SQLCom.CommandText = ls_Sql
                                    SQLCom.ExecuteNonQuery()
                                End If
                            End If
                        End If
                    Next
                End If

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

    Private Sub up_Save2()
        Dim k As Integer
        Dim ls_Check As Boolean = False
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim ls_OrderNoErr As String = ""
        Dim ls_Forecast1 As Integer = 0
        Dim ls_Forecast2 As Integer = 0
        Dim ls_forecast3 As Integer = 0
        Dim ls_PrevForecast As Integer = 0
        Dim ls_Variance As Integer = 0
        Dim ls_VariancePercentage As Double = 0
        Dim ls_Week1 As Integer = 0

        Try
            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()
            SqlTran = SqlCon.BeginTransaction

            Dim SQLCom As SqlCommand = SqlCon.CreateCommand
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran

            ''delete data first
            'ls_Sql = "delete PO_Detail_Export where PONo in (" & Session("PONoUpload") & ") " & vbCrLf & _
            '    "and exists( " & vbCrLf & _
            '    "    select * from PO_Master_Export a where a.PONo in (" & Session("PONoUpload") & ") " & vbCrLf & _
            '    "    and a.PASISendToSupplierDate is null " & vbCrLf & _
            '    "    and a.PONo = PO_Detail_Export.PONo and a.AffiliateID = PO_Detail_Export.AffiliateID and a.SupplierID = PO_Detail_Export.SupplierID " & vbCrLf & _
            '    ") "
            'SQLCom.CommandText = ls_Sql
            'SQLCom.ExecuteNonQuery()

            'ls_Sql = "delete PO_Master_Export where PONo in (" & Session("PONoUpload") & ") " & vbCrLf & _
            '    "    and PASISendToSupplierDate is null "
            'SQLCom.CommandText = ls_Sql
            'SQLCom.ExecuteNonQuery()

            Try
                ls_Sql = "select * from UploadPOExport where PONo in (" & Session("PONoUpload") & ") order by PONo, AffiliateID, SupplierID "

                SQLCom.CommandText = ls_Sql
                Dim da11 As New SqlDataAdapter(SQLCom)
                Dim ds11 As New DataSet
                da11.Fill(ds11)

                If ds11.Tables(0).Rows.Count > 0 Then
                    For k = 0 To ds11.Tables(0).Rows.Count - 1
                        If ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim = "M" Then
                            ls_Week1 = ds11.Tables(0).Rows(k)("Week1").ToString.Trim
                            ls_PrevForecast = 0 'ds11.Tables(0).Rows(k)("PreviousForecast").ToString.Trim
                            ls_Forecast1 = ds11.Tables(0).Rows(k)("ForecastN1").ToString.Trim
                            ls_Forecast2 = ds11.Tables(0).Rows(k)("ForecastN2").ToString.Trim
                            ls_forecast3 = ds11.Tables(0).Rows(k)("ForecastN3").ToString.Trim
                            ls_Variance = 0 'ls_PrevForecast - ls_Week1
                            ls_VariancePercentage = 0 'IIf(ls_Week1 = 0, 0, (ls_Variance / ls_Week1) * 100)
                        Else
                            ls_Week1 = ds11.Tables(0).Rows(k)("Week1").ToString.Trim
                            ls_PrevForecast = 0
                            ls_Forecast1 = 0
                            ls_Forecast2 = 0
                            ls_forecast3 = 0
                            ls_Variance = 0
                            ls_VariancePercentage = 0
                        End If

                        If ds11.Tables(0).Rows(k)("ErrorCls").ToString.Trim = "" Then
                            Try
                                ls_Sql = "select TOP 1 AcceptPONo, PONo from PO_Master_Export where AcceptPONo = '" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' and AffiliateID = '" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' and SupplierID = '" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' order by PONo desc"
                                SQLCom.CommandText = ls_Sql
                                Dim da12 As New SqlDataAdapter(SQLCom)
                                Dim ds12 As New DataSet
                                da12.Fill(ds12)

                                If ds12.Tables(0).Rows.Count > 0 Then
                                    'ls_OrderNoErr = Split(ds12.Tables(0).Rows(k)("AcceptPONo").ToString.Trim, "-")(-1)
                                    ls_OrderNoErr = ds12.Tables(0).Rows(k)("PONo").ToString.Trim
                                    ls_OrderNoErr = Split(ls_OrderNoErr, "-")(1)
                                    If ls_OrderNoErr = "" Then
                                        ls_OrderNoErr = ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "-1"
                                    Else
                                        ls_OrderNoErr = ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "-" & CDbl(ls_OrderNoErr) + 1
                                    End If
                                Else
                                    ls_OrderNoErr = ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "-1"
                                End If
                            Catch ex As Exception
                                ls_OrderNoErr = ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "-1"
                            End Try

                            If ls_Week1 > 0 Then
                                ls_Sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                      "            ([PONo] " & vbCrLf & _
                                      "            ,[AffiliateID] " & vbCrLf & _
                                      "            ,[SupplierID] " & vbCrLf & _
                                      "            ,[ForwarderID] " & vbCrLf & _
                                      "            ,[OrderNo1] " & vbCrLf & _
                                      "            ,[PartNo] " & vbCrLf & _
                                      "            ,[Week1] " & vbCrLf & _
                                      "            ,[TotalPOQty] " & vbCrLf

                                ls_Sql = ls_Sql + "            ,[PreviousForecast] " & vbCrLf & _
                                                  "            ,[Forecast1] " & vbCrLf & _
                                                  "            ,[Forecast2] " & vbCrLf & _
                                                  "            ,[Forecast3] " & vbCrLf & _
                                                  "            ,[Variance] " & vbCrLf & _
                                                  "            ,[VariancePercentage] " & vbCrLf & _
                                                  "            ,[EntryDate] " & vbCrLf & _
                                                  "            ,[EntryUser]) " & vbCrLf

                                ls_Sql = ls_Sql + "      VALUES " & vbCrLf & _
                                                  "            ('" & ls_OrderNoErr & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim.ToString & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ls_OrderNoErr & "' "

                                ls_Sql = ls_Sql + "            ,'" & ds11.Tables(0).Rows(k)("PartNo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_PrevForecast & "' "

                                ls_Sql = ls_Sql + "            ,'" & ls_Forecast1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Forecast2 & "' " & vbCrLf & _
                                                  "            ,'" & ls_forecast3 & "' " & vbCrLf & _
                                                  "            ,'" & ls_Variance & "' " & vbCrLf & _
                                                  "            ,'" & ls_VariancePercentage & "' " & vbCrLf & _
                                                  "            , getdate() " & vbCrLf & _
                                                  "            ,'" & Session("UserID") & "' ) "

                                SQLCom.CommandText = ls_Sql
                                SQLCom.ExecuteNonQuery()
                                ls_MsgID = "1001"
                                ls_Detail = "ada"

                                ls_Sql = "select * from PO_Master_Export where PONo = '" & ls_OrderNoErr & "' and AffiliateID = '" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' and SupplierID = '" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' and OrderNo1 ='" & ls_OrderNoErr & "'"

                                SQLCom.CommandText = ls_Sql
                                Dim da7 As New SqlDataAdapter(SQLCom)
                                Dim ds7 As New DataSet
                                da7.Fill(ds7)

                                If ds7.Tables(0).Rows.Count = 0 Then
                                    ls_Sql = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                              "            ([PONo] " & vbCrLf & _
                                              "            ,[AffiliateID] " & vbCrLf & _
                                              "            ,[SupplierID] " & vbCrLf & _
                                              "            ,[ForwarderID] " & vbCrLf & _
                                              "            ,[Period] " & vbCrLf & _
                                              "            ,[CommercialCls] " & vbCrLf & _
                                              "            ,[EmergencyCls] " & vbCrLf & _
                                              "            ,[ShipCls] " & vbCrLf & _
                                              "            ,[ErrorStatus], [AcceptPONo] " & vbCrLf & _
                                              "            ,[OrderNo1] " & vbCrLf & _
                                              "            ,[ETDVendor1] " & vbCrLf & _
                                              "            ,[ETDPort1] " & vbCrLf & _
                                              "            ,[ETAPort1] " & vbCrLf & _
                                              "            ,[ETAFactory1] " & vbCrLf & _
                                              "            ,[UploadDate] " & vbCrLf & _
                                              "            ,[UploadUser] " & vbCrLf & _
                                              "            ,[EntryDate] " & vbCrLf & _
                                              "            ,[EntryUser]) " & vbCrLf & _
                                              "      VALUES " & vbCrLf & _
                                              "            ('" & ls_OrderNoErr & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("Period").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("CommercialCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ShipCls").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'OK', '" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ls_OrderNoErr & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETDVendor1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETDPort1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETAPort1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,'" & ds11.Tables(0).Rows(k)("ETAFactory1").ToString.Trim & "' " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "' " & vbCrLf & _
                                              "            ,getdate() " & vbCrLf & _
                                              "            ,'" & Session("UserID") & "') "

                                    SQLCom.CommandText = ls_Sql
                                    SQLCom.ExecuteNonQuery()
                                End If
                            End If
                        End If
                    Next
                End If

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

    Private Sub up_Save()
        Dim k As Integer
        Dim ls_Check As Boolean = False
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim ls_OrderNoErr As String = ""
        Dim ls_Forecast1 As Integer = 0
        Dim ls_Forecast2 As Integer = 0
        Dim ls_forecast3 As Integer = 0
        Dim ls_PrevForecast As Integer = 0
        Dim ls_Variance As Integer = 0
        Dim ls_VariancePercentage As Double = 0
        Dim ls_Week1 As Integer = 0

        Try
            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()
            SqlTran = SqlCon.BeginTransaction

            Dim SQLCom As SqlCommand = SqlCon.CreateCommand
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran

            Try
                '01. Cari ada data yg disubmit
                For i = 0 To grid.VisibleRowCount - 1
                    If grid.GetRowValues(i, "ErrorDesc").ToString.Trim <> "" Then
                        ls_Check = True
                        Exit For
                    End If
                Next i

                If ls_Check = True Then
                    lblInfo.Text = "[9999] Invalid data in this File Upload, please check the file again!"
                    Session("YA010IsSubmit") = lblInfo.Text
                    grid.JSProperties("cpMessage") = lblInfo.Text
                    Exit Sub
                End If

                For i = 0 To grid.VisibleRowCount - 1
                    
                    If InStr(1, grid.GetRowValues(i, "ErrorDesc"), "PO Already Approve by Supplier") = 1 Or _
                        InStr(1, grid.GetRowValues(i, "ErrorDesc"), "PO Already Send to Supplier") = 1 Or _
                        grid.GetRowValues(i, "ErrorDesc").ToString.Trim = "" Then

                        '01.1 delete data detail first
                        ls_Sql = "delete PO_Detail_Export where PONo = '" & grid.GetRowValues(i, "PONo").ToString.Trim & "' and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString.Trim & "'"
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()

                        '01.2 delete data master first
                        ls_Sql = "delete PO_Master_Export where PONo = '" & grid.GetRowValues(i, "PONo").ToString.Trim & "' and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString.Trim & "'"
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()

                        ls_Sql = "select * from UploadPOExport where PONo = '" & grid.GetRowValues(i, "PONo").ToString.Trim & "' and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString.Trim & "' order by PONo, AffiliateID, SupplierID "

                        SQLCom.CommandText = ls_Sql
                        Dim da11 As New SqlDataAdapter(SQLCom)
                        Dim ds11 As New DataSet
                        da11.Fill(ds11)

                        If ds11.Tables(0).Rows.Count > 0 Then
                            For k = 0 To ds11.Tables(0).Rows.Count - 1
                                If ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim = "M" Then
                                    ls_Week1 = ds11.Tables(0).Rows(k)("Week1").ToString.Trim
                                    ls_PrevForecast = 0 'ds11.Tables(0).Rows(k)("PreviousForecast").ToString.Trim
                                    ls_Forecast1 = ds11.Tables(0).Rows(k)("ForecastN1").ToString.Trim
                                    ls_Forecast2 = ds11.Tables(0).Rows(k)("ForecastN2").ToString.Trim
                                    ls_forecast3 = ds11.Tables(0).Rows(k)("ForecastN3").ToString.Trim
                                    ls_Variance = 0 'ls_PrevForecast - ls_Week1
                                    ls_VariancePercentage = 0 'IIf(ls_Week1 = 0, 0, (ls_Variance / ls_Week1) * 100)
                                Else
                                    ls_Week1 = ds11.Tables(0).Rows(k)("Week1").ToString.Trim
                                    ls_PrevForecast = 0
                                    ls_Forecast1 = 0
                                    ls_Forecast2 = 0
                                    ls_forecast3 = 0
                                    ls_Variance = 0
                                    ls_VariancePercentage = 0
                                End If
                                ' MsgBox(ds11.Tables(0).Rows(k)("Week1").ToString.Trim)
                                If ls_Week1 > 0 Then
                                    ls_Sql = " INSERT INTO [dbo].[PO_Detail_Export] " & vbCrLf & _
                                          "            ([PONo] " & vbCrLf & _
                                          "            ,[AffiliateID] " & vbCrLf & _
                                          "            ,[SupplierID] " & vbCrLf & _
                                          "            ,[ForwarderID] " & vbCrLf & _
                                          "            ,[OrderNo1] " & vbCrLf & _
                                          "            ,[PartNo] " & vbCrLf & _
                                          "            ,[Week1] " & vbCrLf & _
                                          "            ,[TotalPOQty] " & vbCrLf

                                    ls_Sql = ls_Sql + "            ,[PreviousForecast] " & vbCrLf & _
                                                      "            ,[Forecast1] " & vbCrLf & _
                                                      "            ,[Forecast2] " & vbCrLf & _
                                                      "            ,[Forecast3] " & vbCrLf & _
                                                      "            ,[Variance] " & vbCrLf & _
                                                      "            ,[VariancePercentage] " & vbCrLf & _
                                                      "            ,[EntryDate] " & vbCrLf & _
                                                      "            ,[EntryUser] " & vbCrLf & _
                                                      "            ,[POMOQ] " & vbCrLf & _
                                                      "            ,[POQtyBox]) " & vbCrLf

                                    ls_Sql = ls_Sql + "      VALUES " & vbCrLf & _
                                                      "            ('" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                                      "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                                      "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim.ToString & "' " & vbCrLf & _
                                                      "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                                      "            ,'" & ds11.Tables(0).Rows(k)("OrderNo1").ToString.Trim & "' "

                                    ls_Sql = ls_Sql + "            ,'" & ds11.Tables(0).Rows(k)("PartNo").ToString.Trim & "' " & vbCrLf & _
                                                      "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                      "            ,'" & ls_Week1 & "' " & vbCrLf & _
                                                      "            ,'" & ls_PrevForecast & "' "

                                    ls_Sql = ls_Sql + "            ,'" & ls_Forecast1 & "' " & vbCrLf & _
                                                      "            ,'" & ls_Forecast2 & "' " & vbCrLf & _
                                                      "            ,'" & ls_forecast3 & "' " & vbCrLf & _
                                                      "            ,'" & ls_Variance & "' " & vbCrLf & _
                                                      "            ,'" & ls_VariancePercentage & "' " & vbCrLf & _
                                                      "            , getdate() " & vbCrLf & _
                                                      "            ,'" & Session("UserID") & "' " & vbCrLf & _
                                                      "            ,'" & uf_GetMOQ(ds11.Tables(0).Rows(k)("PartNo").ToString.Trim, ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim.ToString, ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim) & "' " & vbCrLf & _
                                                      "            ,'" & uf_GetQtybox(ds11.Tables(0).Rows(k)("PartNo").ToString.Trim, ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim.ToString, ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim) & "' ) "

                                    SQLCom.CommandText = ls_Sql
                                    SQLCom.ExecuteNonQuery()
                                    ls_MsgID = "1001"
                                    ls_Detail = "ada"

                                    ls_Sql = "select * from PO_Master_Export where PONo = '" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' and AffiliateID = '" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' and SupplierID = '" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' and OrderNo1 ='" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "'"

                                    SQLCom.CommandText = ls_Sql
                                    Dim da7 As New SqlDataAdapter(SQLCom)
                                    Dim ds7 As New DataSet
                                    da7.Fill(ds7)

                                    If ds7.Tables(0).Rows.Count = 0 Then
                                        ls_Sql = " INSERT INTO [dbo].[PO_Master_Export] " & vbCrLf & _
                                                  "            ([PONo] " & vbCrLf & _
                                                  "            ,[AffiliateID] " & vbCrLf & _
                                                  "            ,[SupplierID] " & vbCrLf & _
                                                  "            ,[ForwarderID] " & vbCrLf & _
                                                  "            ,[Period] " & vbCrLf & _
                                                  "            ,[CommercialCls] " & vbCrLf & _
                                                  "            ,[EmergencyCls] " & vbCrLf & _
                                                  "            ,[ShipCls] " & vbCrLf & _
                                                  "            ,[ErrorStatus] " & vbCrLf & _
                                                  "            ,[OrderNo1] " & vbCrLf & _
                                                  "            ,[ETDVendor1] " & vbCrLf & _
                                                  "            ,[ETDPort1] " & vbCrLf & _
                                                  "            ,[ETAPort1] " & vbCrLf & _
                                                  "            ,[ETAFactory1] " & vbCrLf & _
                                                  "            ,[UploadDate] " & vbCrLf & _
                                                  "            ,[UploadUser] " & vbCrLf & _
                                                  "            ,[EntryDate] " & vbCrLf & _
                                                  "            ,[EntryUser]) " & vbCrLf & _
                                                  "      VALUES " & vbCrLf & _
                                                  "            ('" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("AffiliateID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("SupplierID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ForwarderID").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("Period").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("CommercialCls").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("EmergencyCls").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ShipCls").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'OK' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("PONo").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ETDVendor1").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ETDPort1").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ETAPort1").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,'" & ds11.Tables(0).Rows(k)("ETAFactory1").ToString.Trim & "' " & vbCrLf & _
                                                  "            ,getdate() " & vbCrLf & _
                                                  "            ,'" & Session("UserID") & "' " & vbCrLf & _
                                                  "            ,getdate() " & vbCrLf & _
                                                  "            ,'" & Session("UserID") & "') "
                                        'MsgBox(ls_Sql)
                                        SQLCom.CommandText = ls_Sql
                                        SQLCom.ExecuteNonQuery()
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next i

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

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click
        Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
    End Sub

    Private Sub btnApprove_Click(sender As Object, e As System.EventArgs) Handles btnApprove.Click
        Try
            Dim ls_Check As Boolean = False
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            'grid.JSProperties("cpMessage") = Session("YA010IsSubmit")
            If grid.VisibleRowCount = 0 Then Exit Sub
            For i = 0 To grid.VisibleRowCount - 1
                If InStr(1, grid.GetRowValues(i, "ErrorDesc"), "PO Already Approve by Supplier") = 1 Or _
                    InStr(1, grid.GetRowValues(i, "ErrorDesc"), "PO Already Send to Supplier") = 1 Then
                    ls_Check = True
                    Exit For
                End If
                'If
            Next i

            If ls_Check = True Then
                popUp2.ShowOnPageLoad = True
            Else
                up_Save()
            End If

EndProcedure:
            Session("YA010IsSubmit") = ""

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Function uf_GetMOQ(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(MOQ,0) MOQ FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

    Private Function uf_GetQtybox(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim Qty As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(QtyBox,0) Qty FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                Qty = dt.Rows(0)("Qty")
            End If
        End Using
        Return Qty
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(clsGlobal.ConnectionString)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function

End Class