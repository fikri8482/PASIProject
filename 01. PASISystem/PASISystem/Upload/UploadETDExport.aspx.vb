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

Public Class UploadETDExport
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A25"
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

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/Master/ETDExport.aspx")
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Uploader.NullText = "Click here to browse files..."

        lblInfo.Text = ""

        Uploader.Enabled = True
        btnSave.Enabled = True
        btnUpload.Enabled = True

        up_GridLoadWhenEventChange()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        up_Import()
    End Sub

    Private Sub ASPxCallback1_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
        Try
            Dim fi As New FileInfo(Server.MapPath("~\Template\TemplateETDExport.xlsx"))
            If Not fi.Exists Then
                lblInfo.Text = "[9999] Excel Template Not Found !"
                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("Template ETD Export.xlsx")

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        End Try

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("ErrorCls") = "" Then
        Else
            e.Cell.BackColor = Color.Red
        End If
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
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

            'ls_SQL = "  SELECT  " & vbCrLf & _
            '      " 	ROW_NUMBER() OVER (ORDER BY SupplierID) AS RowNumber, " & vbCrLf & _
            '      " 	Period, SupplierID, AffiliateID, Week, ETDVendor, ETAForwarder, ETDPort, ETAPort, ETAFactory, ErrorCls "

            'ls_SQL = ls_SQL + " from UploadETDExport order by SupplierID, AffiliateID "

            ls_SQL = " SELECT   " & vbCrLf & _
                  "  	ROW_NUMBER() OVER (ORDER BY a.SupplierID) AS RowNumber,  " & vbCrLf & _
                  "  	a.Period, a.SupplierID, a.AffiliateID, a.Week, a.ETDVendor, a.ETAForwarder, a.ETDPort, a.ETAPort, a.ETAFactory, a.CutOfDate, a.ErrorCls,  " & vbCrLf & _
                  " 	xETDVendor = b.ETDVendor, xETAForwarder = b.ETAForwarder, xETDPort = b.ETDPort, xETAPort = b.ETAPort, xETAFactory = b.ETAFactory, xCutOfDate = b.CutOfDate " & vbCrLf & _
                  " from UploadETDExport a  " & vbCrLf & _
                  " left join MS_ETD_Export b on a.Period = b.Period and a.SupplierID = b.SupplierID and a.AffiliateID = b.AffiliateID and a.Week = b.Week " & vbCrLf & _
                  " order by a.SupplierID, a.AffiliateID  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' RowNumber, '' SupplierID, '' AffiliateID, '' Week,'' ETAForwarder, '' ETDVendor , '' ETDPort, '' ETAPort , '' ETAFactory, '' ErrorCls, '' xETDVendor, '' xETAForwarder, '' xETDPort, '' xETAPort, '' xETAFactory "

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
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = """"
        Dim tempDate As Date

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
                Select Case FileExt
                    Case ".xls"
                        'Excel 97-03
                        connStr = ConfigurationManager.ConnectionStrings("Excel03ConString").ConnectionString
                    Case ".xlsx"
                        'Excel 07
                        'connStr = ConfigurationManager.ConnectionStrings("Excel07ConString").ConnectionString
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

                        'Get Detail Data
                        Dim dtUploadDetailList As New List(Of clsMaster)

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A3:J65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsMaster

                                    Try
                                        'tempDate = "01-" & dtDetail.Rows(i).Item(0)
                                        tempDate = Format(dtDetail.Rows(i).Item(0), "MM") & "-01-" & Format(dtDetail.Rows(i).Item(0), "yyyy")
                                        dtUploadDetail.Period = tempDate
                                        'Session("Period") = tempDate
                                    Catch ex As Exception
                                        lblInfo.Text = "[9999] Invalid Period, please check the file again!"
                                        grid.JSProperties("cpMessage") = lblInfo.Text
                                        Exit Sub
                                    End Try

                                    dtUploadDetail.SupplierID = dtDetail.Rows(i).Item(2)
                                    dtUploadDetail.AffiliateID = dtDetail.Rows(i).Item(1)
                                    dtUploadDetail.Week = dtDetail.Rows(i).Item(3)

                                    'dtUploadDetail.ETDVendor = dtDetail.Rows(i).Item(4)
                                    If IsDBNull(dtDetail.Rows(i).Item(4)) = False Then
                                        dtUploadDetail.ETDVendor = dtDetail.Rows(i).Item(4)
                                    Else
                                        dtUploadDetail.ETDVendor = "NULL"
                                    End If

                                    'dtUploadDetail.ETAForwarder = dtDetail.Rows(i).Item(5)
                                    If IsDBNull(dtDetail.Rows(i).Item(5)) = False Then
                                        dtUploadDetail.ETAForwarder = dtDetail.Rows(i).Item(5)
                                    Else
                                        dtUploadDetail.ETAForwarder = "NULL"
                                    End If

                                    'dtUploadDetail.ETDPort = dtDetail.Rows(i).Item(6)
                                    If IsDBNull(dtDetail.Rows(i).Item(6)) = False Then
                                        dtUploadDetail.ETDPort = dtDetail.Rows(i).Item(6)
                                    Else
                                        dtUploadDetail.ETDPort = "NULL"
                                    End If

                                    'dtUploadDetail.ETAPort = dtDetail.Rows(i).Item(7)
                                    If IsDBNull(dtDetail.Rows(i).Item(7)) = False Then
                                        dtUploadDetail.ETAPort = dtDetail.Rows(i).Item(7)
                                    Else
                                        dtUploadDetail.ETAPort = "NULL"
                                    End If

                                    'dtUploadDetail.ETAFactory = dtDetail.Rows(i).Item(8)
                                    If IsDBNull(dtDetail.Rows(i).Item(8)) = False Then
                                        dtUploadDetail.ETAFactory = dtDetail.Rows(i).Item(8)
                                    Else
                                        dtUploadDetail.ETAFactory = "NULL"
                                    End If

                                    'dtUploadDetail.ETAFactory = dtDetail.Rows(i).Item(8)
                                    If IsDBNull(dtDetail.Rows(i).Item(9)) = False Then
                                        dtUploadDetail.CutOfDate = dtDetail.Rows(i).Item(9)
                                    Else
                                        dtUploadDetail.CutOfDate = "NULL"
                                    End If

                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadETDExport")
                            ''01.01 Delete TempoaryData
                            ls_sql = "delete uploadETDExport"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            ''02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim PO As clsMaster = dtUploadDetailList(i)

                                If IsDate(PO.Period) = False Then
                                    If ls_error = "" Then
                                        ls_error = ls_error & "Invalid format date, please check again"
                                    Else
                                        ls_error = ls_error & "; " & "Invalid format date, please check again"
                                    End If
                                End If

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID = '" & PO.SupplierID & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    ls_error = "Supplier ID not found in Supplier Master, please check again."
                                End If

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID = '" & PO.AffiliateID & "' "
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)

                                If ds3.Tables(0).Rows.Count = 0 Then
                                    ls_error = "Affiliate ID not found in Affiliate Master, please check again."
                                End If

                                '03.1 Check ETD yang sama dalam 1 periode di MS_ETDExport
                                ls_sql = "Select * From [dbo].[MS_ETD_Export] Where Period = '" & Format(PO.Period, "yyyy-MM-01") & "' " & vbCrLf & _
                                         "And SupplierID = '" & Trim(PO.SupplierID) & "' And AffiliateID = '" & Trim(PO.AffiliateID) & "' " & vbCrLf & _
                                         "And [ETDVendor] = '" & Format(CDate(PO.ETDVendor), "yyyy-MM-dd") & "' And [ETAForwarder] = '" & Format(CDate(PO.ETAForwarder), "yyyy-MM-dd") & "'" & vbCrLf & _
                                         "And [ETDPort] = '" & Format(CDate(PO.ETDPort), "yyyy-MM-dd") & "' And [ETAPort] = '" & Format(CDate(PO.ETAPort), "yyyy-MM-dd") & "'" & vbCrLf & _
                                         "And [ETAFactory] = '" & Format(CDate(PO.ETAFactory), "yyyy-MM-dd") & "'"
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                If ds4.Tables(0).Rows.Count > 0 Then
                                    ls_error = "Data already Exist, please check again"
                                End If

                                If PO.ETDVendor <> "NULL" Then
                                    If IsDate(PO.ETDVendor) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                If PO.ETAForwarder <> "NULL" Then
                                    If IsDate(PO.ETAForwarder) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                If PO.ETDPort <> "NULL" Then
                                    If IsDate(PO.ETDPort) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                If PO.ETAPort <> "NULL" Then
                                    If IsDate(PO.ETAPort) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                If PO.ETAFactory <> "NULL" Then
                                    If IsDate(PO.ETAFactory) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                '03 CHECK ETD YANG SAMA PADA 1 TEMPLATE
                                For a = 0 To dtUploadDetailList.Count - 1
                                    Dim data As clsMaster = dtUploadDetailList(a)

                                    If PO.Period = data.Period And PO.SupplierID = data.SupplierID And PO.AffiliateID = data.AffiliateID And PO.Week <> data.Week Then
                                        If PO.Period = data.Period And PO.SupplierID = data.SupplierID And PO.AffiliateID = data.AffiliateID And PO.ETDVendor = data.ETDVendor Then
                                            ls_error = ls_error & "Data already Exist on another week, please check again"
                                        End If
                                    End If
                                Next

                                ls_sql = " INSERT INTO [dbo].[UploadETDExport] " & vbCrLf & _
                                          "            ([Period], [SupplierID], [AffiliateID], [Week], [ETDVendor], [ETAForwarder], [ETDPort], [ETAPort], [ETAFactory], [CutOfDate] ,[ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & PO.Period & "'" & vbCrLf & _
                                          "            ,'" & PO.SupplierID & "' " & vbCrLf & _
                                          "            ,'" & PO.AffiliateID & "' " & vbCrLf & _
                                          "            ,'" & PO.Week & "' " & vbCrLf & _
                                          "            ," & IIf(PO.ETDVendor = "NULL", PO.ETDVendor, "'" & PO.ETDVendor & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.ETAForwarder = "NULL", PO.ETAForwarder, "'" & PO.ETAForwarder & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.ETDPort = "NULL", PO.ETDPort, "'" & PO.ETDPort & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.ETAPort = "NULL", PO.ETAPort, "'" & PO.ETAPort & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.ETAFactory = "NULL", PO.ETAFactory, "'" & PO.ETAFactory & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.CutOfDate = "NULL", PO.CutOfDate, "'" & PO.CutOfDate & "'") & "" & vbCrLf & _
                                          "            ,'" & ls_error & "') " & vbCrLf
                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()
                            Next
                            sqlTran.Commit()

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
        Dim i As Integer ', j As Integer
        Dim ls_Check As Boolean = False
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim shostname As String = System.Net.Dns.GetHostName
        Dim ls_Remarks As String = ""

        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "ErrorCls").ToString.Trim <> "" Then
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

            Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
            Dim SqlTran As SqlTransaction

            SqlCon.Open()

            SqlTran = SqlCon.BeginTransaction

            Try
                '2.1 delete data 
                Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                SQLCom.Connection = SqlCon
                SQLCom.Transaction = SqlTran

                '2.2 Insert New Detail Data
                Dim ls_SuppCount As Integer = 1
                For i = 0 To grid.VisibleRowCount - 1
                    If IsNothing(grid.GetRowValues(i, "Period")) = False Then
                        If grid.GetRowValues(i, "SupplierID") <> "" Then
                            If grid.GetRowValues(i, "AffiliateID") <> "" Then
                                ls_Sql = " IF NOT EXISTS (select * from MS_ETD_Export where Period = '" & grid.GetRowValues(i, "Period") & "' AND SupplierID = '" & Trim(grid.GetRowValues(i, "SupplierID")) & "' and AffiliateID = '" & Trim(grid.GetRowValues(i, "AffiliateID")) & "' and Week = '" & Trim(grid.GetRowValues(i, "Week")) & "')" & vbCrLf & _
                                          " BEGIN" & vbCrLf & _
                                          "      INSERT INTO [dbo].[MS_ETD_Export] " & vbCrLf & _
                                          "            ([Period] " & vbCrLf & _
                                          "            ,[SupplierID] " & vbCrLf & _
                                          "            ,[AffiliateID] " & vbCrLf & _
                                          "            ,[Week] " & vbCrLf & _
                                          "            ,[ETDVendor] " & vbCrLf & _
                                          "            ,[ETAForwarder] " & vbCrLf & _
                                          "            ,[ETDPort] " & vbCrLf & _
                                          "            ,[ETAPort] " & vbCrLf & _
                                          "            ,[ETAFactory] " & vbCrLf & _
                                          "            ,[CutOfDate]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & grid.GetRowValues(i, "Period") & "' " & vbCrLf & _
                                          "            ,'" & Trim(grid.GetRowValues(i, "SupplierID")) & "' " & vbCrLf & _
                                          "            ,'" & Trim(grid.GetRowValues(i, "AffiliateID")) & "' " & vbCrLf & _
                                          "            ,'" & Trim(grid.GetRowValues(i, "Week")) & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "ETDVendor") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "ETAForwarder") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "ETDPort") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "ETAPort") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "ETAFactory") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "CutOfDate") & "') " & vbCrLf & _
                                          " END " & vbCrLf & _
                                          " ELSE " & vbCrLf & _
                                          " BEGIN" & vbCrLf & _
                                          "      UPDATE [dbo].[MS_ETD_Export] SET " & vbCrLf & _
                                          "            [ETDVendor] ='" & grid.GetRowValues(i, "ETDVendor") & "', " & vbCrLf & _
                                          "            [ETDPort] ='" & grid.GetRowValues(i, "ETDPort") & "', " & vbCrLf & _
                                          "            [ETAForwarder] ='" & grid.GetRowValues(i, "ETAForwarder") & "', " & vbCrLf & _
                                          "            [ETAPort] ='" & grid.GetRowValues(i, "ETAPort") & "', " & vbCrLf & _
                                          "            [ETAFactory] ='" & grid.GetRowValues(i, "ETAFactory") & "', " & vbCrLf & _
                                          "            [CutOfDate] ='" & grid.GetRowValues(i, "CutOfDate") & "' " & vbCrLf & _
                                          "      WHERE [Period] = '" & grid.GetRowValues(i, "Period") & "' and [SupplierID] = '" & Trim(grid.GetRowValues(i, "SupplierID")) & "' and [AffiliateID] = '" & Trim(grid.GetRowValues(i, "AffiliateID")) & "' and Week = '" & Trim(grid.GetRowValues(i, "Week")) & "'" & vbCrLf & _
                                          " END"

                                SQLCom.CommandText = ls_Sql
                                SQLCom.ExecuteNonQuery()
                            End If
                        End If
                    End If
                    'End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ETDVendor")) And Not IsDBNull(grid.GetRowValues(i, "xETDVendor"))) And (grid.GetRowValues(i, "ETDVendor").ToString <> "" And grid.GetRowValues(i, "xETDVendor").ToString <> "") Then
                        If (Trim(Format(grid.GetRowValues(i, "ETDVendor"), "dd-MM-yyyy")) <> Trim(Format(grid.GetRowValues(i, "xETDVendor"), "dd-MM-yyyy"))) Then
                            ls_Remarks = ls_Remarks + "ETDVendor " + Trim(Format(grid.GetRowValues(i, "xETDVendor"), "dd-MMM-yyyy")) & " " & "->" & " " & Trim(Format(grid.GetRowValues(i, "ETDVendor"), "dd-MMM-yyyy")) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ETAForwarder")) And Not IsDBNull(grid.GetRowValues(i, "xETAForwarder"))) And (grid.GetRowValues(i, "ETAForwarder").ToString <> "" And grid.GetRowValues(i, "xETAForwarder").ToString <> "") Then
                        If (Trim(Format(grid.GetRowValues(i, "ETAForwarder"), "dd-MM-yyyy")) <> Trim(Format(grid.GetRowValues(i, "xETAForwarder"), "dd-MM-yyyy"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ETAForwarder " + Trim(Format(grid.GetRowValues(i, "xETAForwarder"), "dd-MMM-yyyy")) & " " & "->" & " " & Trim(Format(grid.GetRowValues(i, "ETAForwarder"), "dd-MMM-yyyy")) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ETDPort")) And Not IsDBNull(grid.GetRowValues(i, "xETDPort"))) And (grid.GetRowValues(i, "ETDPort").ToString <> "" And grid.GetRowValues(i, "xETDPort").ToString <> "") Then
                        If (Trim(Format(grid.GetRowValues(i, "ETDPort"), "dd-MM-yyyy")) <> Trim(Format(grid.GetRowValues(i, "xETDPort"), "dd-MM-yyyy"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ETDPort " + Trim(Format(grid.GetRowValues(i, "xETDPort"), "dd-MMM-yyyy")) & " " & "->" & " " & Trim(Format(grid.GetRowValues(i, "ETDPort"), "dd-MMM-yyyy")) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ETAPort")) And Not IsDBNull(grid.GetRowValues(i, "xETAPort"))) And (grid.GetRowValues(i, "ETAPort").ToString <> "" And grid.GetRowValues(i, "xETAPort").ToString <> "") Then
                        If (Trim(Format(grid.GetRowValues(i, "ETAPort"), "dd-MM-yyyy")) <> Trim(Format(grid.GetRowValues(i, "xETAPort"), "dd-MM-yyyy"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ETAPort " + Trim(Format(grid.GetRowValues(i, "xETAPort"), "dd-MMM-yyyy")) & " " & "->" & " " & Trim(Format(grid.GetRowValues(i, "ETAPort"), "dd-MMM-yyyy")) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "ETAFactory")) And Not IsDBNull(grid.GetRowValues(i, "xETAFactory"))) And (grid.GetRowValues(i, "ETAFactory").ToString <> "" And grid.GetRowValues(i, "xETAFactory").ToString <> "") Then
                        If (Trim(Format(grid.GetRowValues(i, "ETAFactory"), "dd-MM-yyyy")) <> Trim(Format(grid.GetRowValues(i, "xETAFactory"), "dd-MM-yyyy"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "ETAFactory " + Trim(Format(grid.GetRowValues(i, "xETAFactory"), "dd-MMM-yyyy")) & " " & "->" & " " & Trim(Format(grid.GetRowValues(i, "ETAFactory"), "dd-MMM-yyyy")) & ""
                        End If
                    End If

                    If ls_Remarks <> "" Then
                        ls_Sql = " INSERT INTO MS_History (PCName, MenuID, OperationID, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U','Update [" & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                        ls_Remarks = ""
                    End If
                Next i

                '2.3.2 Commit transaction
                ls_MsgID = "1001"
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