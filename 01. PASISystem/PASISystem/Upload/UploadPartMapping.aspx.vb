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

Public Class UploadPartMapping
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A07"
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

        'ls_AllowUpdate = clsGlobal.Auth_UserUpdate(Session("UserID"), menuID)

        If (Not IsPostBack) AndAlso (Not IsCallback) Then
            'If ls_AllowUpdate = False Then
            'btnUpload.Enabled = False
            'btnClear.Enabled = False
            'btnSave.Enabled = False
            'btnDownload.Enabled = False
            'End If
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/Master/PartMapping.aspx")
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
            Dim fi As New FileInfo(Server.MapPath("~\Template\TemplatePO.xlsx"))
            If Not fi.Exists Then
                lblInfo.Text = "[9999] Excel Template Not Found !"
                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("Template PO.xlsx")

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
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "save"
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "6021", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    Else
                        Call up_Save()
                    End If
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

            ls_SQL = " select " & vbCrLf & _
                    " row_number() over (order by a.AffiliateID asc) as NoUrut, " & vbCrLf & _
                    " a.[AffiliateID], c.[AffiliateName], a.[PartNo], b.[PartName], a.[SupplierID], d.[SupplierName], a.[Quota], " & vbCrLf & _
                    " a.[LocationID], a.[PackingCls], e.[Description] as PackingDesc, a.[MOQ], a.[QtyBox], a.[BoxPallet], " & vbCrLf & _
                    " a.[NetWeight], a.[GrossWeight], a.[Length], a.[Width], a.[Height], [ErrorCls], " & vbCrLf & _
                    " xaffiliateid = mpm.[AffiliateID], xpartno = mpm.[PartNo], xsupplierid = mpm.[SupplierID], xquota = " & vbCrLf & _
                    " mpm.[Quota], xlocationid = mpm.[LocationID], xpackingcls = mpm.[PackingCls], xmoq = mpm.[MOQ], xqtybox " & vbCrLf & _
                    " = mpm.[QtyBox], xboxpallet = mpm.[BoxPallet], xnetweight = mpm.[NetWeight], xgrossweight = " & vbCrLf & _
                    " mpm.[GrossWeight], xlength = mpm.[Length], xwidth = mpm.[Width], xheight = mpm.[Height] " & vbCrLf & _
                    " from [UploadPartMapping] a " & vbCrLf & _
                    " left join MS_Parts b on a.PartNo = b.PartNo "

            ls_SQL = ls_SQL + " left join MS_Affiliate c on a.AffiliateID = c.AffiliateID " & vbCrLf & _
                                " left join MS_Supplier d on a.SupplierID = d.SupplierID " & vbCrLf & _
                                " left join MS_PackingCls e on a.PackingCls = e.PackingCls " & vbCrLf & _
                                " left join MS_PartMapping mpm on mpm.PartNo = a.PartNo and mpm.AffiliateID =  " & vbCrLf & _
                                " a.AffiliateID and mpm.SupplierID = a.SupplierID " & vbCrLf & _
                                " order by a.AffiliateID  " & vbCrLf

            'ls_SQL = " select  " & vbCrLf & _
            '      " 	row_number() over (order by a.AffiliateID asc) as NoUrut, " & vbCrLf & _
            '      " 	a.[AffiliateID], c.[AffiliateName], a.[PartNo], b.[PartName], a.[SupplierID], d.[SupplierName], a.[Quota], a.[LocationID], " & _
            '      " 	a.[PackingCls], e.[Description] as PackingDesc, a.[MOQ], a.[QtyBox], a.[BoxPallet], a.[NetWeight], a.[GrossWeight], a.[Length], a.[Width], a.[Height], [ErrorCls] "

            'ls_SQL = ls_SQL + " from [UploadPartMapping] a" & vbCrLf & _
            '                  " left join MS_Parts b on a.PartNo = b.PartNo" & vbCrLf & _
            '                  " left join MS_Affiliate c on a.AffiliateID = c.AffiliateID" & vbCrLf & _
            '                  " left join MS_Supplier d on a.SupplierID = d.SupplierID" & vbCrLf & _
            '                  " left join MS_PackingCls e on a.PackingCls = e.PackingCls" & vbCrLf & _
            '                  " order by a.AffiliateID "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' [AffiliateID],'' [LocationID],'' [AffiliateName],'' [PartNo],'' [PartName],'' [SupplierID],'' [SupplierName],'' [Quota]" & _
                " ,'' [PackingCls],'' [PackingDesc],'' [QtyBox],'' [BoxPallet],'' [NetWeight],'' [GrossWeight],'' [Length],'' [Width],'' [Height],'' [ErrorCls]"

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
        'Dim tempDate As Date
        Dim ls_MOQ As Double = 0
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = """"

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

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A3:N65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsMaster
                                    dtUploadDetail.PartNo = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.AffiliateID = IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))
                                    dtUploadDetail.SupplierID = IIf(IsDBNull(dtDetail.Rows(i).Item(2)), "", dtDetail.Rows(i).Item(2))
                                    dtUploadDetail.Quota = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), 100, dtDetail.Rows(i).Item(3))
                                    dtUploadDetail.Location = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), "", dtDetail.Rows(i).Item(4))
                                    dtUploadDetail.PackingCls = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), "", dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.MOQ = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), 0, dtDetail.Rows(i).Item(6))
                                    dtUploadDetail.QtyBox = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), 0, dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.BoxPallet = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), 0, dtDetail.Rows(i).Item(8))
                                    dtUploadDetail.NetWeight = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), 0, dtDetail.Rows(i).Item(9))
                                    dtUploadDetail.GrossWeight = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), 0, dtDetail.Rows(i).Item(10))
                                    dtUploadDetail.ItmLength = IIf(IsDBNull(dtDetail.Rows(i).Item(11)), 0, dtDetail.Rows(i).Item(11))
                                    dtUploadDetail.ItmWidth = IIf(IsDBNull(dtDetail.Rows(i).Item(12)), 0, dtDetail.Rows(i).Item(12))
                                    dtUploadDetail.ItmHeight = IIf(IsDBNull(dtDetail.Rows(i).Item(13)), 0, dtDetail.Rows(i).Item(13))

                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            ''01.01 Delete TempoaryData
                            ls_sql = "delete UploadPartMapping"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            ''02. Check PartNo, AffiliateID, SupplierID, PackingCls
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim PartMapping As clsMaster = dtUploadDetailList(i)

                                '02.1 Check PartNo
                                ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & PartMapping.PartNo & "' "
                                Dim sqlCmd1 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA1 As New SqlDataAdapter(sqlCmd1)
                                Dim ds1 As New DataSet
                                sqlDA1.Fill(ds1)

                                If ds1.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = ls_error & "PartNo not found in Part Master, please check again."
                                    Else
                                        ls_error = ls_error & "; " & "PartNo not found in Part Master, please check again."
                                    End If
                                End If

                                '02.2 Check AffiliateID
                                ls_sql = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID = '" & PartMapping.AffiliateID & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = ls_error & "Affiliate ID not found in Affiliate Master, please check again."
                                    Else
                                        ls_error = ls_error & "; " & "Affiliate ID not found in Affiliate Master, please check again."
                                    End If
                                End If

                                '02.3 Check SupplierID
                                ls_sql = "SELECT * FROM dbo.MS_Supplier WHERE SupplierID = '" & PartMapping.SupplierID & "' "
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)

                                If ds3.Tables(0).Rows.Count = 0 Then
                                    ls_error = "Supplier ID not found in Supplier Master, please check again."
                                End If

                                '02.3 Check PackingCls
                                ls_sql = "SELECT * FROM dbo.MS_PackingCls WHERE PackingCls = '" & PartMapping.PackingCls & "' "
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                If ds4.Tables(0).Rows.Count = 0 Then
                                    ls_error = "Supplier ID not found in Supplier Master, please check again."
                                End If

                                ls_sql = " INSERT INTO [dbo].[UploadPartMapping]( " & vbCrLf & _
                                          "            [PartNo], [AffiliateID], [SupplierID], [Quota], [LocationID], [PackingCls], [MOQ], [QtyBox], [BoxPallet], " & vbCrLf & _
                                          "            [NetWeight], [GrossWeight], [Length], [Width], [Height], [ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & PartMapping.PartNo & "'" & vbCrLf & _
                                          "            ,'" & PartMapping.AffiliateID & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.SupplierID & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.Quota & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.Location & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.PackingCls & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.MOQ & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.QtyBox & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.BoxPallet & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.NetWeight & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.GrossWeight & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.ItmLength & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.ItmWidth & "' " & vbCrLf & _
                                          "            ,'" & PartMapping.ItmHeight & "' " & vbCrLf & _
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
                        MyConnection.Close()
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
        'Dim tampung As String = ""
        Dim ls_Check As Boolean = False
        'Dim ls_PONo As String = ""
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""
        Dim ls_SupplierID As String = ""
        'Dim ls_Period As Date
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        'Dim ls_DoubleSupplier As Boolean = False
        'Dim ls_TempSupplierID As String = ""

        Dim shostname As String = System.Net.Dns.GetHostName
        Dim ls_Remarks As String = ""
        Dim PartNo As String = ""


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
                    If grid.GetRowValues(i, "PartNo") <> "" Then
                        ls_Sql = " IF NOT EXISTS (select * from MS_PartMapping where AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and SupplierID = '" & grid.GetRowValues(i, "SupplierID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "')" & vbCrLf & _
                                " BEGIN" & vbCrLf & _
                                " INSERT INTO [dbo].[MS_PartMapping] " & vbCrLf & _
                                "            ([AffiliateID] " & vbCrLf & _
                                "            ,[PartNo] " & vbCrLf & _
                                "            ,[SupplierID] " & vbCrLf & _
                                "            ,[Quota] " & vbCrLf & _
                                "            ,[LocationID] " & vbCrLf & _
                                "            ,[PackingCls] " & vbCrLf & _
                                "            ,[MOQ] " & vbCrLf & _
                                "            ,[QtyBox] " & vbCrLf & _
                                "            ,[BoxPallet] " & vbCrLf & _
                                "            ,[NetWeight] " & vbCrLf & _
                                "            ,[GrossWeight] " & vbCrLf & _
                                "            ,[Length] " & vbCrLf & _
                                "            ,[Width] " & vbCrLf & _
                                "            ,[Height] " & vbCrLf & _
                                "            ,[EntryDate] " & vbCrLf & _
                                "            ,[EntryUser] ) " & vbCrLf & _
                                "      VALUES " & vbCrLf & _
                                "            ('" & grid.GetRowValues(i, "AffiliateID") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "PartNo") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "SupplierID") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "Quota") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "LocationID") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "PackingCls") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "MOQ") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "QtyBox") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "BoxPallet") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "NetWeight") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "GrossWeight") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "Length") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "Width") & "' " & vbCrLf & _
                                "            ,'" & grid.GetRowValues(i, "Height") & "' " & vbCrLf & _
                                "            ,getdate() " & vbCrLf & _
                                "            ,'UPLOAD') " & vbCrLf & _
                                " END " & vbCrLf & _
                                " ELSE " & vbCrLf & _
                                " BEGIN" & vbCrLf & _
                                "       UPDATE [dbo].[MS_PartMapping] SET " & vbCrLf & _
                                "       [Quota] = '" & grid.GetRowValues(i, "Quota") & "', " & vbCrLf & _
                                "       [LocationID] = '" & grid.GetRowValues(i, "LocationID") & "', " & vbCrLf & _
                                "       [PackingCls] = '" & grid.GetRowValues(i, "PackingCls") & "', " & vbCrLf & _
                                "       [MOQ] = '" & grid.GetRowValues(i, "MOQ") & "', " & vbCrLf & _
                                "       [QtyBox] = '" & grid.GetRowValues(i, "QtyBox") & "', " & vbCrLf & _
                                "       [BoxPallet] = '" & grid.GetRowValues(i, "BoxPallet") & "', " & vbCrLf & _
                                "       [NetWeight] = '" & grid.GetRowValues(i, "NetWeight") & "', " & vbCrLf & _
                                "       [GrossWeight] = '" & grid.GetRowValues(i, "GrossWeight") & "', " & vbCrLf & _
                                "       [Length] = '" & grid.GetRowValues(i, "Length") & "', " & vbCrLf & _
                                "       [Width] = '" & grid.GetRowValues(i, "Width") & "', " & vbCrLf & _
                                "       [Height] = '" & grid.GetRowValues(i, "Height") & "' " & vbCrLf & _
                                "       WHERE AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and SupplierID = '" & grid.GetRowValues(i, "SupplierID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "' " & vbCrLf & _
                                " END"

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()


                    End If

                    PartNo = grid.GetRowValues(i, "PartNo")

                    If Not IsDBNull(grid.GetRowValues(i, "AffiliateID")) And Not IsDBNull(grid.GetRowValues(i, "xaffiliateid")) Then
                        If Trim(grid.GetRowValues(i, "AffiliateID")) <> Trim(grid.GetRowValues(i, "xaffiliateid")) Then
                            ls_Remarks = ls_Remarks + "AffiliateID " + Trim(grid.GetRowValues(i, "xaffiliateid")) & "->" & Trim(grid.GetRowValues(i, "AffiliateID")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "SupplierID")) And Not IsDBNull(grid.GetRowValues(i, "xsupplierid")) Then
                        If Trim(grid.GetRowValues(i, "SupplierID")) <> Trim(grid.GetRowValues(i, "xsupplierid")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "SupplierID " + Trim(grid.GetRowValues(i, "xsupplierid")) & "->" & Trim(grid.GetRowValues(i, "SupplierID")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Quota")) And Not IsDBNull(grid.GetRowValues(i, "xquota")) Then
                        If Trim(grid.GetRowValues(i, "Quota")) <> Trim(grid.GetRowValues(i, "xquota")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Quota " + Trim(grid.GetRowValues(i, "xquota")) & "->" & Trim(grid.GetRowValues(i, "Quota")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "LocationID")) And Not IsDBNull(grid.GetRowValues(i, "xlocationid")) Then
                        If Trim(grid.GetRowValues(i, "LocationID")) <> Trim(grid.GetRowValues(i, "xlocationid")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "LocationID " + Trim(grid.GetRowValues(i, "xlocationid")) & "->" & Trim(grid.GetRowValues(i, "LocationID")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "PackingCls")) And Not IsDBNull(grid.GetRowValues(i, "xpackingcls")) Then
                        If Trim(grid.GetRowValues(i, "PackingCls")) <> Trim(grid.GetRowValues(i, "xpackingcls")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PackingCls " + Trim(grid.GetRowValues(i, "xpackingcls")) & "->" & Trim(grid.GetRowValues(i, "PackingCls")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "MOQ")) And Not IsDBNull(grid.GetRowValues(i, "xmoq")) Then
                        If Trim(grid.GetRowValues(i, "MOQ")) <> Trim(grid.GetRowValues(i, "xmoq")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "MOQ " + Trim(grid.GetRowValues(i, "xmoq")) & "->" & Trim(grid.GetRowValues(i, "MOQ")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "QtyBox")) And Not IsDBNull(grid.GetRowValues(i, "xqtybox")) Then
                        If Trim(grid.GetRowValues(i, "QtyBox")) <> Trim(grid.GetRowValues(i, "xqtybox")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "QtyBox " + Trim(grid.GetRowValues(i, "xqtybox")) & "->" & Trim(grid.GetRowValues(i, "QtyBox")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "BoxPallet")) And Not IsDBNull(grid.GetRowValues(i, "xboxpallet")) Then
                        If Trim(grid.GetRowValues(i, "BoxPallet")) <> Trim(grid.GetRowValues(i, "xboxpallet")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "BoxPallet " + Trim(grid.GetRowValues(i, "xboxpallet")) & "->" & Trim(grid.GetRowValues(i, "BoxPallet")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "NetWeight")) And Not IsDBNull(grid.GetRowValues(i, "xnetweight")) Then
                        If Trim(grid.GetRowValues(i, "NetWeight")) <> Trim(grid.GetRowValues(i, "xnetweight")) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "NetWeight " + Trim(grid.GetRowValues(i, "xnetweight")) & "->" & Trim(grid.GetRowValues(i, "NetWeight")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "GrossWeight")) And Not IsDBNull(grid.GetRowValues(i, "xgrossweight")) Then
                        If CDbl(Trim(grid.GetRowValues(i, "GrossWeight"))) <> CDbl(Trim(grid.GetRowValues(i, "xgrossweight"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "GrossWeight " + Trim(grid.GetRowValues(i, "xgrossweight")) & "->" & Trim(grid.GetRowValues(i, "GrossWeight")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Length")) And Not IsDBNull(grid.GetRowValues(i, "xlength")) Then
                        If CDbl(Trim(grid.GetRowValues(i, "Length"))) <> CDbl(Trim(grid.GetRowValues(i, "xlength"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Length " + Trim(grid.GetRowValues(i, "xlength")) & "->" & Trim(grid.GetRowValues(i, "Length")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Width")) And Not IsDBNull(grid.GetRowValues(i, "xwidth")) Then
                        If CDbl(Trim(grid.GetRowValues(i, "Width"))) <> CDbl(Trim(grid.GetRowValues(i, "xwidth"))) Then
                            ls_Remarks = ls_Remarks + "Width " + Trim(grid.GetRowValues(i, "xwidth")) & "->" & Trim(grid.GetRowValues(i, "Width")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Height")) And Not IsDBNull(grid.GetRowValues(i, "xheight")) Then
                        If CDbl(Trim(grid.GetRowValues(i, "Height"))) <> CDbl(Trim(grid.GetRowValues(i, "xheight"))) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Height " + Trim(grid.GetRowValues(i, "xheight") & "->" & grid.GetRowValues(i, "Height")) & ""
                        End If
                    End If

                    If ls_Remarks <> "" Then
                        ls_Sql = " INSERT INTO MS_History (PCName, MenuID, OperationID, PartNo, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                        "VALUES ('" & shostname & "','" & menuID & "','U','" & PartNo & "','Update [" & ls_Remarks & "]', " & vbCrLf & _
                        "GETDATE(), '" & Session("UserID") & "') "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                        ls_Remarks = ""
                    End If

                Next i

                '2.3.1 Habis save semua,.. delete tada di tempolary table
                ls_Sql = "delete UploadPartMapping "

                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                '2.3.3 Commit transaction
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