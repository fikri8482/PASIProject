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

Public Class UploadDeliveryLocation
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A18"
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
        Response.Redirect("~/Master/DeliveryLocationMaster.aspx")
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

            ls_SQL = " select    " & vbCrLf & _
                     "   	row_number() over (order by a.AffiliateID asc) as NoUrut,   " & vbCrLf & _
                     "   	a.[AffiliateID],a.[DeliveryLocationCode],a.[DeliveryLocationName],a.[Address],a.[City],a.[PostalCode],   " & vbCrLf & _
                     "      a.[Phone1],a.[Phone2],a.[Fax],a.[NPWP],a.[PODeliveryBy],a.[DefaultCls],a.[ErrorCls],  " & vbCrLf & _
                     "      xaffiliateid = b.AffiliateID, xdeliverylocationcode = b.DeliveryLocationCode, xdeliverylocationname = b.DeliveryLocationName,  " & vbCrLf & _
                     "  	xaddress = b.Address, xcity = b.City, xpostalcode = b.PostalCode, xphone1 = b.Phone1, xphone2 = b.Phone2, xfax = b.Fax, xnpwp = b.NPWP,   " & vbCrLf & _
                     "  	xpodeliveryby = b.PODeliveryBy, xdefaultcls = b.DefaultCls    " & vbCrLf & _
                     "  from [UploadDeliveryPlace] a  " & vbCrLf & _
                     "  left join MS_DeliveryPlace b on a.AffiliateID = b.AffiliateID and a.DeliveryLocationCode = b.DeliveryLocationCode " & vbCrLf & _
                     "  order by a.AffiliateID "

            'ls_SQL = " select  " & vbCrLf & _
            '      " 	row_number() over (order by AffiliateID asc) as NoUrut, " & vbCrLf & _
            '      " 	[AffiliateID],[DeliveryLocationCode],[DeliveryLocationName],[Address],[City],[PostalCode], " & vbCrLf & _
            '      "     [Phone1],[Phone2],[Fax],[NPWP],[PODeliveryBy],[DefaultCls],[ErrorCls] "

            'ls_SQL = ls_SQL + " from [UploadDeliveryPlace] order by AffiliateID "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 1, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' [AffiliateID],'' [DeliveryLocationCode],'' [DeliveryLocationName],'' [Address],'' [City],'' [PostalCode],'' [Phone1],'' [Phone2],'' [Fax],'' [NPWP],'' [PODeliveryBy],'' [DefaultCls],'' [ErrorCls]"

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

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A3:L65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsMaster
                                    dtUploadDetail.AffiliateID = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.DeliveryLocationCode = IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))
                                    dtUploadDetail.DeliveryLocationName = IIf(IsDBNull(dtDetail.Rows(i).Item(2)), "", dtDetail.Rows(i).Item(2))
                                    dtUploadDetail.Address = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), "", dtDetail.Rows(i).Item(3))
                                    dtUploadDetail.City = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), "", dtDetail.Rows(i).Item(4))
                                    dtUploadDetail.PostalCode = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), "", dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.Phone1 = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), "", dtDetail.Rows(i).Item(6))
                                    dtUploadDetail.Phone2 = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), "", dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.Fax = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), "", dtDetail.Rows(i).Item(8))
                                    dtUploadDetail.NPWP = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), "", dtDetail.Rows(i).Item(9))
                                    dtUploadDetail.PODeliveryBy = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), "", dtDetail.Rows(i).Item(10))
                                    dtUploadDetail.DefaultCls = IIf(IsDBNull(dtDetail.Rows(i).Item(11)), "", dtDetail.Rows(i).Item(11))
                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            ''01.01 Delete TempoaryData
                            ls_sql = "delete UploadDeliveryPlace"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            ''02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim PO As clsMaster = dtUploadDetailList(i)

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID = '" & PO.AffiliateID & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    ls_error = "Affiliate ID not found in Affiliate Master, please check again."
                                End If

                                ''02.1 Check PartNo di MS_Part
                                'ls_sql = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID = '" & PO.AffiliateID & "' "
                                'Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                'Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                'Dim ds3 As New DataSet
                                'sqlDA3.Fill(ds3)

                                'If ds3.Tables(0).Rows.Count = 0 Then
                                '    If ls_error = "" Then
                                '        ls_error = ls_error & "Affiliate ID not found in Affiliate Master, please check again."
                                '    Else
                                '        ls_error = ls_error & "; " & "Affiliate ID not found in Affiliate Master, please check again."
                                '    End If
                                'End If

                                'If IsDate(PO.ETAAffiliate) = False Then
                                '    If ls_error = "" Then
                                '        ls_error = ls_error & "Invalid format date, please check again"
                                '    Else
                                '        ls_error = ls_error & "; " & "Invalid format date, please check again"
                                '    End If
                                'End If

                                'If IsDate(PO.ETDSupplier) = False Then
                                '    If ls_error = "" Then
                                '        ls_error = ls_error & "Invalid format date, please check again"
                                '    Else
                                '        ls_error = ls_error & "; " & "Invalid format date, please check again"
                                '    End If
                                'End If

                                ls_sql = " INSERT INTO [dbo].[UploadDeliveryPlace] " & vbCrLf & _
                                          "            ([AffiliateID], [DeliveryLocationCode],[DeliveryLocationName], [PODeliveryBy], [DefaultCls], [Address],[City],[PostalCode], " & vbCrLf & _
                                          "             [Phone1],[Phone2],[Fax],[NPWP],[ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & PO.AffiliateID & "'" & vbCrLf & _
                                          "            ,'" & PO.DeliveryLocationCode & "' " & vbCrLf & _
                                          "            ,'" & PO.DeliveryLocationName & "' " & vbCrLf & _
                                          "            ,'" & PO.PODeliveryBy & "' " & vbCrLf & _
                                          "            ,'" & PO.DefaultCls & "' " & vbCrLf & _
                                          "            ,'" & PO.Address & "' " & vbCrLf & _
                                          "            ,'" & PO.City & "' " & vbCrLf & _
                                          "            ,'" & PO.PostalCode & "' " & vbCrLf & _
                                          "            ,'" & PO.Phone1 & "' " & vbCrLf & _
                                          "            ,'" & PO.Phone2 & "' " & vbCrLf & _
                                          "            ,'" & PO.Fax & "' " & vbCrLf & _
                                          "            ,'" & PO.NPWP & "' " & vbCrLf & _
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
        Dim i As Integer, j As Integer
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
        'Dim PartNo As String = ""

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
                    If grid.GetRowValues(i, "AffiliateID") <> "" Then
                        ls_Sql = " IF NOT EXISTS (select * from MS_DeliveryPlace where AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and DeliveryLocationCode = '" & grid.GetRowValues(i, "DeliveryLocationCode") & "')" & vbCrLf & _
                                  " BEGIN" & vbCrLf & _
                                  " INSERT INTO [dbo].[MS_DeliveryPlace] " & vbCrLf & _
                                  "            ([AffiliateID] " & vbCrLf & _
                                  "            ,[DeliveryLocationCode] " & vbCrLf & _
                                  "            ,[DeliveryLocationName] " & vbCrLf & _
                                  "            ,[Address] " & vbCrLf & _
                                  "            ,[City] " & vbCrLf & _
                                  "            ,[PostalCode] " & vbCrLf & _
                                  "            ,[Phone1] " & vbCrLf & _
                                  "            ,[Phone2] " & vbCrLf & _
                                  "            ,[Fax] " & vbCrLf & _
                                  "            ,[NPWP] " & vbCrLf & _
                                  "            ,[PODeliveryBy] "

                        ls_Sql = ls_Sql + "            ,[DefaultCls] " & vbCrLf & _
                                          "            ,[EntryDate] " & vbCrLf & _
                                          "            ,[EntryUser] ) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & grid.GetRowValues(i, "AffiliateID") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "DeliveryLocationCode") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "DeliveryLocationName") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "Address") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "City") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "PostalCode") & "' "

                        ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "Phone1") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "Phone2") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "Fax") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "NPWP") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "PODeliveryBy") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "DefaultCls") & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'UPLOAD') " & vbCrLf & _
                                          " END " & vbCrLf & _
                                          " ELSE " & vbCrLf & _
                                          " BEGIN" & vbCrLf & _
                                          "      UPDATE [dbo].[MS_DeliveryPlace] SET " & vbCrLf & _
                                          "       [DeliveryLocationName] = '" & grid.GetRowValues(i, "DeliveryLocationName") & "' " & vbCrLf & _
                                          "       ,[Address] = '" & grid.GetRowValues(i, "Address") & "' " & vbCrLf & _
                                          "       ,[City] = '" & grid.GetRowValues(i, "City") & "' " & vbCrLf & _
                                          "       ,[PostalCode] = '" & grid.GetRowValues(i, "PostalCode") & "' " & vbCrLf & _
                                          "       ,[Phone1] = '" & grid.GetRowValues(i, "Phone1") & "' " & vbCrLf & _
                                          "       ,[Phone2] = '" & grid.GetRowValues(i, "Phone2") & "' " & vbCrLf & _
                                          "       ,[Fax] = '" & grid.GetRowValues(i, "Fax") & "' " & vbCrLf & _
                                          "       ,[NPWP] = '" & grid.GetRowValues(i, "NPWP") & "' " & vbCrLf & _
                                          "       ,[PODeliveryBy] = '" & grid.GetRowValues(i, "PODeliveryBy") & "' " & vbCrLf & _
                                          "       ,[DefaultCls] = '" & grid.GetRowValues(i, "DefaultCls") & "' " & vbCrLf & _
                                          "      WHERE [AffiliateID] = '" & grid.GetRowValues(i, "AffiliateID") & "' and DeliveryLocationCode = '" & grid.GetRowValues(i, "DeliveryLocationCode") & "' " & vbCrLf & _
                                          " END"

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    End If

                    'PartNo = grid.GetRowValues(i, "PartNo")

                    'If Trim(grid.GetRowValues(i, "AffiliateID")) <> Trim(grid.GetRowValues(i, "xaffiliateid")) Then
                    '    ls_Remarks = ls_Remarks + "AffiliateID " + Trim(grid.GetRowValues(i, "xaffiliateid")) & "->" & Trim(grid.GetRowValues(i, "AffiliateID")) & " "
                    'End If

                    If Not IsDBNull(grid.GetRowValues(i, "DeliveryLocationName")) And Not IsDBNull(grid.GetRowValues(i, "xdeliverylocationname")) Then
                        If Trim(grid.GetRowValues(i, "DeliveryLocationName").ToString) <> Trim(grid.GetRowValues(i, "xdeliverylocationname").ToString) Then
                            ls_Remarks = ls_Remarks + "DeliveryLocationName " + Trim(grid.GetRowValues(i, "xdeliverylocationname").ToString) & "->" & Trim(grid.GetRowValues(i, "DeliveryLocationName").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Address")) And Not IsDBNull(grid.GetRowValues(i, "xaddress")) Then
                        If Trim(grid.GetRowValues(i, "Address").ToString) <> Trim(grid.GetRowValues(i, "xaddress").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Address " + Trim(grid.GetRowValues(i, "xaddress").ToString) & "->" & Trim(grid.GetRowValues(i, "Address").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "City")) And Not IsDBNull(grid.GetRowValues(i, "xcity")) Then
                        If Trim(grid.GetRowValues(i, "City").ToString) <> Trim(grid.GetRowValues(i, "xcity").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "City " + Trim(grid.GetRowValues(i, "xcity").ToString) & "->" & Trim(grid.GetRowValues(i, "City").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "PostalCode")) And Not IsDBNull(grid.GetRowValues(i, "xpostalcode")) Then
                        If Trim(grid.GetRowValues(i, "PostalCode").ToString) <> Trim(grid.GetRowValues(i, "xpostalcode").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PostalCode " + Trim(grid.GetRowValues(i, "PostalCode").ToString) & "->" & Trim(grid.GetRowValues(i, "PostalCode").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Phone1")) And Not IsDBNull(grid.GetRowValues(i, "xphone1")) Then
                        If Trim(grid.GetRowValues(i, "Phone1").ToString) <> Trim(grid.GetRowValues(i, "xphone1").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Phone1 " + Trim(grid.GetRowValues(i, "xphone1").ToString) & "->" & Trim(grid.GetRowValues(i, "Phone1").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Phone2")) And Not IsDBNull(grid.GetRowValues(i, "xphone2")) Then
                        If Trim(grid.GetRowValues(i, "Phone2").ToString) <> Trim(grid.GetRowValues(i, "xphone2").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Phone2 " + Trim(grid.GetRowValues(i, "xphone2").ToString) & "->" & Trim(grid.GetRowValues(i, "Phone2").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Fax")) And Not IsDBNull(grid.GetRowValues(i, "xfax")) Then
                        If Trim(grid.GetRowValues(i, "Fax").ToString) <> Trim(grid.GetRowValues(i, "xfax").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Fax " + Trim(grid.GetRowValues(i, "xfax").ToString) & "->" & Trim(grid.GetRowValues(i, "Fax").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "NPWP")) And Not IsDBNull(grid.GetRowValues(i, "xnpwp")) Then
                        If Trim(grid.GetRowValues(i, "NPWP").ToString) <> Trim(grid.GetRowValues(i, "xnpwp").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "NPWP " + Trim(grid.GetRowValues(i, "xnpwp").ToString) & "->" & Trim(grid.GetRowValues(i, "NPWP").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "PODeliveryBy")) And Not IsDBNull(grid.GetRowValues(i, "xpodeliveryby")) Then
                        If Trim(grid.GetRowValues(i, "PODeliveryBy").ToString) <> Trim(grid.GetRowValues(i, "xpodeliveryby").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PODeliveryBy " + Trim(grid.GetRowValues(i, "xpodeliveryby").ToString) & "->" & Trim(grid.GetRowValues(i, "PODeliveryBy").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "DefaultCls")) And Not IsDBNull(grid.GetRowValues(i, "xdefaultcls")) Then
                        If Trim(grid.GetRowValues(i, "DefaultCls").ToString) <> Trim(grid.GetRowValues(i, "xdefaultcls").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "DefaultCls " + Trim(grid.GetRowValues(i, "xdefaultcls").ToString) & "->" & Trim(grid.GetRowValues(i, "DefaultCls").ToString) & ""
                        End If
                    End If

                    If ls_Remarks <> "" Then
                        ls_Sql = " INSERT INTO MS_History (PCName, MenuID, OperationID, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                        "VALUES ('" & shostname & "','" & menuID & "','U','Update [" & ls_Remarks & "]', " & vbCrLf & _
                        "GETDATE(), '" & Session("UserID") & "') "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                        ls_Remarks = ""
                    End If


                Next i

                '2.3.1 Habis save semua,.. delete tada di tempolary table
                ls_Sql = "delete UploadDeliveryPlace "

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