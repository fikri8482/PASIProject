'Update By Robby
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

Public Class UploadSupplier
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A03"
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
        Response.Redirect("~/Master/SupplierMaster.aspx")
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

            ls_SQL = " select   " & vbCrLf & _
                  " 	row_number() over (order by a.SupplierID asc) as NoUrut,  " & vbCrLf & _
                  " 	a.[SupplierID], a.[SupplierName], a.[SupplierType], a.[SupplierCode], a.[Address],  " & vbCrLf & _
                  " 	a.[City], a.[PostalCode], a.[Phone1], a.[Phone2], a.[Fax], a.[NPWP], " & vbCrLf & _
                  " 	a.[ErrorCls], a.[LabelNo], xSupplierID  = b.[SupplierID], xSupplierName = b.[SupplierName], " & vbCrLf & _
                  " 	xSupplierType = b.[SupplierType], xSupplierCode = b.[SupplierCode], xAddress = b.[Address], " & vbCrLf & _
                  " 	xCity = b.[City], xPostalCode = b.[PostalCode], xPhone1 = b.[Phone1], xPhone2 = b.[Phone2],  " & vbCrLf & _
                  " 	xFax = b.[Fax], xNPWP = b.[NPWP], xLabelCode = b.[LabelCode] " & vbCrLf & _
                  " from [UploadSupplier] a  " & vbCrLf & _
                  " left join MS_Supplier b on a.SupplierID = b.SupplierID "

            ls_SQL = ls_SQL + " order by a.SupplierID " 

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

            'ls_SQL = " select top 0 '' NoUrut, '' [SupplierID],'' [SupplierName],'' [SupplierType], '' [SupplierCode],'' [Address],'' [City],'' [PostalCode],'' [Phone1],'' [Phone2],'' [Fax],'' [NPWP],'' [ErrorCls]"

            ls_SQL = " select top 0 '' NoUrut, '' [SupplierID],'' [SupplierName],'' [SupplierType], '' [SupplierCode],'' [Address],'' [City],'' [PostalCode],'' [Phone1],'' [Phone2],'' [Fax],'' [NPWP],'' [ErrorCls], '' [LabelNo], " & vbCrLf & _
                     " '' [xSupplierID],'' [xSupplierName],'' [xSupplierType], '' [xSupplierCode],'' [xAddress],'' [xCity],'' [xPostalCode],'' [xPhone1],'' [xPhone2],'' [xFax],'' [xNPWP], '' [xLabelCode] "

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
                                    dtUploadDetail.SupplierID = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.SupplierName = IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))
                                    dtUploadDetail.SupplierType = IIf(IsDBNull(dtDetail.Rows(i).Item(2)), "", dtDetail.Rows(i).Item(2))
                                    dtUploadDetail.SupplierCode = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), "", dtDetail.Rows(i).Item(3))
                                    dtUploadDetail.LabelCode = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), "", dtDetail.Rows(i).Item(4))
                                    dtUploadDetail.Address = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), "", dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.City = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), "", dtDetail.Rows(i).Item(6))
                                    dtUploadDetail.PostalCode = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), "", dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.Phone1 = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), "", dtDetail.Rows(i).Item(8))
                                    dtUploadDetail.Phone2 = IIf(IsDBNull(dtDetail.Rows(i).Item(9)), "", dtDetail.Rows(i).Item(9))
                                    dtUploadDetail.Fax = IIf(IsDBNull(dtDetail.Rows(i).Item(10)), "", dtDetail.Rows(i).Item(10))
                                    dtUploadDetail.NPWP = IIf(IsDBNull(dtDetail.Rows(i).Item(11)), "", dtDetail.Rows(i).Item(11))
                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            ''01.01 Delete TempoaryData
                            ls_sql = "delete UploadSupplier"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            ''02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim PO As clsMaster = dtUploadDetailList(i)

                                If PO.SupplierID.Trim.Length > 20 Then
                                    ls_error = "Supplier ID max length only 20"
                                End If

                                'If PO.SupplierName.Trim.Length > 100 Then
                                '    If ls_error = "" Then
                                '        ls_error = ls_error & "Affiliate ID not found in Affiliate Master, please check again."
                                '    Else
                                '        ls_error = ls_error & "; " & "Affiliate ID not found in Affiliate Master, please check again."
                                '    End If
                                'End If
                                

                                ls_sql = " INSERT INTO [dbo].[UploadSupplier] " & vbCrLf & _
                                          "            ([SupplierID], [SupplierName], [SupplierType], [SupplierCode], [LabelNo], [Address],[City],[PostalCode], " & vbCrLf & _
                                          "             [Phone1],[Phone2],[Fax],[NPWP],[ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & PO.SupplierID & "'" & vbCrLf & _
                                          "            ,'" & PO.SupplierName & "' " & vbCrLf & _
                                          "            ,'" & PO.SupplierType & "' " & vbCrLf & _
                                          "            ,'" & PO.SupplierCode & "' " & vbCrLf & _
                                          "            ,'" & PO.LabelCode & "' " & vbCrLf & _
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
        Dim shostname As String = System.Net.Dns.GetHostName
        Dim ls_Remarks As String = "", ls_SupplierType As String = "", ls_xSupplierType As String = ""
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
                    If grid.GetRowValues(i, "SupplierID") <> "" Then
                        ls_Sql = " IF NOT EXISTS (select * from MS_Supplier where SupplierID = '" & grid.GetRowValues(i, "SupplierID") & "')" & vbCrLf & _
                                  " BEGIN" & vbCrLf & _
                                  " INSERT INTO [dbo].[MS_Supplier] " & vbCrLf & _
                                  "            ([SupplierID] " & vbCrLf & _
                                  "            ,[SupplierName] " & vbCrLf & _
                                  "            ,[SupplierType] " & vbCrLf & _
                                  "            ,[SupplierCode] " & vbCrLf & _
                                  "            ,[Address] " & vbCrLf & _
                                  "            ,[City] " & vbCrLf & _
                                  "            ,[PostalCode] " & vbCrLf & _
                                  "            ,[Phone1] " & vbCrLf & _
                                  "            ,[Phone2] " & vbCrLf & _
                                  "            ,[Fax] " & vbCrLf & _
                                  "            ,[NPWP] "

                        ls_Sql = ls_Sql + "            ,[EntryDate] " & vbCrLf & _
                                          "            ,[EntryUser] ) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & grid.GetRowValues(i, "SupplierID") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "SupplierName") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "SupplierType") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "SupplierCode") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "Address") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "City") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "PostalCode") & "' "

                        ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "Phone1") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "Phone2") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "Fax") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "NPWP") & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'UPLOAD') " & vbCrLf & _
                                          " END " & vbCrLf & _
                                          " ELSE " & vbCrLf & _
                                          " BEGIN" & vbCrLf & _
                                          "      UPDATE [dbo].[MS_Supplier] SET " & vbCrLf & _
                                          "       [SupplierName] = '" & grid.GetRowValues(i, "SupplierName") & "' " & vbCrLf & _
                                          "       ,[SupplierType] = '" & grid.GetRowValues(i, "SupplierType") & "' " & vbCrLf & _
                                          "       ,[SupplierCode] = '" & grid.GetRowValues(i, "SupplierCode") & "' " & vbCrLf & _
                                          "       ,[Address] = '" & grid.GetRowValues(i, "Address") & "' " & vbCrLf & _
                                          "       ,[City] = '" & grid.GetRowValues(i, "City") & "' " & vbCrLf & _
                                          "       ,[PostalCode] = '" & grid.GetRowValues(i, "PostalCode") & "' " & vbCrLf & _
                                          "       ,[Phone1] = '" & grid.GetRowValues(i, "Phone1") & "' " & vbCrLf & _
                                          "       ,[Phone2] = '" & grid.GetRowValues(i, "Phone2") & "' " & vbCrLf & _
                                          "       ,[Fax] = '" & grid.GetRowValues(i, "Fax") & "' " & vbCrLf & _
                                          "       ,[NPWP] = '" & grid.GetRowValues(i, "NPWP") & "' " & vbCrLf & _
                                          "      WHERE [SupplierID] = '" & grid.GetRowValues(i, "SupplierID") & "' " & vbCrLf & _
                                          " END"

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "SupplierName")) And Not IsDBNull(grid.GetRowValues(i, "xSupplierName"))) And (grid.GetRowValues(i, "SupplierName").ToString <> "" And grid.GetRowValues(i, "xSupplierName").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "SupplierName").ToString) <> Trim(grid.GetRowValues(i, "xSupplierName").ToString)) Then
                            ls_Remarks = ls_Remarks + "SupplierName " + Trim(grid.GetRowValues(i, "xSupplierName").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "SupplierName").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "SupplierType")) And Not IsDBNull(grid.GetRowValues(i, "xSupplierType"))) And (grid.GetRowValues(i, "SupplierType").ToString <> "" And grid.GetRowValues(i, "xSupplierType").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "SupplierType").ToString) <> Trim(grid.GetRowValues(i, "xSupplierType").ToString)) Then
                            If grid.GetRowValues(i, "xSupplierType").ToString.Trim = "1" Then ls_xSupplierType = "PASI Supplier" Else ls_xSupplierType = "Potential Supplier"
                            If grid.GetRowValues(i, "SupplierType").ToString.Trim = "1" Then ls_SupplierType = "PASI Supplier" Else ls_SupplierType = "Potential Supplier"
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            'ls_Remarks = ls_Remarks + "SupplierType " + Trim(grid.GetRowValues(i, "xSupplierType").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "SupplierType").ToString) & ""
                            ls_Remarks = ls_Remarks + "SupplierType " + ls_xSupplierType & " " & "->" & " " & ls_SupplierType & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "SupplierCode")) And Not IsDBNull(grid.GetRowValues(i, "xSupplierCode"))) And (grid.GetRowValues(i, "SupplierCode").ToString <> "" And grid.GetRowValues(i, "xSupplierCode").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "SupplierCode").ToString) <> Trim(grid.GetRowValues(i, "xSupplierCode").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "SupplierCode " + Trim(grid.GetRowValues(i, "xSupplierCode").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "SupplierCode").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Address")) And Not IsDBNull(grid.GetRowValues(i, "xAddress"))) And (grid.GetRowValues(i, "Address").ToString <> "" And grid.GetRowValues(i, "xAddress").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Address").ToString) <> Trim(grid.GetRowValues(i, "xAddress").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Address " + Trim(grid.GetRowValues(i, "xAddress").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Address").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "City")) And Not IsDBNull(grid.GetRowValues(i, "xCity"))) And (grid.GetRowValues(i, "City").ToString <> "" And grid.GetRowValues(i, "xCity").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "City").ToString) <> Trim(grid.GetRowValues(i, "xCity").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "City " + Trim(grid.GetRowValues(i, "xCity").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "City").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "PostalCode")) And Not IsDBNull(grid.GetRowValues(i, "xPostalCode"))) And (grid.GetRowValues(i, "PostalCode").ToString <> "" And grid.GetRowValues(i, "xPostalCode").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "PostalCode").ToString) <> Trim(grid.GetRowValues(i, "xPostalCode").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PostalCode " + Trim(grid.GetRowValues(i, "xPostalCode").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "PostalCode").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Phone1")) And Not IsDBNull(grid.GetRowValues(i, "xPhone1"))) And (grid.GetRowValues(i, "Phone1").ToString <> "" And grid.GetRowValues(i, "xPhone1").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Phone1").ToString) <> Trim(grid.GetRowValues(i, "xPhone1").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Phone1 " + Trim(grid.GetRowValues(i, "xPhone1").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Phone1").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Phone2")) And Not IsDBNull(grid.GetRowValues(i, "xPhone2"))) And (grid.GetRowValues(i, "Phone2").ToString <> "" And grid.GetRowValues(i, "xPhone2").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Phone2").ToString) <> Trim(grid.GetRowValues(i, "xPhone2").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Phone2 " + Trim(grid.GetRowValues(i, "xPhone2").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Phone2").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "Fax")) And Not IsDBNull(grid.GetRowValues(i, "xFax"))) And (grid.GetRowValues(i, "Fax").ToString <> "" And grid.GetRowValues(i, "xFax").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "Fax").ToString) <> Trim(grid.GetRowValues(i, "xFax").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Fax " + Trim(grid.GetRowValues(i, "xFax").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "Fax").ToString) & ""
                        End If
                    End If

                    If (Not IsDBNull(grid.GetRowValues(i, "NPWP")) And Not IsDBNull(grid.GetRowValues(i, "xNPWP"))) And (grid.GetRowValues(i, "NPWP").ToString <> "" And grid.GetRowValues(i, "xNPWP").ToString <> "") Then
                        If (Trim(grid.GetRowValues(i, "NPWP").ToString) <> Trim(grid.GetRowValues(i, "xNPWP").ToString)) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "NPWP " + Trim(grid.GetRowValues(i, "xNPWP").ToString) & " " & "->" & " " & Trim(grid.GetRowValues(i, "NPWP").ToString) & ""
                        End If
                    End If

                    If ls_Remarks <> "" Then
                        'insert into history
                        ls_Sql = " INSERT INTO MS_History (PCName, MenuID, OperationID, Remarks, RegisterDate, RegisterUserID) " & vbCrLf & _
                                 "VALUES ('" & shostname & "','" & menuID & "','U', 'Update [" & ls_Remarks & "]', " & vbCrLf & _
                                 "GETDATE(), '" & Session("UserID") & "')  "
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                        ls_Remarks = ""
                    End If
                Next i

                '2.3.1 Habis save semua,.. delete tada di tempolary table
                ls_Sql = "delete UploadSupplier "

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