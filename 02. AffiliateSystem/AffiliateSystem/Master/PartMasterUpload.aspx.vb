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

Public Class PartMasterUpload
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A01"
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
            End If
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Response.Redirect("~/Master/PartMasterSetting.aspx")
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

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If Trim(e.GetValue("remarks")) = "" Then

        Else
            e.Cell.BackColor = Color.Red
        End If
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)
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
            grid.JSProperties("cpMessage") = lblInfo.Text
            Session("YA010IsSubmit") = lblInfo.Text
        End Try
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim id As String = Session("AffiliateID").ToString

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT ROW_NUMBER() OVER (ORDER BY UPL.PartNo) NoUrut, UPL.PartNo, MP.PartName, CASE WHEN UPL.ShowCls = '1' THEN 'YES' ELSE 'NO' END ShowCls, UPL.MOQ, UPL.LocationID, UPL.remarks " & vbCrLf & _
                "FROM UploadPartLocation UPL " & vbCrLf & _
                "LEFT JOIN MS_Parts MP ON UPL.PartNo = MP.PartNo " & vbCrLf & _
                "WHERE UPL.AffiliateID = '" & Session("AffiliateID") & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet

            sqlDA.Fill(ds)

            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False)
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' PartNo, '' PartName, '' ShowCls, '' MOQ, '' LocationID, '' remarks"

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
        Dim ls_sql As String = ""

        Dim connStr As String = ""
        Dim pAffiliateID As String = Session("AffiliateID")

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

            For Each drSheet In dtSheets.Rows
                If InStr("_xlnm#_FilterDatabase", drSheet("TABLE_NAME").ToString(), CompareMethod.Text) = 0 Then
                    If InStr("_xlnm#Print_Titles", drSheet("TABLE_NAME").ToString(), CompareMethod.Text) = 0 Then
                        listSheet.Add(drSheet("TABLE_NAME").ToString())
                    End If
                End If
            Next

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Dim pTableCode As String = listSheet(0)

                Try
                    'Get Data
                    MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A2:F65536]")
                    MyAdapter.SelectCommand = MyCommand
                    MyAdapter.Fill(dt)

                    Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPartLocation")

                        'Delete TemporaryData
                        ls_sql = "delete UploadPartLocation where AffiliateID = '" & pAffiliateID & "'"

                        Dim sqlCmd = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        sqlCmd.ExecuteNonQuery()
                        sqlCmd.Dispose()

                        Dim strPartNo As String = ""
                        Dim strPartName As String = ""
                        Dim strMOQ As String = ""
                        Dim strShow As String = ""
                        Dim strLocation As String = ""
                        Dim strRemarks As String = ""

                        If dt.Rows.Count > 0 Then
                            For i = 1 To dt.Rows.Count - 1
                                If CStr(IIf(IsDBNull(dt.Rows(i).Item(0)), "", dt.Rows(i).Item(0))) <> "" Then
                                    strPartNo = CStr(IIf(IsDBNull(dt.Rows(i).Item(1)), "", dt.Rows(i).Item(1)))
                                    strPartName = CStr(IIf(IsDBNull(dt.Rows(i).Item(2)), "", dt.Rows(i).Item(2)))
                                    strMOQ = CStr(IIf(IsDBNull(dt.Rows(i).Item(3)), "", dt.Rows(i).Item(3)))
                                    strShow = CStr(IIf(IsDBNull(dt.Rows(i).Item(4)), "", dt.Rows(i).Item(4)))
                                    strLocation = CStr(IIf(IsDBNull(dt.Rows(i).Item(5)), "", dt.Rows(i).Item(5)))
                                    strRemarks = ""

                                    ls_sql = "SELECT * FROM MS_PartMapping WHERE PartNo = '" & strPartNo & "' AND AffiliateID = '" & pAffiliateID & "'"

                                    Dim sqlCmd1 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA As New SqlDataAdapter(sqlCmd1)
                                    Dim ds As New DataSet
                                    sqlDA.Fill(ds)

                                    If ds.Tables(0).Rows.Count = 0 Then
                                        strRemarks = "Part No. not exists !"
                                    End If

                                    ls_sql = "INSERT INTO UploadPartLocation(AffiliateID, PartNo, MOQ, ShowCls, LocationID, Remarks) " & vbCrLf & _
                                        "VALUES('" & pAffiliateID & "', '" & strPartNo & "', '" & strMOQ & "', '" & IIf(UCase(strShow) = "YES", "1", "0") & "', '" & strLocation & "', '" & strRemarks & "')"

                                    Dim sqlCmd2 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlCmd2.ExecuteNonQuery()
                                    sqlCmd2.Dispose()
                                End If
                            Next
                        End If

                        sqlTran.Commit()

                        lblInfo.Text = "[7001] Data Checking Done!"
                        lblInfo.ForeColor = Color.Blue
                        grid.JSProperties("cpMessage") = lblInfo.Text

                        Call bindData()
                    End Using
                Catch ex As Exception
                    MyConnection.Close()
                    lblInfo.Text = ex.Message
                    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                    If InStr(ex.Message, "Cannot find column") = 1 Then
                        lblInfo.Text = "Format Template Tidak Sesuai"
                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, lblInfo.Text.ToString())
                    End If
                Finally
                    MyConnection.Close()
                    sqlConn.Close()
                End Try
                dt.Reset()
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
        Dim ls_Check As Boolean
        Dim ls_Sql As String = ""
        Dim ls_MsgID As String = ""
        Dim ls_Affiliate As String = Session("AffiliateID").ToString

        Dim SqlCon As New SqlConnection(clsGlobal.ConnectionString)
        Dim SqlTran As SqlTransaction

        SqlCon.Open()
        SqlTran = SqlCon.BeginTransaction

        Try
            If grid.VisibleRowCount = 0 Then Exit Sub

            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "remarks").ToString.Trim <> "" Then
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

            Dim SQLCom As SqlCommand = SqlCon.CreateCommand
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran

            For i = 0 To grid.VisibleRowCount - 1
                ls_Sql = "UPDATE MS_PartMapping SET LocationID = '" & grid.GetRowValues(i, "LocationID").ToString.Trim & "', " & vbCrLf & _
                    "UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID") & "' " & vbCrLf & _
                    "WHERE AffiliateID = '" & ls_Affiliate & "' " & vbCrLf & _
                    "AND PartNo = '" & grid.GetRowValues(i, "PartNo").ToString.Trim & "' "

                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()
            Next

            ls_Sql = "Delete UploadPartLocation Where AffiliateID = '" & ls_Affiliate & "'"

            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            SqlTran.Commit()
            ls_MsgID = "1001"

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