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

Public Class ReceivingUpload
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "F04"
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
            Dim fi As New FileInfo(Server.MapPath("~\Template\Template Upload Receiving.xlsx"))
            If Not fi.Exists Then
                lblInfo.Text = "[9999] Excel Template Not Found !"
                ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Template\Template Upload Receiving.xlsx")
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        End Try

    End Sub

    Private Sub grid_HtmlDataCellPrepared(sender As Object, e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")

        If e.GetValue("ErrorCls") = "" Then
        Else
            If InStr(1, e.GetValue("ErrorCls").ToString.Trim, "Warning") = 0 Then
                e.Cell.BackColor = Color.Red
            Else
                e.Cell.BackColor = Color.Yellow
            End If

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
                Case "saveReplace"

            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

#End Region

#Region "PROCEDURE"

    Private Sub bindData(ByVal pPO As String, ByVal pAff As String)
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                      " 	row_number() over (order by a.SupplierID, a.SuratJalanNo, a.PONo, a.KanbanNo, a.PartNo asc) as NoUrut, " & vbCrLf & _
                      " 	a.SuratJalanNo, " & vbCrLf & _
                      " 	a.AffiliateID, " & vbCrLf & _
                      " 	a.SupplierID, " & vbCrLf & _
                      " 	a.PONo, " & vbCrLf & _
                      " 	a.KanbanNo, " & vbCrLf & _
                      " 	a.PartNo, " & vbCrLf & _
                      " 	b.PartName, " & vbCrLf & _
                      " 	a.RecQty, " & vbCrLf & _
                      " 	a.DefectQty, "

            ls_SQL = ls_SQL + " 	a.ErrorCls " & vbCrLf & _
                              " from UploadReceiving a " & vbCrLf & _
                              " left join MS_Parts b on a.partno = b.partno " & vbCrLf & _
                              " where a.SuratJalanNo in (" & pPO & ") and a.AffiliateID = '" & pAff & "'"
            '" where a.SuratJalanNo in (" & Session("PONoUpload") & ") and a.AffiliateID = '" & Session("AffiliateID") & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()
            'Session.Remove("PONoUpload")
            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, '' SuratJalanNo, '' AffiliateID, '' SupplierID, '' PONo, " & vbCrLf & _
                     " '' KanbanNo, '' PartNo, '' PartName, " & vbCrLf & _
                     " 0 RecQty, 0 DefectQty, '' ErrorCls"

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

    Private Sub up_Import()
        Dim dt As New System.Data.DataTable
        Dim dtDetail As New System.Data.DataTable
        Dim ls_sql As String = ""
        Session.Remove("SuratJalanNoAda")
        Session.Remove("PONoUpload")
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

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A1:G2]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dt)

                        If dt.Rows.Count > 0 Then
                            '1.Surat Jalan No
                            If IsDBNull(dt.Rows(0).Item(0)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            '2.Affiliate
                            If IsDBNull(dt.Rows(0).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            '3.PONO
                            If IsDBNull(dt.Rows(0).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            '4.Kanban
                            If IsDBNull(dt.Rows(0).Item(3)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            '5.PartNo
                            If IsDBNull(dt.Rows(0).Item(4)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            '6.GR
                            If IsDBNull(dt.Rows(0).Item(5)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            '7.DF
                            If IsDBNull(dt.Rows(0).Item(6)) Then
                                lblInfo.Text = "[9999] Invalid Template Upload Receiving, please check the file again with Original Template Upload Receiving!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'Check ada data yang perlu diupload atau tidak
                            '8.Surat Jalan No
                            If IsDBNull(dt.Rows(1).Item(0)) Then
                                lblInfo.Text = "[9998] No data you want to Upload!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                        End If

                        Dim dtUploadReceivingList As New List(Of clsUploadReceiving)

                        'Get Detail Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A2:G65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsUploadReceiving

                                    '1. Check Surat Jalan No
                                    If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                        dtUploadDetail.SuratJalanNo = dtDetail.Rows(i).Item(0).ToString().Trim()
                                    Else
                                        dtUploadDetail.SuratJalanNo = ""
                                    End If

                                    '2. Check AffiliateID
                                    If IsDBNull(dtDetail.Rows(i).Item(1)) = False Then
                                        dtUploadDetail.AffiliateID = dtDetail.Rows(i).Item(1).ToString().Trim()
                                    Else
                                        dtUploadDetail.AffiliateID = ""
                                    End If

                                    '3. Check PONo
                                    If IsDBNull(dtDetail.Rows(i).Item(2)) = False Then
                                        dtUploadDetail.PONo = dtDetail.Rows(i).Item(2).ToString().Trim()
                                    Else
                                        dtUploadDetail.PONo = ""
                                    End If

                                    '4. Check KanbanNo
                                    If IsDBNull(dtDetail.Rows(i).Item(3)) = False Then
                                        dtUploadDetail.KanbanNo = dtDetail.Rows(i).Item(3).ToString().Trim()
                                    Else
                                        dtUploadDetail.KanbanNo = ""
                                    End If

                                    '5. Check PartNo
                                    If IsDBNull(dtDetail.Rows(i).Item(4)) = False Then
                                        dtUploadDetail.PartNo = dtDetail.Rows(i).Item(4).ToString().Trim()
                                    Else
                                        dtUploadDetail.PartNo = ""
                                    End If

                                    '6. Check GR
                                    If IsDBNull(dtDetail.Rows(i).Item(5)) = False Then
                                        dtUploadDetail.GoodReceiveQty = dtDetail.Rows(i).Item(5)
                                    Else
                                        dtUploadDetail.GoodReceiveQty = 0
                                    End If

                                    '7. Check DF
                                    If IsDBNull(dtDetail.Rows(i).Item(6)) = False Then
                                        dtUploadDetail.DefectQty = dtDetail.Rows(i).Item(6)
                                    Else
                                        dtUploadDetail.DefectQty = 0
                                    End If

                                    dtUploadReceivingList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            '01.01 Delete TempoaryData
                            ls_sql = "delete UploadReceiving where AffiliateID = '" & Session("AffiliateID") & "'"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()

                            Dim checkKosong As Boolean = False

                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadReceivingList.Count - 1
                                Dim ls_SupplierID As String
                                checkKosong = False

                                Dim ls_error As String = ""
                                Dim PO As clsUploadReceiving = dtUploadReceivingList(i)

                                '02.0.1 Check Kosong Affiliate dan Master
                                If PO.AffiliateID = "" Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "Affiliate Code is blank, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", Affiliate Code is blank, please check a file again!"
                                    End If
                                End If

                                '02.0.2
                                ls_sql = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID = '" & PO.AffiliateID & "'"
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)
                                If ds3.Tables(0).Rows.Count = 0 Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "Affiliate Code not found in Master Affiliate, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", Affiliate Code not found in Master Affiliate, please check a file again!"
                                    End If
                                End If

                                '02.0.3 Check Kosong PONo
                                If PO.PONo = "" Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "PO No. is blank, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", PO No. is blank, please check a file again!"
                                    End If
                                End If

                                '02.0.4 Check Kosong Kanban No
                                If PO.KanbanNo = "" Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "Kanban No. is blank, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", Kanban No. is blank, please check a file again!"
                                    End If
                                End If

                                '02.0.5 Check Kosong Kanban No
                                If PO.PartNo = "" Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "Part No. is blank, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", Part No. is blank, please check a file again!"
                                    End If
                                End If

                                '02.0.6 Check Kosong GR Qty
                                If IsNumeric(PO.GoodReceiveQty) = False Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "GoodReceiveQty must be numeric, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", GoodReceiveQty must be numeric, please check a file again!"
                                    End If
                                End If

                                '02.0.7 Check Kosong DF Qty
                                If IsNumeric(PO.DefectQty) = False Then
                                    checkKosong = True
                                    If ls_error = "" Then
                                        ls_error = "DefectQty must be numeric, please check a file again!"
                                    Else
                                        ls_error = ls_error + ", DefectQty must be numeric, please check a file again!"
                                    End If
                                End If

                                If checkKosong = False Then
                                    '02.1 Check Kanban No.
                                    'ls_sql = "select * from DOPASI_Detail " & vbCrLf & _
                                    '         "where SuratJalanNo = '" & PO.SuratJalanNo & "' and AffiliateID = '" & PO.AffiliateID & "' " & vbCrLf & _
                                    '         "      and KanbanNo = '" & PO.KanbanNo & "' and PONo = '" & PO.PONo & "' and PartNo = '" & PO.PartNo & "'"
                                    ls_sql = "select * from DOPASI_Detail " & vbCrLf & _
                                             "where SuratJalanNo = '" & PO.SuratJalanNo & "' and AffiliateID = '" & PO.AffiliateID & "' " & vbCrLf & _
                                             "      and KanbanNo = '" & PO.KanbanNo & "' and PONo = '" & PO.PONo & "' "
                                    Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                    Dim ds2 As New DataSet
                                    sqlDA2.Fill(ds2)

                                    If ds2.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "This Kanban No. not found in DN PASI, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", This Kanban No. not found in DN PASI, please check again with PASI!"
                                        End If
                                        GoTo step99
                                    End If

                                    '02.2 Check PartNo
                                    ls_sql = "select SupplierID, SUM(DOQty) DOQty from DOPASI_Detail " & vbCrLf & _
                                             "where SuratJalanNo = '" & PO.SuratJalanNo & "' and AffiliateID = '" & PO.AffiliateID & "' " & vbCrLf & _
                                             "      and KanbanNo = '" & PO.KanbanNo & "' and PONo = '" & PO.PONo & "' and PartNo = '" & PO.PartNo & "' " & vbCrLf & _
                                             "group by SupplierID "
                                    Dim sqlCmd33 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA33 As New SqlDataAdapter(sqlCmd33)
                                    Dim ds33 As New DataSet
                                    sqlDA33.Fill(ds33)

                                    If ds33.Tables(0).Rows.Count = 0 Then
                                        If ls_error = "" Then
                                            ls_error = "This Part No. not found in DN PASI, please check again with PASI!"
                                        Else
                                            ls_error = ls_error + ", This Part No. not found in DN PASI, please check again with PASI!"
                                        End If
                                        GoTo step99
                                    Else
                                        Dim ls_DOPASI As Double = CDbl(ds33.Tables(0).Rows(0)("DOQty"))
                                        Dim ls_GRAffiliate As Double = CDbl(PO.GoodReceiveQty)
                                        Dim ls_DFAffiliate As Double = CDbl(PO.DefectQty)

                                        ls_SupplierID = ds33.Tables(0).Rows(0)("SupplierID")

                                        '02.3 Check MOQ
                                        ls_sql = "select ISNULL(a.POQtyBox,b.QtyBox) QtyBox from PO_Detail a Left Join" & vbCrLf & _
                                                 "MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                                                 "where a.SupplierID = '" & ls_SupplierID & "' and a.AffiliateID = '" & PO.AffiliateID & "' " & vbCrLf & _
                                                 "     and a.PartNo = '" & PO.PartNo & "' and a.PoNo = '" & PO.PONo & "' "
                                        Dim sqlCmd66 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                        Dim sqlDA66 As New SqlDataAdapter(sqlCmd66)
                                        Dim ds66 As New DataSet
                                        sqlDA66.Fill(ds66)

                                        If ds66.Tables(0).Rows.Count > 0 Then
                                            Dim ls_QtyBox As Double = CDbl(ds66.Tables(0).Rows(0)("QtyBox"))
                                            If ls_GRAffiliate Mod ls_QtyBox <> 0 Then
                                                If ls_error = "" Then
                                                    ls_error = "This QTY must be multiple from Box/Qty (" & ls_QtyBox & "), please check again with PASI!"
                                                Else
                                                    ls_error = ls_error + ", This QTY must be multiple from Qty/Box (" & ls_QtyBox & "), please check again with PASI!"
                                                End If
                                                GoTo step99
                                            End If
                                            If ls_DFAffiliate Mod ls_QtyBox <> 0 Then
                                                If ls_error = "" Then
                                                    ls_error = "This QTY must be multiple from Box/Qty (" & ls_QtyBox & "), please check again with PASI!"
                                                Else
                                                    ls_error = ls_error + ", This QTY must be multiple from Qty/Box (" & ls_QtyBox & "), please check again with PASI!"
                                                End If
                                                GoTo step99
                                            End If
                                        Else
                                            If ls_error = "" Then
                                                ls_error = "This QTY must be multiple from Box/Qty, please check again with PASI!"
                                            Else
                                                ls_error = ls_error + ", This QTY must be multiple from Qty/Box, please check again with PASI!"
                                            End If
                                            GoTo step99
                                        End If

                                        '02.1.1 Check GR + DF < DO
                                        If (ls_GRAffiliate + ls_DFAffiliate) <> ls_DOPASI Then
                                            If ls_error = "" Then
                                                ls_error = "Warning!!! Good Receiving Qty + Defect Qty not same with Delivery Qty in DN PASI!"
                                            Else
                                                ls_error = ls_error + ", Warning!!! Good Receiving Qty + Defect Qty not same with Delivery Qty in DN PASI!"
                                            End If
                                        End If
                                    End If

                                    'If ds2.Tables(0).Rows.Count = 0 Then
                                    '    If ls_error = "" Then
                                    '        ls_error = "This Receiving not found in DN PASI, please check again with PASI!"
                                    '    Else
                                    '        ls_error = ls_error + ", This Receiving not found in DN PASI, please check again with PASI!"
                                    '    End If
                                    'Else
                                    '    Dim ls_DOPASI As Double = CDbl(ds2.Tables(0).Rows(0)("DOQty"))
                                    '    Dim ls_GRAffiliate As Double = CDbl(PO.GoodReceiveQty)
                                    '    Dim ls_DFAffiliate As Double = CDbl(PO.DefectQty)

                                    '    ls_SupplierID = ds2.Tables(0).Rows(0)("SupplierID")

                                    '    '02.1.1 Check GR + DF < DO
                                    '    If (ls_GRAffiliate + ls_DFAffiliate) <> ls_DOPASI Then
                                    '        If ls_error = "" Then
                                    '            ls_error = "Good Receiving Qty + Defect Qty not same with Delivery Qty in DN PASI, please check again with PASI!"
                                    '        Else
                                    '            ls_error = ls_error + ", Good Receiving Qty + Defect Qty not same with Delivery Qty in DN PASI, please check again with PASI!"
                                    '        End If
                                    '    End If
                                    'End If

                                    '02.2 Check DATA SURAT JALAN NO RECEIVING
                                    ls_sql = "select * from ReceiveAffiliate_Master " & vbCrLf & _
                                             "where SuratJalanNo = '" & PO.SuratJalanNo & "' and AffiliateID = '" & PO.AffiliateID & "' "
                                    Dim sqlCmd11 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    Dim sqlDA11 As New SqlDataAdapter(sqlCmd11)
                                    Dim ds11 As New DataSet
                                    sqlDA11.Fill(ds11)

                                    If ds11.Tables(0).Rows.Count > 0 Then
                                        Session("SuratJalanNoAda") = "ada"
                                    End If
                                End If

step99:
                                '02.3 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.UploadReceiving WHERE PartNo = '" & PO.PartNo & "' and AffiliateID = '" & PO.AffiliateID & "'" & vbCrLf & _
                                         "          and PONo = '" & PO.PONo & "' and SupplierID = '" & ls_SupplierID & "' and KanbanNo = '" & PO.KanbanNo & "'"
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                If ds4.Tables(0).Rows.Count > 0 Then
                                    ls_sql = "delete UploadReceiving WHERE PartNo = '" & PO.PartNo & "' and AffiliateID = '" & PO.AffiliateID & "'" & vbCrLf & _
                                         "          and PONo = '" & PO.PONo & "' and SupplierID = '" & ls_SupplierID & "' and KanbanNo = '" & PO.KanbanNo & "'"
                                    Dim sqlComm1 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm1.ExecuteNonQuery()
                                    sqlComm1.Dispose()
                                End If

                                ls_sql = " INSERT INTO [dbo].[UploadReceiving] " & vbCrLf & _
                                          "            ([SuratJalanNo],[SupplierID],[AffiliateID],[PONo],[KanbanNo],[PartNo],[RecQty],[DefectQty],[ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & PO.SuratJalanNo & "' " & vbCrLf & _
                                          "            ,'" & ls_SupplierID & "' " & vbCrLf & _
                                          "            ,'" & PO.AffiliateID & "' " & vbCrLf

                                ls_sql = ls_sql + "            ,'" & PO.PONo & "' " & vbCrLf & _
                                                  "            ,'" & PO.KanbanNo & "' " & vbCrLf & _
                                                  "            ,'" & PO.PartNo & "' " & vbCrLf & _
                                                  "            ," & PO.GoodReceiveQty & " " & vbCrLf & _
                                                  "            ," & PO.DefectQty & " " & vbCrLf & _                                                  
                                                  "            ,'" & ls_error & "') " & vbCrLf
                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()

                                If InStr(1, Session("PONoUpload"), PO.SuratJalanNo.Trim) = 0 Then
                                    If Session("PONoUpload") = "" Then
                                        Session("PONoUpload") = "'" & PO.SuratJalanNo.Trim & "'"
                                    Else
                                        Session("PONoUpload") = Session("PONoUpload") & ",'" & PO.SuratJalanNo.Trim & "'"
                                    End If
                                End If

                                Session("AffiliateUpload") = PO.AffiliateID

                            Next
                            sqlTran.Commit()

                            lblInfo.Text = "[7001] Data Checking Done!"
                            lblInfo.ForeColor = Color.Blue
                            grid.JSProperties("cpMessage") = lblInfo.Text

                            Call bindData(Session("PONoUpload"), Session("AffiliateUpload"))
                        End Using
                    Catch ex As Exception
                        lblInfo.Text = ex.Message
                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                        Exit Sub
                    End Try
                    dt.Reset()
                    dtDetail.Reset()
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
        Finally

        End Try
    End Sub

    Private Sub up_Save()
        Dim i As Integer
        Dim ls_Check As Boolean = False
        Dim ls_Sql As String
        Dim ls_MsgID As String = ""


        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "ErrorCls").ToString.Trim <> "" Then
                    If InStr(1, grid.GetRowValues(i, "ErrorCls").ToString.Trim, "Warning") = 0 Then
                        ls_Check = True
                        Exit For
                    End If                    
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
                Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                SQLCom.Connection = SqlCon
                SQLCom.Transaction = SqlTran

                ls_Sql = "delete ReceiveAffiliate_Detail where SuratJalanNo in (" & Session("PONoUpload") & ") and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                ls_Sql = "delete ReceiveAffiliate_Master where SuratJalanNo in (" & Session("PONoUpload") & ") and AffiliateID = '" & Session("AffiliateID") & "'"
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                For i = 0 To grid.VisibleRowCount - 1
                    ls_Sql = "select * from ReceiveAffiliate_Detail where SuratJalanNo = '" & grid.GetRowValues(i, "SuratJalanNo").ToString & "'" & vbCrLf & _
                              "        and AffiliateID = '" & grid.GetRowValues(i, "AffiliateID").ToString & "' " & vbCrLf & _
                              "        and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString & "' " & vbCrLf & _
                              "        and KanbanNo = '" & grid.GetRowValues(i, "KanbanNo").ToString & "' " & vbCrLf & _
                              "        and PONo = '" & grid.GetRowValues(i, "PONo").ToString & "' " & vbCrLf & _
                              "        and PartNo = '" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf

                    SQLCom.CommandText = ls_Sql
                    Dim da10 As New SqlDataAdapter(SQLCom)
                    Dim ds10 As New DataSet
                    da10.Fill(ds10)

                    If ds10.Tables(0).Rows.Count = 0 Then
                        ls_Sql = " INSERT INTO [dbo].[ReceiveAffiliate_Detail] " & vbCrLf & _
                                  "            ([SuratJalanNo] " & vbCrLf & _
                                  "            ,[AffiliateID] " & vbCrLf & _
                                  "            ,[SupplierID] " & vbCrLf & _
                                  "            ,[PONo] " & vbCrLf & _
                                  "            ,[KanbanNo] " & vbCrLf & _
                                  "            ,[PartNo] " & vbCrLf & _
                                  "            ,[RecQty] " & vbCrLf & _
                                  "            ,[DefectQty] ) " & vbCrLf

                        ls_Sql = ls_Sql + "      VALUES " & vbCrLf & _
                                        "            ('" & grid.GetRowValues(i, "SuratJalanNo").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "AffiliateID").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "SupplierID").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "PONo").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "KanbanNo").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "PartNo").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "RecQty").ToString & "' " & vbCrLf & _
                                        "            ,'" & grid.GetRowValues(i, "DefectQty").ToString & "' ) "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()                        
                    End If

                    ls_Sql = "select * from ReceiveAffiliate_Master where SuratJalanNo = '" & grid.GetRowValues(i, "SuratJalanNo").ToString & "'" & vbCrLf & _
                              "        and AffiliateID = '" & grid.GetRowValues(i, "AffiliateID").ToString & "' " & vbCrLf & _
                              "        and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString & "' "

                    SQLCom.CommandText = ls_Sql
                    Dim da7 As New SqlDataAdapter(SQLCom)
                    Dim ds7 As New DataSet
                    da7.Fill(ds7)

                    If ds7.Tables(0).Rows.Count = 0 Then
                        ls_Sql = " insert into ReceiveAffiliate_Master " & vbCrLf & _
                                  " select SuratJalanNo, AffiliateID, SupplierID, 0, GETDATE(), '" & Session("AffiliateID") & "', JenisArmada, " & vbCrLf & _
                                  " DriverName, DriverContact, NoPol, " & vbCrLf & _
                                  " TotalBox =  (select sum(RecQty / ISNULL(POD.POQtyBox,b.QtyBox))  " & vbCrLf & _
                                  "             from ReceiveAffiliate_Detail a  " & vbCrLf & _
                                  "             LEFT JOIN PO_Detail POD ON a.PONo = POD.PONo And a.AffiliateID = POD.AffiliateID And a.SupplierID = POD.SupplierID And a.PartNo = POD.PartNo  " & vbCrLf & _
                                  "             left join MS_PartMapping b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo  " & vbCrLf & _
                                  "             where a.SuratJalanNo = '" & grid.GetRowValues(i, "SuratJalanNo").ToString & "' and a.AffiliateID = '" & grid.GetRowValues(i, "AffiliateID").ToString & "' and a.SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString & "')  " & vbCrLf & _
                                  " ,GETDATE(), '" & Session("AffiliateID") & "', GETDATE(),'" & Session("AffiliateID") & "', 0, 'D01' from DOPASI_Master " & vbCrLf & _
                                  " where SuratJalanNo = '" & grid.GetRowValues(i, "SuratJalanNo").ToString & "'" & vbCrLf & _
                                  "        and AffiliateID = '" & grid.GetRowValues(i, "AffiliateID").ToString & "' " & vbCrLf & _
                                  "        and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString & "' "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    Else
                        ls_Sql = " Update ReceiveAffiliate_Master set " & vbCrLf & _
                                  " TotalBox =  (select sum(RecQty / ISNULL(POD.POQtyBox,b.QtyBox))  " & vbCrLf & _
                                  "             from ReceiveAffiliate_Detail a  " & vbCrLf & _
                                  "             LEFT JOIN PO_Detail POD ON a.PONo = POD.PONo And a.AffiliateID = POD.AffiliateID And a.SupplierID = POD.SupplierID And a.PartNo = POD.PartNo  " & vbCrLf & _
                                  "             left join MS_PartMapping b on a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID and a.PartNo = b.PartNo  " & vbCrLf & _
                                  "             where a.SuratJalanNo = '" & grid.GetRowValues(i, "SuratJalanNo").ToString & "' and a.AffiliateID = '" & grid.GetRowValues(i, "AffiliateID").ToString & "' and a.SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString & "')  " & vbCrLf & _
                                  " where SuratJalanNo = '" & grid.GetRowValues(i, "SuratJalanNo").ToString & "'" & vbCrLf & _
                                  "        and AffiliateID = '" & grid.GetRowValues(i, "AffiliateID").ToString & "' " & vbCrLf & _
                                  "        and SupplierID = '" & grid.GetRowValues(i, "SupplierID").ToString & "' "

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    End If
                    ls_MsgID = "1001"
                Next

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