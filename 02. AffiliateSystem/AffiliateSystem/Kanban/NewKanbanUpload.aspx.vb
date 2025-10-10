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

Public Class NewKanbanUpload
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
                'btnDownload.Enabled = False
            End If
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If

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
                Case "send"
                    Call up_ApproveData()
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

    Private Sub bindData(ByVal kanban As String)
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select  " & vbCrLf & _
                  " 	row_number() over (order by a.SupplierID, a.PartNo asc) as no, " & vbCrLf & _
                  " 	kanbanno, a.PartNo as partno, b.PartName as partname, a.Cycle1 as qty, a.remarks as remarks, a.supplierid as supplier, Convert(char(11),convert(date,deliverydate),120) as deliverydate, cycle = kanbancycle, direct, uom = unitcls, ISNULL(d.MOQ,0) MOQ " & vbCrLf

            ls_SQL = ls_SQL + " from UploadKanban a  " & vbCrLf & _
                              " left join MS_Parts b on a.PartNo = b.PartNo " & vbCrLf & _
                              " left join MS_PartMapping d on a.AffiliateID = d.AffiliateID and a.SupplierID = d.SupplierID and a.PartNo = d.PartNo " & vbCrLf & _
                              " where a.kanbanno IN (" & kanban & ") and a.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
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

            ls_SQL = " select top 0 '' no, '' kanbanno, '' PartNo, '' PartName, '' MOQ, '' qty, '' Remarks "

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
        Dim ls_TempKanban As String = ""
        Dim pDeleteSupplier As String = ""
        Dim pDeleteKanban As String = ""

        Dim connStr As String = ""


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


                        '    Session.Remove("KanbanTime1")
                        '    Session.Remove("KanbanTime2")
                        '    Session.Remove("KanbanTime3")
                        '    Session.Remove("KanbanTime4")

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


                        'End If

                        Dim dtUploadHeader As New clsKanbanHeader
                        Dim dtUploadHeaderList As New List(Of clsKanbanHeader)

                        'Dim dtUploadDetail As New clsPODetail
                        Dim dtUploadDetailList As New List(Of clsKanbanDetail)


                        ''Get Header Data
                        'MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B6:B9]")
                        'MyAdapter.SelectCommand = MyCommand
                        'MyAdapter.Fill(dtHeader)

                        'If dtHeader.Rows.Count > 0 Then
                        '    dtUploadHeader.H_AffiliateID = Session("AffiliateID")
                        '    dtUploadHeader.H_Vendor = dtHeader.Rows(0).Item(0)
                        '    dtUploadHeader.H_DeliveryDate = Microsoft.VisualBasic.Left(dt.Rows(5).Item(5), 8) 'dtHeader.Rows(2).Item(0)
                        '    dtUploadHeader.H_ShipBy = dtHeader.Rows(3).Item(0)
                        '    dtUploadHeader.H_kanbanNo = dt.Rows(5).Item(5)
                        '    dtUploadHeader.H_Cycle = Microsoft.VisualBasic.Mid(dt.Rows(5).Item(5), 10, 2)
                        'End If

                        'Get Detail Data
                        dtUploadHeader.H_AffiliateID = Session("AffiliateID")
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A2:G65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If CStr(IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))) <> "" Then
                                    Dim dtUploadDetail As New clsKanbanDetail

                                    'GET SUPPLIER
                                    ls_sql = "SELECT * FROM MS_PartMapping WHERE AffiliateID = '" & Session("AffiliateID") & "' AND PartNo = '" & dtDetail.Rows(i).Item(3) & "' "

                                    Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn)
                                    Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                    Dim ds2 As New DataSet
                                    sqlDA2.Fill(ds2)
                                    'GET SUPPLIER
                                    If ds2.Tables(0).Rows.Count > 0 Then
                                        dtUploadDetail.D_supp = ds2.Tables(0).Rows(0)("supplierID")
                                        If pDeleteSupplier = "" Then
                                            pDeleteSupplier = "'" + Trim(dtUploadDetail.D_supp) + "'"
                                        Else
                                            pDeleteSupplier = pDeleteSupplier + ",'" + Trim(dtUploadDetail.D_supp) + "'"
                                        End If
                                    Else
                                        pDeleteSupplier = "''"
                                    End If
                                    dtUploadDetail.D_Kanbanno = dtDetail.Rows(i).Item(1)
                                    If pDeleteKanban = "" Then
                                        pDeleteKanban = "'" + Trim(dtUploadDetail.D_Kanbanno) + "'"
                                    Else
                                        pDeleteKanban = pDeleteKanban + ",'" + Trim(dtUploadDetail.D_Kanbanno) + "'"
                                    End If

                                    dtUploadDetail.D_DeliveryDate = Left(dtUploadHeader.H_kanbanNo, 8)
                                    dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(3)
                                    dtUploadDetail.D_c1 = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), 0, dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.D_Direct = dtDetail.Rows(i).Item(4)
                                    dtUploadDetail.D_DeliveryDate = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.D_c2 = dtDetail.Rows(i).Item(2) 'cycle

                                    If UCase(Trim(dtDetail.Rows(i).Item(4))) = "DIRECT" Then dtUploadDetail.D_Direct = "0"
                                    If UCase(Trim(dtDetail.Rows(i).Item(4))) <> "DIRECT" Then dtUploadDetail.D_Direct = "1"

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

                            ls_sql = "delete UploadKanban where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'" & vbCrLf & _
                                     " and KanbanNO IN (" & pDeleteKanban & ") " & vbCrLf & _
                                     " --and (SupplierID IN (" & pDeleteSupplier & ") or SupplierID = '') " & vbCrLf

                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim Kanban As clsKanbanDetail = dtUploadDetailList(i)
                                Dim ls_Qty As Integer

                                '01. Check Kanban already Exists
                                ls_sql = "SELECT * FROM Kanban_Master WHERE KanbaNNo = '" & Kanban.D_Kanbanno & "' " & vbCrLf & _
                                         " and AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                                         " and SupplierID = '" & Kanban.D_supp & "' "

                                Dim sqlCm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlD As New SqlDataAdapter(sqlCm)
                                Dim dsq As New DataSet
                                sqlD.Fill(dsq)

                                If dsq.Tables(0).Rows.Count > 0 Then
                                    ls_error = "This PO No. Already exist!"
                                End If

                                ''01. Check Kanban already Exists 202010
                                'ls_sql = "SELECT * FROM Kanban_Master WHERE KanbaNNo = '" & Kanban.D_Kanbanno & "' " & vbCrLf & _
                                '         " and AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                                '         " and SupplierID = '" & Kanban.D_supp & "' " & vbCrLf & _
                                '         " and isnull(AffiliateApproveUser,'') <> '' "

                                'Dim sqlCm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                'Dim sqlD As New SqlDataAdapter(sqlCm)
                                'Dim dsq As New DataSet
                                'sqlD.Fill(dsq)

                                'If dsq.Tables(0).Rows.Count > 0 Then
                                '    If Not IsDBNull(dsq.Tables(0).Rows(0)("KanbanStatus")) Then
                                '        Call clsMsg.DisplayMessage(lblInfo, "5012", clsMessage.MsgType.ErrorMessage)
                                '        Exit Sub
                                '    End If
                                'End If

                                ''01.01 Delete TempoaryData

                                'If OLDKanban = "" And OLDSupp = "" Then
                                '    ls_sql = "delete UploadKanban where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'" & vbCrLf & _
                                '             " and KanbanNO = '" & Kanban.D_Kanbanno & "'" & vbCrLf & _
                                '             " and (SupplierID = '" & Trim(Kanban.D_supp) & "' or SupplierID = '') " & vbCrLf & _
                                '             " --and PartNo = '" & Kanban.D_PartNo & "' " & vbCrLf

                                '    Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                '    sqlComm9.ExecuteNonQuery()
                                '    sqlComm9.Dispose()
                                'ElseIf Trim(OLDKanban) <> Trim(Kanban.D_Kanbanno) And Trim(OLDSupp) <> Trim(Kanban.D_supp) Then
                                '    ls_sql = "delete UploadKanban where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'" & vbCrLf & _
                                '             " and KanbanNO = '" & Kanban.D_Kanbanno & "'" & vbCrLf & _
                                '             " and (SupplierID = '" & Trim(Kanban.D_supp) & "' or SupplierID = '') " & vbCrLf & _
                                '             " --and PartNo = '" & Kanban.D_PartNo & "' " & vbCrLf

                                '    Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                '    sqlComm9.ExecuteNonQuery()
                                '    sqlComm9.Dispose()
                                'End If
                                'OLDKanban = Kanban.D_Kanbanno
                                'OLDSupp = Kanban.D_supp

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & Kanban.D_PartNo & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "PartNo not found in Part Master, please check again with PASI!"
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

                                    ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                    ls_Qty = 0
                                    If CDbl(Kanban.D_c1) <> 0 Then ls_Qty = CDbl(Kanban.D_c1)

                                    If (ls_Qty Mod ls_MOQ) <> 0 And ls_Qty <> 0 Then
                                        If ls_error = "" Then
                                            ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                        End If
                                    End If
                                End If

                                '02.2.1 Check PartNo di Ms_Price
                                'ls_sql = "select * from ms_partmapping WHERE PartNo = '" & Kanban.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "'"
                                ls_sql = "select * from MS_Price where AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and PartNo = '" & Kanban.D_PartNo & "' and ('" & Kanban.D_DeliveryDate & "' between StartDate and EndDate)"
                                Dim sqlCmd9 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA9 As New SqlDataAdapter(sqlCmd9)
                                Dim ds9 As New DataSet
                                sqlDA9.Fill(ds9)

                                If ds9.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "This PartNo not found or expired in Price Master, please check again with PASI!"
                                    End If
                                End If

                                '02.3 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.UploadKanban WHERE PartNo = '" & Kanban.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and KanbanNo = '" & Kanban.D_Kanbanno & "' and SupplierID = '" & Kanban.D_supp & "'"
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                'If ds4.Tables(0).Rows.Count > 0 Then
                                '    ls_sql = "delete UploadKanban where PartNo = '" & Kanban.D_PartNo & "' and AffiliateID = '" & dtUploadHeader.H_AffiliateID & "' and KanbanNo = '" & Kanban.D_Kanbanno & "' and SupplierID = '" & Kanban.D_supp & "'"
                                '    Dim sqlComm1 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                '    sqlComm1.ExecuteNonQuery()
                                '    sqlComm1.Dispose()
                                'End If

                                ls_sql = " INSERT INTO [dbo].[UploadKanban] " & vbCrLf & _
                                          "            ([AffiliateID], [ShipBy], [SupplierID], [Partno],[Cycle1],[Remarks],[KanbanNo],[DeliveryDate],[KanbanCycle],[direct] )" & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & dtUploadHeader.H_AffiliateID & "' " & vbCrLf & _
                                          "            ,'' " & vbCrLf & _
                                          "            ,'" & Trim(Kanban.D_supp) & "' " & vbCrLf

                                ls_sql = ls_sql + "            ,'" & Trim(Kanban.D_PartNo) & "' " & vbCrLf & _
                                                  "            ,'" & Kanban.D_c1 & "' " & vbCrLf & _
                                                  "            ,'" & ls_error & "' " & vbCrLf & _
                                                  "            , '" & Trim(Kanban.D_Kanbanno) & "' " & vbCrLf & _
                                                  "            , '" & Kanban.D_DeliveryDate & "' " & vbCrLf & _
                                                  "            , '" & Trim(Kanban.D_c2) & "' " & vbCrLf & _
                                                  "            , '" & Trim(Kanban.D_Direct) & "' ) " & vbCrLf
                                Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()

                                If i = 0 Then
                                    ls_TempKanban = "'" & Kanban.D_Kanbanno & "'"
                                End If

                                If Trim(ls_TempKanban) <> Trim(Kanban.D_Kanbanno) Then
                                    ls_TempKanban = ls_TempKanban + ",'" + Kanban.D_Kanbanno + "'"
                                End If

                            Next
                            sqlTran.Commit()

                            Session("DeliveryDate") = dtUploadHeader.H_DeliveryDate
                            Session("PONoUpload") = dtUploadHeader.H_kanbanNo
                            Session("FilterKanbanNoNew") = ls_TempKanban

                            lblInfo.Text = "[7001] Data Checking Done!"
                            lblInfo.ForeColor = Color.Blue
                            grid.JSProperties("cpMessage") = lblInfo.Text


                            Call bindData(ls_TempKanban)
                        End Using
                    Catch ex As Exception
                        MyConnection.Close()
                        lblInfo.Text = ex.Message
                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                        'Exit Sub
                    Finally
                        MyConnection.Close()
                        sqlConn.Close()
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
            lblInfo.Text = ex.Message
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
        Dim ls_TempKanban As String = ""
        Dim ls_KanbanNo As String
        Dim ls_cycle As Integer
        Dim ls_qty As Integer
        Dim ls_Date As String
        Dim ls_time As String
        Dim ls_seq As Integer
        Dim ls_KanbanDate As Date

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
                If i = 0 Then
                    ls_TempSupplierID = "'" & grid.GetRowValues(i, "supplier").ToString.Trim & "'"
                    ls_TempKanban = "'" & grid.GetRowValues(i, "kanbanno").ToString.Trim & "'"
                    countSupplier = 1
                End If

                If ls_TempSupplierID <> grid.GetRowValues(i, "supplier").ToString.Trim Then
                    ls_DoubleSupplier = True
                    ls_TempSupplierID = ls_TempSupplierID + ",'" + grid.GetRowValues(i, "supplier").ToString.Trim + "'"
                    countSupplier = countSupplier + 1
                End If

                If Trim(ls_TempKanban) <> Trim(grid.GetRowValues(i, "kanbanno").ToString.Trim) Then
                    ls_TempKanban = ls_TempKanban + ",'" + grid.GetRowValues(i, "kanbanno").ToString.Trim + "'"
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
            SQLCom.CommandTimeout = 60
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran
            Dim ls_KanbanAsli As String = Trim(Session("PONoUpload"))
            Dim ls_deliveryDate As String = Session("DeliveryDate")

            ls_Sql = "Delete Kanban_Detail where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            ls_Sql = ls_Sql + "Delete Kanban_Master where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            'ls_Sql = ls_Sql + "Delete Kanban_Barcode where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            '2. Insert New Data
            ls_KanbanNo = ""

            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "qty") <> 0 Then
                    'For x = 1 To 4
                    ls_KanbanNo = Trim(grid.GetRowValues(i, "kanbanno")) '& "-" & x
                    ls_deliveryDate = Trim(grid.GetRowValues(i, "deliverydate"))
                    ls_KanbanDate = DateTime.Parse(ls_deliveryDate)
                    'ls_KanbanDate = Right(ls_deliveryDate, 4) + "-" + Mid(ls_deliveryDate, 4, 2) + "-" + Left(ls_deliveryDate, 2)
                    'ls_KanbanDate = Left(ls_KanbanNo, 4) + "-" + Mid(ls_KanbanNo, 5, 2) + "-" + Mid(ls_KanbanNo, 7, 2)
                    ls_Date = Day(ls_KanbanDate) 'IIf(Mid(ls_KanbanNo, 7, 1) = 0, Mid(ls_KanbanNo, 8, 1), Mid(ls_KanbanNo, 7, 2))
                    ls_qty = grid.GetRowValues(i, "qty")

                    'If Len(Trim(ls_KanbanNo)) = 10 Then
                    '    ls_cycle = Mid(Trim(ls_KanbanNo), 10, 1)
                    'ElseIf Len(ls_KanbanNo) = 11 Then
                    '    ls_cycle = Mid(Trim(ls_KanbanNo), 10, 1)
                    'Else
                    '    ls_cycle = Mid(Trim(ls_KanbanNo), 10, 2)
                    'End If
                    ls_cycle = Trim(grid.GetRowValues(i, "cycle"))

                    ls_time = Session("KanbanTime1")

                    If CInt(ls_cycle) <= 4 Then
                        ls_seq = "1"
                    ElseIf CInt(ls_cycle) Mod 4 >= 1 Then
                        ls_seq = (CInt(ls_cycle) \ 4) + 1
                    End If

                    If ls_qty <> 0 Then

                        ls_Sql = " IF NOT EXISTS (select * From Kanban_Master where KanbanNo = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                 "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                 "              and AffiliateID = '" & Session("AffiliateID") & "') BEGIN " & vbCrLf & _
                                 " Insert Into Kanban_Master Values( " & vbCrLf & _
                                 " '" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                 " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                 " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                 " '" & ls_cycle & "', " & vbCrLf & _
                                 " '" & ls_KanbanDate & "', " & vbCrLf & _
                                 " (select kanbantime from ms_kanbantime where affiliateID = '" & Session("AffiliateID") & "' and KanbanCycle = '" & ls_cycle & "'), " & vbCrLf & _
                                 " '0', " & vbCrLf & _
                                 " '', " & vbCrLf & _
                                 " NULL, " & vbCrLf & _
                                 " '' , " & vbCrLf & _
                                 " NULL, " & vbCrLf & _
                                 " GetDate(), " & vbCrLf & _
                                 " '" & Session("UserID").ToString & "', " & vbCrLf & _
                                 " NULL, NULL, " & vbCrLf & _
                                 " '" & Session("DeliveryLoc") & "', " & vbCrLf & _
                                 " '0', " & ls_seq & ") END" & vbCrLf
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()

                        ls_Sql = " IF NOT EXISTS (select * From Kanban_Detail where KanbanNo = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                 "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                 "              and AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                                 "              and partno = '" & grid.GetRowValues(i, "partno") & "' ) BEGIN " & vbCrLf
                        ls_Sql = ls_Sql + "Insert into Kanban_Detail values ( " & vbCrLf & _
                                          " '" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                          " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                          " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                          " '" & grid.GetRowValues(i, "partno") & "', " & vbCrLf & _
                                          " '" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                          " '" & Trim(Session("DeliveryLoc")) & "', " & vbCrLf & _
                                          " '" & Trim(grid.GetRowValues(i, "uom")) & "', " & vbCrLf & _
                                          "  " & CDbl(ls_qty) & ", " & vbCrLf & _
                                          " '" & uf_GetMOQ(grid.GetRowValues(i, "partno"), grid.GetRowValues(i, "supplier"), Session("AffiliateID")) & "', " & vbCrLf & _
                                          " '" & uf_GetQtybox(grid.GetRowValues(i, "partno"), grid.GetRowValues(i, "supplier"), Session("AffiliateID")) & "' ) END " & vbCrLf & _
                                          " ELSE BEGIN " & vbCrLf
                        ls_Sql = ls_Sql + " UPDATE Kanban_Detail set KanbanQty = " & CDbl(ls_qty) & " " & vbCrLf & _
                                          " where KanbanNo = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                          "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                          "              and AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                                          "              and partno = '" & grid.GetRowValues(i, "partno") & "' END "

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
                                 " '" & Format(ls_KanbanDate, "yyyy-MM-01") & "', " & vbCrLf & _
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
                                          " ) END " & vbCrLf

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()

                        ls_Sql = " IF NOT EXISTS (select * From PO_Detail where PoNo = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                 "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                 "              and AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                                 "              and partno = '" & grid.GetRowValues(i, "partno") & "' ) BEGIN " & vbCrLf & _
                                 " INSERT INTO PO_Detail " & vbCrLf & _
                                 " (PONO, AffiliateID, SupplierID, PartNo, KanbanCls, POQty, DeliveryD" & ls_Date & ", EntryDate,EntryUser, UpdateDate, UpdateUser, POMOQ, POQtyBox) " & vbCrLf & _
                                 " Values ('" & Trim(ls_KanbanNo) & "', " & vbCrLf & _
                                 " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                 " '" & grid.GetRowValues(i, "supplier") & "', " & vbCrLf & _
                                 " '" & grid.GetRowValues(i, "partno") & "', " & vbCrLf & _
                                 " '1', " & vbCrLf & _
                                 " " & CDbl(ls_qty) & ", " & vbCrLf & _
                                 " " & CDbl(ls_qty) & ", " & vbCrLf & _
                                 " GETDATE(), " & vbCrLf & _
                                 " '" & Session("AffiliateID") & "', " & vbCrLf & _
                                 " GETDATE(), "
                        ls_Sql = ls_Sql + " '" & Session("AffiliateID") & "', " & vbCrLf & _
										  " '" & uf_GetMOQ(grid.GetRowValues(i, "partno"), grid.GetRowValues(i, "supplier"), Session("AffiliateID")) & "', " & vbCrLf & _
                                          " '" & uf_GetQtybox(grid.GetRowValues(i, "partno"), grid.GetRowValues(i, "supplier"), Session("AffiliateID")) & "' )" & vbCrLf & _
                                          " END ELSE BEGIN " & vbCrLf & _
                                          " UPDATE PO_Detail set POQty = " & CDbl(ls_qty) & ", " & vbCrLf & _
                                          " DeliveryD" & ls_Date & " = " & CDbl(ls_qty) & " " & vbCrLf & _
                                          " where PoNo = '" & Trim(ls_KanbanNo) & "'" & vbCrLf & _
                                          "              and SupplierID = '" & grid.GetRowValues(i, "supplier") & "' " & vbCrLf & _
                                          "              and AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                                          "              and partno = '" & grid.GetRowValues(i, "partno") & "' END "
                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()

                        ls_MsgID = "1001"
                        ls_Detail = "ada"
                    End If
                    'Next
                End If
            Next

            'insert affiliate Master
            ls_Sql = "Delete Affiliate_master where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                     "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                     "              and AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf
            ls_Sql = ls_Sql + "Insert Into Affiliate_master (PONO, AffiliateID, SupplierID, Excelcls,Entrydate, EntryUser, UpdateDate, UpdateUSer) " & vbCrLf & _
                              " select PONO, AffiliateID, SupplierID, '2',Entrydate, EntryUser, UpdateDate, UpdateUSer " & vbCrLf & _
                              " From PO_Master where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                              "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                              "              and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            'Insert Affiliate Detail
            ls_Sql = "Delete Affiliate_Detail where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                     "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                     "              and AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf
            ls_Sql = ls_Sql + " Insert Into Affiliate_Detail " & vbCrLf & _
                              " SELECT PONO, AffiliateID, SupplierID, a.PartNo, '1','','',POQty,POQty, '',0,0, " & vbCrLf & _
                              " DeliveryD1, DeliveryD1, " & vbCrLf & _
                              " DeliveryD2, DeliveryD2, DeliveryD3, DeliveryD3, DeliveryD4, DeliveryD4, DeliveryD5, DeliveryD5," & vbCrLf & _
                              " DeliveryD6, DeliveryD6, DeliveryD7, DeliveryD7, DeliveryD8, DeliveryD8, DeliveryD9, DeliveryD9, " & vbCrLf & _
                              " DeliveryD10, DeliveryD10, DeliveryD11, DeliveryD11, DeliveryD12, DeliveryD12, DeliveryD13, DeliveryD13, " & vbCrLf & _
                              " DeliveryD14, DeliveryD14, DeliveryD15, DeliveryD15, DeliveryD16, DeliveryD16, DeliveryD17, DeliveryD17," & vbCrLf & _
                              " DeliveryD18, DeliveryD18, DeliveryD19, DeliveryD19, DeliveryD20, DeliveryD20, DeliveryD21, DeliveryD21," & vbCrLf & _
                              " DeliveryD22, DeliveryD22, DeliveryD23, DeliveryD23, DeliveryD24, DeliveryD24, DeliveryD25, DeliveryD25," & vbCrLf & _
                              " DeliveryD26, DeliveryD26, DeliveryD27, DeliveryD27, DeliveryD28, DeliveryD28, DeliveryD29, DeliveryD29, DeliveryD30," & vbCrLf & _
                              " DeliveryD30, DeliveryD31, DeliveryD31, a.Entrydate, a.entryuser, a.updatedate, a.updateUser" & vbCrLf & _
                              " FROM PO_Detail a LEFT JOIN MS_Parts b ON a.partno = b.partno" & vbCrLf & _
                              " where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                              "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                              "              and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            'Insert PO_MasterUpload
            ls_Sql = "Delete PO_MasterUpload where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                     "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                     "              and AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf
            ls_Sql = ls_Sql + "Insert Into PO_MasterUpload (PONO, AffiliateID, SupplierID, Remarks,Entrydate, EntryUser, UpdateDate, UpdateUSer)" & vbCrLf & _
                              " select PONO, AffiliateID, SupplierID, '',Entrydate, EntryUser, UpdateDate, UpdateUSer " & vbCrLf & _
                              " From PO_Master where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                              "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                              "              and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            'Insert PO_DetailUpload
            ls_Sql = "Delete PO_DetailUpload where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                     "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                     "              and AffiliateID = '" & Session("AffiliateID") & "'"
            ls_Sql = ls_Sql + " Insert Into PO_DetailUpload " & vbCrLf & _
                              " SELECT distinct PONO, AffiliateID, SupplierID, a.PartNo, '1','','', POQty,POQty,'',0,0, " & vbCrLf & _
                              " DeliveryD1, DeliveryD1, " & vbCrLf & _
                              " DeliveryD2, DeliveryD2, DeliveryD3, DeliveryD3, DeliveryD4, DeliveryD4, DeliveryD5, DeliveryD5," & vbCrLf & _
                              " DeliveryD6, DeliveryD6, DeliveryD7, DeliveryD7, DeliveryD8, DeliveryD8, DeliveryD9, DeliveryD9, " & vbCrLf & _
                              " DeliveryD10, DeliveryD10, DeliveryD11, DeliveryD11, DeliveryD12, DeliveryD12, DeliveryD13, DeliveryD13, " & vbCrLf & _
                              " DeliveryD14, DeliveryD14, DeliveryD15, DeliveryD15, DeliveryD16, DeliveryD16, DeliveryD17, DeliveryD17," & vbCrLf & _
                              " DeliveryD18, DeliveryD18, DeliveryD19, DeliveryD19, DeliveryD20, DeliveryD20, DeliveryD21, DeliveryD21," & vbCrLf & _
                              " DeliveryD22, DeliveryD22, DeliveryD23, DeliveryD23, DeliveryD24, DeliveryD24, DeliveryD25, DeliveryD25," & vbCrLf & _
                              " DeliveryD26, DeliveryD26, DeliveryD27, DeliveryD27, DeliveryD28, DeliveryD28, DeliveryD29, DeliveryD29, DeliveryD30," & vbCrLf & _
                              " DeliveryD30, DeliveryD31, DeliveryD31, a.Entrydate, a.entryuser, a.updatedate, a.updateUser" & vbCrLf & _
                              " FROM PO_Detail a LEFT JOIN MS_Parts b ON a.partno = b.partno" & vbCrLf & _
                              " where PONO IN (" & Trim(ls_TempKanban) & ")" & vbCrLf & _
                              "              and SupplierID IN (" & ls_TempSupplierID & ") " & vbCrLf & _
                              "              and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            'delete tada di tempolary table
            ls_Sql = "delete UploadKanban where AffiliateID = '" & Session("AffiliateID") & "' and KanbanNo IN (" & Trim(ls_TempKanban) & ")"

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

    Private Sub up_ApproveData()
        Dim ls_sql As String
        Dim status As String
        Dim iRow As Integer

        status = "nothing"
        ls_sql = ""

        If grid.VisibleRowCount = 0 Then Exit Sub

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            ls_sql = "SELECT * FROM dbo.Kanban_Master " & vbCrLf & _
                "WHERE AffiliateID = '" & Session("AffiliateID") & "' " & vbCrLf & _
                "AND KanbanNo IN (" & Trim(Session("FilterKanbanNoNew")) & ") " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count = 0 Then
                lblInfo.Text = "[9999] Please Press SAVE it first!"
                Session("YA010IsSubmit") = lblInfo.Text
                grid.JSProperties("cpMessage") = lblInfo.Text
                Exit Sub
            End If

            iRow = 0
            Using sqlTran As SqlTransaction = cn.BeginTransaction()
                Dim sqlComm As New SqlCommand(ls_sql, cn, sqlTran)
                sqlComm.CommandTimeout = 90

                If ds.Tables(0).Rows.Count > 0 Then
                    For iRow = 0 To ds.Tables(0).Rows.Count - 1
                        'approve
                        ls_sql = "UPDATE Kanban_Master " & vbCrLf & _
                            "SET AffiliateApproveUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                            "AffiliateApproveDate = GETDATE(), " & vbCrLf & _
                            "ExcelCls = '1' " & vbCrLf & _
                            "WHERE AffiliateID = '" & Session("AffiliateID") & " ' " & vbCrLf & _
                            "AND SupplierID = '" & ds.Tables(0).Rows(iRow)("SupplierID") & "' " & vbCrLf & _
                            "AND Kanbanno = '" & ds.Tables(0).Rows(iRow)("KanbanNo") & "' " & vbCrLf

                        ls_sql = ls_sql + " DECLARE @KanbanNo AS VARCHAR(25) , " & vbCrLf & _
                                          "     @SupplierID AS VARCHAR(10) , " & vbCrLf & _
                                          "     @SupplierName AS VARCHAR(100) , " & vbCrLf & _
                                          "     @DockID AS VARCHAR(20) , " & vbCrLf & _
                                          "     @PartNo AS VARCHAR(50) , " & vbCrLf & _
                                          "     @PartName AS VARCHAR(100) , " & vbCrLf & _
                                          "     @Qty AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @Cust AS VARCHAR(50) , " & vbCrLf & _
                                          "     @DeliveryDate AS VARCHAR(10) , " & vbCrLf & _
                                          "     @TIME AS VARCHAR(10) , " & vbCrLf & _
                                          "     @Location AS VARCHAR(50) , " & vbCrLf & _
                                          "     @AffCode As Varchar(20), " & vbCrLf

                        ls_sql = ls_sql + "     @DeliveryLocationCode AS VARCHAR(50) , " & vbCrLf & _
                                          "     @PONo AS VARCHAR(50) , " & vbCrLf & _
                                          "     @Barcode AS VARCHAR(1000) , " & vbCrLf & _
                                          "     @Barcode2 AS VARCHAR(1000) , " & vbCrLf & _
                                          "     @QtyBox AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @Loop AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @StartNo AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @EndNo AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @Total AS NUMERIC(10, 2) , " & vbCrLf & _
                                          "     @PartNoSave AS CHAR(50) , " & vbCrLf & _
                                          "     @ETAAffiliate AS VARCHAR(10) , " & vbCrLf

                        ls_sql = ls_sql + "     @ETAPASI AS VARCHAR(10) , " & vbCrLf & _
                                          "     @BoxNo AS VARCHAR(10) , " & vbCrLf & _
                                          "     @AffiliateID AS VARCHAR(10) , " & vbCrLf & _
                                          "     @Cycle AS VARCHAR(3), " & vbCrLf & _
                                          "     @sequence AS Numeric(10,0), " & vbCrLf & _
                                          "     @LabelCode AS VARCHAR(10) " & vbCrLf & _
                                          "     SET @sequence = 0 " & vbCrLf & _
                                          "     SET @AffCode = (Select Top 1 Case When ISNULL(RTRIM(AffiliateCode),'') = '' Then '32G8' Else RTRIM(AffiliateCode) End from MS_Affiliate where AffiliateID = '" & Session("AffiliateID") & "') " & vbCrLf & _
                                          "   SELECT TOP 1 " & vbCrLf & _
                                          "             KanbanNo = CONVERT(CHAR(25), '') , " & vbCrLf & _
                                          "             AffiliateID = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             SupplierID = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             SupplierName = CONVERT(CHAR(100), '') , " & vbCrLf & _
                                          "             PartNo = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             PartName = CONVERT(CHAR(100), '') , " & vbCrLf

                        ls_sql = ls_sql + "             Qty = 0 , " & vbCrLf & _
                                          "             Cust = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             DeliveryDate = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             TIME = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             Location = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             PONo = CONVERT(CHAR(50), '') , " & vbCrLf & _
                                          "             Barcode2 = CONVERT(CHAR(1000), '') , " & vbCrLf & _
                                          "             qtybox = 0 , " & vbCrLf & _
                                          "             startno = 0 , " & vbCrLf & _
                                          "             EndNo = 0 , " & vbCrLf & _
                                          "             total = 0 , " & vbCrLf

                        ls_sql = ls_sql + "             DockID = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             DeliveryLocationCode = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             EtaPasi = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             EtaAffiliate = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             BoxNo = CONVERT(CHAR(10), '') , " & vbCrLf & _
                                          "             Barcode = CONVERT(CHAR(1000), '') , " & vbCrLf & _
                                          "             Cycle = CONVERT(CHAR(3), '') , " & vbCrLf & _
                                          "             LabelCode = CONVERT(VARCHAR(10), '') " & vbCrLf & _
                                          "   INTO      #data " & vbCrLf & _
                                          "   FROM      dbo.Kanban_Master KM   " & vbCrLf & _
                                          "      " & vbCrLf & _
                                          "   DELETE    FROM #data " & vbCrLf & _
                                          "   WHERE     KanbanNo = ''   " & vbCrLf

                        ls_sql = ls_sql + "   DECLARE cur_Print CURSOR FOR   " & vbCrLf & _
                                          "   SELECT  KM.KanbanNo AS kanbanNo ,Km.AffiliateID,   " & vbCrLf & _
                                          "   KM.SupplierID AS SupplierID ,   " & vbCrLf & _
                                          "   MSS.SupplierName AS SupplierName ,  KD.PartNo AS PartNo ,   " & vbCrLf & _
                                          "   MSP.PartName AS PartName ,   " & vbCrLf & _
                                          "   KD.KanbanQty Qty ,   " & vbCrLf & _
                                          "   KM.AffiliateID AS Cust ,   " & vbCrLf & _
                                          "   DeliveryDate = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) ,   " & vbCrLf & _
                                          "   CONVERT(CHAR(5), KM.KanbanTime) AS TIME ,   " & vbCrLf

                        ls_sql = ls_sql + "   '' LocationID ,   " & vbCrLf & _
                                          "   KD.PONo ,   " & vbCrLf & _
                                          "   Barcode2 = @AffCode + ',' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + RTRIM(MSP.PartCarMaker) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','') , " & vbCrLf & _
                                          "   --Barcode2 = '32G8,' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','') , " & vbCrLf & _
                                          "   QtyBox = KD.POQtyBox , " & vbCrLf & _
                                          "   DockID = '', " & vbCrLf & _
                                          "   DeliveryLocationCode = ISNULL(KM.DeliveryLocationCode,''), " & vbCrLf & _
                                          "   ETAAffiliate = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(KM.Kanbandate,'')), 112) , " & vbCrLf

                        ls_sql = ls_sql + "   ETAPASI = CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(ETDPASI,'')), 112) , " & vbCrLf & _
                                          "   BoxNo = '', " & vbCrLf & _
                                          "   Barcode = 'http://zxing.org/w/chart?cht=qr&chs=120x120&chld=L&choe=ISO-8859-1&chl=' + @AffCode + ',' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + RTRIM(MSP.PartCarMaker) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','')   " & vbCrLf & _
                                          "   --Barcode = 'http://zxing.org/w/chart?cht=qr&chs=120x120&chld=L&choe=ISO-8859-1&chl=32G8,' + RTRIM(KD.PONO) + ',' + RTRIM(KM.KanbanNo) + ',' + Rtrim(CONVERT(CHAR(10), CONVERT(DATETIME, ISNULL(KM.Kanbandate, '')), 103)) + ',' + RTRIM(KD.PartNo) + ',' + Replace(Rtrim(KD.POQtyBox),'.00','')   " & vbCrLf & _
                                          "   ,Cycle = KM.KanbanCycle " & vbCrLf & _
                                          "   ,MSS.LabelCode " & vbCrLf & _
                                          "   FROM    dbo.Kanban_Master KM   " & vbCrLf & _
                                          "   LEFT JOIN dbo.Kanban_Detail KD ON KM.AffiliateID = KD.AffiliateID   " & vbCrLf & _
                                          "   AND KM.KanbanNo = KD.KanbanNo   " & vbCrLf & _
                                          "   AND KM.SupplierID = KD.SupplierID   " & vbCrLf & _
                                          "   LEFT JOIN dbo.MS_Supplier MSS ON MSS.SupplierID = KM.SupplierID   " & vbCrLf & _
                                          "   LEFT JOIN dbo.MS_Parts MSP ON MSP.PartNo = KD.PartNo   " & vbCrLf & _
                                          "   LEFT JOIN dbo.MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf

                        ls_sql = ls_sql + "   LEFT JOIN MS_ETD_PASI MEP ON MEP.AffiliateID = KM.AffiliateID " & vbCrLf & _
                                          "   AND CONVERT(CHAR(8), CONVERT(DATETIME, ETAAFfiliate),112) = CONVERT(CHAR(8), CONVERT(DATETIME, KanbanDate),112) " & vbCrLf & _
                                          "   WHERE KanbanQty <> 0 " & vbCrLf & _
                                          "   AND KD. AffiliateID = '" & Session("AffiliateID") & " '" & vbCrLf & _
                                          "   AND KD.SupplierID  = '" & ds.Tables(0).Rows(iRow)("SupplierID") & "'   " & vbCrLf & _
                                          "   AND KM.Kanbanno = '" & ds.Tables(0).Rows(iRow)("KanbanNo") & "'" & vbCrLf

                        ls_sql = ls_sql + " OPEN cur_Print   " & vbCrLf & _
                                          "   FETCH NEXT FROM cur_Print   " & vbCrLf & _
                                          "      INTO @KanbanNo, @AffiliateID, @SupplierID, @SupplierName, @PartNo, @PartName,   " & vbCrLf & _
                                          "   		@Qty, @Cust, @DeliveryDate, @TIME, @Location, @PONo, @Barcode2, @QtyBox, @DockID,  " & vbCrLf & _
                                          "   		@DeliveryLocationCode, @ETAAffiliate, @ETAPASI, @BoxNo, @Barcode, @Cycle, @LabelCode " & vbCrLf & _
                                          "      " & vbCrLf & _
                                          "   WHILE @@Fetch_Status = 0  " & vbCrLf & _
                                          "     BEGIN   " & vbCrLf & _
                                          "         SET @StartNo = 0 " & vbCrLf & _
                                          "         SET @total = 0   " & vbCrLf & _
                                          "         WHILE @Total < @Qty  " & vbCrLf

                        ls_sql = ls_sql + "             BEGIN   " & vbCrLf & _
                                          "                 BEGIN   " & vbCrLf & _
                                          "                      SET @StartNo = @StartNo + 1  " & vbCrLf & _
                                          "                      SET @sequence = @sequence + 1      " & vbCrLf & _
                                          "                      INSERT  INTO #Data  " & vbCrLf & _
                                          "                      VALUES  ( @KanbanNo, @AffiliateID, @SupplierID,  " & vbCrLf & _
                                          "                                @SupplierName, @PartNo, @PartName, @Qty, @Cust,  " & vbCrLf & _
                                          "                                @DeliveryDate, @Time, @Location, @PONo,  " & vbCrLf & _
                                          "                                ( Rtrim(@Barcode2) + ',' + @LabelCode + RIGHT(RTRIM('00000' + REPLACE(CONVERT(NUMERIC, @sequence), '.00', '')), 5))   " & vbCrLf & _
                                          "                               , @QtyBox,  " & vbCrLf & _
                                          "                                RTRIM(CONVERT(NUMERIC, @StartNo)),  " & vbCrLf & _
                                          "                                RTRIM(CONVERT(NUMERIC, ( @Qty / @QtyBox ))),  "

                        ls_sql = ls_sql + "                                ( @Qty / @QtyBox ), @dockID,  " & vbCrLf & _
                                          "                                @DeliveryLocationCode, @ETAPASI, @EtaAffiliate,  " & vbCrLf & _
                                          "                                @LabelCode + RIGHT(RTRIM('00000' + REPLACE(CONVERT(NUMERIC, @sequence), '.00', '')), 5),  " & vbCrLf & _
                                          "                                Rtrim(@Barcode) + ',' + @LabelCode + RIGHT(RTRIM('00000' + REPLACE(CONVERT(NUMERIC, @sequence), '.00', '')), 5) , @Cycle, @LabelCode )  " & vbCrLf & _
                                          "                      SET @Total = @Total + @QtyBox   " & vbCrLf & _
                                          "                 END   " & vbCrLf & _
                                          "             END " & vbCrLf & _
                                          "         FETCH NEXT FROM cur_Print   " & vbCrLf & _
                                          "   		 INTO @KanbanNo, @AffiliateID, @SupplierID, @SupplierName, @PartNo, @PartName,   " & vbCrLf & _
                                          "   			@Qty, @Cust, @DeliveryDate, @TIME, @Location, @PONo, @Barcode2, @QtyBox,@DockID,  " & vbCrLf & _
                                          "   		@DeliveryLocationCode, @ETAAffiliate, @ETAPASI, @BoxNo, @Barcode, @Cycle, @LabelCode   " & vbCrLf

                        ls_sql = ls_sql + "      " & vbCrLf & _
                                          "     END   " & vbCrLf & _
                                          "   CLOSE cur_Print   " & vbCrLf & _
                                          "   DEALLOCATE cur_Print    " & vbCrLf & _
                                          "   INSERT    INTO Kanban_Barcode " & vbCrLf & _
                                          "             SELECT  AffiliateID , " & vbCrLf & _
                                          "                     SupplierID , " & vbCrLf & _
                                          "                     DockID , " & vbCrLf & _
                                          "                     Location , " & vbCrLf & _
                                          "                     EtaAffiliate , " & vbCrLf & _
                                          "                     ETAPasi , " & vbCrLf

                        ls_sql = ls_sql + "                     POno , " & vbCrLf &
                                          "                     KanbanNo , " & vbCrLf &
                                          "                     Cycle , " & vbCrLf &
                                          "                     partNo , " & vbCrLf &
                                          "                     BoxNo ,  " & vbCrLf &
                                          "                     Startno , " & vbCrLf &
                                          "                     EndNo , " & vbCrLf &
                                          "                     QtyBox , " & vbCrLf &
                                          "                     Barcode , " & vbCrLf &
                                          "                     DeliveryLocationCode , " & vbCrLf &
                                          "                     Barcode2 "

                        ls_sql = ls_sql + "             FROM    #data    " & vbCrLf

                        ls_sql = ls_sql + "             " & vbCrLf & _
                                          "             DROP TABLE  #data    " & vbCrLf

                        sqlComm = New SqlCommand(ls_sql, cn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                    Next
                End If

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using
        End Using

        Call clsMsg.DisplayMessage(lblInfo, "1012", clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text
    End Sub

#End Region

    Private Sub grid_RowInserting(sender As Object, e As DevExpress.Web.Data.ASPxDataInsertingEventArgs) Handles grid.RowInserting
        e.Cancel = True
    End Sub

    Private Sub grid_RowUpdating(sender As Object, e As DevExpress.Web.Data.ASPxDataUpdatingEventArgs) Handles grid.RowUpdating
        e.Cancel = True
    End Sub
End Class