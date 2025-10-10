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

Public Class ForecastUploadMonthly
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
                'btnDownload.Enabled = False
            End If
            lblInfo.Text = ""
        Else
            lblInfo.Text = ""
            Ext = Server.MapPath("")
        End If
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowPager)
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

    Private Sub bindData(ByVal Part As String, ByVal Period As Integer, ByVal Affiliate As String, ByVal Rev As Integer)
        Dim ls_SQL As String = ""
        Dim pWhereKanban As String = ""
        Dim pWhereDifference As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim PrevRev As Integer
            If Rev = 0 Then
                PrevRev = Rev
            Else
                PrevRev = Rev - 1
            End If

            grid.VisibleColumns(5).Caption = "Forecast Quantity Jul-" & Period
            grid.VisibleColumns(6).Caption = "Forecast Quantity Aug-" & Period
            grid.VisibleColumns(7).Caption = "Forecast Quantity Sep-" & Period
            grid.VisibleColumns(8).Caption = "Forecast Quantity Oct-" & Period
            grid.VisibleColumns(9).Caption = "Forecast Quantity Nov-" & Period
            grid.VisibleColumns(10).Caption = "Forecast Quantity Dec-" & Period
            grid.VisibleColumns(11).Caption = "Forecast Quantity Jan-" & Period + 1
            grid.VisibleColumns(12).Caption = "Forecast Quantity Feb-" & Period + 1
            grid.VisibleColumns(13).Caption = "Forecast Quantity Mar-" & Period + 1
            grid.VisibleColumns(14).Caption = "Forecast Quantity Apr-" & Period + 1
            grid.VisibleColumns(15).Caption = "Forecast Quantity May-" & Period + 1
            grid.VisibleColumns(16).Caption = "Forecast Quantity Jun-" & Period + 1
      

            ls_SQL = " Select row_number() over (order by UFM.[Year],UFM.[AffiliateID],UFM.[Rev],UFM.[PartNo] asc) as no,UFM.[Year],UFM.[AffiliateID],UFM.[Rev],UFM.[PartNo] " & vbCrLf & _
                     "       ,UFM.[Jul],UFM.[Aug],UFM.[Sep],UFM.[Oct],UFM.[Nov],UFM.[Dec],UFM.[Jan],UFM.[Feb],UFM.[Mar],UFM.[Apr],UFM.[May],UFM.[Jun],remarks " & vbCrLf & _
                     "       ,C7=ISNULL(UFM.[Jul]-FM.[Jul],0),C8=ISNULL(UFM.[Aug]-FM.[Aug],0),C9=ISNULL(UFM.[Sep]-FM.[Sep],0),C10=ISNULL(UFM.[Oct]-FM.[Oct],0),C11=ISNULL(UFM.[Nov]-FM.[Nov],0),C12=ISNULL(UFM.[Dec]-FM.[Dec],0) " & vbCrLf & _
                     "       ,C1=ISNULL(UFM.[Jan]-FM.[Jan],0),C2=ISNULL(UFM.[Feb]-FM.[Feb],0),C3=ISNULL(UFM.[Mar]-FM.[Mar],0),C4=ISNULL(UFM.[Apr]-FM.[Apr],0),C5=ISNULL(UFM.[May]-FM.[May],0),C6=ISNULL(UFM.[Jun]-FM.[Jun],0)" & vbCrLf & _
                     " From UploadForecastMonthly UFM" & vbCrLf & _
                     " Left Join ForecastMonthly FM On UFM.Year = FM.Year And UFM.AffiliateID = FM.AffiliateID And UFM.PartNo = FM.PartNo And FM.Rev = '" & PrevRev & "' " & vbCrLf & _
                     " Where UFM.Year = '" & Period & "' And UFM.AffiliateID = '" & Affiliate & "' And UFM.Rev = '" & Rev & "' " & vbCrLf & _
                     " And UFM.PartNo IN (" & Part & ") "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 4, False, clsAppearance.PagerMode.ShowAllRecord)
            End With
            sqlConn.Close()

            'clsGlobal.HideColumTanggal1(Session("Period"), grid)
        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' no, '' kanbanno, '' PartNo, '' PartName, '' qty"

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
        'Dim dt As New System.Data.DataTable
        Dim dtHeader As New System.Data.DataTable
        Dim dtDetail As New System.Data.DataTable
        Dim tempDate As Date
        Dim ls_MOQ As Double = 0
        Dim ls_sql As String = ""
        Dim ls_SupplierID As String = """"
        Dim ls_Part As String = ""
        Dim pDeleteSupplier As String = ""
        Dim pDeletePart As String = ""

        Dim connStr As String = ""

        Dim pYear As Integer
        Dim pAffiliateID As String = Session("AffiliateID")
        Dim pRev As Integer

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

                    'Get Header Data
                    MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A4:C6]")
                    MyAdapter.SelectCommand = MyCommand
                    MyAdapter.Fill(dtHeader)

                    If dtHeader.Rows.Count > 0 Then
                        'PERIOD
                        If IsDBNull(dtHeader.Rows(0).Item(2)) Then
                            lblInfo.Text = "[9999] Please input Year, please check the file again!"
                            grid.JSProperties("cpMessage") = lblInfo.Text
                            MyConnection.Close()
                            Exit Sub
                        Else
                            Try
                                pYear = dtHeader.Rows(0).Item(2)
                            Catch ex As Exception
                                lblInfo.Text = "[9999] Invalid Format Year, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End Try
                        End If

                        ''AFFILIATE
                        'If IsDBNull(dtHeader.Rows(1).Item(2)) Then
                        '    lblInfo.Text = "[9999] Please input Affiliate Code!"
                        '    grid.JSProperties("cpMessage") = lblInfo.Text
                        '    MyConnection.Close()
                        '    Exit Sub
                        'Else
                        '    '03.1 Check Affiliate di PO MASTER
                        '    ls_sql = "SELECT * FROM MS_Affiliate WHERE AffiliateID = '" & dtHeader.Rows(1).Item(2) & "'"
                        '    Dim sqlCmd5 As New SqlCommand(ls_sql, sqlConn)
                        '    Dim sqlDA5 As New SqlDataAdapter(sqlCmd5)
                        '    Dim ds5 As New DataSet
                        '    sqlDA5.Fill(ds5)

                        '    If ds5.Tables(0).Rows.Count = 0 Then
                        '        lblInfo.Text = "[9999] Affiliate Code not valid, please check the file again!"
                        '        grid.JSProperties("cpMessage") = lblInfo.Text
                        '        MyConnection.Close()
                        '        Exit Sub
                        '    End If

                        '    pAffiliateID = dtHeader.Rows(1).Item(2)
                        'End If

                        'REVISION
                        If IsDBNull(dtHeader.Rows(2).Item(2)) Then
                            lblInfo.Text = "[9999] Please input Revision, please check the file again!"
                            grid.JSProperties("cpMessage") = lblInfo.Text
                            MyConnection.Close()
                            Exit Sub
                        Else
                            Try
                                pRev = dtHeader.Rows(2).Item(2)
                                If pRev > "3" Then
                                    lblInfo.Text = "[9999] Revision > 3, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                lblInfo.Text = "[9999] Invalid Format Revision, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End Try
                        End If

                        ''Check Revision yang sudah pernah diinput
                        'ls_sql = "SELECT * FROM ForecastMonthly WHERE Year = '" & pYear & "' AND AffiliateID = '" & pAffiliateID & "' AND Rev = '" & pRev & "' "
                        'Dim sqlCmd6 As New SqlCommand(ls_sql, sqlConn)
                        'Dim sqlDA6 As New SqlDataAdapter(sqlCmd6)
                        'Dim ds6 As New DataSet
                        'sqlDA6.Fill(ds6)

                        'If ds6.Tables(0).Rows.Count > 0 Then
                        '    If ds6.Tables(0).Rows(0).Item("Rev") = pRev Then
                        '        lblInfo.Text = "[9999] Rev " & pRev & " Already Upload, please check the file again!"
                        '        grid.JSProperties("cpMessage") = lblInfo.Text
                        '        MyConnection.Close()
                        '        Exit Sub
                        '    End If
                        '    If ds6.Tables(0).Rows(0).Item("Rev") + 1 <> pRev Then
                        '        lblInfo.Text = "[9999] Please Upload Previous Rev First, please check the file again!"
                        '        grid.JSProperties("cpMessage") = lblInfo.Text
                        '        MyConnection.Close()
                        '        Exit Sub
                        '    End If
                        'Else
                        '    'Pertama Upload Harus Rev 0
                        '    If pRev > 0 Then
                        '        If ds6.Tables(0).Rows(0).Item("Rev") + 1 <> pRev Then
                        '            lblInfo.Text = "[9999] Please Upload Rev 0 First, please check the file again!"
                        '            grid.JSProperties("cpMessage") = lblInfo.Text
                        '            MyConnection.Close()
                        '            Exit Sub
                        '        End If

                        '    End If
                        'End If
                        'Check Revision yang sudah pernah diinput
                        ls_sql = "SELECT DISTINCT Rev FROM ForecastMonthly WHERE Year = '" & pYear & "' AND AffiliateID = '" & pAffiliateID & "' AND Rev = '" & pRev & "' " & vbCrLf & _
                                 "Order By Rev Desc"
                        Dim sqlCmd6 As New SqlCommand(ls_sql, sqlConn)
                        Dim sqlDA6 As New SqlDataAdapter(sqlCmd6)
                        Dim ds6 As New DataSet
                        sqlDA6.Fill(ds6)

                        If ds6.Tables(0).Rows.Count > 0 Then
                            If ds6.Tables(0).Rows(0).Item("Rev") = pRev Then
                                lblInfo.Text = "[9999] Rev " & pRev & " Already Upload, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End If
                        End If

                        'Check Urutan Revision
                        ls_sql = "SELECT DISTINCT Rev FROM ForecastMonthly WHERE Year = '" & pYear & "' AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                 "Order By Rev Desc"
                        Dim sqlCmd7 As New SqlCommand(ls_sql, sqlConn)
                        Dim sqlDA7 As New SqlDataAdapter(sqlCmd7)
                        Dim ds7 As New DataSet
                        sqlDA7.Fill(ds7)

                        If ds7.Tables(0).Rows.Count > 0 Then
                            If ds7.Tables(0).Rows(0).Item("Rev") + 1 <> pRev Then
                                lblInfo.Text = "[9999] Please Upload Previous Rev First, please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                MyConnection.Close()
                                Exit Sub
                            End If
                        Else
                            'Pertama Upload Harus Rev 0
                            If pRev > 0 Then
                                If ds7.Tables(0).Rows(0).Item("Rev") + 1 <> pRev Then
                                    lblInfo.Text = "[9999] Please Upload Rev 0 First, please check the file again!"
                                    grid.JSProperties("cpMessage") = lblInfo.Text
                                    MyConnection.Close()
                                    Exit Sub
                                End If

                            End If
                        End If
                    End If


                    Dim dtUploadDetailList As New List(Of clsForecastMonthly)

                    'Get Detail Data
                    MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A11:AL65536]")
                    MyAdapter.SelectCommand = MyCommand
                    MyAdapter.Fill(dtDetail)

                    If dtDetail.Rows.Count > 0 Then
                        For i = 0 To dtDetail.Rows.Count - 1
                            If CStr(IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))) <> "" Then
                                Dim dtUploadDetail As New clsForecastMonthly
                                dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(1)
                                If pDeletePart = "" Then
                                    pDeletePart = "'" + Trim(dtUploadDetail.D_PartNo) + "'"
                                Else
                                    pDeletePart = pDeletePart + ",'" + Trim(dtUploadDetail.D_PartNo) + "'"
                                End If

                                dtUploadDetail.D_Year = pYear
                                dtUploadDetail.D_AffiliateID = pAffiliateID
                                dtUploadDetail.D_Rev = pRev

                                'dtUploadDetail.D_PartNo = dtDetail.Rows(i).Item(1)
                                dtUploadDetail.D_Jul = IIf(IsDBNull(dtDetail.Rows(i).Item(2)) = True, 0, dtDetail.Rows(i).Item(2))
                                dtUploadDetail.D_Aug = IIf(IsDBNull(dtDetail.Rows(i).Item(3)) = True, 0, dtDetail.Rows(i).Item(3))
                                dtUploadDetail.D_Sep = IIf(IsDBNull(dtDetail.Rows(i).Item(4)) = True, 0, dtDetail.Rows(i).Item(4))
                                dtUploadDetail.D_Oct = IIf(IsDBNull(dtDetail.Rows(i).Item(5)) = True, 0, dtDetail.Rows(i).Item(5))
                                dtUploadDetail.D_Nov = IIf(IsDBNull(dtDetail.Rows(i).Item(6)) = True, 0, dtDetail.Rows(i).Item(6))
                                dtUploadDetail.D_Dec = IIf(IsDBNull(dtDetail.Rows(i).Item(7)) = True, 0, dtDetail.Rows(i).Item(7))
                                dtUploadDetail.D_Jan = IIf(IsDBNull(dtDetail.Rows(i).Item(8)) = True, 0, dtDetail.Rows(i).Item(8))
                                dtUploadDetail.D_Feb = IIf(IsDBNull(dtDetail.Rows(i).Item(9)) = True, 0, dtDetail.Rows(i).Item(9))
                                dtUploadDetail.D_Mar = IIf(IsDBNull(dtDetail.Rows(i).Item(10)) = True, 0, dtDetail.Rows(i).Item(10))
                                dtUploadDetail.D_Apr = IIf(IsDBNull(dtDetail.Rows(i).Item(11)) = True, 0, dtDetail.Rows(i).Item(11))
                                dtUploadDetail.D_May = IIf(IsDBNull(dtDetail.Rows(i).Item(12)) = True, 0, dtDetail.Rows(i).Item(12))
                                dtUploadDetail.D_Jun = IIf(IsDBNull(dtDetail.Rows(i).Item(13)) = True, 0, dtDetail.Rows(i).Item(13))

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

                        ls_sql = "delete UploadForecastMonthly where AffiliateID = '" & pAffiliateID & "'" & vbCrLf & _
                                 " and Year = '" & pYear & "' " & vbCrLf & _
                                 " and Rev = '" & pRev & "' " & vbCrLf & _
                                 " and PartNo IN (" & pDeletePart & ") " & vbCrLf

                        Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                        sqlComm9.ExecuteNonQuery()
                        sqlComm9.Dispose()


                        '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                        For i = 0 To dtUploadDetailList.Count - 1
                            Dim ls_error As String = ""
                            Dim Part As clsForecastMonthly = dtUploadDetailList(i)
                            'Dim ls_Qty As Integer


                            '02.1 Check PartNo di MS_Part
                            ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & Part.D_PartNo & "' "
                            Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                            Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                            Dim ds2 As New DataSet
                            sqlDA2.Fill(ds2)

                            If ds2.Tables(0).Rows.Count = 0 Then
                                ls_error = "PartNo not found in Part Master, please check again with PASI!"
                            End If

                            '02.2 Check PartNo di Ms_PartMapping
                            ls_sql = "select * from ms_partmapping WHERE PartNo = '" & Part.D_PartNo & "' and AffiliateID = '" & pAffiliateID & "'"
                            Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                            Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                            Dim ds3 As New DataSet
                            sqlDA3.Fill(ds3)

                            If ds3.Tables(0).Rows.Count = 0 Then
                                If ls_error = "" Then
                                    ls_error = "PartNo not found in Part Mapping, please check again with PASI!"
                                End If
                            Else
                                'ls_SupplierID = ds3.Tables(0).Rows(0)("SupplierID")

                                'ls_MOQ = IIf(IsDBNull(ds3.Tables(0).Rows(0)("MOQ")), 0, ds3.Tables(0).Rows(0)("MOQ"))
                                'ls_Qty = 0
                                'If CDbl(Kanban.D_c1) <> 0 Then ls_Qty = CDbl(Kanban.D_c1)

                                'If (ls_Qty Mod ls_MOQ) <> 0 And ls_Qty <> 0 Then
                                '    If ls_error = "" Then
                                '        ls_error = "Total Firm Qty must be same or multiple of the MOQ!, please check the file again!"
                                '    End If
                                'End If
                            End If

                            ls_sql = " INSERT INTO [dbo].[UploadForecastMonthly] " & vbCrLf & _
                                      "            ([Year],[AffiliateID],[Rev],[PartNo],[Jan],[Feb],[Mar],[Apr] " & vbCrLf & _
                                      "             ,[May],[Jun],[Jul],[Aug],[Sep],[Oct],[Nov],[Dec],remarks) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & Part.D_Year & "' " & vbCrLf & _
                                      "            ,'" & Part.D_AffiliateID & "' " & vbCrLf & _
                                      "            ,'" & Part.D_Rev & "' " & vbCrLf & _
                                      "            ,'" & Part.D_PartNo & "' " & vbCrLf & _
                                      "            ,'" & Part.D_Jan & "' " & vbCrLf & _
                                      "            ,'" & Part.D_Feb & "' " & vbCrLf & _
                                      "            ,'" & Part.D_Mar & "' " & vbCrLf & _
                                      "            ,'" & Part.D_Apr & "' " & vbCrLf

                            ls_sql = ls_sql + "            ,'" & Part.D_May & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Jun & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Jul & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Aug & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Sep & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Oct & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Nov & "' " & vbCrLf & _
                                              "            ,'" & Part.D_Dec & "' " & vbCrLf & _
                                              "            ,'" & ls_error & "') " & vbCrLf
                            Dim sqlComm = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                            If i = 0 Then
                                ls_Part = "'" & Part.D_PartNo & "'"
                            End If

                            If Trim(ls_Part) <> Trim(Part.D_PartNo) Then
                                ls_Part = ls_Part + ",'" + Part.D_PartNo + "'"
                            End If

                        Next
                        sqlTran.Commit()

                        'Session("DeliveryDate") = dtUploadHeader.H_DeliveryDate
                        'Session("PONoUpload") = dtUploadHeader.H_kanbanNo
                        'Session("FilterKanbanNoNew") = ls_TempKanban

                        lblInfo.Text = "[7001] Data Checking Done!"
                        lblInfo.ForeColor = Color.Blue
                        grid.JSProperties("cpMessage") = lblInfo.Text


                        Call bindData(ls_Part, pYear, pAffiliateID, pRev)
                    End Using
                Catch ex As Exception
                    MyConnection.Close()
                    lblInfo.Text = ex.Message
                    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
                    If InStr(ex.Message, "Cannot find column") = 1 Then
                        lblInfo.Text = "Format Template Tidak Sesuai"
                        Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, lblInfo.Text.ToString())
                    End If
                    'Exit Sub
                Finally
                    MyConnection.Close()
                    sqlConn.Close()
                End Try
                'dt.Reset()
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
        'Catch ex As Exception
        '    MyConnection.Close()
        '    Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        'End Try
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
        'Dim ls_Period As Date
        Dim ls_ShipBy As String = ""
        Dim ls_Detail As String = ""
        Dim ls_DoubleSupplier As Boolean = False
        Dim ls_TempSupplierID As String = ""
        Dim ls_TempKanban As String = ""
        'Dim ls_KanbanNo As String
        Dim ls_cycle As Integer
        Dim ls_qty As Integer
        Dim ls_Date As String
        Dim ls_time As String
        Dim ls_seq As Integer
        'Dim ls_KanbanDate As Date

        Dim ls_Year As Integer
        Dim ls_Rev As Integer
        Dim ls_Affiliate As String = ""
        Dim ls_TempPart As String = ""
        Dim ls_Part As String = ""

        Dim C1 As Integer
        Dim C2 As Integer
        Dim C3 As Integer
        Dim C4 As Integer
        Dim C5 As Integer
        Dim C6 As Integer
        Dim C7 As Integer
        Dim C8 As Integer
        Dim C9 As Integer
        Dim C10 As Integer
        Dim C11 As Integer
        Dim C12 As Integer


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
                ls_Year = grid.GetRowValues(0, "Year").ToString.Trim
                ls_Rev = grid.GetRowValues(0, "Rev").ToString.Trim
                ls_Affiliate = grid.GetRowValues(0, "AffiliateID").ToString.Trim

                If i = 0 Then
                    'ls_TempSupplierID = "'" & grid.GetRowValues(i, "supplier").ToString.Trim & "'"
                    'ls_TempKanban = "'" & grid.GetRowValues(i, "kanbanno").ToString.Trim & "'"
                    'countSupplier = 1
                    ls_TempPart = "'" & grid.GetRowValues(i, "PartNo").ToString.Trim & "'"
                End If

                'If ls_TempSupplierID <> grid.GetRowValues(i, "supplier").ToString.Trim Then
                '    ls_DoubleSupplier = True
                '    ls_TempSupplierID = ls_TempSupplierID + ",'" + grid.GetRowValues(i, "supplier").ToString.Trim + "'"
                '    countSupplier = countSupplier + 1
                'End If

                'If Trim(ls_TempKanban) <> Trim(grid.GetRowValues(i, "kanbanno").ToString.Trim) Then
                '    ls_TempKanban = ls_TempKanban + ",'" + grid.GetRowValues(i, "kanbanno").ToString.Trim + "'"
                'End If
                If Trim(ls_TempPart) <> Trim(grid.GetRowValues(i, "PartNo").ToString.Trim) Then
                    ls_TempPart = ls_TempPart + ",'" + grid.GetRowValues(i, "PartNo").ToString.Trim + "'"
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
            SQLCom.Connection = SqlCon
            SQLCom.Transaction = SqlTran
            'Dim ls_KanbanAsli As String = Trim(Session("PONoUpload"))
            'Dim ls_deliveryDate As String = Session("DeliveryDate")

            ls_Sql = "Delete ForecastMonthly where Year = '" & ls_Year & "' and Rev = '" & ls_Rev & "' and AffiliateID = '" & ls_Affiliate & "' And PartNo IN (" & ls_TempPart & ")"
            'ls_Sql = "Delete Kanban_Detail where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            'ls_Sql = ls_Sql + "Delete Kanban_Master where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            'ls_Sql = ls_Sql + "Delete Kanban_Barcode where KanbanNo IN (" & ls_TempKanban & ") and SupplierID IN (" & ls_TempSupplierID & ") and AffiliateID = '" & Session("AffiliateID") & "'"
            SQLCom.CommandText = ls_Sql
            SQLCom.ExecuteNonQuery()

            '2. Insert New Data
            'ls_KanbanNo = ""

            For i = 0 To grid.VisibleRowCount - 1
                ls_Year = Trim(grid.GetRowValues(i, "Year"))
                ls_Rev = Trim(grid.GetRowValues(i, "Rev"))
                ls_Affiliate = Trim(grid.GetRowValues(i, "AffiliateID"))
                ls_Part = Trim(grid.GetRowValues(i, "PartNo"))

                C1 = IIf(Trim(grid.GetRowValues(i, "C1")) = 0, 0, 1)
                C2 = IIf(Trim(grid.GetRowValues(i, "C2")) = 0, 0, 1)
                C3 = IIf(Trim(grid.GetRowValues(i, "C3")) = 0, 0, 1)
                C4 = IIf(Trim(grid.GetRowValues(i, "C4")) = 0, 0, 1)
                C5 = IIf(Trim(grid.GetRowValues(i, "C5")) = 0, 0, 1)
                C6 = IIf(Trim(grid.GetRowValues(i, "C6")) = 0, 0, 1)
                C7 = IIf(Trim(grid.GetRowValues(i, "C7")) = 0, 0, 1)
                C8 = IIf(Trim(grid.GetRowValues(i, "C8")) = 0, 0, 1)
                C9 = IIf(Trim(grid.GetRowValues(i, "C9")) = 0, 0, 1)
                C10 = IIf(Trim(grid.GetRowValues(i, "C10")) = 0, 0, 1)
                C11 = IIf(Trim(grid.GetRowValues(i, "C11")) = 0, 0, 1)
                C12 = IIf(Trim(grid.GetRowValues(i, "C12")) = 0, 0, 1)

                ls_Sql = " INSERT INTO [dbo].[ForecastMonthly] " & vbCrLf & _
                                      "            ([Year],[YearTo],[AffiliateID],[Rev],[PartNo],[Jan],[Feb],[Mar],[Apr] " & vbCrLf & _
                                      "             ,[May],[Jun],[Jul],[Aug],[Sep],[Oct],[Nov],[Dec],EntryDate,EntryUser " & vbCrLf & _
                                      "             ,[C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10],[C11],[C12]) " & vbCrLf & _
                                      "      VALUES " & vbCrLf & _
                                      "            ('" & ls_Year & "' " & vbCrLf & _
                                      "            ,'" & ls_Year + 1 & "' " & vbCrLf & _
                                      "            ,'" & ls_Affiliate & "' " & vbCrLf & _
                                      "            ,'" & ls_Rev & "' " & vbCrLf & _
                                      "            ,'" & ls_Part & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "Jan") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "Feb") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "Mar") & "' " & vbCrLf & _
                                      "            ,'" & grid.GetRowValues(i, "Apr") & "' " & vbCrLf

                ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "May") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Jun") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Jul") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Aug") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Sep") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Oct") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Nov") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Dec") & "' " & vbCrLf & _
                                  "            ,GETDATE(),'" & Session("UserID").ToString & "' " & vbCrLf & _
                                  "            ,'" & C1 & "' " & vbCrLf & _
                                  "            ,'" & C2 & "' " & vbCrLf & _
                                  "            ,'" & C3 & "' " & vbCrLf & _
                                  "            ,'" & C4 & "' " & vbCrLf & _
                                  "            ,'" & C5 & "' " & vbCrLf & _
                                  "            ,'" & C6 & "' " & vbCrLf & _
                                  "            ,'" & C7 & "' " & vbCrLf & _
                                  "            ,'" & C8 & "' " & vbCrLf & _
                                  "            ,'" & C9 & "' " & vbCrLf & _
                                  "            ,'" & C10 & "' " & vbCrLf & _
                                  "            ,'" & C11 & "' " & vbCrLf & _
                                  "            ,'" & C12 & "') " & vbCrLf

                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()


                ls_MsgID = "1001"
                ls_Detail = "ada"
            Next

            'delete tada di tempolary table
            'ls_Sql = "delete UploadKanban where AffiliateID = '" & Session("AffiliateID") & "' and KanbanNo IN (" & Trim(ls_TempKanban) & ")"
            ls_Sql = "Delete UploadForecastMonthly where Year = '" & ls_Year & "' and Rev = '" & ls_Rev & "' and AffiliateID = '" & ls_Affiliate & "' And PartNo IN (" & ls_TempPart & ")"

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
#End Region
End Class