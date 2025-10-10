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

Public Class UploadPASIPrice
#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "A13"
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
        Response.Redirect("~/Master/PASIPriceMaster.aspx")
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
                     "   	a.AffiliateID,b.AffiliateName,a.PartNo,c.PartName,d.Description CurrCls, e.Description PackingCls, f.Description PriceCls,   " & vbCrLf & _
                     "      a.[Price], a.[StartDate], a.[EndDate],a.[EntryDate],[ErrorCls],  " & vbCrLf & _
                     "  	xaffiliateid = mp.AffiliateID, xpartno = mp.PartNo, xpackingtype = g.Description, xpricecategory = i.Description, xcurrcls = h.Description, xprice = mp.price, xstartdate = mp.StartDate, xenddate = mp.EndDate, xeffectivedate = mp.EffectiveDate	    " & vbCrLf & _
                     " from [UploadPrice] a   " & vbCrLf & _
                     "       left join MS_Affiliate b on a.AffiliateID = b.AffiliateID   " & vbCrLf & _
                     "       left join MS_Parts c on c.PartNo = a.PartNo   " & vbCrLf & _
                     "       left join MS_CurrCls d on d.CurrCls = a.CurrCls   " & vbCrLf & _
                     "       left join MS_PackingCls e on a.PackingCls = e.PackingCls   " & vbCrLf & _
                     "       left join MS_PriceCls f on a.PriceCls = f.PriceCls    "

            ls_SQL = ls_SQL + " 	  left join MS_Price mp  on mp.AffiliateID = a.AffiliateID and mp.PartNo = a.PartNo and mp.PackingCls = a.PackingCls and mp.StartDate = a.StartDate and mp.CurrCls = a.Currcls and mp.PriceCls = a.PriceCls " & vbCrLf & _
                              " 	  left join MS_PackingCls g on mp.PackingCls = g.PackingCls " & vbCrLf & _
                              " 	  left join MS_CurrCls h on h.CurrCls = MP.CurrCls   " & vbCrLf & _
                              "       left join MS_PriceCls i on mp.PriceCls = i.PriceCls" & vbCrLf & _
                              " order by a.AffiliateID " & vbCrLf & _
                              "  "

            'ls_SQL = " select  " & vbCrLf & _
            '      " 	row_number() over (order by a.AffiliateID asc) as NoUrut, " & vbCrLf & _
            '      " 	a.AffiliateID,b.AffiliateName,a.PartNo,c.PartName,d.Description CurrCls, e.Description PackingCls, f.Description PriceCls, " & vbCrLf & _
            '      "     [Price],[StartDate],[EndDate],a.[EntryDate],[ErrorCls] "

            'ls_SQL = ls_SQL + " from [UploadPrice] a left join MS_Affiliate b on a.AffiliateID = b.AffiliateID " & vbCrLf & _
            '                  " left join MS_Parts c on c.PartNo = a.PartNo " & vbCrLf & _
            '                  " left join MS_CurrCls d on d.CurrCls = a.CurrCls " & vbCrLf & _
            '                  " left join MS_PackingCls e on a.PackingCls = e.PackingCls " & vbCrLf & _
            '                  " left join MS_PriceCls f on a.PriceCls = f.PriceCls " & vbCrLf & _
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

            ls_SQL = " select top 0 '' NoUrut, '' [PriceCls], ''[PackingCls], '' [PartNo],'' [PartName], '' [AffiliateID], '' [AffiliateName],'' [CurrCls],'' [Price],'' [StartDate],'' [EndDate],'' [EntryDate],'' [ErrorCls]"

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

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "A3:I65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IsDBNull(dtDetail.Rows(i).Item(0)) = False Then
                                    Dim dtUploadDetail As New clsMaster
                                    dtUploadDetail.AffiliateID = dtDetail.Rows(i).Item(0)
                                    dtUploadDetail.PartNo = IIf(IsDBNull(dtDetail.Rows(i).Item(1)), "", dtDetail.Rows(i).Item(1))

                                    If IsDBNull(dtDetail.Rows(i).Item(2)) = False Then
                                        If Trim(dtDetail.Rows(i).Item(2)).ToUpper = "CARTON BOX" Then
                                            dtUploadDetail.PackingCls = "01"
                                        ElseIf Trim(dtDetail.Rows(i).Item(2)).ToUpper = "IMPRABOARD" Then
                                            dtUploadDetail.PackingCls = "02"
                                        ElseIf Trim(dtDetail.Rows(i).Item(2)).ToUpper = "POLYBOX" Then
                                            dtUploadDetail.PackingCls = "03"
                                        ElseIf Trim(dtDetail.Rows(i).Item(2)).ToUpper = "KMT 40" Then
                                            dtUploadDetail.PackingCls = "04"
                                        Else
                                            dtUploadDetail.PackingCls = ""
                                        End If
                                    Else
                                        dtUploadDetail.PackingCls = ""
                                    End If

                                    If Not IsDBNull(dtDetail.Rows(i).Item(3)) Then
                                        If Trim(dtDetail.Rows(i).Item(3)).ToUpper = "FCA - AIR" Then
                                            dtUploadDetail.PriceCategory = "'1'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "FCA - BOAT" Then
                                            dtUploadDetail.PriceCategory = "'2'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "CIF - AIR" Then
                                            dtUploadDetail.PriceCategory = "'3'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "CIF - BOAT" Then
                                            dtUploadDetail.PriceCategory = "'4'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "DDU PASI" Then
                                            dtUploadDetail.PriceCategory = "'5'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "DDU AFFILIATE" Then
                                            dtUploadDetail.PriceCategory = "'6'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "EX-WORK" Then
                                            dtUploadDetail.PriceCategory = "'7'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "FCA" Then
                                            dtUploadDetail.PriceCategory = "'9'"
                                        ElseIf Trim(dtDetail.Rows(i).Item(3)).ToUpper = "CIF" Then
                                            dtUploadDetail.PriceCategory = "'10'"
                                        Else
                                            dtUploadDetail.PriceCategory = "''"
                                        End If
                                    Else
                                        dtUploadDetail.PriceCategory = "''"
                                    End If

                                    'dtUploadDetail.PriceCategory = IIf(IsDBNull(dtDetail.Rows(i).Item(3)), "", dtDetail.Rows(i).Item(3))

                                    Dim tempCls As String = IIf(IsDBNull(dtDetail.Rows(i).Item(4)), "IDR", dtDetail.Rows(i).Item(4))

                                    If tempCls = "JPY" Then
                                        dtUploadDetail.CurrCls = "01"
                                    ElseIf tempCls = "USD" Then
                                        dtUploadDetail.CurrCls = "02"
                                    ElseIf tempCls = "IDR" Then
                                        dtUploadDetail.CurrCls = "03"
                                    ElseIf tempCls = "SGD" Then
                                        dtUploadDetail.CurrCls = "04"
                                    ElseIf tempCls = "EUR" Then
                                        dtUploadDetail.CurrCls = "05"
                                    End If

                                    dtUploadDetail.Price = IIf(IsDBNull(dtDetail.Rows(i).Item(5)), 0, dtDetail.Rows(i).Item(5))
                                    dtUploadDetail.StartDate = IIf(IsDBNull(dtDetail.Rows(i).Item(6)), "NULL", dtDetail.Rows(i).Item(6)) 'dtDetail.Rows(i).Item(4)
                                    dtUploadDetail.EndDate = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), "NULL", dtDetail.Rows(i).Item(7)) 'dtDetail.Rows(i).Item(5)
                                    dtUploadDetail.EffectiveDate = IIf(IsDBNull(dtDetail.Rows(i).Item(8)), "NULL", dtDetail.Rows(i).Item(8)) 'dtDetail.Rows(i).Item(5)
                                    dtUploadDetailList.Add(dtUploadDetail)
                                End If
                            Next
                        End If

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            ''01.01 Delete TempoaryData
                            ls_sql = "delete UploadPrice"
                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            ''02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim PO As clsMaster = dtUploadDetailList(i)

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Parts WHERE PartNo = '" & PO.PartNo & "' "
                                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                                Dim ds2 As New DataSet
                                sqlDA2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count = 0 Then
                                    ls_error = "Part No not found in Part Master, please check again."
                                End If

                                '02.1 Check PartNo di MS_Part
                                ls_sql = "SELECT * FROM dbo.MS_Affiliate WHERE AffiliateID = '" & PO.AffiliateID & "' "
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)

                                If ds3.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = ls_error & "Affiliate ID not found in Affiliate Master, please check again."
                                    Else
                                        ls_error = ls_error & "; " & "Affiliate ID not found in Affiliate Master, please check again."
                                    End If
                                End If

                                If PO.StartDate <> "NULL" Then
                                    If IsDate(PO.StartDate) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If
                                
                                If PO.EndDate <> "NULL" Then
                                    If IsDate(PO.EndDate) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                If PO.EffectiveDate <> "NULL" Then
                                    If IsDate(PO.EffectiveDate) = False Then
                                        If ls_error = "" Then
                                            ls_error = ls_error & "Invalid format date, please check again"
                                        Else
                                            ls_error = ls_error & "; " & "Invalid format date, please check again"
                                        End If
                                    End If
                                End If

                                If IsNumeric(PO.Price) = False Then
                                    If ls_error = "" Then
                                        ls_error = ls_error & "Invalid price, please check again"
                                    Else
                                        ls_error = ls_error & "; " & "Invalid price, please check again"
                                    End If
                                End If

                                ls_sql = " INSERT INTO [dbo].[UploadPrice] " & vbCrLf & _
                                          "            ([AffiliateID], [PartNo], [CurrCls], [Price], [PackingCls], [PriceCls], [StartDate], [EndDate], [EntryDate], [ErrorCls]) " & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & PO.AffiliateID & "'" & vbCrLf & _
                                          "            ,'" & PO.PartNo & "' " & vbCrLf & _
                                          "            ,'" & PO.CurrCls & "' " & vbCrLf & _
                                          "            ,'" & PO.Price & "' " & vbCrLf & _
                                          "            ,'" & PO.PackingCls & "' " & vbCrLf & _
                                          "            ," & PO.PriceCategory & " " & vbCrLf & _
                                          "            ," & IIf(PO.StartDate = "NULL", PO.StartDate, "'" & PO.StartDate & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.EndDate = "NULL", PO.EndDate, "'" & PO.EndDate & "'") & "" & vbCrLf & _
                                          "            ," & IIf(PO.EffectiveDate = "NULL", PO.EffectiveDate, "'" & PO.EffectiveDate & "'") & "" & vbCrLf & _
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
        Dim ls_Sql, ls_PriceCategory, ls_PackingType As String
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
                Dim tempCurrCls As String, tempPackingCls As String, tempPriceCategory As String

                For i = 0 To grid.VisibleRowCount - 1
                    If grid.GetRowValues(i, "AffiliateID") <> "" Then
                        Dim tempCls As String = grid.GetRowValues(i, "CurrCls")
                        If tempCls.Trim = "JPY" Then
                            tempCurrCls = "01"
                        ElseIf tempCls.Trim = "USD" Then
                            tempCurrCls = "02"
                        ElseIf tempCls.Trim = "IDR" Then
                            tempCurrCls = "03"
                        ElseIf tempCls.Trim = "SGD" Then
                            tempCurrCls = "04"
                        ElseIf tempCls.Trim = "EUR" Then
                            tempCurrCls = "05"
                        End If

                        If Not IsDBNull(grid.GetRowValues(i, "PackingCls")) Then
                            If Trim(grid.GetRowValues(i, "PackingCls")).ToUpper = "CARTON BOX" Then
                                tempPackingCls = "01"
                            ElseIf Trim(grid.GetRowValues(i, "PackingCls")).ToUpper = "IMPRABOARD" Then
                                tempPackingCls = "02"
                            ElseIf Trim(grid.GetRowValues(i, "PackingCls")).ToUpper = "POLYBOX" Then
                                tempPackingCls = "03"
                            ElseIf Trim(grid.GetRowValues(i, "PackingCls")).ToUpper = "KMT 40" Then
                                tempPackingCls = "04"
                            End If
                        Else
                            tempPackingCls = ""
                        End If

                        If Not IsDBNull(grid.GetRowValues(i, "PriceCls")) Then
                            If Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "FCA - AIR" Then
                                tempPriceCategory = "1"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "FCA - BOAT" Then
                                tempPriceCategory = "2"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "CIF - AIR" Then
                                tempPriceCategory = "3"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "CIF - BOAT" Then
                                tempPriceCategory = "4"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "DDU PASI" Then
                                tempPriceCategory = "5"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "DDU AFFILIATE" Then
                                tempPriceCategory = "6"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "EX-WORK" Then
                                tempPriceCategory = "7"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "FCA" Then
                                tempPriceCategory = "9"
                            ElseIf Trim(grid.GetRowValues(i, "PriceCls")).ToUpper = "CIF" Then
                                tempPriceCategory = "10"
                            End If
                        Else
                            tempPriceCategory = ""
                        End If

                        ls_Sql = " IF NOT EXISTS (select * from MS_Price where AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "' and CurrCls = '" & tempCurrCls & "' and StartDate = '" & grid.GetRowValues(i, "StartDate") & "' and PackingCls in ('" & tempPackingCls & "','') and PriceCls = '" & tempPriceCategory & "')" & vbCrLf & _
                                  " BEGIN" & vbCrLf & _
                                  " INSERT INTO [dbo].[MS_Price] " & vbCrLf & _
                                  "            ([AffiliateID] " & vbCrLf & _
                                  "            ,[PartNo] " & vbCrLf & _
                                  "            ,[CurrCls] " & vbCrLf & _
                                  "            ,[Price] " & vbCrLf & _
                                  "            ,[PackingCls] " & vbCrLf & _
                                  "            ,[PriceCls] " & vbCrLf & _
                                  "            ,[StartDate] " & vbCrLf & _
                                  "            ,[EndDate] " & vbCrLf & _
                                  "            ,[EffectiveDate] " & vbCrLf & _
                                  "            ,[EntryDate] " & vbCrLf & _
                                  "            ,[EntryUser] ) " & vbCrLf & _
                                  "      VALUES " & vbCrLf & _
                                  "            ('" & grid.GetRowValues(i, "AffiliateID") & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "PartNo") & "' " & vbCrLf & _
                                  "            ,'" & tempCurrCls & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "Price") & "' " & vbCrLf & _
                                  "            ,'" & tempPackingCls & "' " & vbCrLf & _
                                  "            ,'" & tempPriceCategory & "' " & vbCrLf & _
                                  "            ,'" & grid.GetRowValues(i, "StartDate") & "' "

                        'ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "EndDate") & "' " & vbCrLf & _
                        '                  "            ,'" & grid.GetRowValues(i, "EntryDate") & "' " & vbCrLf & _
                        '                  "            ,getdate() " & vbCrLf & _
                        '                  "            ,'UPLOAD') " & vbCrLf & _
                        '                  " END " & vbCrLf & _
                        '                  " ELSE " & vbCrLf & _
                        '                  " BEGIN" & vbCrLf & _
                        '                  "      IF EXISTS (select * from MS_Price where AffiliateID = '" & grid.GetRowValues(i, "SupplierID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "' and CurrCls = '" & tempCurrCls & "' and StartDate = '" & grid.GetRowValues(i, "StartDate") & "' and PackingCls in (''))" & vbCrLf & _
                        '                  "      BEGIN " & vbCrLf & _
                        '                  "      UPDATE [dbo].[MS_Price] SET " & vbCrLf & _
                        '                  "       [Price] = '" & grid.GetRowValues(i, "Price") & "' " & vbCrLf & _
                        '                  "       ,[PackingCls] = '" & tempPackingCls & "' " & vbCrLf & _
                        '                  "       ,[PriceCls] = '" & tempPriceCategory & "' " & vbCrLf & _
                        '                  "       ,[EndDate] = '" & grid.GetRowValues(i, "EndDate") & "' " & vbCrLf & _
                        '                  "       ,[EffectiveDate] = '" & grid.GetRowValues(i, "EntryDate") & "' " & vbCrLf & _
                        '                  "      WHERE AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "' and CurrCls = '" & tempCurrCls & "' and StartDate = '" & grid.GetRowValues(i, "StartDate") & "' " & vbCrLf & _
                        '                  "      END " & vbCrLf & _
                        '                  "      ELSE " & vbCrLf & _
                        '                  "      BEGIN " & vbCrLf & _
                        '                  "      UPDATE [dbo].[MS_Price] SET " & vbCrLf & _
                        '                  "       [Price] = '" & grid.GetRowValues(i, "Price") & "' " & vbCrLf & _
                        '                  "       ,[PriceCls] = '" & tempPriceCategory & "' " & vbCrLf & _
                        '                  "       ,[EndDate] = '" & grid.GetRowValues(i, "EndDate") & "' " & vbCrLf & _
                        '                  "       ,[EffectiveDate] = '" & grid.GetRowValues(i, "EntryDate") & "' " & vbCrLf & _
                        '                  "      WHERE AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "' and CurrCls = '" & tempCurrCls & "' and StartDate = '" & grid.GetRowValues(i, "StartDate") & "' and PackingCls = '" & tempPackingCls & "' " & vbCrLf & _
                        '                  "      END " & vbCrLf & _
                        '                  " END"

                        ls_Sql = ls_Sql + "            ,'" & grid.GetRowValues(i, "EndDate") & "' " & vbCrLf & _
                                          "            ,'" & grid.GetRowValues(i, "EntryDate") & "' " & vbCrLf & _
                                          "            ,getdate() " & vbCrLf & _
                                          "            ,'" & Session("UserID") & "') " & vbCrLf & _
                                          " END " & vbCrLf & _
                                          " ELSE " & vbCrLf & _
                                          " BEGIN" & vbCrLf & _
                                          "      UPDATE [dbo].[MS_Price] SET " & vbCrLf & _
                                          "       [Price] = '" & grid.GetRowValues(i, "Price") & "' " & vbCrLf & _
                                          "       ,[PriceCls] = '" & tempPriceCategory & "' " & vbCrLf & _
                                          "       ,[EndDate] = '" & grid.GetRowValues(i, "EndDate") & "' " & vbCrLf & _
                                          "       ,[EffectiveDate] = '" & grid.GetRowValues(i, "EntryDate") & "' " & vbCrLf & _
                                          "       ,[UpdateDate] = GETDATE() " & vbCrLf & _
                                          "       ,[UpdateUser] = '" & Session("UserID") & "' " & vbCrLf & _
                                          "      WHERE AffiliateID = '" & grid.GetRowValues(i, "AffiliateID") & "' and PartNo = '" & grid.GetRowValues(i, "PartNo") & "' and CurrCls = '" & tempCurrCls & "' and StartDate = '" & grid.GetRowValues(i, "StartDate") & "' and PackingCls = '" & tempPackingCls & "' and PriceCls = '" & tempPriceCategory & "'" & vbCrLf & _
                                          " END"

                        SQLCom.CommandText = ls_Sql
                        SQLCom.ExecuteNonQuery()
                    End If

                    'ls_PriceCategory = grid.GetRowValues(i, "xpricecategory")
                    'ls_PriceCategory = grid.GetRowValues(i, "PriceCls")
                    'ls_PackingType = grid.GetRowValues(i, "xpackingtype")
                    'ls_PackingType = grid.GetRowValues(i, "PackingCls")

                    PartNo = grid.GetRowValues(i, "PartNo")

                    'If Trim(grid.GetRowValues(i, "AffiliateID")) <> Trim(grid.GetRowValues(i, "xaffiliateid")) Then
                    '    ls_Remarks = ls_Remarks + "AffiliateID " + Trim(grid.GetRowValues(i, "xaffiliateid")) & "->" & Trim(grid.GetRowValues(i, "AffiliateID")) & " "
                    'End If

                    If Not IsDBNull(grid.GetRowValues(i, "PackingCls").ToString) And Not IsDBNull(grid.GetRowValues(i, "xpackingtype").ToString) Then
                        If Trim(grid.GetRowValues(i, "PackingCls").ToString) <> Trim(grid.GetRowValues(i, "xpackingtype").ToString) Then
                            ls_Remarks = ls_Remarks + "PackingCls " + Trim(grid.GetRowValues(i, "xpackingtype").ToString) & "->" & Trim(grid.GetRowValues(i, "PackingCls").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "PriceCls")) And Not IsDBNull(grid.GetRowValues(i, "xpricecategory")) Then
                        If Trim(grid.GetRowValues(i, "PriceCls").ToString) <> Trim(grid.GetRowValues(i, "xpricecategory").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "PriceCls " + Trim(grid.GetRowValues(i, "xpricecategory").ToString) & "->" & Trim(grid.GetRowValues(i, "PriceCls").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "Price")) And Not IsDBNull(grid.GetRowValues(i, "xprice")) Then
                        If Trim(grid.GetRowValues(i, "Price").ToString) <> Trim(grid.GetRowValues(i, "xprice").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "Price " + Trim(grid.GetRowValues(i, "xprice").ToString) & "->" & Trim(grid.GetRowValues(i, "Price").ToString) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "EndDate")) And Not IsDBNull(grid.GetRowValues(i, "xenddate")) Then
                        If Trim(grid.GetRowValues(i, "EndDate").ToString) <> Trim(grid.GetRowValues(i, "xenddate").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "EndDate " + Trim(Format(grid.GetRowValues(i, "xenddate"), "dd-MMM-yyyy")) & "->" & Trim(Format(grid.GetRowValues(i, "EndDate"), "dd-MMM-yyyy")) & ""
                        End If
                    End If

                    If Not IsDBNull(grid.GetRowValues(i, "EntryDate")) And Not IsDBNull(grid.GetRowValues(i, "xeffectivedate")) Then
                        If Trim(grid.GetRowValues(i, "EntryDate").ToString) <> Trim(grid.GetRowValues(i, "xeffectivedate").ToString) Then
                            If ls_Remarks <> "" Then ls_Remarks = ls_Remarks + ", "
                            ls_Remarks = ls_Remarks + "EffectiveDate " + Trim(Format(grid.GetRowValues(i, "xeffectivedate"), "dd-MMM-yyyy")) & "->" & Trim(Format(grid.GetRowValues(i, "EntryDate"), "dd-MMM-yyyy")) & ""
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
                ls_Sql = "delete UploadPrice "

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