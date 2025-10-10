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

Public Class DeliveryExportUpload
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
        Response.Redirect("~/DeliveryExport/DeliveryToAffListExport.aspx")
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

            ls_SQL = "  SELECT NoUrut = ROW_NUMBER() OVER ( ORDER BY partno ) , " & vbCrLf & _
                  "         * " & vbCrLf & _
                  "  FROM   ( SELECT DISTINCT  " & vbCrLf & _
                  "                     partno = ISNULL(DOP.PartNo,'') , " & vbCrLf & _
                  "                     partname = ISNULL(MP.PartName,'') , " & vbCrLf & _
                  "                     labelno = '' , " & vbCrLf & _
                  "                     uom = ISNULL(UC.DESCRIPTION,'') , " & vbCrLf & _
                  "                     qtybox = ISNULL(MP.Qtybox,0) , " & vbCrLf & _
                  "                     boxpalet = ISNULL(boxpallet,0) , " & vbCrLf & _
                  "                     deliveryplanqty = CASE DOP.Week " & vbCrLf & _
                  "                                         WHEN '1' THEN POD.week1 "

            ls_SQL = ls_SQL + "                                         WHEN '2' THEN POD.week2 " & vbCrLf & _
                              "                                         WHEN '3' THEN POD.week3 " & vbCrLf & _
                              "                                         WHEN '4' THEN POD.week4 " & vbCrLf & _
                              "                                         WHEN '5' THEN POD.week5 " & vbCrLf & _
                              "                                       END , " & vbCrLf & _
                              "                     remainingqty = ( CASE DOP.Week " & vbCrLf & _
                              "                                        WHEN '1' THEN POD.week1 " & vbCrLf & _
                              "                                        WHEN '2' THEN POD.week2 " & vbCrLf & _
                              "                                        WHEN '3' THEN POD.week3 " & vbCrLf & _
                              "                                        WHEN '4' THEN POD.week4 " & vbCrLf & _
                              "                                        WHEN '5' THEN POD.week5 "

            ls_SQL = ls_SQL + "                                      END ) - DOP.DOQty , " & vbCrLf & _
                              "                     deliveryqty = ISNULL(DOP.DOQty,0) , " & vbCrLf & _
                              "                     deliveryqtybox = ISNULL(DOP.DOQty / MP.QtyBox,0) , " & vbCrLf & _
                              "                     deliveryqtypallet = ISNULL(DOP.DOQty / MP.BoxPallet,0), " & vbCrLf & _
                              "                     remarks = isnull(error,'') " & vbCrLf & _
                              "           FROM      DOSupplier_Upload DOP " & vbCrLf & _
                              "                     LEFT JOIN PO_Master_Export POM ON POM.PONo = DOP.POno " & vbCrLf & _
                              "                     LEFT JOIN PO_Detail_Export POD ON POD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              "                                                       AND POD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "                                                       AND POD.PONo = POM.PONo " & vbCrLf & _
                              "                                                       AND DOP.PartNo = POD.Partno " & vbCrLf & _
                              "                     LEFT JOIN MS_Parts MP ON DOP.Partno = MP.Partno " & vbCrLf & _
                              "                     LEFT JOIN MS_UnitCls UC ON UC.unitcls = MP.UnitCls "

            ls_SQL = ls_SQL + "         ) x "


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

            ls_SQL = " select top 0 '' NoUrut, '' Period, '' PONo, '' ShipBy, '' PartNo, '' PartName, '' UnitDesc, '' MOQ, '' Maker, " & vbCrLf & _
                  " '' Project, '' SupplierID, 0 POQty, 0 ForecastN1, 0 ForecastN2, 0 ForecastN3,   " & vbCrLf & _
                  " 0 DeliveryD1, 0 DeliveryD2, 0 DeliveryD3, 0 DeliveryD4, 0 DeliveryD5, " & vbCrLf & _
                  " 0 DeliveryD6, 0 DeliveryD7, 0 DeliveryD8, 0 DeliveryD9, 0 DeliveryD10, " & vbCrLf & _
                  " 0 DeliveryD11, 0 DeliveryD12, 0 DeliveryD13, 0 DeliveryD14, 0 DeliveryD15, " & vbCrLf & _
                  " 0 DeliveryD16, 0 DeliveryD17, 0 DeliveryD18, 0 DeliveryD19, 0 DeliveryD20, " & vbCrLf & _
                  " 0 DeliveryD21, 0 DeliveryD22, 0 DeliveryD23, 0 DeliveryD24, 0 DeliveryD25, " & vbCrLf & _
                  " 0 DeliveryD26, 0 DeliveryD27, 0 DeliveryD28, 0 DeliveryD29, 0 DeliveryD30, " & vbCrLf & _
                  " 0 DeliveryD31, '' ErrorCls"

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
                    Case ".xlsm"
                        'Excel xlsm
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

                Dim dtUploadHeader As New clsDeliveryExportHeader
                Dim dtUploadHeaderList As New List(Of clsDeliveryExportHeader)

                For Each drSheet In dtSheets.Rows
                    listSheet.Add(drSheet("TABLE_NAME").ToString())
                Next

                Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                    sqlConn.Open()

                    ''==========Table EXCEL Master==========
                    Dim pTableCode As String = listSheet(0)

                    Try
                        Session.Remove("suratjalanno")
                        Session.Remove("supplier")
                        Session.Remove("affiliateID")
                        Session.Remove("orderno")
                        Session.Remove("pic")
                        Session.Remove("jenisarmada")
                        Session.Remove("drivername")
                        Session.Remove("drivercontact")
                        Session.Remove("nopol")
                        Session.Remove("totalbox")
                        Session.Remove("totalpallet")

                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "H3:AV36]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dt)

                        If dt.Rows.Count > 0 Then

                            'setting header
                            dtUploadHeader.H_suratjalanno = ""
                            dtUploadHeader.H_supplier = ""
                            dtUploadHeader.H_affilaiteid = ""
                            dtUploadHeader.H_orderno = ""
                            dtUploadHeader.H_pono = ""
                            dtUploadHeader.H_deliverydate = ""
                            dtUploadHeader.H_pic = ""
                            dtUploadHeader.H_jenisarmada = ""
                            dtUploadHeader.H_drivername = ""
                            dtUploadHeader.H_drivercontact = ""
                            dtUploadHeader.H_nopol = ""
                            dtUploadHeader.H_WEEK = 0
                            'setting header

                            'SuratJalanNo
                            If IsDBNull(dt.Rows(27).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid column ""Surat Jalan No."", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_suratjalanno = Trim(dt.Rows(27).Item(2))

                            'Supplier
                            If IsDBNull(dt.Rows(8).Item(1)) Then
                                lblInfo.Text = "[9999] Invalid column ""Supplier"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_supplier = Trim(dt.Rows(8).Item(1))

                            'AffiliateID
                            If IsDBNull(dt.Rows(0).Item(0)) Then
                                lblInfo.Text = "[9999] Invalid column ""Affiliate ID"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_affilaiteid = Trim(dt.Rows(0).Item(0))

                            'OrderNo
                            If IsDBNull(dt.Rows(12).Item(23)) Then
                                lblInfo.Text = "[9999] Invalid column ""Order No"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_orderno = Trim(dt.Rows(12).Item(23))

                            'PIC
                            If IsDBNull(dt.Rows(25).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid colum ""PIC"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_pic = Trim(dt.Rows(25).Item(2))

                            'Jenis Armada
                            If IsDBNull(dt.Rows(27).Item(11)) Then
                                lblInfo.Text = "[9999] Invalid colum ""Jenis Armada"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_jenisarmada = Trim(dt.Rows(27).Item(11))

                            'Driver Name
                            If IsDBNull(dt.Rows(29).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid colum ""Driver Name"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_drivername = Trim(dt.Rows(29).Item(2))

                            'Driver Contact
                            If IsDBNull(dt.Rows(31).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid colum ""Driver Contact"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_drivercontact = Trim(dt.Rows(31).Item(2))

                            'Nopol
                            If IsDBNull(dt.Rows(33).Item(2)) Then
                                lblInfo.Text = "[9999] Invalid colum ""No Polisi"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_nopol = Trim(dt.Rows(33).Item(2))

                            'Total Box
                            If IsDBNull(dt.Rows(29).Item(16)) Then
                                lblInfo.Text = "[9999] Invalid colum ""total Box"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_totalbox = Trim(dt.Rows(29).Item(16))

                            'Total Pallet
                            If IsDBNull(dt.Rows(31).Item(16)) Then
                                lblInfo.Text = "[9999] Invalid colum ""total Pallet"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_totalpalet = Trim(dt.Rows(31).Item(16))

                            'WEEEK
                            If IsDBNull(dt.Rows(10).Item(23)) Then
                                lblInfo.Text = "[9999] Invalid colum ""WEEK"", please check the file again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If
                            dtUploadHeader.H_WEEK = Trim(dt.Rows(10).Item(23))

                            'AffiliateID
                            ls_sql = "select AffiliateID from MS_Affiliate where affiliateID = '" & Trim(dtUploadHeader.H_affilaiteid) & "' "

                            Dim sqlCmd As New SqlCommand(ls_sql, sqlConn)
                            Dim sqlDA As New SqlDataAdapter(sqlCmd)
                            Dim ds As New DataSet
                            sqlDA.Fill(ds)

                            If ds.Tables(0).Rows.Count > 0 Then

                            Else
                                lblInfo.Text = "[9999] Invalid AffiliateID, please Check Again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                            'SupplierID
                            ls_sql = "select SupplierID from MS_Supplier where SupplierID = '" & Trim(dtUploadHeader.H_supplier) & "' "

                            Dim sqlCmd1 As New SqlCommand(ls_sql, sqlConn)
                            Dim sqlDA1 As New SqlDataAdapter(sqlCmd1)
                            Dim ds1 As New DataSet
                            sqlDA.Fill(ds1)

                            If ds1.Tables(0).Rows.Count > 0 Then

                            Else
                                lblInfo.Text = "[9999] Invalid SupplierID, please Check Again!"
                                grid.JSProperties("cpMessage") = lblInfo.Text
                                Exit Sub
                            End If

                        End If

                        Session("suratjalanno") = dtUploadHeader.H_suratjalanno
                        Session("supplier") = dtUploadHeader.H_supplier
                        Session("affiliateID") = dtUploadHeader.H_affilaiteid
                        Session("orderno") = dtUploadHeader.H_orderno
                        Session("pic") = dtUploadHeader.H_pic
                        Session("jenisarmada") = dtUploadHeader.H_jenisarmada
                        Session("drivername") = dtUploadHeader.H_drivername
                        Session("drivercontact") = dtUploadHeader.H_drivercontact
                        Session("nopol") = dtUploadHeader.H_nopol
                        Session("totalbox") = dtUploadHeader.H_totalbox
                        Session("totalpallet") = dtUploadHeader.H_totalpalet

                        'Dim dtUploadDetail As New clsPODetail
                        Dim dtUploadDetailList As New List(Of clsDeliveryExportDetail)

                        'Get Detail Data
                        MyCommand.CommandText = ("SELECT * FROM [" & pTableCode & "B42:AO65536]")
                        MyAdapter.SelectCommand = MyCommand
                        MyAdapter.Fill(dtDetail)

                        If dtDetail.Rows.Count > 0 Then
                            For i = 0 To dtDetail.Rows.Count - 1
                                If IIf(IsDBNull(dtDetail.Rows(i).Item(0)), "", dtDetail.Rows(i).Item(0)) <> "E" Then
                                    Dim dtUploadDetail As New clsDeliveryExportDetail

                                    dtUploadDetail.D_partno = IIf(IsDBNull(dtDetail.Rows(i).Item(2)), 0, dtDetail.Rows(i).Item(2))
                                    dtUploadDetail.D_partname = IIf(IsDBNull(dtDetail.Rows(i).Item(7)), 0, dtDetail.Rows(i).Item(7))
                                    dtUploadDetail.D_labelno = IIf(IsDBNull(dtDetail.Rows(i).Item(16)), 0, dtDetail.Rows(i).Item(16))
                                    dtUploadDetail.D_uom = IIf(IsDBNull(dtDetail.Rows(i).Item(20)), 0, dtDetail.Rows(i).Item(20))
                                    dtUploadDetail.D_qtybox = IIf(IsDBNull(dtDetail.Rows(i).Item(22)), 0, dtDetail.Rows(i).Item(22))
                                    dtUploadDetail.D_boxpallet = IIf(IsDBNull(dtDetail.Rows(i).Item(24)), 0, dtDetail.Rows(i).Item(24))
                                    dtUploadDetail.D_deliveryplanqty = IIf(IsDBNull(dtDetail.Rows(i).Item(27)), 0, dtDetail.Rows(i).Item(27))
                                    dtUploadDetail.D_remainingqty = IIf(IsDBNull(dtDetail.Rows(i).Item(31)), 0, dtDetail.Rows(i).Item(31)) - IIf(IsDBNull(dtDetail.Rows(i).Item(27)), 0, dtDetail.Rows(i).Item(27))
                                    dtUploadDetail.D_deliveryqty = IIf(IsDBNull(dtDetail.Rows(i).Item(31)), 0, dtDetail.Rows(i).Item(31))
                                    dtUploadDetail.D_deliveryqtybox = IIf(IsDBNull(dtDetail.Rows(i).Item(35)), 0, dtDetail.Rows(i).Item(35))
                                    dtUploadDetail.D_deliveryqtypallet = IIf(IsDBNull(dtDetail.Rows(i).Item(39)), 0, dtDetail.Rows(i).Item(39))

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
                        Dim ls_POno As String = ""

                        Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                            '01. Check Kanban already Exists
                            ls_sql = "SELECT * FROM DOSupplier_Master_Export WHERE suratjalanno = '" & dtUploadHeader.H_suratjalanno & "' " & vbCrLf & _
                                     " and AffiliateID = '" & dtUploadHeader.H_affilaiteid & "'" & vbCrLf & _
                                     " and SupplierID = '" & dtUploadHeader.H_supplier & "' " & vbCrLf & _
                                     " and OrderNo = '" & dtUploadHeader.H_orderno & "' "

                            Dim sqlCmd As New SqlCommand(ls_sql, sqlConn, sqlTran)
                            Dim sqlDA As New SqlDataAdapter(sqlCmd)
                            Dim ds As New DataSet
                            sqlDA.Fill(ds)

                            If ds.Tables(0).Rows.Count > 0 Then
                                'If Not IsDBNull(ds.Tables(0).Rows(0)("KanbanStatus")) Then
                                '    Call clsMsg.DisplayMessage(lblInfo, "5012", clsMessage.MsgType.ErrorMessage)
                                '    Exit Sub
                                'End If
                            End If

                            '01.01 Delete TempoaryData
                            ls_sql = "delete [DOSupplier_UPLOAD] WHERE suratjalanno = '" & dtUploadHeader.H_suratjalanno & "' " & vbCrLf & _
                                     " and AffiliateID = '" & dtUploadHeader.H_affilaiteid & "'" & vbCrLf & _
                                     " and SupplierID = '" & dtUploadHeader.H_supplier & "' " & vbCrLf & _
                                     " and OrderNo = '" & dtUploadHeader.H_orderno & "' "

                            Dim sqlComm9 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                            sqlComm9.ExecuteNonQuery()
                            sqlComm9.Dispose()


                            '02. Check PartNo di Part Master MS_Parts dan Check Item di Part Mapping MS_PartMapping
                            For i = 0 To dtUploadDetailList.Count - 1
                                Dim ls_error As String = ""
                                Dim Delivery As clsDeliveryExportDetail = dtUploadDetailList(i)
                                Dim ls_Qty As Integer
                                Dim ls_FieldOrderNo As String

                                '1. PONO

                                If dtUploadHeader.H_WEEK = "1" Then ls_FieldOrderNo = "OrderNo1"
                                If dtUploadHeader.H_WEEK = "2" Then ls_FieldOrderNo = "OrderNo2"
                                If dtUploadHeader.H_WEEK = "3" Then ls_FieldOrderNo = "OrderNo3"
                                If dtUploadHeader.H_WEEK = "4" Then ls_FieldOrderNo = "OrderNo4"

                                ls_sql = "select * from PO_Master_Export WHERE SupplierID = '" & dtUploadHeader.H_supplier & "' and AffiliateID = '" & Session("affiliateID") & "'" & vbCrLf & _
                                         " and " & ls_FieldOrderNo & " = '" & dtUploadHeader.H_orderno & "' "
                                Dim sqlCmd3 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA3 As New SqlDataAdapter(sqlCmd3)
                                Dim ds3 As New DataSet
                                sqlDA3.Fill(ds3)

                                If ds3.Tables(0).Rows.Count = 0 Then
                                    If ls_error = "" Then
                                        ls_error = "PO No. not found in PO Master, please check again !"
                                    End If
                                Else
                                    ls_POno = ds3.Tables(0).Rows(0)("PONo")
                                    Session("POno") = ls_POno
                                End If

                                '02.3 Check PartNo di MS_Part
                                ls_sql = "Select * from [DOSupplier_UPLOAD] WHERE suratjalanno = '" & dtUploadHeader.H_suratjalanno & "' " & vbCrLf & _
                                         " and AffiliateID = '" & dtUploadHeader.H_affilaiteid & "'" & vbCrLf & _
                                         " and SupplierID = '" & dtUploadHeader.H_supplier & "' " & vbCrLf & _
                                         " and OrderNo = '" & dtUploadHeader.H_orderno & "' " & vbCrLf & _
                                         " and PartNo = '" & Delivery.D_partno & "'" & vbCrLf & _
                                         " and POno = '" & ls_POno & "'"
                                Dim sqlCmd4 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                                Dim sqlDA4 As New SqlDataAdapter(sqlCmd4)
                                Dim ds4 As New DataSet
                                sqlDA4.Fill(ds4)

                                If ds4.Tables(0).Rows.Count > 0 Then
                                    ls_sql = "delete [DOSupplier_UPLOAD] WHERE suratjalanno = '" & dtUploadHeader.H_suratjalanno & "' " & vbCrLf & _
                                             " and AffiliateID = '" & dtUploadHeader.H_affilaiteid & "'" & vbCrLf & _
                                             " and SupplierID = '" & dtUploadHeader.H_supplier & "' " & vbCrLf & _
                                             " and OrderNo = '" & dtUploadHeader.H_orderno & "' " & vbCrLf & _
                                             " and PartNo = '" & Delivery.D_partno & "'" & vbCrLf & _
                                             " and POno = '" & ls_POno & "'"
                                    Dim sqlComm1 = New SqlCommand(ls_sql, sqlConn, sqlTran)
                                    sqlComm1.ExecuteNonQuery()
                                    sqlComm1.Dispose()
                                End If

                                ls_sql = " INSERT INTO [dbo].[DOSupplier_UPLOAD] " & vbCrLf & _
                                          "            (SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, DoQty, Error, week)" & vbCrLf & _
                                          "      VALUES " & vbCrLf & _
                                          "            ('" & dtUploadHeader.H_suratjalanno & "' " & vbCrLf & _
                                          "            ,'" & dtUploadHeader.H_supplier & "' " & vbCrLf & _
                                          "            ,'" & dtUploadHeader.H_affilaiteid & "' " & vbCrLf

                                ls_sql = ls_sql + "            ,'" & ls_POno & "' " & vbCrLf & _
                                                  "            ,'" & Delivery.D_partno & "' " & vbCrLf & _
                                                  "            ,'" & dtUploadHeader.H_orderno & "' " & vbCrLf & _
                                                  "            , '" & Delivery.D_deliveryqty & "' " & vbCrLf & _
                                                  "            , '" & ls_error & "' " & vbCrLf & _
                                                  "            , '" & dtUploadHeader.H_WEEK & "') " & vbCrLf
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
        Dim ls_KanbanNo As String
        Dim ls_cycle As Integer
        Dim ls_Date As String
        Dim ls_time As String
        Dim ls_seq As Integer

        Try
            '01. Cari ada data yg disubmit
            For i = 0 To grid.VisibleRowCount - 1
                If grid.GetRowValues(i, "remarks").ToString.Trim <> "" Then
                    ls_Check = True
                    Exit For
                End If
            Next i

            Dim countSupplier As Integer = 0

            'For i = 0 To grid.VisibleRowCount - 1
            '    If i = 0 Then
            '        ls_TempSupplierID = grid.GetRowValues(i, "supplier").ToString.Trim
            '        countSupplier = 1
            '    End If

            '    If ls_TempSupplierID <> grid.GetRowValues(i, "supplier").ToString.Trim Then
            '        ls_DoubleSupplier = True
            '        ls_TempSupplierID = grid.GetRowValues(i, "supplier").ToString.Trim
            '        countSupplier = countSupplier + 1
            '    End If
            'Next i

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
                'dian
                '1. Delete Data
                Dim SQLCom As SqlCommand = SqlCon.CreateCommand
                SQLCom.Connection = SqlCon
                SQLCom.Transaction = SqlTran
                Dim ls_KanbanAsli As String = Trim(Session("PONoUpload"))
                Dim ls_deliveryDate As String = Session("DeliveryDate")

                ls_Sql = "Delete DOSupplier_Detail_Export where suratjalanno = '" & Session("suratjalanno") & "' and supplierID = '" & Session("supplier") & "' " & vbCrLf & _
                         " and AffiliateID = '" & Session("affiliateID") & "' and OrderNo = '" & Session("orderno") & "' " & vbCrLf & _
                         " and POno = '" & Session("POno") & "' "
                ls_Sql = ls_Sql + "Delete DOSupplier_Master_Export where suratjalanno = '" & Session("suratjalanno") & "' and supplierID = '" & Session("supplier") & "' " & vbCrLf & _
                         " and AffiliateID = '" & Session("affiliateID") & "' and OrderNo = '" & Session("orderno") & "' " & vbCrLf & _
                         " and POno = '" & Session("POno") & "' "
                SQLCom.CommandText = ls_Sql
                SQLCom.ExecuteNonQuery()

                '2. Insert New Data
                For i = 0 To grid.VisibleRowCount - 1
                    ls_Sql = " IF NOT EXISTS (select * From DOSupplier_Master_Export where suratjalanno = '" & Session("suratjalanno") & "' and supplierID = '" & Session("supplier") & "' " & vbCrLf & _
                             " and AffiliateID = '" & Session("affiliateID") & "' and OrderNo = '" & Session("orderno") & "' " & vbCrLf & _
                             " and POno = '" & Session("POno") & "') BEGIN " & vbCrLf & _
                             " Insert Into DOSupplier_Master_Export Values( " & vbCrLf & _
                             " '" & Session("suratjalanno") & "', " & vbCrLf & _
                             " '" & Session("supplier") & "', " & vbCrLf & _
                             " '" & Session("AffiliateID") & "', " & vbCrLf & _
                             " '" & Session("pono") & "', " & vbCrLf & _
                             " '" & Session("orderno") & "', " & vbCrLf & _
                             " Getdate(), " & vbCrLf & _
                             " '" & Session("pic") & "', " & vbCrLf & _
                             " '" & Session("jenisarmada") & "', " & vbCrLf & _
                             " '" & Session("DriverName") & "', " & vbCrLf & _
                             " '" & Session("DriverContact") & "' , " & vbCrLf & _
                             " '" & Session("nopol") & "', " & vbCrLf & _
                             " " & Session("totalbox") & ", " & vbCrLf & _
                             " getdate(), " & vbCrLf & _
                             " '" & Session("UserID").ToString & "', " & vbCrLf & _
                             " getdate(), " & vbCrLf & _
                             " '" & Session("UserID").ToString & "') END" & vbCrLf
                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()

                    ls_Sql = "Insert into DOSupplier_Detail_Export values ( " & vbCrLf & _
                             " '" & Session("suratjalanno") & "', " & vbCrLf & _
                             " '" & Session("supplier") & "', " & vbCrLf & _
                             " '" & Session("AffiliateID") & "', " & vbCrLf & _
                             " '" & Session("pono") & "', " & vbCrLf & _
                             " '" & grid.GetRowValues(i, "partno") & "', " & vbCrLf & _
                             " '" & Session("orderno") & "', " & vbCrLf & _
                             "  " & grid.GetRowValues(i, "deliveryqty") & " )"

                    SQLCom.CommandText = ls_Sql
                    SQLCom.ExecuteNonQuery()

                    ls_MsgID = "1001"
                    ls_Detail = "ada"
                Next

                'delete tada di tempolary table
                ls_Sql = " Delete DOSupplier_Upload where suratjalanno = '" & Session("suratjalanno") & "' and supplierID = '" & Session("supplier") & "' " & vbCrLf & _
                         " and AffiliateID = '" & Session("affiliateID") & "' and OrderNo = '" & Session("orderno") & "' " & vbCrLf & _
                         " and POno = '" & Session("POno") & "' "

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
                Exit Sub
            End Try

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("YA010IsSubmit") = lblInfo.Text
        End Try
    End Sub

#End Region
End Class