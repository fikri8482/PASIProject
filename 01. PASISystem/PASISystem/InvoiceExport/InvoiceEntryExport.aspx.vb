Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Public Class InvoiceEntryExport
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramSupplier As String
    Dim paramaffiliate As String
    Dim sstatus As Boolean

    'parameter
    Dim pSuratjalan As String
    Dim pInvoiceNo As String
    Dim pAffiliate As String
    Dim pSupplier As String
    Dim pStatus As Boolean
    
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Not IsNothing(Request.QueryString("prm")) Then
                    Dim param As String = Request.QueryString("prm").ToString

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session("sstatus") = "TRUE"
                            pInvoiceNo = Split(param, "|")(0)
                            pSuratjalan = Split(param, "|")(1)
                            pAffiliate = Split(param, "|")(2)
                            pSupplier = Split(param, "|")(3)

                            txtinv.Text = pInvoiceNo
                            txtsuratjalanno.Text = pSuratjalan
                            txtaffiliatecode.Text = pAffiliate

                            If pSuratjalan <> "" Then btnsubmenu.Text = "BACK"

                            pStatus = True
                            Call fillHeader("load")
                            Call up_GridLoad()

                        End If
                    End If

                    btnsubmenu.Text = "BACK"
                End If
            End If
            '===============================================================================

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblerrmessage.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try

    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/InvoiceExport/InvFromSuppListExport.aspx")
    End Sub

    Private Sub fillHeader(ByVal pstatus As String)
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)


        i = 0

        ls_sql = " SELECT  DISTINCT " & vbCrLf & _
                  "         InvoiceDate = ISNULL(CONVERT(CHAR(12), CONVERT(DATETIME, IM.InvoiceDate), 106), " & vbCrLf & _
                  "                              '') , " & vbCrLf & _
                  "         PONO = RM.PONo, " & vbCrLf & _
                  "         SupplierID = RM.SupplierID, " & vbCrLf & _
                  "         AffiliateCode = RM.AffiliateID , " & vbCrLf & _
                  "         AffiliateName = MA.AffiliateName , " & vbCrLf & _
                  "         SupplierSJ = RM.Suratjalanno , " & vbCrLf & _
                  "         InvNo = ISNULL(IM.InvoiceNo, '') , " & vbCrLf & _
                  "         PaymentTerm = ISNULL(IM.PaymentTerm, '') , " & vbCrLf & _
                  "         DueDate = ISNULL(CONVERT(CHAR(12), CONVERT(DATETIME, IM.DueDate), 106), " & vbCrLf & _
                  "                          '') , " & vbCrLf & _
                  "         TotalAmount = ISNULL(TotalAmount, 0) "

        ls_sql = ls_sql + " FROM    dbo.ReceiveForwarder_Master RM " & vbCrLf & _
                          "         LEFT JOIN ReceiveForwarder_Detail RD ON RM.Suratjalanno = RD.Suratjalanno " & vbCrLf & _
                          "                                                 AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          "                                                 AND RM.SupplierID = RD.SupplierID " & vbCrLf & _
                          "                                                 AND RM.POno = RD.POno " & vbCrLf & _
                          "                                                 AND RM.OrderNo = RD.OrderNo " & vbCrLf & _
                          "         LEFT JOIN DOSupplier_Detail_Export SD ON SD.suratjalanno = RM.suratjalanno " & vbCrLf & _
                          "                                                  AND SD.AffiliateID = RM.AffiliateID " & vbCrLf & _
                          "                                                  AND SD.SupplierID = RM.SupplierID " & vbCrLf & _
                          "                                                  AND SD.POno = RM.POno " & vbCrLf & _
                          "                                                  AND SD.OrderNo = RM.OrderNo "

        ls_sql = ls_sql + "                                                  AND SD.Partno = RD.PartNo " & vbCrLf & _
                          "         LEFT JOIN DOSupplier_Master_Export SM ON SM.suratjalanno = SD.suratjalanno " & vbCrLf & _
                          "                                                  AND SM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                          "                                                  AND SM.SupplierID = SD.SupplierID " & vbCrLf & _
                          "                                                  AND SM.POno = SD.POno " & vbCrLf & _
                          "                                                  AND SM.OrderNo = SD.OrderNo " & vbCrLf & _
                          "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                          "                                                       AND IM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                          "                                                       AND IM.SupplierID = RD.SupplierID " & vbCrLf & _
                          "                                                       AND IM.POno = RD.POno " & vbCrLf & _
                          "                                                       AND IM.OrderNo = RD.OrderNo "

        ls_sql = ls_sql + "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                          "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                          "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                          "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                          "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                          "                                                       AND ID.PartNo = RD.PartNo " & vbCrLf & _
                          "         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = SD.AffiliateID " & vbCrLf & _
                          "         LEFT JOIN ms_supplier MS ON MS.SupplierID = SD.SupplierID " & vbCrLf & _
                          "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                          "         LEFT JOIN MS_Parts MP ON MP.Partno = RD.Partno " & vbCrLf & _
                          "         LEFT JOIN MS_UnitCls UC ON UC.unitcls = MP.UnitCls "

        ls_sql = ls_sql + " WHERE RM.AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf & _
                          " AND RM.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "' "

        If txtinv.Text <> "" Then
            ls_sql = ls_sql + " AND IM.InvoiceNo = '" & Trim(txtinv.Text) & "' "
        End If

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            Dim ls_tot As Double
            ls_tot = 0
            txttotalamount.Text = 0
            If ds.Tables(0).Rows.Count > 0 Then

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If pstatus = "grid" Then
                        Grid.JSProperties("s.cpDate") = ds.Tables(0).Rows(i)("invoiceDate")
                        If ds.Tables(0).Rows(i)("invoiceDate") = "" Then Grid.JSProperties("s.cpDate") = Format(Now, "dd MMM yyyy")
                        Grid.JSProperties("cpInvoiceNo") = ds.Tables(0).Rows(i)("InvNo")
                        Grid.JSProperties("s.cpaffcode") = ds.Tables(0).Rows(i)("affiliatecode")
                        Grid.JSProperties("s.cpaffname") = ds.Tables(0).Rows(i)("affiliatename")
                        Grid.JSProperties("s.cpsj") = ds.Tables(0).Rows(i)("suppliersj")
                        Grid.JSProperties("s.cppayment") = ds.Tables(0).Rows(i)("paymentterm")
                        Grid.JSProperties("s.cpduedate") = ds.Tables(0).Rows(i)("Duedate")
                        ls_tot = ls_tot + (ds.Tables(0).Rows(i)("Totalamount"))
                        Grid.JSProperties("s.cptotalamount") = Format(ls_tot, "#,###,###.00")
                    Else
                        txtinvdate.Value = Format(ds.Tables(0).Rows(i)("invoiceDate"), "dd MMM yyyy")
                        If Trim(ds.Tables(0).Rows(i)("invoiceDate")) = "" Then txtinvdate.Text = Format(Now, "dd MMM yyyy")
                        txtinv.Text = ds.Tables(0).Rows(i)("InvNo")
                        txtaffiliatecode.Text = ds.Tables(0).Rows(i)("AffiliateCode")
                        txtaffiliatename.Text = ds.Tables(0).Rows(i)("affiliatename")
                        txtsuratjalanno.Text = ds.Tables(0).Rows(i)("suppliersj")
                        txtpono.Text = ds.Tables(0).Rows(i)("PONO")
                        txtpayment.Text = ds.Tables(0).Rows(i)("paymentterm")
                        txtsupplier.Text = ds.Tables(0).Rows(i)("supplierID")
                        dt2.Value = Format(ds.Tables(0).Rows(i)("Duedate"), "dd MMM yyyy")
                        If Trim(ds.Tables(0).Rows(i)("Duedate")) = "" Then dt2.Value = Format(Now, "dd MMM yyyy")
                        txttotalamount.Text = Format(txttotalamount.Text + (ds.Tables(0).Rows(i)("Totalamount")), "#,###,###.00")

                    End If
                Next

            Else
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            End If
            cn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_sql = " SELECT " & vbCrLf & _
                  " no = ROW_NUMBER() OVER ( ORDER BY RM.OrderNo ), " & vbCrLf & _
                  " orderno = ISNULL(RM.OrderNo,''), " & vbCrLf & _
                  " partno = ISNULL(RD.PartNo,''), " & vbCrLf & _
                  " partname = ISNULL(MP.PartName,''), " & vbCrLf & _
                  " uom = ISNULL(UC.DESCRIPTION,''), " & vbCrLf & _
                  " qtybox = ISNULL(MPM.QtyBox,0), " & vbCrLf & _
                  " suppdelqty = ISNULL(SD.DOQty,0), " & vbCrLf & _
                  " goodrecqty = ISNULL(RD.GoodRecQty,0), " & vbCrLf & _
                  " suppqty = COALESCE(ID.Qty,RD.GoodRecQty), " & vbCrLf & _
                  " diffqty = COALESCE(ID.Qty,RD.GoodRecQty) - ISNULL(SD.DOQty,0), " & vbCrLf

            ls_sql = ls_sql + " delqty = CEILING(ISNULL(SD.DOQty,0)/ISNULL(MPM.QtyBox,0)), " & vbCrLf & _
                              " suppcurr = ISNULL(MC.DESCRIPTION,''), " & vbCrLf & _
                              " suppprice = ISNULL(Coalesce(ID.Price,MPR.Price),0), " & vbCrLf & _
                              " suppamount = ISNULL(ID.Amount,0) " & vbCrLf & _
                              " FROM dbo.ReceiveForwarder_Master RM " & vbCrLf & _
                              "         LEFT JOIN ReceiveForwarder_Detail RD ON RM.Suratjalanno = RD.Suratjalanno " & vbCrLf & _
                              "                                                 AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                 AND RM.SupplierID = RD.SupplierID " & vbCrLf & _
                              "                                                 AND RM.POno = RD.POno " & vbCrLf & _
                              "                                                 AND RM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "         LEFT JOIN DOSupplier_Detail_Export SD ON SD.suratjalanno = RM.suratjalanno " & vbCrLf

            ls_sql = ls_sql + "                                                  AND SD.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              "                                                  AND SD.SupplierID = RM.SupplierID " & vbCrLf & _
                              "                                                  AND SD.POno = RM.POno " & vbCrLf & _
                              "                                                  AND SD.OrderNo = RM.OrderNo " & vbCrLf & _
                              "                                                  AND SD.Partno = RD.PartNo " & vbCrLf & _
                              "         LEFT JOIN DOSupplier_Master_Export SM ON SM.suratjalanno = SD.suratjalanno " & vbCrLf & _
                              "                                                  AND SM.AffiliateID = SD.AffiliateID " & vbCrLf & _
                              "                                                  AND SM.SupplierID = SD.SupplierID " & vbCrLf & _
                              "                                                  AND SM.POno = SD.POno " & vbCrLf & _
                              "                                                  AND SM.OrderNo = SD.OrderNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Master_Export IM ON IM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf

            ls_sql = ls_sql + "                                                       AND IM.AffiliateID = RD.AffiliateID " & vbCrLf & _
                              "                                                       AND IM.SupplierID = RD.SupplierID " & vbCrLf & _
                              "                                                       AND IM.POno = RD.POno " & vbCrLf & _
                              "                                                       AND IM.OrderNo = RD.OrderNo " & vbCrLf & _
                              "         LEFT JOIN InvoiceSupplier_Detail_Export ID ON ID.InvoiceNo = IM.InvoiceNo " & vbCrLf & _
                              "                                                       AND ID.AffiliateID = IM.AffiliateID " & vbCrLf & _
                              "                                                       AND ID.SupplierID = IM.SupplierID " & vbCrLf & _
                              "                                                       AND ID.POno = IM.POno " & vbCrLf & _
                              "                                                       AND ID.OrderNo = IM.OrderNo " & vbCrLf & _
                              "                                                       AND ID.PartNo = RD.PartNo " & vbCrLf & _
                              " 		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = SD.AffiliateID " & vbCrLf

            ls_sql = ls_sql + "         LEFT JOIN MS_Price  MPR ON MPR.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              "                                    AND MPR.partno = ID.PartNo " & vbCrLf & _
                              "                                    AND MPR.CurrCls = ID.Curr " & vbCrLf & _
                              "                                    AND IM.InvoiceDate BETWEEN MPR.StartDate and MPR.EndDate " & vbCrLf

            ls_sql = ls_sql + "         LEFT JOIN ms_supplier MS ON MS.SupplierID = SD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN MS_CurrCls MC ON MC.CurrCls = ID.Curr " & vbCrLf & _
                              "         LEFT JOIN MS_Parts MP ON MP.Partno = RD.Partno " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = RD.PartNo and MPM.AffiliateID = RD.AffiliateID and MPM.SupplierID = RD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN MS_UnitCls UC ON UC.unitcls = MP.UnitCls " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()

            If Grid.VisibleRowCount = 0 Then
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                'Call colorGrid()
            End If
        End Using
    End Sub

    'Private Sub colorGrid()

    '    Grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(1).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(2).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(3).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(4).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(5).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(6).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(7).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(8).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(9).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(10).CellStyle.BackColor = Drawing.Color.White
    '    Grid.VisibleColumns(11).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(12).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(13).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(14).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(15).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(16).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(17).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(18).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(19).CellStyle.BackColor = Drawing.Color.LightYellow
    '    Grid.VisibleColumns(20).CellStyle.BackColor = Drawing.Color.LightYellow

    'End Sub

    'Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
    '    Dim ls_SQL As String = "", ls_MsgID As String = ""
    '    Dim ls_Active As String = "", iLoop As Long = 1
    '    Dim isStatusNew As Boolean
    '    Dim pIsUpdate As Boolean
    '    Dim sqlstring As String
    '    Dim i As Long = 0
    '    Dim pReceiveDate As Date
    '    Dim pPokanban As String
    '    isStatusNew = False

    '    Session.Remove("sstatus")

    '    Session("sstatus") = "TRUE"
    '    If txttotalamount.Text = "" Then txttotalamount.Text = 0
    '    pReceiveDate = txtinvdate.Text

    '    Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '        cn.Open()

    '        Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
    '            Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)
    '            With Grid
    '                For iLoop = 0 To e.UpdateValues.Count - 1
    '                    'cek QTY tidak boleh melebihi Qty
    '                    If CDbl(e.UpdateValues(iLoop).NewValues("suppqty").ToString()) > CDbl(e.UpdateValues(iLoop).NewValues("pasirecqty").ToString()) Then
    '                        Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
    '                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                        Session("sstatus") = "FALSE"
    '                        Exit Sub
    '                    End If
    '                    'cek QTY tidak boleh melebihi Qty

    '                    If Trim(e.UpdateValues(iLoop).NewValues("pokanban").ToString()) = "YES" Then pPokanban = "1" Else pPokanban = "0"

    '                    sqlstring = "SELECT * FROM dbo.InvoiceSupplier_Detail WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                                " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    '                                " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf & _
    '                                " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("partno").ToString()) & "' " & vbCrLf

    '                    sqlComm = New SqlCommand(sqlstring, cn, sqlTran)
    '                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

    '                    If sqlRdr.Read Then
    '                        pIsUpdate = True
    '                    Else
    '                        pIsUpdate = False
    '                    End If
    '                    sqlRdr.Close()

    '                    If pIsUpdate = False Then
    '                        ls_SQL = ""
    '                        ''INSERT KANBAN
    '                        'ls_SQL = " INSERT INTO dbo.ReceivePASI_Detail " & vbCrLf & _
    '                        '          "         ( SuratJalanNo , " & vbCrLf & _
    '                        '          "           SupplierID , " & vbCrLf & _
    '                        '          "           PONo , " & vbCrLf & _
    '                        '          "           POKanbanCls , " & vbCrLf & _
    '                        '          "           KanbanNo , " & vbCrLf & _
    '                        '          "           PartNo , " & vbCrLf & _
    '                        '          "           UnitCls , " & vbCrLf & _
    '                        '          "           GoodRecQty, " & vbCrLf & _
    '                        '          "           DefectRecQty, AffiliateID " & vbCrLf & _
    '                        '          "         ) " & vbCrLf & _
    '                        '          " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

    '                        'ls_SQL = ls_SQL + "           '" & Trim(txtaffiliatecode.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
    '                        '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' , -- PONo - char(20) " & vbCrLf & _
    '                        '                  "           '" & pPokanban & "' , -- POKansbanCls - char(1) " & vbCrLf & _
    '                        '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "' , -- KanbanNo - char(20) " & vbCrLf & _
    '                        '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' , -- PartNo - char(120) " & vbCrLf & _
    '                        '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colunitcls").ToString()) & "' , -- UnitCls - char(3) " & vbCrLf & _
    '                        '                  "           " & CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
    '                        '                  "           " & CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
    '                        '                  "           '" & txtaffiliate.Text & "'" & vbCrLf & _
    '                        '                  "         ) "


    '                    ElseIf pIsUpdate = True Then
    '                        'Update Data
    '                        ls_SQL = " UPDATE dbo.InvoiceSupplier_Detail SET " & vbCrLf & _
    '                                     " InvQty = " & CDbl(e.UpdateValues(iLoop).NewValues("suppqty").ToString()) & ", " & vbCrLf & _
    '                                     " InvAmount = InvPrice * " & CDbl(e.UpdateValues(iLoop).NewValues("suppqty").ToString()) & " " & vbCrLf & _
    '                                     " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                                     " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    '                                     " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf & _
    '                                     " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("partno").ToString()) & "' " & vbCrLf
    '                    End If

    '                    sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
    '                    sqlComm.ExecuteNonQuery()
    '                    Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
    '                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                Next iLoop
    '            End With

    '            sqlComm.Dispose()
    '            sqlTran.Commit()
    '        End Using

    '        cn.Close()
    '    End Using
    '    Call colorGrid()
    'End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)

        Try
            Select Case pAction

                Case "gridload"
                    Call fillHeader("grid")
                    Call up_GridLoad()
                    If pAction = "" Then
                        If Grid.VisibleRowCount = 0 Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        Else
                            Grid.JSProperties("cpMessage") = ""
                            lblerrmessage.Text = ""
                        End If
                    End If
                    'Call colorGrid()
                Case "save"
                    If Session("sstatus") Is Nothing Then Session("sstatus") = "TRUE"
                    Call up_GridLoad()
                    If Session("sstatus") = "TRUE" Then Call saveData()
                    Call fillHeader("grid")
                    Call up_GridLoad()
                    If Session("sstatus") = "TRUE" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                    Else
                        Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                    End If

                Case "kosong"

            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Grid.FocusedRowIndex = -1

        Finally
            'If (Not IsNothing(Session("YA010Msg"))) Then Grid.JSProperties("cpMessage") = Session("YA010Msg") : Session.Remove("YA010Msg")
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
        End Try
    End Sub

    ''Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
    ''    Dim x As Integer = CInt(e.VisibleIndex.ToString())
    ''    Dim pRemaining As Double

    ''    If x > Grid.VisibleRowCount Then Exit Sub
    ''    If e.DataColumn.FieldName = "diffqty" Then
    ''        pRemaining = e.GetValue("diffqty")
    ''    End If

    ''    With Grid
    ''        If .VisibleRowCount > 0 Then
    ''            If pRemaining > 0 Then
    ''                If e.DataColumn.FieldName = "diffqty" Then
    ''                    e.Cell.BackColor = Color.HotPink
    ''                End If
    ''            End If

    ''        End If
    ''    End With
    ''End Sub


    Private Sub saveData()
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 1
        Dim isStatusNew As Boolean
        Dim pIsUpdate As Boolean
        Dim sqlstring As String
        Dim i As Long = 0
        Dim pInvoiceDate As Date
        'Dim pPokanban As String
        isStatusNew = False

        If txttotalamount.Text = "" Then txttotalamount.Text = 0
        pInvoiceDate = txtinvdate.Text

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
                Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)
                With Grid
                    For i = 0 To Grid.VisibleRowCount - 1
                        'cek QTY tidak boleh melebihi Qty
                        If CDbl(Grid.GetRowValues(i, "goodrecqty").ToString) > CDbl(Grid.GetRowValues(i, "suppdelqty").ToString) Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("sstatus") = "FALSE"
                            Exit Sub
                        Else
                            txtstatus.Text = "TRUE"
                        End If
                        'cek QTY tidak boleh melebihi Qty

                        sqlstring = "SELECT * FROM dbo.InvoiceSupplier_Detail_Export  with(nolock) WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " --AND SupplierID = '" & Trim(txtsupplier.Text) & "' " & vbCrLf & _
                                    " and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
                                    " AND PartNo = '" & Trim(Grid.GetRowValues(i, "partno").ToString) & "' " & vbCrLf & _
                                    " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'"

                        sqlComm = New SqlCommand(sqlstring, cn, sqlTran)
                        Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                        If sqlRdr.Read Then
                            pIsUpdate = True
                        Else
                            pIsUpdate = False
                        End If
                        sqlRdr.Close()

                        If pIsUpdate = False Then
                            ls_SQL = ""
                            ls_SQL = " INSERT INTO dbo.InvoiceSupplier_Detail_Export " & vbCrLf & _
                                      "         ( InvoiceNo , " & vbCrLf & _
                                      "           SupplierID , " & vbCrLf & _
                                      "           AffiliateID , " & vbCrLf & _
                                      "           PONo , " & vbCrLf & _
                                      "           PartNo , " & vbCrLf & _
                                      "           OrderNo , " & vbCrLf & _
                                      "           Qty , " & vbCrLf & _
                                      "           Curr, " & vbCrLf & _
                                        "           Price, Amount " & vbCrLf & _
                                      "         ) " & vbCrLf & _
                                      " VALUES  ( '" & txtinv.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                            ls_SQL = ls_SQL + "           '" & Trim(txtsupplier.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                                              "           '" & Trim(txtaffiliatecode.Text) & "' , -- PONo - char(20) " & vbCrLf & _
                                              "           '" & Trim(txtsupplier.Text) & "' , --- char(1) " & vbCrLf & _
                                              "           '" & Trim(txtaffiliatecode.Text) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                                              "           '" & Trim(txtpono.Text) & "' , -- PartNo - char(120) " & vbCrLf & _
                                              "           '" & Trim(Grid.GetRowValues(i, "orderno").ToString) & "' , -- UnitCls - char(3) " & vbCrLf & _
                                              "           " & CDbl(Grid.GetRowValues(i, "suppqty").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                                              "           " & (Grid.GetRowValues(i, "curr").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                                              "           " & CDbl(Grid.GetRowValues(i, "price").ToString) & ", " & vbCrLf & _
                                              "           " & CDbl(Grid.GetRowValues(i, "amount").ToString) & " ) "

                        ElseIf pIsUpdate = True Then
                            'Update Data
                            ls_SQL = " Update InvoiceSupplier_Detail_Export set " & vbCrLf & _
                                     " Qty = " & CDbl(Grid.GetRowValues(i, "suppqty").ToString) & ", " & vbCrLf & _
                                     " Amount = price * " & CDbl(Grid.GetRowValues(i, "suppqty").ToString) & " " & vbCrLf & _
                                     " WHERE InvoiceNo ='" & Trim(txtinv.Text) & "'" & vbCrLf & _
                                     " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and pono = '" & Trim(txtpono.Text) & "'" & vbCrLf & _
                                     " and orderNo = '" & Trim(Grid.GetRowValues(i, "orderno").ToString) & "' " & vbCrLf & _
                                     " AND PartNo = '" & Trim(Grid.GetRowValues(i, "partno").ToString) & "' " & vbCrLf & _
                                     " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'"
                        End If

                        sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        'End If
                    Next i

                End With

                sqlComm.Dispose()
                sqlTran.Commit()

            End Using

            cn.Close()
        End Using
        'Call colorGrid()
    End Sub

    ''Protected Sub btndelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btndelete.Click
    ''    Dim ls_sql As String

    ''    ls_sql = ""
    ''    Using cn As New SqlConnection(clsGlobal.ConnectionString)
    ''        cn.Open()

    ''        'Using sqlTran As SqlTransaction = cn.BeginTransaction("Cols")
    ''        Dim sqlComm As New SqlCommand(ls_sql, cn)
    ''        ls_sql = "SELECT * FROM dbo.InvoiceSupplier_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    ''                 " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    ''                 " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'"

    ''        sqlComm = New SqlCommand(ls_sql, cn)
    ''        Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

    ''        If sqlRdrM.Read Then
    ''            ls_sql = "delete from InvoiceSupplier_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    ''                     " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    ''                     " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf
    ''            ls_sql = ls_sql + "Delete from ReceivePASI_Detail WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    ''                              " AND SupplierID = '" & Trim(txtaffiliatecode.Text) & "' " & vbCrLf
    ''            sqlRdrM.Close()
    ''            sqlComm = New SqlCommand(ls_sql, cn)
    ''            sqlComm.ExecuteNonQuery()
    ''            Call fillHeader("load")
    ''            Call up_GridLoad()

    ''            Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
    ''            Grid.JSProperties("cpMessage") = lblerrmessage.Text

    ''            txtinv.Text = ""
    ''            txtpayment.Text = ""
    ''            txtnopol.Text = ""
    ''            txtjenisarmada.Text = ""
    ''            txttotalamount.Text = ""

    ''        Else
    ''            'data ga ada
    ''            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    ''            Grid.JSProperties("cpMessage") = lblerrmessage.Text
    ''        End If

    ''        sqlComm.Dispose()
    ''        sqlRdrM.Close()

    ''    End Using
    ''End Sub

End Class