Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Public Class ReceivingEntry
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
    Dim pReceivedate As Date
    Dim psuppID As String
    Dim psuppname As String
    Dim pSuratjalanNo As String
    Dim pPlandelivery As Date
    Dim pDeldate As Date
    Dim pKanbanno As String
    Dim pStatus As Boolean
    Dim pAffiliate As String
    Dim ppono As String
    Dim pSJPasi As String

    Dim ReveivingKanagata As Boolean = False
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Not IsNothing(Request.QueryString("prm")) Then
                    Dim param As String = Request.QueryString("prm").ToString
                    FillCombo()
                    Session("E02ParamPageLoad") = Request.QueryString("prm").ToString()

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session("MenuDesc") = "RECEIVING ENTRY"
                            Session("sstatus") = "TRUE"
                            Session.Remove("SJPasi")
                            pReceivedate = Split(param, "|")(0)
                            psuppID = Split(param, "|")(1)
                            psuppname = Split(param, "|")(2)
                            'pSuratjalanNo = Replace(Split(param, "|")(3), "DAN", "&")
                            pSuratjalanNo = Split(param, "|")(3)
                            'pSJPasi = Replace(Split(param, "|")(4), "DAM", "&")
                            pSJPasi = Split(param, "|")(4)
                            pPlandelivery = Split(param, "|")(14)
                            pDeldate = Split(param, "|")(5)
                            pKanbanno = Split(param, "|")(6)
                            pAffiliate = Split(param, "|")(7)
                            ppono = Split(param, "|")(8)

                            If psuppID <> "" Then btnsubmenu.Text = "BACK"
                            If pReceivedate = "#1/1/1900#" Then pReceivedate = Format(Now, "dd MMM yyyy")
                            'txtreceivedate.Text = Format(pReceivedate, "dd MMM yyyy")
                            dt1.Text = Format(pReceivedate, "dd MMM yyyy")
                            txtsuppliercode.Text = psuppID
                            txtsuppliername.Text = psuppname
                            txtsuratjalanno.Text = pSuratjalanNo
                            'txtplandeliverydate.Text = Format(pPlandelivery, "dd MMM yyyy")

                            If Trim(pSJPasi) = "" Then Session("SJPasi") = pSuratjalanNo Else Session("SJPasi") = pSJPasi
                            Session("AFF") = pAffiliate
                            Session("SUPP") = psuppID

                            txtsupplierdeliverydate.Text = Format(pDeldate, "dd MMM yyyy")
                            txtkanbanno.Text = pKanbanno
                            txtaffiliate.Text = pAffiliate
                            txtpono.Text = ppono
                            txtdrivername.Text = Split(param, "|")(9)
                            txtdrivercontact.Text = Split(param, "|")(10)
                            txtnopol.Text = Split(param, "|")(11)
                            txtjenisarmada.Text = Split(param, "|")(12)
                            txttotalbox.Text = Split(param, "|")(13)
                            txtplandeliverydate.Text = Format(pPlandelivery, "dd MMM yyyy") 'Split(param, "|")(14)

                            pStatus = True
                            'Call fillHeader("load")
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

            Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try

    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/Receiving/SuppDeliveryConf.aspx")
    End Sub

    Private Sub fillHeader(ByVal pstatus As String)
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        Grid.JSProperties("cpDate") = Format(pReceivedate, "dd MMM yyyy")
        Grid.JSProperties("cpScode") = psuppID
        Grid.JSProperties("cpSname") = psuppname
        Grid.JSProperties("cpSJ") = pSuratjalanNo
        Grid.JSProperties("cpPlandeldate") = Format(pPlandelivery, "dd MMM yyyy")
        Grid.JSProperties("cpSdeldate") = Format(pDeldate, "dd MMM yyyy")
        pKanbanno = pKanbanno

        i = 0
        ls_sql = ""
        ls_sql = "    SELECT  " & vbCrLf & _
                  "     supplierID , " & vbCrLf & _
                  "     suppliername , " & vbCrLf & _
                  "     plandeldate, " & vbCrLf & _
                  "     deldate, " & vbCrLf & _
                  "     sj, " & vbCrLf & _
                  "     receivedate , " & vbCrLf & _
                  "     drivername, " & vbCrLf & _
                  "     drivercontact, " & vbCrLf & _
                  "     nopol, " & vbCrLf & _
                  "     jenisarmada, "

        ls_sql = ls_sql + "     totalbox = ISNULL(CEILING(SUM(totalbox)),0),PerformanceCls, PerformanceName" & vbCrLf & _
                          "    FROM( "
        ls_sql = ls_sql + "   " & vbCrLf & _
                  "     SELECT " & vbCrLf & _
                  "             supplierID = POM.SupplierID , " & vbCrLf & _
                  "             suppliername = MS.SupplierName , " & vbCrLf & _
                  "             plandeldate = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(KM.KanbanDate,'')), 106) , " & vbCrLf & _
                  "             deldate = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(DOM.DeliveryDate,'')), 106) , " & vbCrLf & _
                  "             sj = ISNULL(RM.SuratJalanNo, DOM.SuratJalanNo)  , " & vbCrLf & _
                  "             receivedate = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(RM.entryDate,'')), 106) , " & vbCrLf & _
                  "             drivername = Coalesce(RM.DriverName, DOM.DriverName),  " & vbCrLf & _
                  "             drivercontact = Coalesce(RM.DriverContact, DOM.DriverContact) ," & vbCrLf & _
                  "             nopol = Coalesce(RM.NoPol,DOM.nopol), "

        ls_sql = ls_sql + "             jenisarmada = Coalesce(RM.JenisArmada, DOM.JenisArmada), " & vbCrLf & _
                          "             totalbox = Coalesce(RD.GoodRecQty, DOD.DOQty) / ISNULL(POD.POQtyBox,MPM.Qtybox), isnull(RM.PerformanceCls,'') as PerformanceCls,isnull(MPC.Description,'') as PerformanceName " & vbCrLf & _
                          "              " & vbCrLf & _
                          "     FROM    dbo.PO_Master POM " & vbCrLf & _
                          "             INNER JOIN dbo.PO_Detail POD ON POD.PONo = POM.PONo " & vbCrLf & _
                          "                                             AND POD.SupplierID = POM.SupplierID " & vbCrLf & _
                          "                                             AND POD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "                                             AND POM.DeliveryByPasiCls = 1 " & vbCrLf & _
                          "             INNER JOIN dbo.DOSupplier_Detail DOD ON DOD.PONo = POD.PONo " & vbCrLf & _
                          "                                                     AND DOD.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                                                     AND DOD.AffiliateID = POD.AffiliateID "

        ls_sql = ls_sql + "                                                     AND DOD.PartNo = POD.PartNo " & vbCrLf & _
                          "             INNER JOIN dbo.DOSupplier_Master DOM ON DOM.SuratJalanNo = DOD.SuratJalanNo " & vbCrLf & _
                          "                                                     AND DOM.SupplierID = DOD.SupplierID " & vbCrLf & _
                          "                                                     AND DOM.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                          "             LEFT JOIN dbo.ReceivePASI_Detail RD ON RD.SuratJalanNo = DOM.SuratJalanNo " & vbCrLf & _
                          "                                                    AND RD.SupplierID = POD.SupplierID " & vbCrLf & _
                          "                                                    AND RD.PartNo = POD.PartNo " & vbCrLf & _
                          "                                                    AND RD.PONo = POD.PONo " & vbCrLf & _
                          "             LEFT JOIN dbo.ReceivePASI_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                          "                                                    AND RM.SupplierID = RD.SupplierID " & vbCrLf & _
                          "             LEFT JOIN dbo.Kanban_Detail KD ON KD.PONo = POM.PONo "

        ls_sql = ls_sql + "                                               AND KD.SupplierID = POM.SupplierID " & vbCrLf & _
                          "                                               AND KD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "                                               AND KD.PartNo = POD.PartNo " & vbCrLf & _
                          "                                               AND KD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                          "             LEFT JOIN dbo.Kanban_Master KM ON KM.KanbanNo = KD.KanbanNo " & vbCrLf & _
                          "                                               AND KM.SupplierID = POM.SupplierID " & vbCrLf & _
                          "                                               AND KM.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "                                               AND KM.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                          "             INNER JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                          "             LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = DOD.PartNo AND MPM.AffiliateID = DOD.AffiliateID AND MPM.SupplierID = DOD.SupplierID " & vbCrLf & _
                          "             LEFT JOIN MS_PerformanceCls MPC ON MPC.PerformanceCls = RM.PerformanceCls " & vbCrLf & _
                          "             INNER JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID " & vbCrLf & _
                          "             INNER JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID    " & vbCrLf & _
                          " WHERE DOM.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "' " & vbCrLf & _
                          " AND DOM.SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf & _
                          " --AND DOD.KanbanNo = '" & txtkanbanno.Text & "' " & vbCrLf

        'If Session("SJPasi") <> "" Then ls_sql = ls_sql + " and RM.SuratJalanNo = '" & Session("SJPasi") & "' " & vbCrLf
        'If Session("SJPasi") = "" Then ls_sql = ls_sql + " and isnull(RM.SuratJalanNo,'') = '" & Session("SJPasi") & "' " & vbCrLf

        ls_sql = ls_sql + " )x " & vbCrLf & _
                          " 	GROUP BY " & vbCrLf & _
                          "     supplierID , " & vbCrLf & _
                          "     suppliername , " & vbCrLf & _
                          "     plandeldate, " & vbCrLf & _
                          "     deldate, " & vbCrLf & _
                          "     sj, " & vbCrLf & _
                          "     receivedate , " & vbCrLf & _
                          "     drivername, " & vbCrLf & _
                          "     drivercontact, " & vbCrLf & _
                          "     nopol, " & vbCrLf & _
                          "     jenisarmada,PerformanceCls, PerformanceName "

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If pstatus = "grid" Then
                        Grid.JSProperties("cpDname") = ds.Tables(0).Rows(i)("DriverName")
                        Grid.JSProperties("cpDContact") = ds.Tables(0).Rows(i)("DriverContact")
                        Grid.JSProperties("cpNopol") = ds.Tables(0).Rows(i)("NoPol")
                        Grid.JSProperties("cpJenisarmada") = ds.Tables(0).Rows(i)("JenisArmada")
                        Grid.JSProperties("cpTotalbox") = ds.Tables(0).Rows(i)("TotalBox")

                        Grid.JSProperties("cpDate") = ds.Tables(0).Rows(i)("receivedate")
                        Grid.JSProperties("cpScode") = ds.Tables(0).Rows(i)("supplierid")
                        Grid.JSProperties("cpSname") = ds.Tables(0).Rows(i)("suppliername")
                        Grid.JSProperties("cpSJ") = ds.Tables(0).Rows(i)("sj")
                        Grid.JSProperties("cpPlandeldate") = ds.Tables(0).Rows(i)("plandeldate")
                        Grid.JSProperties("cpSdeldate") = ds.Tables(0).Rows(i)("deldate")
                        Grid.JSProperties("cpCls") = ds.Tables(0).Rows(i)("PerformanceCls")
                        Grid.JSProperties("cpClsN") = ds.Tables(0).Rows(i)("PerformanceName")

                        'txtdrivername.Text = ds.Tables(0).Rows(i)("DriverName")
                        'txtdrivercontact.Text = ds.Tables(0).Rows(i)("DriverContact")
                        'txtnopol.Text = ds.Tables(0).Rows(i)("NoPol")
                        'txtjenisarmada.Text = ds.Tables(0).Rows(i)("JenisArmada")
                        'txttotalbox.Text = ds.Tables(0).Rows(i)("TotalBox")

                        'If Trim(ds.Tables(0).Rows(i)("receivedate")) = "01 Jan 1900" Then
                        '    txtreceivedate.Text = Format(Now, "dd MMM yyyy")
                        'Else
                        '    txtreceivedate.Text = Trim(ds.Tables(0).Rows(i)("receivedate"))
                        'End If
                        'txtsuppliercode.Text = ds.Tables(0).Rows(i)("supplierid")
                        'txtsuppliername.Text = ds.Tables(0).Rows(i)("suppliername")
                        'txtsuratjalanno.Text = ds.Tables(0).Rows(i)("sj")
                        'txtplandeliverydate.Text = ds.Tables(0).Rows(i)("plandeldate")
                        'txtsupplierdeliverydate.Text = ds.Tables(0).Rows(i)("deldate")
                        'cbocls.Text = ds.Tables(0).Rows(i)("PerformanceCls")
                        'txtcls.Text = ds.Tables(0).Rows(i)("PerformanceName")

                    Else
                        txtdrivername.Text = ds.Tables(0).Rows(i)("DriverName")
                        txtdrivercontact.Text = ds.Tables(0).Rows(i)("DriverContact")
                        txtnopol.Text = ds.Tables(0).Rows(i)("NoPol")
                        txtjenisarmada.Text = ds.Tables(0).Rows(i)("JenisArmada")
                        txttotalbox.Text = ds.Tables(0).Rows(i)("TotalBox")

                        If Trim(ds.Tables(0).Rows(i)("receivedate")) = "01 Jan 1900" Then
                            'txtreceivedate.Text = Format(Now, "dd MMM yyyy")
                            dt1.Text = Format(Now, "dd MMM yyyy")
                        Else
                            'txtreceivedate.Text = Trim(ds.Tables(0).Rows(i)("receivedate"))
                            dt1.Text = Trim(ds.Tables(0).Rows(i)("receivedate"))
                        End If
                        txtsuppliercode.Text = ds.Tables(0).Rows(i)("supplierid")
                        txtsuppliername.Text = ds.Tables(0).Rows(i)("suppliername")
                        txtsuratjalanno.Text = ds.Tables(0).Rows(i)("sj")
                        txtplandeliverydate.Text = ds.Tables(0).Rows(i)("plandeldate")
                        txtsupplierdeliverydate.Text = ds.Tables(0).Rows(i)("deldate")
                        cbocls.Text = ds.Tables(0).Rows(i)("PerformanceCls")
                        txtcls.Text = ds.Tables(0).Rows(i)("PerformanceName")
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
            ls_sql = " SELECT colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY DOD.PONo )) , " & vbCrLf & _
                     "         colpono = DOD.PONo , " & vbCrLf & _
                     "         colpokanban = CASE WHEN ISNULL(DOD.POKanbanCls, 0) = 0 THEN 'NO' " & vbCrLf & _
                     "                            ELSE 'YES' " & vbCrLf & _
                     "                       END , " & vbCrLf & _
                     "         colkanbanno = DOD.KanbanNo , " & vbCrLf & _
                     "         colpartno = DOD.PartNo , " & vbCrLf & _
                     "         colpartname = MP.PartName , " & vbCrLf & _
                     "         coluom = UC.Description , " & vbCrLf & _
                     "         colqtybox = ISNULL(DOD.POQtyBox,MPM.Qtybox) , " & vbCrLf & _
                     "         colsupplierqty = DOD.DOQty , " & vbCrLf & _
                     "         colPrice = ISNULL(MPR.Price, 0) , "

            ls_sql = ls_sql + "         colreceivingqty = ISNULL(RD.goodRecQty,DOD.DOQty), " & vbCrLf & _
                              "         coldefect = ISNULL(RD.DefectRecQty, 0) , " & vbCrLf & _
                              "         colremaining = ISNULL(POD.POQty, 0) - ISNULL(RD.goodRecQty,DOD.DOQty) , " & vbCrLf & _
                              "         colboxqty = CEILING(ISNULL(RD.goodRecQty,DOD.DOQty) / ISNULL(DOD.POQtyBox,MPM.Qtybox)) , " & vbCrLf & _
                              "         colunitcls = ISNULL(MP.unitcls, '') , " & vbCrLf & _
                              "         colHgood = ISNULL(RD.goodRecQty, 0) , " & vbCrLf & _
                              "         colHdefect = ISNULL(RD.defectRecQty, 0), " & vbCrLf & _
                              "         colKanbanQty = kanbanQty " & vbCrLf & _
                              "  FROM   dbo.DOSupplier_Master DOM " & vbCrLf & _
                              "         INNER JOIN dbo.DOSupplier_Detail DOD ON DOM.SuratJalanNo = DOD.SuratJalanNo " & vbCrLf & _
                              "                                                 AND DOM.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                              "                                                 AND DOM.SupplierID = DOD.SupplierID " & vbCrLf

            ls_sql = ls_sql + "         LEFT JOIN ReceivePasi_Detail RD ON RD.suratJalanNo = DOD.SuratJalanNo " & vbCrLf & _
                              "                                            AND RD.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                              "                                            AND RD.SupplierID = DOD.SupplierID " & vbCrLf & _
                              "                                            AND RD.PartNo = DOD.PartNo " & vbCrLf & _
                              "                                            AND RD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              "                                            AND RD.PONO = DOD.PONO " & vbCrLf & _
                              "         LEFT JOIN ReceivePasi_Master RM ON RM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
                              "                                             AND RM.supplierID= RD.SupplierID " & vbCrLf & _
                              "                             				AND RM.affiliateID = RD.AffiliateID " & vbCrLf & _
                              "         INNER JOIN Kanban_Detail KD ON KD.KanbanNo = DOD.KanbanNo " & vbCrLf & _
                              "                                         AND KD.SupplierID = DOD.SupplierID " & vbCrLf & _
                              "                                         AND KD.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                              "                                         AND KD.PONo = DOD.PONo " & vbCrLf & _
                              "                                         AND KD.PartNo = DOD.PartNo " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = DOD.PartNo " & vbCrLf & _
                              "         LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                              "         LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
                              "         LEFT JOIN PO_Detail POD ON POD.PONO = DOD.PONO " & vbCrLf & _
                              "                                    AND POD.AffiliateID = DOD.AffiliateID " & vbCrLf & _
                              "                                    AND POD.SupplierID = DOD.SupplierID "

            ls_sql = ls_sql + "                                    AND POD.PartNo = DOD.PartNo " & vbCrLf & _
                              "         LEFT JOIN PO_Master POM ON POM.PONO = POD.PONO " & vbCrLf & _
                              "                                    AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              "                                    AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
         "         Left Join MS_Price MPR ON DOD.PartNo = MPR.PartNo  " & vbCrLf & _
                              " 									AND '" & Format(CDate(dt1.Text), "yyyy-MM-dd") & "' between MPR.Startdate and MPR.Enddate " & vbCrLf & _
                              " 									AND MPR.AffiliateID = DOD.SupplierID " & vbCrLf & _
                              " 									AND MPR.DeliveryLocationID = DOD.AffiliateID " & vbCrLf & _
                              "                                     AND MPM.PackingCls = MPR.PackingCls " & vbCrLf & _
                              " WHERE DeliveryByPASICls = 1 and DOM.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "' AND DOM.SupplierID = '" & Trim(txtsuppliercode.Text) & "' --AND DOD.KanbanNo = '" & txtkanbanno.Text & "' " & vbCrLf & _
                              " AND DOD.AffiliateID = '" & Session("AFF") & "' " & vbCrLf & _
                              " AND DOD.SupplierID = '" & Session("SUPP") & "' " & vbCrLf & _
                              ""

            'If Session("SJPasi") <> "" Then ls_sql = ls_sql + " and RM.SuratJalanNo = '" & Session("SJPasi") & "' " & vbCrLf
            'If Session("SJPasi") = "" Then ls_sql = ls_sql + " and isnull(RM.SuratJalanNo,'') = '" & Session("SJPasi") & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
                'Call ColorGrid()
            End With
            sqlConn.Close()

            If Grid.VisibleRowCount = 0 Then
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                Call colorGrid()
            End If
        End Using
    End Sub

    Private Sub colorGrid()

        Grid.VisibleColumns(0).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(1).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(2).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(3).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(4).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(5).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(6).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(7).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(8).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(9).CellStyle.BackColor = Drawing.Color.White
        Grid.VisibleColumns(10).CellStyle.BackColor = Drawing.Color.White
        Grid.VisibleColumns(11).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(12).CellStyle.BackColor = Drawing.Color.LightYellow

    End Sub

    Private Sub FillCombo()
        Dim ls_sql As String = ""

        ls_sql = "SELECT [Performance Cls] = RTRIM(PerformanceCls) ,[Performance Name] = RTRIM(Description) FROM MS_PerformanceCls " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbocls
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Performance Cls")
                .Columns(0).Width = 90
                .Columns.Add("Performance Name")
                .Columns(1).Width = 240

                .TextField = "Performance Cls"
                .DataBind()
                .SelectedIndex = 0
                txtcls.Text = ds.Tables(0).Rows(0)("Performance Name")
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 1
        Dim isStatusNew As Boolean
        Dim pIsUpdate As Boolean
        Dim sqlstring As String
        Dim i As Long = 0
        Dim pReceiveDate As Date
        Dim pPokanban As String
        Dim totalBox As Integer = 0
        isStatusNew = False

        Session.Remove("sstatus")

        Session("sstatus") = "TRUE"
        If txttotalbox.Text = "" Then txttotalbox.Text = 0
        'pReceiveDate = txtreceivedate.Text
        pReceiveDate = dt1.Text

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()
            Dim sqlComm As New SqlCommand '(ls_SQL, cn, sqlTran)
            With Grid
                totalBox = 0
                For iLoop = 0 To e.UpdateValues.Count - 1

                    'cek QTY tidak boleh melebihi Qty
                    If (CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) + CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString())) > CDbl(e.UpdateValues(iLoop).NewValues("colsupplierqty").ToString()) Then
                        'Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                        lblerrmessage.Text = "Qty Can't bigger than Supplier Delivery Qty !"
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                        Session("YA010IsSubmit") = lblerrmessage.Text
                        Session("sstatus") = "FALSE"
                        Exit Sub
                    End If
                    'cek QTY tidak boleh melebihi Qty

                    'cek QTY Receiving tidak boleh melebihi Qty Remaining
                    If CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) > CDbl(e.UpdateValues(iLoop).NewValues("colremaining").ToString()) Then
                        'Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                        lblerrmessage.Text = "Part " & e.UpdateValues(iLoop).NewValues("colpartno").ToString() & " Over PO "
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                        Session("YA010IsSubmit") = lblerrmessage.Text
                        Session("sstatus") = "FALSE"
                        Exit Sub
                    End If
                    'cek QTY Receiving tidak boleh melebihi Qty Remaining

                    If Trim(e.UpdateValues(iLoop).NewValues("colpokanban").ToString()) = "YES" Then pPokanban = "1" Else pPokanban = "0"

                    sqlstring = "SELECT * FROM dbo.ReceivePASI_Detail WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' " & vbCrLf & _
                                    " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' and affiliateID = '" & txtaffiliate.Text & "'" & vbCrLf & _
                                    " and PONO = '" & Trim(txtpono.Text) & "'" & vbCrLf & _
                                    " AND KanbanNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "' " & vbCrLf

                    sqlComm = New SqlCommand(sqlstring, cn)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If pIsUpdate = False Then
                        ls_SQL = ""
                        'INSERT KANBAN
                        ls_SQL = " INSERT INTO dbo.ReceivePASI_Detail " & vbCrLf & _
                                  "         ( SuratJalanNo , " & vbCrLf & _
                                  "           SupplierID , " & vbCrLf & _
                                  "           PONo , " & vbCrLf & _
                                  "           POKanbanCls , " & vbCrLf & _
                                  "           KanbanNo , " & vbCrLf & _
                                  "           PartNo , " & vbCrLf & _
                                  "           UnitCls , " & vbCrLf & _
                                  "           GoodRecQty, " & vbCrLf & _
                                  "           DefectRecQty, AffiliateID " & vbCrLf & _
								  "           ,Price " & vbCrLf & _
                                  "         ) " & vbCrLf & _
                                  " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                        ls_SQL = ls_SQL + "           '" & Trim(txtsuppliercode.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                                          "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' , -- PONo - char(20) " & vbCrLf & _
                                          "           '" & pPokanban & "' , -- POKansbanCls - char(1) " & vbCrLf & _
                                          "           '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                                          "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' , -- PartNo - char(120) " & vbCrLf & _
                                          "           '" & Trim(e.UpdateValues(iLoop).NewValues("colunitcls").ToString()) & "' , -- UnitCls - char(3) " & vbCrLf & _
                                          "           " & CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
                                          "           " & CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
                                          "           '" & txtaffiliate.Text & "' , " & vbCrLf & _
                                          "           " & CDbl(e.UpdateValues(iLoop).NewValues("colPrice").ToString()) & vbCrLf & _
                                          "         ) "
                        totalBox = totalBox + (CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) / CDbl(e.UpdateValues(iLoop).NewValues("colqtybox").ToString()))

                        ''save to remaining
                        'If CDbl(Trim(e.UpdateValues(iLoop).NewValues("colKanbanQty").ToString())) - CDbl(Trim(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString())) Then
                        '    ls_SQL = ls_SQL + "Insert Into ReceivePASI_Remaining VALUES ( " & vbCrLf & _
                        '                      " '" & txtsuratjalanno.Text & "', " & vbCrLf & _
                        '                      " '" & Trim(txtaffiliate.Text) & "', " & vbCrLf & _
                        '                      " '" & Trim(txtsuppliercode.Text) & "', " & vbCrLf & _
                        '                      " '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "', " & vbCrLf & _
                        '                      " '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "', " & vbCrLf & _
                        '                      " '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "', " & vbCrLf & _
                        '                      " " & CDbl(Trim(e.UpdateValues(iLoop).NewValues("colKanbanQty").ToString())) & " - " & CDbl(Trim(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString())) & " " & vbCrLf & _
                        '                      " ,'0' " & vbCrLf & _
                        '                      " ) "
                        'End If
                        ''save to remaining

                    ElseIf pIsUpdate = True Then
                        'Update Data
                        ls_SQL = " Update ReceivePASI_Detail set " & vbCrLf & _
                                 " GoodRecQty = " & CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) & ", " & vbCrLf & _
                                 " DefectRecQty = " & CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString()) & " " & vbCrLf & _
                                 " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                 " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' " & vbCrLf & _
                                 " AND poNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' " & vbCrLf & _
                                 " AND AffiliateID = '" & txtaffiliate.Text & "'" & vbCrLf & _
                                 " AND KanbanNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "' " & vbCrLf
                        totalBox = totalBox + (CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) / CDbl(e.UpdateValues(iLoop).NewValues("colqtybox").ToString()))

                        If CDbl(Trim(e.UpdateValues(iLoop).NewValues("colKanbanQty").ToString())) - CDbl(Trim(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString())) Then
                            ls_SQL = ls_SQL + " Update ReceivePASI_Remaining set " & vbCrLf & _
                                     " status = '0', " & vbCrLf & _
                                     " Qty = " & CDbl(Trim(e.UpdateValues(iLoop).NewValues("colKanbanQty").ToString())) & " - " & CDbl(Trim(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString())) & " " & vbCrLf & _
                                     " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' " & vbCrLf & _
                                     " AND poNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' " & vbCrLf & _
                                     " AND AffiliateID = '" & txtaffiliate.Text & "'" & vbCrLf & _
                                     " AND KanbanNo = '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "' " & vbCrLf
                        End If

                        If ReveivingKanagata = True Then
                            Dim ls_Kanagata As String
                            Dim ls_TotalActualProd As Double = CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) + CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString())
                            Dim ls_TotalActualProdOld As Double = CDbl(e.UpdateValues(iLoop).OldValues("colreceivingqty").ToString()) + CDbl(e.UpdateValues(iLoop).OldValues("coldefect").ToString())

                            ls_Kanagata = "UPDATE PASIKANAGATA.PASI_KANAGATA.dbo.PartControlDetail_AmortizationPASI " & vbCrLf & _
                                            "SET ActualProductionQty = ISNULL(ActualProductionQty,0) - " & ls_TotalActualProdOld & " + " & ls_TotalActualProd & ", ActualAmortizationAmount = (ISNULL(ActualProductionQty,0) - " & ls_TotalActualProdOld & " + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0)" & vbCrLf & _
                                            ", RemainingAmortizationAmount = ISNULL(TotalMoldCost,0) - ((ISNULL(ActualProductionQty,0) - " & ls_TotalActualProdOld & " + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0))" & vbCrLf & _
                                            "WHERE PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "'"
                            '7009-1862-02
                            'ls_Kanagata = "UPDATE PASIKANAGATA.PASI_KANAGATA.dbo.PartControlDetail_AmortizationPASI " & vbCrLf & _
                            '                "SET ActualProductionQty = ISNULL(ActualProductionQty,0) - " & ls_TotalActualProdOld & " + " & ls_TotalActualProd & ", ActualAmortizationAmount = (ISNULL(ActualProductionQty,0) - " & ls_TotalActualProdOld & " + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0)" & vbCrLf & _
                            '                ",RemainingAmortizationAmount = ISNULL(TotalMoldCost,0) - ((ISNULL(ActualProductionQty,0) - " & ls_TotalActualProdOld & " + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0))" & vbCrLf & _
                            '                "WHERE PartNo = '7009-1862-02'"

                            sqlComm = New SqlCommand(ls_Kanagata, cn)
                            sqlComm.ExecuteNonQuery()
                        End If
                    End If

                    sqlComm = New SqlCommand(ls_SQL, cn)
                    sqlComm.ExecuteNonQuery()

                    'insert master
                    sqlstring = "SELECT * FROM dbo.ReceivePASI_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' and AffiliateID = '" & txtaffiliate.Text & "' " & vbCrLf

                    sqlComm = New SqlCommand(sqlstring, cn)
                    Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdrM.Read Then
                        'UPDATE
                        ls_SQL = " UPDATE dbo.ReceivePASI_Master SET " & vbCrLf & _
                                     " DriverName = '" & Trim(txtdrivername.Text) & "', " & vbCrLf & _
                                     " DriverContact = '" & Trim(txtdrivercontact.Text) & "', " & vbCrLf & _
                                     " NoPol = '" & Trim(txtnopol.Text) & "', " & vbCrLf & _
                                     " JenisArmada = '" & Trim(txtjenisarmada.Text) & "', " & vbCrLf & _
                                     " TotalBox= " & totalBox & "," & vbCrLf & _
                                     " UpdateUser = '" & Session("UserID") & "', " & vbCrLf & _
                                     " UpdateDate = GETDATE(), " & vbCrLf & _
                                     " PerformanceCls = '" & Trim(cbocls.Text) & "' " & vbCrLf & _
                                     " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf & _
                                     " AND AffiliateID = '" & txtaffiliate.Text & "'"
                    ElseIf Not sqlRdrM.Read Then
                        'INSERT
                        ls_SQL = " INSERT INTO dbo.ReceivePASI_Master " & vbCrLf & _
                                    "         ( SuratJalanNo , " & vbCrLf & _
                                    "           SupplierID , " & vbCrLf & _
                                    "           ReceiveDate , " & vbCrLf & _
                                    "           ReceiveBy , " & vbCrLf & _
                                    "           JenisArmada , " & vbCrLf & _
                                    "           DriverName , " & vbCrLf & _
                                    "           DriverContact , " & vbCrLf & _
                                    "           NoPol , " & vbCrLf & _
                                    "           TotalBox , " & vbCrLf & _
                                    "           EntryDate , "

                        ls_SQL = ls_SQL + "           EntryUser , " & vbCrLf & _
                                          "           UpdateDate , " & vbCrLf & _
                                          "           UpdateUser, AffiliateID, PerformanceCls " & vbCrLf & _
                                          "         ) " & vbCrLf & _
                                          " VALUES  ( '" & Trim(txtsuratjalanno.Text) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                                          "           '" & Trim(txtsuppliercode.Text) & "' , -- SupplierID - char(10) " & vbCrLf & _
                                          "           '" & Format(pReceiveDate, "yyyyMMdd") & "' , -- ReceiveDate - date " & vbCrLf & _
                                          "           '" & Session("UserID") & "' , -- ReceiveBy - char(15) " & vbCrLf & _
                                          "           '" & Trim(txtjenisarmada.Text) & "' , -- JenisArmada - char(15) " & vbCrLf & _
                                          "           '" & Trim(txtdrivername.Text) & "' , -- DriverName - char(15) " & vbCrLf & _
                                          "           '" & Trim(txtdrivercontact.Text) & "' , -- DriverContact - char(15) " & vbCrLf

                        ls_SQL = ls_SQL + "           '" & Trim(txtnopol.Text) & "' , -- NoPol - char(10) " & vbCrLf & _
                                          "           " & totalBox & " , -- TotalBox - numeric " & vbCrLf & _
                                          "           Getdate() , -- EntryDate - datetime " & vbCrLf & _
                                          "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
                                          "           Getdate() , -- UpdateDate - datetime " & vbCrLf & _
                                          "           '" & Session("UserID") & "', '" & txtaffiliate.Text & "',  -- UpdateUser - char(15) " & vbCrLf & _
                                          "           '" & Trim(cbocls.Text) & "' " & vbCrLf & _
                                          "         ) "

                    End If
                    sqlRdrM.Close()
                    sqlComm = New SqlCommand(ls_SQL, cn)
                    sqlComm.ExecuteNonQuery()
                    sqlRdrM.Close()
                    'insert master

                    Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Session("YA010IsSubmit") = lblerrmessage.Text

                Next iLoop
            End With
            'Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")


            '    sqlComm.Dispose()
            '    sqlTran.Commit()
            'End Using

            cn.Close()
        End Using
        Call colorGrid()
    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

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
                    Call colorGrid()
                    Grid.JSProperties("cpMessage") = Session("MsgRec")
                    lblerrmessage.Text = Session("MsgRec")
                    Session.Remove("MsgRec")
                Case "save"
                    Session.Remove("MsgRec")
                    If Session("sstatus") Is Nothing Then Session("sstatus") = "TRUE"
                    Call up_GridLoad()
                    If Session("sstatus") = "TRUE" Then Call saveData()
                    Call fillHeader("grid")
                    Call up_GridLoad()
                    If Session("sstatus") = "TRUE" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                    ElseIf Session("sstatus") = "FALSE" Then
                        'Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                        'Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        'lblerrmessage.Text = lblerrmessage.Text
                        lblerrmessage.Text = Grid.JSProperties("cpMessage")
                    End If
                    Session("MsgRec") = lblerrmessage.Text

                Case "kosong"

                Case "sendtosupplier"
                    Call UpdateExcel(True, txtaffiliate.Text, txtsuratjalanno.Text, txtsuppliercode.Text)
                    'Call clsMsg.DisplayMessage(lblerrmessage, "1010", clsMessage.MsgType.InformationMessage)
                    'Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    'lblerrmessage.Text = lblerrmessage.Text
                    Call clsMsg.DisplayMessage(lblerrmessage, "1010", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
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

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        Dim pRemaining As Double

        If x > Grid.VisibleRowCount Then Exit Sub
        If e.DataColumn.FieldName = "colremaining" Then
            pRemaining = e.GetValue("colremaining")
        End If

        With Grid
            If .VisibleRowCount > 0 Then
                If pRemaining > 0 Then
                    If e.DataColumn.FieldName = "colremaining" Then
                        e.Cell.BackColor = Color.HotPink
                    End If
                End If

                If e.GetValue("colHgood") = 0 Then
                    If e.DataColumn.FieldName = "colreceivingqty" Then
                        e.Cell.BackColor = Color.Yellow
                    End If
                    If e.DataColumn.FieldName = "coldefect" Then
                        e.Cell.BackColor = Color.Yellow
                    End If
                Else
                    If e.DataColumn.FieldName = "colreceivingqty" Then
                        e.Cell.BackColor = Color.White
                    End If
                    If e.DataColumn.FieldName = "coldefect" Then
                        e.Cell.BackColor = Color.White
                    End If
                End If

            End If
        End With
    End Sub

    Private Sub saveData()
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 1
        Dim isStatusNew As Boolean
        Dim pIsUpdate As Boolean
        Dim sqlstring As String
        Dim i As Long = 0
        Dim pReceiveDate As Date
        Dim pPokanban As String
        Dim ls_totalbox As Integer
        isStatusNew = False

        ls_totalbox = 0
        If txttotalbox.Text = "" Then txttotalbox.Text = 0
        'pReceiveDate = txtreceivedate.Text
        pReceiveDate = dt1.Text

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()
            Dim sqlComm As New SqlCommand
            With Grid
                For i = 0 To Grid.VisibleRowCount - 1
                    'cek QTY tidak boleh melebihi Qty
                    If CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) > CDbl(Grid.GetRowValues(i, "colsupplierqty").ToString) Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        Session("sstatus") = "FALSE"
                        Exit Sub
                    Else
                        txtstatus.Text = "TRUE"
                    End If
                    'cek QTY tidak boleh melebihi Qty

                    ''CEK OVER RECEIVING
                    'ls_SQL = " Select Remaining = POD.POQty - ISNULL(RD.Qty,0) From PO_Detail POD " & vbCrLf & _
                    '            " LEFT JOIN (Select SupplierID,AffiliateID,PONo,PartNo,Qty=SUM(GoodRecQty)  " & vbCrLf & _
                    '            " 		   From ReceivePASI_Detail " & vbCrLf & _
                    '            " 		   Group By SupplierID,AffiliateID,PONo,PartNo) RD  " & vbCrLf & _
                    '            " 	   ON POD.PONo = RD.PONo " & vbCrLf & _
                    '            " 	   AND POD.AffiliateID = RD.AffiliateID " & vbCrLf & _
                    '            " 	   AND POD.SupplierID = RD.SupplierID " & vbCrLf & _
                    '            " 	   AND POD.PartNo = RD.PartNo " & vbCrLf & _
                    '            " Where POD.PONo = '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "' " & vbCrLf & _
                    '            " AND POD.AffiliateID = '" & Trim(txtaffiliate.Text) & "' " & vbCrLf & _
                    '            " AND POD.SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & _
                    '            " AND POD.PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' "
                    'Dim sqlCmd2 As New SqlCommand(ls_SQL, cn)
                    'Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                    'Dim ds2 As New DataSet
                    'sqlDA2.Fill(ds2)
                    'If ds2.Tables(0).Rows.Count > 0 Then
                    '    If (ds2.Tables(0).Rows(0)("Remaining") - CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString)) > 0 Then
                    '        'Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                    '        'Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    '        'lblerrmessage.Text = lblerrmessage.Text
                    '        'Session("YA010IsSubmit") = lblerrmessage.Text
                    '        'Session("sstatus") = "FALSE"
                    '        'Exit Sub
                    '    End If
                    'End If



                    'If CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) <> CDbl(Grid.GetRowValues(i, "colHgood").ToString) _
                    '    Or CDbl(Grid.GetRowValues(i, "coldefect").ToString) <> CDbl(Grid.GetRowValues(i, "colHdefect").ToString) Then

                    ls_totalbox = ls_totalbox + (CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) / CDbl(Grid.GetRowValues(i, "colqtybox").ToString))

                    If Trim(Grid.GetRowValues(i, "colpokanban").ToString) = "YES" Then pPokanban = "1" Else pPokanban = "0"
                    'insert master
                    sqlstring = "SELECT * FROM dbo.ReceivePASI_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                " AND AffiliateID = '" & Trim(txtaffiliate.Text) & "'" & vbCrLf & _
                                " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf

                    sqlComm = New SqlCommand(sqlstring, cn)
                    Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdrM.Read Then
                        'UPDATE
                        ls_SQL = " UPDATE dbo.ReceivePASI_Master SET " & vbCrLf & _
                                     " DriverName = '" & Trim(txtdrivername.Text) & "', " & vbCrLf & _
                                     " DriverContact = '" & Trim(txtdrivercontact.Text) & "', " & vbCrLf & _
                                     " NoPol = '" & Trim(txtnopol.Text) & "', " & vbCrLf & _
                                     " JenisArmada = '" & Trim(txtjenisarmada.Text) & "', " & vbCrLf & _
                                     " TotalBox = " & Trim(ls_totalbox) & " , " & vbCrLf & _
                                     " UpdateUser = '" & Session("UserID") & "', " & vbCrLf & _
                                     " UpdateDate = GETDATE(), " & vbCrLf & _
                                     " PerformanceCls = '" & Trim(cbocls.Text) & "' " & vbCrLf & _
                                     " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf
                    ElseIf Not sqlRdrM.Read Then
                        'INSERT
                        ls_SQL = " INSERT INTO dbo.ReceivePASI_Master " & vbCrLf & _
                                    "         ( SuratJalanNo , " & vbCrLf & _
                                    "           SupplierID , " & vbCrLf & _
                                    "           ReceiveDate , " & vbCrLf & _
                                    "           ReceiveBy , " & vbCrLf & _
                                    "           JenisArmada , " & vbCrLf & _
                                    "           DriverName , " & vbCrLf & _
                                    "           DriverContact , " & vbCrLf & _
                                    "           NoPol , " & vbCrLf & _
                                    "           TotalBox , " & vbCrLf & _
                                    "           EntryDate , "

                        ls_SQL = ls_SQL + "           EntryUser , " & vbCrLf & _
                                          "           UpdateDate , " & vbCrLf & _
                                          "           UpdateUser, AffiliateID" & vbCrLf & _
                                          "         ) " & vbCrLf & _
                                          " VALUES  ( '" & Trim(txtsuratjalanno.Text) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                                          "           '" & Trim(txtsuppliercode.Text) & "' , -- SupplierID - char(10) " & vbCrLf & _
                                          "           '" & Format(pReceiveDate, "yyyyMMdd") & "' , -- ReceiveDate - date " & vbCrLf & _
                                          "           '" & Session("UserID") & "' , -- ReceiveBy - char(15) " & vbCrLf & _
                                          "           '" & Trim(txtjenisarmada.Text) & "' , -- JenisArmada - char(15) " & vbCrLf & _
                                          "           '" & Trim(txtdrivername.Text) & "' , -- DriverName - char(15) " & vbCrLf & _
                                          "           '" & Trim(txtdrivercontact.Text) & "' , -- DriverContact - char(15) " & vbCrLf

                        ls_SQL = ls_SQL + "           '" & Trim(txtnopol.Text) & "' , -- NoPol - char(10) " & vbCrLf & _
                                          "           " & ls_totalbox & " , -- TotalBox - numeric " & vbCrLf & _
                                          "           Getdate() , -- EntryDate - datetime " & vbCrLf & _
                                          "           '" & Session("UserID") & "' , -- EntryUser - char(15) " & vbCrLf & _
                                          "           Getdate() , -- UpdateDate - datetime " & vbCrLf & _
                                          "           '" & Session("UserID") & "',  -- UpdateUser - char(15) " & vbCrLf & _
                                          "           '" & txtaffiliate.Text & "') "

                    End If
                    sqlRdrM.Close()
                    sqlComm = New SqlCommand(ls_SQL, cn)
                    sqlComm.ExecuteNonQuery()
                    sqlRdrM.Close()
                    'insert master

                    sqlstring = "SELECT * FROM dbo.ReceivePASI_Detail WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " AND PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' " & vbCrLf & _
                                    " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf & _
                                    " AND AffiliateID = '" & Trim(txtaffiliate.Text) & "'" & vbCrLf & _
                                    " AND PONO = '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "'" & vbCrLf & _
                                    " and KanbanNo = '" & Trim(Grid.GetRowValues(i, "colkanbanno").ToString) & "' " & vbCrLf

                    sqlComm = New SqlCommand(sqlstring, cn)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If pIsUpdate = False And (CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) <> 0 Or CDbl(Grid.GetRowValues(i, "coldefect").ToString) <> 0) Then
                        ls_SQL = ""
                        'INSERT KANBAN
                        ls_SQL = " INSERT INTO dbo.ReceivePASI_Detail " & vbCrLf & _
                                  "         ( SuratJalanNo , " & vbCrLf & _
                                  "           SupplierID , " & vbCrLf & _
                                  "           PONo , " & vbCrLf & _
                                  "           POKanbanCls , " & vbCrLf & _
                                  "           KanbanNo , " & vbCrLf & _
                                  "           PartNo , " & vbCrLf & _
                                  "           UnitCls , " & vbCrLf & _
                                  "           GoodRecQty, " & vbCrLf & _
                                  "           DefectRecQty, AffiliateID " & vbCrLf & _
								  "           ,Price " & vbCrLf & _
                                  "         ) " & vbCrLf & _
                                  " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                        ls_SQL = ls_SQL + "           '" & Trim(txtsuppliercode.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                                          "           '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "' , -- PONo - char(20) " & vbCrLf & _
                                          "           '" & pPokanban & "' , -- POKansbanCls - char(1) " & vbCrLf & _
                                          "           '" & Trim(Grid.GetRowValues(i, "colkanbanno").ToString) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                                          "           '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' , -- PartNo - char(120) " & vbCrLf & _
                                          "           '" & Trim(Grid.GetRowValues(i, "colunitcls").ToString) & "' , -- UnitCls - char(3) " & vbCrLf & _
                                          "           " & CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                                          "           " & CDbl(Grid.GetRowValues(i, "coldefect").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                                          "           '" & txtaffiliate.Text & "' , " & vbCrLf & _
                                          "           " & CDbl(Grid.GetRowValues(i, "colPrice").ToString) & vbCrLf & _
                                          "             )"
                        If ReveivingKanagata = True Then
                            Dim ls_Kanagata As String
                            Dim ls_TotalActualProd As Double = CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) + CDbl(Grid.GetRowValues(i, "coldefect").ToString)

                            ls_Kanagata = "UPDATE PASIKANAGATA.PASI_KANAGATA.dbo.PartControlDetail_AmortizationPASI " & vbCrLf & _
                                            "SET ActualProductionQty = ISNULL(ActualProductionQty,0) + " & ls_TotalActualProd & ", ActualAmortizationAmount = (ISNULL(ActualProductionQty,0) + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0)" & vbCrLf & _
                                            ", RemainingAmortizationAmount = ISNULL(TotalMoldCost,0) - ((ISNULL(ActualProductionQty,0) + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0))" & vbCrLf & _
                                            "WHERE PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "'"
                            ''7009-1862-02
                            'ls_Kanagata = "UPDATE PASIKANAGATA.PASI_KANAGATA.dbo.PartControlDetail_AmortizationPASI " & vbCrLf & _
                            '                "SET ActualProductionQty = ISNULL(ActualProductionQty,0) + " & ls_TotalActualProd & ", ActualAmortizationAmount = (ISNULL(ActualProductionQty,0) + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0)" & vbCrLf & _
                            '                ",RemainingAmortizationAmount = ISNULL(TotalMoldCost,0) - ((ISNULL(ActualProductionQty,0) + " & ls_TotalActualProd & ") * ISNULL(AmortizationPerPcs,0))" & vbCrLf & _
                            '                "WHERE PartNo = '7009-1862-02'"

                            sqlComm = New SqlCommand(ls_Kanagata, cn)
                            sqlComm.ExecuteNonQuery()
                        End If
                    ElseIf pIsUpdate = True And (CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) <> 0 Or CDbl(Grid.GetRowValues(i, "coldefect").ToString) <> 0) Then
                        'Update Data
                        ls_SQL = " Update ReceivePASI_Detail set " & vbCrLf & _
                                 " GoodRecQty = '" & CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) & "', " & vbCrLf & _
                                 " DefectRecQty = '" & CDbl(Grid.GetRowValues(i, "coldefect").ToString) & "', Price = " & CDbl(Grid.GetRowValues(i, "colPrice").ToString) & " " & vbCrLf & _
                                 " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                 " AND PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' " & vbCrLf & _
                                 " AND supplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf & _
                                 " AND AffiliateID = '" & txtaffiliate.Text & "'" & vbCrLf & _
                                 " AND PONO = '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "'" & vbCrLf & _
                                 " and KanbanNo = '" & Trim(Grid.GetRowValues(i, "colkanbanno").ToString) & "' " & vbCrLf

                        If CDbl(Grid.GetRowValues(i, "colKanbanQty").ToString) - CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) Then
                            ls_SQL = ls_SQL + " Update ReceivePASI_Remaining set " & vbCrLf & _
                                     " status = '0', " & vbCrLf & _
                                     " Qty = " & CDbl(Grid.GetRowValues(i, "colKanbanQty").ToString) & " - " & CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) & " " & vbCrLf & _
                                     " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' " & vbCrLf & _
                                     " AND supplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf & _
                                     " AND AffiliateID = '" & txtaffiliate.Text & "'" & vbCrLf & _
                                     " AND PONO = '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "'" & vbCrLf & _
                                     " and KanbanNo = '" & Trim(Grid.GetRowValues(i, "colkanbanno").ToString) & "' " & vbCrLf
                        End If
                    End If

                    sqlComm = New SqlCommand(ls_SQL, cn)
                    sqlComm.ExecuteNonQuery()
                    Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    'End If
                Next i
                'Cek apakah ada remaining
                ls_SQL = "insert into ReceivePASI_Remaining " & vbCrLf & _
                         " Select distinct * from ( " & vbCrLf & _
                         " select distinct " & vbCrLf & _
                         " suratjalan = '" & Trim(txtsuratjalanno.Text) & "', " & vbCrLf & _
                         " KD.AffiliateID, " & vbCrLf & _
                         " KD.SupplierID, " & vbCrLf & _
                         " KD.PONo, " & vbCrLf & _
                         " KD.KanbanNo, " & vbCrLf & _
                         " KD.PartNo, " & vbCrLf & _
                         " Remaining = Sum(KD.KanbanQty) - SUM(isnull(DSD.GoodRecQty,0)) " & vbCrLf & _
                         " ,'0' Status" & vbCrLf & _
                         " from Kanban_Detail KD LEFT JOIN ReceivePasi_Detail DSD " & vbCrLf & _
                         " ON KD.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                         " AND KD.SupplierID = DSD.SupplierID " & vbCrLf & _
                         " AND KD.PartNo = DSD.PartNo " & vbCrLf

                ls_SQL = ls_SQL + " AND KD.KanbanNo = DSD.KanbanNo " & vbCrLf & _
                                  " where KD.PONo IN (select distinct PONo from ReceivePasi_Detail where suratjalanno = '" & Trim(txtsuratjalanno.Text) & "' and affiliateID= '" & Trim(txtaffiliate.Text) & "' and supplierID = '" & Trim(txtsuppliercode.Text) & "') " & vbCrLf & _
                                  " AND KD.AffiliateID = '" & Trim(txtaffiliate.Text) & "'" & vbCrLf & _
                                  " AND KD.SupplierID = '" & Trim(txtsuppliercode.Text) & "'" & vbCrLf & _
                                  " AND RTRIM(KD.AFFILIATEID)+RTRIM(KD.SUPPLIERID)+RTRIM(KD.PARTNO)+RTRIM(KD.PONO)+RTRIM(KD.KANBANNO) NOT IN " & vbCrLf & _
                                  "     (select distinct RTRIM(AFFILIATEID)+RTRIM(SUPPLIERID)+RTRIM(PARTNO)+RTRIM(PONO)+RTRIM(KANBANNO) from ReceivePasi_Detail where suratjalanno = '" & Trim(txtsuratjalanno.Text) & "' and affiliateID= '" & Trim(txtaffiliate.Text) & "' and supplierID = '" & Trim(txtsuppliercode.Text) & "') " & vbCrLf & _
                                  " Group by KD.PONo, KD.PartNO, " & vbCrLf & _
                                  " KD.AffiliateID, " & vbCrLf & _
                                  " KD.SupplierID, " & vbCrLf & _
                                  " KD.KanbanNo " & vbCrLf & _
                                  " having (Sum(KD.KanbanQty) -SUM(isnull(DSD.GoodRecQty,0))) > 0 "
                ls_SQL = ls_SQL + " UNION ALL " & vbCrLf & _
                                  " select distinct " & vbCrLf & _
                                  " suratjalan = '" & Trim(txtsuratjalanno.Text) & "', " & vbCrLf & _
                                  " KD.AffiliateID, " & vbCrLf & _
                                  " KD.SupplierID, " & vbCrLf & _
                                  " KD.PONo, " & vbCrLf & _
                                  " KD.KanbanNo, " & vbCrLf & _
                                  " KD.PartNo, " & vbCrLf & _
                                  " Remaining = Sum(KD.KanbanQty) - SUM(isnull(DSD.GoodRecQty,0)) " & vbCrLf & _
                                  " ,'0' Status" & vbCrLf & _
                                  " from Kanban_Detail KD LEFT JOIN ReceivePasi_Detail DSD " & vbCrLf & _
                                  " ON KD.AffiliateID = DSD.AffiliateID " & vbCrLf & _
                                  " AND KD.SupplierID = DSD.SupplierID " & vbCrLf & _
                                  " AND KD.PartNo = DSD.PartNo " & vbCrLf

                ls_SQL = ls_SQL + " AND KD.KanbanNo = DSD.KanbanNo " & vbCrLf & _
                                  " where DSD.PONo IN (select distinct PONo from ReceivePasi_Detail where suratjalanno = '" & Trim(txtsuratjalanno.Text) & "' and affiliateID= '" & Trim(txtaffiliate.Text) & "' and supplierID = '" & Trim(txtsuppliercode.Text) & "') " & vbCrLf & _
                                  " AND DSD.AffiliateID = '" & Trim(txtaffiliate.Text) & "'" & vbCrLf & _
                                  " AND DSD.SupplierID = '" & Trim(txtsuppliercode.Text) & "'" & vbCrLf & _
                                  " Group by KD.PONo, KD.PartNO, " & vbCrLf & _
                                  " KD.AffiliateID, " & vbCrLf & _
                                  " KD.SupplierID, " & vbCrLf & _
                                  " KD.KanbanNo " & vbCrLf & _
                                  " having (Sum(KD.KanbanQty) -SUM(isnull(DSD.GoodRecQty,0))) > 0 " & vbCrLf & _
                                  " ) a "
                sqlComm = New SqlCommand(ls_SQL, cn)
                sqlComm.ExecuteNonQuery()
                'Cek apakah ada remaining
            End With
            'Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
            '    Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)


            '    sqlComm.Dispose()
            '    sqlTran.Commit()
            '    Session("SJPasi") = Trim(txtsuratjalanno.Text)
            'End Using

            cn.Close()
        End Using
        Call colorGrid()
    End Sub

    Protected Sub btndelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btndelete.Click
        Dim ls_sql As String

        ls_sql = ""
        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            'Using sqlTran As SqlTransaction = cn.BeginTransaction("Cols")
            Dim sqlComm As New SqlCommand(ls_sql, cn)
            ls_sql = "SELECT * FROM dbo.ReceivePASI_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                     " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf

            sqlComm = New SqlCommand(ls_sql, cn)
            Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

            If sqlRdrM.Read Then
                ls_sql = "delete from ReceivePASI_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                     " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf
                ls_sql = ls_sql + "Delete from ReceivePASI_Detail WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                  " AND SupplierID = '" & Trim(txtsuppliercode.Text) & "' " & vbCrLf
                sqlRdrM.Close()
                sqlComm = New SqlCommand(ls_sql, cn)
                sqlComm.ExecuteNonQuery()
                Call fillHeader("load")
                Call up_GridLoad()

                Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text

                txtdrivername.Text = ""
                txtdrivercontact.Text = ""
                txtnopol.Text = ""
                txtjenisarmada.Text = ""
                txttotalbox.Text = ""

            Else
                'data ga ada
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
            End If

            sqlComm.Dispose()
            sqlRdrM.Close()

        End Using
    End Sub

    Private Sub UpdateExcel(ByVal pIsNewData As Boolean, _
                        Optional ByVal pAffCode As String = "", _
                        Optional ByVal pSuratJalan As String = "", _
                        Optional ByVal pSuppCode As String = "")

        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim admin As String = Session("UserID").ToString

        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = " UPDATE dbo.ReceivePASI_Master " & vbCrLf & _
                          " SET ExcelCls='1'" & vbCrLf & _
                          " WHERE SuratJalanNo='" & pSuratJalan & "'  " & vbCrLf & _
                          " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
                          " AND SupplierID='" & pSuppCode & "' " & vbCrLf & _
                          "Exec Receive_UpdPrice '" & pSuratJalan & "', '" & pAffCode & "', '" & pSuppCode & "' "

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlConn.Close()
            End Using
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
        Call UpdateExcel(True, txtaffiliate.Text, txtsuratjalanno.Text, txtsuppliercode.Text)
        Call clsMsg.DisplayMessage(lblerrmessage, "1010", clsMessage.MsgType.InformationMessage)
        ButtonApprove.JSProperties("cpMessage") = lblerrmessage.Text
    End Sub

    Private Sub btnPrintGR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintGR.Click
        'If txtPASISJNo.Text <> "" Then
        Session("E02SupplierSJNo") = txtsuratjalanno.Text
        Session("E02AffiliateID") = txtaffiliate.Text
        Session("E02KanbanNo") = txtkanbanno.Text
        'Else
        '    Session("E02SupplierSJNo") = txtSupplierSJNo.Text
        'End If

        'LOG HEADER
        'Session("E02ParamPageLoad") = Trim(txtRecDate.Text) & "|" & _
        '                              Trim(txtSupplierCode.Text) & "|" & Trim(txtSupplierName.Text) & "|" & _
        '                              Trim(txtDeliveryLocationCode.Text) & "|" & Trim(txtDeliveryLocationName.Text) & "|" & _
        '                              Trim(txtSupplierSJNo.Text) & "|" & Trim(txtSupplierPlanDeliveryDate.Text) & "|" & Trim(txtSupplierDeliveryDate.Text) & "|" & _
        '                              Trim(txtPASISJNo.Text) & "|" & Trim(txtPASIDeliveryDate.Text) & "|" & _
        '                              Trim(txtDriverName.Text) & "|" & Trim(txtDriverContact.Text) & "|" & Trim(txtNoPol.Text) & "|" & Trim(txtJenisArmada.Text) & "|" & Trim(txtTotalBox.Text)

        Response.Redirect("~/Receiving/GoodReceivingReport.aspx")
    End Sub
End Class