Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel

Public Class InvFromPASIDetail
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
    Dim pInvdate As Date
    Dim pAffCode As String
    Dim pAffName As String
    Dim pSJ As String
    Dim pPONO As String
    Dim pKanbanNo As String
    Dim pStatus As Boolean
    Dim pInvoiceNo As String
    Dim psupplier As String

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtHeader As DataTable
    Dim dtDetail As DataTable
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Not IsNothing(Request.QueryString("prm")) Or Not IsNothing(Session("param")) Then
                    Dim param As String = ""

                    If Not IsNothing(Request.QueryString("prm")) Then param = Request.QueryString("prm").ToString : Session("param") = Request.QueryString("prm")
                    If Not IsNothing(Session("param")) Then param = Session("param")

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session("MenuDesc") = "INVOICE FROM PASI DETAIL"
                            Session("sstatus") = "TRUE"
                            pInvdate = Split(param, "|")(0)
                            pAffCode = Split(param, "|")(1)
                            pAffName = Split(param, "|")(2)
                            pSJ = Split(param, "|")(3)
                            pPONO = Split(param, "|")(4)
                            pKanbanNo = Split(param, "|")(5)
                            pInvoiceNo = Split(param, "|")(6)

                            If pAffCode <> "" Then btnsubmenu.Text = "BACK"
                            If pInvdate = "#1/1/1900#" Then pInvdate = Format(Now, "dd MMM yyyy")
                            txtinvdate.Text = Format(pInvdate, "dd MMM yyyy")
                            txtaffiliatecode.Text = pAffCode
                            txtaffiliatename.Text = pAffName
                            txtsuratjalanno.Text = pSJ
                            txtkanbanno.Text = pKanbanNo
                            txtpono.Text = pPONO
                            txtinv.Text = pInvoiceNo

                            pStatus = True
                            Call fillHeader("load")
                            Call up_GridLoad(pInvoiceNo, pAffCode, pSJ)

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
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
        End Try

    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Session.Remove("param")
        Response.Redirect("~/InvPASI/InvFromPASIList.aspx")
    End Sub

    Private Sub fillHeader(ByVal pstatus As String)
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        Grid.JSProperties("cpDate") = Format(pInvdate, "dd MMM yyyy")
        Grid.JSProperties("cpScode") = pAffCode
        Grid.JSProperties("cpSname") = pAffName
        Grid.JSProperties("cpSJ") = pSJ

        pKanbanNo = pKanbanNo

        i = 0
        ls_sql = ""
        ls_sql = " select DISTINCT InvoiceDate = DeliveryDate, " & vbCrLf & _
                  " IM.AffiliateID, " & vbCrLf & _
                  " AffiliateName, " & vbCrLf & _
                  " SuppSJ = IM.SuratJalanNo, " & vbCrLf & _
                  " SupInvNo = IM.InvoiceNo, " & vbCrLf & _
                  " PaymentTerm = IM.PaymentTerms , " & vbCrLf & _
                  " DueDate = ISNULL(CONVERT(CHAR,dateadd(day,30,deliverydate),106),''), " & vbCrLf & _
                  " SUM(isnull(ID.DOQty,0)*isnull(MP.Price,0)) totalamount  " & vbCrLf & _
                  " From PLPASI_Master IM Left Join PLPASI_Detail ID " & vbCrLf & _
                  " ON IM.AffiliateID = ID.AffiliateID " & vbCrLf & _
                  " AND IM.Suratjalanno = ID.SuratJalanNo " & vbCrLf & _
                  " Left Join MS_Affiliate MA ON MA.AffiliateID = IM.AffiliateID " & vbCrLf & _
                  " LEFT JOIN MS_Price MP ON MP.PartNo = ID.PartNo and MP.AffiliateID = ID.AffiliateID and (DeliveryDate between StartDate and EndDate) " & vbCrLf & _
                  " WHERE IM.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "' " & vbCrLf & _
                  " AND IM.InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
                  " AND IM.AffiliateID = '" & Session("AffiliateID") & "'" & vbCrLf & _
                  " GROUP BY DeliveryDate, IM.AffiliateID, AffiliateName, IM.SuratJalanNo, IM.InvoiceNo, IM.PaymentTerms"

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If pstatus = "grid" Then

                        Grid.JSProperties("cpDate") = ds.Tables(0).Rows(i)("invoiceDate")
                        Grid.JSProperties("cpInvoiceNo") = ds.Tables(0).Rows(i)("supInvNo")
                        Grid.JSProperties("cpaffcode") = ds.Tables(0).Rows(i)("affiliateid")
                        Grid.JSProperties("cpaffname") = ds.Tables(0).Rows(i)("affiliatename")
                        Grid.JSProperties("cpsj") = ds.Tables(0).Rows(i)("suppsj")
                        Grid.JSProperties("cppayment") = ds.Tables(0).Rows(i)("paymentterm")
                        Grid.JSProperties("cpduedate") = ds.Tables(0).Rows(i)("Duedate")
                        'Grid.JSProperties("cpKanbanno") = ds.Tables(0).Rows(i)("Kanbanno")
                        'Grid.JSProperties("cppono") = ds.Tables(0).Rows(i)("pono")
                        Grid.JSProperties("cptotalamount") = ds.Tables(0).Rows(i)("Totalamount")
                    Else
                        Grid.JSProperties("cpDate") = ds.Tables(0).Rows(i)("invoiceDate")
                        Grid.JSProperties("cpInvoiceNo") = ds.Tables(0).Rows(i)("supInvNo")
                        Grid.JSProperties("cpaffcode") = ds.Tables(0).Rows(i)("affiliateid")
                        Grid.JSProperties("cpaffname") = ds.Tables(0).Rows(i)("affiliatename")
                        Grid.JSProperties("cpsj") = ds.Tables(0).Rows(i)("suppsj")
                        Grid.JSProperties("cppayment") = ds.Tables(0).Rows(i)("paymentterm")
                        Grid.JSProperties("cpduedate") = ds.Tables(0).Rows(i)("Duedate")
                        'Grid.JSProperties("cpKanbanno") = ds.Tables(0).Rows(i)("Kanbanno")
                        'Grid.JSProperties("cppono") = ds.Tables(0).Rows(i)("pono")
                        Grid.JSProperties("cptotalamount") = ds.Tables(0).Rows(i)("Totalamount")

                        txtinvdate.Value = Format(ds.Tables(0).Rows(i)("invoiceDate"), "dd MMM yyyy")
                        txtinv.Text = ds.Tables(0).Rows(i)("supInvNo")
                        txtaffiliatecode.Text = ds.Tables(0).Rows(i)("affiliateid")
                        txtaffiliatename.Text = ds.Tables(0).Rows(i)("affiliatename")
                        txtsuratjalanno.Text = ds.Tables(0).Rows(i)("suppsj")
                        txtpayment.Text = ds.Tables(0).Rows(i)("paymentterm")
                        txtduedate.Text = ds.Tables(0).Rows(i)("DueDate")
                        'txtkanbanno.Text = ds.Tables(0).Rows(i)("Kanbanno")
                        'txtpono.Text = ds.Tables(0).Rows(i)("pono")
                        txttotalamount.Text = Format(ds.Tables(0).Rows(i)("Totalamount"), "###,###,###.00")

                    End If
                Next

            Else
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            End If
            cn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad(ByVal pInvoice, ByVal pAff, ByVal pSJ)
        Dim ls_sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'ls_sql = "  SELECT  DISTINCT  " & vbcrlf & _ 
            '      "    				no = CONVERT(Numeric,ROW_NUMBER() OVER (ORDER BY PARTNO, KanbanNo, PONO)),    " & vbCrLf & _
            '      "    						pono,    " & vbCrLf & _
            '      "    						pokanban,    " & vbCrLf & _
            '      "    						kanbanno,    " & vbCrLf & _
            '      "    						partno,    " & vbCrLf & _
            '      "    						partname,    " & vbCrLf & _
            '      "    						uom ,    " & vbCrLf & _
            '      "    						qtybox,    " & vbCrLf & _
            '      "    						pasidelqty ,    " & vbCrLf & _
            '      "    						recqty,  						   						    						 "

            'ls_sql = ls_sql + "    						invqty ,     " & vbCrLf & _
            '                  "    						delqty,    " & vbCrLf & _
            '                  "    						pasicurr,    " & vbCrLf & _
            '                  "    						pasiprice,    " & vbCrLf & _
            '                  "    						pasiamount    " & vbCrLf & _
            '                  "    				FROM (   				      " & vbCrLf & _
            '                  "       				  SELECT DISTINCT IPM.InvoiceNo,     						   " & vbCrLf & _
            '                  "       				  no = '',    " & vbCrLf & _
            '                  "    						pono = KD.PONO,    " & vbCrLf & _
            '                  "    						pokanban = (Case when ISNULL(POD.KanbanCls,'0') = '1' then 'YES' else 'NO' END),    " & vbCrLf & _
            '                  "    						kanbanno = ISNULL(KD.KanbanNo,''),    						partno = POD.PartNo,      						partname = MP.PartName,    "

            'ls_sql = ls_sql + "    						uom = UC.Description,    " & vbCrLf & _
            '                  "    						qtybox = MPM.QtyBox,    " & vbCrLf & _
            '                  "    						pasidelqty = Round(CONVERT(CHAR,Round(ISNULL(PDD.DOQty,0),0)),0),    " & vbCrLf & _
            '                  "    						recqty=Round(convert(char, Round(Isnull(RAD.RecQty,0),0)),0),    " & vbCrLf & _
            '                  "    						invqty = Round(convert(char, Round(Isnull(IPD.INVQty,0),0),0),0),   " & vbCrLf & _
            '                  "    						delqty= Round(Convert(char,(ISNULL(PDD.DOQty,0)/MPM.QtyBox)),0),    " & vbCrLf & _
            '                  "    						pasicurr=isnull(MC.Description,''),    " & vbCrLf & _
            '                  "    						pasiprice=isnull(IPD.InvPrice,0),    " & vbCrLf & _
            '                  "    						pasiamount = Isnull(IPD.INVQty,0) * Isnull(IPD.InvPrice,0)      					  " & vbCrLf & _
            '                  "    						FROM    dbo.PO_Master POM     " & vbCrLf & _
            '                  "             LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID     "

            'ls_sql = ls_sql + "                                        AND POM.PoNo = POD.PONo     " & vbCrLf & _
            '                  "                                        AND POM.SupplierID = POD.SupplierID     " & vbCrLf & _
            '                  "             LEFT JOIN dbo.Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID     " & vbCrLf & _
            '                  "                                               AND KD.PoNo = POD.PONo     " & vbCrLf & _
            '                  "                                               AND KD.SupplierID = POD.SupplierID     " & vbCrLf & _
            '                  "                                               AND KD.PartNo = POD.PartNo     " & vbCrLf & _
            '                  "             LEFT JOIN dbo.Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID     " & vbCrLf & _
            '                  "                                               AND KD.KanbanNo = KM.KanbanNo                                                   " & vbCrLf & _
            '                  "                                               AND KD.SupplierID = KM.SupplierID     " & vbCrLf & _
            '                  "                                               AND KD.DeliveryLocationCode = KM.DeliveryLocationCode     " & vbCrLf & _
            '                  "             LEFT JOIN (SELECT SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls, DoQty = SUM(ISNULL(DoQty,0))    "

            'ls_sql = ls_sql + "             			FROM DOPasi_Detail GROUP BY SuratJalanno, SupplierID, AffiliateID, PONO, KanbanNO, Partno, UnitCls) PDD    " & vbCrLf & _
            '                  "                                                ON KD.AffiliateID = PDD.AffiliateID     " & vbCrLf & _
            '                  "                                                AND KD.KanbanNo = PDD.KanbanNo                                                    " & vbCrLf & _
            '                  "                                                AND KD.SupplierID = PDD.SupplierID     " & vbCrLf & _
            '                  "                                                AND KD.PartNo = PDD.PartNo     " & vbCrLf & _
            '                  "                                                AND KD.PoNo = PDD.PoNo " & vbCrLf & _
            '                  "                                                AND KD.SupplierID = PDD.SupplierID      " & vbCrLf & _
            '                  "             LEFT JOIN dbo.DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID     " & vbCrLf & _
            '                  "                                                AND PDD.SuratJalanNo = PDM.SuratJalanNo      " & vbCrLf & _
            '                  "             LEFT JOIN dbo.ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
            '                  "                                                         AND PDD.KanbanNo = RAD.KanbanNo    "

            'ls_sql = ls_sql + "                                                         AND PDD.SupplierID = RAD.SupplierID    " & vbCrLf & _
            '                  "                                                         AND PDD.PartNo = RAD.PartNo " & vbCrLf & _
            '                  "                                                         AND PDD.PoNo = RAD.PoNo    " & vbCrLf & _
            '                  "             LEFT JOIN dbo.ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo    " & vbCrLf & _
            '                  "                                                         AND RAM.AffiliateID = RAD.AffiliateID     " & vbCrLf & _
            '                  "    		 INNER JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID    " & vbCrLf & _
            '                  "    													AND RAD.KanbanNo = IPD.KanbanNo    " & vbCrLf & _
            '                  "    													AND RAD.PartNo = IPD.PartNo    " & vbCrLf & _
            '                  "    													AND RAD.PONo = IPD.PONo     " & vbCrLf & _
            '                  "    													AND RAD.SuratJalanNo = PDD.SuratJalanNo " & vbCrLf & _
            '                  "    		 INNER JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID    "

            'ls_sql = ls_sql + "    													AND IPD.InvoiceNo = IPM.InvoiceNo      													 " & vbCrLf & _
            '                  "    													AND IPD.SuratJalanNo = IPM.SuratJalanNo    " & vbCrLf & _
            '                  "             LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo     " & vbCrLf & _
            '                  "             LEFT JOIN dbo.MS_PartMapping MPM ON MPM.PartNo = POD.PartNo AND MPM.SupplierID = POD.SupplierID AND MPM.AffiliateID = POD.AffiliateID " & vbCrLf & _
            '                  "             LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls     " & vbCrLf & _
            '                  "             LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID     " & vbCrLf & _
            '                  "             LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID     " & vbCrLf & _
            '                  "             LEFT JOIN dbo.MS_DeliveryPlace MD ON MD.DeliveryLocationCode = KM.DeliveryLocationCode  " & vbCrLf & _
            '                  "             LEFT JOIn MS_CurrCls MC ON MC.CurrCls = IPD.InvCurrCls   " & vbCrLf & _
            '                  "         WHERE isnull(IPM.InvoiceNo, '') = '" & Trim(pInvoice) & "'  )A "

            'ls_sql = ls_sql + "   "
			
			ls_sql = "sp_Affiliate_InvFromPASIDetail_GridLoad"
            Dim cmd As New SqlCommand(ls_sql, sqlConn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("InvoiceNo", pInvoice)
            Dim sqlDA As New SqlDataAdapter(cmd)

            
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With Grid
                .DataSource = ds.Tables(0)
                .DataBind()
                'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
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

        Grid.VisibleColumns(0).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(1).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(2).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(3).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(4).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(5).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(6).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(7).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(8).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(9).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(10).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(11).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(12).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(13).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(14).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(15).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(16).CellStyle.BackColor = Color.LightYellow
        Grid.VisibleColumns(17).CellStyle.BackColor = Color.LightYellow

    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)

        Try
            Select Case pAction

                Case "gridload"
                    Call fillHeader("grid")
                    Call up_GridLoad(txtinv.Text, txtaffiliatecode.Text, txtsuratjalanno.Text)
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
                Case "save"
                    Call Excel()
                    txtinv.Text = txtinv.Text
                Case "kosong"
                Case "bc40"
                    Call ExcelBC40()
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

    Private Sub HeaderExcelBC40()

        Dim ds As New DataSet
        Dim ls_sql As String = ""

        ls_sql = " SELECT DISTINCT " & vbCrLf & _
                  " IzinTPB = (Select Rtrim(IzinTPB) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " BCPerson = (Select Rtrim(BCPerson) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " KantorPabean = (Select Rtrim(KantorPabean) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " NPWP = (Select Rtrim(NPWP) from MS_Affiliate where AffiliateID = PLM.AffiliateID), " & vbCrLf & _
                  " Buyer = (Select Rtrim(AffiliateName) from MS_Affiliate where AffiliateID = PLM.AffiliateID),  " & vbCrLf & _
                  " AlamatBuyer =(Select Rtrim(Address) from MS_Affiliate where AffiliateID = PLM.AffiliateID),   " & vbCrLf & _
                  " City =(Select Rtrim(City) from MS_Affiliate where AffiliateID = PLM.AffiliateID),   " & vbCrLf & _
                  " NPWPPengirim = (Select Rtrim(NPWP) from MS_Affiliate where AffiliateID = 'PASI'), " & vbCrLf & _
                  " Pengirim = (Select Rtrim(AffiliateName) from MS_Affiliate where AffiliateID = 'PASI'),  " & vbCrLf & _
                  " AlamatPengirim =(Select Rtrim(Address) from MS_Affiliate where AffiliateID = 'PASI'),   " & vbCrLf & _
                  " ShipCls = Rtrim(isnull(POM.ShipCls,'')), " & vbCrLf & _
                  " NoPol = Rtrim(isnull(PLM.NoPol,'')),   " & vbCrLf & _
                  " InvoiceNo = Rtrim(coalesce(PLM.InvoiceNo,'-')),   " & vbCrLf & _
                  " Invdate = Coalesce(DPM.DeliveryDate, DSM.DeliveryDate),  PLD.PONo, PODate = POM.EntryDate, " & vbCrLf

        ls_sql = ls_sql + " Currency = MC.Description, " & vbCrLf & _
                          " JumlahHarga = (SELECT SUM(DPD.Price*DOQty) FROM PLPASI_Detail PLPD " & vbCrLf & _
                          " LEFT JOIN MS_Price MPr1 ON MPr1.PartNo = PLPD.PartNo  " & vbCrLf & _
                          " AND PLPD.AffiliateID = MPr1.AffiliateID WHERE PLPD.SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND PLPD.AffiliateID='" & Trim(txtaffiliatecode.Text) & "'), " & vbCrLf & _
                          " JumlahKemasan =  (SELECT SUM(CONVERT(NUMERIC,ISNULL(CartonQty,0))) FROM PLPASI_Detail WHERE SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND AffiliateID='" & Trim(txtaffiliatecode.Text) & "'), " & vbCrLf & _
                          " JumlahQty =  (SELECT SUM(CONVERT(NUMERIC,ISNULL(DOQty,0))) FROM PLPASI_Detail WHERE SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND AffiliateID='" & Trim(txtaffiliatecode.Text) & "'), " & vbCrLf & _
                          " BeratBersih = (SELECT SUM(ISNULL(CartonQty,0) * (b.NetWeight/1000)) FROM PLPASI_Detail a left join MS_Parts b on a.PartNo = b.PartNo WHERE SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND AffiliateID='" & Trim(txtaffiliatecode.Text) & "'), " & vbCrLf & _
                          " BeratKotor = (SELECT SUM(ISNULL(CartonQty,0) * (b.GrossWeight/1000)) FROM PLPASI_Detail a left join MS_Parts b on a.PartNo = b.PartNo WHERE SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND AffiliateID='" & Trim(txtaffiliatecode.Text) & "') " & vbCrLf & _
                          " FROM PLPASI_Detail PLD " & vbCrLf & _
                          " LEFT JOIN PLPASI_Master PLM ON PLM.SuratJalanNo = PLD.SuratJalanNo AND PLM.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
                          " LEFT JOIN PO_Master POM ON POM.PONo = PLD.PONo AND POM.AffiliateID = PLD.AffiliateID AND POM.SupplierID = PLD.SupplierID " & vbCrLf & _
                          " LEFT JOIN DOPasi_Detail DPD  ON DPD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DPD.SupplierID = PLD.SupplierID   " & vbCrLf

        ls_sql = ls_sql + "   	AND DPD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DPD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOPASI_Master DPM  ON DPM.SuratJalanNo = DPD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DPD.SupplierID = DPM.SupplierID   " & vbCrLf & _
                          "   	AND DPD.AffiliateID = DPM.AffiliateID     " & vbCrLf & _
                          " LEFT JOIN DOSupplier_Detail DSD ON DSD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DSD.SupplierID = PLD.SupplierID   " & vbCrLf & _
                          "   	AND DSD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DSD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOSUPPLIER_Master DSM ON DSM.SuratJalanNo = DSD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DSD.SupplierID = DSM.SupplierID   " & vbCrLf

        ls_sql = ls_sql + "   	AND DSD.AffiliateID = DSM.AffiliateID   " & vbCrLf & _
                          " LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo  " & vbCrLf & _
                          " LEFT JOIN MS_Price MPr ON MPr.PartNo = PLD.PartNo AND PLD.AffiliateID = MPr.AffiliateID " & vbCrLf & _
                          " 	AND MPR.PartNo = PLD.PartNo and COALESCE(DPM.DeliveryDate,DSM.DeliveryDate) between MPR.StartDate and MPR.EndDate " & vbCrLf & _
                          " LEFT JOIN MS_CurrCls MC ON MC.CurrCls = MPr.CurrCls " & vbCrLf

        ls_sql = ls_sql + " WHERE PLM.SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND PLM.AffiliateID='" & Trim(txtaffiliatecode.Text) & "' " & vbCrLf & _
                          " GROUP BY POM.EntryDate, PLD.PONo, PLM.AffiliateID,PLM.ViaDelivery,PLM.InvoiceNo,POM.ShipCls,DPM.DeliveryDate,DSM.DeliveryDate,MC.DESCRIPTION,PLM.NoPol " & vbCrLf


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtHeader = ds.Tables(0)
    End Sub

    Private Sub DetailExcelBC40()

        Dim ds As New DataSet
        Dim ls_sql As String = ""

        ls_sql = " SELECT DISTINCT " & vbCrLf & _
                  " Row =  ROW_NUMBER() OVER (ORDER BY PLD.PartNo), " & vbCrLf & _
                  " PartNo = Rtrim(MP.PartNo), PartName = Rtrim(MP.PartName), " & vbCrLf & _
                  " Qty =  CONVERT(NUMERIC,ISNULL(PLD.DOQty,0)), " & vbCrLf & _
                  " Currency = MC.Description, " & vbCrLf & _
                  " Harga = DPD.Price*PLD.DOQty " & vbCrLf & _
                  " FROM PLPASI_Detail PLD " & vbCrLf & _
                  " LEFT JOIN PLPASI_Master PLM ON PLM.SuratJalanNo = PLD.SuratJalanNo  " & vbCrLf & _
                  " 	AND PLM.AffiliateID = PLD.AffiliateID  " & vbCrLf & _
                  " LEFT JOIN MS_Parts MP ON MP.PartNo = PLD.PartNo  " & vbCrLf

        ls_sql = ls_sql + " LEFT JOIN MS_Price MPr ON MPr.PartNo = PLD.PartNo  " & vbCrLf & _
                          " 	AND PLD.AffiliateID = MPr.AffiliateID " & vbCrLf & _
                          " LEFT JOIN MS_CurrCls MC ON MC.CurrCls = MPr.CurrCls " & vbCrLf & _
                          " LEFT JOIN PO_Master POM ON POM.PONo = PLD.PONo  " & vbCrLf & _
                          " 	AND POM.AffiliateID = PLD.AffiliateID AND POM.SupplierID = PLD.SupplierID " & vbCrLf & _
                          " LEFT JOIN DOPasi_Detail DPD  ON DPD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DPD.SupplierID = PLD.SupplierID   " & vbCrLf & _
                          "   	AND DPD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DPD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOPASI_Master DPM  ON DPM.SuratJalanNo = DPD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DPD.SupplierID = DPM.SupplierID   " & vbCrLf

        ls_sql = ls_sql + "   	AND DPD.AffiliateID = DPM.AffiliateID     " & vbCrLf & _
                          " LEFT JOIN DOSupplier_Detail DSD ON DSD.SuratJalanNo = PLD.SuratJalanNo   " & vbCrLf & _
                          "   	AND DSD.SupplierID = PLD.SupplierID   " & vbCrLf & _
                          "   	AND DSD.AffiliateID = PLD.AffiliateID   " & vbCrLf & _
                          "   	AND DSD.PONo = PLD.PONo   " & vbCrLf & _
                          " LEFT JOIN DOSUPPLIER_Master DSM ON DSM.SuratJalanNo = DSD.SuratJalanNo  	  " & vbCrLf & _
                          "   	AND DSD.SupplierID = DSM.SupplierID   " & vbCrLf & _
                          "   	AND DSD.AffiliateID = DSM.AffiliateID   " & vbCrLf

        ls_sql = ls_sql + " WHERE PLM.SuratJalanNo='" & Trim(txtsuratjalanno.Text) & "' AND PLM.AffiliateID='" & Trim(txtaffiliatecode.Text) & "' " & vbCrLf & _
                          " GROUP BY PLD.PartNo,MP.PartNo,MP.PartName,PLD.DOQty,MC.DESCRIPTION,DPD.Price" & vbCrLf


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.Fill(ds)
            sqlConn.Close()
        End Using
        dtDetail = ds.Tables(0)
    End Sub

    Private Sub ExcelBC40()
        Call HeaderExcelBC40()
        Call DetailExcelBC40()
        FileName = "Template BC4.0.xlsx"
        FilePath = Server.MapPath("~\Template\" & FileName)
        Call epplusExportHeaderExcel(FilePath, "", dtHeader, "A:17", "")

        'Call epplusExportExcel(FilePath, "KEDUA", dtDetail, "A:7", "")
    End Sub

    Private Sub epplusExportHeaderExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "BC 4.0 " & Trim(txtinv.Text) & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\InvPASI\" & tempFile)

            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet
            Dim wsCount As ExcelWorksheet

            ws = exl.Workbook.Worksheets("form (2)")
            Dim irow As Integer = 0
            Dim rowstart As Integer

            With ws
                If pData.Rows.Count > 0 Then
                    .Cells("C5").Value = ": " & pData.Rows(irow)("KantorPabean") 'KantorPabean
                    .Cells("C11").Value = ": " & pData.Rows(irow)("NPWP")
                    .Cells("C12").Value = ": " & pData.Rows(irow)("Buyer")
                    .Cells("C13:E14").Merge = True
                    .Cells("C13:E14").Style.WrapText = True
                    .Cells("C13:E14").Value = ": " & pData.Rows(irow)("AlamatBuyer")
                    .Cells("I12").Value = "" & pData.Rows(irow)("NPWPPengirim")
                    .Cells("I13").Value = ": " & pData.Rows(irow)("Pengirim")
                    .Cells("I14:J16").Merge = True
                    .Cells("I14:J16").Style.WrapText = True
                    .Cells("I14:J16").Value = ": " & pData.Rows(irow)("AlamatPengirim")

                    .Cells("C18").Value = ": " & pData.Rows(irow)("InvoiceNo")
                    .Cells("E18").Value = Format(pData.Rows(irow)("Invdate"), "dd-MMM-yyyy")

                    .Cells("C21").Value = ": " & pData.Rows(irow)("PONo")
                    .Cells("E21").Value = Format(pData.Rows(irow)("PODate"), "dd-MMM-yyyy")

                    .Cells("B23").Value = "Jenis Sarana Pengangkut darat : " & pData.Rows(irow)("ShipCls")
                    .Cells("I23").Value = ": " & pData.Rows(irow)("NoPol")

                    .Cells("C25").Value = pData.Rows(irow)("JumlahHarga")
                    If .Cells("C25").Value <> "0" Then
                        .Cells("C25").Style.Numberformat.Format = "#,###"
                    End If

                    .Cells("H28").Value = pData.Rows(irow)("JumlahKemasan")

                    .Cells("H31").Value = Format(pData.Rows(irow)("BeratKotor"), "###,##0.0#") & " Kg"
                    .Cells("J31").Value = Format(pData.Rows(irow)("BeratBersih"), "###,##0.0#") & " Kg"

                    .Cells("I54").Value = pData.Rows(irow)("City")

                    .Cells("J54").FormulaR1C1 = "=R[-36]C[-5]"
                    .Cells("H60").Value = "(.........." & pData.Rows(irow)("BCPerson") & "..........)"

                End If
            End With

            ws = exl.Workbook.Worksheets("Lembar Lanjutan (2)")
            With ws
                If pData.Rows.Count > 0 Then
                    .Cells("C8").Value = ": " & pData.Rows(irow)("KantorPabean")                   
                    .Cells("F62").Value = pData.Rows(irow)("City")
                    .Cells("G62").FormulaR1C1 = "='form (2)'!R[-8]C[3]" ' "='form (2)'!R[-8]C[3]"
                    .Cells("F65").Value = "(.........." & pData.Rows(irow)("BCPerson") & "..........)"
                End If
            End With



            wsCount = exl.Workbook.Worksheets("form (2)")
            For irow = 0 To dtDetail.Rows.Count - 1
                If irow <= 12 Then
                    rowstart = 38
                    'Sheet pertama muat 13 item
                    For icol = 1 To dtDetail.Columns.Count
                        wsCount.Cells("A" & irow + rowstart).Value = dtDetail.Rows(irow)("Row")
                        wsCount.Cells("B" & irow + rowstart).Value = dtDetail.Rows(irow)("PartName")
                        wsCount.Cells("D" & irow + rowstart).Value = dtDetail.Rows(irow)("PartNo")
                        wsCount.Cells("H" & irow + rowstart).Value = Format(dtDetail.Rows(irow)("Qty"), "#,##0") & " PCS"
                        wsCount.Cells("J" & irow + rowstart).Value = dtDetail.Rows(irow)("Harga")
                        If wsCount.Cells("J" & irow + rowstart).Value <> "0" Then
                            'wsCount.Cells("J" & irow + rowstart).Style.Numberformat.Format = "#,##0.00"
                            wsCount.Cells("J" & irow + rowstart).Style.Numberformat.Format = "_([$Rp-421]* #,##0_);_([$Rp-421]* (#,##0);_([$Rp-421]* ""-""_);_(@_)"
                        End If
                    Next
                Else
                    'Sheet kedua muat 26 item
                    rowstart = 17 - 13
                    wsCount = exl.Workbook.Worksheets("Lembar Lanjutan (2)")
                    For icol = 1 To dtDetail.Columns.Count
                        wsCount.Cells("A" & irow + rowstart).Value = dtDetail.Rows(irow)("Row")
                        wsCount.Cells("B" & irow + rowstart).Value = dtDetail.Rows(irow)("PartName")
                        wsCount.Cells("C" & irow + rowstart).Value = dtDetail.Rows(irow)("PartNo")
                        wsCount.Cells("F" & irow + rowstart).Value = Format(dtDetail.Rows(irow)("Qty"), "#,##0") & " PCS"
                        wsCount.Cells("H" & irow + rowstart).Value = dtDetail.Rows(irow)("Harga")
                        If wsCount.Cells("H" & irow + rowstart).Value <> "0" Then
                            'wsCount.Cells("H" & irow + rowstart).Style.Numberformat.Format = "#,##0.00"
                            wsCount.Cells("H" & irow + rowstart).Style.Numberformat.Format = "_([$Rp-421]* #,##0_);_([$Rp-421]* (#,##0);_([$Rp-421]* ""-""_);_(@_)"
                        End If
                    Next
                End If
                'If irow = 0 Then 'ROW 39 SHEET PERTAMA
                '    iRowTmp = 21
                'ElseIf irow < 29 Then 'NEXT ROW SHEET PERTAMA
                '    iRowTmp = iRowTmp + 1
                'ElseIf irow = 29 Then 'NEXT ROW SHEET KEDUA
                '    wsCount = exl.Workbook.Worksheets("KEDUA")
                '    rowstart = Split(pCellStart, ":")(1) - 3
                '    iRowTmp = 1
                'ElseIf irow > 29 And irow < 55 Then 'NEXT ROW SHEET KEDUA
                '    iRowTmp = iRowTmp + 1
                'ElseIf irow = 55 Then 'NEXT ROW SHEET KETIGA
                '    wsCount = exl.Workbook.Worksheets("KETIGA")
                '    rowstart = Split(pCellStart, ":")(1) - 3
                '    iRowTmp = 0
                'ElseIf irow > 55 Then
                '    iRowTmp = iRowTmp + 1
                'End If

                'For icol = 1 To dtDetail.Columns.Count
                '    wsCount.Cells("A" & iRowTmp + rowstart + 1).Value = dtDetail.Rows(irow)("Row")
                '    wsCount.Cells("B" & iRowTmp + rowstart + 1).Value = dtDetail.Rows(irow)("Description")
                '    wsCount.Cells("F" & iRowTmp + rowstart + 1).Value = dtDetail.Rows(irow)("Qty")
                '    wsCount.Cells("F" & iRowTmp + rowstart + 1).Style.Numberformat.Format = "#,###"
                '    wsCount.Cells("G" & iRowTmp + rowstart + 1).Value = "PCS"
                '    wsCount.Cells("H" & iRowTmp + rowstart + 1).Value = dtDetail.Rows(irow)("Currency")
                '    wsCount.Cells("I" & iRowTmp + rowstart + 1).Value = dtDetail.Rows(irow)("Harga")
                '    If wsCount.Cells("I" & iRowTmp + rowstart + 1).Value <> "0" Then
                '        wsCount.Cells("I" & iRowTmp + rowstart + 1).Style.Numberformat.Format = "#,##0.00"
                '    End If
                'Next
                'TotalQty = TotalQty + dtDetail.Rows(irow)("Qty")
                'TotalPrice = TotalPrice + dtDetail.Rows(irow)("Harga")
            Next
            'iRowTmp = iRowTmp + 1
            'wsCount.Cells("F" & iRowTmp + rowstart + 1).Formula = TotalQty '"=SUM(F16:F" & irow + rowstart & ")"
            'wsCount.Cells("F" & iRowTmp + rowstart + 1).Style.Numberformat.Format = "#,###"
            'wsCount.Cells("G" & iRowTmp + rowstart + 1).Value = "PCS"
            'wsCount.Cells("F" & iRowTmp + rowstart + 1).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin

            'wsCount.Cells("H" & iRowTmp + rowstart + 1).Value = dtDetail.Rows(irow - 1)("Currency")
            'wsCount.Cells("I" & iRowTmp + rowstart + 1).Formula = TotalPrice '"=SUM(I16:I" & irow + rowstart & ")"
            'If wsCount.Cells("I" & iRowTmp + rowstart + 1).Value <> "0" Then
            '    wsCount.Cells("I" & iRowTmp + rowstart + 1).Style.Numberformat.Format = "#,##0.00"
            'End If
            'wsCount.Cells("I" & iRowTmp + rowstart + 1).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            'wsCount.Workbook.CalcMode = ExcelCalcMode.Automatic

            exl.Save()

            'DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)
            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\InvPASI\" & tempFile & "")
            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub epplusExportDetailExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try


            Dim NewFileName As String = Server.MapPath("~\InvPASI\" & pFilename)
            If (System.IO.File.Exists(pFileName)) Then
                System.IO.File.Copy(pFileName, NewFileName, True)
            End If


            'Dim NewFileName As String = Server.MapPath("~\Template\TemplateBC40.xlsx")
            'If (System.IO.File.Exists(pFilename)) Then
            '    System.IO.File.Copy(pFilename, NewFileName, True)
            'End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets(pSheetName)
            Dim irow As Integer = 0
            Dim icol As Integer = 0

            With ws
                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count
                        .Cells(irow + rowstart + 1, icol).Value = pData.Rows(irow)(icol - 1)
                    Next
                Next

                'ALIGNMENT
                .Cells(rowstart + 1, icol, irow, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(iRow + space, colKanbanSeqNo).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                '.Cells(irow + rowstart + 1, iCol, irow + rowstart, colSupplierDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIDelDate).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPartName).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                '.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvQty).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
                '.Cells(irow + rowstart + 1, iCol, irow + rowstart, colPASIInvCurr).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                .Cells(irow + rowstart + 1, icol).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub Excel()
        Dim strFileSize As String = ""
        Dim pFileName As String
        Dim FileNames As String

        FileNames = "INVOICE FROM PASI" & txtinv.Text.Trim & ".xlsx"
        pFileName = Server.MapPath("~\Template\INVOICE FROM PASI.xlsx")

        Dim NewFileName As String = Server.MapPath("~\InvPASI\" & FileNames)
        If (System.IO.File.Exists(pFilename)) Then
            System.IO.File.Copy(pFileName, NewFileName, True)
        End If

        Dim fi As New FileInfo(NewFileName)

        Dim exl As New ExcelPackage(fi)
        Dim ws As ExcelWorksheet
        Dim space As Integer = 1
        ws = exl.Workbook.Worksheets("Sheet1")
        ws.Cells(1, 1, 100, 100).Style.Font.Name = "Arial"
        ws.Cells(1, 1, 100, 100).Style.Font.Size = 9

        With ws

            If Grid.VisibleRowCount > 0 Then
                .Cells(3, 3).Value = ":" & txtinv.Text.Trim
                .Cells(4, 3).Value = ":" & txtsuratjalanno.Text.Trim
                .Cells(5, 3).Value = ":" & txtpayment.Text.Trim
                .Cells(6, 3).Value = ":" & txtduedate.Text.Trim

                space = 10
                For i = 0 To Grid.VisibleRowCount - 1
                    .Cells(i + space, 1).Value = Trim(Grid.GetRowValues(i, "no"))
                    .Cells(i + space, 2).Value = Trim(Grid.GetRowValues(i, "pono"))
                    .Cells(i + space, 3).Value = Trim(Grid.GetRowValues(i, "kanbanno"))
                    .Cells(i + space, 4).Value = Trim(Grid.GetRowValues(i, "pokanban"))
                    .Cells(i + space, 5).Value = Trim(Grid.GetRowValues(i, "partno"))
                    .Cells(i + space, 6).Value = Trim(Grid.GetRowValues(i, "partname"))
                    .Cells(i + space, 7).Value = Trim(Grid.GetRowValues(i, "uom"))
                    .Cells(i + space, 8).Value = CDbl(Trim(Grid.GetRowValues(i, "qtybox")))
                    .Cells(i + space, 9).Value = CDbl(Trim(Grid.GetRowValues(i, "pasidelqty")))
                    .Cells(i + space, 10).Value = CDbl(Trim(Grid.GetRowValues(i, "recqty")))
                    .Cells(i + space, 11).Value = CDbl(Trim(Grid.GetRowValues(i, "invqty")))
                    .Cells(i + space, 12).Value = CDbl(Trim(Grid.GetRowValues(i, "delqty")))
                    .Cells(i + space, 13).Value = Trim(Grid.GetRowValues(i, "pasicurr"))
                    .Cells(i + space, 14).Value = CDbl(Trim(Grid.GetRowValues(i, "pasiprice")))
                    .Cells(i + space, 15).Value = CDbl(Trim(Grid.GetRowValues(i, "pasiamount")))
                    .Cells(i + space, 16).Value = CDbl(Trim(Grid.GetRowValues(i, "netweight")))
                    .Cells(i + space, 17).Value = CDbl(Trim(Grid.GetRowValues(i, "grossweight")))

                    .Cells(i + space, 8).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 9).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 10).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 11).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 12).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 14).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 15).Style.Numberformat.Format = "#,###"
                    .Cells(i + space, 16).Style.Numberformat.Format = "#,##0.00"
                    .Cells(i + space, 17).Style.Numberformat.Format = "#,##0.00"
                Next

                'Dim rgAll As ExcelRange = ws.Cells(10, 1, Grid.VisibleRowCount + 2, 15)
                Dim rgAll As ExcelRange = ws.Cells(10, 1, Grid.VisibleRowCount + 2 + 10 - 3, 17)
                EpPlusDrawAllBorders(rgAll)

                'save to file
                exl.Save()
            End If
            'redirect to file download
            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\InvPASI\" + fi.Name)
            txtinv.Text = txtinv.Text
        End With

        
        'DrawAllBorders(ExcelSheet.Range("A1" & ": V" & i))


        Exit Sub
ErrHandler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub


    Private Sub EpPlusDrawAllBorders(ByVal Rg As ExcelRange)
        With Rg
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
        End With
    End Sub
    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
        With Rg
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
    End Sub

    Protected Sub btnprint_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnprint.Click
        Session("PrintSJ") = txtsuratjalanno.Text
        Session("PrintAffID") = txtaffiliatecode.Text
        'Session("PrintSuppID") = txtSupplierCode.Text

        Response.Redirect("~/InvPASI/PackingListViewReport.aspx")
    End Sub
End Class