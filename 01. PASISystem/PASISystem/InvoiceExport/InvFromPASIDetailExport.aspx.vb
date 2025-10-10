Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO

Public Class InvFromPASIDetailExport
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
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If

            '=============================================================
            'If (Not IsPostBack) AndAlso (Not IsCallback) Then
            '    If Not IsNothing(Request.QueryString("prm")) Then
            '        Dim param As String = Request.QueryString("prm").ToString

            '        If param = "  'back'" Then
            '            btnsubmenu.Text = "BACK"
            '        Else
            '            If pStatus = False Then
            '                Session("MenuDesc") = "INVOICE FROM PASI DETAIL"
            '                Session("sstatus") = "TRUE"
            '                pInvdate = Split(param, "|")(0)
            '                pAffCode = Split(param, "|")(1)
            '                pAffName = Split(param, "|")(2)
            '                pSJ = Split(param, "|")(3)
            '                pPONO = Split(param, "|")(4)
            '                pKanbanNo = Split(param, "|")(5)
            '                pInvoiceNo = Split(param, "|")(6)

            '                If pAffCode <> "" Then btnsubmenu.Text = "BACK"
            '                If pInvdate = "#1/1/1900#" Then pInvdate = Format(Now, "dd MMM yyyy")
            '                txtinvdate.Text = Format(pInvdate, "dd MMM yyyy")
            '                txtaffiliatecode.Text = pAffCode
            '                txtaffiliatename.Text = pAffName
            '                txtsuratjalanno.Text = pSJ
            '                txtkanbanno.Text = pKanbanNo
            '                txtpono.Text = pPONO
            '                txtinv.Text = pInvoiceNo

            '                pStatus = True
            '                Call fillHeader("load")
            '                Call up_GridLoad(pInvoiceNo, pAffCode, pSJ)

            '            End If
            '        End If

            '        btnsubmenu.Text = "BACK"
            '    End If
            'End If
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
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
        End Try

    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/InvPASI/InvFromPASIList.aspx")
    End Sub

    '    Private Sub fillHeader(ByVal pstatus As String)
    '        Dim ls_sql As String
    '        Dim i As Integer
    '        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

    '        Grid.JSProperties("cpDate") = Format(pInvdate, "dd MMM yyyy")
    '        Grid.JSProperties("cpScode") = pAffCode
    '        Grid.JSProperties("cpSname") = pAffName
    '        Grid.JSProperties("cpSJ") = pSJ

    '        pKanbanNo = pKanbanNo

    '        i = 0
    '        ls_sql = ""
    '        ls_sql = " select DISTINCT InvoiceDate, " & vbCrLf & _
    '                  " IM.AffiliateID, " & vbCrLf & _
    '                  " AffiliateName, " & vbCrLf & _
    '                  " SuppSJ = IM.SuratJalanNo, " & vbCrLf & _
    '                  " SupInvNo = IM.InvoiceNo, " & vbCrLf & _
    '                  " PaymentTerm , " & vbCrLf & _
    '                  " DueDate = ISNULL(CONVERT(CHAR,DueDate,106),''), " & vbCrLf & _
    '                  " kanbanNo, " & vbCrLf & _
    '                  " PoNo,isnull(totalamount,0) totalamount " & vbCrLf & _
    '                  " From InvoicePasi_Master IM Left Join InvoicePasi_Detail ID " & vbCrLf & _
    '                  " ON IM.InvoiceNo = ID.InvoiceNo "

    '        ls_sql = ls_sql + " and IM.AffiliateID = ID.AffiliateID " & vbCrLf & _
    '                          " AND IM.Suratjalanno = ID.SuratJalanNo " & vbCrLf & _
    '                          " Left Join MS_Affiliate MA ON MA.AffiliateID = IM.AffiliateID " & vbCrLf & _
    '                          " WHERE IM.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "' " & vbCrLf & _
    '                          " AND IM.InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    '                          " AND IM.AffiliateID = '" & Session("AffiliateID") & "'"

    '        Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '            cn.Open()

    '            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
    '            Dim ds As New DataSet
    '            sqlDA.Fill(ds)

    '            If ds.Tables(0).Rows.Count > 0 Then

    '                For i = 0 To ds.Tables(0).Rows.Count - 1
    '                    If pstatus = "grid" Then

    '                        Grid.JSProperties("s.cpDate") = ds.Tables(0).Rows(i)("invoiceDate")
    '                        Grid.JSProperties("cp.cpInvoiceNo") = ds.Tables(0).Rows(i)("supInvNo")
    '                        Grid.JSProperties("s.cpaffcode") = ds.Tables(0).Rows(i)("affiliateid")
    '                        Grid.JSProperties("s.cpaffname") = ds.Tables(0).Rows(i)("affiliatename")
    '                        Grid.JSProperties("s.cpsj") = ds.Tables(0).Rows(i)("suppsj")
    '                        Grid.JSProperties("s.cppayment") = ds.Tables(0).Rows(i)("paymentterm")
    '                        Grid.JSProperties("s.cpduedate") = ds.Tables(0).Rows(i)("Duedate")
    '                        Grid.JSProperties("s.cpKanbanno") = ds.Tables(0).Rows(i)("Kanbanno")
    '                        Grid.JSProperties("s.cppono") = ds.Tables(0).Rows(i)("pono")
    '                        Grid.JSProperties("s.cptotalamount") = ds.Tables(0).Rows(i)("Totalamount")
    '                    Else
    '                        txtinvdate.Value = Format(ds.Tables(0).Rows(i)("invoiceDate"), "dd MMM yyyy")
    '                        txtinv.Text = ds.Tables(0).Rows(i)("supInvNo")
    '                        txtaffiliatecode.Text = ds.Tables(0).Rows(i)("affiliateid")
    '                        txtaffiliatename.Text = ds.Tables(0).Rows(i)("affiliatename")
    '                        txtsuratjalanno.Text = ds.Tables(0).Rows(i)("suppsj")
    '                        txtpayment.Text = ds.Tables(0).Rows(i)("paymentterm")
    '                        txtduedate.Text = ds.Tables(0).Rows(i)("DueDate")
    '                        txtkanbanno.Text = ds.Tables(0).Rows(i)("Kanbanno")
    '                        txtpono.Text = ds.Tables(0).Rows(i)("pono")
    '                        txttotalamount.Text = Format(ds.Tables(0).Rows(i)("Totalamount"), "###,###,###.00")

    '                    End If
    '                Next

    '            Else
    '                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    '            End If
    '            cn.Close()
    '        End Using
    '    End Sub

    '    Private Sub up_GridLoad(ByVal pInvoice, ByVal pAff, ByVal pSJ)
    '        Dim ls_sql As String = ""

    '        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '            sqlConn.Open()
    '            ls_sql = " SELECT  DISTINCT " & vbCrLf & _
    '                  "   				no = CONVERT(Numeric,ROW_NUMBER() OVER (ORDER BY PARTNO, KanbanNo, PONO)),   " & vbCrLf & _
    '                  "   						pono,   " & vbCrLf & _
    '                  "   						pokanban,   " & vbCrLf & _
    '                  "   						kanbanno,   " & vbCrLf & _
    '                  "   						partno,   " & vbCrLf & _
    '                  "   						partname,   " & vbCrLf & _
    '                  "   						uom ,   " & vbCrLf & _
    '                  "   						qtybox,   " & vbCrLf & _
    '                  "   						pasidelqty ,   " & vbCrLf & _
    '                  "   						recqty,  						   						 "

    '            ls_sql = ls_sql + "   						invqty ,    " & vbCrLf & _
    '                              "   						delqty,   " & vbCrLf & _
    '                              "   						pasicurr,   " & vbCrLf & _
    '                              "   						pasiprice,   " & vbCrLf & _
    '                              "   						pasiamount   " & vbCrLf & _
    '                              "   				FROM (   				     " & vbCrLf & _
    '                              "      				  SELECT DISTINCT IPM.InvoiceNo,     						  " & vbCrLf & _
    '                              "      				  no = '',   " & vbCrLf & _
    '                              "   						pono = KD.PONO,   " & vbCrLf & _
    '                              "   						pokanban = (Case when ISNULL(POD.KanbanCls,'0') = '1' then 'YES' else 'NO' END),   " & vbCrLf & _
    '                              "   						kanbanno = ISNULL(KD.KanbanNo,''),    						partno = POD.PartNo,   "

    '            ls_sql = ls_sql + "   						partname = MP.PartName,   " & vbCrLf & _
    '                              "   						uom = MU.Description,   " & vbCrLf & _
    '                              "   						qtybox = MP.QtyBox,   " & vbCrLf & _
    '                              "   						pasidelqty = Round(CONVERT(CHAR,Round(ISNULL(DPD.DOQty,0),0)),0),   " & vbCrLf & _
    '                              "   						recqty=Round(convert(char, Round(Isnull(RAD.RecQty,0),0)),0),   " & vbCrLf & _
    '                              "   						invqty = Round(convert(char, Round(Isnull(IPD.INVQty,0),0),0),0),  " & vbCrLf & _
    '                              "   						delqty= Round(Convert(char,(ISNULL(DPD.DOQty,0)/MP.QtyBox)),0),   " & vbCrLf & _
    '                              "   						pasicurr=isnull(MC.Description,''),   " & vbCrLf & _
    '                              "   						pasiprice=isnull(IPD.InvPrice,0),   " & vbCrLf & _
    '                              "   						pasiamount = Isnull(IPD.INVQty,0) * Isnull(IPD.InvPrice,0)      					 " & vbCrLf & _
    '                              "   						FROM PO_DETAIL POD       "

    '            ls_sql = ls_sql + "      						 INNER JOIN PO_Master POM ON POM.AffiliateID =POD.AffiliateID      " & vbCrLf & _
    '                              "      							AND POM.SupplierID =POD.SupplierID        							 " & vbCrLf & _
    '                              "      							AND POM.PONO =POD.PONO      " & vbCrLf & _
    '                              "      						 INNER JOIN Kanban_Detail KD ON KD.AffiliateID =POD.AffiliateID      " & vbCrLf & _
    '                              "      							AND KD.SupplierID =POD.SupplierID      " & vbCrLf & _
    '                              "      							AND KD.PONO =POD.PONO      " & vbCrLf & _
    '                              "      							AND KD.PartNo =POD.PartNo      " & vbCrLf & _
    '                              "      						 INNER JOIN Kanban_Master KM ON KD.AffiliateID =KM.AffiliateID      " & vbCrLf & _
    '                              "      							AND KD.SupplierID =KM.SupplierID      " & vbCrLf & _
    '                              "      							AND KD.KanbanNo =KM.KanbanNo      " & vbCrLf & _
    '                              "                                  AND KD.DeliveryLocationCode = KM.DeliveryLocationCode          						  "

    '            ls_sql = ls_sql + "                              INNER JOIN DOSupplier_Detail DSD ON KD.AffiliateID =DSD.AffiliateID      " & vbCrLf & _
    '                              "      							AND KD.SupplierID =DSD.SupplierID        							  " & vbCrLf & _
    '                              "      							AND KD.PONO =DSD.PONO      " & vbCrLf & _
    '                              "      							AND KD.PartNo =DSD.PartNo      " & vbCrLf & _
    '                              "      							AND KD.KanbanNo =DSD.KanbanNo      " & vbCrLf & _
    '                              "      						 INNER JOIN DOSupplier_Master DSM ON DSM.AffiliateID =DSD.AffiliateID      " & vbCrLf & _
    '                              "      							AND DSM.SupplierID =DSD.SupplierID      " & vbCrLf & _
    '                              "      							AND DSM.SuratJalanNo =DSD.SuratJalanNo      " & vbCrLf & _
    '                              "      						 LEFT JOIN DOPASI_Detail DPD ON DPD.AffiliateID =KD.AffiliateID      " & vbCrLf & _
    '                              "      							AND DPD.SupplierID =KD.SupplierID      " & vbCrLf & _
    '                              "      							AND DPD.PONO =KD.PONO          							 "

    '            ls_sql = ls_sql + "      							AND KD.PartNo =DPD.PartNo      " & vbCrLf & _
    '                              "      							AND KD.KanbanNo =DPD.KanbanNo        						   " & vbCrLf & _
    '                              "                                  AND DPD.SuratJalanNoSupplier = DSM.SuratJalanNo  " & vbCrLf & _
    '                              "      						INNER JOIN DOPASI_Master DPM ON DPM.AffiliateID =DPD.AffiliateID      " & vbCrLf & _
    '                              "      							AND DPM.SupplierID =DPD.SupplierID      " & vbCrLf & _
    '                              "      							AND DPM.SuratJalanNo =DPD.SuratJalanNo      " & vbCrLf & _
    '                              "      						 LEFT JOIN ReceivePASI_Detail RPD ON RPD.AffiliateID = DPM.AffiliateID      " & vbCrLf & _
    '                              "      							AND RPD.SupplierID = DPM.SupplierID      " & vbCrLf & _
    '                              "      							AND RPD.PONo = POD.PONo      " & vbCrLf & _
    '                              "      							AND RPD.PartNo = POD.PartNo      " & vbCrLf & _
    '                              "      							AND RPD.KanbanNo = KD.KanbanNo      "

    '            ls_sql = ls_sql + "                                  AND RPD.SuratJalanNo = DSM.SuratJalanNo  " & vbCrLf & _
    '                              "      				         LEFT JOIN ReceiveAffiliate_Detail RAD ON RAD.AffiliateID = KD.AffiliateID          					        AND RAD.SupplierID = KD.SupplierID      " & vbCrLf & _
    '                              "      					        AND RAD.KanbanNo = KD.KanbanNo        					        AND RAD.PONo = KD.PONo      " & vbCrLf & _
    '                              "      					        AND RAD.PartNo = KD.PartNo      " & vbCrLf & _
    '                              "      				         LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.AffiliateID = RAD.AffiliateID      " & vbCrLf & _
    '                              "      					        AND RAM.SupplierID = RAD.SupplierID      " & vbCrLf & _
    '                              "      					        AND RAM.SuratJalanNo = RAD.SuratJalanNo      " & vbCrLf & _
    '                              "      					     LEFT JOIN InvoiceSupplier_Detail INVSD ON INVSD.SupplierID = KD.SupplierID     " & vbCrLf & _
    '                              "      							AND INVSD.AffiliateID = KD.AffiliateID       " & vbCrLf & _
    '                              "      							AND INVSD.KanbanNo = KD.kanbanNo     " & vbCrLf & _
    '                              "      							AND INVSD.PONo = KD.PONo         							 "

    '            ls_sql = ls_sql + "      							AND INVSD.PartNo = KD.PartNo      						   " & vbCrLf & _
    '                              "      					LEFT JOIN InvoiceSupplier_Master INVSM ON INVSM.InvoiceNo = INVSD.InvoiceNo     " & vbCrLf & _
    '                              "     								AND INVSM.SupplierID = INVSD.SupplierID     " & vbCrLf & _
    '                              "     								AND INVSM.AffiliateID = INVSD.AffiliateID     " & vbCrLf & _
    '                              "     								AND INVSM.SuratJalanNo = INVSD.SuratJalanNo     " & vbCrLf & _
    '                              "     						 Left Join MS_Price C ON RPD.PartNo = C.PartNo AND RAM.Receivedate between C.Startdate and C.Enddate     " & vbCrLf & _
    '                              "     									  AND POD.CurrCls = C.CurrCls and C.AffiliateID = RPD.AffiliateID   " & vbCrLf & _
    '                              "     									  AND RPD.PartNo = C.PartNo   " & vbCrLf & _
    '                              "     						 INNER JOIN dbo.InvoicePASI_Detail IPD ON RAD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
    '                              "   													AND RAD.KanbanNo = IPD.KanbanNo   " & vbCrLf & _
    '                              "   													AND RAD.PartNo = IPD.PartNo    													AND RAD.PONo = IPD.PONo   "

    '            ls_sql = ls_sql + "   													--AND RAD.SuratJalanNo = IPD.SuratJalanNo   " & vbCrLf & _
    '                              "   						INNER JOIN dbo.InvoicePASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID   " & vbCrLf & _
    '                              "   													AND IPD.InvoiceNo = IPM.InvoiceNo   " & vbCrLf & _
    '                              "   													AND IPD.SuratJalanNo = IPM.SuratJalanNo  " & vbCrLf & _
    '                              "     						 LEFT JOIn MS_CurrCls MC ON MC.CurrCls = IPD.InvCurrCls     " & vbCrLf & _
    '                              "  						 LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo      " & vbCrLf & _
    '                              "  						 LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls      " & vbCrLf & _
    '                              "                           LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = KM.AffiliateID                                   " & vbCrLf & _
    '                              "                           LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID       " & vbCrLf & _
    '                              "     		           --WHERE POD.AffiliateID = 'JAI' AND POD.pono='PO20150501-KMK '  " & vbCrLf & _
    '                              "                  WHERE isnull(IPM.InvoiceNo, '') = '" & Trim(pInvoice) & "' " & vbCrLf & _
    '                              "                     AND Isnull(IPM.AffiliateID,'') = '" & Trim(pAff) & "' " & vbCrLf & _
    '                              "                     --AND isnull(IPM.SupplierID, '') = '" & Trim(txtsupplier.Text) & "' " & vbCrLf & _
    '                              "                     AND Isnull(IPM.SuratJalanNo,'') = '" & Trim(pSJ) & "' )A "

    '            ls_sql = ls_sql + "   "


    '            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
    '            Dim ds As New DataSet
    '            sqlDA.Fill(ds)
    '            With Grid
    '                .DataSource = ds.Tables(0)
    '                .DataBind()
    '                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
    '                'Call ColorGrid()
    '            End With
    '            sqlConn.Close()

    '            If Grid.VisibleRowCount = 0 Then
    '                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    '                Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                Call colorGrid()
    '            End If
    '        End Using
    '    End Sub

    '    Private Sub colorGrid()

    '        Grid.VisibleColumns(0).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(1).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(2).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(3).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(4).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(5).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(6).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(7).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(8).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(9).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(10).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(11).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(12).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(13).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(14).CellStyle.BackColor = Color.LightYellow
    '        Grid.VisibleColumns(15).CellStyle.BackColor = Color.LightYellow

    '    End Sub

    '    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
    '        Dim pAction As String = Split(e.Parameters, "|")(0)

    '        Try
    '            Select Case pAction

    '                Case "gridload"
    '                    Call fillHeader("grid")
    '                    Call up_GridLoad(txtinv.Text, txtaffiliatecode.Text, txtsuratjalanno.Text)
    '                    If pAction = "" Then
    '                        If Grid.VisibleRowCount = 0 Then
    '                            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    '                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                        Else
    '                            Grid.JSProperties("cpMessage") = ""
    '                            lblerrmessage.Text = ""
    '                        End If
    '                    End If
    '                    Call colorGrid()
    '                Case "save"
    '                    'If Session("sstatus") Is Nothing Then Session("sstatus") = "TRUE"
    '                    'Call up_GridLoad(txtinv.Text, txtaffiliatecode.Text, txtsuratjalanno.Text)
    '                    'Call fillHeader("grid")
    '                    Call Excel()

    '                    'If Session("sstatus") = "TRUE" Then
    '                    '    Call clsMsg.DisplayMessage(lblerrmessage, "1010", clsMessage.MsgType.InformationMessage)
    '                    '    Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                    '    lblerrmessage.Text = lblerrmessage.Text
    '                    'Else
    '                    '    Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.ErrorMessage)
    '                    '    Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '                    '    lblerrmessage.Text = lblerrmessage.Text
    '                    'End If

    '                Case "kosong"

    '            End Select
    '        Catch ex As Exception
    '            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '            Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '            Grid.FocusedRowIndex = -1

    '        Finally
    '            'If (Not IsNothing(Session("YA010Msg"))) Then Grid.JSProperties("cpMessage") = Session("YA010Msg") : Session.Remove("YA010Msg")
    '            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
    '        End Try
    '    End Sub


    '    Private Sub Excel()
    '        Dim strFileSize As String = ""

    '        Dim fi As New FileInfo(Server.MapPath("~\InvPASI\INVOICE FROM PASI.xlsx"))
    '        If fi.Exists Then
    '            fi.Delete()
    '            fi = New FileInfo(Server.MapPath("~\InvPASI\INVOICE FROM PASI.xlsx"))
    '        End If
    '        Dim exl As New ExcelPackage(fi)
    '        Dim ws As ExcelWorksheet
    '        Dim space As Integer = 1
    '        ws = exl.Workbook.Worksheets.Add("Sheet1")
    '        ws.Cells(1, 1, 100, 100).Style.Font.Name = "Calibri"
    '        ws.Cells(1, 1, 100, 100).Style.Font.Size = 9

    '        With ws

    '            If Grid.VisibleRowCount > 0 Then
    '                .Cells(space, 1).Value = "NO"
    '                .Cells(space, 2).Value = "PASI INVOICE NO"
    '                .Cells(space, 3).Value = "PASI INVOICE DATE"
    '                .Cells(space, 4).Value = "PAYMENT TERM"
    '                .Cells(space, 5).Value = "DUE DATE"
    '                .Cells(space, 6).Value = "PASI INVOICE CURR"
    '                .Cells(space, 7).Value = "PASI INVOICE TOTAL AMOUNT"
    '                .Cells(space, 8).Value = "PO NO"
    '                .Cells(space, 9).Value = "PO KANBAN"
    '                .Cells(space, 10).Value = "KANBAN NO"
    '                .Cells(space, 11).Value = "PART NO"
    '                .Cells(space, 12).Value = "PART NAME"
    '                .Cells(space, 13).Value = "UOM"
    '                .Cells(space, 14).Value = "QTY/BOX"
    '                .Cells(space, 15).Value = "PASI SURAT JALAN NO"
    '                .Cells(space, 16).Value = "PASI DELIVERY QTY"
    '                .Cells(space, 17).Value = "AFFILIATE RECEIVING QTY"
    '                .Cells(space, 18).Value = "PASI INVOICE QTY"
    '                .Cells(space, 19).Value = "DELIVERY QTY (BOX)"
    '                .Cells(space, 20).Value = "PASI INVOICE"
    '                .Cells(space, 21).Value = "PASI INVOICE"
    '                .Cells(space, 22).Value = "PASI INVOICE"

    '                space = 2

    '                .Cells(space, 1).Value = "NO"
    '                .Cells(space, 2).Value = "PASI INVOICE NO"
    '                .Cells(space, 3).Value = "PASI INVOICE DATE"
    '                .Cells(space, 4).Value = "PAYMENT TERM"
    '                .Cells(space, 5).Value = "DUE DATE"
    '                .Cells(space, 6).Value = "PASI INVOICE CURR"
    '                .Cells(space, 7).Value = "PASI INVOICE TOTAL AMOUNT"
    '                .Cells(space, 8).Value = "PO NO"
    '                .Cells(space, 9).Value = "PO KANBAN"
    '                .Cells(space, 10).Value = "KANBAN NO"
    '                .Cells(space, 11).Value = "PART NO"
    '                .Cells(space, 12).Value = "PART NAME"
    '                .Cells(space, 13).Value = "UOM"
    '                .Cells(space, 14).Value = "QTY/BOX"
    '                .Cells(space, 15).Value = "PASI SURAT JALAN NO"
    '                .Cells(space, 16).Value = "PASI DELIVERY QTY"
    '                .Cells(space, 17).Value = "AFFILIATE RECEIVING QTY"
    '                .Cells(space, 18).Value = "PASI INVOICE QTY"
    '                .Cells(space, 19).Value = "DELIVERY QTY (BOX)"
    '                .Cells(space, 20).Value = "QTY"
    '                .Cells(space, 21).Value = "PRICE"
    '                .Cells(space, 22).Value = "AMOUNT"

    '                space = 1
    '                Dim iCol As Integer = 0, iNextCol As Integer = 0
    '                For iCol = 1 To 19
    '                    .Cells(space, (1 + iNextCol), space + 1, (1 + iNextCol)).Merge = True
    '                    iNextCol = iNextCol + 1
    '                Next iCol
    '                iNextCol = 0

    '                ws.Cells(space, 20, space, 22).Merge = True

    '                .Cells(space, 1, space + 1, 22).Style.WrapText = True
    '                .Cells(space, 1, space + 1, 22).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
    '                .Cells(space, 1, space + 1, 22).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

    '                space = 3
    '                For i = 0 To Grid.VisibleRowCount - 1
    '                    .Cells(i + space, 1).Value = Trim(Grid.GetRowValues(i, "no"))
    '                    .Cells(i + space, 2).Value = Trim(txtinv.Text)
    '                    .Cells(i + space, 3).Value = Trim(txtinvdate.Text)
    '                    .Cells(i + space, 4).Value = Trim(txtpayment.Text)
    '                    .Cells(i + space, 5).Value = Trim(txtduedate.Text)
    '                    .Cells(i + space, 6).Value = Trim(Grid.GetRowValues(i, "pasicurr"))
    '                    .Cells(i + space, 7).Value = Trim(txttotalamount.Text)
    '                    .Cells(i + space, 8).Value = Trim(Grid.GetRowValues(i, "pono"))
    '                    .Cells(i + space, 9).Value = Trim(Grid.GetRowValues(i, "pokanban"))
    '                    .Cells(i + space, 10).Value = Trim(Grid.GetRowValues(i, "kanbanno"))
    '                    .Cells(i + space, 11).Value = Trim(Grid.GetRowValues(i, "partno"))
    '                    .Cells(i + space, 12).Value = Trim(Grid.GetRowValues(i, "partname"))
    '                    .Cells(i + space, 13).Value = Trim(Grid.GetRowValues(i, "uom"))
    '                    .Cells(i + space, 14).Value = Trim(Grid.GetRowValues(i, "qtybox"))
    '                    .Cells(i + space, 15).Value = Trim(txtsuratjalanno.Text)
    '                    .Cells(i + space, 16).Value = Trim(Grid.GetRowValues(i, "pasidelqty"))
    '                    .Cells(i + space, 17).Value = Trim(Grid.GetRowValues(i, "recqty"))
    '                    .Cells(i + space, 18).Value = Trim(Grid.GetRowValues(i, "invqty"))
    '                    .Cells(i + space, 19).Value = Trim(Grid.GetRowValues(i, "delqty"))
    '                    .Cells(i + space, 20).Value = Trim(Grid.GetRowValues(i, "pasicurr"))
    '                    .Cells(i + space, 21).Value = Trim(Grid.GetRowValues(i, "pasiprice"))
    '                    .Cells(i + space, 22).Value = Trim(Grid.GetRowValues(i, "pasiamount"))



    '                    .Cells(i + space, 14).Style.Numberformat.Format = "#,###.00"
    '                    .Cells(i + space, 16).Style.Numberformat.Format = "#,###.00"
    '                    .Cells(i + space, 17).Style.Numberformat.Format = "#,###.00"
    '                    .Cells(i + space, 18).Style.Numberformat.Format = "#,###.00"
    '                    .Cells(i + space, 21).Style.Numberformat.Format = "#,###.00"
    '                    .Cells(i + space, 22).Style.Numberformat.Format = "#,###.00"

    '                    .Cells(i + space, 14).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
    '                    .Cells(i + space, 16).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
    '                    .Cells(i + space, 17).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
    '                    .Cells(i + space, 18).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
    '                    .Cells(i + space, 19).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
    '                    .Cells(i + space, 21).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
    '                    .Cells(i + space, 22).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

    '                    .Cells(i + space, 1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 2).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 4).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 5).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 6).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 7).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 8).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 9).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 10).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 11).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 12).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
    '                    .Cells(i + space, 13).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left


    '                Next

    '                Dim rgAll As ExcelRange = ws.Cells(1, 1, Grid.VisibleRowCount + 2, 22)
    '                EpPlusDrawAllBorders(rgAll)

    '                'save to file
    '                exl.Save()
    '            End If
    '            'redirect to file download
    '            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)
    '        End With


    '        'DrawAllBorders(ExcelSheet.Range("A1" & ": V" & i))


    '        Exit Sub
    'ErrHandler:
    '        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    '    End Sub


    '    Private Sub EpPlusDrawAllBorders(ByVal Rg As ExcelRange)
    '        With Rg
    '            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
    '            .Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
    '            .Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
    '            .Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
    '            .Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
    '        End With
    '    End Sub
    '    Private Sub DrawAllBorders(ByVal Rg As Microsoft.Office.Interop.Excel.Range)
    '        With Rg
    '            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle() = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '        End With
    '    End Sub
End Class