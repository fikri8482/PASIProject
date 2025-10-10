Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Public Class InvoiceEntry
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
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Not IsNothing(Request.QueryString("prm")) Then
                    Dim param As String = Request.QueryString("prm").ToString

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session("sstatus") = "TRUE"
                            pInvdate = Split(param, "|")(0)
                            pAffCode = Split(param, "|")(1)
                            pAffName = Split(param, "|")(2)
                            pSJ = Split(param, "|")(3)
                            pPONO = Split(param, "|")(4)
                            pKanbanNo = Split(param, "|")(5)
                            pInvoiceNo = Split(param, "|")(6)
                            pSupplier = Split(param, "|")(7)

                            If pAffCode <> "" Then btnsubmenu.Text = "BACK"
                            If pInvdate = "#1/1/1900#" Then pInvdate = Format(Now, "dd MMM yyyy")
                            txtinvdate.Text = Format(pInvdate, "dd MMM yyyy")
                            txtaffiliatecode.Text = pAffCode
                            txtaffiliatename.Text = pAffName
                            txtsuratjalanno.Text = pSJ
                            txtkanbanno.Text = pKanbanno
                            txtpono.Text = pPONO
                            txtinv.Text = pInvoiceNo
                            txtsupplier.Text = pSupplier

                            pStatus = True
                            Call fillHeader("load")
                            Call up_TotalAmount()
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
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
        End Try

    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Response.Redirect("~/Invoice/InvFromSuppList.aspx")
    End Sub

    Private Sub fillHeader(ByVal pstatus As String)
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        Grid.JSProperties("cpDate") = Format(pInvdate, "dd MMM yyyy")
        Grid.JSProperties("cpScode") = pAffCode
        Grid.JSProperties("cpSname") = pAffName
        Grid.JSProperties("cpSJ") = pSJ

        pKanbanno = pKanbanno

        i = 0

        ls_sql = ""
        ls_sql = " select InvoiceDate, " & vbCrLf & _
                  " IM.AffiliateID,IM.SupplierID, " & vbCrLf & _
                  " AffiliateName, " & vbCrLf & _
                  " SuppSJ = IM.SuratJalanNo, " & vbCrLf & _
                  " SupInvNo = IM.InvoiceNo, " & vbCrLf & _
                  " IM.PaymentTerm , " & vbCrLf & _
                  " DueDate, " & vbCrLf & _
                  " kanbanNo, " & vbCrLf & _
                  " PoNo,isnull(ID.InvAmount,0) totalamount " & vbCrLf & _
                  " From InvoiceSupplier_Master IM with(nolock) Left Join InvoiceSupplier_Detail ID with(nolock)  " & vbCrLf & _
                  " ON IM.InvoiceNo = ID.InvoiceNo "

        ls_sql = ls_sql + " and IM.AffiliateID = ID.AffiliateID " & vbCrLf & _
                          " AND IM.SupplierID = ID.SupplierID " & vbCrLf & _
                          " AND IM.Suratjalanno = ID.SuratJalanNo " & vbCrLf & _
                          " Left Join MS_Affiliate MA with(nolock) ON MA.AffiliateID = IM.AffiliateID " & vbCrLf & _
                          " Where IM.InvoiceNo = '" & Trim(txtinv.Text) & "' AND IM.AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf & _
                          " AND IM.SupplierID = '" & Trim(txtsupplier.Text) & "' AND IM.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "' "

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
                        Grid.JSProperties("cpInvoiceNo") = ds.Tables(0).Rows(i)("supInvNo")
                        'txtinv.Text = txtinv.Text
                        Grid.JSProperties("s.cpaffcode") = ds.Tables(0).Rows(i)("affiliateid")
                        Grid.JSProperties("s.cpaffname") = ds.Tables(0).Rows(i)("affiliatename")
                        Grid.JSProperties("s.cpsj") = ds.Tables(0).Rows(i)("suppsj")
                        Grid.JSProperties("s.cppayment") = ds.Tables(0).Rows(i)("paymentterm")
                        Grid.JSProperties("s.cpduedate") = ds.Tables(0).Rows(i)("Duedate")
                        Grid.JSProperties("s.cpKanbanno") = ds.Tables(0).Rows(i)("Kanbanno")
                        Grid.JSProperties("s.cppono") = ds.Tables(0).Rows(i)("pono")
                        ls_tot = ls_tot + (ds.Tables(0).Rows(i)("Totalamount"))
                        Grid.JSProperties("s.cptotalamount") = Format(ls_tot, "#,###,###.00")
                    Else
                        txtinvdate.Value = Format(ds.Tables(0).Rows(i)("invoiceDate"), "dd MMM yyyy")
                        txtinv.Text = ds.Tables(0).Rows(i)("supInvNo")
                        txtaffiliatecode.Text = ds.Tables(0).Rows(i)("affiliateid")
                        txtaffiliatename.Text = ds.Tables(0).Rows(i)("affiliatename")
                        txtsuratjalanno.Text = ds.Tables(0).Rows(i)("suppsj")
                        txtpayment.Text = ds.Tables(0).Rows(i)("paymentterm")
                        dt2.Value = Format(ds.Tables(0).Rows(i)("Duedate"), "dd MMM yyyy")
                        txtkanbanno.Text = ds.Tables(0).Rows(i)("Kanbanno")
                        txtpono.Text = ds.Tables(0).Rows(i)("pono")
                        'txttotalamount.Text = Format(txttotalamount.Text + (ds.Tables(0).Rows(i)("Totalamount")), "#,###,###.00")

                    End If
                Next

            Else
                Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
            End If
            cn.Close()
        End Using
    End Sub

    Private Sub up_TotalAmount()
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        i = 0

        ls_sql = ""
        ls_sql = " select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,tot = sum(GoodRecQty * isnull(C.Price,0)), C.currcls    " & vbCrLf & _
                  " From ReceivePasi_Detail A   with(nolock)  " & vbCrLf & _
                  " Left join ReceivePasi_Master B    with(nolock)   " & vbCrLf & _
                  "    	ON A.SuratJalanNo = B.SuratJalanNo    " & vbCrLf & _
                  " AND A.SupplierID = B.SupplierID    " & vbCrLf & _
                  " Left Join PO_detail D   with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo      							     " & vbCrLf & _
                  " Left Join MS_Price C   with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate    " & vbCrLf & _
                  "    		and C.AffiliateID = A.AffiliateID    " & vbCrLf & _
                  " WHERE c.price IS NOT null  " & vbCrLf & _
                  " 	AND A.AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'     " & vbCrLf & _
                  "     AND A.SupplierID = '" & Trim(txtsupplier.Text) & "' " & vbCrLf

        ls_sql = ls_sql + "     AND A.PONo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                          "     AND A.KanbanNo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                          "     AND A.SuratJalanNo = '" & Trim(txtsuratjalanno.Text) & "'   " & vbCrLf & _
                          " Group by A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,C.currcls "


        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            Dim ls_tot As Double
            ls_tot = 0
            txttotalamount.Text = 0
            If ds.Tables(0).Rows.Count > 0 Then
                txttotalamount.Text = Format(txttotalamount.Text + (ds.Tables(0).Rows(0)("tot")), "#,###,###.00")
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
            ls_sql = " 				SELECT  " & vbCrLf & _
                  " 				no = CONVERT(CHAR,ROW_NUMBER() OVER (ORDER BY PARTNO, KanbanNo, PONO)), " & vbCrLf & _
                  " 						pono, " & vbCrLf & _
                  " 						pokanban, " & vbCrLf & _
                  " 						kanbanno, " & vbCrLf & _
                  " 						partno, " & vbCrLf & _
                  " 						partname, " & vbCrLf & _
                  " 						uom , " & vbCrLf & _
                  " 						qtybox, " & vbCrLf & _
                  " 						suppdelqty , " & vbCrLf & _
                  " 						pasirecqty, "

            ls_sql = ls_sql + " 						suppqty , " & vbCrLf & _
                              " 						diffqty, " & vbCrLf & _
                              " 						delqty, " & vbCrLf & _
                              " 						pasicurr, " & vbCrLf & _
                              " 						pasiprice, " & vbCrLf & _
                              " 						pasiamount, " & vbCrLf & _
                              " 						suppcurr, " & vbCrLf & _
                              " 						suppprice, " & vbCrLf & _
                              " 						suppamount " & vbCrLf & _
                              " 				FROM (   				   " & vbCrLf & _
                              "    				  SELECT DISTINCT     "

            ls_sql = ls_sql + "    						no = '', " & vbCrLf & _
                              " 						pono = KD.PONO, " & vbCrLf & _
                              " 						pokanban = (Case when ISNULL(POD.KanbanCls,'0') = '1' then 'YES' else 'NO' END), " & vbCrLf & _
                              " 						kanbanno = ISNULL(KD.KanbanNo,''), " & vbCrLf & _
                              " 						partno = POD.PartNo, " & vbCrLf & _
                              " 						partname = MP.PartName, " & vbCrLf & _
                              " 						uom = MU.Description, " & vbCrLf & _
                              " 						qtybox = MPM.QtyBox, " & vbCrLf & _
                              " 						suppdelqty = Round(CONVERT(CHAR,Round(ISNULL(DSD.DOQty,0),0)),0), " & vbCrLf & _
                              " 						pasirecqty= case when deliveryByPasiCls = 1 then Round(convert(char, Round(Isnull(RPD.GoodRecQty,0),0)),0) else Round(convert(char, Round(Isnull(RAD.RecQty,0),0)),0) end, " & vbCrLf & _
                              " 						suppqty = Round(convert(char, Round(Isnull(INVSD.INVQty,0),0),0),0), "

            ls_sql = ls_sql + " 						diffqty=Round(Convert(Char, Round((ISNULL(DSD.DOQty,'0') - Isnull(INVSD.INVQty,0)),0)),0), " & vbCrLf & _
                              " 						delqty=convert(numeric(32,0),Isnull(INVSD.INVQty,0) / MPM.QtyBox), " & vbCrLf & _
                              " 						pasicurr= case when deliveryByPasiCls = 1 then MCPasi.Description else MCAff.Description end, " & vbCrLf & _
                              " 						pasiprice= case when deliveryByPasiCls = 1 then sumpasirec.Price else sumaffrec.Price end, " & vbCrLf & _
                              " 						pasiamount =case when deliveryByPasiCls = 1 then Convert(Varchar,cast(isnull(SumPasiRec.tot,0) as money),1) else Convert(Varchar,cast(isnull(SumAffRec.tot,0) as money),1) end, " & vbCrLf & _
                              " 						suppcurr=(select Description From MS_CurrCls where currcls = isnull(INVSD.InvCurrCls,'')), " & vbCrLf & _
                              " 						suppprice=Convert(numeric(32,0),Isnull(C.Price,0)), " & vbCrLf & _
                              " 						suppamount= Convert(numeric(32,0),Isnull(INVSD.INVQty,0) * isnull(C.Price,0)) " & vbCrLf & _
                              "    					FROM PO_DETAIL POD with(nolock)     " & vbCrLf & _
                              "    						 LEFT JOIN PO_Master POM with(nolock) ON POM.AffiliateID =POD.AffiliateID    " & vbCrLf & _
                              "    							AND POM.SupplierID =POD.SupplierID    "

            ls_sql = ls_sql + "    							AND POM.PONO =POD.PONO    " & vbCrLf & _
                              "    						 LEFT JOIN Kanban_Detail KD with(nolock) ON KD.AffiliateID =POD.AffiliateID    " & vbCrLf & _
                              "    							AND KD.SupplierID =POD.SupplierID    " & vbCrLf & _
                              "    							AND KD.PONO =POD.PONO    " & vbCrLf & _
                              "    							AND KD.PartNo =POD.PartNo    " & vbCrLf & _
                              "    						 LEFT JOIN Kanban_Master KM with(nolock) ON KD.AffiliateID =KM.AffiliateID    " & vbCrLf & _
                              "    							AND KD.SupplierID =KM.SupplierID    " & vbCrLf & _
                              "    							AND KD.KanbanNo =KM.KanbanNo    " & vbCrLf & _
                              "                                AND KD.DeliveryLocationCode = KM.DeliveryLocationCode    " & vbCrLf & _
                              "    						 LEFT JOIN DOSupplier_Detail DSD with(nolock) ON KD.AffiliateID =DSD.AffiliateID    " & vbCrLf & _
                              "    							AND KD.SupplierID =DSD.SupplierID    "

            ls_sql = ls_sql + "    							AND KD.PONO =DSD.PONO    " & vbCrLf & _
                              "    							AND KD.PartNo =DSD.PartNo    " & vbCrLf & _
                              "    							AND KD.KanbanNo =DSD.KanbanNo    " & vbCrLf & _
                              "    						 LEFT JOIN DOSupplier_Master DSM with(nolock) ON DSM.AffiliateID =DSD.AffiliateID    " & vbCrLf & _
                              "    							AND DSM.SupplierID =DSD.SupplierID    " & vbCrLf & _
                              "    							AND DSM.SuratJalanNo =DSD.SuratJalanNo    " & vbCrLf & _
                              "    						 LEFT JOIN DOPASI_Detail DPD with(nolock) ON DPD.AffiliateID =KD.AffiliateID    " & vbCrLf & _
                              "    							AND DPD.SupplierID =KD.SupplierID    " & vbCrLf & _
                              "    							AND DPD.PONO =KD.PONO    " & vbCrLf & _
                              "    							--AND KD.PartNo =DPD.PartNo    " & vbCrLf & _
                              "    							AND KD.KanbanNo =DPD.KanbanNo    " & vbCrLf & _
                              "                             AND DPD.SuratJalanNoSupplier = DSM.SuratJalanNo " & vbCrLf

            ls_sql = ls_sql + "    						 LEFT JOIN DOPASI_Master DPM with(nolock) ON DPM.AffiliateID =DPD.AffiliateID    " & vbCrLf & _
                              "    							AND DPM.SupplierID =DPD.SupplierID    " & vbCrLf & _
                              "    							AND DPM.SuratJalanNo =DPD.SuratJalanNo    " & vbCrLf & _
                              "    						 LEFT JOIN ReceivePASI_Detail RPD with(nolock) ON RPD.AffiliateID = DPM.AffiliateID    " & vbCrLf & _
                              "    							AND RPD.SupplierID = DPM.SupplierID    " & vbCrLf & _
                              "    							AND RPD.PONo = POD.PONo    " & vbCrLf & _
                              "    							AND RPD.PartNo = POD.PartNo    " & vbCrLf & _
                              "    							AND RPD.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "                             AND RPD.SuratJalanNo = DSM.SuratJalanNo " & vbCrLf & _
                              "    				         LEFT JOIN ReceiveAffiliate_Detail RAD with(nolock) ON RAD.AffiliateID = KD.AffiliateID    " & vbCrLf & _
                              "    					        AND RAD.SupplierID = KD.SupplierID    " & vbCrLf & _
                              "    					        AND RAD.KanbanNo = KD.KanbanNo    "

            ls_sql = ls_sql + "    					        AND RAD.PONo = KD.PONo    " & vbCrLf & _
                              "    					        AND RAD.PartNo = KD.PartNo    " & vbCrLf & _
                              "    				         LEFT JOIN ReceiveAffiliate_Master RAM with(nolock) ON RAM.AffiliateID = RAD.AffiliateID    " & vbCrLf & _
                              "    					        AND RAM.SupplierID = RAD.SupplierID    " & vbCrLf & _
                              "    					        AND RAM.SuratJalanNo = RAD.SuratJalanNo    " & vbCrLf & _
                              "    					     LEFT JOIN InvoiceSupplier_Detail INVSD with(nolock) ON INVSD.SupplierID = KD.SupplierID   " & vbCrLf & _
                              "    							AND INVSD.AffiliateID = KD.AffiliateID   " & vbCrLf & _
                              "    							--AND INVSD.SuratJalanNo = DSD.SuratJalanNo   " & vbCrLf & _
                              "    							AND INVSD.KanbanNo = KD.kanbanNo   " & vbCrLf & _
                              "    							AND INVSD.PONo = KD.PONo   " & vbCrLf & _
                              "    							AND INVSD.PartNo = KD.PartNo   "

            ls_sql = ls_sql + "   						 LEFT JOIN InvoiceSupplier_Master INVSM with(nolock) ON INVSM.InvoiceNo = INVSD.InvoiceNo   " & vbCrLf & _
                              "   							AND INVSM.SupplierID = INVSD.SupplierID   " & vbCrLf & _
                              "   							AND INVSM.AffiliateID = INVSD.AffiliateID   " & vbCrLf & _
                              "   							AND INVSM.SuratJalanNo = INVSD.SuratJalanNo   " & vbCrLf & _
                              "                          LEFT JOIN MS_ETD_PASI MEP ON INVSM.AffiliateID = MEP.AffiliateID AND CONVERT(CHAR(11),isnull(MEP.ETAAFFILIATE,''),106) = CONVERT(CHAR(11),isnull(KM.KanbanDate,''),106) " & vbCrLf & _
                              "                          LEFT JOIN MS_ETD_Supplier_Pasi ES ON MEP.ETDPASI = ES.ETAPASI " & vbCrLf & _
                              "   						 Left Join MS_Price C with(nolock) ON RPD.PartNo = C.PartNo AND ES.ETDSUPPLIER between C.Startdate and C.Enddate  " & vbCrLf & _
                              "   									  AND C.EffectiveDate >= ES.ETDSupplier " & vbCrLf & _
                              "   									  AND C.DeliveryLocationID = RPD.AffiliateID " & vbCrLf & _
                              "                                       AND INVSM.SupplierID = C.AffiliateID " & vbCrLf & _
                              "   						 LEFT JOIn MS_CurrCls MC with(nolock) ON MC.CurrCls = INVSD.InvCurrCls   " & vbCrLf & _
                              "    						 LEFT JOIN MS_Parts MP with(nolock) ON MP.PartNo = POD.PartNo    " & vbCrLf & _
                              "                          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = KD.PartNo AND MPM.AffiliateID = KD.AffiliateID AND MPM.SupplierID = KD.SupplierID " & vbCrLf & _
                              "    						 LEFT JOIN MS_UnitCls MU with(nolock) ON MU.UnitCls = MP.UnitCls    " & vbCrLf & _
                              "                             LEFT JOIN MS_Affiliate MA with(nolock)  ON MA.AffiliateID = KM.AffiliateID    "

            ls_sql = ls_sql + "                             LEFT JOIN dbo.MS_Supplier MS with(nolock) ON MS.SupplierID = POM.SupplierID     " & vbCrLf & _
                              "                             LEFT JOIN     " & vbCrLf & _
                              "     				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(RecQty),0) + ISNULL(SUM(DefectQty),0)    " & vbCrLf & _
                              "     				            FROM ReceiveAffiliate_Detail with(nolock)    " & vbCrLf & _
                              "     				           GROUP BY PONo, KanbanNo    " & vbCrLf & _
                              "     				         ) SumKanban ON SumKanban.PONo = KD.PONo AND SumKanban.KanbanNo = KD.KanbanNo    " & vbCrLf & _
                              "     				         LEFT JOIN     " & vbCrLf & _
                              "     				         (SELECT PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)    " & vbCrLf & _
                              "     				            FROM DOSupplier_Detail with(nolock)     " & vbCrLf & _
                              "     				           GROUP BY PONo, KanbanNo    " & vbCrLf & _
                              "     				         ) SumDSD ON SumDSD.PONo = KD.PONo AND SumDSD.KanbanNo = KD.KanbanNo    "

            ls_sql = ls_sql + "     				         LEFT JOIN     " & vbCrLf & _
                              "     				         (SELECT SuratJalanNoSupplier, PONo, KanbanNo, TQty = ISNULL(SUM(DOQty),0)    " & vbCrLf & _
                              "     				            FROM DOPASI_Detail with(nolock)    " & vbCrLf & _
                              "     				           GROUP BY SuratJalanNoSupplier, PONo, KanbanNo    " & vbCrLf & _
                              "     				         ) SumDPD ON SumDPD.PONo = KD.PONo AND SumDPD.KanbanNo = KD.KanbanNo AND SumDPD.SuratJalanNoSupplier = DSM.SuratJalanNo   " & vbCrLf & _
                              "     				         LEFT JOIN   " & vbCrLf & _
                              "     				         (select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,A.PartNo,tot = (GoodRecQty * isnull(C.Price,0)) ,C.currcls, c.Price  " & vbCrLf & _
                              "   							From ReceivePasi_Detail A with(nolock) Left join ReceivePasi_Master B  with(nolock)  " & vbCrLf & _
                              "   							ON A.SuratJalanNo = B.SuratJalanNo   " & vbCrLf & _
                              "   							AND A.SupplierID = B.SupplierID   " & vbCrLf & _
                              "   							Left Join PO_detail D with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo and D.SupplierID = B.SupplierID "

            ls_sql = ls_sql + "   							Left Join MS_Price C with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate   " & vbCrLf & _
                              "   									  and C.AffiliateID = A.AffiliateID " & vbCrLf & _
                              "   									  AND D.PartNo = C.PartNo  " & vbCrLf & _
                              "   							) SumPasiRec   " & vbCrLf & _
                              "   						  ON SumPasiRec.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "                             AND SumPasiRec.SuratJalanNo = DSM.SuratJalanNo " & vbCrLf & _
                              "    							AND SumPasiRec.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "    							AND SumPasiRec.PONo = POD.PONo    " & vbCrLf & _
                              "    							AND SumPasiRec.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "    							AND SumPasiRec.PartNo = KD.PartNO   " & vbCrLf & _
                              "     				         LEFT JOIN   " & vbCrLf & _
                              "     				         (select A.suratJalanno, A.supplierID, A.AffiliateID, A.PONO, A.KanbanNo,A.PartNo,tot = (RecQty * isnull(C.Price,0)) ,C.currcls, c.Price  " & vbCrLf & _
                              "   							From ReceiveAffiliate_Detail A with(nolock) Left join ReceiveAffiliate_Master B  with(nolock)  " & vbCrLf & _
                              "   							ON A.SuratJalanNo = B.SuratJalanNo   " & vbCrLf & _
                              "   							AND A.SupplierID = B.SupplierID   " & vbCrLf & _
                              "   							Left Join PO_detail D with(nolock) ON D.PoNo = A.Pono and D.PartNo = A.PartNo and D.SupplierID = B.SupplierID "

            ls_sql = ls_sql + "   							Left Join MS_Price C with(nolock) ON A.PartNo = C.PartNo AND B.Receivedate between C.Startdate and C.Enddate   " & vbCrLf & _
                              "   									  and C.AffiliateID = A.AffiliateID " & vbCrLf & _
                              "   									  AND D.PartNo = C.PartNo  " & vbCrLf & _
                              "   							) SumAffRec   " & vbCrLf & _
                              "   						  ON SumAffRec.AffiliateID = KM.AffiliateID    " & vbCrLf & _
                              "    							AND SumAffRec.SupplierID = POM.SupplierID    " & vbCrLf & _
                              "    							AND SumAffRec.PONo = POD.PONo    " & vbCrLf & _
                              "    							AND SumAffRec.KanbanNo = KD.KanbanNo  " & vbCrLf & _
                              "    							AND SumAffRec.PartNo = KD.PartNO   " & vbCrLf & _
                              "                         LEFT JOIn MS_CurrCls MCPasi   with(nolock) ON MCPasi.CurrCls = SumPasiRec.CurrCls  " & vbCrLf & _
                              "                         LEFT JOIn MS_CurrCls MCAff   with(nolock) ON MCAff.CurrCls = SumAffRec.CurrCls  " & vbCrLf & _
                              "    		           --WHERE POD.AffiliateID = 'JAI' AND POD.pono='PO20150501-KMK '    " & vbCrLf & _
                              "                  WHERE isnull(INVSM.InvoiceNo, '') = '" & Trim(txtinv.Text) & "' " & vbCrLf & _
                              "                     AND Isnull(INVSM.AffiliateID,'') = '" & Trim(txtaffiliatecode.Text) & "' " & vbCrLf & _
                              "                     AND isnull(INVSM.SupplierID, '') = '" & Trim(txtsupplier.Text) & "' " & vbCrLf & _
                              "                     AND Isnull(INVSM.SuratJalanNo,'') = '" & Trim(txtsuratjalanno.Text) & "' " & vbCrLf & _
                              "                     --AND isnull(INVSD.KanbanNo,'') = '" & Trim(txtkanbanno.Text) & "'" & vbCrLf & _
                              " )A "

            ls_sql = ls_sql + "   "


            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            sqlDA.SelectCommand.CommandTimeout = 200
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
        Grid.VisibleColumns(9).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(10).CellStyle.BackColor = Drawing.Color.White
        Grid.VisibleColumns(11).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(12).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(13).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(14).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(15).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(16).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(17).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(18).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(19).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(20).CellStyle.BackColor = Drawing.Color.LightYellow

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
        isStatusNew = False

        Session.Remove("sstatus")

        Session("sstatus") = "TRUE"
        If txttotalamount.Text = "" Then txttotalamount.Text = 0
        pReceiveDate = txtinvdate.Text

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
                Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)
                With Grid
                    For iLoop = 0 To e.UpdateValues.Count - 1
                        'cek QTY tidak boleh melebihi Qty
                        If CDbl(e.UpdateValues(iLoop).NewValues("suppqty").ToString()) > CDbl(e.UpdateValues(iLoop).NewValues("pasirecqty").ToString()) Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("sstatus") = "FALSE"
                            Exit Sub
                        End If
                        'cek QTY tidak boleh melebihi Qty

                        If Trim(e.UpdateValues(iLoop).NewValues("pokanban").ToString()) = "YES" Then pPokanban = "1" Else pPokanban = "0"

                        sqlstring = "SELECT * FROM dbo.InvoiceSupplier_Detail WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
                                    " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf & _
                                    " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("partno").ToString()) & "' " & vbCrLf

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
                            ''INSERT KANBAN
                            'ls_SQL = " INSERT INTO dbo.ReceivePASI_Detail " & vbCrLf & _
                            '          "         ( SuratJalanNo , " & vbCrLf & _
                            '          "           SupplierID , " & vbCrLf & _
                            '          "           PONo , " & vbCrLf & _
                            '          "           POKanbanCls , " & vbCrLf & _
                            '          "           KanbanNo , " & vbCrLf & _
                            '          "           PartNo , " & vbCrLf & _
                            '          "           UnitCls , " & vbCrLf & _
                            '          "           GoodRecQty, " & vbCrLf & _
                            '          "           DefectRecQty, AffiliateID " & vbCrLf & _
                            '          "         ) " & vbCrLf & _
                            '          " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                            'ls_SQL = ls_SQL + "           '" & Trim(txtaffiliatecode.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                            '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpono").ToString()) & "' , -- PONo - char(20) " & vbCrLf & _
                            '                  "           '" & pPokanban & "' , -- POKansbanCls - char(1) " & vbCrLf & _
                            '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colkanbanno").ToString()) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                            '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colpartno").ToString()) & "' , -- PartNo - char(120) " & vbCrLf & _
                            '                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("colunitcls").ToString()) & "' , -- UnitCls - char(3) " & vbCrLf & _
                            '                  "           " & CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
                            '                  "           " & CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
                            '                  "           '" & txtaffiliate.Text & "'" & vbCrLf & _
                            '                  "         ) "


                        ElseIf pIsUpdate = True Then
                            'Update Data
                            ls_SQL = " UPDATE dbo.InvoiceSupplier_Detail SET " & vbCrLf & _
                                         " InvQty = " & CDbl(e.UpdateValues(iLoop).NewValues("suppqty").ToString()) & ", " & vbCrLf & _
                                         " InvPrice = " & CDbl(e.UpdateValues(iLoop).NewValues("suppprice").ToString()) & ", " & vbCrLf & _
                                         " InvAmount = " & CDbl(e.UpdateValues(iLoop).NewValues("suppqty").ToString()) * CDbl(e.UpdateValues(iLoop).NewValues("suppprice").ToString()) & " " & vbCrLf & _
                                         " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
                                         " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("partno").ToString()) & "' " & vbCrLf
                        End If

                        sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Next iLoop
                End With

                sqlComm.Dispose()
                sqlTran.Commit()
            End Using

            cn.Close()
        End Using
        Call colorGrid()
    End Sub

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
                    Call colorGrid()
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

    'Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
    '    Dim x As Integer = CInt(e.VisibleIndex.ToString())
    '    Dim pRemaining As Double

    '    If x > Grid.VisibleRowCount Then Exit Sub
    '    If e.DataColumn.FieldName = "diffqty" Then
    '        pRemaining = e.GetValue("diffqty")
    '    End If

    '    With Grid
    '        If .VisibleRowCount > 0 Then
    '            If pRemaining > 0 Then
    '                If e.DataColumn.FieldName = "diffqty" Then
    '                    e.Cell.BackColor = Color.HotPink
    '                End If
    '            End If

    '        End If
    '    End With
    'End Sub


    Private Sub saveData()
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 1
        Dim isStatusNew As Boolean
        Dim pIsUpdate As Boolean
        Dim sqlstring As String
        Dim i As Long = 0
        Dim pInvoiceDate As Date
        Dim pPokanban As String
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
                        If CDbl(Grid.GetRowValues(i, "suppqty").ToString) > CDbl(Grid.GetRowValues(i, "pasirecqty").ToString) Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("sstatus") = "FALSE"
                            Exit Sub
                        Else
                            txtstatus.Text = "TRUE"
                        End If
                        'cek QTY tidak boleh melebihi Qty

                        If Trim(Grid.GetRowValues(i, "pokanban").ToString) = "YES" Then pPokanban = "1" Else pPokanban = "0"

                        sqlstring = "SELECT * FROM dbo.InvoiceSupplier_Detail  with(nolock) WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
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
                            'ls_SQL = ""
                            ''INSERT KANBAN
                            'ls_SQL = " INSERT INTO dbo.ReceivePASI_Detail " & vbCrLf & _
                            '          "         ( SuratJalanNo , " & vbCrLf & _
                            '          "           SupplierID , " & vbCrLf & _
                            '          "           PONo , " & vbCrLf & _
                            '          "           POKanbanCls , " & vbCrLf & _
                            '          "           KanbanNo , " & vbCrLf & _
                            '          "           PartNo , " & vbCrLf & _
                            '          "           UnitCls , " & vbCrLf & _
                            '          "           GoodRecQty, " & vbCrLf & _
                            '          "           DefectRecQty, AffiliateID " & vbCrLf & _
                            '          "         ) " & vbCrLf & _
                            '          " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                            'ls_SQL = ls_SQL + "           '" & Trim(txtaffiliatecode.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                            '                  "           '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "' , -- PONo - char(20) " & vbCrLf & _
                            '                  "           '" & pPokanban & "' , -- POKansbanCls - char(1) " & vbCrLf & _
                            '                  "           '" & Trim(Grid.GetRowValues(i, "colkanbanno").ToString) & "' , -- KanbanNo - char(20) " & vbCrLf & _
                            '                  "           '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' , -- PartNo - char(120) " & vbCrLf & _
                            '                  "           '" & Trim(Grid.GetRowValues(i, "colunitcls").ToString) & "' , -- UnitCls - char(3) " & vbCrLf & _
                            '                  "           " & CDbl(Grid.GetRowValues(i, "colreceivingqty").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                            '                  "           " & CDbl(Grid.GetRowValues(i, "coldefect").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                            '                  "           '" & txtaffiliate.Text & "') "


                        ElseIf pIsUpdate = True Then
                            'Update Data
                            ls_SQL = " Update InvoiceSupplier_Detail set " & vbCrLf & _
                                     " InvQty = " & CDbl(Grid.GetRowValues(i, "suppqty").ToString) & ", " & vbCrLf & _
                                     " InvPrice = " & CDbl(Grid.GetRowValues(i, "suppprice").ToString) & ", " & vbCrLf & _
                                     " InvAmount = " & CDbl(Grid.GetRowValues(i, "suppqty").ToString) * CDbl(Grid.GetRowValues(i, "suppprice").ToString) & " " & vbCrLf & _
                                     " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
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
        Call colorGrid()
    End Sub

    'Protected Sub btndelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btndelete.Click
    '    Dim ls_sql As String

    '    ls_sql = ""
    '    Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '        cn.Open()

    '        'Using sqlTran As SqlTransaction = cn.BeginTransaction("Cols")
    '        Dim sqlComm As New SqlCommand(ls_sql, cn)
    '        ls_sql = "SELECT * FROM dbo.InvoiceSupplier_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                 " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    '                 " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'"

    '        sqlComm = New SqlCommand(ls_sql, cn)
    '        Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

    '        If sqlRdrM.Read Then
    '            ls_sql = "delete from InvoiceSupplier_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                     " AND SupplierID = '" & Trim(txtsupplier.Text) & "' and InvoiceNo = '" & Trim(txtinv.Text) & "'" & vbCrLf & _
    '                     " AND AffiliateID = '" & Trim(txtaffiliatecode.Text) & "'" & vbCrLf
    '            ls_sql = ls_sql + "Delete from ReceivePASI_Detail WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                              " AND SupplierID = '" & Trim(txtaffiliatecode.Text) & "' " & vbCrLf
    '            sqlRdrM.Close()
    '            sqlComm = New SqlCommand(ls_sql, cn)
    '            sqlComm.ExecuteNonQuery()
    '            Call fillHeader("load")
    '            Call up_GridLoad()

    '            Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
    '            Grid.JSProperties("cpMessage") = lblerrmessage.Text

    '            txtinv.Text = ""
    '            txtpayment.Text = ""
    '            txtnopol.Text = ""
    '            txtjenisarmada.Text = ""
    '            txttotalamount.Text = ""

    '        Else
    '            'data ga ada
    '            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    '            Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '        End If

    '        sqlComm.Dispose()
    '        sqlRdrM.Close()

    '    End Using
    'End Sub
End Class