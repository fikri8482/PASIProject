Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Public Class ReceivingEntryExport
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
    Dim pSuratJalanNoSupplier As String
    Dim pOrderNo As String
    Dim pAffiliateName As String
    Dim pSuratjalanNo As String
    Dim pPlandelivery As Date
    Dim pDeldate As Date
    Dim pKanbanno As String
    Dim pStatus As Boolean
    Dim pAffiliate As String
    Dim ppono As String
    Dim pOrder As String
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not String.IsNullOrEmpty(Request.QueryString("id"))) Then
                Session("M01Url") = Request.QueryString("Session")
            End If
            ''=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                If Not IsNothing(Request.QueryString("prm")) Then
                    Dim param As String = Request.QueryString("prm").ToString
                    Session("E02ParamPageLoad") = Request.QueryString("prm").ToString()

                    If param = "  'back'" Then
                        btnsubmenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session("MenuDesc") = "RECEIVING ENTRY"
                            Session("sstatus") = "TRUE"
                            pOrder = Split(param, "|")(0)
                            pSuratJalanNoSupplier = Split(param, "|")(1)
                            pOrderNo = Split(param, "|")(2)
                            pAffiliate = Split(param, "|")(3)
                            pAffiliateName = Split(param, "|")(4)
                            ppono = Split(param, "|")(5)

                            Session("PONO") = ppono
                            Session("AFFID") = pAffiliate
                            Session("SuratJalanNoSupplier") = pSuratJalanNoSupplier
                            Session("OrderNo") = pOrder

                            If pSuratJalanNoSupplier <> "" Then btnsubmenu.Text = "BACK"

                            pStatus = True
                            Call fillHeader("load")
                            Call up_GridLoad()
                            FillCombo()

                        End If
                    End If

                    btnsubmenu.Text = "BACK"
                End If
            End If
            '===============================================================================
            btnsubmenu.Text = "BACK"
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblerrmessage.Text = ""
                'dt1.Value = Format(txtkanbandate.text, "MMM yyyy")
            End If

            'Call colorGrid()

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, False, False, False, False, True)
        End Try

    End Sub

    Protected Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnsubmenu.Click
        Session.Remove("PONO")
        Session.Remove("AFFID")
        Session.Remove("SuratJalanNoSupplier")
        Session.Remove("OrderNo")
        'Response.Redirect("~/DeliveryExport/DeliveryToAffListExport.aspx")

        If btnsubmenu.Text = "BACK" Then
            Response.Redirect("~/DeliveryExport/DeliveryToAffListExport.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If
    End Sub

    Private Sub fillHeader(ByVal pstatus As String)
        Dim ls_sql As String
        Dim i As Integer
        Dim sqlcom As New SqlCommand(clsGlobal.ConnectionString)

        i = 0
        ls_sql = ""
        'ls_sql = "     SELECT   " & vbCrLf & _
        '          "     Receiveddate, " & vbCrLf & _
        '          "     Supplier, " & vbCrLf & _
        '          "     SupplierName, " & vbCrLf & _
        '          "     SupplierPlan, " & vbCrLf & _
        '          "     SupplierDelivery, " & vbCrLf & _
        '          "     SuratJalan, " & vbCrLf & _
        '          "     DriverName, " & vbCrLf & _
        '          "     DriverContact, " & vbCrLf & _
        '          "     Nopol, " & vbCrLf & _
        '          "     JenisArmada, "

        'ls_sql = ls_sql + "     totalbox = isnull(CEILING(SUM(totalbox)),0),PerformanceCls, PerformanceName, " & vbCrLf & _
        '                  "     AffiliateID, " & vbCrLf & _
        '                  "     AffiliateName, " & vbCrLf & _
        '                  "     DeliveryLocation, " & vbCrLf & _
        '                  "     DeliveryLocationName " & vbCrLf & _
        '                  "      " & vbCrLf & _
        '                  "     FROM(     " & vbCrLf & _
        '                  "      SELECT  " & vbCrLf & _
        '                  " 			Receiveddate = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(RM.ReceiveDate,'')), 106)  , " & vbCrLf & _
        '                  " 			Supplier = ISNULL(DSM.SupplierID,''), " & vbCrLf & _
        '                  " 			SupplierName = ISNULL(MS.SupplierName,''), " & vbCrLf

        'ls_sql = ls_sql + " 			SupplierPlan = CONVERT(CHAR(12), CONVERT(DATETIME, POM.ETDVendor), 106) , " & vbCrLf & _
        '                  " 			SupplierDelivery = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(DSM.DeliveryDate,'')), 106) , " & vbCrLf & _
        '                  " 			SuratJalan = ISNULL(DSM.SuratJalanNo,''), " & vbCrLf & _
        '                  " 			DriverName = ISNULL(DSM.DriverName,''), " & vbCrLf & _
        '                  " 			DriverContact = ISNULL(DSM.DriverContact,''), " & vbCrLf & _
        '                  " 			Nopol = ISNULL(DSM.Nopol,''), " & vbCrLf & _
        '                  " 			JenisArmada = ISNULL(DSM.JenisArmada,''), " & vbCrLf & _
        '                  " 			totalbox = Coalesce(RD.GoodRecQty, DSD.DOQty) / MPM.Qtybox, " & vbCrLf & _
        '                  " 			PerformanceCls = '' , PerformanceName = '', " & vbCrLf & _
        '                  " 			AffiliateID =  ISNULL(DSM.AffiliateID,''), " & vbCrLf & _
        '                  " 			AffiliateName = ISNULL(MA.AffiliateName,''), " & vbCrLf

        'ls_sql = ls_sql + " 			DeliveryLocation = POM.ForwarderID, " & vbCrLf & _
        '                  " 			DeliveryLocationName = ISNULL(MF.ForwarderName,'') " & vbCrLf & _
        '                  "                " & vbCrLf & _
        '                  "      FROM    DOSupplier_Detail_Export DSD  " & vbCrLf & _
        '                  "          LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo  " & vbCrLf & _
        '                  "                                                    AND DSM.AffiliateID = DSD.AffiliateID  " & vbCrLf & _
        '                  "                                                    AND DSM.SupplierID = DSD.SupplierID  " & vbCrLf & _
        '                  "                                                    AND DSM.PONO = DSD.PONO  " & vbCrLf & _
        '                  "          LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO  " & vbCrLf & _
        '                  "                                            AND POD.AffiliateID = DSM.AffiliateID  " & vbCrLf & _
        '                  "                                            AND POD.SupplierID = DSM.SupplierID  "

        'ls_sql = ls_sql + "                                            AND POD.PartNo = DSD.PartNo  " & vbCrLf & _
        '                  "          LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo  " & vbCrLf & _
        '                  "                                                  AND DSD.affiliateID = RM.affiliateID  " & vbCrLf & _
        '                  "                                                  AND DSD.SupplierID = RM.SupplierID  " & vbCrLf & _
        '                  "          LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo  " & vbCrLf & _
        '                  "                                                  AND RM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
        '                  "                                                  AND RM.SupplierID = RD.SupplierID  " & vbCrLf & _
        '                  "                                                  AND RM.PONO = RD.PONO  " & vbCrLf & _
        '                  "                                                  AND DSD.PartNo = RD.PartNo  " & vbCrLf & _
        '                  "                                                  AND DSD.PONO = RD.PONO  " & vbCrLf & _
        '                  "  		LEFT JOIN (Select *, OrderNO = OrderNo1, ETDVendor = ETDVendor1, ETAPort = ETAPort1, ETAFactory = ETAFactory1  "

        'ls_sql = ls_sql + "  					from Po_Master_Export  " & vbCrLf & _
        '                  "  				    UNION ALL   " & vbCrLf & _
        '                  "  				   Select *, OrderNO = OrderNo2, ETDVendor = ETDVendor2, ETAPort = ETAPort2, ETAFactory = ETAFactory2  " & vbCrLf & _
        '                  "  					from Po_Master_Export  " & vbCrLf & _
        '                  "  					UNION ALL   " & vbCrLf & _
        '                  "  				   Select *, OrderNO = OrderNo3, ETDVendor = ETDVendor3, ETAPort = ETAPort3, ETAFactory = ETAFactory3  " & vbCrLf & _
        '                  "  					from Po_Master_Export  " & vbCrLf & _
        '                  "  					UNION ALL   " & vbCrLf & _
        '                  "  				   Select *, OrderNO = OrderNo4, ETDVendor = ETDVendor4, ETAPort = ETAPort4, ETAFactory = ETAFactory4  " & vbCrLf & _
        '                  "  					from Po_Master_Export) POM ON POM.PONO = POD.PONO  " & vbCrLf & _
        '                  "  										  AND POM.AffiliateID = POD.AffiliateID  "

        'ls_sql = ls_sql + "  										  AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
        '                  "  		  " & vbCrLf & _                          
        '                  "  		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID   " & vbCrLf & _
        '                  "  		LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID  "

        'ls_sql = ls_sql + "  		LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID  " & vbCrLf & _
        '                  "  		LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo  " & vbCrLf & _
        '                  "         LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID and MPM.PartNo = RD.PartNo " & vbCrLf & _
        '                  "  		LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls     " & vbCrLf & _
        '                  "  WHERE DSM.SuratJalanNo = '" & Session("SuratJalanNoSupplier") & "'  " & vbCrLf & _
        '                  "  AND DSM.AffiliateID = '" & Session("AFFID") & "'  " & vbCrLf & _
        '                  "  AND POM.OrderNo = '" & Session("OrderNo") & "'" & vbCrLf & _
        '                  "  " & vbCrLf & _
        '                  "  )x  " & vbCrLf & _
        '                  "  	GROUP BY  " & vbCrLf & _
        '                  "      Receiveddate, " & vbCrLf & _
        '                  "     Supplier, " & vbCrLf & _
        '                  "     SupplierName, "

        'ls_sql = ls_sql + "     SupplierPlan, " & vbCrLf & _
        '                  "     SupplierDelivery, " & vbCrLf & _
        '                  "     SuratJalan, " & vbCrLf & _
        '                  "     DriverName, " & vbCrLf & _
        '                  "     DriverContact, " & vbCrLf & _
        '                  "     Nopol, " & vbCrLf & _
        '                  "     JenisArmada, " & vbCrLf & _
        '                  "     PerformanceCls, PerformanceName, " & vbCrLf & _
        '                  "     AffiliateID, " & vbCrLf & _
        '                  "     AffiliateName, " & vbCrLf & _
        '                  "     DeliveryLocation, "

        'ls_sql = ls_sql + "     DeliveryLocationName "
        ls_sql = "      SELECT    " & vbCrLf & _
                  "      Receiveddate,  " & vbCrLf & _
                  "      Supplier,  " & vbCrLf & _
                  "      SupplierName,  " & vbCrLf & _
                  "      SupplierPlan,  " & vbCrLf & _
                  "      SupplierDelivery,  " & vbCrLf & _
                  "      SuratJalan,  " & vbCrLf & _
                  "      DriverName,  " & vbCrLf & _
                  "      DriverContact,  " & vbCrLf & _
                  "      Nopol,  " & vbCrLf & _
                  "      JenisArmada, " & vbCrLf

        ls_sql = ls_sql + " 	 totalbox = isnull(CEILING(SUM(totalbox)),0), " & vbCrLf & _
                          " 	 PerformanceCls,  " & vbCrLf & _
                          " 	 PerformanceName,  " & vbCrLf & _
                          "      AffiliateID,  " & vbCrLf & _
                          "      AffiliateName,  " & vbCrLf & _
                          "      DeliveryLocation,  " & vbCrLf & _
                          "      DeliveryLocationName  " & vbCrLf & _
                          "        " & vbCrLf & _
                          "      FROM(      " & vbCrLf & _
                          "       SELECT   " & vbCrLf & _
                          "  			Receiveddate = CONVERT(CHAR(12), CONVERT(DATETIME, ISNULL(RM.ReceiveDate,'')), 106)  ,  " & vbCrLf

        ls_sql = ls_sql + "  			Supplier = ISNULL(RM.SupplierID,''),  " & vbCrLf & _
                          "  			SupplierName = ISNULL(MS.SupplierName,''),  " & vbCrLf & _
                          "  			SupplierPlan = CONVERT(CHAR(12), CONVERT(DATETIME, POM.ETDVendor), 106) ,  " & vbCrLf & _
                          "  			SupplierDelivery = ISNULL(CONVERT(CHAR(12), CONVERT(DATETIME, DSM.DeliveryDate), 106),''),  " & vbCrLf & _
                          "  			SuratJalan = ISNULL(RM.SuratJalanNo,''),  " & vbCrLf & _
                          "  			DriverName = ISNULL(RM.DriverName,''),  " & vbCrLf & _
                          "  			DriverContact = ISNULL(RM.DriverContact,''),  " & vbCrLf & _
                          "  			Nopol = ISNULL(RM.Nopol,''),  " & vbCrLf & _
                          "  			JenisArmada = ISNULL(RM.JenisArmada,''),  " & vbCrLf & _
                          "  			totalbox = Coalesce(RD.GoodRecQty, DSD.DOQty) / ISNULL(POD.POQtyBox,MPM.Qtybox),  " & vbCrLf & _
                          "  			PerformanceCls = '' , PerformanceName = '',  " & vbCrLf

        ls_sql = ls_sql + "  			AffiliateID =  ISNULL(RM.AffiliateID,''),  " & vbCrLf & _
                          "  			AffiliateName = ISNULL(MA.AffiliateName,''),  " & vbCrLf & _
                          "  			DeliveryLocation = POM.ForwarderID,  " & vbCrLf & _
                          "  			DeliveryLocationName = ISNULL(MF.ForwarderName,'')  " & vbCrLf & _
                          "                  " & vbCrLf & _
                          "       FROM    ReceiveForwarder_Detail RD " & vbCrLf & _
                          "           LEFT JOIN ReceiveForwarder_Master RM ON RM.SuratJalanNo = RD.SuratjalanNo   " & vbCrLf & _
                          "                                                     AND RM.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                          "                                                     AND RM.SupplierID = RD.SupplierID   " & vbCrLf & _
                          "                                                     AND RM.PONO = RD.PONO   " & vbCrLf & _
                          "           LEFT JOIN po_detail_Export POD ON POD.PONO = RM.PONO   " & vbCrLf

        ls_sql = ls_sql + "                                             AND POD.AffiliateID = RM.AffiliateID   " & vbCrLf & _
                          "                                             AND POD.SupplierID = RM.SupplierID    " & vbCrLf & _
                          " 											AND POD.PartNo = RD.PartNo   " & vbCrLf & _
                          "           LEFT JOIN DOSupplier_Master_Export DSM ON RD.suratJalanNo = DSM.SuratJalanNo   " & vbCrLf & _
                          "                                                   AND RD.affiliateID = DSM.affiliateID   " & vbCrLf & _
                          "                                                   AND RD.SupplierID = DSM.SupplierID   " & vbCrLf & _
                          "           LEFT JOIN DOSupplier_Detail_Export DSD ON DSM.SuratJalanNo = DSD.SuratjalanNo   " & vbCrLf & _
                          "                                                   AND DSM.AffiliateID = DSD.AffiliateID   " & vbCrLf & _
                          "                                                   AND DSM.SupplierID = DSD.SupplierID   " & vbCrLf & _
                          "                                                   AND DSM.PONO = DSD.PONO   " & vbCrLf & _
                          "                                                   AND RD.PartNo = DSD.PartNo   " & vbCrLf

        ls_sql = ls_sql + "                                                   AND RD.PONO = DSD.PONO   " & vbCrLf & _
                          "   		LEFT JOIN (Select *, OrderNO = OrderNo1, ETDVendor = ETDVendor1, ETAPort = ETAPort1, ETAFactory = ETAFactory1   " & vbCrLf & _
                          " 					from Po_Master_Export   " & vbCrLf & _
                          "   				    UNION ALL    " & vbCrLf & _
                          "   				   Select *, OrderNO = OrderNo2, ETDVendor = ETDVendor2, ETAPort = ETAPort2, ETAFactory = ETAFactory2   " & vbCrLf & _
                          "   					from Po_Master_Export   " & vbCrLf & _
                          "   					UNION ALL    " & vbCrLf & _
                          "   				   Select *, OrderNO = OrderNo3, ETDVendor = ETDVendor3, ETAPort = ETAPort3, ETAFactory = ETAFactory3   " & vbCrLf & _
                          "   					from Po_Master_Export   " & vbCrLf & _
                          "   					UNION ALL    " & vbCrLf & _
                          "   				   Select *, OrderNO = OrderNo4, ETDVendor = ETDVendor4, ETAPort = ETAPort4, ETAFactory = ETAFactory4   " & vbCrLf

        ls_sql = ls_sql + "   					from Po_Master_Export) POM ON POM.PONO = POD.PONO   " & vbCrLf & _
                          "   										  AND POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          " 										  AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                          "   		   " & vbCrLf & _
                          "   		LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = RM.AffiliateID    " & vbCrLf & _
                          "   		LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID    		 " & vbCrLf & _
                          " 		LEFT JOIN ms_supplier MS ON MS.SupplierID = RM.SupplierID   " & vbCrLf & _
                          "   		LEFT JOIN MS_Parts MP ON MP.PartNo = RD.PartNo   " & vbCrLf & _
                          "          LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID and MPM.PartNo = RD.PartNo  " & vbCrLf & _
                          "   		LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls      " & vbCrLf & _
                          "   WHERE RM.SuratJalanNo = '" & Session("SuratJalanNoSupplier") & "'   " & vbCrLf

        ls_sql = ls_sql + "   AND RM.AffiliateID = '" & Session("AFFID") & "'   " & vbCrLf & _
                          "   AND POM.OrderNo = '" & Session("OrderNo") & "' " & vbCrLf & _
                          "    " & vbCrLf & _
                          "   )x   " & vbCrLf & _
                          "   	GROUP BY  " & vbCrLf & _
                          "      Receiveddate,  " & vbCrLf & _
                          "      Supplier,  " & vbCrLf & _
                          "      SupplierName, " & vbCrLf & _
                          " 	 SupplierPlan,  " & vbCrLf & _
                          "      SupplierDelivery,  " & vbCrLf & _
                          "      SuratJalan,  " & vbCrLf

        ls_sql = ls_sql + "      DriverName,  " & vbCrLf & _
                          "      DriverContact,  " & vbCrLf & _
                          "      Nopol,  " & vbCrLf & _
                          "      JenisArmada,  " & vbCrLf & _
                          "      PerformanceCls, PerformanceName,  " & vbCrLf & _
                          "      AffiliateID,  " & vbCrLf & _
                          "      AffiliateName,  " & vbCrLf & _
                          "      DeliveryLocation, " & vbCrLf & _
                          " 	 DeliveryLocationName  "



        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If pstatus = "grid" Then
                        'Grid.JSProperties("s.cpreceivedate") = ds.Tables(0).Rows(i)("Receiveddate")
                        txtrecdate.Text = Trim(ds.Tables(0).Rows(i)("Receiveddate"))
                        'Grid.JSProperties("s.cpsuppliercode") = ds.Tables(0).Rows(i)("Supplier")
                        Grid.JSProperties("s.cpsuppliername") = ds.Tables(0).Rows(i)("SupplierName")
                        Grid.JSProperties("s.cpsuratjalanno") = ds.Tables(0).Rows(i)("SuratJalan")
                        Grid.JSProperties("s.cpclscode") = "" 'ds.Tables(0).Rows(i)("DriverName")
                        Grid.JSProperties("s.cpclsname") = "" 'ds.Tables(0).Rows(i)("DriverName")
                        Grid.JSProperties("s.cpdrivername") = ds.Tables(0).Rows(i)("DriverName")
                        Grid.JSProperties("s.cpdrivercontact") = ds.Tables(0).Rows(i)("Drivercontact")
                        Grid.JSProperties("s.cpnopol") = ds.Tables(0).Rows(i)("Nopol")
                        Grid.JSProperties("s.cpjenisarmada") = ds.Tables(0).Rows(i)("JenisArmada")
                        Grid.JSProperties("s.cpplan") = ds.Tables(0).Rows(i)("SupplierPlan")
                        Grid.JSProperties("s.cpdeliverydate") = ds.Tables(0).Rows(i)("SupplierDelivery")
                        Grid.JSProperties("s.cpaffiliatecode") = ds.Tables(0).Rows(i)("AffiliateID")
                        Grid.JSProperties("s.cpaffiliatename") = ds.Tables(0).Rows(i)("AffiliateName")
                        Grid.JSProperties("s.cpdeliverycode") = ds.Tables(0).Rows(i)("DeliveryLocation")
                        Grid.JSProperties("s.cpdeliveryname") = ds.Tables(0).Rows(i)("DeliveryLocationName")
                        Grid.JSProperties("s.cptotalbox") = ds.Tables(0).Rows(i)("totalbox")

                    Else

                        If Trim(ds.Tables(0).Rows(i)("Receiveddate")) = "01 Jan 1900" Then
                            txtrecdate.Text = Format(Now, "dd MMM yyyy")
                        Else
                            txtrecdate.Text = Trim(ds.Tables(0).Rows(i)("Receiveddate"))
                        End If
                        txtsupp.Text = ds.Tables(0).Rows(i)("Supplier")
                        txtsuppliername.Text = ds.Tables(0).Rows(i)("SupplierName")
                        txtsuratjalanno.Text = ds.Tables(0).Rows(i)("SuratJalan")
                        cbocls.Text = "" 'ds.Tables(0).Rows(i)("DriverName")
                        txtcls.Text = "" 'ds.Tables(0).Rows(i)("DriverName")
                        txtdrivername.Text = ds.Tables(0).Rows(i)("DriverName")
                        txtdrivercontact.Text = ds.Tables(0).Rows(i)("Drivercontact")
                        txtnopol.Text = ds.Tables(0).Rows(i)("Nopol")
                        txtjenisarmada.Text = ds.Tables(0).Rows(i)("JenisArmada")
                        txtplandeliverydate.Text = ds.Tables(0).Rows(i)("SupplierPlan")
                        txtsupplierdeliverydate.Text = ds.Tables(0).Rows(i)("SupplierDelivery")
                        txtaffiliatecode.Text = ds.Tables(0).Rows(i)("AffiliateID")
                        txtaffiliate.Text = ds.Tables(0).Rows(i)("AffiliateName")
                        txtdeliverycode.Text = ds.Tables(0).Rows(i)("DeliveryLocation")
                        txtdeliveryname.Text = ds.Tables(0).Rows(i)("DeliveryLocationName")
                        txttotalbox.Text = ds.Tables(0).Rows(i)("totalbox")
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
            'ls_sql = " SELECT  colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colOrderNO, colpartno, LabelNo1, LabelNo2 )) , " & vbCrLf & _
            '      "             * " & vbCrLf & _
            '      "     FROM    ( SELECT DISTINCT " & vbCrLf & _
            '      "                         idx = 0 , " & vbCrLf & _
            '      "                         colpilih = 0 , " & vbCrLf & _
            '      "                         colorderno = RD.OrderNo , " & vbCrLf & _
            '      "                         collabelno = '' , " & vbCrLf & _
            '      "                         colpartno = RD.partno , " & vbCrLf & _
            '      "                         colpartname = MP.PartName , " & vbCrLf & _
            '      "                         coluom = ISNULL(MU.DESCRIPTION, '') , " & vbCrLf & _
            '      "                         colqtybox = ISNULL(MPM.QtyBox, 0) , "

            'ls_sql = ls_sql + "                         coldelqty = ISNULL(DSD.DOQty, 0) , " & vbCrLf & _
            '                  "                         colgoodreceiving = ISNULL(MPM.QtyBox, 0)*ISNULL(RB.box, 0) , " & vbCrLf & _
            '                  "                         coldefectreceiving = 0 , " & vbCrLf & _
            '                  "                         coldefect = 0 , " & vbCrLf & _
            '                  "                         colreceivingbox = ISNULL(RB.box, 0) , " & vbCrLf & _
            '                  "                         colHgood = ISNULL(MPM.QtyBox, 0)*ISNULL(RB.box, 0) , " & vbCrLf & _
            '                  "                         colHdefect = 0 , " & vbCrLf & _
            '                  "                         colpono = ISNULL(POD.PONo, '') , " & vbCrLf & _
            '                  "                         LabelNo1 = ISNULL(RTRIM(RB.Label1), '') , " & vbCrLf & _
            '                  "                         LabelNo2 = ISNULL(RTRIM(RB.Label2), '') , " & vbCrLf & _
            '                  "                         PART = RD.PartNo , "

            'ls_sql = ls_sql + "                         PO = POM.PONo , " & vbCrLf & _
            '                  "                         LABEL = ISNULL(RTRIM(RB.Label1), '') + '-' " & vbCrLf & _
            '                  "                         + ISNULL(RTRIM(RB.Label2), '') , " & vbCrLf & _
            '                  "                         StatusDefect = RB.StatusDefect " & vbCrLf & _
            '                  "               FROM      DOSupplier_Detail_Export DSD " & vbCrLf & _
            '                  "                         LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo " & vbCrLf & _
            '                  "                                                               AND DSM.AffiliateID = DSD.AffiliateID " & vbCrLf & _
            '                  "                                                               AND DSM.SupplierID = DSD.SupplierID " & vbCrLf & _
            '                  "                                                               AND DSM.PONO = DSD.PONO " & vbCrLf & _
            '                  "                         LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO " & vbCrLf & _
            '                  "                                                           AND POD.AffiliateID = DSM.AffiliateID "

            'ls_sql = ls_sql + "                                                           AND POD.SupplierID = DSM.SupplierID " & vbCrLf & _
            '                  "                                                           AND POD.PartNo = DSD.PartNo " & vbCrLf & _
            '                  "                         LEFT JOIN ( SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo1 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor1 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort1 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory1 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                     UNION ALL " & vbCrLf & _
            '                  "                                     SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo2 , "

            'ls_sql = ls_sql + "                                             ETDVendor = ETDVendor2 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort2 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory2 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                     UNION ALL " & vbCrLf & _
            '                  "                                     SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo3 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor3 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort3 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory3 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export "

            'ls_sql = ls_sql + "                                     UNION ALL " & vbCrLf & _
            '                  "                                     SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo4 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor4 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort4 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory4 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                   ) POM ON POM.PONO = POD.PONO " & vbCrLf & _
            '                  "                                            AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
            '                  "                                            AND POM.SupplierID = POD.SupplierID " & vbCrLf & _                             
            '                  "                         LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo " & vbCrLf & _
            '                  "                                                               AND DSD.affiliateID = RM.affiliateID " & vbCrLf & _
            '                  "                                                               AND DSD.SupplierID = RM.SupplierID "

            'ls_sql = ls_sql + "                         LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo " & vbCrLf & _
            '                  "                                                               AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  "                                                               AND RM.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "                                                               AND RM.PONO = RD.PONO " & vbCrLf & _
            '                  "                                                               AND DSD.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                                                               AND DSD.PONO = RD.PONO " & vbCrLf & _
            '                  "                         LEFT JOIN ( SELECT  suratjalanno , " & vbCrLf & _
            '                  "                                             supplierid , " & vbCrLf & _
            '                  "                                             affiliateID , " & vbCrLf & _
            '                  "                                             PONO , " & vbCrLf & _
            '                  "                                             partno , "

            'ls_sql = ls_sql + "                                             goodRecQty = SUM(ISNULL(goodRecQty, " & vbCrLf & _
            '                  "                                                               0)) , " & vbCrLf & _
            '                  "                                             DefectRecQty = SUM(ISNULL(DefectRecQty, " & vbCrLf & _
            '                  "                                                               0)) " & vbCrLf & _
            '                  "                                     FROM    ReceiveForwarder_Detail " & vbCrLf & _
            '                  "                                     GROUP BY suratjalanno , " & vbCrLf & _
            '                  "                                             supplierid , " & vbCrLf & _
            '                  "                                             affiliateID , " & vbCrLf & _
            '                  "                                             PONO , " & vbCrLf & _
            '                  "                                             partno " & vbCrLf & _
            '                  "                                   ) REM ON REM.SuratJalanNo = RD.SuratjalanNo "

            'ls_sql = ls_sql + "                                            AND REM.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  "                                            AND REM.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "                                            AND REM.PONO = RD.PONO " & vbCrLf & _
            '                  "                                            AND REM.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                                            AND REM.PONO = RD.PONO " & vbCrLf & _
            '                  "                         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID " & vbCrLf & _
            '                  "                         LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID " & vbCrLf & _
            '                  "                         LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID " & vbCrLf & _
            '                  "                         LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo " & vbCrLf & _
            '                  "                         LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID " & vbCrLf & _
            '                  "                                                         AND MPM.PartNo = RD.PartNo "

            'ls_sql = ls_sql + "                         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls " & vbCrLf & _
            '                  "                         LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                  "                                                               AND RB.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  "                                                               AND RB.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "                                                               AND RB.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                                                               AND RB.StatusDefect = '0' " & vbCrLf & _
            '                  "                                                               AND RB.PONo = RD.PONo " & vbCrLf & _
            '                  "                                                               AND RB.OrderNo = RD.OrderNo " & vbCrLf & _
            '                  "  WHERE RB.SuratJalanNo <> '' " & vbCrLf & _
            '                  "  AND RD.AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
            '                  "  AND RD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "' " & vbCrLf & _
            '                  "  AND RD.OrderNO = '" & Session("OrderNo") & "' " & vbCrLf & _
            '                  "  UNION ALL " & vbCrLf & _
            '                   " SELECT DISTINCT " & vbCrLf & _
            '      "                         idx = 0 , " & vbCrLf & _
            '      "                         colpilih = 0 , " & vbCrLf & _
            '      "                         colorderno = RD.OrderNo , " & vbCrLf & _
            '      "                         collabelno = '' , " & vbCrLf & _
            '      "                         colpartno = RD.partno , " & vbCrLf & _
            '      "                         colpartname = MP.PartName , " & vbCrLf & _
            '      "                         coluom = ISNULL(MU.DESCRIPTION, '') , " & vbCrLf & _
            '      "                         colqtybox = ISNULL(MPM.QtyBox, 0) , " & vbCrLf & _
            '      "                         coldelqty = ISNULL(DSD.DOQty, 0) , "

            'ls_sql = ls_sql + "                         colgoodreceiving = 0 , " & vbCrLf & _
            '                  "                         coldefectreceiving = ISNULL(RB.Box, 0)*ISNULL(MPM.QtyBox, 0) , " & vbCrLf & _
            '                  "                         coldefect = ISNULL(RB.Box, 0) , " & vbCrLf & _
            '                  "                         colreceivingbox = 0 , " & vbCrLf & _
            '                  "                         colHgood = 0, " & vbCrLf & _
            '                  "                         colHdefect = ISNULL(RB.Box, 0)*ISNULL(MPM.QtyBox, 0) , " & vbCrLf & _
            '                  "                         colpono = ISNULL(POD.PONo, '') , " & vbCrLf & _
            '                  "                         LabelNo1 = ISNULL(RTRIM(RB.Label1), '') , " & vbCrLf & _
            '                  "                         LabelNo2 = ISNULL(RTRIM(RB.Label2), '') , " & vbCrLf & _
            '                  "                         PART = RD.PartNo , " & vbCrLf & _
            '                  "                         PO = POM.PONo , "

            'ls_sql = ls_sql + "                         LABEL = ISNULL(RTRIM(RB.Label1), '') + '-' " & vbCrLf & _
            '                  "                         + ISNULL(RTRIM(RB.Label2), '') , " & vbCrLf & _
            '                  "                         StatusDefect = RB.StatusDefect " & vbCrLf & _
            '                  "               FROM      DOSupplier_Detail_Export DSD " & vbCrLf & _
            '                  "                         LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo " & vbCrLf & _
            '                  "                                                               AND DSM.AffiliateID = DSD.AffiliateID " & vbCrLf & _
            '                  "                                                               AND DSM.SupplierID = DSD.SupplierID " & vbCrLf & _
            '                  "                                                               AND DSM.PONO = DSD.PONO " & vbCrLf & _
            '                  "                         LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO " & vbCrLf & _
            '                  "                                                           AND POD.AffiliateID = DSM.AffiliateID " & vbCrLf & _
            '                  "                                                           AND POD.SupplierID = DSM.SupplierID "

            'ls_sql = ls_sql + "                                                           AND POD.PartNo = DSD.PartNo " & vbCrLf & _
            '                  "                         LEFT JOIN ( SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo1 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor1 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort1 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory1 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                     UNION ALL " & vbCrLf & _
            '                  "                                     SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo2 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor2 , "

            'ls_sql = ls_sql + "                                             ETAPort = ETAPort2 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory2 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                     UNION ALL " & vbCrLf & _
            '                  "                                     SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo3 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor3 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort3 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory3 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                     UNION ALL "

            'ls_sql = ls_sql + "                                     SELECT  * , " & vbCrLf & _
            '                  "                                             OrderNO = OrderNo4 , " & vbCrLf & _
            '                  "                                             ETDVendor = ETDVendor4 , " & vbCrLf & _
            '                  "                                             ETAPort = ETAPort4 , " & vbCrLf & _
            '                  "                                             ETAFactory = ETAFactory4 " & vbCrLf & _
            '                  "                                     FROM    Po_Master_Export " & vbCrLf & _
            '                  "                                   ) POM ON POM.PONO = POD.PONO " & vbCrLf & _
            '                  "                                            AND POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
            '                  "                                            AND POM.SupplierID = POD.SupplierID " & vbCrLf & _                              
            '                  "                         LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo " & vbCrLf & _
            '                  "                                                               AND DSD.affiliateID = RM.affiliateID " & vbCrLf & _
            '                  "                                                               AND DSD.SupplierID = RM.SupplierID " & vbCrLf & _
            '                  "                         LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo "

            'ls_sql = ls_sql + "                                                               AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  "                                                               AND RM.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "                                                               AND RM.PONO = RD.PONO " & vbCrLf & _
            '                  "                                                               AND DSD.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                                                               AND DSD.PONO = RD.PONO " & vbCrLf & _
            '                  "                         LEFT JOIN ( SELECT  suratjalanno , " & vbCrLf & _
            '                  "                                             supplierid , " & vbCrLf & _
            '                  "                                             affiliateID , " & vbCrLf & _
            '                  "                                             PONO , " & vbCrLf & _
            '                  "                                             partno , " & vbCrLf & _
            '                  "                                             goodRecQty = SUM(ISNULL(goodRecQty, "

            'ls_sql = ls_sql + "                                                               0)) , " & vbCrLf & _
            '                  "                                             DefectRecQty = SUM(ISNULL(DefectRecQty, " & vbCrLf & _
            '                  "                                                               0)) " & vbCrLf & _
            '                  "                                     FROM    ReceiveForwarder_Detail " & vbCrLf & _
            '                  "                                     GROUP BY suratjalanno , " & vbCrLf & _
            '                  "                                             supplierid , " & vbCrLf & _
            '                  "                                             affiliateID , " & vbCrLf & _
            '                  "                                             PONO , " & vbCrLf & _
            '                  "                                             partno " & vbCrLf & _
            '                  "                                   ) REM ON REM.SuratJalanNo = RD.SuratjalanNo " & vbCrLf & _
            '                  "                                            AND REM.AffiliateID = RD.AffiliateID "

            'ls_sql = ls_sql + "                                            AND REM.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "                                            AND REM.PONO = RD.PONO " & vbCrLf & _
            '                  "                                            AND REM.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                                            AND REM.PONO = RD.PONO " & vbCrLf & _
            '                  "                         LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID " & vbCrLf & _
            '                  "                         LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID " & vbCrLf & _
            '                  "                         LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID " & vbCrLf & _
            '                  "                         LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo " & vbCrLf & _
            '                  "                         LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID " & vbCrLf & _
            '                  "                                                         AND MPM.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                         LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls "

            'ls_sql = ls_sql + "                         LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                  "                                                               AND RB.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  "                                                               AND RB.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "                                                               AND RB.PartNo = RD.PartNo " & vbCrLf & _
            '                  "                                                               AND RB.StatusDefect = '1' " & vbCrLf & _
            '                  "                                                               AND RB.PONo = RD.PONo " & vbCrLf & _
            '                  "                                                               AND RB.OrderNo = RD.OrderNo " & vbCrLf & _
            '                  "  WHERE RB.SuratJalanNo <> '' "

            'ls_sql = ls_sql + " AND RD.AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
            '                  " AND RD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "' " & vbCrLf & _
            '                  " AND RD.OrderNO = '" & Session("OrderNo") & "' " & vbCrLf & _
            '                  "  )Y Order By colOrderNO,colpartno,LabelNo1, LabelNo2 "

            ls_sql = " SELECT  colno = CONVERT(CHAR, ROW_NUMBER() OVER ( ORDER BY colOrderNO, colpartno, LabelNo1, LabelNo2 )) ,	 " & vbCrLf & _
                  " 		* " & vbCrLf & _
                  " 		FROM    (			 " & vbCrLf & _
                  " 				SELECT DISTINCT  " & vbCrLf & _
                  "                          idx = 0 ,  " & vbCrLf & _
                  "                          colpilih = 0 ,  " & vbCrLf & _
                  "                          colorderno = RD.OrderNo ,  " & vbCrLf & _
                  "                          collabelno = '' ,  " & vbCrLf & _
                  "                          colpartno = RD.partno ,  " & vbCrLf & _
                  "                          colpartname = MP.PartName ,  " & vbCrLf & _
                  "                          coluom = ISNULL(MU.DESCRIPTION, '') ,  " & vbCrLf

            ls_sql = ls_sql + "                          colqtybox = ISNULL(POD.POQtyBox,MPM.QtyBox) , " & vbCrLf & _
                              " 						 coldelqty = ISNULL(DSD.DOQty, 0) ,  " & vbCrLf & _
                              "                          colgoodreceiving = ISNULL(POD.POQtyBox,MPM.QtyBox)*ISNULL(RB.box, 0) ,  " & vbCrLf & _
                              "                          coldefectreceiving = 0 ,  " & vbCrLf & _
                              "                          coldefect = 0 ,  " & vbCrLf & _
                              "                          colreceivingbox = ISNULL(RB.box, 0) ,  " & vbCrLf & _
                              "                          colHgood = ISNULL(POD.POQtyBox,MPM.QtyBox)*ISNULL(RB.box, 0) ,  " & vbCrLf & _
                              "                          colHdefect = 0 ,  " & vbCrLf & _
                              "                          colpono = ISNULL(POD.PONo, '') ,  " & vbCrLf & _
                              "                          LabelNo1 = ISNULL(RTRIM(RB.Label1), '') ,  " & vbCrLf & _
                              "                          LabelNo2 = ISNULL(RTRIM(RB.Label2), '') ,  " & vbCrLf

            ls_sql = ls_sql + "                          PART = RD.PartNo , " & vbCrLf & _
                              " 						 PO = POM.PONo ,  " & vbCrLf & _
                              "                          LABEL = ISNULL(RTRIM(RB.Label1), '') + '-'  " & vbCrLf & _
                              "                          + ISNULL(RTRIM(RB.Label2), '') ,  " & vbCrLf & _
                              "                          StatusDefect = RB.StatusDefect  " & vbCrLf & _
                              "                FROM      ReceiveForwarder_Detail RD  " & vbCrLf & _
                              "                          LEFT JOIN ReceiveForwarder_Master RM ON RM.SuratJalanNo = RD.SuratjalanNo  " & vbCrLf & _
                              "                                                                AND RM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "                                                                AND RM.SupplierID = RD.SupplierID  " & vbCrLf & _
                              "                                                                AND RM.PONO = RD.PONO  " & vbCrLf & _
                              "                          LEFT JOIN po_detail_Export POD ON POD.PONO = RM.PONO  " & vbCrLf

            ls_sql = ls_sql + "                                                            AND POD.AffiliateID = RM.AffiliateID " & vbCrLf & _
                              " 														   AND POD.SupplierID = RM.SupplierID  " & vbCrLf & _
                              "                                                            AND POD.PartNo = RD.PartNo  " & vbCrLf & _
                              "                          LEFT JOIN ( SELECT  * ,  " & vbCrLf & _
                              "                                              OrderNO = OrderNo1 ,  " & vbCrLf & _
                              "                                              ETDVendor = ETDVendor1 ,  " & vbCrLf & _
                              "                                              ETAPort = ETAPort1 ,  " & vbCrLf & _
                              "                                              ETAFactory = ETAFactory1  " & vbCrLf & _
                              "                                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                                      UNION ALL  " & vbCrLf & _
                              "                                      SELECT  * ,  " & vbCrLf

            ls_sql = ls_sql + "                                              OrderNO = OrderNo2 , " & vbCrLf & _
                              " 											 ETDVendor = ETDVendor2 ,  " & vbCrLf & _
                              "                                              ETAPort = ETAPort2 ,  " & vbCrLf & _
                              "                                              ETAFactory = ETAFactory2  " & vbCrLf & _
                              "                                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                                      UNION ALL  " & vbCrLf & _
                              "                                      SELECT  * ,  " & vbCrLf & _
                              "                                              OrderNO = OrderNo3 ,  " & vbCrLf & _
                              "                                              ETDVendor = ETDVendor3 ,  " & vbCrLf & _
                              "                                              ETAPort = ETAPort3 ,  " & vbCrLf & _
                              "                                              ETAFactory = ETAFactory3  " & vbCrLf

            ls_sql = ls_sql + "                                      FROM    Po_Master_Export " & vbCrLf & _
                              " 									 UNION ALL  " & vbCrLf & _
                              "                                      SELECT  * ,  " & vbCrLf & _
                              "                                              OrderNO = OrderNo4 ,  " & vbCrLf & _
                              "                                              ETDVendor = ETDVendor4 ,  " & vbCrLf & _
                              "                                              ETAPort = ETAPort4 ,  " & vbCrLf & _
                              "                                              ETAFactory = ETAFactory4  " & vbCrLf & _
                              "                                      FROM    Po_Master_Export  " & vbCrLf & _
                              "                                    ) POM ON POM.PONO = POD.PONO  " & vbCrLf & _
                              "                                             AND POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                              "                                             AND POM.SupplierID = POD.SupplierID  " & vbCrLf

            ls_sql = ls_sql + "                          LEFT JOIN DOSupplier_Master_Export DSM ON RD.suratJalanNo = DSM.SuratJalanNo  " & vbCrLf & _
                              "                                                                AND RD.affiliateID = DSM.affiliateID  " & vbCrLf & _
                              "                                                                AND RD.SupplierID = DSM.SupplierID " & vbCrLf & _
                              " 						 LEFT JOIN DOSupplier_Detail_Export DSD ON DSM.SuratJalanNo = DSD.SuratjalanNo  " & vbCrLf & _
                              "                                                                AND DSM.AffiliateID = DSD.AffiliateID  " & vbCrLf & _
                              "                                                                AND DSM.SupplierID = DSD.SupplierID  " & vbCrLf & _
                              "                                                                AND DSM.PONO = DSD.PONO  " & vbCrLf & _
                              "                                                                AND RD.PartNo = DSD.PartNo  " & vbCrLf & _
                              "                                                                AND RD.PONO = DSD.PONO  " & vbCrLf & _
                              "                          LEFT JOIN ( SELECT  suratjalanno ,  " & vbCrLf & _
                              "                                              supplierid ,  " & vbCrLf

            ls_sql = ls_sql + "                                              affiliateID ,  " & vbCrLf & _
                              "                                              PONO ,  " & vbCrLf & _
                              "                                              partno , " & vbCrLf & _
                              " 											 goodRecQty = SUM(ISNULL(goodRecQty,  " & vbCrLf & _
                              "                                                                0)) ,  " & vbCrLf & _
                              "                                              DefectRecQty = SUM(ISNULL(DefectRecQty,  " & vbCrLf & _
                              "                                                                0))  " & vbCrLf & _
                              "                                      FROM    ReceiveForwarder_Detail  " & vbCrLf & _
                              "                                      GROUP BY suratjalanno ,  " & vbCrLf & _
                              "                                              supplierid ,  " & vbCrLf & _
                              "                                              affiliateID , " & vbCrLf

            ls_sql = ls_sql + "                                              PONO ,  " & vbCrLf & _
                              "                                              partno  " & vbCrLf & _
                              "                                    ) REM ON REM.SuratJalanNo = RD.SuratjalanNo " & vbCrLf & _
                              " 											AND REM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "                                             AND REM.SupplierID = RD.SupplierID  " & vbCrLf & _
                              "                                             AND REM.PONO = RD.PONO  " & vbCrLf & _
                              "                                             AND REM.PartNo = RD.PartNo  " & vbCrLf & _
                              "                                             AND REM.PONO = RD.PONO  " & vbCrLf & _
                              "                          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                              "                          LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID  " & vbCrLf & _
                              "                          LEFT JOIN ms_supplier MS ON MS.SupplierID = RM.SupplierID  " & vbCrLf

            ls_sql = ls_sql + "                          LEFT JOIN MS_Parts MP ON MP.PartNo = RD.PartNo  " & vbCrLf & _
                              "                          LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                              "                                                          AND MPM.PartNo = RD.PartNo                           " & vbCrLf & _
                              " 						 LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls  " & vbCrLf & _
                              "                          LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo  " & vbCrLf & _
                              "                                                                AND RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "                                                                AND RB.SupplierID = RD.SupplierID  " & vbCrLf & _
                              "                                                                AND RB.PartNo = RD.PartNo  " & vbCrLf & _
                              "                                                                AND RB.StatusDefect IN ('0','1')  " & vbCrLf & _
                              "                                                                AND RB.PONo = RD.PONo  " & vbCrLf & _
                              "                                                                AND RB.OrderNo = RD.OrderNo  " & vbCrLf

            ls_sql = ls_sql + "   WHERE RB.SuratJalanNo <> ''  " & vbCrLf & _
                              "   AND RD.AffiliateID = '" & Session("AFFID") & "'  " & vbCrLf & _
                              "   AND RD.SuratJalanNo = '" & Session("SuratJalanNoSupplier") & "'  " & vbCrLf & _
                              "   AND RD.OrderNO = '" & Session("OrderNo") & "'  " & vbCrLf & _
                              "   )Y Order By colOrderNO,colpartno,LabelNo1, LabelNo2, StatusDefect "


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
                'Call colorGrid()
            End If
        End Using
    End Sub

    Private Sub Up_AddCarton()
        Dim ls_sql As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_sql = "     Select colno = CONVERT(char,ROW_NUMBER() OVER(ORDER BY colOrderNO,colpartno,LabelNo1, LabelNo2)), * FROM (  " & vbCrLf & _
                              "     SELECT distinct idx = 0,colpilih = 0,  " & vbCrLf & _
                              "            colorderno = RD.OrderNo ,    " & vbCrLf & _
                              "            collabelno = '' ,    " & vbCrLf & _
                              "            colpartno = RD.partno ,    " & vbCrLf & _
                              "            colpartname = MP.PartName ,    " & vbCrLf & _
                              "            coluom = ISNULL(MU.DESCRIPTION, '') ,    " & vbCrLf & _
                              "            colqtybox = ISNULL(DSD.POQtyBox,MPM.QtyBox) ,    " & vbCrLf & _
                              "            coldelqty = ISNULL(DSD.DOQty, 0) ,    " & vbCrLf & _
                              "            colgoodreceiving = ISNULL(DSD.POQtyBox,MPM.QtyBox)*ISNULL(RB.Box, 0),    " & vbCrLf & _
                              "            coldefectreceiving = 0 , coldefect = 0 ,              "

            ls_sql = ls_sql + "            colreceivingbox = ISNULL(RB.Box, 0) ,    " & vbCrLf & _
                              "            colHgood = RD.GoodRecQty,  " & vbCrLf & _
                              "            colHdefect = ISNULL(RD.DefectrecQty, 0),    " & vbCrLf & _
                              "            colpono = isnull(POD.PONo,''), LabelNo1 = isnull(Rtrim(RB.Label1),''), LabelNo2= isnull(Rtrim(RB.Label2),''),   " & vbCrLf & _
                              "            PART = RD.PartNo, PO = POM.PONo,LABEL =  isnull(Rtrim(RB.Label1),'') + '-' + isnull(Rtrim(RB.Label2),''), StatusDefect = RB.StatusDefect   " & vbCrLf & _
                              "     FROM   DOSupplier_Detail_Export DSD    " & vbCrLf & _
                              "            LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf & _
                              "                                                      AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf & _
                              "                                                      AND DSM.SupplierID = DSD.SupplierID    " & vbCrLf & _
                              "                                                      AND DSM.PONO = DSD.PONO              LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO    " & vbCrLf & _
                              "                                              AND POD.AffiliateID = DSM.AffiliateID                                              AND POD.SupplierID = DSM.SupplierID    "

            ls_sql = ls_sql + "                                              AND POD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo1 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,                                  OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2    "

            ls_sql = ls_sql + "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    Po_Master_Export                        UNION ALL                          SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort4 ,    "

            ls_sql = ls_sql + "                                ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                      ) POM ON POM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT TOP 1                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo1 ,                                  ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    "

            ls_sql = ls_sql + "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2                          FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    "

            ls_sql = ls_sql + "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL                          SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort4 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export                        ORDER BY PORevNo    "

            ls_sql = ls_sql + "                      ) PRM ON PRM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND PRM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND PRM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO                                                   AND PRD.AffiliateID = PRM.AffiliateID    " & vbCrLf & _
                              "                                                 AND PRD.SupplierID = PRM.SupplierID    " & vbCrLf & _
                              "                                                 AND PRD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo    " & vbCrLf & _
                              "                                                    AND DSD.affiliateID = RM.affiliateID    " & vbCrLf & _
                              "                                                    AND DSD.SupplierID = RM.SupplierID            LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo    " & vbCrLf & _
                              "                                                    AND RM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                                                    AND RM.SupplierID = RD.SupplierID    "

            ls_sql = ls_sql + "                                                    AND RM.PONO = RD.PONO    " & vbCrLf & _
                              "                                                    AND DSD.PartNo = RD.PartNo    " & vbCrLf & _
                              "                                                    AND DSD.PONO = RD.PONO              LEFT JOIN ( SELECT  suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,    " & vbCrLf & _
                              "                                partno ,                                goodRecQty = SUM(ISNULL(goodRecQty, 0)) ,    " & vbCrLf & _
                              "                                DefectRecQty = SUM(ISNULL(DefectRecQty, 0))    " & vbCrLf & _
                              "                        FROM    ReceiveForwarder_Detail    " & vbCrLf & _
                              "                        GROUP BY suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    "

            ls_sql = ls_sql + "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,                                  partno    " & vbCrLf & _
                              "                      ) REM ON REM.SuratJalanNo = RD.SuratjalanNo    " & vbCrLf & _
                              "                               AND REM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                               AND REM.SupplierID = RD.SupplierID                               AND REM.PONO = RD.PONO    " & vbCrLf & _
                              "                               AND REM.PartNo = RD.PartNo    " & vbCrLf & _
                              "                               AND REM.PONO = RD.PONO    " & vbCrLf & _
                              "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID    " & vbCrLf & _
                              "            LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID    " & vbCrLf & _
                              "            LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo    "

            ls_sql = ls_sql + "            LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID and MPM.PartNo = RD.PartNo             LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo and RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "  					AND RB.SupplierID = RD.SupplierID and RB.PartNo = RD.PartNo and RB.StatusDefect = '0'  " & vbCrLf & _
                              "                      AND RB.PONo = RD.PONo and RB.OrderNo = RD.OrderNo  " & vbCrLf & _
                              "   WHERE RB.SuratJalanNo <> ''  " & vbCrLf & _
                              "   AND RD.AffiliateID = '" & Session("AFFID") & "'  " & vbCrLf & _
                              "   AND RD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "'  " & vbCrLf & _
                              "   AND RD.OrderNO = '" & Session("OrderNo") & "'  " & vbCrLf & _
                              "   UNION ALL  " & vbCrLf & _
                              "   SELECT distinct idx = 0,colpilih = 0,    " & vbCrLf & _
                              "            colorderno = RD.OrderNo ,    "

            ls_sql = ls_sql + "            collabelno = '' ,              colpartno = RD.partno ,    " & vbCrLf & _
                              "            colpartname = MP.PartName ,    " & vbCrLf & _
                              "            coluom = ISNULL(MU.DESCRIPTION, '') ,    " & vbCrLf & _
                              "            colqtybox = ISNULL(DSD.POQtyBox,MPM.QtyBox) ,    " & vbCrLf & _
                              "            coldelqty = ISNULL(DSD.DOQty, 0) ,    " & vbCrLf & _
                              "            colgoodreceiving = 0 ,    " & vbCrLf & _
                              "            coldefectreceiving = ISNULL(DSD.POQtyBox,MPM.QtyBox)*ISNULL(RB.Box, 0) ,            " & vbCrLf & _
                              "            coldefect = ISNULL(RB.Box, 0) , " & vbCrLf & _
                              "            colreceivingbox = 0 ,    " & vbCrLf & _
                              "            colHgood = RD.GoodRecQty,    " & vbCrLf & _
                              "            colHdefect = ISNULL(RD.DefectrecQty, 0),              colpono = isnull(POD.PONo,''), LabelNo1 = isnull(Rtrim(RB.Label1),''), LabelNo2= isnull(Rtrim(RB.Label2),''),   "

            ls_sql = ls_sql + "            PART = RD.PartNo, PO = POM.PONo,LABEL =  isnull(Rtrim(RB.Label1),'') + '-' + isnull(Rtrim(RB.Label2),''), StatusDefect = RB.StatusDefect   " & vbCrLf & _
                              "     FROM   DOSupplier_Detail_Export DSD    " & vbCrLf & _
                              "            LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf & _
                              "                                                      AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf & _
                              "                                                      AND DSM.SupplierID = DSD.SupplierID    " & vbCrLf & _
                              "                                                      AND DSM.PONO = DSD.PONO    " & vbCrLf & _
                              "            LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO    " & vbCrLf & _
                              "                                              AND POD.AffiliateID = DSM.AffiliateID                                              AND POD.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "                                              AND POD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT  * ,                                  OrderNO = OrderNo1 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor1 ,    "

            ls_sql = ls_sql + "                                ETAPort = ETAPort1 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2    " & vbCrLf & _
                              "                        FROM    Po_Master_Export                          UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    "

            ls_sql = ls_sql + "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    Po_Master_Export                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort4 ,                                  ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                      ) POM ON POM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND POM.AffiliateID = POD.AffiliateID    "

            ls_sql = ls_sql + "                               AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT TOP 1                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo1 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export                          ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    "

            ls_sql = ls_sql + "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1                                  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    "

            ls_sql = ls_sql + "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,                                  ETAPort = ETAPort4 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export                        ORDER BY PORevNo    " & vbCrLf & _
                              "                      ) PRM ON PRM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND PRM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND PRM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO    "

            ls_sql = ls_sql + "                                                 AND PRD.AffiliateID = PRM.AffiliateID    " & vbCrLf & _
                              "                                                 AND PRD.SupplierID = PRM.SupplierID    " & vbCrLf & _
                              "                                                 AND PRD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo                                                      AND DSD.affiliateID = RM.affiliateID    " & vbCrLf & _
                              "                                                    AND DSD.SupplierID = RM.SupplierID            LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo    " & vbCrLf & _
                              "                                                    AND RM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                                                    AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                    AND RM.PONO = RD.PONO    " & vbCrLf & _
                              "                                                    AND DSD.PartNo = RD.PartNo    " & vbCrLf & _
                              "                                                    AND DSD.PONO = RD.PONO    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT  suratjalanno ,    "

            ls_sql = ls_sql + "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,                                  partno ,                                goodRecQty = SUM(ISNULL(goodRecQty, 0)) ,    " & vbCrLf & _
                              "                                DefectRecQty = SUM(ISNULL(DefectRecQty, 0))    " & vbCrLf & _
                              "                        FROM    ReceiveForwarder_Detail    " & vbCrLf & _
                              "                        GROUP BY suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,    " & vbCrLf & _
                              "                                partno    " & vbCrLf & _
                              "                      ) REM ON REM.SuratJalanNo = RD.SuratjalanNo    "

            ls_sql = ls_sql + "                               AND REM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                               AND REM.SupplierID = RD.SupplierID AND REM.PONO = RD.PONO                                 AND REM.PartNo = RD.PartNo    " & vbCrLf & _
                              "                               AND REM.PONO = RD.PONO    " & vbCrLf & _
                              "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID    " & vbCrLf & _
                              "            LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID    " & vbCrLf & _
                              "            LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID and MPM.PartNo = RD.PartNo   " & vbCrLf & _
                              "            LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo and RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "  					AND RB.SupplierID = RD.SupplierID and RB.PartNo = RD.PartNo and RB.StatusDefect = '1'  "

            ls_sql = ls_sql + "                      AND RB.PONo = RD.PONo and RB.OrderNo = RD.OrderNo  " & vbCrLf & _
                              "   WHERE RB.SuratJalanNo <> ''  AND RD.AffiliateID = '" & Session("AFFID") & "'  " & vbCrLf & _
                              "  AND RD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "'  " & vbCrLf & _
                              "  AND RD.OrderNO = '" & Session("OrderNo") & "' " & vbCrLf & _
                              "   )Y  " & vbCrLf & _
                              "   --------------------- DETAIL ----------------------------------------- " & vbCrLf & _
                              "   UNION ALL " & vbCrLf & _
                              "   SELECT distinct colno = '',idx = 1,colpilih = 0,  " & vbCrLf & _
                              "            colorderno = '' ,    " & vbCrLf & _
                              "            collabelno = '' ,    " & vbCrLf & _
                              "            colpartno = '' ,    "

            ls_sql = ls_sql + "            colpartname = '' ,    " & vbCrLf & _
                              "            coluom = ISNULL(MU.DESCRIPTION, '') ,    " & vbCrLf & _
                              "            colqtybox = ISNULL(DSD.POQtyBox,MPM.QtyBox) ,    " & vbCrLf & _
                              "            coldelqty = ISNULL(DSD.DOQty, 0) ,    " & vbCrLf & _
                              "            colgoodreceiving = ISNULL(DSD.POQtyBox,MPM.QtyBox)*ISNULL(RB.Box, 0),    " & vbCrLf & _
                              "            coldefectreceiving = 0 , coldefect = 0 ,              " & vbCrLf & _
                              "            colreceivingbox = ISNULL(RB.Box, 0) ,    " & vbCrLf & _
                              "            colHgood = RD.GoodRecQty,  " & vbCrLf & _
                              "            colHdefect = ISNULL(RD.DefectrecQty, 0),    " & vbCrLf & _
                              "            colpono = isnull(POD.PONo,''), LabelNo1 = isnull(Rtrim(RB.Label1),''), LabelNo2= isnull(Rtrim(RB.Label2),''),   " & vbCrLf & _
                              "            PART = RD.PartNo, PO = POM.PONo,LABEL =  isnull(Rtrim(RB.Label1),'') + '-' + isnull(Rtrim(RB.Label2),''), StatusDefect = RB.StatusDefect   "

            ls_sql = ls_sql + "     FROM   DOSupplier_Detail_Export DSD    " & vbCrLf & _
                              "            LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf & _
                              "                                                      AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf & _
                              "                                                      AND DSM.SupplierID = DSD.SupplierID    " & vbCrLf & _
                              "                                                      AND DSM.PONO = DSD.PONO              LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO    " & vbCrLf & _
                              "                                              AND POD.AffiliateID = DSM.AffiliateID                                              AND POD.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "                                              AND POD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo1 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    "

            ls_sql = ls_sql + "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,                                  OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor3 ,    "

            ls_sql = ls_sql + "                                ETAPort = ETAPort3 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    Po_Master_Export                        UNION ALL                          SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort4 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                      ) POM ON POM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND POM.SupplierID = POD.SupplierID    "

            ls_sql = ls_sql + "            LEFT JOIN ( SELECT TOP 1                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo1 ,                                  ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    "

            ls_sql = ls_sql + "                                ETAFactory = ETAFactory2                          FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL                          SELECT TOP 1    "

            ls_sql = ls_sql + "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort4 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export                        ORDER BY PORevNo    " & vbCrLf & _
                              "                      ) PRM ON PRM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND PRM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND PRM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO                                                   AND PRD.AffiliateID = PRM.AffiliateID    " & vbCrLf & _
                              "                                                 AND PRD.SupplierID = PRM.SupplierID    "

            ls_sql = ls_sql + "                                                 AND PRD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo    " & vbCrLf & _
                              "                                                    AND DSD.affiliateID = RM.affiliateID    " & vbCrLf & _
                              "                                                    AND DSD.SupplierID = RM.SupplierID            LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo    " & vbCrLf & _
                              "                                                    AND RM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                                                    AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                    AND RM.PONO = RD.PONO    " & vbCrLf & _
                              "                                                    AND DSD.PartNo = RD.PartNo    " & vbCrLf & _
                              "                                                    AND DSD.PONO = RD.PONO              LEFT JOIN ( SELECT  suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    "

            ls_sql = ls_sql + "                                PONO ,    " & vbCrLf & _
                              "                                partno ,                                goodRecQty = SUM(ISNULL(goodRecQty, 0)) ,    " & vbCrLf & _
                              "                                DefectRecQty = SUM(ISNULL(DefectRecQty, 0))    " & vbCrLf & _
                              "                        FROM    ReceiveForwarder_Detail    " & vbCrLf & _
                              "                        GROUP BY suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,                                  partno    " & vbCrLf & _
                              "                      ) REM ON REM.SuratJalanNo = RD.SuratjalanNo    " & vbCrLf & _
                              "                               AND REM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                               AND REM.SupplierID = RD.SupplierID                               AND REM.PONO = RD.PONO    "

            ls_sql = ls_sql + "                               AND REM.PartNo = RD.PartNo    " & vbCrLf & _
                              "                               AND REM.PONO = RD.PONO    " & vbCrLf & _
                              "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID    " & vbCrLf & _
                              "            LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID    " & vbCrLf & _
                              "            LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID and MPM.PartNo = RD.PartNo             LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo and RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "  					AND RB.SupplierID = RD.SupplierID and RB.PartNo = RD.PartNo and RB.StatusDefect = '0'  " & vbCrLf & _
                              "                      AND RB.PONo = RD.PONo and RB.OrderNo = RD.OrderNo  " & vbCrLf & _
                              "   WHERE RB.SuratJalanNo <> ''  "

            ls_sql = ls_sql + "   AND RD.AffiliateID = '" & Session("AFFID") & "'  " & vbCrLf & _
                              "   AND RD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "'  " & vbCrLf & _
                              "   AND RD.OrderNO = '" & Session("OrderNo") & "'  " & vbCrLf & _
                              "   AND Rtrim(RM.PONo)+Rtrim(RD.PartNo)+(isnull(Rtrim(RB.Label1),'') + '-' + isnull(Rtrim(RB.Label2),'')) IN( " & Trim(Session("combination")) & ")" & vbCrLf & _
                              "   UNION ALL  " & vbCrLf & _
                              "   SELECT distinct colno = '', idx = 1,colpilih = 0,    " & vbCrLf & _
                              "            colorderno = '' ,    " & vbCrLf & _
                              "            collabelno = '' ,              colpartno ='' ,    " & vbCrLf & _
                              "            colpartname ='' ,    " & vbCrLf & _
                              "            coluom = ISNULL(MU.DESCRIPTION, '') ,    " & vbCrLf & _
                              "            colqtybox = ISNULL(DSD.POQtyBox,MPM.QtyBox) ,    " & vbCrLf & _
                              "            coldelqty = ISNULL(DSD.DOQty, 0) ,    "

            ls_sql = ls_sql + "            colgoodreceiving = 0 ,    " & vbCrLf & _
                              "            coldefectreceiving = ISNULL(DSD.POQtyBox,MPM.QtyBox)*ISNULL(RB.Box, 0) ,            " & vbCrLf & _
                              "            coldefect = ISNULL(RB.Box, 0) , " & vbCrLf & _
                              "            colreceivingbox = 0,    " & vbCrLf & _
                              "            colHgood = RD.GoodRecQty,    " & vbCrLf & _
                              "            colHdefect = ISNULL(RD.DefectrecQty, 0),              colpono = isnull(POD.PONo,''), LabelNo1 = isnull(Rtrim(RB.Label1),''), LabelNo2= isnull(Rtrim(RB.Label2),''),   " & vbCrLf & _
                              "            PART = RD.PartNo, PO = POM.PONo,LABEL =  isnull(Rtrim(RB.Label1),'') + '-' + isnull(Rtrim(RB.Label2),''), StatusDefect = RB.StatusDefect   " & vbCrLf & _
                              "     FROM   DOSupplier_Detail_Export DSD    " & vbCrLf & _
                              "            LEFT JOIN DOSupplier_Master_Export DSM ON DSM.SuratJalanNo = DSD.SuratjalanNo    " & vbCrLf & _
                              "                                                      AND DSM.AffiliateID = DSD.AffiliateID    " & vbCrLf & _
                              "                                                      AND DSM.SupplierID = DSD.SupplierID    "

            ls_sql = ls_sql + "                                                      AND DSM.PONO = DSD.PONO    " & vbCrLf & _
                              "            LEFT JOIN po_detail_Export POD ON POD.PONO = DSM.PONO    " & vbCrLf & _
                              "                                              AND POD.AffiliateID = DSM.AffiliateID                                              AND POD.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "                                              AND POD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT  * ,                                  OrderNO = OrderNo1 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    "

            ls_sql = ls_sql + "                                OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2    " & vbCrLf & _
                              "                        FROM    Po_Master_Export                          UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    Po_Master_Export                        UNION ALL    " & vbCrLf & _
                              "                        SELECT  * ,    "

            ls_sql = ls_sql + "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort4 ,                                  ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    Po_Master_Export    " & vbCrLf & _
                              "                      ) POM ON POM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND POM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND POM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT TOP 1                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo1 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor1 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort1 ,    "

            ls_sql = ls_sql + "                                ETAFactory = ETAFactory1    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export                          ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo2 ,                                ETDVendor = ETDVendor2 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort2 ,    " & vbCrLf & _
                              "                                ETAFactory = ETAFactory2    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    "

            ls_sql = ls_sql + "                        SELECT TOP 1                                  * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo3 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor3 ,    " & vbCrLf & _
                              "                                ETAPort = ETAPort3 ,                                ETAFactory = ETAFactory3    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export    " & vbCrLf & _
                              "                        ORDER BY PORevNo    " & vbCrLf & _
                              "                        UNION ALL    " & vbCrLf & _
                              "                        SELECT TOP 1    " & vbCrLf & _
                              "                                * ,    " & vbCrLf & _
                              "                                OrderNO = OrderNo4 ,    " & vbCrLf & _
                              "                                ETDVendor = ETDVendor4 ,                                  ETAPort = ETAPort4 ,    "

            ls_sql = ls_sql + "                                ETAFactory = ETAFactory4    " & vbCrLf & _
                              "                        FROM    PoRev_Master_Export                        ORDER BY PORevNo    " & vbCrLf & _
                              "                      ) PRM ON PRM.PONO = POD.PONO    " & vbCrLf & _
                              "                               AND PRM.AffiliateID = POD.AffiliateID    " & vbCrLf & _
                              "                               AND PRM.SupplierID = POD.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN poRev_detail_Export PRD ON PRD.PONO = PRM.PONO    " & vbCrLf & _
                              "                                                 AND PRD.AffiliateID = PRM.AffiliateID    " & vbCrLf & _
                              "                                                 AND PRD.SupplierID = PRM.SupplierID    " & vbCrLf & _
                              "                                                 AND PRD.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_Master RM ON DSD.suratJalanNo = RM.SuratJalanNo                                                      AND DSD.affiliateID = RM.affiliateID    " & vbCrLf & _
                              "                                                    AND DSD.SupplierID = RM.SupplierID            LEFT JOIN ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratjalanNo    "

            ls_sql = ls_sql + "                                                    AND RM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                                                    AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              "                                                    AND RM.PONO = RD.PONO    " & vbCrLf & _
                              "                                                    AND DSD.PartNo = RD.PartNo    " & vbCrLf & _
                              "                                                    AND DSD.PONO = RD.PONO    " & vbCrLf & _
                              "            LEFT JOIN ( SELECT  suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,                                  partno ,                                goodRecQty = SUM(ISNULL(goodRecQty, 0)) ,    " & vbCrLf & _
                              "                                DefectRecQty = SUM(ISNULL(DefectRecQty, 0))    " & vbCrLf & _
                              "                        FROM    ReceiveForwarder_Detail    "

            ls_sql = ls_sql + "                        GROUP BY suratjalanno ,    " & vbCrLf & _
                              "                                supplierid ,    " & vbCrLf & _
                              "                                affiliateID ,    " & vbCrLf & _
                              "                                PONO ,    " & vbCrLf & _
                              "                                partno    " & vbCrLf & _
                              "                      ) REM ON REM.SuratJalanNo = RD.SuratjalanNo    " & vbCrLf & _
                              "                               AND REM.AffiliateID = RD.AffiliateID    " & vbCrLf & _
                              "                               AND REM.SupplierID = RD.SupplierID AND REM.PONO = RD.PONO                                 AND REM.PartNo = RD.PartNo    " & vbCrLf & _
                              "                               AND REM.PONO = RD.PONO    " & vbCrLf & _
                              "            LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = DSM.AffiliateID    " & vbCrLf & _
                              "            LEFT JOIN ms_forwarder MF ON MF.ForwarderID = POM.ForwarderID    "

            ls_sql = ls_sql + "            LEFT JOIN ms_supplier MS ON MS.SupplierID = DSM.SupplierID    " & vbCrLf & _
                              "            LEFT JOIN MS_Parts MP ON MP.PartNo = DSD.PartNo    " & vbCrLf & _
                              "            LEFT JOIN Ms_PartMapping MPM ON MPM.AffiliateID = RM.AffiliateID and MPM.PartNo = RD.PartNo   " & vbCrLf & _
                              "            LEFT JOIN MS_UnitCls MU ON MU.UnitCls = MP.UnitCls   " & vbCrLf & _
                              "            LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo and RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              "  					AND RB.SupplierID = RD.SupplierID and RB.PartNo = RD.PartNo and RB.StatusDefect = '1'  " & vbCrLf & _
                              "                      AND RB.PONo = RD.PONo and RB.OrderNo = RD.OrderNo  " & vbCrLf & _
                              "   WHERE RB.SuratJalanNo <> ''  AND RD.AffiliateID = '" & Session("AFFID") & "'  " & vbCrLf & _
                              "  AND RD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "'  " & vbCrLf & _
                              "  AND RD.OrderNO = '" & Session("OrderNo") & "' " & vbCrLf & _
                              "  AND Rtrim(RM.PONo)+Rtrim(RD.PartNo)+(isnull(Rtrim(RB.Label1),'') + '-' + isnull(Rtrim(RB.Label2),'')) IN( " & Trim(Session("combination")) & ")" & vbCrLf & _
                              "  Order By PART, PO,LABEL, idx  "

            '" WHERE DSD.AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
            '" AND DSD.SuratJalanNo = '" & Session("SuratJalanNOSupplier") & "' " & vbCrLf & _
            '" AND COALESCE(PRM.OrderNO, POM.OrderNO) = '" & Session("OrderNo") & "' " & vbCrLf & _
            '" AND Rtrim(RM.PONo)+Rtrim(RD.PartNo) IN( " & Trim(Session("combination")) & ")" & vbCrLf & _
            '"  Order By PART, PO,LABEL, idx " & vbCrLf

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
                'Call colorGrid()
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
        Grid.VisibleColumns(8).CellStyle.BackColor = Drawing.Color.White
        Grid.VisibleColumns(9).CellStyle.BackColor = Drawing.Color.White
        Grid.VisibleColumns(10).CellStyle.BackColor = Drawing.Color.LightYellow
        Grid.VisibleColumns(11).CellStyle.BackColor = Drawing.Color.LightYellow
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
        Dim ls_Label1 As Integer
        Dim ls_Label2 As Integer
        Dim ls_Label1O As Integer
        Dim ls_Label2O As Integer
        Dim ls_labelNo As String

        If HF.Get("hfTest") = "save" Then 'Save
            Session.Remove("sstatus")
            Session("sstatus") = "TRUE"
            If txttotalbox.Text = "" Then txttotalbox.Text = 0
            'pReceiveDate = txtrecdate.Text

            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
                    Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)
                    Dim ls_Good As Integer
                    Dim ls_Defect As Integer
                    With Grid
                        totalBox = 0
                        For iLoop = 0 To e.UpdateValues.Count - 1

                            'cek QTY tidak boleh melebihi Qty
                            If (CDbl(e.UpdateValues(iLoop).NewValues("colgoodreceiving").ToString()) + CDbl(e.UpdateValues(iLoop).NewValues("coldefectreceiving").ToString())) > CDbl(e.UpdateValues(iLoop).NewValues("coldelqty").ToString()) Then
                                Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                                lblerrmessage.Text = lblerrmessage.Text
                                Session("YA010IsSubmit") = lblerrmessage.Text
                                Session("sstatus") = "FALSE"
                                Exit Sub
                            End If
                            'cek QTY tidak boleh melebihi Qty
                            If (CDbl(e.UpdateValues(iLoop).NewValues("colreceivingbox").ToString()) <> 0 And CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString()) <> 0) Then
                                Call clsMsg.DisplayMessage(lblerrmessage, "7014", clsMessage.MsgType.ErrorMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                                lblerrmessage.Text = lblerrmessage.Text
                                Session("YA010IsSubmit") = lblerrmessage.Text
                                Session("sstatus") = "FALSE"
                                Exit Sub
                            End If
                            'cek BoxNo
                            ls_Label1 = CDbl(Microsoft.VisualBasic.Right(e.UpdateValues(iLoop).NewValues("LabelNo1").ToString(), 7))
                            ls_Label2 = CDbl(Microsoft.VisualBasic.Right(e.UpdateValues(iLoop).NewValues("LabelNo2").ToString(), 7))
                            ls_Label1O = CDbl(Microsoft.VisualBasic.Right(e.UpdateValues(iLoop).OldValues("LabelNo1").ToString(), 7))
                            ls_Label2O = CDbl(Microsoft.VisualBasic.Right(e.UpdateValues(iLoop).OldValues("LabelNo2").ToString(), 7))

                            If (ls_Label2 - ls_Label2) < 0 Then
                                Call clsMsg.DisplayMessage(lblerrmessage, "7016", clsMessage.MsgType.ErrorMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                                lblerrmessage.Text = lblerrmessage.Text
                                Session("YA010IsSubmit") = lblerrmessage.Text
                                Session("sstatus") = "FALSE"
                                Exit Sub
                            End If

                            If (ls_Label2 < ls_Label1) Then
                                Call clsMsg.DisplayMessage(lblerrmessage, "7016", clsMessage.MsgType.ErrorMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                                lblerrmessage.Text = lblerrmessage.Text
                                Session("YA010IsSubmit") = lblerrmessage.Text
                                Session("sstatus") = "FALSE"
                                Exit Sub
                            End If

                            If (ls_Label1 >= ls_Label1O And ls_Label1 <= ls_Label2O) Then
                            Else
                                Call clsMsg.DisplayMessage(lblerrmessage, "7016", clsMessage.MsgType.ErrorMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                                lblerrmessage.Text = lblerrmessage.Text
                                Session("YA010IsSubmit") = lblerrmessage.Text
                                Session("sstatus") = "FALSE"
                                Exit Sub
                            End If

                            If (ls_Label2 >= ls_Label1O And ls_Label2 <= ls_Label2O) Then
                            Else
                                Call clsMsg.DisplayMessage(lblerrmessage, "7016", clsMessage.MsgType.ErrorMessage)
                                Grid.JSProperties("cpMessage") = lblerrmessage.Text
                                lblerrmessage.Text = lblerrmessage.Text
                                Session("YA010IsSubmit") = lblerrmessage.Text
                                Session("sstatus") = "FALSE"
                                Exit Sub
                            End If

                            'If lblStatus.Text = "" Then
                            sqlstring = "SELECT * FROM dbo.ReceiveForwarder_Detail WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                            " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                            " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                            " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                            " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf

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
                                'INSERT KANBAN
                                ls_SQL = " INSERT INTO dbo.ReceiveForwarder_Detail " & vbCrLf & _
                                          "         ( SuratJalanNo , " & vbCrLf & _
                                          "           SupplierID , " & vbCrLf & _
                                          "           AffiliateID, " & vbCrLf & _
                                          "           PONo , " & vbCrLf & _
                                          "           PartNo , " & vbCrLf & _
                                          "           OrderNo , " & vbCrLf & _
                                          "           GoodRecQty, " & vbCrLf & _
                                          "           DefectRecQty " & vbCrLf & _
                                          "         ) " & vbCrLf & _
                                          " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                                ls_SQL = ls_SQL + "           '" & Trim(txtsupp.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                                                  "           '" & Session("AFFID") & "' , " & vbCrLf & _
                                                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("PO").ToString()) & "' , " & vbCrLf & _
                                                  "           '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' , -- PartNo - char(120) " & vbCrLf & _
                                                  "           '" & Session("ORDERNO") & "' , -- UnitCls - char(3) " & vbCrLf & _
                                                  "           " & CDbl(e.UpdateValues(iLoop).NewValues("colgoodreceiving").ToString()) & ",  -- RecQty - numeric " & vbCrLf & _
                                                  "           " & CDbl(e.UpdateValues(iLoop).NewValues("coldefectreceiving").ToString()) & "  -- RecQty - numeric " & vbCrLf & _
                                                  "         ) "
                                totalBox = totalBox + (CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) / CDbl(e.UpdateValues(iLoop).NewValues("colqtybox").ToString()))
                                sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                sqlComm.ExecuteNonQuery()

                            ElseIf pIsUpdate = True Then
                                ls_Good = CDbl(e.UpdateValues(iLoop).NewValues("colreceivingbox").ToString()) * CDbl(e.UpdateValues(iLoop).NewValues("colqtybox").ToString())
                                ls_Defect = CDbl(e.UpdateValues(iLoop).NewValues("coldefect").ToString()) * CDbl(e.UpdateValues(iLoop).NewValues("colqtybox").ToString())
                                If e.UpdateValues(iLoop).NewValues("StatusDefect") = 0 Then
                                    'update good ke Defect
                                    ls_SQL = " Update ReceiveForwarder_Detail set " & vbCrLf & _
                                         " GoodRecQty = GoodRecQty - " & CDbl(ls_Defect) & ", " & vbCrLf & _
                                         " DefectRecQty = DefectRecQty + " & CDbl(ls_Defect) & " " & vbCrLf & _
                                         " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                         " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                         " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf
                                Else
                                    'Defect ke Good
                                    ls_SQL = " Update ReceiveForwarder_Detail set " & vbCrLf & _
                                         " GoodRecQty = GoodRecQty + " & CDbl(ls_Good) & ", " & vbCrLf & _
                                         " DefectRecQty = DefectRecQty - " & CDbl(ls_Good) & " " & vbCrLf & _
                                         " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                         " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                         " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf
                                End If
                                sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                sqlComm.ExecuteNonQuery()

                                'Delete Data box
                                ls_SQL = " Delete ReceiveForwarder_Detailbox " & vbCrLf & _
                                         " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                         " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                         " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf & _
                                         " and rtrim(label1)+'-'+rtrim(label2) = '" & Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()) & "'"
                                sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                sqlComm.ExecuteNonQuery()

                                'Update Data
                                ls_SQL = "INSERT INTO ReceiveForwarder_Detailbox Values ( " & vbCrLf & _
                                                  " '" & Trim(txtsuratjalanno.Text) & "', " & vbCrLf & _
                                                  " '" & Trim(txtsupp.Text) & "', " & vbCrLf & _
                                                  " '" & Trim(txtaffiliatecode.Text) & "', " & vbCrLf & _
                                                  " '" & Trim(e.UpdateValues(iLoop).NewValues("PO").ToString()) & "', " & vbCrLf & _
                                                  " '" & Session("ORDERNO") & "', " & vbCrLf & _
                                                  " '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "', " & vbCrLf & _
                                                  " '" & Trim(e.UpdateValues(iLoop).NewValues("LabelNo1").ToString()) & "', " & vbCrLf & _
                                                  " '" & Trim(e.UpdateValues(iLoop).NewValues("LabelNo2").ToString()) & "', " & vbCrLf
                                If e.UpdateValues(iLoop).OldValues("coldefect").ToString() > 0 Then 'defect
                                    ls_SQL = ls_SQL + " " & (ls_Label2 - ls_Label1) + 1 & ",'0') " & vbCrLf
                                Else
                                    ls_SQL = ls_SQL + " " & (ls_Label2 - ls_Label1) + 1 & ",'1') " & vbCrLf
                                End If
                                sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                sqlComm.ExecuteNonQuery()

                                'insert sebelum
                                If Microsoft.VisualBasic.Right((e.UpdateValues(iLoop).NewValues("LabelNo1").ToString()), 7) - Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 9), 7) > 0 Then
                                    Dim L2 As String
                                    L2 = Microsoft.VisualBasic.Left(e.UpdateValues(iLoop).NewValues("LABEL").ToString(), 2)
                                    L2 = L2 + Microsoft.VisualBasic.Right("0000000" + Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Right(Trim(e.UpdateValues(iLoop).NewValues("LabelNo1").ToString()), 7) - 1, 7), 7)
                                    ls_SQL = "INSERT INTO ReceiveForwarder_Detailbox Values ( " & vbCrLf & _
                                                      " '" & Trim(txtsuratjalanno.Text) & "', " & vbCrLf & _
                                                      " '" & Trim(txtsupp.Text) & "', " & vbCrLf & _
                                                      " '" & Trim(txtaffiliatecode.Text) & "', " & vbCrLf & _
                                                      " '" & Trim(e.UpdateValues(iLoop).NewValues("PO").ToString()) & "', " & vbCrLf & _
                                                      " '" & Session("ORDERNO") & "', " & vbCrLf & _
                                                      " '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "', " & vbCrLf & _
                                                      " '" & Microsoft.VisualBasic.Left(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 9) & "', " & vbCrLf & _
                                                      " '" & Trim(L2) & "', " & vbCrLf
                                    If e.UpdateValues(iLoop).OldValues("coldefect").ToString() > 0 Then 'defect
                                        ls_SQL = ls_SQL + " " & (Microsoft.VisualBasic.Right(L2, 7) - Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 9), 7)) + 1 & ",'1') " & vbCrLf
                                    Else
                                        ls_SQL = ls_SQL + " " & (Microsoft.VisualBasic.Right(L2, 7) - Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 9), 7)) + 1 & ",'0') " & vbCrLf
                                    End If
                                    sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                End If
                                'insert sesudah
                                If Microsoft.VisualBasic.Right(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 7) - Microsoft.VisualBasic.Right((e.UpdateValues(iLoop).NewValues("LabelNo2").ToString()), 7) > 0 Then
                                    Dim L2 As String
                                    L2 = Microsoft.VisualBasic.Left(e.UpdateValues(iLoop).NewValues("LABEL").ToString(), 2)
                                    L2 = L2 + Microsoft.VisualBasic.Right("0000000" + Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Right(Trim(e.UpdateValues(iLoop).NewValues("LabelNo2").ToString()), 7) + 1, 7), 7)
                                    ls_SQL = "INSERT INTO ReceiveForwarder_Detailbox Values ( " & vbCrLf & _
                                                      " '" & Trim(txtsuratjalanno.Text) & "', " & vbCrLf & _
                                                      " '" & Trim(txtsupp.Text) & "', " & vbCrLf & _
                                                      " '" & Trim(txtaffiliatecode.Text) & "', " & vbCrLf & _
                                                      " '" & Trim(e.UpdateValues(iLoop).NewValues("PO").ToString()) & "', " & vbCrLf & _
                                                      " '" & Session("ORDERNO") & "', " & vbCrLf & _
                                                      " '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "', " & vbCrLf & _
                                                      " '" & Trim(L2) & "', " & vbCrLf & _
                                                      " '" & Trim(Microsoft.VisualBasic.Right(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 9)) & "', " & vbCrLf
                                    If e.UpdateValues(iLoop).OldValues("coldefect").ToString() > 0 Then 'defect
                                        ls_SQL = ls_SQL + " " & (Microsoft.VisualBasic.Right(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 7) - Microsoft.VisualBasic.Right(L2, 7)) + 1 & ",'1') " & vbCrLf
                                    Else
                                        ls_SQL = ls_SQL + " " & (Microsoft.VisualBasic.Right(Trim(e.UpdateValues(iLoop).NewValues("LABEL").ToString()), 7) - Microsoft.VisualBasic.Right(L2, 7)) + 1 & ",'0') " & vbCrLf
                                    End If
                                    sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                                    sqlComm.ExecuteNonQuery()
                                End If

                                'totalBox = totalBox + (CDbl(e.UpdateValues(iLoop).NewValues("colreceivingqty").ToString()) / CDbl(e.UpdateValues(iLoop).NewValues("colqtybox").ToString()))
                            End If

                            'sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                            'sqlComm.ExecuteNonQuery()

                            'insert master
                            sqlstring = "SELECT * FROM dbo.ReceiveForwarder_Master WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                            " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                            " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                            " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf

                            sqlComm = New SqlCommand(sqlstring, cn, sqlTran)
                            Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

                            If sqlRdrM.Read Then
                                'UPDATE
                                ls_SQL = " UPDATE dbo.ReceiveForwarder_Master SET " & vbCrLf & _
                                             " DriverName = '" & Trim(txtdrivername.Text) & "', " & vbCrLf & _
                                             " DriverContact = '" & Trim(txtdrivercontact.Text) & "', " & vbCrLf & _
                                             " NoPol = '" & Trim(txtnopol.Text) & "', " & vbCrLf & _
                                             " JenisArmada = '" & Trim(txtjenisarmada.Text) & "', " & vbCrLf & _
                                             " TotalBox= " & totalBox & "," & vbCrLf & _
                                             " UpdateUser = '" & Session("UserID") & "', " & vbCrLf & _
                                             " UpdateDate = GETDATE() " & vbCrLf & _
                                             " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                             " AND SupplierID = '" & Trim(txtsupp.Text) & "' " & vbCrLf & _
                                             " and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                             " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                             " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf
                            ElseIf Not sqlRdrM.Read Then
                                'INSERT
                                ls_SQL = " INSERT INTO dbo.ReceiveForwarder_Master " & vbCrLf & _
                                            "         ( SuratJalanNo , " & vbCrLf & _
                                            "           AffiliateID, " & vbCrLf & _
                                            "           SupplierID , " & vbCrLf & _
                                            "           PONo, " & vbCrLf & _
                                            "           OrderNo, " & vbCrLf & _
                                            "           ExcelCls, " & vbCrLf & _
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
                                                  "           UpdateUser " & vbCrLf & _
                                                  "         ) " & vbCrLf & _
                                                  " VALUES  ( '" & Trim(txtsuratjalanno.Text) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                                                  "           '" & Session("AFFID") & "', " & vbCrLf & _
                                                  "           '" & Trim(txtsupp.Text) & "' , -- SupplierID - char(10) " & vbCrLf & _
                                                  "           '" & Session("PONO") & "', " & vbCrLf & _
                                                  "           '" & Session("OrderNo") & "', " & vbCrLf & _
                                                  "           '', " & vbCrLf & _
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
                                                  "           '" & Session("UserID") & "'" & vbCrLf & _
                                                  "         ) "

                            End If
                            sqlRdrM.Close()
                            sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlRdrM.Close()
                            'insert master
                            txtstatus.Text = ""
                            Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("YA010IsSubmit") = lblerrmessage.Text

                        Next iLoop
                    End With

                    sqlComm.Dispose()
                    sqlTran.Commit()
                End Using

                cn.Close()
            End Using
            Call colorGrid()
        ElseIf HF.Get("hfTest") = "add" Then 'Add
            Dim ls_PartNos As String = ""
            Dim ls_PartNo As String = ""
            Dim ls_PONos As String = ""
            Dim ls_PO As String = ""
            Dim ls_Label As String = ""
            Dim ls_combination As String = ""

            For iLoop = 0 To e.UpdateValues.Count - 1
                ls_PO = Trim(e.UpdateValues(iLoop).OldValues("PO").ToString())
                Session("PONO") = ls_PO
                ls_PartNo = Trim(e.UpdateValues(iLoop).OldValues("PART").ToString())
                Session("PartNo") = ls_PartNo

                ls_Label = Trim(e.UpdateValues(iLoop).OldValues("LABEL").ToString())
                Session("LABEL") = ls_Label

                If ls_combination = "" Then
                    ls_combination = "'" + ls_PO + ls_PartNo + ls_Label + "'"
                Else
                    ls_combination = ls_combination + ",'" + ls_PO + ls_PartNo + ls_Label + "'"
                End If
                Session("combination") = ls_combination
            Next
        ElseIf HF.Get("hfTest") = "delete" Then 'Delete
            With Grid
                Dim ls_good As Integer
                Dim ls_Defect As Integer
                totalBox = 0
                For iLoop = 0 To e.UpdateValues.Count - 1
                    Using cn As New SqlConnection(clsGlobal.ConnectionString)
                        cn.Open()
                        Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
                            Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)
                            'Delete Data box
                            ls_SQL = " Delete ReceiveForwarder_Detailbox " & vbCrLf & _
                                     " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                     " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                     " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                     " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf & _
                                     " and Label1 = '" & Trim(e.UpdateValues(iLoop).NewValues("LabelNo1").ToString()) & "'" & vbCrLf & _
                                     " and Label2 = '" & Trim(e.UpdateValues(iLoop).NewValues("LabelNo2").ToString()) & "'"
                            sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                            sqlComm.ExecuteNonQuery()

                            ls_good = CDbl(e.UpdateValues(iLoop).NewValues("colgoodreceiving").ToString())
                            ls_Defect = CDbl(e.UpdateValues(iLoop).NewValues("coldefectreceiving").ToString())

                            If e.UpdateValues(iLoop).NewValues("StatusDefect") = 0 Then
                                'update good ke Defect
                                ls_SQL = " Update ReceiveForwarder_Detail set " & vbCrLf & _
                                         " GoodRecQty = GoodRecQty - " & CDbl(ls_good) & " " & vbCrLf & _
                                         " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                         " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                         " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf
                            Else
                                'Defect ke Good
                                ls_SQL = " Update ReceiveForwarder_Detail set " & vbCrLf & _
                                         " DefectRecQty = DefectRecQty - " & CDbl(ls_Defect) & " " & vbCrLf & _
                                         " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND PartNo = '" & Trim(e.UpdateValues(iLoop).NewValues("PART").ToString()) & "' " & vbCrLf & _
                                         " AND SupplierID = '" & Trim(txtsupp.Text) & "' and affiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                         " and PONO = '" & Session("PONO") & "'" & vbCrLf & _
                                         " AND OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf
                            End If
                            sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                        End Using
                    End Using

                Next
            End With
        End If
    End Sub

    Private Sub Grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles Grid.CellEditorInitialize
        If (e.Column.FieldName = "colno" Or e.Column.FieldName = "colorderno" Or e.Column.FieldName = "collabelno" _
            Or e.Column.FieldName = "colpartno" Or e.Column.FieldName = "colpartname" Or e.Column.FieldName = "coluom" _
            Or e.Column.FieldName = "colqtybox" Or e.Column.FieldName = "coldelqty" _
            ) _
        And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
            'If Not IsNothing(e.KeyValue("colpartno")) = "" Then e.Editor.ReadOnly = False
        Else
            e.Editor.ReadOnly = False
        End If

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
                    Else
                        Grid.JSProperties("cpMessage") = Session("MsgRec")
                        lblerrmessage.Text = Session("MsgRec")
                    End If
                    Call colorGrid()
                    Grid.JSProperties("cpMessage") = Session("MsgRec")
                    lblerrmessage.Text = Session("MsgRec")
                    Session.Remove("MsgRec")
                Case "save"
                    Session.Remove("MsgRec")
                    If Session("sstatus") Is Nothing Then Session("sstatus") = "TRUE"
                    'Call up_GridLoad()
                    'If Session("sstatus") = "TRUE" Then Call saveData()
                    Call fillHeader("load")
                    Call up_GridLoad()
                    If Session("sstatus") = "TRUE" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                    ElseIf Session("sstatus") = "FALSE" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                        lblerrmessage.Text = lblerrmessage.Text
                    End If
                    Session("MsgRec") = lblerrmessage.Text
                    Session.Remove("sstatus")

                Case "kosong"

                Case "sendtosupplier"
                    'Call UpdateExcel(True, txtaffiliate.Text, txtsuratjalanno.Text, txtsupp.Text)

                    Call clsMsg.DisplayMessage(lblerrmessage, "1010", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                Case "Delete"
                    Call up_Delete(txtsuratjalanno.Text)
                    'Call fillHeader("load")
                    Call up_GridLoad()
                    txttotalbox.Text = 0

                    Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                Case "addrow"
                    Call fillHeader("load")
                    Call Up_AddCarton()
            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
            Grid.FocusedRowIndex = -1
            Session.Remove("sstatus")

        Finally
            'If (Not IsNothing(Session("YA010Msg"))) Then Grid.JSProperties("cpMessage") = Session("YA010Msg") : Session.Remove("YA010Msg")
            'Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
        End Try
    End Sub

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        Dim x As Integer = CInt(e.VisibleIndex.ToString())
        Dim pPartNo As String

        If x > Grid.VisibleRowCount Then Exit Sub
        If e.DataColumn.FieldName = "colpartno" Then
            pPartNo = e.GetValue("colpartno")
        End If

        With Grid
            If .VisibleRowCount > 0 Then
                If e.GetValue("colpartno") = "" Then
                    If e.DataColumn.FieldName = "LabelNo1" Or e.DataColumn.FieldName = "LabelNo2" Then
                        e.Cell.BackColor = Color.Yellow
                    End If

                    If e.DataColumn.FieldName = "colreceivingbox" Then
                        e.Cell.BackColor = Color.Yellow
                    End If
                    If e.DataColumn.FieldName = "coldefect" Then
                        e.Cell.BackColor = Color.Yellow
                    End If
                End If

                'If e.GetValue("colHgood") = 0 Then
                '    If e.DataColumn.FieldName = "colgoodreceiving" Then
                '        e.Cell.BackColor = Color.Yellow
                '    End If
                '    If e.DataColumn.FieldName = "coldefect" Then
                '        e.Cell.BackColor = Color.Yellow
                '    End If
                'Else
                '    If e.DataColumn.FieldName = "colgoodreceiving" Then
                '        e.Cell.BackColor = Color.White
                '    End If
                '    If e.DataColumn.FieldName = "coldefectreceiving" Then
                '        e.Cell.BackColor = Color.White
                '    End If
                'End If

            End If
        End With
    End Sub

    Private Sub up_Delete(ByVal pSJ As String)
        Dim ls_sql As String

        ls_sql = ""
        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            'Using sqlTran As SqlTransaction = cn.BeginTransaction("Cols")
            Dim sqlComm As New SqlCommand(ls_sql, cn)
            ls_sql = " SELECT * FROM dbo.ReceiveForwarder_Master WHERE SuratJalanNo ='" & pSJ & "'" & vbCrLf & _
                     " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                     " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
                     " AND OrderNo = '" & Session("OrderNo") & "'"

            sqlComm = New SqlCommand(ls_sql, cn)
            Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

            If sqlRdrM.Read Then
                ls_sql = " delete from ReceiveForwarder_Master WHERE SuratJalanNo ='" & pSJ & "'" & vbCrLf & _
                         " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                         " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
                         " AND OrderNo = '" & Session("OrderNo") & "'" & vbCrLf
                ls_sql = ls_sql + " Delete from ReceiveForwarder_Detail WHERE SuratJalanNo ='" & pSJ & "'" & vbCrLf & _
                                  " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                                  " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
                                  " AND OrderNo = '" & Session("OrderNo") & "'" & vbCrLf
                ls_sql = ls_sql + " Delete from ReceiveForwarder_DetailBox WHERE SuratJalanNo ='" & pSJ & "'" & vbCrLf & _
                                  " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                                  " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
                                  " AND OrderNo = '" & Session("OrderNo") & "'" & vbCrLf
                ls_sql = ls_sql + " Update printlabelexport SET suratjalanno_fwd = '' WHERE suratjalanno_fwd ='" & pSJ & "'" & vbCrLf & _
                                  " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                                  " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
                                  " AND OrderNo = '" & Session("OrderNo") & "'"
                sqlRdrM.Close()
                sqlComm = New SqlCommand(ls_sql, cn)
                sqlComm.ExecuteNonQuery()

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

    Private Sub saveData()
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 1
        Dim isStatusNew As Boolean
        Dim pIsUpdate As Boolean
        Dim sqlstring As String
        Dim i As Long = 0
        Dim pReceiveDate As Date
        Dim ls_totalbox As Integer
        isStatusNew = False

        ls_totalbox = 0
        If txttotalbox.Text = "" Then txttotalbox.Text = 0
        pReceiveDate = txtrecdate.Text

        Using cn As New SqlConnection(clsGlobal.ConnectionString)
            cn.Open()

            Using sqlTran As SqlTransaction = cn.BeginTransaction("cols")
                Dim sqlComm As New SqlCommand(ls_SQL, cn, sqlTran)
                With Grid
                    For i = 0 To Grid.VisibleRowCount - 1
                        'cek QTY tidak boleh melebihi Qty
                        If CDbl(Grid.GetRowValues(i, "colgoodreceiving").ToString) > CDbl(Grid.GetRowValues(i, "colHgood").ToString) Then
                            Call clsMsg.DisplayMessage(lblerrmessage, "7013", clsMessage.MsgType.ErrorMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                            Session("sstatus") = "FALSE"
                            Exit Sub
                        Else
                            txtstatus.Text = "TRUE"
                        End If
                        'cek QTY tidak boleh melebihi Qty

                        ls_totalbox = ls_totalbox + (CDbl(Grid.GetRowValues(i, "colgoodreceiving").ToString) / CDbl(Grid.GetRowValues(i, "colqtybox").ToString))

                        'insert master
                        sqlstring = "SELECT * FROM dbo.ReceiveForwarder_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                    " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                                    " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
                                    " AND OrderNo = '" & Session("OrderNo") & "'"

                        sqlComm = New SqlCommand(sqlstring, cn, sqlTran)
                        Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

                        If sqlRdrM.Read Then
                            'UPDATE
                            ls_SQL = " UPDATE dbo.ReceiveForwarder_Master SET " & vbCrLf & _
                                         " DriverName = '" & Trim(txtdrivername.Text) & "', " & vbCrLf & _
                                         " DriverContact = '" & Trim(txtdrivercontact.Text) & "', " & vbCrLf & _
                                         " NoPol = '" & Trim(txtnopol.Text) & "', " & vbCrLf & _
                                         " JenisArmada = '" & Trim(txtjenisarmada.Text) & "', " & vbCrLf & _
                                         " TotalBox = " & Trim(ls_totalbox) & " , " & vbCrLf & _
                                         " UpdateUser = '" & Session("UserID") & "', " & vbCrLf & _
                                         " UpdateDate = GETDATE() " & vbCrLf & _
                                         " WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                         " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
                                         " AND PONO = '" & Session("PONO") & "' " & vbCrLf & _
                                         " AND OrderNo = '" & Session("OrderNo") & "'"
                        ElseIf Not sqlRdrM.Read Then
                            'INSERT
                            ls_SQL = " INSERT INTO dbo.ReceiveForwarder_Master " & vbCrLf & _
                                     " (SuratJalanNo, " & vbCrLf & _
                                     " AffiliateID, " & vbCrLf & _
                                     " SupplierID, " & vbCrLf & _
                                     " PONo, " & vbCrLf & _
                                     " ExcelCls, " & vbCrLf & _
                                     " ReceiveDate, " & vbCrLf & _
                                     " ReceiveBy, " & vbCrLf & _
                                     " JenisArmada, " & vbCrLf & _
                                     " DriverName, " & vbCrLf & _
                                     " DriverContact, "

                            ls_SQL = ls_SQL + " Nopol, " & vbCrLf & _
                                              " TotalBox, " & vbCrLf & _
                                              " EntryDate, " & vbCrLf & _
                                              " EntryUser, " & vbCrLf & _
                                              " UpdateDate, " & vbCrLf & _
                                              " UpdateUser, OrderNo " & vbCrLf & _
                                              "         ) " & vbCrLf & _
                                              " VALUES  ( '" & Trim(txtsuratjalanno.Text) & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                                              "           '" & Session("AFFID") & "', " & vbCrLf & _
                                              "           '" & Trim(txtsupp.Text) & "' , -- SupplierID - char(10) " & vbCrLf & _
                                              "           '" & Session("PONO") & "', " & vbCrLf & _
                                              "           '', " & vbCrLf & _
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
                                              "           '" & Session("UserID") & "', '" & Session("OrderNo") & "' -- UpdateUser - char(15) " & vbCrLf & _
                                              "           ) "

                        End If
                        sqlRdrM.Close()
                        sqlComm = New SqlCommand(ls_SQL, cn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlRdrM.Close()
                        'insert master

                        sqlstring = "SELECT * FROM dbo.ReceiveForwarder_Detail WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                        " AND PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' " & vbCrLf & _
                                        " AND SupplierID = '" & Trim(txtsupp.Text) & "' " & vbCrLf & _
                                        " AND AffiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                        " AND PONO = '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "'" & vbCrLf & _
                                        " and OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf

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
                            'INSERT KANBAN
                            ls_SQL = " INSERT INTO dbo.ReceiveForwarder_Detail " & vbCrLf & _
                                      "         ( SuratJalanNo , " & vbCrLf & _
                                      "           SupplierID , " & vbCrLf & _
                                      "           AffiliateID, " & vbCrLf & _
                                      "           PONo , " & vbCrLf & _
                                      "           PartNo , " & vbCrLf & _
                                      "           OrderNo, " & vbCrLf & _
                                      "           GoodRecQty, " & vbCrLf & _
                                      "           DefectRecQty " & vbCrLf & _
                                      "         ) " & vbCrLf & _
                                      " VALUES  ( '" & txtsuratjalanno.Text & "' , -- SuratJalanNo - char(20) " & vbCrLf

                            ls_SQL = ls_SQL + "           '" & Trim(txtsupp.Text) & "' , -- SupplierID - char(15) " & vbCrLf & _
                                              "           '" & Session("AFFID") & "' , -- PONo - char(20) " & vbCrLf & _
                                              "           '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "' , -- POKansbanCls - char(1) " & vbCrLf & _
                                              "           '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' , -- PartNo - char(120) " & vbCrLf & _
                                              "           '" & Trim(Grid.GetRowValues(i, "colorderno").ToString) & "' , -- UnitCls - char(3) " & vbCrLf & _
                                              "           " & CDbl(Grid.GetRowValues(i, "colgoodreceiving").ToString) & ",  -- RecQty - numeric " & vbCrLf & _
                                              "           " & CDbl(Grid.GetRowValues(i, "coldefectreceiving").ToString) & "  -- RecQty - numeric " & vbCrLf & _
                                              "           ) "


                        ElseIf pIsUpdate = True Then
                            'Update Data
                            ls_SQL = " Update ReceiveForwarder_Detail set " & vbCrLf & _
                                     " GoodRecQty = '" & CDbl(Grid.GetRowValues(i, "colgoodreceiving").ToString) & "', " & vbCrLf & _
                                     " DefectRecQty = '" & CDbl(Grid.GetRowValues(i, "coldefectreceiving").ToString) & "' " & vbCrLf & _
                                     " WHERE suratjalanno ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
                                     " AND PartNo = '" & Trim(Grid.GetRowValues(i, "colpartno").ToString) & "' " & vbCrLf & _
                                     " AND SupplierID = '" & Trim(txtsupp.Text) & "' " & vbCrLf & _
                                     " AND AffiliateID = '" & Session("AFFID") & "'" & vbCrLf & _
                                     " AND PONO = '" & Trim(Grid.GetRowValues(i, "colpono").ToString) & "'" & vbCrLf & _
                                     " and OrderNo = '" & Session("ORDERNO") & "' " & vbCrLf
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
    '        ls_sql = " SELECT * FROM dbo.ReceiveForwarder_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                 " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
    '                 " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
    '                 " AND OrderNo = '" & Session("OrderNo") & "'"

    '        sqlComm = New SqlCommand(ls_sql, cn)
    '        Dim sqlRdrM As SqlDataReader = sqlComm.ExecuteReader()

    '        If sqlRdrM.Read Then
    '            ls_sql = " delete from ReceiveForwarder_Master WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                     " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
    '                     " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
    '                     " AND OrderNo = '" & Session("OrderNo") & "'" & vbCrLf
    '            ls_sql = ls_sql + " Delete from ReceiveForwarder_Detail WHERE SuratJalanNo ='" & Trim(txtsuratjalanno.Text) & "'" & vbCrLf & _
    '                              " AND AffiliateID = '" & Session("AFFID") & "' " & vbCrLf & _
    '                              " AND PONo = '" & Session("PONO") & "'" & vbCrLf & _
    '                              " AND OrderNo = '" & Session("OrderNo") & "'"
    '            sqlRdrM.Close()
    '            sqlComm = New SqlCommand(ls_sql, cn)
    '            sqlComm.ExecuteNonQuery()
    '            Call fillHeader("load")
    '            Call up_GridLoad()

    '            Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
    '            Grid.JSProperties("cpMessage") = lblerrmessage.Text

    '            txtdrivername.Text = ""
    '            txtdrivercontact.Text = ""
    '            txtnopol.Text = ""
    '            txtjenisarmada.Text = ""
    '            txttotalbox.Text = ""

    '        Else
    '            'data ga ada
    '            Call clsMsg.DisplayMessage(lblerrmessage, "2001", clsMessage.MsgType.InformationMessage)
    '            Grid.JSProperties("cpMessage") = lblerrmessage.Text
    '        End If

    '        sqlComm.Dispose()
    '        sqlRdrM.Close()

    '    End Using
    'End Sub

    'Private Sub UpdateExcel(ByVal pIsNewData As Boolean, _
    '                    Optional ByVal pAffCode As String = "", _
    '                    Optional ByVal pSuratJalan As String = "", _
    '                    Optional ByVal pSuppCode As String = "")

    '    Dim ls_SQL As String = "", ls_MsgID As String = ""
    '    Dim admin As String = Session("UserID").ToString

    '    Try
    '        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
    '            sqlConn.Open()

    '            ls_SQL = " UPDATE dbo.ReceivePASI_Master " & vbCrLf & _
    '                      " SET ExcelCls='1'" & vbCrLf & _
    '                      " WHERE SuratJalanNo='" & pSuratJalan & "'  " & vbCrLf & _
    '                      " AND AffiliateID='" & pAffCode & "' " & vbCrLf & _
    '                      " AND SupplierID='" & pSuppCode & "' "

    '            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
    '            sqlComm.ExecuteNonQuery()
    '            sqlComm.Dispose()
    '            sqlConn.Close()
    '        End Using
    '    Catch ex As Exception
    '        Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
    '    End Try
    'End Sub

    'Private Sub ButtonApprove_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ButtonApprove.Callback
    '    Call UpdateExcel(True, txtaffiliate.Text, txtsuratjalanno.Text, txtsupp.Text)
    '    Call clsMsg.DisplayMessage(lblerrmessage, "1010", clsMessage.MsgType.InformationMessage)
    '    ButtonApprove.JSProperties("cpMessage") = lblerrmessage.Text
    'End Sub

    'Private Sub btnPrintGR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintGR.Click
    '    'If txtPASISJNo.Text <> "" Then
    '    Session("E02SupplierSJNo") = txtsuratjalanno.Text
    '    Session("E02AffiliateID") = txtaffiliate.Text
    '    Session("E02KanbanNo") = txtkanbanno.Text
    '    'Else
    '    '    Session("E02SupplierSJNo") = txtSupplierSJNo.Text
    '    'End If

    '    'LOG HEADER
    '    'Session("E02ParamPageLoad") = Trim(txtRecDate.Text) & "|" & _
    '    '                              Trim(txtsupp.Text) & "|" & Trim(txtSupplierName.Text) & "|" & _
    '    '                              Trim(txtDeliveryLocationCode.Text) & "|" & Trim(txtDeliveryLocationName.Text) & "|" & _
    '    '                              Trim(txtSupplierSJNo.Text) & "|" & Trim(txtSupplierPlanDeliveryDate.Text) & "|" & Trim(txtSupplierDeliveryDate.Text) & "|" & _
    '    '                              Trim(txtPASISJNo.Text) & "|" & Trim(txtPASIDeliveryDate.Text) & "|" & _
    '    '                              Trim(txtDriverName.Text) & "|" & Trim(txtDriverContact.Text) & "|" & Trim(txtNoPol.Text) & "|" & Trim(txtJenisArmada.Text) & "|" & Trim(txtTotalBox.Text)

    '    Response.Redirect("~/Receiving/GoodReceivingReport.aspx")
    'End Sub

    Protected Sub btnPrintGR_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPrintGR.Click
        'Session("E02SupplierSJNo") = txtsuratjalanno.Text
        'Session("E02AffiliateID") = txtaffiliate.Text
        'Session("E02KanbanNo") = txtkanbanno.Text

        'Response.Redirect("~/DeliveryExport/GoodReceivingReportExport.aspx")
    End Sub
End Class