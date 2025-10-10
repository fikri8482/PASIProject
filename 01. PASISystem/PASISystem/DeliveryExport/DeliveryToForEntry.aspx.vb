Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports OfficeOpenXml
Imports Microsoft.Office.Interop
Imports System.Net
Imports System.Net.Mail
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO
Imports System.Transactions
Imports System.Collections
Imports System.Reflection

Public Class DeliveryToForEntry
    Inherits System.Web.UI.Page

#Region "DECLARATION"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramaffiliate As String

    'parameter
    Dim pDeliverydate As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pSuratjalanNo As String
    Dim pForwaderCode As String
    Dim pForwaderName As String
    Dim pDriverName As String
    Dim pDriverContact As String
    Dim pNoPol As String
    Dim pJenisArmada As String
    Dim pPO As String
    Dim pPallet As String
    Dim pSupplier As String
    Dim pSupplierName As String
    Dim pStatus As Boolean
    Dim pSuratJalan As String
    Dim pRSJ As String
    Dim pSuppSuratJalan As String
    Dim pFilter As String
    Dim pub_Param As String
    Dim pstatusInsert As String

    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String

    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "E02"

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim param As String = ""
        Dim filterQty As String = ""
        Dim ls_GenerateNo As String = ""

        ls_AllowDelete = clsGlobal.Auth_UserDelete(Session("UserID"), menuID)

        Try
            '=============================================================
            If (Not IsPostBack) AndAlso (Not IsCallback) Then

                Session("MenuDesc") = "DELIVERY TO FORWARDER ENTRY"

                If Session("POList") <> "" Then
                    param = Session("POList").ToString()
                ElseIf Session("TampungDelivery") <> "" Then
                    param = Session("TampungDelivery").ToString()
                Else
                    If IsNothing(Request.QueryString("prm")) = True Then
                        lblerrmessage.Text = ""
                        Exit Sub
                    End If
                    param = Request.QueryString("prm").ToString
                End If

                If param = "  'back'" Then
                    btnsubmenu.Text = "BACK"
                Else
                    If pStatus = False Then
                        pDeliverydate = IIf(IsNothing(Split(param, "|")(0)) = True, "", Split(param, "|")(0))
                        pAffiliateCode = IIf(IsNothing(Split(param, "|")(1)) = True, "", Split(param, "|")(1))
                        pAffiliateName = IIf(IsNothing(Split(param, "|")(2)) = True, "", Split(param, "|")(2))
                        'pSuratjalanNo = Replace(IIf(IsNothing(Split(param, "|")(3)) = True, "", Split(param, "|")(3)), "DAN", "&")
                        pSuratjalanNo = IIf(IsNothing(Split(param, "|")(3)) = True, "", Split(param, "|")(3))
                        pForwaderCode = IIf(IsNothing(Split(param, "|")(4)) = True, "", Split(param, "|")(4))
                        pForwaderName = IIf(IsNothing(Split(param, "|")(5)) = True, "", Split(param, "|")(5))
                        pDriverName = IIf(IsNothing(Split(param, "|")(6)) = True, "", Split(param, "|")(6))
                        pDriverContact = IIf(IsNothing(Split(param, "|")(7)) = True, "", Split(param, "|")(7))
                        pNoPol = IIf(IsNothing(Split(param, "|")(8)) = True, "", Split(param, "|")(8))
                        pJenisArmada = IIf(IsNothing(Split(param, "|")(9)) = True, "", Split(param, "|")(9))
                        pPO = IIf(IsNothing(Split(param, "|")(10)) = True, "", Split(param, "|")(10))
                        pPallet = IIf(IsNothing(Split(param, "|")(11)) = True, "", Split(param, "|")(11))
                        pSupplier = IIf(IsNothing(Split(param, "|")(12)) = True, "", Split(param, "|")(12))
                        pSupplierName = IIf(IsNothing(Split(param, "|")(13)) = True, "", Split(param, "|")(13))
                        'pSuratJalan = Trim(Replace(IIf(IsNothing(Split(param, "|")(14)) = True, "", Split(param, "|")(14)), "DAN", "&"))
                        pSuratJalan = Trim(IIf(IsNothing(Split(param, "|")(14)) = True, "", Split(param, "|")(14)))

                        Session("PONO") = pPO
                        If Session("POList") <> "" Then
                            pRSJ = Trim(IIf(IsNothing(Split(param, "|")(15)) = True, "", Split(param, "|")(15)))
                        Else
                            pRSJ = "'" & Trim(IIf(IsNothing(Split(param, "|")(15)) = True, "", Split(param, "|")(15))) & "'"
                        End If

                        If Session("POList") <> "" Then
                            'pSuppSuratJalan = Trim(Replace(IIf(IsNothing(Split(param, "|")(16)) = True, "", Split(param, "|")(16)), "DAN", "&"))
                            pSuppSuratJalan = Trim(IIf(IsNothing(Split(param, "|")(16)) = True, "", Split(param, "|")(16)))
                            pFilter = Trim(IIf(IsNothing(Split(param, "|")(17)) = True, "", Split(param, "|")(17)))
                        Else
                            'pSuppSuratJalan = "'" & Trim(Replace(IIf(IsNothing(Split(param, "|")(16)) = True, "", Split(param, "|")(16)), "DAN", "&")) & "'"
                            pSuppSuratJalan = "'" & Trim(IIf(IsNothing(Split(param, "|")(16)) = True, "", Split(param, "|")(16))) & "'"
                        End If

                        pstatusInsert = IIf(IsNothing(Split(param, "|")(18)) = True, "", Split(param, "|")(18))
                        Session("statusinsert") = pstatusInsert

                        If Session("POList") <> "" Then pPallet = Session("PalletList")
                        If Session("POList") <> "" Then pPO = Session("POList2")

                        If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"
                        If Trim(pDeliverydate) = "01 Jan 1900" Then pDeliverydate = Format(Now, "dd MMM yyyy")
                        If Trim(pDeliverydate) = "" Then pDeliverydate = Format(Now, "dd MMM yyyy")

                        dt1.Text = Format(Now, "dd MMM yyyy")
                        txtaffiliatecode.Text = pAffiliateCode
                        txtaffiliatename.Text = pAffiliateName
                        txtForwarderCode.Text = pForwaderCode
                        txtForwarderName.Text = pForwaderName
                        txtSupplierCode.Text = pSupplier
                        txtSupplierName.Text = pSupplierName
                        Session("sSuppID") = pSupplier

                        pStatus = True

                        If pSuratjalanNo <> "" Then
                            txtsuratjalanno.Text = pSuratjalanNo
                            txtdrivername.Text = pDriverName
                            txtdrivercontact.Text = pDriverContact
                            txtnopol.Text = pNoPol
                            txtjenisarmada.Text = pJenisArmada
                            Call up_IsiInvoice(pSuratjalanNo)
                        ElseIf Session("TampungDelivery") <> "" Then
                            Call up_IsiInvoice(Session("Sj"))
                        End If

                        If pRSJ <> "" Then Call up_IsiInvoice(pSuratJalan)

                        Call up_GridLoad(pPO, pSuratJalan, pRSJ, pSuppSuratJalan, pPallet, pFilter, pstatusInsert)
                        txttotalbox.Text = uf_SumQty(pPO, pPallet, pSuratJalan, pFilter, pstatusInsert)
                        If pstatusInsert = "DELIVERY" Then
                            ls_GenerateNo = CreateInvoiceNo(pAffiliateCode)
                            txtsuratjalanno.Text = ls_GenerateNo
                            txtInvoiceNo.Text = ls_GenerateNo
                        End If
                        Session("PO") = pPO
                        Session("Pallet") = pPallet
                        Session("Sj") = pSuratJalan
                        Session("TampungDelivery") = param

                        Session("pRSJ") = pRSJ
                        Session("pSuppSJ") = pSuppSuratJalan
                        Session("pFilter") = pFilter
                        Session.Remove("POList")
                    End If
                End If

                btnsubmenu.Text = "BACK"
            Else
                If txtsuratjalanno.Text <> "" Then

                    txttotalbox.Text = uf_SumQty(Session("POList2"), Session("PalletList"), Session("Sj"), pFilter, pstatusInsert)
                End If
            End If
            '===============================================================================

            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                lblerrmessage.Text = ""
            End If

            If ls_AllowDelete = False Then btndelete.Enabled = False

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Grid.JSProperties("cpMessage") = lblerrmessage.Text
        Finally
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
            'Session.Remove("PO")
            'Session.Remove("Pallet")
            'Session.Remove("sSuppID")
            'Session.Remove("Sj")
            'Session.Remove("TampungDelivery")

            Session.Remove("POList")
            Session.Remove("PalletList")
            Session.Remove("pFilter")
        End Try
    End Sub

    Private Sub up_GridLoad(ByVal pPO As String, ByVal pSJ As String, ByVal pRSJ As String, ByVal pSupSJ As String, ByVal pPallet As String, ByVal pFilter As String, ByVal pStatus As String)
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim PasiSuratJalan As String

        PasiSuratJalan = txtsuratjalanno.Text

        pSJ = Replace(pSJ, "'", "")

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If pStatus = "DELIVERY" Then
                ls_SQL = ls_SQL + "SELECT DISTINCT colno = 0 , " & vbCrLf
            Else
                ls_SQL = ls_SQL + "SELECT DISTINCT colno = (case when isnull(PDM.SuratJalanNo,'') = '' then 0 else 1 END) , " & vbCrLf
            End If
            ls_SQL = ls_SQL + "         colpono = POM.PONo , " & vbCrLf & _
                              "         colpartno = POD.PartNo , " & vbCrLf & _
                              "         colpartname = MP.PartName , " & vbCrLf & _
                              "         coluom = UC.Description , " & vbCrLf & _
                              "         colCls = UC.unitcls ," & vbCrLf & _
                              "         colForwarder = KD.ForwarderID, "

            ls_SQL = ls_SQL + "         colQtyBox = ROUND(CONVERT(CHAR, ISNULL(MPM.QtyBox, 0)),0) , " & vbCrLf & _
                              "         colsuppdelqty = ROUND(CONVERT(CHAR,ISNULL(REM.DOQty, 0),0),0) , " & vbCrLf & _
                              "         colpasigoodrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.GoodRecQty,0)))), " & vbCrLf & _
                              "         colpasidefectrec = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),ISNULL(PRD.DefectRecQty,0)))), " & vbCrLf & _
                              "         colpasiremaining = ROUND(CONVERT(CHAR, ( ISNULL(REM.DOQty, 0) " & vbCrLf & _
                              "                                            - ( ISNULL(PRD.GoodRecQty, 0) " & vbCrLf & _
                              "                                         + ISNULL(PRD.DefectRecQty, 0) ) )),0) , " & vbCrLf

            If pStatus = "DELIVERY" Then
                ls_SQL = ls_SQL + "         colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN 0 ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END))), " & vbCrLf
            Else
                ls_SQL = ls_SQL + "         colpasideliveryqty = CONVERT(CHAR,(CONVERT(NUMERIC(9,0),isnull(REM.DOQty,0)))), " & vbCrLf
            End If

            ls_SQL = ls_SQL + "         colremainingdelqty = CASE WHEN ISNULL(PDD.SuratJalanNo,'') = '' THEN 0 else CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(REM.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END))) END,  " & vbCrLf

            If pStatus = "DELIVERY" Then
                ls_SQL = ls_SQL + "         coldelqtybox = CEILING(CONVERT(CHAR,CASE MPM.QtyBox " & vbCrLf & _
                                  "                          WHEN 0 THEN 0 " & vbCrLf & _
                                  "                          ELSE (CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(REM.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END)))) / MPM.QtyBox " & vbCrLf & _
                                  "                        END,0)), " & vbCrLf
            Else
                ls_SQL = ls_SQL + "         coldelqtybox = CEILING(CONVERT(CHAR,CASE MPM.QtyBox " & vbCrLf & _
                                  "                          WHEN 0 THEN 0 " & vbCrLf & _
                                  "                          ELSE ISNULL(REM.DOQty, 0) / MPM.QtyBox " & vbCrLf & _
                                  "                        END,0)), " & vbCrLf
            End If


            If pStatus = "DELIVERY" Then
                ls_SQL = ls_SQL + "         colstsDO = '' , " & vbCrLf
            Else
                ls_SQL = ls_SQL + "         colstsDO = ISNULL(PDD.PartNo,''), " & vbCrLf
            End If

            ls_SQL = ls_SQL + "         suppsj = isnull(SDM.SuratJalanNo,''), SupplierID = isnull(sdd.supplierID,'') " & vbCrLf & _
                              "         FROM   dbo.PO_Master_Export POM  "

            ls_SQL = ls_SQL + "          LEFT JOIN dbo.PO_Detail_Export POD ON POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                              "                                     AND POM.PoNo = POD.PONo  " & vbCrLf & _
                              "                                     AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                              " 		 LEFT JOIN dbo.ShippingInstruction_DetailPallet KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 												AND KD.AffiliateID = POM.AffiliateID " & vbCrLf & _
                              " 												AND KD.ForwarderID = POD.ForwarderID " & vbCrLf & _
                              " 												AND KD.OrderNo = POM.PONo " & vbCrLf & _
                              " 												AND KD.SupplierID = POM.SupplierID " & vbCrLf & _
                              "          LEFT JOIN dbo.DOSupplier_Detail_Export SDD ON KD.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
                              "                                                 AND KD.SupplierID = SDD.SupplierID  " & vbCrLf & _
                              "                                                 AND KD.PartNo = SDD.PartNo  "

            ls_SQL = ls_SQL + "                                                 AND KD.OrderNo = SDD.PONo  " & vbCrLf & _
                              "          LEFT JOIN dbo.DOSupplier_Master_Export SDM ON SDM.AffiliateID = SDD.AffiliateID  " & vbCrLf & _
                              "                                                 AND SDM.SuratJalanNo = SDD.SuratJalanNo  " & vbCrLf & _
                              "                                                 AND SDM.SupplierID = SDD.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.ReceiveForwarder_Detail PRD ON KD.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                              "                                                  AND KD.SupplierID = PRD.SupplierID  " & vbCrLf & _
                              "                                                  AND KD.PartNo = PRD.PartNo  " & vbCrLf & _
                              "                                                  AND KD.OrderNo = PRD.PONO  " & vbCrLf & _
                              "                                                  AND SDM.SuratJalanno = PRD.SuratJalanNo  " & vbCrLf & _
                              "          LEFT JOIN dbo.ReceiveForwarder_Master PRM ON PRM.AffiliateID = PRD.AffiliateID  " & vbCrLf & _
                              "                                                  AND PRM.SuratJalanNo = PRD.SuratJalanNo  "

            ls_SQL = ls_SQL + "                                                  AND PRM.SupplierID = PRD.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.DOPASI_Detail_Export PDD ON KD.AffiliateID = PDD.AffiliateID  " & vbCrLf & _
                              "                                             AND KD.SupplierID = PDD.SupplierID  " & vbCrLf & _
                              "                                             AND KD.PartNo = PDD.PartNo  " & vbCrLf & _
                              "                                             AND KD.OrderNo = PDD.OrderNo " & vbCrLf & _
                              "          LEFT JOIN dbo.DOPASI_Master_Export PDM ON PDD.AffiliateID = PDM.AffiliateID  " & vbCrLf & _
                              "                                             AND PDD.SuratJalanNo = PDM.SuratJalanNo  " & vbCrLf & _
                              "          LEFT JOIN (SELECT ShippingInstructionNo,PalletNo,AffiliateID,ForwarderID,SupplierID,PartNo,OrderNo, DOQty = Count(ISNULL(LabelNo,0)), SuratJalanNo  " & vbCrLf & _
                              "                      FROM dbo.DOPASI_Detail_Export GROUP BY ShippingInstructionNo,PalletNo,AffiliateID,ForwarderID,SupplierID,PartNo,OrderNo, SuratJalanNo) REM    " & vbCrLf & _
                              "              ON KD.AffiliateID = REM.AffiliateID         "

            ls_SQL = ls_SQL + "              AND KD.SupplierID = REM.SupplierID     " & vbCrLf & _
                              "              AND KD.PartNo = REM.PartNo     " & vbCrLf & _
                              "              AND KD.OrderNo = REM.OrderNo " & vbCrLf & _
                              "              AND SDM.SuratJalanNo = REM.SuratJalanNo  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SDD.PartNo AND MPM.AffiliateID = SDD.AffiliateID AND MPM.SupplierID = SDD.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID           " & vbCrLf & _
                              "   WHERE  'A' = 'A' "

            If pPallet <> "" Then
                ls_SQL = ls_SQL + " AND KD.PalletNo IN (" & Trim(pPallet) & ")  " & vbCrLf
            End If
            If pStatus = "DELIVERY" Then
                If pPO <> "" Then
                    ls_SQL = ls_SQL + " AND KD.OrderNo IN (" & Trim(pPO) & ") " & vbCrLf
                End If
            Else
                ls_SQL = ls_SQL + " AND KD.OrderNo IN (" & Trim(pPO) & ") " & vbCrLf
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count <> 0 Then

                With Grid
                    .DataSource = ds.Tables(0)
                    .DataBind()
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowPager, False, False, False, True)
                    'Call ColorGrid()
                End With

            End If
            sqlConn.Close()

        End Using
    End Sub

    Private Function EmailToEmailCCNotif(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                     "  select 'AFF' flag,affiliatepocc, affiliatepoto,FromEmail = '' from ms_emailaffiliate where AffiliateID='" & Trim(pAfffCode) & "'" & vbCrLf & _
                     " union all " & vbCrLf & _
                     " --PASI TO -CC " & vbCrLf & _
                     " select 'PASI' flag,affiliatepocc,affiliatepoto='',FromEmail = affiliatepoto from ms_emailPASI where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Sub GetSettingEmail()
        Dim ls_SQL As String = ""
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.Ms_EmailSetting"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                smtpClient = Trim(ds.Tables(0).Rows(0)("SMTP"))
                portClient = Trim(ds.Tables(0).Rows(0)("PORTSMTP"))
                usernameSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("usernameSMTP")), "", ds.Tables(0).Rows(0)("usernameSMTP"))
                PasswordSMTP = If(IsDBNull(ds.Tables(0).Rows(0)("passwordSMTP")), "", ds.Tables(0).Rows(0)("passwordSMTP"))
            End If
        End Using
    End Sub

    Private Sub sendEmailtoAffiliate()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""
            Dim ls_Line1 As String = "", ls_Line2 As String = "", ls_Line3 As String = "", ls_Line4 As String = "", ls_Line5 As String = ""
            Dim ls_Line6 As String = "", ls_Line7 As String = "", ls_Line8 As String = ""
            Dim ls_Body As String = ""

            Dim ls_URl As String = "http://" & clsNotification.pub_ServerNameAffiliate & "/Receiving/ReceivingEntry.aspx?id2=" & clsNotification.EncryptURL(txtsuratjalanno.Text.Trim)

            ls_Body = clsNotification.GetNotification("13", ls_URl, Trim(Session("PPONO")))

            Dim dsEmail As New DataSet
            dsEmail = EmailToEmailCCNotif(Trim(txtaffiliatecode.Text), "PASI", "")
            '1 CC Affiliate,'2 CC PASI,'3 CC & TO Supplier
            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("FromEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next
            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            If receiptEmail = "" Then
                MsgBox("Recipient's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            End If

            If fromEmail = "" Then
                MsgBox("Mailer's e-mail address is not found", vbCritical, "Warning")
                Exit Sub
            End If

            Dim mailMessage As New Mail.MailMessage()
            mailMessage.From = New MailAddress(fromEmail)
            mailMessage.Subject = "PASI Approval PONo: " & Trim(Session("PPONo"))

            If receiptEmail <> "" Then
                For Each recipient In receiptEmail.Split(";"c)
                    If recipient <> "" Then
                        Dim mailAddress As New MailAddress(recipient)
                        mailMessage.To.Add(mailAddress)
                    End If
                Next
            End If
            If receiptCCEmail <> "" Then
                For Each recipientCC In receiptCCEmail.Split(";"c)
                    If recipientCC <> "" Then
                        Dim mailAddress As New MailAddress(recipientCC)
                        mailMessage.CC.Add(mailAddress)
                    End If
                Next
            End If
            GetSettingEmail()
            mailMessage.Body = ls_Body
            mailMessage.IsBodyHtml = False
            Dim smtp As New SmtpClient
            smtp.Host = smtpClient
            If smtp.UseDefaultCredentials = True Then
                smtp.EnableSsl = True
            Else
                smtp.EnableSsl = False
                Dim myCredential As System.Net.NetworkCredential = New System.Net.NetworkCredential(Trim(usernameSMTP), Trim(PasswordSMTP))
                smtp.Credentials = myCredential
            End If

            smtp.Port = portClient
            smtp.Send(mailMessage)
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblerrmessage, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try


    End Sub

    Private Sub up_IsiInvoice(ByVal pSJ As String)
        Dim ls_SQL As String = ""
        Dim ls_HT As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()


            ls_SQL = "  SELECT InvoiceNo = ISNULL(InvoiceNo,''), " & vbCrLf & _
                  " 		SuratJalanNo = ISNULL(SuratJalanNo,''), " & vbCrLf & _
                  "         DeliveryDate = CONVERT(char(11), CONVERT(datetime, DeliveryDate),106), " & vbCrLf & _
                  " 		DriverName = ISNULL(DriverName,''), " & vbCrLf & _
                  " 		DriverContact = ISNULL(DriverContact,''), " & vbCrLf & _
                  " 		NoPol = ISNULL(NoPol,''), " & vbCrLf & _
                  " 		JenisArmada = ISNULL(JenisArmada,''), " & vbCrLf & _
                  " 		HT_Cls = ISNULL(HT_Cls,'0'), " & vbCrLf & _
                  "         TotalPalet = (SELECT (COUNT(A.PalletNo)) FROM ( " & vbCrLf & _
                  "			              SELECT DISTINCT PalletNo FROM dbo.DOPASI_Detail_Export " & vbCrLf & _
                  "                 	  WHERE SuratJalanNo = '" & pSJ & "' " & vbCrLf & _
                  "	                      GROUP BY PalletNo) A) " & vbCrLf & _
                  "  FROM DOPASI_Master_Export WHERE SuratJalanNo = '" & pSJ & "' " & vbCrLf & _
                  "  "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Try
                    With ds.Tables(0)
                        txtInvoiceNo.Text = Trim(.Rows(0).Item("InvoiceNo"))
                        txtsuratjalanno.Text = Trim(.Rows(0).Item("SuratJalanNo"))
                        txtdrivername.Text = Trim(.Rows(0).Item("DriverName"))
                        txtdrivercontact.Text = Trim(.Rows(0).Item("DriverContact"))
                        txtnopol.Text = Trim(.Rows(0).Item("NoPol"))
                        txtjenisarmada.Text = Trim(.Rows(0).Item("JenisArmada"))
                        txttotalpalet.Text = Trim(.Rows(0).Item("TotalPalet"))
                        ls_HT = Trim(.Rows(0).Item("HT_Cls"))
                        HF.Set("HTcls", ls_HT)
                        dt1.Text = Trim(.Rows(0).Item("DeliveryDate"))
                    End With
                Catch ex As Exception

                End Try
            End If
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_UPDATEMASTER()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_PIC As String = Trim(Session("UserID").ToString)

        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)
        Session("Sj") = ls_Sjno
        Dim ls_SupplierID As String = Session("sSuppID")
        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)

        Dim ls_DeliveryDate As Date = dt1.Text
        ls_DeliveryDate = Format(CDate(ls_DeliveryDate), "yyyy-MM-dd")
        Dim ls_JenisArmada As String = Trim(txtjenisarmada.Text)
        Dim ls_DriverName As String = Trim(txtdrivername.Text)
        Dim ls_DriverContact As String = Trim(txtdrivercontact.Text)
        Dim ls_NoPol As String = Trim(txtnopol.Text)
        Dim ls_TotalBox As String = Trim(txttotalbox.Text)
        Dim ls_InvoiceNo As String = Trim(txtInvoiceNo.Text)
        Dim ls_palet As String = Trim(txttotalpalet.Text)
        Dim ls_forwarder As String = Trim(txtForwarderCode.Text)

        Session("SJPASI") = txtsuratjalanno.Text
        If ls_palet = "" Then ls_palet = 0

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlScope As New TransactionScope
                Try
                    ls_SQL = " UPDATE dbo.DOPASI_Master_Export " & vbCrLf & _
                             " SET DeliveryDate ='" & ls_DeliveryDate & "', " & vbCrLf & _
                             " 	PIC ='" & ls_PIC & "', " & vbCrLf & _
                             " 	JenisArmada ='" & ls_JenisArmada & "', " & vbCrLf & _
                             " 	DriverName ='" & ls_DriverName & "', " & vbCrLf & _
                             "  DriverContact ='" & ls_DriverContact & "', " & vbCrLf & _
                             "     NoPol ='" & ls_NoPol & "', " & vbCrLf & _
                             "     TotalBox ='" & ls_TotalBox & "', " & vbCrLf & _
                             "     InvoiceNo ='" & ls_InvoiceNo & "', " & vbCrLf & _
                             "     UpdateDate = GETDATE(), " & vbCrLf

                    ls_SQL = ls_SQL + "     UpdateUser ='" & ls_PIC & "', " & vbCrLf & _
                                      "     TotalPalet = (SELECT (COUNT(A.PalletNo)) FROM ( " & vbCrLf & _
                                      "			         SELECT DISTINCT PalletNo FROM dbo.DOPASI_Detail_Export " & vbCrLf & _
                                      "            	     WHERE SuratJalanNo = '" & ls_Sjno & "' " & vbCrLf & _
                                      "	                 GROUP BY PalletNo) A) " & vbCrLf & _
                                      " WHERE SuratJalanNo = '" & ls_Sjno & "'  " & vbCrLf & _
                                      "   AND AffiliateID = '" & ls_AffiliateID & "' " & vbCrLf & _
                                      "   AND SupplierID  = '" & ls_SupplierID & "' " & vbCrLf & _
                                      "   AND ForwarderID = '" & ls_forwarder & "' "

                    Dim sqlCommH As New SqlCommand(ls_SQL, sqlConn)
                    sqlCommH.ExecuteNonQuery()
                    sqlCommH.Dispose()
                    sqlScope.Complete()

                Catch ex As Exception

                End Try

            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Function uf_SumQty(ByVal pPO As String, ByVal pPalet As String, ByVal pSJ As String, ByVal pFilter As String, ByVal pStatus As String)
        Dim ls_SQL As String = ""
        Dim PasiSuratJalan As String

        PasiSuratJalan = txtsuratjalanno.Text
        pFilter = "'" & (Replace(pFilter, "'", "")) & "'"

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If pStatus = "DELIVERY" Then
                ls_SQL = ls_SQL + " SELECT coldelqtybox = isnull(SUM(CEILING(CONVERT(CHAR,CASE MPM.QtyBox " & vbCrLf & _
                                  "                          WHEN 0 THEN 0 " & vbCrLf & _
                                  "                          ELSE (CONVERT(CHAR,(CONVERT(NUMERIC(9,0),CASE WHEN ISNULL(PRD.GoodRecQty,0) = 0 THEN (ISNULL(REM.DOQty,0) - ISNULL(REM.DOQty,0)) ELSE (ISNULL(PRD.GoodRecQty,0) - ISNULL(REM.DOQty,0)) END)))) / MPM.QtyBox " & vbCrLf & _
                                  "                        END,0))),0) " & vbCrLf
            Else
                ls_SQL = ls_SQL + " SELECT coldelqtybox = isnull(SUM(CEILING(CONVERT(CHAR,CASE MPM.QtyBox " & vbCrLf & _
                                  "                          WHEN 0 THEN 0 " & vbCrLf & _
                                  "                          ELSE ISNULL(REM.DOQty, 0) / MPM.QtyBox " & vbCrLf & _
                                  "                        END,0))),0) " & vbCrLf
            End If

            ls_SQL = ls_SQL + " FROM   dbo.PO_Master_Export POM             " & vbCrLf & _
                  " 		LEFT JOIN dbo.PO_Detail_Export POD ON POM.AffiliateID = POD.AffiliateID   " & vbCrLf & _
                  "                                      AND POM.PoNo = POD.PONo   " & vbCrLf & _
                  "                                      AND POM.SupplierID = POD.SupplierID   " & vbCrLf & _
                  "  		 LEFT JOIN dbo.ShippingInstruction_DetailPallet KD ON KD.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                  "  												AND KD.AffiliateID = POM.AffiliateID  " & vbCrLf & _
                  "  												AND KD.ForwarderID = POD.ForwarderID  " & vbCrLf & _
                  "  												AND KD.OrderNo = POM.PONo  " & vbCrLf & _
                  "  												AND KD.SupplierID = POM.SupplierID  " & vbCrLf & _
                  "           LEFT JOIN dbo.DOSupplier_Detail_Export SDD ON KD.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
                  "                                                  AND KD.SupplierID = SDD.SupplierID   "

            ls_SQL = ls_SQL + "                                                  AND KD.PartNo = SDD.PartNo                                                    " & vbCrLf & _
                              " 												 AND KD.OrderNo = SDD.PONo   " & vbCrLf & _
                              "           LEFT JOIN dbo.DOSupplier_Master_Export SDM ON SDM.AffiliateID = SDD.AffiliateID   " & vbCrLf & _
                              "                                                  AND SDM.SuratJalanNo = SDD.SuratJalanNo   " & vbCrLf & _
                              "                                                  AND SDM.SupplierID = SDD.SupplierID   " & vbCrLf & _
                              "           LEFT JOIN dbo.ReceiveForwarder_Detail PRD ON KD.AffiliateID = PRD.AffiliateID   " & vbCrLf & _
                              "                                                   AND KD.SupplierID = PRD.SupplierID   " & vbCrLf & _
                              "                                                   AND KD.PartNo = PRD.PartNo   " & vbCrLf & _
                              "                                                   AND KD.OrderNo = PRD.PONO   " & vbCrLf & _
                              "                                                   AND SDM.SuratJalanno = PRD.SuratJalanNo   " & vbCrLf & _
                              "           LEFT JOIN dbo.ReceiveForwarder_Master PRM ON PRM.AffiliateID = PRD.AffiliateID   "

            ls_SQL = ls_SQL + "                                                   AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                              " 												  AND PRM.SupplierID = PRD.SupplierID   " & vbCrLf & _
                              "           LEFT JOIN dbo.DOPASI_Detail_Export PDD ON KD.AffiliateID = PDD.AffiliateID   " & vbCrLf & _
                              "                                              AND KD.SupplierID = PDD.SupplierID   " & vbCrLf & _
                              "                                              AND KD.PartNo = PDD.PartNo   " & vbCrLf & _
                              "                                              AND KD.OrderNo = PDD.OrderNo  " & vbCrLf & _
                              "  											AND SDM.SuratJalanNo = PDD.SuratJalanNo  " & vbCrLf & _
                              "           LEFT JOIN dbo.DOPASI_Master_Export PDM ON PDD.AffiliateID = PDM.AffiliateID   " & vbCrLf & _
                              "                                              AND PDD.SuratJalanNo = PDM.SuratJalanNo   " & vbCrLf & _
                              "          LEFT JOIN (SELECT ShippingInstructionNo,PalletNo,AffiliateID,ForwarderID,SupplierID,PartNo,OrderNo, DOQty = Count(ISNULL(LabelNo,0)), SuratJalanNo   " & vbCrLf & _
                              "                       FROM dbo.ShippingInstruction_DetailPallet GROUP BY ShippingInstructionNo,PalletNo,AffiliateID,ForwarderID,SupplierID,PartNo,OrderNo, SuratJalanNo) REM     "

            ls_SQL = ls_SQL + "               ON KD.AffiliateID = REM.AffiliateID                       AND KD.SupplierID = REM.SupplierID      " & vbCrLf & _
                              "               AND KD.PartNo = REM.PartNo      " & vbCrLf & _
                              "               AND KD.OrderNo = REM.OrderNo  " & vbCrLf & _
                              "               AND SDM.SuratJalanNo = REM.SuratJalanNo   " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Parts MP ON MP.PartNo = POD.PartNo  " & vbCrLf & _
                              "          LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SDD.PartNo AND MPM.AffiliateID = SDD.AffiliateID AND MPM.SupplierID = SDD.SupplierID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Affiliate MA ON MA.AffiliateID = POM.AffiliateID  " & vbCrLf & _
                              "          LEFT JOIN dbo.MS_Supplier MS ON MS.SupplierID = POM.SupplierID  " & vbCrLf & _
                              "          Where 'A' = 'A'"

            If pPO <> "" Then
                ls_SQL = ls_SQL + " And REM.OrderNo in (" & pPO.Trim & ")  "
            End If

            If pSJ <> "" Then
                ls_SQL = ls_SQL + " And REM.SuratJalanNo = '" & pSJ.Trim & "'  "
            End If

            If pPalet <> "" Then
                ls_SQL = ls_SQL + " And REM.PalletNo in (" & pPalet.Trim & ")  "
            End If

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            uf_SumQty = ds.Tables(0).Rows(0)("coldelqtybox")
            sqlConn.Close()


        End Using
    End Function

    Private Function Save_Master(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pForwaderID As String, ByVal pDeliveryDate As String, ByVal pPIC As String, _
                            ByVal pjenisArmada As String, ByVal pDriverName As String, ByVal pDriverContact As String, ByVal pNopol As String, ByVal pTotalBox As String, ByVal pInvoiceNo As String, ByVal pPalet As String)

        Dim ls_sql As String
        If pPalet = "" Then pPalet = 0

        ls_sql = ""
        ls_sql = ls_sql + " IF EXISTS(SELECT * FROM DOPASI_Master_Export WHERE SuratJalanNo = '" & pSjno & "' AND AffiliateID = '" & pAffiliateID & "' AND ForwarderID = '" & pForwaderID & "' AND SupplierID = '" & pSupplierID & "') " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " UPDATE dbo.DOPASI_Master_Export " & vbCrLf & _
                          " SET DeliveryDate ='" & pDeliveryDate & "', " & vbCrLf & _
                          " 	PIC ='" & pPIC & "', " & vbCrLf & _
                          " 	JenisArmada ='" & pjenisArmada & "', " & vbCrLf & _
                          " 	DriverName ='" & pDriverName & "', " & vbCrLf & _
                          "     DriverContact ='" & pDriverContact & "', " & vbCrLf & _
                          "     NoPol ='" & pNopol & "', " & vbCrLf & _
                          "     TotalBox ='" & pTotalBox & "', " & vbCrLf & _
                          "     InvoiceNo ='" & pInvoiceNo & "', " & vbCrLf & _
                          "     UpdateDate = GETDATE(), " & vbCrLf

        ls_sql = ls_sql + "     UpdateUser ='" & pPIC & "', " & vbCrLf & _
                          "     TotalPalet = " & pPalet & " " & vbCrLf & _
                          " WHERE SuratJalanNo = '" & pSjno & "'  " & vbCrLf & _
                          "   AND SupplierID = '" & pSupplierID & "'  " & vbCrLf & _
                          "   AND AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                          "   AND ForwarderID = '" & pForwaderID & "' " & vbCrLf & _
                          " END " & vbCrLf & _
                          " ELSE " & vbCrLf & _
                          " BEGIN " & vbCrLf & _
                          " INSERT INTO dbo.DOPASI_Master_Export " & vbCrLf & _
                          "         ( SuratJalanNo ,SupplierID ,AffiliateID ,ForwarderID ,DeliveryDate ,PIC ,JenisArmada ,DriverName , " & vbCrLf & _
                          "           DriverContact ,NoPol ,TotalBox ,InvoiceNo ,EntryDate ,EntryUser, totalpalet " & vbCrLf & _
                          "         ) " & vbCrLf

        ls_sql = ls_sql + " VALUES  ( '" & pSjno & "' , -- SuratJalanNo - char(20) " & vbCrLf & _
                          "           '" & pSupplierID & "' , -- SupplierID - char(20) " & vbCrLf & _
                          "           '" & pAffiliateID & "' , -- AffiliateID - char(20) " & vbCrLf & _
                          "           '" & pForwaderID & "' , -- ForwaderCode - char(20) " & vbCrLf & _
                          "           '" & pDeliveryDate & "' , -- DeliveryDate - date " & vbCrLf & _
                          "           '" & pPIC & "' , -- PIC - char(15) " & vbCrLf & _
                          "           '" & pjenisArmada & "' , -- JenisArmada - char(15) " & vbCrLf & _
                          "           '" & pDriverName & "' , -- DriverName - char(15) " & vbCrLf & _
                          "           '" & pDriverContact & "' , -- DriverContact - char(15) " & vbCrLf & _
                          "           '" & pNopol & "' , -- NoPol - char(10) " & vbCrLf & _
                          "           " & pTotalBox & " , -- TotalBox - numeric " & vbCrLf & _
                          "           '" & pInvoiceNo & "' , " & vbCrLf & _
                          "           GETDATE() , -- EntryDate - datetime " & vbCrLf

        ls_sql = ls_sql + "           '" & pPIC & "',  -- EntryUser - char(15) " & vbCrLf & _
                          "           " & pPalet & " " & vbCrLf & _
                          "         )	 " & vbCrLf & _
                          " END " & vbCrLf

        Save_Master = ls_sql
    End Function

    Private Function CreateInvoiceNo(ByVal pAffiliateID As String) As String
        Dim ls_sql As String = ""
        Dim ls_Sfx As String = ""
        Dim ls_SeqNo As String = ""
        Dim ls_Year As String = ""
        Dim ls_Temp As Integer = 0

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                'ls_sql = " select  " & vbCrLf & _
                '      " 	SuffixInvoice =  " & vbCrLf & _
                '      " 	CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + " & vbCrLf & _
                '      " 	AffiliateCls + POCode, SeqNO = SeqNo + 1   " & vbCrLf & _
                '      " from MS_Affiliate where AffiliateID ='" & pAffiliateID & "'" & vbCrLf

                ls_sql = " IF EXISTS " & vbCrLf & _
                          " ( " & vbCrLf & _
                          " 		SELECT CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + AffiliateCls + POCode + convert(varchar(5),SeqNo) + 'T' as CreateNomor " & vbCrLf & _
                          " 		FROM " & vbCrLf & _
                          " 		( " & vbCrLf & _
                          " 		SELECT MA.AffiliateID, MAX(SUBSTRING(SuratJalanNo,5,5)) + 1 as SeqNo FROM DOPASI_Master_Export DPM " & vbCrLf & _
                          " 		LEFT JOIN MS_Affiliate MA on DPM.AffiliateID = MA.AffiliateID " & vbCrLf & _
                          " 		WHERE MA.AffiliateID = '" & pAffiliateID & "' and SUBSTRING(SuratJalanNo,1,4) = CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + AffiliateCls + POCode " & vbCrLf & _
                          " 		GROUP BY MA.AffiliateID " & vbCrLf & _
                          " 		) x LEFT JOIN " & vbCrLf & _
                          " 		( " & vbCrLf

                ls_sql = ls_sql + " 			SELECT AffiliateID, AffiliateCls, OverseasCls, POCode FROM MS_Affiliate " & vbCrLf & _
                                  " 		)z ON z.AffiliateID = x.AffiliateID " & vbCrLf & _
                                  " ) " & vbCrLf & _
                                  " 	BEGIN " & vbCrLf & _
                                  " 		SELECT CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + AffiliateCls + POCode + convert(varchar(5),SeqNo) + 'T' as CreateNomor " & vbCrLf & _
                                  " 		FROM " & vbCrLf & _
                                  " 		( " & vbCrLf & _
                                  " 		SELECT MA.AffiliateID, MAX(SUBSTRING(SuratJalanNo,5,5)) + 1 as SeqNo FROM DOPASI_Master_Export DPM " & vbCrLf & _
                                  " 		LEFT JOIN MS_Affiliate MA on DPM.AffiliateID = MA.AffiliateID " & vbCrLf & _
                                  " 		WHERE MA.AffiliateID = '" & pAffiliateID & "' and SUBSTRING(SuratJalanNo,1,4) = CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + AffiliateCls + POCode " & vbCrLf & _
                                  " 		GROUP BY MA.AffiliateID " & vbCrLf

                ls_sql = ls_sql + " 		) x LEFT JOIN " & vbCrLf & _
                                  " 		( " & vbCrLf & _
                                  " 			SELECT AffiliateID, AffiliateCls, OverseasCls, POCode FROM MS_Affiliate " & vbCrLf & _
                                  " 		)z ON z.AffiliateID = x.AffiliateID " & vbCrLf & _
                                  " 	END " & vbCrLf & _
                                  " ELSE " & vbCrLf & _
                                  " 	BEGIN " & vbCrLf & _
                                  " 		SELECT CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + AffiliateCls + POCode + RIGHT(CONVERT(char(4),YEAR(GETDATE())),1) + convert(varchar(5),SeqNo) + 'T' as CreateNomor " & vbCrLf & _
                                  " 		FROM " & vbCrLf & _
                                  " 		( " & vbCrLf & _
                                  " 		SELECT MA.AffiliateID, '0001' as SeqNo FROM  " & vbCrLf

                ls_sql = ls_sql + " 		MS_Affiliate MA  " & vbCrLf & _
                                  " 		WHERE MA.AffiliateID = '" & pAffiliateID & "' " & vbCrLf & _
                                  " 		GROUP BY MA.AffiliateID " & vbCrLf & _
                                  " 		) x LEFT JOIN " & vbCrLf & _
                                  " 		( " & vbCrLf & _
                                  " 			SELECT AffiliateID, AffiliateCls, OverseasCls, POCode FROM MS_Affiliate " & vbCrLf & _
                                  " 		)z ON z.AffiliateID = x.AffiliateID " & vbCrLf & _
                                  " 	END " & vbCrLf & _
                                  "  "


                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                Dim ds2 As New DataSet
                sqlDA2.Fill(ds2)

                'If ds2.Tables(0).Rows.Count > 0 Then
                '    ls_Sfx = ds2.Tables(0).Rows(0)("SuffixInvoice")
                '    ls_Year = Right(CStr(Year(Now)), 1)
                '    ls_Temp = CDbl(ds2.Tables(0).Rows(0)("SeqNo"))
                '    If ls_Temp <= 9 Then
                '        ls_SeqNo = "000" & ds2.Tables(0).Rows(0)("SeqNo")
                '    ElseIf ls_Temp < 99 Then
                '        ls_SeqNo = "00" & ds2.Tables(0).Rows(0)("SeqNo")
                '    ElseIf ls_Temp < 999 Then
                '        ls_SeqNo = "0" & ds2.Tables(0).Rows(0)("SeqNo")
                '    ElseIf ls_Temp < 9999 Then
                '        ls_SeqNo = "0001"
                '    End If
                'Else
                '    ls_Sfx = ""
                'End If
                'CreateInvoiceNo = ls_Sfx + ls_Year + ls_SeqNo + "T"
                If ds2.Tables(0).Rows.Count > 0 Then
                    CreateInvoiceNo = ds2.Tables(0).Rows(0)("CreateNomor")
                Else
                    CreateInvoiceNo = ""
                End If

            End Using
        End Using
    End Function

    Private Function Save_Detail(ByVal pSjno As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pPOno As String, _
                            ByVal pForwarder As String)

        Dim ls_sql As String

        ls_sql = ""
        ls_sql = " IF NOT EXISTS (SELECT * FROM dbo.DOPASI_Detail_Export WHERE AffiliateID = '" & pAffiliateID.Trim & "' AND " & vbCrLf & _
                   "  ForwarderID = '" & pForwarder.Trim & "' AND SuratJalanNo = '" & pSjno.Trim & "' AND OrderNo = '" & pPOno.Trim & "'  AND SupplierID = '" & pSupplierID.Trim & "' ) " & vbCrLf & _
                   " BEGIN " & vbCrLf & _
                   " INSERT INTO dbo.DOPASI_Detail_Export ( AffiliateID ,ForwarderID ,ShippingInstructionNo ,SuratJalanNo ,OrderNo ,LabelNo ,PalletNo ,PalletType ,Location,  " & vbCrLf & _
                   " 									   GrossWeight ,Height ,EntryDate ,EntryUser ,UpdateDate ,UpdateUser ,PrintCount ,PartNo ,SupplierID ,Print_Cls) " & vbCrLf & _
                   " SELECT AffiliateID ,ForwarderID ,ShippingInstructionNo = '" & pSjno.Trim & "',SuratJalanNo = '" & pSjno.Trim & "',OrderNo ,LabelNo ,PalletNo ,PalletType ,Location,  " & vbCrLf & _
                   " GrossWeight ,Height ,EntryDate ,EntryUser ,UpdateDate ,UpdateUser ,PrintCount ,PartNo ,SupplierID ,Print_Cls  " & vbCrLf & _
                   " FROM dbo.ShippingInstruction_DetailPallet " & vbCrLf & _
                   " WHERE AffiliateID = '" & pAffiliateID.Trim & "' AND " & vbCrLf & _
                   " ForwarderID = '" & pForwarder.Trim & "' AND OrderNo = '" & pPOno.Trim & "'  AND SupplierID = '" & pSupplierID.Trim & "' " & vbCrLf & _
                   " END "

        Save_Detail = ls_sql
    End Function

    Private Sub up_Delete(ByVal pSJ As String)
        Dim ls_SQL As String = ""

        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " DELETE dbo.DOPASI_Master_Export WHERE SuratJalanNo = '" & pSJ & "'" & vbCrLf & _
                     " DELETE dbo.DOPASI_Detail_Export WHERE SuratJalanNo = '" & pSJ & "'" & vbCrLf & _
                     " DELETE dbo.DOPASISeq_Detail WHERE SuratJalanNo = '" & pSJ & "'"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)


            Dim sqlComm As New SqlCommand(ls_SQL, sqlConn)
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()

            sqlConn.Close()

            txtsuratjalanno.Text = ""
        End Using
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click
        Session.Remove("PO")
        Session.Remove("Pallet")
        Session.Remove("sSuppID")
        Session.Remove("Sj")
        Session.Remove("TampungDelivery")

        Session.Remove("POList")
        Session.Remove("PalletList")
        Session.Remove("pFilter")

        'remove Request.QueryString("prm")
        '-------------------------------------------
        Dim isreadonly As PropertyInfo = _
        GetType(System.Collections.Specialized.NameValueCollection).GetProperty("IsReadOnly", BindingFlags.Instance Or BindingFlags.NonPublic)

        ' make collection editable
        isreadonly.SetValue(Me.Request.QueryString, False, Nothing)

        ' remove
        If IsNothing(Me.Request.QueryString("prm")) = False Then
            Me.Request.QueryString.Remove("prm")
        End If
        '-------------------------------------------

        If btnsubmenu.Text = "BACK" Then
            Response.Redirect("~/DeliveryExport/DeliveryToForList.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If

    End Sub

    Private Sub Grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles Grid.BatchUpdate
        Dim ls_MsgID As String = ""
        Dim ls_sql As String = ""
        Dim iRow As Integer = 0
        Dim ls_PIC As String = Trim(Session("UserID").ToString)


        Dim ls_Sjno As String = Trim(txtsuratjalanno.Text)
        Dim ls_SupplierID As String = Session("sSuppID")
        Dim ls_AffiliateID As String = Trim(txtaffiliatecode.Text)

        Dim ls_DeliveryDate As Date = dt1.Text
        ls_DeliveryDate = Format(CDate(ls_DeliveryDate), "yyyy-MM-dd")
        Dim ls_JenisArmada As String = Trim(txtjenisarmada.Text)
        Dim ls_DriverName As String = Trim(txtdrivername.Text)
        Dim ls_DriverContact As String = Trim(txtdrivercontact.Text)
        Dim ls_NoPol As String = Trim(txtnopol.Text)
        Dim ls_TotalBox As String = Trim(txttotalbox.Text)
        Dim ls_InvoiceNo As String = Trim(txtInvoiceNo.Text)
        Dim ls_Palet As String = Trim(txttotalpalet.Text)

        Dim ls_POno As String
        Dim ls_Forwarder As String
        Dim ls_PartNo As String
        Dim ls_UnitCls As String
        Dim ls_DOqty As String
        Dim ls_SuppSJ As String
        Dim ls_supID As String


        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Grid.VisibleRowCount = 0 Then Exit Sub
            Session.Remove("msgDelivery")
            Using sqlTran As New TransactionScope
                If e.UpdateValues.Count = 0 Then
                    Session("MsgDelivery") = "NOT SAVED"
                    Exit Sub
                End If

                'Validasi Duplicate Surat Jalan
                ls_sql = "SELECT SuratJalanNo FROM DOPASI_Master_Export WHERE SuratJalanNo = '" & ls_Sjno & "'"
                Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    Session("MsgDelivery") = "NOT SAVED"
                    Exit Sub
                End If

                If Session("statusinsert") = "DELIVERY" Then
                    If Left(ls_Sjno, 2) = "DA" Then
                        ls_sql = "update ms_affiliate set seqno = " & Mid(ls_Sjno, 6, 4) & " where AffiliateID = '" & ls_AffiliateID & "'"
                        Dim sqlComm22 As New SqlCommand(ls_sql, sqlConn)
                        sqlComm22.ExecuteNonQuery()
                        sqlComm22.Dispose()
                    End If
                End If

                For iRow = 0 To e.UpdateValues.Count - 1

                    'cek QTY tidak boleh melebihi Qty
                    If (CDbl(e.UpdateValues(iRow).NewValues("colpasideliveryqty").ToString())) > CDbl(e.UpdateValues(iRow).NewValues("colpasigoodrec").ToString()) Then
                        Session("MsgDelivery") = "QTY"
                        Exit Sub
                    End If
                    'cek QTY tidak boleh melebihi Qty

                    ls_POno = e.UpdateValues(iRow).NewValues("colpono").ToString()
                    ls_Forwarder = e.UpdateValues(iRow).NewValues("colForwarder").ToString()
                    ls_PartNo = e.UpdateValues(iRow).NewValues("colpartno").ToString()
                    ls_UnitCls = e.UpdateValues(iRow).NewValues("colCls").ToString()
                    ls_DOqty = Trim(CDbl(IIf(e.UpdateValues(iRow).NewValues("colpasideliveryqty").ToString(), e.UpdateValues(iRow).NewValues("colpasideliveryqty").ToString(), 0)))
                    ls_SuppSJ = e.UpdateValues(iRow).NewValues("suppsj").ToString()
                    ls_supID = e.UpdateValues(iRow).NewValues("SupplierID").ToString()

                    If (e.UpdateValues(iRow).NewValues("colno").ToString()) = True Then
                        'save Detail
                        ls_sql = Save_Detail(ls_Sjno, ls_supID, ls_AffiliateID, ls_POno, ls_Forwarder)

                        'save master
                        ls_sql = ls_sql + vbCrLf & _
                                 Save_Master(ls_Sjno, ls_supID, ls_AffiliateID, ls_Forwarder, ls_DeliveryDate, ls_PIC, ls_JenisArmada, ls_DriverName, ls_DriverContact, ls_NoPol, ls_TotalBox, ls_InvoiceNo, ls_Palet)

                        Dim sqlComm2 As New SqlCommand(ls_sql, sqlConn)
                        sqlComm2.ExecuteNonQuery()
                        sqlComm2.Dispose()
                        Session("statusinsert") = "DETAIL"
                        Session("MsgDelivery") = "SAVED"
                    Else
                        'delete data detail
                        ls_sql = " Delete FROM dbo.DOPASI_Detail_Export " & vbCrLf & _
                                 " WHERE AffiliateID = '" & ls_AffiliateID.Trim & "' AND " & vbCrLf & _
                                 " ForwarderID = '" & ls_Forwarder.Trim & "' AND SuratJalanNo = '" & ls_Sjno.Trim & "' " & vbCrLf & _
                                 " AND OrderNo = '" & ls_POno.Trim & "'  AND SupplierID = '" & ls_SupplierID.Trim & "' "

                        Dim sqlComm3 As New SqlCommand(ls_sql, sqlConn)
                        sqlComm3.ExecuteNonQuery()
                        sqlComm3.Dispose()
                    End If
                Next iRow
                sqlTran.Complete()

                pub_Param = Session("TampungDelivery")


                'pDeliverydate = txtdeliverydate.Text
                pDeliverydate = dt1.Text
                pAffiliateCode = Split(pub_Param, "|")(1)
                pAffiliateName = Split(pub_Param, "|")(2)
                pSuratjalanNo = txtsuratjalanno.Text.Trim
                pForwaderCode = Split(pub_Param, "|")(4)
                pForwaderName = Split(pub_Param, "|")(5)
                pDriverName = txtdrivername.Text.Trim

                pDriverContact = txtdrivercontact.Text.Trim
                pNoPol = txtnopol.Text.Trim
                pJenisArmada = txtjenisarmada.Text.Trim
                pPO = Split(pub_Param, "|")(10)
                pPallet = Split(pub_Param, "|")(11)
                pSupplier = Split(pub_Param, "|")(12)
                pSupplierName = Split(pub_Param, "|")(13)
                pSuratJalan = txtsuratjalanno.Text.Trim
                pRSJ = txtsuratjalanno.Text.Trim
                pSuppSuratJalan = Trim(Split(pub_Param, "|")(16))
                pFilter = Trim(Split(pub_Param, "|")(17))


                Session("TampungDelivery") = pDeliverydate & "|" & pAffiliateCode & "|" & pAffiliateName & "|" & pSuratjalanNo & "|" & pForwaderCode _
                                            & "|" & pForwaderName & "|" & pDriverName & "|" & pDriverContact & "|" & pNoPol & "|" & pJenisArmada & "|" & pPO _
                                            & "|" & pPallet & "|" & pSupplier & "|" & pSupplierName & "|" & pSuratJalan & "|" & pRSJ & "|" & pSuppSuratJalan & "|" & pFilter & "|" & Session("statusinsert")
            End Using
            sqlConn.Close()
            Call up_GridLoad(Session("POList2"), Session("pRSJ"), Session("pRSJ"), Session("pSuppSJ"), Session("PalletList"), Session("pFilter"), Session("statusinsert"))
        End Using

        Session.Remove("TampungDelivery")
    End Sub

    Private Sub Grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles Grid.CustomCallback
        Dim pAction As String = Split(e.Parameters, "|")(0)
        Select Case pAction
            Case "gridload"

                If Session("MsgDelivery") = "SAVED" Then
                    Call up_UPDATEMASTER()

                    If Not IsNothing(Session("SJPASI")) Then Session("pRSJ") = Session("SJPASI")
                    Call up_GridLoad(Session("POList2"), Session("pRSJ"), Session("pRSJ"), Session("pSuppSJ"), Session("PalletList"), Session("pFilter"), Session("statusinsert"))
                    Call up_IsiInvoice(Session("pRSJ"))
                    txttotalbox.Text = uf_SumQty(Session("POList2"), Session("PalletList"), Session("pRSJ"), Session("pFilter"), "INSERT")
                    Grid.JSProperties("cptotalbox") = txttotalbox.Text

                    'Call sendEmailtoAffiliate()
                    Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                    Grid.JSProperties("cpMessage") = lblerrmessage.Text
                Else
                    Session.Remove("SJPASI")
                    If Session("MsgDelivery") = "SAVED" Then
                        'Call sendEmailtoAffiliate()
                        Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    ElseIf Session("MsgDelivery") = "QTY" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "7014", clsMessage.MsgType.ErrorMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    ElseIf Session("MsgDelivery") = "NOT SAVED" Then
                        Call clsMsg.DisplayMessage(lblerrmessage, "6018", clsMessage.MsgType.ErrorMessage)
                        Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    Else
                            Call up_UPDATEMASTER()
                            Call clsMsg.DisplayMessage(lblerrmessage, "1002", clsMessage.MsgType.InformationMessage)
                            Grid.JSProperties("cpMessage") = lblerrmessage.Text
                    End If
                End If
                Session.Remove("MsgDelivery")

            Case "Delete"
                Call up_Delete(txtsuratjalanno.Text)
                Call up_GridLoad(Session("POList2"), Session("pRSJ"), Session("pRSJ"), Session("pSuppSJ"), Session("PalletList"), Session("pFilter"), "Delivery")
                txttotalbox.Text = 0 'uf_SumQty(Session("PO"), Session("Pallet"), Session("pRSJ"), Session("pRSJ"), Session("pSuppSJ"), Session("pFilter"), Session("statusinsert"))

                Call clsMsg.DisplayMessage(lblerrmessage, "1003", clsMessage.MsgType.InformationMessage)
                Grid.JSProperties("cpMessage") = lblerrmessage.Text
        End Select
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, False, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False, False, False, True)
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Session("SJAffiliate") = txtsuratjalanno.Text
        Session("RPTAffiliateID") = txtaffiliatecode.Text
        Response.Redirect("~/DeliveryExport/viewDeliveryToFor.aspx")
    End Sub

    Private Sub Grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles Grid.HtmlDataCellPrepared
        If Not (e.DataColumn.FieldName = "colpasideliveryqty" Or e.DataColumn.FieldName = "colno") Then
            e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
        Else
            e.Cell.BackColor = Color.White
        End If

        If e.DataColumn.FieldName = "colremainingdelqty" Then
            If (CDbl(e.GetValue("colpasigoodrec")) > e.GetValue("colpasideliveryqty")) Then
                e.Cell.BackColor = Color.Fuchsia
            End If
        End If

        'Delivery Qty Not save
        If e.DataColumn.FieldName = "colpasideliveryqty" Then
            If (Trim(e.GetValue("colstsDO")) = "") Then
                e.Cell.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub Grid_HtmlRowPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableRowEventArgs) Handles Grid.HtmlRowPrepared
        Try
            Dim getRowValues As String = e.GetValue("colpasideliveryqty")
            If Not IsNothing(getRowValues) Then
                If getRowValues.Trim() <> "" Then
                    e.Row.BackColor = Color.FromName("#E0E0E0")
                End If
            End If

        Catch ex As Exception
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class