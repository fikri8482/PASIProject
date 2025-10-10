Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxEditors
Imports System.Web.UI
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxClasses
Imports System.Drawing
Imports DevExpress.Web.ASPxMenu
Imports OfficeOpenXml
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Net

Public Class ShippingInstructionToForwarder
    Inherits System.Web.UI.Page
    Private processAddNewRow As Boolean

#Region "DECLARATION"
    Dim IsNewLabel As Boolean = True
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim paramDT1 As Date
    Dim paramDT2 As Date
    Dim paramSupplier As String
    Dim paramLocation As String

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "M00"

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
    Dim dtHeader As DataTable
    Dim dtDetail As DataTable

    Dim pCboCreateUpdate As String
    Dim pETDPort As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pForwarderCode As String
    Dim pForwarderName As String
    Dim pInstructionNo As String
    Dim pInstructionDate As String
    Dim pBLNo As String
    Dim pBLDate As String
    Dim pSupplierCode As String
    Dim pSupplierName As String
    Dim pPartCode As String
    Dim pPartName As String
    Dim pPONo As String
    Dim pStatusShipping As String
    Dim errMsg As String
#End Region

#Region "CONTROL EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ls_GenerateNo As String = ""
        Dim pShipCode As String = "B"

        Try
            Dim param As String = ""
            If (Not IsPostBack) AndAlso (Not IsCallback) Then

                'uf_ButtonSendEDI()

                Session("MenuDesc") = "SHIPPING INSTRUCTION TO FORWARDER"
                If IsNothing(Request.QueryString("prm")) Then
                    param = ""
                Else
                    param = Request.QueryString("prm").ToString()
                End If

                If param = "" Then
                    If Session("SHAFFILIATEID") <> "" Then
                        Call up_fillcombo()
                        Call up_fillcombocreateupdate()
                        'Call up_fillcombopackinglist()
                        Call ColorGrid()
                        etdport.Value = Now
                        etdvendor.Value = Now
                        etafactory.Value = Now
                        etaport.Value = Now
                        dtShippingDate.Value = Now
                        txtSupplierName.Text = "==ALL=="
                        txtPartName.Text = "==ALL=="
                        lblErrMsg.Text = ""
                        txtmeasurement.Text = 0
                        txtTotalPallet.Text = 0
                        txtGrossWeight.Text = 0
                        cboAffiliateCode.Text = Session("SHAFFILIATEID")
                        'txtOrderNo.Text = Session("SHORDERNO")
                        'cboSupplierCode.Text = Session("SHSUPPLIERID")
                        cboForwarder.Text = Session("SHFWD")
                        etdport.Value = Now
                        etdvendor.Value = Now
                        etafactory.Value = Now
                        etaport.Value = Now
                        dtShippingDate.Value = Now
                        cboPartNo.Text = "==ALL=="
                        txtPartName.Text = "==ALL=="
                        txtBLNo.Text = ""
                        dtBLDate.Value = Now
                        btnsubmenu.Text = "BACK"
                        cboCreate.Text = "CREATE"
                        Call ColorGrid()
                        Call up_GridLoad()
                        'Call up_GridLoadCreate()

                        ls_GenerateNo = CreateInvoiceNo(Session("SHAFFILIATEID"), pShipCode)
                        cboShippingNo.Text = ls_GenerateNo

                        'cboAffiliateCode.Enabled = False
                        'cboForwarder.Enabled = False
                    Else
                        Call up_fillcombo()
                        Call up_fillcombocreateupdate()
                        'Call up_fillcombopackinglist()
                        Call ColorGrid()
                        etdport.Value = Now
                        etdvendor.Value = Now
                        etafactory.Value = Now
                        etaport.Value = Now
                        dtShippingDate.Value = Now
                        txtSupplierName.Text = "==ALL=="
                        txtPartName.Text = "==ALL=="
                        txtBLNo.Text = ""
                        dtBLDate.Value = Now
                        lblErrMsg.Text = ""

                        'cboAffiliateCode.Enabled = True
                        'cboForwarder.Enabled = True
                    End If
                ElseIf param <> "" And Session("GOTOStatus") = "6" Then
                    lblErrMsg.Text = ""
                    pCboCreateUpdate = Split(param, "|")(1)
                    pETDPort = Split(param, "|")(2)
                    pAffiliateCode = Split(param, "|")(3)
                    pAffiliateName = Split(param, "|")(4)
                    pForwarderCode = Split(param, "|")(5)
                    pForwarderName = Split(param, "|")(6)
                    pInstructionNo = Split(param, "|")(7)
                    pInstructionDate = Split(param, "|")(8)
                    pSupplierCode = Split(param, "|")(9)
                    pSupplierName = Split(param, "|")(10)
                    pPartCode = Split(param, "|")(11)
                    pPartName = Split(param, "|")(12)
                    pStatusShipping = Split(param, "|")(13)
                    pPONo = Split(param, "|")(14)
                    pBLNo = Split(param, "|")(15)
                    pBLDate = Split(param, "|")(16)

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"

                    cboCreate.Text = pCboCreateUpdate
                    cboAffiliateCode.Text = pAffiliateCode
                    txtAffiliateName.Text = pAffiliateName
                    cboForwarder.Text = pForwarderCode
                    txtForwarder.Text = pForwarderName
                    cboShippingNo.Text = pInstructionNo
                    dtShippingDate.Text = pInstructionDate
                    txtBLNo.Text = pBLNo
                    dtBLDate.Text = pBLDate

                    If pSupplierCode = "" Then
                        cboSupplierCode.Text = "==ALL=="
                    Else
                        cboSupplierCode.Text = pSupplierCode
                    End If

                    If pSupplierName = "" Then
                        txtSupplierName.Text = "==ALL=="
                    Else
                        txtSupplierName.Text = pSupplierName
                    End If

                    If pPartCode = "" Then
                        cboPartNo.Text = "==ALL=="
                    Else
                        cboPartNo.Text = pPartCode
                    End If

                    If pPartName = "" Then
                        txtPartName.Text = "==ALL=="
                    Else
                        txtPartName.Text = pPartName
                    End If

                    txtOrderNo.Text = pPONo
                    txtSend.Text = pStatusShipping

                    Call up_GridLoadUpdate()
                    'Call up_GridLoad()
                    up_IsiInvoice(Trim(cboShippingNo.Text))

                    Session.Remove("EmergencyUrl")
                    btnsubmenu.Text = "BACK"

                ElseIf param <> "" And Session("GOTOStatus") = "enam" Then
                    lblErrMsg.Text = ""
                    pCboCreateUpdate = Split(param, "|")(1)
                    pETDPort = Split(param, "|")(2)
                    pAffiliateCode = Split(param, "|")(3)
                    pAffiliateName = Split(param, "|")(4)
                    pForwarderCode = Split(param, "|")(5)
                    pForwarderName = Split(param, "|")(6)
                    pInstructionNo = Split(param, "|")(7)
                    pInstructionDate = Split(param, "|")(8)
                    pSupplierCode = Split(param, "|")(9)
                    pSupplierName = Split(param, "|")(10)
                    pPartCode = Split(param, "|")(11)
                    pPartName = Split(param, "|")(12)
                    pStatusShipping = Split(param, "|")(13)
                    pPONo = Split(param, "|")(14)
                    pBLNo = Split(param, "|")(15)
                    pBLDate = Split(param, "|")(16)

                    If pAffiliateCode <> "" Then btnsubmenu.Text = "BACK"

                    cboCreate.Text = pCboCreateUpdate
                    cboAffiliateCode.Text = pAffiliateCode
                    txtAffiliateName.Text = pAffiliateName
                    cboForwarder.Text = pForwarderCode
                    txtForwarder.Text = pForwarderName
                    cboShippingNo.Text = pInstructionNo
                    dtShippingDate.Text = pInstructionDate
                    txtBLNo.Text = pBLNo
                    dtBLDate.Text = pBLDate

                    If pSupplierCode = "" Then
                        cboSupplierCode.Text = "==ALL=="
                    Else
                        cboSupplierCode.Text = pSupplierCode
                    End If

                    If pSupplierName = "" Then
                        txtSupplierName.Text = "==ALL=="
                    Else
                        txtSupplierName.Text = pSupplierName
                    End If

                    If pPartCode = "" Then
                        cboPartNo.Text = "==ALL=="
                    Else
                        cboPartNo.Text = pPartCode
                    End If

                    If pPartName = "" Then
                        txtPartName.Text = "==ALL=="
                    Else
                        txtPartName.Text = pPartName
                    End If

                    txtOrderNo.Text = pPONo
                    txtSend.Text = pStatusShipping

                    Call up_GridLoadUpdate()
                    'Call up_GridLoad()
                    up_IsiInvoice(Trim(cboShippingNo.Text))

                    Session.Remove("EmergencyUrl")
                    btnsubmenu.Text = "BACK"
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Protected Sub btnPrintTally_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPrintTally.Click
        Dim ls_sql As String = ""
        Session.Remove("REPORT")
        Session.Remove("Query")

        'ls_sql = " select distinct  " & vbCrLf & _
        '          " 	ContainerNo = TM.ContainerNo,  " & vbCrLf & _
        '          " 	SealNo = TM.SealNo,  " & vbCrLf & _
        '          " 	Tare = TM.Tare,  " & vbCrLf & _
        '          " 	Gross = TM.Gross,  " & vbCrLf & _
        '          " 	InvoiceNo = TM.ShippingInstructionNo,  " & vbCrLf & _
        '          " 	PalletNo = TD.PalletNo,  " & vbCrLf & _
        '          " 	OrderNo = TD.OrderNo,  " & vbCrLf & _
        '          " 	PartNo = TD.PartNo,  " & vbCrLf & _
        '          " 	CaseNo = Rtrim(TD.CaseNo) + '-' + Rtrim(TD.CaseNo2)," & vbCrLf & _
        '          "     jmlCTN = TD1.totalBox, " & vbCrLf & _
        '          " 	Length = (SUMTally.Length),  " & vbCrLf

        'ls_sql = ls_sql + "     Width = (SUMTally.Width),  " & vbCrLf & _
        '                  "     Height = (SUMTally.Height),  " & vbCrLf & _
        '                  "     M3 = (SUMTally.M3),  " & vbCrLf & _
        '                  "     WGT = (SUMTally.WeightPallet)  " & vbCrLf & _
        '                  "     From Tally_master TM   " & vbCrLf & _
        '                  "     LEFT JOIN Tally_Detail TD  " & vbCrLf & _
        '                  "     ON TM.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
        '                  "     AND TM.ForwarderID = TD.ForwarderID  " & vbCrLf & _
        '                  "     AND TM.AffiliateID = TD.AffiliateID  " & vbCrLf & _
        '                  "     LEFT JOIN (select ShippingInstructionNo, ForwarderID, AffiliateID, SUM(TotalBox) as totalBox from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID)TD1 " & vbCrLf

        'ls_sql = ls_sql + " 	ON TD1.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
        '                  " 		AND TD1.ForwarderID = TD.ForwarderID  " & vbCrLf & _
        '                  " 		AND TD1.AffiliateID = TD.AffiliateID  " & vbCrLf & _
        '                  " 		--AND TD1.PalletNo = TD.PalletNo  " & vbCrLf & _
        '                  " 		--AND TD1.OrderNo = TD.OrderNo  " & vbCrLf & _
        '                  " 		--AND TD1.PartNO = TD.PartNo  " & vbCrLf & _
        '                  " 	LEFT JOIN (select ShippingInstructionNo, ForwarderID, AffiliateID, PalletNo, OrderNo,PartNo,  " & vbCrLf & _
        '                  " 				Max(CaseNo) as CaseNo1 from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID, PalletNo, OrderNo,PartNo)TD2  " & vbCrLf & _
        '                  " 	ON TD2.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
        '                  " 		AND TD2.ForwarderID = TD.ForwarderID  " & vbCrLf & _
        '                  " 		AND TD2.AffiliateID = TD.AffiliateID  " & vbCrLf

        'ls_sql = ls_sql + " 		AND TD2.PalletNo = TD.PalletNo  " & vbCrLf & _
        '                  " 		AND TD2.OrderNo = TD.OrderNo  " & vbCrLf & _
        '                  " 		AND TD2.PartNO = TD.PartNo  " & vbCrLf & _
        '                  "     LEFT JOIN (select JMLCTN = Sum(JMLCTN),ShippingInstructionNo, ForwarderID, AffiliateID From( " & vbCrLf & _
        '                  "                     select ShippingInstructionNo, ForwarderID, AffiliateID,  " & vbCrLf & _
        '                  "                         Count(CaseNo) as JMLCTN from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID) x " & vbCrLf & _
        '                  "                     group by ShippingInstructionNo, ForwarderID, AffiliateID)TD3  " & vbCrLf & _
        '                  "         ON TD3.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
        '                  "         AND TD3.ForwarderID = TD.ForwarderID " & vbCrLf & _
        '                  "         AND TD3.AffiliateID = TD.AffiliateID " & vbCrLf & _
        '                  " 	LEFT JOIN (select distinct ShippingInstructionNo, ForwarderID, AffiliateID, OrderNo, palletno, weightpallet,  " & vbCrLf & _
        '                  " 				width, height, length, M3 from Tally_Detail) SUMTally " & vbCrLf & _
        '                  " 	ON SUMTally.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
        '                  " 		AND SUMTally.ForwarderID = TD.ForwarderID  " & vbCrLf & _
        '                  " 		AND SUMTally.AffiliateID = TD.AffiliateID  " & vbCrLf & _
        '                  " 		AND SUMTally.PalletNo = TD.PalletNo  " & vbCrLf & _
        '                  " 		AND SUMTally.OrderNo = TD.OrderNo  " & vbCrLf

        'ls_sql = ls_sql + "  " & vbCrLf & _
        '                  " WHERE TM.ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' " & vbCrLf & _
        '                  " AND TM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "'" & vbCrLf & _
        '                  " AND TM.ForwarderID = '" & Trim(cboForwarder.Text) & "' "

        ls_sql = " select distinct  " & vbCrLf & _
                  " 	Company.Adress1 + ' ' + Company.City1 + ISNULL('. Phone : ' + Company.Phone1,'') + ISNULL(' FAX : ' + Company.Fax1,'') AS Adress1,  " & vbCrLf & _
                  " 	Company.Adress2 + ' ' + Company.City2 + ISNULL('. Phone : ' + Company.Phone2,'') + ISNULL(' FAX : ' + Company.Fax2,'') AS Adress2,   " & vbCrLf & _
                  " 	ContainerNo = TM.ContainerNo,  " & vbCrLf & _
                  " 	SealNo = TM.SealNo,  " & vbCrLf & _
                  " 	Tare = TM.Tare,  " & vbCrLf & _
                  " 	Gross = TM.Gross,  " & vbCrLf & _
                  " 	InvoiceNo = TM.ShippingInstructionNo,  " & vbCrLf & _
                  " 	PalletNo = TD.PalletNo,  " & vbCrLf & _
                  " 	OrderNo = TD.OrderNo,  " & vbCrLf & _
                  " 	PartNo = TD.PartNo,  partCust = isnull(PartGroupName,'')," & vbCrLf & _
                  " 	CaseNo = Rtrim(TD.CaseNo) + CASE WHEN Rtrim(TD.CaseNo2) = '' then '' else + '-' + Rtrim(TD.CaseNo2) END," & vbCrLf & _
                  "     jmlCTN = TD1.totalBox, " & vbCrLf & _
                  " 	Length = (SUMTally.Length),  " & vbCrLf

        ls_sql = ls_sql + "     Width = (SUMTally.Width),  " & vbCrLf & _
                          "     Height = (SUMTally.Height),  " & vbCrLf & _
                          "     M3 = (SUMTally.M3),  " & vbCrLf & _
                          "     WGT = (SUMTally.WeightPallet),  " & vbCrLf & _
                          "     QtyCarton = TD.TotalBox, " & vbCrLf & _
                          "     QtyPart = TD.TotalBox * SD.QtyBox" & vbCrLf & _
                          "     From Tally_master TM   " & vbCrLf & _
                          "     LEFT JOIN Tally_Detail TD  " & vbCrLf & _
                          "     ON TM.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          "     AND TM.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          "     AND TM.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                          "     LEFT JOIN MS_Parts MSP ON MSP.PartNo = TD.PartNo " & vbCrLf & _
                          "     LEFT JOIN (select ShippingInstructionNo, ForwarderID, AffiliateID, SUM(TotalBox) as totalBox from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID)TD1 " & vbCrLf

        ls_sql = ls_sql + " 	ON TD1.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          " 		AND TD1.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          " 		AND TD1.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                          " 		--AND TD1.PalletNo = TD.PalletNo  " & vbCrLf & _
                          " 		--AND TD1.OrderNo = TD.OrderNo  " & vbCrLf & _
                          " 		--AND TD1.PartNO = TD.PartNo  " & vbCrLf & _
                          " 	LEFT JOIN (select ShippingInstructionNo, ForwarderID, AffiliateID, PalletNo, OrderNo,PartNo,  " & vbCrLf & _
                          " 				Max(CaseNo) as CaseNo1 from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID, PalletNo, OrderNo,PartNo)TD2  " & vbCrLf & _
                          " 	ON TD2.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          " 		AND TD2.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          " 		AND TD2.AffiliateID = TD.AffiliateID  " & vbCrLf

        ls_sql = ls_sql + " 		AND TD2.PalletNo = TD.PalletNo  " & vbCrLf & _
                          " 		AND TD2.OrderNo = TD.OrderNo  " & vbCrLf & _
                          " 		AND TD2.PartNO = TD.PartNo  " & vbCrLf & _
                          "     LEFT JOIN (select JMLCTN = Sum(JMLCTN),ShippingInstructionNo, ForwarderID, AffiliateID From( " & vbCrLf & _
                          "                     select ShippingInstructionNo, ForwarderID, AffiliateID,  " & vbCrLf & _
                          "                         Count(CaseNo) as JMLCTN from Tally_Detail group by ShippingInstructionNo, ForwarderID, AffiliateID) x " & vbCrLf & _
                          "                     group by ShippingInstructionNo, ForwarderID, AffiliateID)TD3  " & vbCrLf & _
                          "         ON TD3.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          "         AND TD3.ForwarderID = TD.ForwarderID " & vbCrLf & _
                          "         AND TD3.AffiliateID = TD.AffiliateID " & vbCrLf & _
                          " 	LEFT JOIN (select distinct ShippingInstructionNo, ForwarderID, AffiliateID, OrderNo, palletno, weightpallet,  " & vbCrLf & _
                          " 				width, height, length, M3 from Tally_Detail) SUMTally " & vbCrLf & _
                          " 	ON SUMTally.ShippingInstructionNo = TD.ShippingInstructionNo  " & vbCrLf & _
                          " 		AND SUMTally.ForwarderID = TD.ForwarderID  " & vbCrLf & _
                          " 		AND SUMTally.AffiliateID = TD.AffiliateID  " & vbCrLf & _
                          " 		AND SUMTally.PalletNo = TD.PalletNo  " & vbCrLf & _
                          " 		AND SUMTally.OrderNo = TD.OrderNo  " & vbCrLf & _
                          "     Left Join ShippingInstruction_Detail SD ON TD.PartNo = SD.PartNo and TD.ShippingInstructionNo =  SD.ShippingInstructionNo and TD.ForwarderID = SD.ForwarderID AND TD.AffiliateID = SD.AffiliateID " & vbCrLf & _
                          " 		OUTER APPLY (SELECT TOP 1 * FROM dbo.CompanyProfile WHERE ActiveDate < '" + dtShippingDate.Text + "' ORDER BY ActiveDate DESC) Company      " & vbCrLf

        ls_sql = ls_sql + "  " & vbCrLf & _
                          " WHERE TM.ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' " & vbCrLf & _
                          " AND TM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "'" & vbCrLf & _
                          " AND TM.ForwarderID = '" & Trim(cboForwarder.Text) & "' "

        Session("REPORT") = "TALLY"
        Session("Query") = ls_sql
        Response.Redirect("~/ShippingInstruction/ShippingViewReportExportCR.aspx")
    End Sub

    Protected Sub btnPrintSI_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPrintSI.Click
        Dim ls_sql As String = ""
        Session.Remove("REPORT")
        Session.Remove("Query")

        ls_sql = "  select       " & vbCrLf & _
                  "     Adress1 = (SELECT TOP 1 Company.Adress1 + ' ' + Company.City1 + ISNULL('. Phone : ' + Company.Phone1, '') + ISNULL(' FAX : ' + Company.Fax1, '') FROM dbo.CompanyProfile company WHERE ActiveDate < '" + dtShippingDate.Text + "' ORDER BY ActiveDate DESC), " & vbCrLf & _
                  "     Adress2 = (SELECT TOP 1 Company.Adress2 + ' ' + Company.City2 + ISNULL('. Phone : ' + Company.Phone2, '') + ISNULL(' FAX : ' + Company.Fax2, '') FROM dbo.CompanyProfile company WHERE ActiveDate < '" + dtShippingDate.Text + "' ORDER BY ActiveDate DESC)," & vbCrLf & _
                  "  ShippingInstructionNo = SIM.ShippingInstructionNo,           " & vbCrLf & _
                  "  FWD = Rtrim(MF.ForwarderName) + ' ' + Rtrim(MF.Address) + ' ' + Rtrim(MF.City) + ' ' + Rtrim(MF.PostalCode),          " & vbCrLf & _
                  "  ATT = isnull(Rtrim(MF.Attn),''),           " & vbCrLf & _
                  "  FAx = isnull(Rtrim(MF.Fax),''),           " & vbCrLf & _
                  "  Tujuan = isnull(TM.DestinationPort,''),           " & vbCrLf & _
                  "  Shipment = Case when TypeOfService = 'FCL' then 'SEA FREIGHT' WHEN TypeOfService = 'LCL' then 'SEA FREIGHT' ELSE 'AIR FREIGHT' END,           " & vbCrLf & _
                  "  Vessel = Vessel,           " & vbCrLf & _
                  "  ETD = Convert(Char(12), convert(Datetime, isnull(SIM.ETDPort,POM.ETDPort1)),106),           " & vbCrLf & _
                  "  ETA = Convert(Char(12), convert(Datetime, isnull(SIM.ETAPort,POM.ETAPort1)),106),           " & vbCrLf & _
                  "  tgltiba = Convert(Char(12), convert(Datetime, isnull(SIM.ETAPort,POM.ETAPort1)),106),          "

        ls_sql = ls_sql + "  part = 'Automotive Component',           " & vbCrLf & _
                          "  jumlah = 0,           " & vbCrLf & _
                          "  pallet = isnull(SUMTally.palletno,0),      " & vbCrLf & _
                          "  Box = isnull(Sumtally.box,0), " & vbCrLf & _
                          "  Qty = SUM(isnull(SD.ShippingQty,0)),          " & vbCrLf & _
                          "  beratBersih = SUM(((netweight/ISNULL(SD.POQtyBox,MPM.QtyBox))* SD.ShippingQty)/1000),           " & vbCrLf & _
                          "  beratKotor = SUM(((grossweight/ISNULL(SD.POQtyBox,MPM.QtyBox))* SD.ShippingQty)/1000),           " & vbCrLf & _
                          "  Buyer = Rtrim(BuyerName), BuyerAddress = Rtrim(BuyerAddress),           " & vbCrLf & _
                          "  Consignee = Rtrim(MA.ConsigneeName), ConsigneeAddress = Rtrim(MA.ConsigneeAddress), Attn = isnull(MSA.AffiliatePOTo,''), " & vbCrLf & _
                          "  Freight = isnull(Freight,''),           " & vbCrLf & _
                          "  Stuffing = Convert(Char(12), convert(Datetime, isnull(TM.Stuffingdate,'')),106)          "

        ls_sql = ls_sql + "  From ShippingInstruction_master SIM           " & vbCrLf & _
                          "  LEFT JOIN ShippingInstruction_Detail SD           " & vbCrLf & _
                          "  ON SIM.ShippingInstructionNo = SD.ShippingInstructionNo           " & vbCrLf & _
                          "  	AND SIM.AffiliateID = SD.AffiliateID         " & vbCrLf & _
                          "  	AND SIM.ForwarderID = SD.ForwarderID           " & vbCrLf

        ls_sql = ls_sql + "  LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SIM.AffiliateID           " & vbCrLf & _
                          "  LEFT JOIN ms_emailAffiliate_Export MSA ON MSA.AffiliateID = MA.AffiliateID " & vbCrLf & _
                          "  LEFT JOIN MS_Parts MP ON MP.PartNo = SD.PartNo           " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SD.PartNo           " & vbCrLf & _
                          "  	AND MPM.AffiliateID = SD.AffiliateID           " & vbCrLf & _
                          "  	AND MPM.SupplierID = SD.SupplierID           " & vbCrLf & _
                          "  LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SIM.ForwarderID      " & vbCrLf & _
                          "  LEFT JOIN Tally_Master TM ON TM.ShippingInstructionNo = SIM.ShippingInstructionNo      "

        ls_sql = ls_sql + "  	AND TM.AffiliateID = SIM.AffiliateID      " & vbCrLf & _
                          "  	AND TM.ForwarderID = SIM.ForwarderID      " & vbCrLf & _
                          "  LEFT JOIN (Select distinct ShippingInstructionNo, ForwarderID, AffiliateID, palletno = Count(palletno), weightpallet = sum(weightpallet),      " & vbCrLf & _
                          "  			width = sum(width), height = sum(height), length = sum(length), M3 = Sum(M3), box = SUM(box)   " & vbCrLf & _
                          "  		   from (     " & vbCrLf & _
                          "  					select ShippingInstructionNo, ForwarderID, AffiliateID, palletno, weightpallet = sum(weightpallet),      " & vbCrLf & _
                          "  					width = sum(width), height = sum(height), length = sum(length), M3 = SUM(M3), box =SUM(TOTALBOX) from Tally_Detail     					 " & vbCrLf & _
                          "  					group by ShippingInstructionNo, ForwarderID, AffiliateID, palletno   " & vbCrLf & _
                          "  				) x group by ShippingInstructionNo, ForwarderID, AffiliateID) SUMTally     " & vbCrLf & _
                          "  ON SUMTally.ShippingInstructionNo = SD.ShippingInstructionNo      " & vbCrLf & _
                          "  AND SUMTally.ForwarderID = SD.ForwarderID " & vbCrLf

        ls_sql = ls_sql + "  AND SUMTally.AffiliateID = SD.AffiliateID " & vbCrLf & _
                          "  " & vbCrLf & _
                          " WHERE TM.ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' " & vbCrLf & _
                          " AND TM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "'" & vbCrLf & _
                          " AND TM.ForwarderID = '" & Trim(cboForwarder.Text) & "' " & vbCrLf & _
                          "  GROUP BY      " & vbCrLf & _
                          "  SIM.ShippingInstructionNo,          " & vbCrLf & _
                          "  Rtrim(MF.ForwarderName) ,Rtrim(MF.Address) , Rtrim(MF.City) , Rtrim(MF.PostalCode),         " & vbCrLf & _
                          "  isnull(Rtrim(MF.Attn),''),  " & vbCrLf & _
                          "  isnull(Rtrim(MF.Fax),''),          " & vbCrLf & _
                          "  isnull(TM.DestinationPort,''), " & vbCrLf & _
                          "  POM.ETDPort1, SIM.ETDPort, SIM.ETAPort, TM.Stuffingdate,MSA.AffiliatePOTo,        "

        ls_sql = ls_sql + "  POM.ETAPort1,              " & vbCrLf & _
                          "  vessel,     " & vbCrLf & _
                          "  Rtrim(BuyerName),Rtrim(BuyerAddress),Rtrim(MA.ConsigneeName), Rtrim(MA.ConsigneeAddress) ,     " & vbCrLf & _
                          "  ShipCls, SUMTally.palletno,Sumtally.box, Freight  "

        ls_sql = " SELECT DISTINCT  " & vbCrLf & _
                  "     Adress1 = (SELECT TOP 1 Company.Adress1 + ' ' + Company.City1 + ISNULL('. Phone : ' + Company.Phone1, '') + ISNULL(' FAX : ' + Company.Fax1, '') FROM dbo.CompanyProfile company WHERE ActiveDate < '" + dtShippingDate.Text + "' ORDER BY ActiveDate DESC), " & vbCrLf & _
                  "     Adress2 = (SELECT TOP 1 Company.Adress2 + ' ' + Company.City2 + ISNULL('. Phone : ' + Company.Phone2, '') + ISNULL(' FAX : ' + Company.Fax2, '') FROM dbo.CompanyProfile company WHERE ActiveDate < '" + dtShippingDate.Text + "' ORDER BY ActiveDate DESC)," & vbCrLf & _
                  "     ShippingInstructionNo = SIM.ShippingInstructionNo,             " & vbCrLf & _
                  "     FWD = Rtrim(MF.ForwarderName) + ' ' + Rtrim(MF.Address) + ' ' + Rtrim(MF.City) + ' ' + Rtrim(MF.PostalCode),            " & vbCrLf & _
                  "     ATT = isnull(Rtrim(MF.Attn),''),             " & vbCrLf & _
                  "     FAx = isnull(Rtrim(MF.Fax),''),             " & vbCrLf & _
                  "     Tujuan = isnull(MA.DestinationPort,''),  " & vbCrLf & _
                  "     Shipment = Case when TypeOfService = 'FCL' then 'SEA FREIGHT' WHEN TypeOfService = 'LCL' then 'SEA FREIGHT' ELSE 'AIR FREIGHT' END,   " & vbCrLf & _
                  "     Vessel = ISNULL(SIM.Vessels,''),  " & vbCrLf & _
                  "     ETD = Convert(Char(12), convert(Datetime, SIM.ETDPort),106),             " & vbCrLf & _
                  "     ETA = Convert(Char(12), convert(Datetime, SIM.ETAPort),106),             " & vbCrLf & _
                  "     tgltiba = Convert(Char(12), convert(Datetime, SIM.ETAPort),106), "

        ls_sql = ls_sql + "     part = 'Automotive Component',             " & vbCrLf & _
                          "     jumlah = 0,  " & vbCrLf & _
                          "     pallet = isnull(SIM.TotalPallet,0),  " & vbCrLf & _
                          "     Box = isnull(SUM(SD.ShippingQty/ISNULL(SD.POQtyBox,MPM.QtyBox)),0),  " & vbCrLf & _
                          "     Qty = SUM(isnull(SD.ShippingQty,0)),  " & vbCrLf & _
                          "     beratBersih = SUM(((netweight/ISNULL(SD.POQtyBox,MPM.QtyBox))* SD.ShippingQty)/1000),             " & vbCrLf & _
                          "     beratKotor = ISNULl(SIM.GrossWeight,0),  " & vbCrLf & _
                          "     Buyer = Rtrim(BuyerName),             " & vbCrLf & _
                          "     BuyerAddress = Rtrim(BuyerAddress),   " & vbCrLf & _
                          "     Consignee = Rtrim(MA.ConsigneeName), ConsigneeAddress = Rtrim(MA.ConsigneeAddress), Attn =isnull(MA.Att,''),   " & vbCrLf & _
                          "     Freight = isnull(Freight,'') "

        ls_sql = ls_sql + "     From ShippingInstruction_master SIM             " & vbCrLf & _
                          "                     LEFT JOIN ShippingInstruction_Detail SD             " & vbCrLf & _
                          "                     ON SIM.ShippingInstructionNo = SD.ShippingInstructionNo             " & vbCrLf & _
                          "                     AND SIM.AffiliateID = SD.AffiliateID           " & vbCrLf & _
                          "                     AND SIM.ForwarderID = SD.ForwarderID  " & vbCrLf & _
                          "                     LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SIM.ForwarderID  " & vbCrLf & _
                          "                     LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SIM.AffiliateID  " & vbCrLf & _
                          "                     LEFT JOIN PO_Master_Export POM ON (POM.PONo = SD.OrderNo or POM.OrderNo1 = SD.OrderNo)  " & vbCrLf & _
                          "                         AND POM.AffiliateID = SD.AffiliateID             " & vbCrLf & _
                          "                         AND POM.SupplierID = SD.SupplierID             " & vbCrLf & _
                          "                     LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SD.PartNo "

        ls_sql = ls_sql + "          				and MPM.SupplierID = SD.SupplierID   " & vbCrLf & _
                          "                         and MPM.AffiliateID = SD.AffiliateID  " & vbCrLf & _
                          "                           WHERE SIM.ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'  " & vbCrLf & _
                          "                           AND SIM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf & _
                          "                           AND SIM.ForwarderID = '" & Trim(cboForwarder.Text) & "'  " & vbCrLf & _
                          "                           GROUP BY SIM.ShippingInstructionNo,  " & vbCrLf & _
                          "                           	Rtrim(MF.ForwarderName) ,Rtrim(MF.Address) , Rtrim(MF.City) , Rtrim(MF.PostalCode),   " & vbCrLf & _
                          "                           	isnull(Rtrim(MF.Attn),''),    " & vbCrLf & _
                          "                           	isnull(Rtrim(MF.Fax),''),            " & vbCrLf & _
                          "                           	isnull(MA.DestinationPort,''),  " & vbCrLf & _
                          "                           	typeofservice, SIM.VesselS, "

        ls_sql = ls_sql + "          					SIM.ETDPort, SIM.ETAPort, SIM.TotalPallet, SIM.GrossWeight,  " & vbCrLf & _
                          "                           	Rtrim(BuyerName),Rtrim(BuyerAddress),Rtrim(MA.ConsigneeName), Rtrim(MA.ConsigneeAddress), MA.Att, SIM.Freight                           " & vbCrLf & _
                          "  "

        Session("REPORT") = "SI"
        Session("Query") = ls_sql
        Response.Redirect("~/ShippingInstruction/ShippingViewReportExportCR.aspx")
    End Sub

    Protected Sub btnprintinvoice_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnprintinvoice.Click
        Dim ls_sql As String = ""
        Session.Remove("REPORT")
        Session.Remove("Query")

        Dim tentukanBoat As String = cboService.Text.Trim
        Dim tentukanTerm As String = cboterm.Text.Trim

        Dim PriceCls As String = 0

        ''If tentukanBoat = "FCL" Or tentukanBoat = "LCL" Then
        'If tentukanTerm = "FCA" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
        '    PriceCls = "2"
        'ElseIf tentukanTerm = "FCA" Then
        '    PriceCls = "1"
        'End If

        'If tentukanTerm = "CIF" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
        '    PriceCls = "4"
        'ElseIf tentukanTerm = "CIF" Then
        '    PriceCls = "3"
        'End If

        'If tentukanTerm = "DDU PASI" Then
        '    PriceCls = "5"
        'ElseIf tentukanTerm = "DDU Affiliate" Then
        '    PriceCls = "6"
        'ElseIf tentukanTerm = "EX-Work" Then
        '    PriceCls = "7"
        'ElseIf tentukanTerm = "FOB" Then
        '    PriceCls = "8"
        'End If

        PriceCls = uf_PriceCls(tentukanTerm)

        ls_sql = "  select distinct  " & vbCrLf & _
                 "  Company.Adress1 + ' ' + Company.City1 + ISNULL('. Phone : ' + Company.Phone1,'') + ISNULL(' FAX : ' + Company.Fax1,'') AS Adress1, " & vbCrLf & _
                 "  Company.Adress2 + ' ' + Company.City2 + ISNULL('. Phone : ' + Company.Phone2,'') + ISNULL(' FAX : ' + Company.Fax2,'') AS Adress2, " & vbCrLf & _
                 "  buyer = Rtrim(MA.BuyerName) + CHAR(13)+CHAR(10) + Rtrim(MA.BuyerAddress),  " & vbCrLf & _
                 "  Consignee = Rtrim(Coalesce(MA.ConsigneeName, MA.AffiliateName)) + CHAR(13)+CHAR(10) + Rtrim(coalesce(MA.ConsigneeAddress, Rtrim(MA.Address) + Rtrim(MA.City) )),  " & vbCrLf & _
                 "  Attn = 'Radhika (Radhika@hesto.co.za)',  " & vbCrLf & _
                 "  Vessel = SHM.VesselS,--TM.Vessel,  " & vbCrLf & _
                 "  Fromto = 'JAKARTA, INDONESIA',  " & vbCrLf & _
                 "  Toto = isnull(MA.DestinationPort,''),  " & vbCrLf & _
                 "  About = Convert(Char(12), convert(Datetime, isnull(SHM.ETAPort,POM.ETAPort1)),106),  " & vbCrLf & _
                 "  ONAbout = Convert(Char(12), convert(Datetime, isnull(SHM.ETDPort,POM.ETDPort1)),106),  " & vbCrLf & _
                 "  Via = SHM.Via,  " & vbCrLf & _
                 "  InvoiceNo = SHM.ShippingInstructionNo,   "

        ls_sql = ls_sql + "  OrderNo = (SELECT (STUFF((SELECT distinct ', ' + RTrim(ShippingInstruction_Detail.orderNo) FROM ShippingInstruction_Detail WHERE ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' AND AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' FOR XML PATH('')), 1, 2, ''))),  " & vbCrLf & _
                          "  InvDate = Convert(Char(12), convert(Datetime, isnull(SHM.ShippingInstructionDate,'')),106),  " & vbCrLf & _
                          "  Place = 'JAKARTA',  " & vbCrLf & _
                          "  Privilege = '',  " & vbCrLf & _
                          "  AWB = '',  " & vbCrLf & _
                          "  ContainerNo = '', --TM.ContainerNo,  " & vbCrLf & _
                          "  Insurance = '',  " & vbCrLf & _
                          "  Remarks = SHM.Remarks,  " & vbCrLf & _
                          "  paymentTerm = CASE WHEN POM.CommercialCls = '1' Then Isnull(MA.PaymentTerm,'') ELSE 'NO COMMERCIAL VALUE' END,  " & vbCrLf & _
                          "  Marks = '',--Description = '',  " & vbCrLf & _
                          "  QtyBox = SHD.QtyBox,   " & vbCrLf

        ls_sql = ls_sql + "  Qty = RB.Box, --td2.Qty,  " & vbCrLf & _
                          "  Price = isnull(SHD.Price,MPR.Price) + '.0000' ,   " & vbCrLf & _
                          "  Amount = 0,  " & vbCrLf & _
                          "  Net =  (isnull(NetWeight,0)/1000),  " & vbCrLf & _
                          "  Gross =(isnull(SHM.GrossWeight,0)),  " & vbCrLf & _
                          "  DocNo = '',  " & vbCrLf & _
                          "  RevNo = '',  " & vbCrLf & _
                          "  partCust = isnull(PartGroupName,''),  " & vbCrLf & _
                          "  PartYazaki = SHD.PartNo,  " & vbCrLf & _
                          "  CaseNo = Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2),  " & vbCrLf & _
                          "  totalCarton = 0, " & vbCrLf & _
                          "  Term = CASE WHEN RTRIM(MPC.Description) = 'FCA - BOAT' THEN 'FCA' " & vbCrLf & _
                          " 			  WHEN RTRIM(MPC.Description) = 'FCA - AIR' THEN 'FCA'	 " & vbCrLf & _
                          " 			  WHEN RTRIM(MPC.Description) = 'CIF - BOAT' THEN 'CIF' " & vbCrLf & _
                          " 			  WHEN RTRIM(MPC.Description) = 'CIF - AIR' THEN 'CIF' " & vbCrLf & _
                          " 			  ELSE RTRIM(MPC.Description) END,	 " & vbCrLf & _
                          "  SHM.TotalPallet," & vbCrLf & _
                          "  SHM.Measurement, SHM.GrossWeight, SHM.Freight, CASE WHEN ISNULL(HSCodeCls,'0')  = '0'  THEN '' else HSCode END HSCode, POD.PONo NewOrderNo, SHM.TypeOfService,  SHM.NamaKapalS From  " & vbCrLf
        '"  SHM.Measurement, SHM.GrossWeight, SHM.Freight, CASE WHEN ISNULL(HSCodeCls,'0')  = '0'  THEN '' else HSCode END HSCode, SHD.OrderNo NewOrderNo, SHM.TypeOfService,  SHM.NamaKapalS From  " & vbCrLf

        ls_sql = ls_sql + "  ShippingInstruction_Detail SHD   " & vbCrLf & _
                          "  LEFT JOIN ShippingInstruction_Master SHM ON SHM.ShippingInstructionNo = SHD.ShippingInstructionNo  " & vbCrLf & _
                          "  AND SHM.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  AND SHM.ForwarderID = SHD.ForwarderID  " & vbCrLf & _
                          "  LEFT JOIN MS_Parts MP ON MP.PartNo = SHD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN MS_PartMapping MPM ON MPM.Partno = SHD.PartNo and MPM.AffiliateID = SHD.AffiliateID and MPM.SupplierID = SHD.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SHD.ForwarderID  LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_master RM ON RM.SuratJalanNo = SHD.SuratJalanno  AND RM.AffiliateID = SHD.AffiliateID  " & vbCrLf & _
                          "  AND RM.OrderNo = SHD.OrderNo  " & vbCrLf & _
                          "  AND SHD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNO = RM.SuratJalanno  " & vbCrLf

        ls_sql = ls_sql + "  AND RD.AffiliateID = RM.AffiliateID  " & vbCrLf & _
                          "  AND RD.SupplierID = RM.SupplierID  " & vbCrLf & _
                          "  AND RD.PONo = RM.PONO  " & vbCrLf & _
                          "  AND RD.OrderNO = Rm.OrderNo  " & vbCrLf & _
                          "  AND RD.PartNo = SHD.PartNo " & vbCrLf & _
                          "  LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = SHD.SuratJalanNo  " & vbCrLf & _
                          "  AND RB.SupplierID = SHD.SupplierID   " & vbCrLf & _
                          "  AND RB.AffiliateID = SHD.AffiliateID   " & vbCrLf & _
                          "  --AND RB.PONo = RD.PONo   " & vbCrLf & _
                          "  AND RB.OrderNo = SHD.OrderNo   " & vbCrLf & _
                          "  AND RB.PartNo = SHD.PartNo   " & vbCrLf

        ls_sql = ls_sql + "  AND RB.StatusDefect = '0'   " & vbCrLf & _
                          "  LEFT JOIN PO_Detail_Export POD ON POD.PONo = RD.PONO  " & vbCrLf & _
                          "  AND POD.OrderNo1 = RD.OrderNo  " & vbCrLf & _
                          "  AND POD.AffiliateID = RD.AffiliateID  AND POD.SupplierID = RD.SupplierID  " & vbCrLf & _
                          "  AND POD.PartNO = RD.PartNo  " & vbCrLf & _
                          "  LEFT JOIN PO_Master_export POM ON POM.PONo = POD.PONO  " & vbCrLf & _
                          "  AND POM.OrderNo1 = POD.OrderNo1  " & vbCrLf & _
                          "  AND POM.AffiliateID = POD.AffiliateID  " & vbCrLf & _
                          "  AND POM.SupplierID = POD.SupplierID  " & vbCrLf & _
                          "  LEFT JOIN MS_Price MPR ON MPR.PartNO = SHD.PartNo  " & vbCrLf & _
                          "  AND MPR.AffiliateID = SHD.AffiliateID  " & vbCrLf

        ls_sql = ls_sql + "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112)  " & vbCrLf & _
                          "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)  " & vbCrLf & _
                          "  AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "'" & vbCrLf & _
                          "  LEFT JOIN MS_PriceCls MPC ON MPC.PriceCls = SHM.TermDelivery " & vbCrLf & _
                          "  OUTER APPLY (SELECT TOP 1 * FROM dbo.CompanyProfile WHERE ActiveDate < SHM.ShippingInstructionDate ORDER BY ActiveDate DESC) Company " & vbCrLf & _
                          "  WHERE SHM.ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'  " & vbCrLf & _
                          "  AND SHM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf & _
                          "  AND SHM.ForwarderID = '" & Trim(cboForwarder.Text) & "'  order by SHD.partno, Rtrim(RB.Label1) + '-' + Rtrim(Rb.Label2)  "

        Session("REPORT") = "INV-EX"
        Session("Query") = ls_sql
        Response.Redirect("~/ShippingInstruction/ShippingViewReportExportCR.aspx")
    End Sub

    Private Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If Trim(cboCreate.Text = "CREATE") Then
            Call up_GridLoad()
        ElseIf Trim(cboCreate.Text = "UPDATE") Then
            'Call up_GridLoad()
            up_IsiInvoice(Trim(cboShippingNo.Text))
            txtBLNo.Text = Session("BLNO")
            etdvendor.Text = Session("etdvendor")
            etdport.Text = Session("etdport")
            etaport.Text = Session("etaport")
            etafactory.Text = Session("etafactory")
            txtSend.Text = Session("sending")

            txtVia.Text = Session("l_Via")
            txtmeasurement.Text = Session("l_Measurement")
            txtTotalPallet.Text = Session("l_TotalPallet")
            txtGrossWeight.Text = Session("l_GrossWeight")

            dtBLDate.Text = Session("l_BLDate")
            dtShippingDate.Text = Session("l_ShippDate")

            txtShippingLine.Text = Session("l_ShipLine")
            txtVoyage.Text = Session("l_Voyage")
            txtVessel.Text = Session("l_Vessel")

            up_fillcombo()
            up_fillcombocreateupdate()
            up_fillcombopackinglist(cboAffiliateCode.Text, cboForwarder.Text)
            cboCreate.SelectedIndex = 1

            cbofreight.Text = Trim(Session("l_Fright"))
            cboterm.Text = Trim(Session("l_Term"))
            cboService.Text = Trim(Session("l_Service"))
            txtterm.Text = Trim(Session("l_TermCls"))

            Call up_GridLoadUpdate()
        End If
        onOffButton()
        'uf_ButtonSendEDI()
    End Sub

#End Region

#Region "PROCEDURE"
    Private Function uf_Approve() As Integer
        Dim ls_sql As String
        Dim x As Integer
        Try
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("AffiliateMaster")
                    ls_sql = " update ShippingInstruction_Master set EDICls = '1'" & vbCrLf & _
                                " Where AffiliateID='" + cboAffiliateCode.Text + "' AND ShippingInstructionNo='" + cboShippingNo.Text + "' AND ForwarderID='" + cboForwarder.Text + "' " & vbCrLf

                    Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                    x = SqlComm.ExecuteNonQuery()

                    SqlComm.Dispose()
                    sqlTran.Commit()
                End Using
                sqlConn.Close()
            End Using
            Return x
        Catch ex As Exception
            errMsg = ex.Message.ToString()
            Return 0
        End Try
    End Function

    Private Sub uf_ButtonSendEDI()
        Dim ds As New DataSet
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim SqlComm As New SqlCommand("sp_ShippingInstruction_SendEDI", sqlConn)
            SqlComm.CommandType = CommandType.StoredProcedure
            SqlComm.Parameters.AddWithValue("AffiliateID", cboAffiliateCode.Text)
            SqlComm.Parameters.AddWithValue("ShippingInstructionNo", cboShippingNo.Text)
            SqlComm.Parameters.AddWithValue("ForwarderID", cboForwarder.Text)
            SqlComm.ExecuteNonQuery()

            Dim da As New SqlDataAdapter(SqlComm)
            da.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(1).Rows.Count > 0 Then
                    Dim a, b As String
                    a = ds.Tables(1).Rows(0)("ExcelCls").ToString()
                    b = ds.Tables(1).Rows(0)("EdiCls").ToString()

                    If a = "2" And b = "0" Then
                        btnSendEDI.Enabled = True
                        'btnDelete.Enabled = True
                    Else
                        btnSendEDI.Enabled = False
                        'btnDelete.Enabled = False
                    End If
                Else
                    btnSendEDI.Enabled = False
                End If
            Else
                btnSendEDI.Enabled = False
            End If

            SqlComm.Dispose()
            sqlConn.Close()
        End Using
    End Sub

    Private Function CreateInvoiceNo(ByVal pAffiliateID As String, ByVal pShipby As String) As String
        Dim ls_sql As String = ""
        Dim ls_Sfx As String = ""
        Dim ls_SeqNo As String = ""
        Dim ls_Year As String = ""
        Dim ls_Temp As Integer = 0

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("uploadPO")
                ls_sql = " select  " & vbCrLf & _
                      " 	SuffixInvoice =  " & vbCrLf & _
                      " 	CASE WHEN OverseasCls = '1' THEN 'E' else 'D' END  + " & vbCrLf & _
                      " 	AffiliateCls + POCode, SeqNO = SeqNo + 1   " & vbCrLf & _
                      " from MS_Affiliate where AffiliateID ='" & pAffiliateID & "'" & vbCrLf

                Dim sqlCmd2 As New SqlCommand(ls_sql, sqlConn, sqlTran)
                Dim sqlDA2 As New SqlDataAdapter(sqlCmd2)
                Dim ds2 As New DataSet
                sqlDA2.Fill(ds2)

                If ds2.Tables(0).Rows.Count > 0 Then
                    ls_Sfx = ds2.Tables(0).Rows(0)("SuffixInvoice")
                    ls_Year = Right(CStr(Year(Now)), 1)
                    ls_Temp = CDbl(ds2.Tables(0).Rows(0)("SeqNo"))
                    If ls_Temp <= 9 Then
                        ls_SeqNo = "000" & ds2.Tables(0).Rows(0)("SeqNo")
                    ElseIf ls_Temp <= 99 Then
                        ls_SeqNo = "00" & ds2.Tables(0).Rows(0)("SeqNo")
                    ElseIf ls_Temp <= 999 Then
                        ls_SeqNo = "0" & ds2.Tables(0).Rows(0)("SeqNo")
                    ElseIf ls_Temp < 9999 Then
                        ls_SeqNo = "0001"
                    End If
                Else
                    ls_Sfx = ""
                End If
                CreateInvoiceNo = ls_Sfx + ls_Year + ls_SeqNo + pShipby
            End Using
        End Using
    End Function

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""

        'Affiliate ID
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(AffiliateName) FROM MS_Affiliate  where isnull(overseascls, '0') = '1'" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliateCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate Code")
                .Columns(0).Width = 100
                .Columns.Add("Affiliate Name")
                .Columns(1).Width = 240
                .TextField = "Affiliate Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Forwarder ID
        ls_sql = "SELECT [Forwarder Code] = RTRIM(ForwarderID) ,[Forwarder Name] = RTRIM(ForwarderName) FROM MS_Forwarder " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboForwarder
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Forwarder Code")
                .Columns(0).Width = 100
                .Columns.Add("Forwarder Name")
                .Columns(1).Width = 240

                .TextField = "Forwarder Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Supplier Code
        ls_sql = "SELECT [Supplier Code] = '==ALL==' , [Supplier Name] = '==ALL==' UNION ALL SELECT [Supplier Code] = RTRIM(supplierID) ,[Supplier Name] = RTRIM(SupplierName) FROM MS_Supplier " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboSupplierCode
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Supplier Code")
                .Columns(0).Width = 100
                .Columns.Add("Supplier Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0
                .TextField = "Supplier Code"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Part No
        ls_sql = "SELECT [Part No] = '==ALL==' , [Part Name] = '==ALL==' UNION ALL SELECT [Part No] = RTRIM(PartNo) ,[Part Name] = RTRIM(PartName) FROM MS_Parts --where FinishGoodCls = '2' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboPartNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Part No")
                .Columns(0).Width = 70
                .Columns.Add("Part Name")
                .Columns(1).Width = 240
                .SelectedIndex = 0

                .TextField = "Part No"
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Term
        ls_sql = " SELECT Term = RTRIM(MPC.Description)/*CASE WHEN RTRIM(MPC.Description) = 'FCA - BOAT' THEN 'FCA' " & vbCrLf & _
                  " 			  WHEN RTRIM(MPC.Description) = 'FCA - AIR' THEN 'FCA'	 " & vbCrLf & _
                  " 			  WHEN RTRIM(MPC.Description) = 'CIF - BOAT' THEN 'CIF' " & vbCrLf & _
                  " 			  WHEN RTRIM(MPC.Description) = 'CIF - AIR' THEN 'CIF' " & vbCrLf & _
                  " 			  ELSE RTRIM(MPC.Description) END*/,	cls = pricecls  " & vbCrLf & _
                  " FROM ms_Pricecls MPC --where PriceCls not in ('2','4')"
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboterm
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Term")
                .Columns(0).Width = 100
                .Columns.Add("cls")
                .Columns(1).Width = 0
                .TextField = "Term"
                .DataBind()
                .Text = "FCA - AIR"
                txtterm.Text = "1"
            End With

            sqlConn.Close()
        End Using

        'Freight
        ls_sql = "SELECT x=0,Freight = 'COLLECT' UNION Select x=1,Freight = 'PREPAID' UNION Select x=2,Freight = 'DLL' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cbofreight
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Freight")
                .Columns(0).Width = 100
                .TextField = "Freight"
                .DataBind()
                .Text = "COLLECT"
            End With

            sqlConn.Close()
        End Using

        'Term of service
        ls_sql = "SELECT x=0,Freight = 'FCL' UNION Select x=1,Freight = 'LCL' UNION Select x=2,Freight = 'AIR FREIGHT' " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboService
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Freight")
                .Columns(0).Width = 100
                .TextField = "Freight"
                .DataBind()
                .Text = "FCL"
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim ls_filter As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim tentukanBoat As String = cboService.Text.Trim
            Dim tentukanTerm As String = cboterm.Text.Trim

            Dim PriceCls As String = 0

            ''If tentukanBoat = "FCL" Or tentukanBoat = "LCL" Then
            'If tentukanTerm = "FCA" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            '    PriceCls = "2"
            'ElseIf tentukanTerm = "FCA" Then
            '    PriceCls = "1"
            'End If

            'If tentukanTerm = "CIF" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            '    PriceCls = "4"
            'ElseIf tentukanTerm = "CIF" Then
            '    PriceCls = "3"
            'End If

            'If tentukanTerm = "DDU PASI" Then
            '    PriceCls = "5"
            'ElseIf tentukanTerm = "DDU Affiliate" Then
            '    PriceCls = "6"
            'ElseIf tentukanTerm = "EX-Work" Then
            '    PriceCls = "7"
            'ElseIf tentukanTerm = "FOB" Then
            '    PriceCls = "8"
            'End If

            PriceCls = uf_PriceCls(tentukanTerm)

            'ls_SQL = " Select ROW_NUMBER() OVER (ORDER BY AffiliateID) AS RowNo, * from ( " & vbCrLf & _
            '      " 	SELECT distinct  " & vbCrLf & _
            '      " 		'1' Act, CASE WHEN isnull(RD.PartNo,'0') = '0' THEN '0' ELSE '0' END AdaData, " & vbCrLf & _
            '      " 		RTRIM(RD.OrderNo)OrderNo, RTRIM(RD.PartNo)PartNo, RTRIM(PartName)PartName, RTRIM(MP.UnitCls + ' - ' + Description)UnitCls, RTRIM(Description)UnitClsDesc,PMP.QtyBox, GoodRecQty = Replace(CONVERT(char,isnull(RB.Box,0) * Isnull(PMP.QtyBox,0)),'.00','') ,   " & vbCrLf & _
            '      " 		ReceiveDate, RM.AffiliateID, RM.ForwarderID, Replace(CONVERT(char,isnull(RB.Box,0) * Isnull(PMP.QtyBox,0)),'.00','')  As ShippingQty,  " & vbCrLf & _
            '      " 		BoxQty = isnull(RB.Box,0), RTRIM(RM.SupplierID)SupplierID, RTRIM(SupplierName)SupplierName,  " & vbCrLf & _
            '      " 		ETDPort = ETDPort1, " & vbCrLf & _
            '      " 		RM.SuratJalanNo, MPR.Price, " & vbCrLf
            'If IsNewLabel = False Then
            '    ls_SQL = ls_SQL + " LabelNo = isnull(Rtrim(PL1.LabelNo) + '-' + Rtrim(PL2.LabelNo ),'') " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + " LabelNo = isnull(SID.BoxNo, Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2)) " & vbCrLf
            'End If
            'ls_SQL = ls_SQL + " 	FROM dbo.ReceiveForwarder_Master RM  " & vbCrLf & _
            '      " 	LEFT JOIN dbo.ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratJalanNo  " & vbCrLf & _
            '      " 		AND RM.SupplierID = RD.SupplierID   " & vbCrLf

            'ls_SQL = ls_SQL + " 		AND RM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
            '                  " 		AND RM.PONo = RD.PONo  " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.MS_Parts MP ON RD.PartNo = MP.PartNo  " & vbCrLf

            'If IsNewLabel = False Then
            '    ls_SQL = ls_SQL + "     LEFT JOIN (select SuratJalanNo_FWD, POno, AffiliateID, SupplierID, PartNo, min(labelNo) as labelno from PrintLabelExport group by SuratJalanNo_FWD,POno, AffiliateID, SupplierID, PartNo) PL1 " & vbCrLf & _
            '                      "         ON PL1.SuratJalanNo_FWD = RM.SuratJalanNo and PL1.AffiliateID = RM.AffiliateID and PL1.SupplierID = RM.SupplierID and PL1.PartNo = RD.PartNo " & vbCrLf & _
            '                      "     LEFT JOIN (select SuratJalanNo_FWD, POno, AffiliateID, SupplierID, PartNo, Max(labelNo) as labelno from PrintLabelExport group by SuratJalanNo_FWD,POno, AffiliateID, SupplierID, PartNo) PL2 " & vbCrLf & _
            '                      "         ON PL2.SuratJalanNo_FWD = RM.SuratJalanNo and PL2.AffiliateID = RM.AffiliateID and PL2.SupplierID = RM.SupplierID and PL2.PartNo = RD.PartNo " & vbCrLf

            'Else
            '    ls_SQL = ls_SQL + "LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                      "AND RB.SupplierID = RD.SupplierID  " & vbCrLf & _
            '                      "AND RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
            '                      "AND RB.PONo = RD.PONo  " & vbCrLf & _
            '                      "AND RB.OrderNo = RD.OrderNo  " & vbCrLf & _
            '                      "AND RB.PartNo = RD.PartNo  " & vbCrLf & _
            '                      "AND RB.StatusDefect = '0'  " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " 	LEFT JOIN dbo.MS_PartMapping PMP ON RD.PartNo = PMP.PartNo and RM.AffiliateID = PMP.AffiliateID and RM.SupplierID = PMP.SupplierID  " & vbCrLf & _
            '                      " 	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf & _
            '                      " 	LEFT JOIN dbo.MS_Supplier MSS ON RM.SupplierID = MSS.SupplierID  " & vbCrLf & _
            '                      " 	LEFT JOIN dbo.ShippingInstruction_Detail SID ON SID.OrderNo = RD.OrderNo   " & vbCrLf & _
            '                      " 		AND SID.PartNo = RD.PartNo AND SID.SupplierID = RD.SupplierID " & vbCrLf & _
            '                      " 		AND SID.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                      " 		AND SID.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                      " 	LEFT JOIN ShippingInstruction_Master SIM ON SIM.ShippingInstructionNo = SID.ShippingInstructionNo " & vbCrLf

            'ls_SQL = ls_SQL + " 		AND SIM.AffiliateID = SID.AffiliateID " & vbCrLf & _
            '                  " 		AND SIM.ForwarderID = SID.ForwarderID " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  " 		AND (RM.OrderNo =  PME.OrderNo1 or RM.OrderNo =  PME.OrderNo2 " & vbCrLf & _
            '                  " 		or RM.OrderNo =  PME.OrderNo3 or RM.OrderNo =  PME.OrderNo4 or RM.OrderNo =  PME.OrderNo5) " & vbCrLf & _
            '                  "         and PME.SupplierID = RM.SupplierID AND PME.PONo = RM.PONo " & vbCrLf & _
            '                  "  LEFT JOIN MS_Price MPR ON MPR.PartNO = SID.PartNo  " & vbCrLf & _
            '                  "  AND MPR.AffiliateID = SID.AffiliateID  " & vbCrLf

            'ls_SQL = ls_SQL + "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SIM.ShippingInstructionDate,'')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112)  " & vbCrLf & _
            '                  "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SIM.ShippingInstructionDate,'')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)  " & vbCrLf & _
            '                  "  AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "'" & vbCrLf & _
            '                  "  	LEFT JOIN dbo.MS_PartMapping PMP ON SID.PartNo = PMP.PartNo and SID.AffiliateID = PMP.AffiliateID and SID.SupplierID = PMP.SupplierID " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.MS_Parts MP ON SID.PartNo = MP.PartNo   " & vbCrLf & _
            '                  " 	WHERE ISNULL(RD.OrderNo, '') <> ''  and SIM.shippingInstructionno = '" & Trim(cboShippingNo.Text) & " ' " & vbCrLf & _
            '                  " 	UNION ALL " & vbCrLf & _
            '                  " 	SELECT distinct  " & vbCrLf & _
            '                  " 		'0' Act, CASE WHEN isnull(RD.PartNo,'0') = '0' THEN '0' ELSE '0' END AdaData, " & vbCrLf & _
            '                  " 		RTRIM(RD.OrderNo)OrderNo, RTRIM(RD.PartNo)PartNo, RTRIM(PartName)PartName, RTRIM(MP.UnitCls + ' - ' + Description)UnitCls, RTRIM(Description)UnitClsDesc,PMP.QtyBox, GoodRecQty = Replace(CONVERT(char,isnull(RB.Box,0) * Isnull(QtyBox,0)),'.00','') ,   " & vbCrLf & _
            '                  " 		ReceiveDate, RM.AffiliateID, RM.ForwarderID, Replace(CONVERT(char,isnull(RB.Box,0) * Isnull(QtyBox,0)),'.00','') As ShippingQty,  " & vbCrLf

            'ls_SQL = ls_SQL + " 		BoxQty = isnull(Rb.Box,0), RTRIM(RM.SupplierID)SupplierID, RTRIM(SupplierName)SupplierName,  " & vbCrLf & _
            '                  " 		ETDPort = ETDPort1, " & vbCrLf & _
            '                  " 		RM.SuratJalanNo, 0 Price, " & vbCrLf

            'If IsNewLabel = False Then
            '    ls_SQL = ls_SQL + " LabelNo = isnull(Rtrim(PL1.LabelNo) + '-' + Rtrim(PL2.LabelNo ),'') " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + " LabelNo = isnull(Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2),'') " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " 	FROM dbo.ReceiveForwarder_Master RM  " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratJalanNo  " & vbCrLf & _
            '                  " 		AND RM.SupplierID = RD.SupplierID   " & vbCrLf & _
            '                  " 		AND RM.AffiliateID = RD.AffiliateID  " & vbCrLf & _
            '                  " 		AND RM.PONo = RD.PONo  " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.MS_Parts MP ON RD.PartNo = MP.PartNo  " & vbCrLf

            'If IsNewLabel = False Then
            '    ls_SQL = ls_SQL + "     LEFT JOIN (select SuratJalanNo_FWD, POno, AffiliateID, SupplierID, PartNo, min(labelNo) as labelno from PrintLabelExport group by SuratJalanNo_FWD,POno, AffiliateID, SupplierID, PartNo) PL1 " & vbCrLf & _
            '                      "         ON PL1.SuratJalanNo_FWD = RM.SuratJalanNo and PL1.AffiliateID = RM.AffiliateID and PL1.SupplierID = RM.SupplierID and PL1.PartNo = RD.PartNo " & vbCrLf & _
            '                      "     LEFT JOIN (select SuratJalanNo_FWD, POno, AffiliateID, SupplierID, PartNo, Max(labelNo) as labelno from PrintLabelExport group by SuratJalanNo_FWD,POno, AffiliateID, SupplierID, PartNo) PL2 " & vbCrLf & _
            '                      "         ON PL2.SuratJalanNo_FWD = RM.SuratJalanNo and PL2.AffiliateID = RM.AffiliateID and PL2.SupplierID = RM.SupplierID and PL2.PartNo = RD.PartNo " & vbCrLf

            'Else
            '    ls_SQL = ls_SQL + "LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                      "AND RB.SupplierID = RD.SupplierID  " & vbCrLf & _
            '                      "AND RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
            '                      "AND RB.PONo = RD.PONo  " & vbCrLf & _
            '                      "AND RB.OrderNo = RD.OrderNo  " & vbCrLf & _
            '                      "AND RB.PartNo = RD.PartNo  " & vbCrLf & _
            '                      "AND RB.StatusDefect = '0'  " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " 	LEFT JOIN dbo.MS_PartMapping PMP ON RD.PartNo = PMP.PartNo and RM.AffiliateID = PMP.AffiliateID and RM.SupplierID = PMP.SupplierID  " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls  " & vbCrLf

            'ls_SQL = ls_SQL + " 	LEFT JOIN dbo.MS_Supplier MSS ON RM.SupplierID = MSS.SupplierID  " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  " 		AND (RM.OrderNo =  PME.OrderNo1 or RM.OrderNo =  PME.OrderNo2 " & vbCrLf & _
            '                  " 		or RM.OrderNo =  PME.OrderNo3 or RM.OrderNo =  PME.OrderNo4 or RM.OrderNo =  PME.OrderNo5) " & vbCrLf & _
            '                  "         and PME.SupplierID = RM.SupplierID AND PME.PONo = RM.PONo " & vbCrLf & _
            '                  " 	WHERE ISNULL(RD.OrderNo, '') <> ''   " & vbCrLf & _
            '                  " 		and RTrim(RD.SuratJalanNo) + Rtrim(RD.AffiliateID) + Rtrim(RD.OrderNo)+RTRIM(RD.SupplierID)+RTRIM(RD.PartNo) " & vbCrLf & _
            '                  " 			NOT IN (SELECT DISTINCT RTrim(SuratJalanNo) + Rtrim(AffiliateID) + Rtrim(OrderNo)+RTRIM(SupplierID)+RTRIM(PartNo) From ShippingInstruction_Detail where " & vbCrLf & _
            '                  " 			suratjalanno = RD.SuratJalanNo and AffiliateID = RD.AffiliateID AND SupplierID = RD.SupplierID and partno = RD.Partno and orderno = RD.OrderNo) " & vbCrLf

            ''Supplier Code
            'If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND RM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            'End If

            ''Part No
            'If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND RD.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            'End If

            ''Order No
            'If Trim(txtOrderNo.Text) <> "" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND RM.OrderNo = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            'End If

            'If Session("isSJ") <> "" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND RM.SuratJalanNo in (" & Trim(Session("isSJ")) & ") " & vbCrLf
            'End If

            'If Session("SHGENERAL") <> "" Then
            '    ls_SQL = ls_SQL + _
            '        "           AND RTRIM(RM.OrderNo) + RTRIM(RM.AffiliateID) + RTRIM(RM.SupplierID) in (" & Session("SHGENERAL") & ") " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + ")x  Order By AffiliateID, LabelNo"
            ls_SQL = "  Select ROW_NUMBER() OVER (ORDER BY AffiliateID) AS RowNo, * from (  " & vbCrLf & _
                      " 	SELECT distinct   " & vbCrLf & _
                      " 		'0' Act, CASE WHEN isnull(RD.PartNo,'0') = '0' THEN '0' ELSE '0' END AdaData,  " & vbCrLf & _
                      " 		RTRIM(RD.OrderNo)OrderNo, RTRIM(RD.PartNo)PartNo, RTRIM(PartName)PartName, RTRIM(MP.UnitCls + ' - ' + Description)UnitCls, RTRIM(Description)UnitClsDesc,QtyBox = ISNULL(PMD.POQtyBox,PMP.QtyBox), GoodRecQty = Replace(CONVERT(char,isnull(RB.Box,0) * ISNULL(PMD.POQtyBox,PMP.QtyBox)),'.00','') ,    " & vbCrLf & _
                      " 		ReceiveDate, RM.AffiliateID, RM.ForwarderID, Replace(CONVERT(char,isnull(RB.Box,0) * ISNULL(PMD.POQtyBox,PMP.QtyBox)),'.00','') As ShippingQty,   " & vbCrLf & _
                      " 		BoxQty = isnull(Rb.Box,0), RTRIM(RM.SupplierID)SupplierID, RTRIM(SupplierName)SupplierName,   " & vbCrLf & _
                      " 		ETDPort = ETDPort1,  " & vbCrLf & _
                      " 		RM.SuratJalanNo, ISNULL(MPR.Price,0) Price,  " & vbCrLf & _
                      " 		LabelNo = isnull(Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2),'')  " & vbCrLf & _
                      " 	FROM dbo.ReceiveForwarder_Master RM   " & vbCrLf & _
                      " 	LEFT JOIN dbo.ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratJalanNo   "

            ls_SQL = ls_SQL + " 		AND RM.SupplierID = RD.SupplierID    " & vbCrLf & _
                              " 		AND RM.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                              " 		AND RM.PONo = RD.PONo   " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_Parts MP ON RD.PartNo = MP.PartNo   " & vbCrLf & _
                              " 	LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo  " & vbCrLf & _
                              " 		AND RB.SupplierID = RD.SupplierID   " & vbCrLf & _
                              " 		AND RB.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                              " 		AND RB.PONo = RD.PONo   " & vbCrLf & _
                              " 		AND RB.OrderNo = RD.OrderNo   " & vbCrLf & _
                              " 		AND RB.PartNo = RD.PartNo   " & vbCrLf & _
                              " 		AND RB.StatusDefect = '0' "

            ls_SQL = ls_SQL + " 	LEFT JOIN MS_Price MPR ON MPR.PartNO = RD.PartNo   " & vbCrLf & _
                              " 		AND MPR.AffiliateID = RD.AffiliateID   " & vbCrLf & _
                              " 		AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL('" & etdport.Text & "','')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112)   " & vbCrLf & _
                              " 		AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL('" & etdport.Text & "','')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)   " & vbCrLf & _
                              " 		AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "'   " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_PartMapping PMP ON RD.PartNo = PMP.PartNo and RM.AffiliateID = PMP.AffiliateID and RM.SupplierID = PMP.SupplierID   " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
                              " 	LEFT JOIN dbo.MS_Supplier MSS ON RM.SupplierID = MSS.SupplierID   " & vbCrLf & _
                              " 	LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = RD.AffiliateID  " & vbCrLf & _
                              " 		AND (RM.OrderNo =  PME.OrderNo1 or RM.OrderNo =  PME.OrderNo2  " & vbCrLf & _
                              " 		or RM.OrderNo =  PME.OrderNo3 or RM.OrderNo =  PME.OrderNo4 or RM.OrderNo =  PME.OrderNo5)  "

            ls_SQL = ls_SQL + " 		and PME.SupplierID = RM.SupplierID AND PME.PONo = RM.PONo  " & vbCrLf & _
                              "     LEFT JOIN dbo.PO_Detail_Export PMD on PMD.PONo = PME.PONo and PME.AffiliateID = PMD.AffiliateID and PME.SupplierID = PMD.SupplierID and PMD.PartNo = RD.PartNo " & vbCrLf & _
                              " 	WHERE ISNULL(RD.OrderNo, '') <> ''    " & vbCrLf & _
                              " 		and RTrim(RD.SuratJalanNo) + Rtrim(RD.AffiliateID) + Rtrim(RD.OrderNo)+RTRIM(RD.SupplierID)+RTRIM(RD.PartNo)  " & vbCrLf & _
                              " 	NOT IN (SELECT DISTINCT RTrim(SuratJalanNo) + Rtrim(AffiliateID) + Rtrim(OrderNo)+RTRIM(SupplierID)+RTRIM(PartNo) From ShippingInstruction_Detail where  " & vbCrLf & _
                              " 	suratjalanno = RD.SuratJalanNo and AffiliateID = RD.AffiliateID AND SupplierID = RD.SupplierID and partno = RD.Partno and orderno = RD.OrderNo)  " & vbCrLf

            'Supplier Code
            If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND RM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            End If

            'Part No
            If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND RD.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            End If

            'Order No
            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "           AND RM.OrderNo = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            If Session("isSJ") <> "" Then
                ls_SQL = ls_SQL + _
                    "           AND RM.SuratJalanNo in (" & Trim(Session("isSJ")) & ") " & vbCrLf
            End If

            If Session("SHGENERAL") <> "" Then
                ls_SQL = ls_SQL + _
                    "           AND RTRIM(RM.OrderNo) + RTRIM(RM.AffiliateID) + RTRIM(RM.SupplierID) + RTRIM(RM.SuratJalanNo) in (" & Session("SHGENERAL") & ") " & vbCrLf
            End If

            ls_SQL = ls_SQL + ")x  Order By AffiliateID, LabelNo"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call ColorGrid()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadUpdate()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            'ls_SQL = " select ROW_NUMBER() OVER (ORDER BY AffiliateID) AS RowNo, * from( " & vbCrLf & _
            '         " SELECT distinct " & vbCrLf & _
            '          " 	CASE WHEN isnull(SID.PartNo,'0') = '0' THEN '0' ELSE '1' END Act, CASE WHEN isnull(SID.PartNo,'0') = '0' THEN '0' ELSE '1' END AdaData," & vbCrLf & _
            '          " 	RTRIM(RD.OrderNo)OrderNo, RTRIM(RD.PartNo)PartNo, RTRIM(PartName)PartName, RTRIM(MP.UnitCls + ' - ' + Description)UnitCls,  RTRIM(Description)UnitClsDesc, MPM.QtyBox, GoodRecQty = Replace(CONVERT(char,isnull(RB.Box,0) * Isnull(MPM.QtyBox,0)),'.00',''),  " & vbCrLf & _
            '          " 	ReceiveDate, RM.AffiliateID, RM.ForwarderID, ShippingQty As ShippingQty, " & vbCrLf & _
            '          " 	BoxQty = RB.Box, RTRIM(RM.SupplierID)SupplierID, RTRIM(SupplierName)SupplierName, " & vbCrLf & _
            '          " 	 ETDPort = CASE RM.OrderNo" & vbCrLf & _
            '          " 	        WHEN PME.OrderNo1 THEN PME.ETDPort1 " & vbCrLf & _
            '          " 	        WHEN PME.OrderNo2 THEN PME.ETDPort2 " & vbCrLf & _
            '          " 	        WHEN PME.OrderNo3 THEN PME.ETDPort3 " & vbCrLf & _
            '          " 	        WHEN PME.OrderNo4 THEN PME.ETDPort4 " & vbCrLf & _
            '          " 	        WHEN PME.OrderNo5 THEN PME.ETDPort5 " & vbCrLf & _
            '          " 	        END, SID.SuratJalanNo, " & vbCrLf

            'If IsNewLabel = False Then
            '    ls_SQL = ls_SQL + " LabelNo = Rtrim(PL1.LabelNo) + '-' + Rtrim(PL2.LabelNo ) " & vbCrLf
            'Else
            '    ls_SQL = ls_SQL + " LabelNo = Rtrim(RB.Label1) + '-' + Rtrim(RB.Label2) " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " FROM dbo.ReceiveForwarder_Master RM " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.ReceiveForwarder_Detail RD ON RM.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                  " 		AND RM.SupplierID = RD.SupplierID  " & vbCrLf & _
            '                  " 		AND RM.AffiliateID = RD.AffiliateID " & vbCrLf & _
            '                  " 		AND RM.PONo = RD.PONo " & vbCrLf

            'If IsNewLabel = False Then
            '    ls_SQL = ls_SQL + "     LEFT JOIN (select SuratJalanNo_FWD, POno, AffiliateID, SupplierID, PartNo, min(labelNo) as labelno from PrintLabelExport group by SuratJalanNo_FWD,POno, AffiliateID, SupplierID, PartNo) PL1 " & vbCrLf & _
            '                      "         ON PL1.SuratJalanNo_FWD = RM.SuratJalanNo and PL1.AffiliateID = RM.AffiliateID and PL1.SupplierID = RM.SupplierID and PL1.PartNo = RD.PartNo " & vbCrLf & _
            '                      "     LEFT JOIN (select SuratJalanNo_FWD, POno, AffiliateID, SupplierID, PartNo, Max(labelNo) as labelno from PrintLabelExport group by SuratJalanNo_FWD,POno, AffiliateID, SupplierID, PartNo) PL2 " & vbCrLf & _
            '                      "         ON PL2.SuratJalanNo_FWD = RM.SuratJalanNo and PL2.AffiliateID = RM.AffiliateID and PL2.SupplierID = RM.SupplierID and PL2.PartNo = RD.PartNo " & vbCrLf

            'Else
            '    ls_SQL = ls_SQL + "LEFT JOIN ReceiveForwarder_DetailBox RB ON RB.SuratJalanNo = RD.SuratJalanNo " & vbCrLf & _
            '                      "AND RB.SupplierID = RD.SupplierID  " & vbCrLf & _
            '                      "AND RB.AffiliateID = RD.AffiliateID  " & vbCrLf & _
            '                      "AND RB.PONo = RD.PONo  " & vbCrLf & _
            '                      "AND RB.OrderNo = RD.OrderNo  " & vbCrLf & _
            '                      "AND RB.PartNo = RD.PartNo  " & vbCrLf & _
            '                      "AND RB.StatusDefect = '0'  " & vbCrLf
            'End If

            'ls_SQL = ls_SQL + " 	LEFT JOIN dbo.MS_Parts MP ON RD.PartNo = MP.PartNo " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls " & vbCrLf & _
            '                  "     LEFT JOIN MS_PARTMAPPING MPM ON MPM.PartNo = RD.PartNo AND MPM.AffiliateID = RD.AffiliateID AND MPM.SupplierID = RD.SupplierID " & vbCrLf

            'ls_SQL = ls_SQL + " 	LEFT JOIN dbo.MS_Supplier MSS ON RM.SupplierID = MSS.SupplierID " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.ShippingInstruction_Detail SID ON SID.OrderNo = RD.OrderNo  " & vbCrLf & _
            '                  " 		AND SID.PartNo = RD.PartNo AND SID.SupplierID = RD.SupplierID " & vbCrLf & _
            '                  "         AND SID.SuratJalanNo = RM.SuratJalanNo  " & vbCrLf & _
            '                  " 	LEFT JOIN dbo.ShippingInstruction_Master SIM ON SIM.AffiliateID = RD.AffiliateID" & vbCrLf & _
            '                  " 		AND SIM.AffiliateID = SID.AffiliateID AND SIM.ForwarderID = SID.ForwarderID " & vbCrLf & _
            '                  "         AND SIM.ShippingInstructionNo = SID.ShippingInstructionNo " & vbCrLf & _
            '                  "     LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = RD.AffiliateID" & vbCrLf & _
            '                  "     AND (RM.OrderNo =  PME.OrderNo1 or RM.OrderNo =  PME.OrderNo2" & vbCrLf & _
            '                  "     or RM.OrderNo =  PME.OrderNo3 or RM.OrderNo =  PME.OrderNo4 or RM.OrderNo =  PME.OrderNo5)" & vbCrLf & _
            '                  "         and PME.SupplierID = RM.SupplierID AND PME.PONo = RM.PONo " & vbCrLf & _
            '                  " WHERE 'A' = 'A' " & vbCrLf & _
            '                  " --AND ISNULL(SID.OrderNo, '') = '' AND ISNULL(SID.PartNo, '') = '' AND ISNULL(SID.SupplierID, '') = '' " & vbCrLf & _
            '                  " --AND ReceiveDate BETWEEN '' AND '' " & vbCrLf & _
            '                  " --AND RM.AffiliateID = 'PEMI' " & vbCrLf & _
            '                  " --AND ForwarderID = 'FORWARDER XYZ' " & vbCrLf & _
            '                  " --AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' " & vbCrLf & _
            '                  " --AND RM.SupplierID = 'TORICA' " & vbCrLf & _
            '                  " --AND RD.PartNo = '7283-1705-30' " & vbCrLf & _
            '                  " --AND RM.OrderNo = 'PC1511E' " & vbCrLf

            Dim tentukanBoat As String = cboService.Text.Trim
            Dim tentukanTerm As String = cboterm.Text.Trim

            Dim PriceCls As String = 0

            ''If tentukanBoat = "FCL" Or tentukanBoat = "LCL" Then
            'If tentukanTerm = "FCA" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            '    PriceCls = "2"
            'ElseIf tentukanTerm = "FCA" Then
            '    PriceCls = "1"
            'End If

            'If tentukanTerm = "CIF" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            '    PriceCls = "4"
            'ElseIf tentukanTerm = "CIF" Then
            '    PriceCls = "3"
            'End If

            'If tentukanTerm = "DDU PASI" Then
            '    PriceCls = "5"
            'ElseIf tentukanTerm = "DDU Affiliate" Then
            '    PriceCls = "6"
            'ElseIf tentukanTerm = "EX-Work" Then
            '    PriceCls = "7"
            'ElseIf tentukanTerm = "FOB" Then
            '    PriceCls = "8"
            'End If

            PriceCls = uf_PriceCls(tentukanTerm)

            ls_SQL = "  select ROW_NUMBER() OVER (ORDER BY AffiliateID) AS RowNo, * from(  " & vbCrLf & _
                      "  SELECT distinct  " & vbCrLf & _
                      "  	CASE WHEN isnull(SID.PartNo,'0') = '0' THEN '0' ELSE '1' END Act, CASE WHEN isnull(SID.PartNo,'0') = '0' THEN '0' ELSE '1' END AdaData, " & vbCrLf & _
                      "  	RTRIM(SID.OrderNo)OrderNo, RTRIM(SID.PartNo)PartNo, RTRIM(PartName)PartName, RTRIM(MP.UnitCls + ' - ' + Description)UnitCls,   " & vbCrLf & _
                      " 	RTRIM(Description)UnitClsDesc, QtyBox = ISNULL(SID.POQtyBox,PMP.QtyBox), GoodRecQty = Replace(CONVERT(char,isnull(SID.GoodReceivingQty,0)),'.00',''),   " & vbCrLf & _
                      "  	ShippingInstructionDate, SID.AffiliateID, SID.ForwarderID, ShippingQty As ShippingQty,  " & vbCrLf & _
                      "  	BoxQty = SID.QtyBox, RTRIM(SID.SupplierID)SupplierID, RTRIM(SupplierName)SupplierName,  " & vbCrLf & _
                      "  	 ETDPort = CASE SID.OrderNo " & vbCrLf & _
                      "  	        WHEN PME.OrderNo1 THEN PME.ETDPort1  " & vbCrLf & _
                      "  	        WHEN PME.OrderNo2 THEN PME.ETDPort2  " & vbCrLf & _
                      "  	        WHEN PME.OrderNo3 THEN PME.ETDPort3  "

            ls_SQL = ls_SQL + "  	        WHEN PME.OrderNo4 THEN PME.ETDPort4  " & vbCrLf & _
                              "  	        WHEN PME.OrderNo5 THEN PME.ETDPort5  " & vbCrLf & _
                              "  	        END, SID.SuratJalanNo,  " & vbCrLf & _
                              "  LabelNo = SID.BoxNo, ISNULL(SID.Price,MPR.Price) Price" & vbCrLf & _
                              "  FROM 	 " & vbCrLf & _
                              " 	dbo.ShippingInstruction_Detail SID  " & vbCrLf & _
                              "  	LEFT JOIN ShippingInstruction_Master SIM ON SIM.ShippingInstructionNo = SID.ShippingInstructionNo  " & vbCrLf & _
                              "  		AND SIM.AffiliateID = SID.AffiliateID  " & vbCrLf & _
                              "  		AND SIM.ForwarderID = SID.ForwarderID  " & vbCrLf & _
                              "  LEFT JOIN MS_Price MPR ON MPR.PartNO = SID.PartNo  " & vbCrLf & _
                              "  AND MPR.AffiliateID = SID.AffiliateID  " & vbCrLf

            ls_SQL = ls_SQL + "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SIM.ETDPort,'')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112)  " & vbCrLf & _
                              "  AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SIM.ETDPort,'')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)  " & vbCrLf & _
                              "  AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "'" & vbCrLf & _
                                  "  	LEFT JOIN dbo.MS_PartMapping PMP ON SID.PartNo = PMP.PartNo and SID.AffiliateID = PMP.AffiliateID and SID.SupplierID = PMP.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN dbo.MS_Parts MP ON SID.PartNo = MP.PartNo   "

            ls_SQL = ls_SQL + "  	LEFT JOIN dbo.MS_UnitCls UC ON UC.UnitCls = MP.UnitCls   " & vbCrLf & _
                              "  	LEFT JOIN dbo.MS_Supplier MSS ON SID.SupplierID = MSS.SupplierID   	 " & vbCrLf & _
                              "  	LEFT JOIN dbo.PO_Master_Export PME ON PME.AffiliateID = SID.AffiliateID  " & vbCrLf & _
                              "  		AND (SID.OrderNo =  PME.OrderNo1 or SID.OrderNo =  PME.OrderNo2  " & vbCrLf & _
                              "  		or SID.OrderNo =  PME.OrderNo3 or SID.OrderNo =  PME.OrderNo4 or SID.OrderNo =  PME.OrderNo5)  " & vbCrLf & _
                              "          and PME.SupplierID = SID.SupplierID  " & vbCrLf & _
                              "  WHERE 'A' = 'A'  " & vbCrLf

            'Shipping Instruction No
            If Trim(cboShippingNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "          AND SIM.ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' " & vbCrLf
            End If

            'Supplier Code
            If Trim(cboSupplierCode.Text) <> "" And Trim(cboSupplierCode.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND SIM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            End If

            'Part No
            If Trim(cboPartNo.Text) <> "" And Trim(cboPartNo.Text) <> "==ALL==" Then
                ls_SQL = ls_SQL + _
                    "           AND SID.PartNo = '" & Trim(cboPartNo.Text) & "' " & vbCrLf
            End If

            'Order No
            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "           AND SIM.OrderNo = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + " )xx  Order By AffiliateID, LabelNo"
            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call ColorGrid()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
                Session.Remove("SHGENERAL")
            End With
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_IsiInvoice(ByVal pSH As String)
        Dim ls_SQL As String = ""
        Dim ls_HT As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  SELECT ShippingInstructionDate, ETDVendor = Isnull(ETDVendor,Getdate()), ETDPort = Isnull(ETDPort,Getdate()), " & vbCrLf & _
                      "  ETAPort = Isnull(ETAPort,Getdate()), ETAFactory = Isnull(ETAFactory,Getdate()), BLNo = ISNULL(BLNo,''), Freight, " & vbCrLf & _
                      " Term = MP.Description /*CASE WHEN RTRIM(MP.Description) = 'FCA - BOAT' THEN 'FCA' " & vbCrLf & _
                      " 			  WHEN RTRIM(MP.Description) = 'FCA - AIR' THEN 'FCA'	 " & vbCrLf & _
                      " 			  WHEN RTRIM(MP.Description) = 'CIF - BOAT' THEN 'CIF' " & vbCrLf & _
                      " 			  WHEN RTRIM(MP.Description) = 'CIF - AIR' THEN 'CIF' " & vbCrLf & _
                      " 			  ELSE RTRIM(MP.Description) END*/, TermDelivery,	 " & vbCrLf & _
                      "  ISNULL(SM.TypeOfService,'')TypeOfService,ISNULL(SM.Via,'')VIA,ISNULL(SM.Measurement,0)Measurement,ISNULL(SM.TotalPallet,0)TotalPallet,ISNULL(SM.GrossWeight,0)GrossWeight," & vbCrLf & _
                      "        ISNULL(SM.VesselS,'')VesselS, ISNULL(SM.NamaKapalS,'')NamaKapalS, ISNULL(SM.ShippingLineS,'')ShippingLineS, BLDate = CONVERT(char(11), CONVERT(datetime, BLDate),106), " & vbCrLf & _
                      " 		/*ExcelCls = Case when isnull(ExcelCls,0)= 1 then 'ALREADY SEND' ELSE 'ALREADY SEND TALLY TEMPLATE' END,*/ " & vbCrLf & _
                      "         ExcelCls = Case isnull(ExcelCls,0) WHEN 1 then 'ALREADY SEND' WHEN 2 THEN 'ALREADY SEND TALLY TEMPLATE' WHEN 0 THEN '' END, " & vbCrLf & _
                      " 		TallyCls = Case isnull(TallyCls,0) WHEN 1 then 'ALREADY RECEIVE TALLY' When 2 then 'ALREADY SEND INVOICE,TALLY,SI' when 0 then '' END " & vbCrLf & _
                      "  FROM ShippingInstruction_Master SM left join ms_priceCls MP ON MP.PriceCls = SM.TermDelivery WHERE ShippingInstructionNo = '" & pSH & "' " & vbCrLf & _
                      "  AND AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf & _
                      "  AND ForwarderID = '" & Trim(cboForwarder.Text) & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Try
                    With ds.Tables(0)
                        etdvendor.Text = Format(.Rows(0).Item("ETDVendor"), "dd MMM yyyy")
                        etdport.Text = Format(.Rows(0).Item("etdport"), "dd MMM yyyy")
                        etaport.Text = Format(.Rows(0).Item("etaport"), "dd MMM yyyy")
                        etafactory.Text = Format(.Rows(0).Item("etafactory"), "dd MMM yyyy")

                        dtShippingDate.Text = Format(.Rows(0).Item("ShippingInstructionDate"), "dd MMM yyyy")
                        dtBLDate.Text = Trim(.Rows(0).Item("BLDate"))

                        cbofreight.Text = .Rows(0).Item("Freight")
                        cboterm.Text = .Rows(0).Item("Term")
                        cboService.Text = .Rows(0).Item("TypeOfService")

                        txtVia.Text = .Rows(0).Item("Via")
                        txtmeasurement.Text = .Rows(0).Item("Measurement")
                        txtTotalPallet.Text = .Rows(0).Item("TotalPallet")
                        txtGrossWeight.Text = .Rows(0).Item("GrossWeight")

                        txtBLNo.Text = Trim(.Rows(0).Item("BLNO"))

                        txtShippingLine.Text = .Rows(0).Item("ShippingLineS")
                        txtVoyage.Text = .Rows(0).Item("NamaKapalS")
                        txtVessel.Text = .Rows(0).Item("VesselS")
                        txtterm.Text = .Rows(0).Item("TermDelivery")

                        Session("BLNO") = txtBLNo.Text
                        Session("etdvendor") = etdvendor.Text
                        Session("etdport") = etdport.Text
                        Session("etaport") = etaport.Text
                        Session("etafactory") = etafactory.Text
                        Session("sending") = txtSend.Text
                        Session("l_Fright") = cbofreight.Text
                        Session("l_Term") = cboterm.Text
                        Session("l_TermCls") = txtterm.Text
                        Session("l_Service") = cboService.Text

                        Session("l_Via") = txtVia.Text
                        Session("l_Measurement") = txtmeasurement.Text
                        Session("l_TotalPallet") = txtTotalPallet.Text
                        Session("l_GrossWeight") = txtGrossWeight.Text

                        Session("l_BLDate") = dtBLDate.Text
                        Session("l_ShippDate") = dtShippingDate.Text

                        Session("l_ShipLine") = txtShippingLine.Text
                        Session("l_Voyage") = txtVoyage.Text
                        Session("l_Vessel") = txtVessel.Text

                        If Trim(.Rows(0).Item("TallyCls")) = "" Then
                            txtSend.Text = Trim(.Rows(0).Item("ExcelCls"))
                            grid.JSProperties("cpSending") = Trim(.Rows(0).Item("ExcelCls"))
                            Session("sending") = txtSend.Text
                        Else
                            txtSend.Text = Trim(.Rows(0).Item("TallyCls"))
                            grid.JSProperties("cpSending") = Trim(.Rows(0).Item("TallyCls"))
                            Session("sending") = txtSend.Text
                        End If
                    End With
                Catch ex As Exception

                End Try
            End If
            sqlConn.Close()


        End Using
    End Sub

    Private Function uf_PriceCls(ByVal pDesc As String)
        Dim ls_SQL As String = ""
        Dim ls_Cls As Integer = 0
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " Select PriceCls From MS_PriceCls Where Description = '" & pDesc & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 200
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                ls_Cls = ds.Tables(0).Rows(0).Item("PriceCls")
            End If
        End Using

        Return ls_Cls
    End Function

    Private Sub ColorGrid()
        grid.VisibleColumns(1).CellStyle.BackColor = Color.White
        grid.VisibleColumns(9).CellStyle.BackColor = Color.LightYellow

        grid.VisibleColumns(0).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(2).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(3).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(4).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(5).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(6).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(7).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(8).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(10).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(11).CellStyle.BackColor = Color.LightYellow
        grid.VisibleColumns(12).CellStyle.BackColor = Color.LightYellow

    End Sub

    Private Sub onOffButton()
        Dim ls_SQL As String = ""
        Dim ls_HT As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "  SELECT * from Tally_Master " & vbCrLf & _
                     "  WHERE ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "' " & vbCrLf & _
                     "  AND AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf & _
                     "  AND ForwarderID = '" & Trim(cboForwarder.Text) & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Try
                    btnPrintTally.Enabled = True
                    btnSendEDI.Enabled = True
                    btnImportEDI.Enabled = True
                Catch ex As Exception

                End Try
            Else
                btnPrintTally.Enabled = False
                btnSendEDI.Enabled = False
                btnImportEDI.Enabled = False
            End If
            sqlConn.Close()


        End Using
    End Sub

    Private Sub up_fillcombopackinglist(ByVal pAffiliateID As String, ByVal pForwarderID As String)
        Dim ls_sql As String

        ls_sql = ""

        'Shipping No
        ls_sql = "SELECT DISTINCT [Shipping No] = RTRIM(ShippingInstructionNo) FROM ShippingInstruction_Master WHERE AffiliateID = '" & pAffiliateID & "' and ForwarderID = '" & pForwarderID & "'" & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboShippingNo
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Shipping No")
                .Columns(0).Width = 100

                .TextField = "Shipping No"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_fillcombocreateupdate()
        Dim ls_sql As String

        ls_sql = ""

        'Shipping No
        ls_sql = "SELECT 'CREATE' As CreateUpdate UNION ALL SELECT 'UPDATE' As CreateUpdate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboCreate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("CreateUpdate")
                .Columns(0).Width = 100
                .SelectedIndex = 0

                .TextField = "CreateUpdate"
                .DataBind()
            End With

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_SaveData()
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Dim pIsUpdate As Boolean

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ExchangeRate")

                    ls_SQL = "SELECT * FROM dbo.ShippingInstruction_Master where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If Trim(cboCreate.Text) = "CREATE" Then
                        If pIsUpdate = True Then
                            'Update
                            ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & _
                                     " SET     ShippingInstructionDate= '" & Convert.ToDateTime(dtShippingDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         BLNO = '" & Trim(txtBLNo.Text) & "'," & vbCrLf & _
                                     "         TermDelivery = '" & Trim(txtterm.Text) & "', " & vbCrLf & _
                                     "         Freight = '" & Trim(cbofreight.Text) & "', " & vbCrLf & _
                                     "         Via = '" & Trim(txtVia.Text) & "'," & vbCrLf & _
                                     "         Measurement = '" & Trim(txtmeasurement.Text) & "'," & vbCrLf & _
                                     "         TotalPallet = '" & Trim(txtTotalPallet.Text) & "'," & vbCrLf & _
                                     "         GrossWeight = '" & Trim(txtGrossWeight.Text) & "'," & vbCrLf & _
                                     "         TypeOfService = '" & Trim(cboService.Text) & "'," & vbCrLf & _
                                     "         ShippingLineS = '" & Trim(txtShippingLine.Text) & "', " & vbCrLf & _
                                     "         NamaKapalS = '" & Trim(txtVoyage.Text) & "', " & vbCrLf & _
                                     "         VesselS = '" & Trim(txtVessel.Text) & "', " & vbCrLf & _
                                     "         BLDate = '" & Convert.ToDateTime(dtBLDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         Remarks = '" & Trim(txtRemarks.Text) & "', " & vbCrLf & _
                                     "         UpdateDate = GETDATE(), " & vbCrLf & _
                                     "         UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                     "         where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                            ls_MsgID = "1002"

                        ElseIf pIsUpdate = False Then
                            'Insert
                            ls_SQL = " INSERT INTO dbo.ShippingInstruction_Master " & _
                                        "(AffiliateID, ForwarderID, ShippingInstructionNo, ShippingInstructionDate, BLNo, BLDate, EntryDate, EntryUser, TallyCls,ETDVendor, ETDPort, ETAPort, ETAFactory, termdelivery,freight,VIA,Measurement,TotalPallet,GrossWeight,TypeOfService,ShippingLineS,NamaKapalS,VesselS,Remarks)" & _
                                        " VALUES ('" & Trim(cboAffiliateCode.Text).Trim & "'," & vbCrLf & _
                                        " '" & Trim(cboForwarder.Text).Trim & "'," & vbCrLf & _
                                        " '" & Trim(cboShippingNo.Text) & "'," & vbCrLf & _
                                        " '" & Convert.ToDateTime(dtShippingDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Trim(txtBLNo.Text) & "'," & vbCrLf & _
                                        " '" & Convert.ToDateTime(dtBLDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " GETDATE()," & vbCrLf & _
                                        " '" & Session("UserID").ToString & "','0', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etdvendor.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etdport.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etaport.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etafactory.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Trim(txtterm.Text) & "', " & vbCrLf & _
                                        " '" & Trim(cbofreight.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtVia.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtmeasurement.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtTotalPallet.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtGrossWeight.Text) & "', " & vbCrLf & _
                                        " '" & Trim(cboService.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtShippingLine.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtVoyage.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtVessel.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtRemarks.Text) & "' " & vbCrLf & _
                                        ")" & vbCrLf
                            ls_MsgID = "1001"
                        End If

                    ElseIf Trim(cboCreate.Text) = "UPDATE" Then
                        If pIsUpdate = True Then
                            'Update
                            ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & _
                                     " SET     ShippingInstructionDate= '" & Convert.ToDateTime(dtShippingDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         BLNO = '" & Trim(txtBLNo.Text) & "'," & vbCrLf & _
                                     "         BLDate = '" & Convert.ToDateTime(dtBLDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         UpdateDate = GETDATE(), " & vbCrLf & _
                                     "         UpdateUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                                     "         ETDVendor = '" & Convert.ToDateTime(etdvendor.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         ETDPort = '" & Convert.ToDateTime(etdport.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         ETAPort = '" & Convert.ToDateTime(etaport.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         ETAFactory = '" & Convert.ToDateTime(etafactory.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                     "         Via = '" & Trim(txtVia.Text) & "'," & vbCrLf & _
                                     "         Measurement = '" & Trim(txtmeasurement.Text) & "'," & vbCrLf & _
                                     "         TotalPallet = '" & Trim(txtTotalPallet.Text) & "'," & vbCrLf & _
                                     "         GrossWeight = '" & Trim(txtGrossWeight.Text) & "'," & vbCrLf & _
                                     "         TypeOfService = '" & Trim(cboService.Text) & "'," & vbCrLf & _
                                     "         TermDelivery = '" & Trim(txtterm.Text) & "', " & vbCrLf & _
                                     "         ShippingLineS = '" & Trim(txtShippingLine.Text) & "', " & vbCrLf & _
                                     "         NamaKapalS = '" & Trim(txtVoyage.Text) & "', " & vbCrLf & _
                                     "         VesselS = '" & Trim(txtVessel.Text) & "', " & vbCrLf & _
                                     "         Remarks = '" & Trim(txtRemarks.Text) & "', " & vbCrLf & _
                                     "         Freight = '" & Trim(cbofreight.Text) & "'" & vbCrLf & _
                                     "         where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                            ls_MsgID = "1002"
                        End If
                    End If

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    sqlTran.Commit()

                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblErrMsg.Text

        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Sub

    Private Sub up_SENDINV()
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Dim pIsUpdate As Boolean

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ExchangeRate")

                    ls_SQL = "SELECT * FROM dbo.ShippingInstruction_Master where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If pIsUpdate = True Then
                        'Update
                        ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & _
                                 " SET     sendInvoice = '1', TallyCls = '1' " & vbCrLf & _
                                 "         where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                        ls_MsgID = "1013"
                        grid.JSProperties("cpSending") = "ALREADY SEND INVOICE"
                    Else
                        ls_MsgID = "6021"
                    End If

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    sqlTran.Commit()

                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblErrMsg.Text

        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Sub

    Private Sub up_SENDTALLY()
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Dim pIsUpdate As Boolean

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ExchangeRate")

                    ls_SQL = "SELECT * FROM dbo.ShippingInstruction_Master where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If pIsUpdate = True Then
                        'Update
                        ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & _
                                 " SET     ExcelCls = '1'  " & vbCrLf & _
                                 "         where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                        ls_MsgID = "1011"
                        grid.JSProperties("cpSending") = "ALREADY SEND TALLY"
                    Else
                        ls_MsgID = "6021"
                    End If


                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    sqlTran.Commit()

                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblErrMsg.Text

        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Sub

    Private Sub up_Delete()
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            'Dim pIsUpdate As Boolean

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ExchangeDelete")

                    Dim sqlComm As New SqlCommand '(ls_SQL, sqlConn, sqlTran)

                    ls_SQL = " DELETE dbo.ShippingInstruction_Master where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'" & vbCrLf & _
                             " DELETE dbo.ShippingInstruction_Detail where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"
                    ls_MsgID = "1003"
                    grid.JSProperties("cpSending") = "Data Deleted Successfully!"


                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    sqlTran.Commit()

                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblErrMsg.Text

        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Sub

    Private Sub up_sendshipping()
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Dim pIsUpdate As Boolean
            Dim pTally As Boolean

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("ExchangeRate")

                    ls_SQL = "SELECT * FROM dbo.ShippingInstruction_Master where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    'cek sudah ada data tally??
                    ls_SQL = "SELECT * FROM dbo.Tally_master where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                    Dim sqlComm1 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlRdr1 As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr1.Read Then
                        pTally = True
                    Else
                        pTally = False
                    End If
                    sqlRdr1.Close()
                    'cek sudah ada data tally??

                    If pTally = True Then
                        If pIsUpdate = True Then
                            'Update
                            ls_SQL = " UPDATE dbo.ShippingInstruction_Master " & _
                                     " SET     TallyCls = '1'" & _
                                     " where AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' AND ForwarderID = '" & Trim(cboForwarder.Text) & "' AND ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"

                            ls_MsgID = "1012"
                            grid.JSProperties("cpSending") = "ALREADY SEND SHIPPING"
                        Else
                            ls_MsgID = "6021"
                        End If
                    Else
                        ls_MsgID = "6021"
                    End If
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    sqlTran.Commit()

                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
            grid.JSProperties("cpMessage") = lblErrMsg.Text

        Catch ex As Exception
            Me.lblErrMsg.Visible = True
            Me.lblErrMsg.Text = ex.Message.ToString
        End Try
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 0
        Dim pIsUpdate As Boolean
        Dim ls_Affiliate As String = "", ls_Forwarder As String = "", ls_ShippingNo As String = "", ls_ShippingDate As String = ""
        Dim ls_OrderNo As String = "", ls_PartNo As String = ""
        Dim ls_PartName As String = "", ls_UOM As String = ""
        Dim ls_QtyBox As String = "", ls_GoodReceivingQty As String = "", ls_ShippingQty As String = ""
        Dim ls_BoxQty As String = "", ls_SupplierCode As String = "", ls_SupplierName As String = ""
        Dim ls_ETDPort As String = ""
        Dim ls_AdaData As String = ""
        Dim ls_AdaDataShipping1 As String = ""
        Dim ls_AdaDataShipping2 As String = ""
        Dim ls_AdaDataShipping3 As String = ""
        Dim ls_SJNo As String = ""
        Dim ls_BLNo As String = ""
        Dim ls_BLDate As String = ""
        Dim ls_LabelNo As String = ""
        Dim pilih As Boolean = False
		Dim ls_Price As String = ""
        Session.Remove("PriceAda")
        Session.Remove("dataExist")
        Session.Remove("YA010IsSubmit")
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("OrderNo")

                If grid.VisibleRowCount = 0 Then
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowPager, False, False)
                    Exit Sub
                End If

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("YA010Msg") = lblErrMsg.Text
                    Exit Sub
                End If

                Dim a As Integer
                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1
                    ls_Active = (e.UpdateValues(iLoop).NewValues("Act").ToString())
                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"
                    If ls_Active = "1" Then
                        pilih = True
                        If CDbl(e.UpdateValues(iLoop).NewValues("Price")) = 0 Then
                            ls_MsgID = "6201"
                            Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                            Session("YA010Msg") = lblErrMsg.Text
                            Session("PriceAda") = "6201"
                            Exit Sub
                        End If
                    End If
                Next

                '=====SAVE MASTER=====                
                If Session("PriceAda") Is Nothing Then
                    If pilih = False Then
                        Exit Sub
                    End If
                End If
                ls_SQL = "Select Top 1 * From ShippingInstruction_Master Where ShippingInstructionNo = '" & Trim(cboShippingNo.Text) & "'"
                Dim sqlComm_Exist As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm_Exist = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                Dim sqlRdr_Exist As SqlDataReader = sqlComm_Exist.ExecuteReader()

                Dim dataExisting As Boolean
                If sqlRdr_Exist.Read Then
                    dataExisting = True
                Else
                    dataExisting = False
                End If
                sqlRdr_Exist.Close()

                If dataExisting = True Then
                    ls_MsgID = "6018"
                    Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("dataExist") = "6018"
                    Exit Sub
                End If

                'Call up_SaveData()
                ls_SQL = " INSERT INTO dbo.ShippingInstruction_Master " & _
                                        "(AffiliateID, ForwarderID, ShippingInstructionNo, ShippingInstructionDate, BLNo, BLDate, EntryDate, EntryUser, TallyCls,ETDVendor, ETDPort, ETAPort, ETAFactory, termdelivery,freight,VIA,Measurement,TotalPallet,GrossWeight,TypeOfService,ShippingLineS,NamaKapalS,VesselS,Remarks)" & _
                                        " VALUES ('" & Trim(cboAffiliateCode.Text).Trim & "'," & vbCrLf & _
                                        " '" & Trim(cboForwarder.Text).Trim & "'," & vbCrLf & _
                                        " '" & Trim(cboShippingNo.Text) & "'," & vbCrLf & _
                                        " '" & Convert.ToDateTime(dtShippingDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Trim(txtBLNo.Text) & "'," & vbCrLf & _
                                        " '" & Convert.ToDateTime(dtBLDate.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " GETDATE()," & vbCrLf & _
                                        " '" & Session("UserID").ToString & "','0', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etdvendor.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etdport.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etaport.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Convert.ToDateTime(etafactory.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                        " '" & Trim(txtterm.Text) & "', " & vbCrLf & _
                                        " '" & Trim(cbofreight.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtVia.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtmeasurement.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtTotalPallet.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtGrossWeight.Text) & "', " & vbCrLf & _
                                        " '" & Trim(cboService.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtShippingLine.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtVoyage.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtVessel.Text) & "', " & vbCrLf & _
                                        " '" & Trim(txtRemarks.Text) & "' " & vbCrLf & _
                                        ")" & vbCrLf
                Dim sqlComm2 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm2 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm2.ExecuteNonQuery()
                sqlComm2.Dispose()
                '=====================

                For iLoop = 0 To a - 1

                    ls_Active = (e.UpdateValues(iLoop).NewValues("Act").ToString())
                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                    ls_Affiliate = Trim(cboAffiliateCode.Text)
                    ls_Forwarder = Trim(cboForwarder.Text)
                    ls_ShippingNo = Trim(cboShippingNo.Text)
                    ls_ShippingDate = Convert.ToDateTime(dtShippingDate.Value).ToString("yyyy-MM-dd")
                    ls_BLNo = Trim(txtBLNo.Text)
                    ls_BLDate = Convert.ToDateTime(dtBLDate.Value).ToString("yyyy-MM-dd")
                    ls_BLNo = Trim(txtBLNo.Text)
                    ls_BLDate = Convert.ToDateTime(dtBLDate.Value).ToString("yyyy-MM-dd")
                    ls_OrderNo = Trim(e.UpdateValues(iLoop).NewValues("OrderNo").ToString())
                    ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
                    ls_PartName = Trim(e.UpdateValues(iLoop).NewValues("PartName").ToString())
                    ls_UOM = Left(e.UpdateValues(iLoop).NewValues("UnitCls"), 2)
                    ls_QtyBox = Trim(e.UpdateValues(iLoop).NewValues("QtyBox").ToString())
                    ls_GoodReceivingQty = Trim(e.UpdateValues(iLoop).NewValues("GoodRecQty").ToString())
                    ls_ShippingQty = Trim(e.UpdateValues(iLoop).NewValues("ShippingQty").ToString())
                    ls_BoxQty = Trim(e.UpdateValues(iLoop).NewValues("BoxQty").ToString())
                    ls_SupplierCode = Trim(e.UpdateValues(iLoop).NewValues("SupplierID").ToString())
                    ls_SupplierName = Trim(e.UpdateValues(iLoop).NewValues("SuratJalanNo").ToString())
                    ls_ETDPort = Format(e.UpdateValues(iLoop).NewValues("ETDPort"), "yyyy-MM-dd")
                    ls_SJNo = Trim(e.UpdateValues(iLoop).NewValues("SuratJalanNo").ToString())
                    ls_AdaData = Trim(e.UpdateValues(iLoop).NewValues("AdaData").ToString())
                    ls_LabelNo = Trim(e.UpdateValues(iLoop).NewValues("LabelNo").ToString())
                    ls_Price = Trim(e.UpdateValues(iLoop).NewValues("Price").ToString())

                    Dim sqlstring As String
                    sqlstring = "SELECT * FROM ShippingInstruction_Detail WHERE AffiliateID ='" & Trim(ls_Affiliate) & "' " & vbCrLf & _
                                " AND ForwarderID = '" & Trim(ls_Forwarder) & "' AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                " AND SuratJalanNo = '" & Trim(ls_SJNo) & "' AND Orderno = '" & Trim(ls_OrderNo) & "' " & vbCrLf & _
                                " AND SupplierID = '" & Trim(ls_SupplierCode) & "' AND BoxNo = '" & ls_LabelNo & "'"

                    Dim sqlComm As New SqlCommand(sqlstring, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If ls_Active = "1" Then
                        If pIsUpdate = False Then
                            'INSERT DATA
                            ls_SQL = " 	INSERT INTO dbo.ShippingInstruction_Detail " & vbCrLf & _
                                     " 	        (AffiliateID, ForwarderID, PartNo, ShippingInstructionNo, SuratJalanNo, ETDPort, OrderNo, UOM, BoxNo, QtyBox, " & vbCrLf & _
                                     " 	        GoodReceivingQty, ShippingQty, BoxQty, SupplierID, EntryDate, EntryUser, POMOQ, POQtyBox, Price) " & vbCrLf & _
                                     " 	VALUES  ( '" & Trim(ls_Affiliate) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_Forwarder) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_PartNo) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_ShippingNo) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_SJNo) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_ETDPort) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_OrderNo) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_UOM) & "', " & vbCrLf & _
                                     " 	          '" & Trim(ls_LabelNo) & "', " & vbCrLf & _
                                     " 	          '" & ls_QtyBox & "', " & vbCrLf & _
                                     " 	          '" & ls_GoodReceivingQty & "', " & vbCrLf & _
                                     " 	          '" & ls_ShippingQty & "', " & vbCrLf & _
                                     " 	          '" & ls_BoxQty & "', " & vbCrLf & _
                                     " 	          '" & ls_SupplierCode & "', " & vbCrLf & _
                                     " 	          GETDATE(), " & vbCrLf & _
                                     " 	          '" & Session("UserID").ToString & "', " & vbCrLf & _
                                     " 	          '" & uf_GetMOQ(Trim(ls_OrderNo), Trim(ls_PartNo), ls_SupplierCode, Trim(ls_Affiliate), Trim(ls_Forwarder)) & "', " & vbCrLf & _
                                     " 	          '" & uf_GetQtybox(Trim(ls_OrderNo), Trim(ls_PartNo), ls_SupplierCode, Trim(ls_Affiliate), Trim(ls_Forwarder)) & "', " & vbCrLf & _
                                     "            '" & ls_Price & "' " & vbCrLf & _
                                     " 	        ) "
                            ls_MsgID = "1001"
                        Else
                            'ls_SQL = " 	UPDATE dbo.ShippingInstruction_Detail " & vbCrLf & _
                            '         " 	   SET ETDPort = '" & Trim(ls_ETDPort) & "' , " & vbCrLf & _
                            '         " 	       OrderNo = '" & Trim(ls_OrderNo) & "' , " & vbCrLf & _
                            '         " 	       UOM = '" & Trim(ls_UOM) & "' , " & vbCrLf & _
                            '         " 	       QtyBox = '" & ls_QtyBox & "' , " & vbCrLf & _
                            '         " 	       GoodReceivingQty = GoodReceivingQty + '" & ls_GoodReceivingQty & "' , " & vbCrLf & _
                            '         " 	       ShippingQty =  ShippingQty + '" & ls_ShippingQty & "' , " & vbCrLf & _
                            '         " 	       BoxQty = BoxQty + '" & ls_BoxQty & "' , " & vbCrLf & _
                            '         " 	       SupplierID = '" & ls_SupplierCode & "' , " & vbCrLf & _
                            '         " 	       UpdateDate = GETDATE(), " & vbCrLf & _
                            '         " 	       UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                            '         " 	 WHERE AffiliateID ='" & Trim(ls_Affiliate) & "' AND BoxNo = '" & ls_LabelNo & "' AND ForwarderID = '" & Trim(ls_Forwarder) & "' AND PartNo = '" & Trim(ls_PartNo) & "' AND ShippingInstructionNo = '" & Trim(ls_ShippingNo) & "' AND SuratJalanNo = '" & Trim(ls_SJNo) & "'"
                            'ls_MsgID = "1002"
                        End If

                    ElseIf ls_Active = "0" And pIsUpdate = True And ls_AdaData = "1" Then
                        ls_SQL = "  DELETE from dbo.ShippingInstruction_Detail" & vbCrLf & _
                                 "  WHERE AffiliateID = '" & Trim(ls_Affiliate) & "'" & vbCrLf & _
                                 "  AND ForwarderID = '" & Trim(ls_Forwarder) & "' " & vbCrLf & _
                                 "  AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                 "  AND ShippingInstructionNo = '" & Trim(ls_ShippingNo) & "' AND SuratJalanNo = '" & Trim(ls_SJNo) & "' AND BoxNo = '" & ls_LabelNo & "'"
                        ls_MsgID = "1003"

                    ElseIf ls_Active = "0" And pIsUpdate = False Then
                        lblErrMsg.Text = "[ Please give a checkmark to save data ! ] "
                        Session("YA010Msg") = lblErrMsg.Text
                        Exit Sub
                    End If

                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

NextLoop:
                Next iLoop

                If Left(ls_ShippingNo, 2) = "EA" Then
                    ls_SQL = "update ms_affiliate set seqno = " & Mid(Trim(cboShippingNo.Text), 6, 4) & " where AffiliateID = '" & Trim(ls_Affiliate) & "'"
                    Dim sqlComm22 As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm22.ExecuteNonQuery()
                    sqlComm22.Dispose()
                    Session("PriceAda") = "ada"
                End If

                sqlTran.Commit()

            End Using

            sqlConn.Close()
        End Using

        Call ColorGrid()
        Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.InformationMessage)
        Session("YA010Msg") = lblErrMsg.Text
        lblErrMsg.Visible = True
        Session("YA010IsSubmit") = "true"
    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "RowNo" Or e.Column.FieldName = "ETDPort" Or e.Column.FieldName = "OrderNo" _
            Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" Or e.Column.FieldName = "UnitClsDesc" _
            Or e.Column.FieldName = "QtyBox" Or e.Column.FieldName = "GoodRecQty" Or e.Column.FieldName = "BoxQty" _
            Or e.Column.FieldName = "SupplierID" Or e.Column.FieldName = "SupplierName" Or e.Column.FieldName = "LabelNo" _
            Or e.Column.FieldName = "SuratJalanNo" Or e.Column.FieldName = "ShippingQty" Or e.Column.FieldName = "UnitCls" _
            ) And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If

        Call ColorGrid()
    End Sub

    Private Sub btnsubmenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsubmenu.Click

        Session.Remove("SHAFFILIATEID")
        Session.Remove("SHORDERNO")
        Session.Remove("SHSUPPLIERID")
        Session.Remove("SHFWD")

        Session.Remove("SHGENERAL")
        Session.Remove("isSJ")

        If btnsubmenu.Text = "BACK" And Session("GOTOStatus") = "6" Then
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        ElseIf btnsubmenu.Text = "BACK" And Session("GOTOStatus") = "enam" Then
            Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
        ElseIf btnsubmenu.Text = "BACK" And Session("SHAFFILIATEID") <> "" Then
            Response.Redirect("~/DeliveryExport/DeliveryToAffListExport.aspx")
        Else
            Response.Redirect("~/MainMenu.aspx")
        End If
        Session.Remove("GOTOStatus")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 3, False, clsAppearance.PagerMode.ShowAllRecord, False)
            Dim ls_MsgID As String

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "SendEDI"
                    If uf_Approve() = 1 Then
                        Call clsMsg.DisplayMessage(lblErrMsg, "2005", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                        grid.JSProperties("cpButton") = "1"
                    Else
                        Call clsMsg.DisplayMessage(lblErrMsg, "8003", clsMessage.MsgType.ErrorMessage, errMsg)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                        grid.JSProperties("cpButton") = "0"
                    End If
                Case "load"
                    Session.Remove("PriceAda")
                    If cboCreate.Text = "CREATE" Then
                        Call up_GridLoad()
                    Else
                        Call up_GridLoadUpdate()
                    End If
                Case "loadaftersubmit"
                    Dim pilih As Boolean = False
                    Dim ls_Active As String

                    If Session("PriceAda") = "6201" Then
                        Call clsMsg.DisplayMessage(lblErrMsg, "6201", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                        'btnSave.JSProperties("pCreate") = cboCreate.Text
                        Exit Sub
                    ElseIf Session("dataExist") = "6018" Then
                        Call clsMsg.DisplayMessage(lblErrMsg, "6018", clsMessage.MsgType.ErrorMessage)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                        'btnSave.JSProperties("pCreate") = cboCreate.Text
                        Exit Sub
                    ElseIf Session("YA010IsSubmit") = "true" Then
                        Call clsMsg.DisplayMessage(lblErrMsg, "1001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                        'btnSave.JSProperties("pCreate") = cboCreate.Text
                        Exit Sub
                    Else
                        If Session("PriceAda") Is Nothing Then
                            For i = 0 To grid.VisibleRowCount - 1
                                ls_Active = grid.GetRowValues(i, "Act") '(e.UpdateValues(iLoop).NewValues("Act").ToString())
                                If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"
                                If ls_Active = "1" Then
                                    pilih = True
                                End If
                            Next
                            If pilih = False Then
                                ls_MsgID = "6010"
                                Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                                'Session("YA010Msg") = lblErrMsg.Text
                                grid.JSProperties("cpMessage") = lblErrMsg.Text
                                Exit Sub
                            End If
                            'ElseIf Session("PriceAda") <> "" Then                            
                            '    For i = 0 To grid.VisibleRowCount - 1
                            '        ls_Active = grid.GetRowValues(i, "Act") '(e.UpdateValues(iLoop).NewValues("Act").ToString())
                            '        If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"
                            '        If ls_Active = "1" Then
                            '            If CDbl(grid.GetRowValues(i, "Price")) = 0 Then
                            '                ls_MsgID = "6201"
                            '                Call clsMsg.DisplayMessage(lblErrMsg, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                            '                'Session("YA010Msg") = lblErrMsg.Text
                            '                grid.JSProperties("cpMessage") = lblErrMsg.Text
                            '                Exit Sub
                            '            End If
                            '        End If
                            '    Next
                        End If
                    End If

                    Call up_SaveData()
                    Call up_GridLoadUpdate()
                    'Call up_GridLoad()
                    up_IsiInvoice(Trim(cboShippingNo.Text))
                    Call ColorGrid()
                    'cboCreate.Text = "UPDATE"
                    txtBLNo.Text = Session("BLNO")
                    etdvendor.Text = Session("etdvendor")
                    etdport.Text = Session("etdport")
                    etaport.Text = Session("etaport")
                    etafactory.Text = Session("etafactory")
                    txtSend.Text = Session("sending")
                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblErrMsg, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblErrMsg.Text
                    End If
                Case "SENDTALLY"
                    Call up_SENDTALLY()
                    Call up_GridLoadUpdate()
                    'Call up_GridLoad()
                    up_IsiInvoice(Trim(cboShippingNo.Text))
                Case "SENDSHIPPING"
                    Call up_sendshipping()
                    Call up_GridLoadUpdate()
                    'Call up_GridLoad()
                    up_IsiInvoice(Trim(cboShippingNo.Text))
                Case "SENDINV"
                    Call up_SENDINV()
                    Call up_GridLoadUpdate()
                    'Call up_GridLoad()
                    up_IsiInvoice(Trim(cboShippingNo.Text))
                Case "delete"
                    Call up_Delete()
                    up_IsiInvoice(Trim(cboShippingNo.Text))
                    Call up_GridLoadUpdate()
                Case "excel"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplateShippingInstruction.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:8", psERR)
                    End If
                    'Case "ImportEDI"
                    '    Call up_ImportEDI()
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblErrMsg, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                            ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "SI No." & cboShippingNo.Text.Trim & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\ProgressReport\Import\" & tempFile & "")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet

            ws = exl.Workbook.Worksheets(pSheetName)
            Dim irow As Integer = 0
            Dim icol As Integer = 0

            With ws
                '.Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 10).AutoFitColumns()

                '.Cells(8, 6, pData.Rows.Count + 7, 6).Style.Numberformat.Format = "dd-mmm-yy"

                '.Cells(8, 5, pData.Rows.Count + 7, 5).Style.Numberformat.Format = "#,##0"

                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 22)
                EpPlusDrawAllBorders(rgAll)

                'For irow = 0 To pData.Rows.Count - 1
                '    For icol = 1 To pData.Columns.Count
                '        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                '        If icol = 7 Or icol = 8 Or icol = 14 Or icol = 15 Or icol = 16 Or icol = 20 Or icol = 23 Or icol = 26 Or icol = 29 Then
                '            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "dd-mmm-yy"
                '        End If
                '        If icol = 11 Or icol = 13 Or icol = 18 Or icol = 19 Or icol = 21 Or icol = 28 Or icol = 30 Or icol = 25 Or icol = 34 Then
                '            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "#,##0"
                '        End If
                '    Next
                'Next

                'Dim rgAll As ExcelRange = .Cells(8, 1, irow + 8, 34)
                'EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

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

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim tentukanBoat As String = cboService.Text.Trim
            Dim tentukanTerm As String = cboterm.Text.Trim

            Dim PriceCls As String = 0

            ''If tentukanBoat = "FCL" Or tentukanBoat = "LCL" Then
            'If tentukanTerm = "FCA" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            '    PriceCls = "2"
            'ElseIf tentukanTerm = "FCA" Then
            '    PriceCls = "1"
            'End If

            'If tentukanTerm = "CIF" And (tentukanBoat = "LCL" Or tentukanBoat = "FCL") Then
            '    PriceCls = "4"
            'ElseIf tentukanTerm = "CIF" Then
            '    PriceCls = "3"
            'End If

            'If tentukanTerm = "DDU PASI" Then
            '    PriceCls = "5"
            'ElseIf tentukanTerm = "DDU Affiliate" Then
            '    PriceCls = "6"
            'ElseIf tentukanTerm = "EX-Work" Then
            '    PriceCls = "7"
            'ElseIf tentukanTerm = "FOB" Then
            '    PriceCls = "8"
            'End If

            PriceCls = uf_PriceCls(tentukanTerm)

            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()

                ls_sql = " SELECT  DISTINCT  " & vbCrLf & _
                          "     InvoiceNo = SHM.ShippingInstructionNo,   " & vbCrLf & _
                          "     Consignee = MA.ConsigneeCode,  " & vbCrLf & _
                          "     Buyer = MA.BuyerCode,  " & vbCrLf & _
                          "     Shipment = case when POM.ShipCls='A' then 'Air Freight' else 'Sea Freight' End,  " & vbCrLf & _
                          "     ShippingLine = ISNULL(SHM.ShippingLineS,''),    " & vbCrLf & _
                          "     Vessel = isnull(SHM.NamaKapalS,''),    " & vbCrLf & _
                          "     Voyage = isnull(SHM.VesselS,''),  " & vbCrLf & _
                          "     FromPort = MF.Port,  " & vbCrLf & _
                          "     VIA = ISNULL(SHM.Via,''),  " & vbCrLf & _
                          "     ToPort = MA.DestinationPort,  " & vbCrLf

                ls_sql = ls_sql + " 	ETD = Replace(CONVERT(CHAR(11), CONVERT(DATETIME, SHM.ETDPort), 104),'.',''),      " & vbCrLf & _
                                  "     ETA = Replace(CONVERT(CHAR(11), CONVERT(DATETIME, SHM.ETAPort), 104),'.',''),  " & vbCrLf & _
                                  "     OrderNo = RD.OrderNo,    " & vbCrLf & _
                                  "     --OriginalNo = RD.PONo,  " & vbCrLf & _
                                  "     PartNo = SDM.PartNo,    " & vbCrLf & _
                                  "     PartGroupName = isnull(PartGroupName,''),  " & vbCrLf & _
                                  "     --SDM.BoxNo,  " & vbCrLf & _
                                  "     CartonCount = SUM(SDM.BoxQty),  " & vbCrLf & _
                                  "     QtyBox = ISNULL(SDM.POQtyBox,MPM.QtyBox),    " & vbCrLf & _
                                  "     Quantity = ISNULL(SDM.POQtyBox,MPM.QtyBox) * SUM(SDM.BoxQty),    " & vbCrLf & _
                                  "     Net = MPM.NetWeight /1000, " & vbCrLf & _
                                  "     TotalNet = SUM(SDM.BoxQty) * (MPM.NetWeight /1000), " & vbCrLf

                ls_sql = ls_sql + " 	Price = isnull(SDM.Price,MPR.Price), " & vbCrLf & _
                                  " 	TotalAmount = Isnull(isnull(SDM.Price,MPR.Price),0) * (ISNULL(SDM.POQtyBox,MPM.QtyBox) * SUM(SDM.BoxQty)) " & vbCrLf & _
                                  "     FROM ShippingInstruction_Master SHM   " & vbCrLf & _
                                  "                     INNER JOIN ShippingInstruction_Detail SDM   " & vbCrLf & _
                                  "                     ON ltrim(SDM.ShippingInstructionNo) = ltrim(SHM.ShippingInstructionNo)     " & vbCrLf & _
                                  "                         AND ltrim(SDM.ForwarderID) = rtrim(SHM.ForwarderID)     " & vbCrLf & _
                                  "                         AND rtrim(SDM.AffiliateID) = rtrim(SHM.AffiliateID)  " & vbCrLf & _
                                  "                     LEFT JOIN PO_Master_Export POM ON POM.PONo = SDM.OrderNo     " & vbCrLf & _
                                  "                         AND POM.AffiliateID = SDM.AffiliateID    " & vbCrLf & _
                                  "                         AND POM.SupplierID = SDM.SupplierID  " & vbCrLf & _
                                  "                         AND POM.OrderNo1 = SDM.OrderNo       " & vbCrLf & _
                                  "                     LEFT JOIN ReceiveForwarder_Detail RD ON RD.SuratJalanNO = SDM.SuratJalanno     " & vbCrLf

                ls_sql = ls_sql + "                         AND RD.AffiliateID = SDM.AffiliateID     	  " & vbCrLf & _
                                  "                     AND RD.SupplierID = SDM.SupplierID     " & vbCrLf & _
                                  "         AND RD.OrderNO = SDM.OrderNo     " & vbCrLf & _
                                  "                         AND RD.PartNo = SDM.PartNo  " & vbCrLf & _
                                  "                     LEFT JOIN MS_Affiliate MA ON MA.AffiliateID = SDM.AffiliateID  " & vbCrLf & _
                                  "                     LEFT JOIN MS_Forwarder MF ON MF.ForwarderID = SHM.ForwarderID  " & vbCrLf & _
                                  "                     LEFT JOIN MS_Parts MP ON MP.PartNo = SDM.PartNo  " & vbCrLf & _
                                  "                     LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = SDM.PartNo   " & vbCrLf & _
                                  " 						AND MPM.AffiliateID = SDM.AffiliateID AND MPM.SupplierID = SDM.SupplierID  " & vbCrLf & _
                                  " 					LEFT JOIN MS_Price MPR ON MPR.PartNO = SDM.PartNo " & vbCrLf & _
                                  " 						AND MPR.AffiliateID = SDM.AffiliateID " & vbCrLf

                ls_sql = ls_sql + " 						AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) >= CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.EffectiveDate,'')), 112) " & vbCrLf & _
                                  " 						AND CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(SHM.ETDPort,'')), 112) between CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Startdate,'')), 112) and CONVERT(CHAR(8), CONVERT(DATETIME, ISNULL(MPR.Enddate,'')), 112)  " & vbCrLf & _
                                  " 						AND MPR.CurrCls = '02' AND MPR.PriceCls = '" & PriceCls & "' " & vbCrLf & _
                                  "                     Where SHM.ShippingInstructionNo = '" & cboShippingNo.Text.Trim & "'  " & vbCrLf & _
                                  "                           AND SHM.AffiliateID = '" & cboAffiliateCode.Text.Trim & "'  " & vbCrLf & _
                                  "                           AND SHM.ForwarderID = '" & cboForwarder.Text.Trim & "'  " & vbCrLf & _
                                  " 	GROUP BY SHM.ShippingInstructionNo, MA.ConsigneeCode, MA.BuyerCode, POM.ShipCls, SHM.ShippingLineS, " & vbCrLf & _
                                  " 			SHM.NamaKapalS, SHM.VesselS, MF.Port, SHM.Via, MA.DestinationPort, SHM.ETDPort, SHM.ETAPort, " & vbCrLf & _
                                  " 			RD.OrderNo, RD.PONo, SDM.PartNo, PartGroupName, MPM.QtyBox, MPM.NetWeight, MPR.Price, SDM.POQtyBox, SDM.price " & vbCrLf & _
                                  "  "


                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                Return dt
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub cboShippingNo_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxClasses.CallbackEventArgsBase) Handles cboShippingNo.Callback
        If cboCreate.Text = "UPDATE" Then
            If String.IsNullOrEmpty(e.Parameter) Then
                Return
            End If

            Dim ls_value1 As String = Split(e.Parameter, "|")(0)
            Dim ls_value2 As String = Split(e.Parameter, "|")(1)

            Call up_fillcombopackinglist(ls_value1, ls_value2)
        End If
    End Sub

    Private Sub up_ImportEDI()
        Dim pAffiliate As String = cboAffiliateCode.Text.ToString.Trim
        Dim pInvoiceNo As String = cboShippingNo.Text.ToString.Trim
        Dim result As String = String.Empty

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim ls_SQL As String = ""
            ls_SQL = "SELECT AF.ConsigneeCode, SM.AffiliateID, SM.ShippingInstructionNo " & vbCrLf & _
                "FROM ShippingInstruction_Master SM " & vbCrLf & _
                "INNER JOIN MS_Affiliate AF ON SM.AffiliateID = AF.AffiliateID " & vbCrLf & _
                "WHERE SM.AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                "AND SM.ShippingInstructionNo = '" & pInvoiceNo & "' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Dim ls_Time As Date = Now
                Dim ls_AffCode As String = Trim(ds.Tables(0).Rows(0)("ConsigneeCode"))
                Dim FileName As String = "INVOICE_DATA_" + Trim(ls_AffCode) + "_32G8_" + Left(Trim(pInvoiceNo), 10) + "_" + Format(ls_Time, "yyyyMMdd") + "_" + Format(ls_Time, "hhmm") + "_" + Format(ls_Time, "ss") + ".txt"
                Dim FilePath As String = Server.MapPath("~\ShippingInstruction\Import\" & FileName)

                ls_SQL = "EXEC sp_SendInvoiceEDI_ASN '" & pInvoiceNo & "', '" & pAffiliate & "'"

                Dim sqlDADetail As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim dsDetail As New DataSet
                sqlDADetail.Fill(dsDetail)
                If dsDetail.Tables(0).Rows.Count > 0 Then
                    For x = 0 To dsDetail.Tables(0).Rows.Count - 1
                        result += dsDetail.Tables(0).Rows(x)("a")
                        result += vbCrLf
                    Next

                    Dim fp As StreamWriter
                    fp = File.CreateText(FilePath)
                    For x = 0 To dsDetail.Tables(0).Rows.Count - 1
                        fp.WriteLine(dsDetail.Tables(0).Rows(x)("a"))
                    Next
                    fp.Close()

                End If

                'DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ShippingInstruction\Import\" & FileName)

                Response.Clear()
                Response.Buffer = True
                Response.AddHeader("content-disposition", "attachment;filename=" + FileName + "")
                Response.Charset = ""
                Response.ContentType = "application/text"
                Response.Output.Write(result)
                Response.Flush()
                Response.End()
            End If

            sqlConn.Close()
        End Using

    End Sub
#End Region

    Private Sub btnImportEDI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnImportEDI.Click
        Call up_ImportEDI()
    End Sub

    Private Function uf_GetMOQ(ByVal pOrderNo As String, ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pForwarderID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(a.POMOQ,b.MOQ) MOQ FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE OrderNo1 ='" + pOrderNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "' AND a.ForwarderID = '" + pForwarderID + "' "
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

    Private Function uf_GetQtybox(ByVal pOrderNo As String, ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String, ByVal pForwarderID As String) As Integer
        Dim Qty As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(a.POQtyBox,b.QtyBox) Qty FROM dbo.PO_Detail_Export a left join MS_PartMapping b on a.PartNo = b.PartNo and a.AffiliateID = b.AffiliateID and a.SupplierID = b.SupplierID " & vbCrLf & _
                     "WHERE OrderNo1 ='" + pOrderNo + "' AND a.PartNo = '" + pPartNo + "' AND a.SupplierID = '" + pSupplierID + "' AND a.AffiliateID = '" + pAffiliateID + "' AND a.ForwarderID = '" + pForwarderID + "' "
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                Qty = dt.Rows(0)("Qty")
            End If
        End Using
        Return Qty
    End Function

    Public Function uf_GetDataTable(ByVal Query As String, Optional ByVal pCon As SqlConnection = Nothing, Optional ByVal pTrans As SqlTransaction = Nothing) As DataTable
        Dim cmd As New SqlCommand(Query)
        If pTrans IsNot Nothing Then
            cmd.Transaction = pTrans
        End If
        If pCon IsNot Nothing Then
            cmd.Connection = pCon
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            Dim dt As New DataTable
            da.Fill(ds)
            Return ds.Tables(0)
        Else
            Using Cn As New SqlConnection(clsGlobal.ConnectionString)
                Cn.Open()
                cmd.Connection = Cn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New DataSet
                Dim dt As New DataTable
                da.Fill(ds)
                Return ds.Tables(0)
            End Using
        End If
    End Function
End Class