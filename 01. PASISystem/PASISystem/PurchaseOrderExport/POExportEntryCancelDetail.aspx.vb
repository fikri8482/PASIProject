Option Explicit On
Imports System.Data.SqlClient
Imports DevExpress.Web.ASPxGridView
Imports DevExpress.Web.ASPxEditors
Imports DevExpress.Web.ASPxPanel
Imports DevExpress.Web.ASPxRoundPanel
Imports System.Transactions
Imports System.Net
Imports System.Net.Mail
Imports DevExpress.Web.ASPxUploadControl
Imports System.IO
Imports System.Data.OleDb
Imports OfficeOpenXml
Imports System.Drawing

Public Class POExportEntryCancelDetail

#Region "DECLARATION"
    Inherits System.Web.UI.Page
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim menuID As String = "B02"
    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim pub_PONo As String, pub_Ship As String, pub_Commercial As String, pub_SupplierID As String
    Dim pub_Period As Date
    Dim pub_HeijunkaSttus As Boolean
    Dim smtpClient As String
    Dim portClient As String
    Dim usernameSMTP As String
    Dim PasswordSMTP As String
    Dim log As String = ""
    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""

    Dim serverPath As String
    Dim fullPath As String
    Dim flag As Boolean = True
    Dim pStatus As Boolean
    Dim pPeriod As Date
    Dim pCommercial As String
    Dim pPOEmergency As String
    Dim pShipBy As String
    Dim pAffiliateCode As String
    Dim pAffiliateName As String
    Dim pSupplierCode As String
    Dim pSupplierName As String
    Dim pOrderNo As String
    Dim pDeliveryCode As String
    Dim pDeliveryName As String
    Dim pETDVendor As String
    Dim pETDPort As String
    Dim pETAPort As String
    Dim pETAFactory As String
    Dim pPO As String
    Dim pConsignee As String
    Dim pSplitRefPONo As String
    Dim pSplitStatus As String

    Dim pFilter As String
    Dim pub_Param As String
    Dim pstatusInsert As String
    Dim ls_TampungError As Integer = 0

#End Region

#Region "FORM EVENTS"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim param As String = ""
            If (Not IsPostBack) AndAlso (Not IsCallback) Then

                Session("MenuDesc") = "CANCELATION LIST DETAIL"
                If IsNothing(Request.QueryString("prm")) Then
                    param = ""
                Else
                    param = Request.QueryString("prm").ToString()
                End If

                If param = "" Then
                    If Session("GOTOStatus") = "" Then
                        Call up_fillcombo()
                        Call ColorGrid()
                        Call up_GridLoad()
                        dtPeriodFrom.Value = Now
                        rdMonthly.Checked = True
                        rdrCom1.Checked = True
                        rdrShipBy2.Checked = True
                        dtETDVendor.Value = Now
                        dtETDPort.Value = Now
                        dtETAPort.Value = Now
                        dtETAFactory.Value = Now
                        lblInfo.Text = ""
                    End If

                ElseIf param <> "" And Session("GOTOStatus") = "1" Then
                    Call up_fillcombo()

                    param = Request.QueryString("prm").ToString
                    If param = "  'back'" Then
                        btnSubMenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session.Remove("LoadSupplier")
                            Session.Remove("SplitStatus")
                            lblInfo.Text = ""
                            pPeriod = Split(param, "|")(1)
                            pCommercial = Split(param, "|")(2)
                            pPOEmergency = Split(param, "|")(3)
                            pShipBy = Split(param, "|")(4)
                            pAffiliateCode = Split(param, "|")(5)
                            pAffiliateName = Split(param, "|")(6)
                            pSupplierCode = Split(param, "|")(7)
                            pSupplierName = Split(param, "|")(8)
                            pOrderNo = Split(param, "|")(9)
                            pETDVendor = Split(param, "|")(10)
                            pETDPort = Split(param, "|")(11)
                            pETAPort = Split(param, "|")(12)
                            pETAFactory = Split(param, "|")(13)
                            pDeliveryCode = Split(param, "|")(14)
                            pDeliveryName = Split(param, "|")(15)
                            pPO = Split(param, "|")(16)
                            pConsignee = Split(param, "|")(17)
                            pSplitRefPONo = Split(param, "|")(18)
                            pSplitStatus = Split(param, "|")(19)

                            Session("SplitReffPONo") = pSplitRefPONo

                            If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"

                            dtPeriodFrom.Value = pPeriod
                            If pCommercial = "1" Then
                                rdrCom1.Checked = True
                            Else
                                rdrCom2.Checked = True
                            End If

                            If pPOEmergency = "E" Then
                                rdEmergency.Checked = True
                            Else
                                rdMonthly.Checked = True
                            End If

                            If pShipBy = "B" Then
                                rdrShipBy2.Checked = True
                            Else
                                rdrShipBy3.Checked = True
                            End If

                            cboAffiliate.Text = pAffiliateCode
                            txtAffiliate.Text = pAffiliateName

                            txtOrderNo.Text = pOrderNo
                            dtETDVendor.Text = pETDVendor
                            dtETDPort.Text = pETDPort
                            dtETAPort.Text = pETAPort
                            dtETAFactory.Text = pETAFactory
                            cboDelLoc.Text = pDeliveryCode
                            txtDelLoc.Text = pDeliveryName
                            txtpono.Text = pPO
                            txtconsignee.Text = pConsignee

                            Session("LoadSupplier") = pSupplierCode
                            Session("SplitStatus") = pSplitStatus

                            pStatus = True

                            Call up_GridLoadUpdate()
                            Session("pCheckError") = "1"

                            Session("pFilter") = pFilter
                            Session.Remove("EmergencyUrl")
                            btnSubMenu.Text = "BACK"
                        End If
                    End If

                ElseIf param <> "" And Session("GOTOStatus") = "satu" Then
                    Call up_fillcombo()

                    param = Request.QueryString("prm").ToString
                    If param = "  'back'" Then
                        btnSubMenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            lblInfo.Text = ""
                            pPeriod = Split(param, "|")(1)
                            pCommercial = Split(param, "|")(2)
                            pPOEmergency = Split(param, "|")(3)
                            pShipBy = Split(param, "|")(4)
                            pAffiliateCode = Split(param, "|")(5)
                            pAffiliateName = Split(param, "|")(6)
                            pSupplierCode = Split(param, "|")(7)
                            pSupplierName = Split(param, "|")(8)
                            pOrderNo = Split(param, "|")(9)
                            pETDVendor = Split(param, "|")(10)
                            pETDPort = Split(param, "|")(11)
                            pETAPort = Split(param, "|")(12)
                            pETAFactory = Split(param, "|")(13)
                            pDeliveryCode = Split(param, "|")(14)
                            pDeliveryName = Split(param, "|")(15)
                            pPO = Split(param, "|")(16)
                            pConsignee = Split(param, "|")(17)
                            pSplitRefPONo = Split(param, "|")(18)
                            pSplitStatus = Split(param, "|")(19)

                            If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"

                            dtPeriodFrom.Value = pPeriod
                            If pCommercial = "1" Then
                                rdrCom1.Checked = True
                            Else
                                rdrCom2.Checked = True
                            End If

                            If pPOEmergency = "E" Then
                                rdEmergency.Checked = True
                            Else
                                rdMonthly.Checked = True
                            End If

                            If pShipBy = "B" Then
                                rdrShipBy2.Checked = True
                            Else
                                rdrShipBy3.Checked = True
                            End If

                            cboAffiliate.Text = pAffiliateCode
                            txtAffiliate.Text = pAffiliateName
                            txtconsignee.Text = pConsignee
                            txtOrderNo.Text = pOrderNo
                            dtETDVendor.Text = pETDVendor
                            dtETDPort.Text = pETDPort
                            dtETAPort.Text = pETAPort
                            dtETAFactory.Text = pETAFactory
                            cboDelLoc.Text = pDeliveryCode
                            txtDelLoc.Text = pDeliveryName
                            txtpono.Text = pPO
                            txtconsignee.Text = pConsignee

                            Session("LoadSupplier") = pSupplierCode
                            Session("SplitStatus") = pSplitStatus

                            pStatus = True

                            Call up_GridLoad()
                            Session("pCheckError") = "1"

                            Session("pFilter") = pFilter
                            Session.Remove("EmergencyUrl")
                            btnSubMenu.Text = "BACK"
                        End If
                    End If

                ElseIf param <> "" Then
                    Call up_fillcombo()

                    param = Request.QueryString("prm").ToString
                    If param = "  'back'" Then
                        btnSubMenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            lblInfo.Text = ""
                            pPeriod = Split(param, "|")(1)
                            pCommercial = Split(param, "|")(2)
                            pPOEmergency = Split(param, "|")(3)
                            pShipBy = Split(param, "|")(4)
                            pAffiliateCode = Split(param, "|")(5)
                            pAffiliateName = Split(param, "|")(6)
                            pSupplierCode = Split(param, "|")(7)
                            pSupplierName = Split(param, "|")(8)
                            pOrderNo = Split(param, "|")(9)
                            pETDVendor = Split(param, "|")(10)
                            pETDPort = Split(param, "|")(11)
                            pETAPort = Split(param, "|")(12)
                            pETAFactory = Split(param, "|")(13)
                            pDeliveryCode = Split(param, "|")(14)
                            pDeliveryName = Split(param, "|")(15)
                            pPO = Split(param, "|")(16)
                            pConsignee = Split(param, "|")(17)
                            pSplitRefPONo = Split(param, "|")(18)
                            pSplitStatus = Split(param, "|")(19)

                            If pAffiliateCode <> "" Then btnSubMenu.Text = "BACK"

                            dtPeriodFrom.Value = pPeriod
                            If pCommercial = "1" Then
                                rdrCom1.Checked = True
                            Else
                                rdrCom2.Checked = True
                            End If

                            If pPOEmergency = "E" Then
                                rdEmergency.Checked = True
                            Else
                                rdMonthly.Checked = True
                            End If

                            If pShipBy = "B" Then
                                rdrShipBy2.Checked = True
                            Else
                                rdrShipBy3.Checked = True
                            End If

                            cboAffiliate.Text = pAffiliateCode
                            txtAffiliate.Text = pAffiliateName
                            txtconsignee.Text = pConsignee
                            txtOrderNo.Text = pOrderNo
                            dtETDVendor.Text = pETDVendor
                            dtETDPort.Text = pETDPort
                            dtETAPort.Text = pETAPort
                            dtETAFactory.Text = pETAFactory
                            cboDelLoc.Text = pDeliveryCode
                            txtDelLoc.Text = pDeliveryName
                            txtpono.Text = pPO
                            txtconsignee.Text = pConsignee

                            Session("LoadSupplier") = pSupplierCode
                            Session("SplitStatus") = pSplitStatus

                            pStatus = True

                            Call up_GridLoadUpdate()

                            Session("pCheckError") = "1"

                            Session("pFilter") = pFilter
                            Session.Remove("EmergencyUrl")
                            btnSubMenu.Text = "BACK"
                        End If
                    End If
                End If
            End If

            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Auto, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub grid_BatchUpdate(ByVal sender As Object, ByVal e As DevExpress.Web.Data.ASPxDataBatchUpdateEventArgs) Handles grid.BatchUpdate
        Dim ls_SQL As String = "", ls_MsgID As String = ""
        Dim ls_Active As String = "", iLoop As Long = 0, iCheckLoop As Long = 0
        Dim pIsUpdate As Boolean
        Dim pIsUpdateMaster As Boolean
        Dim ls_PONo As String = "", ls_Affiliate As String = "", ls_Supplier As String = "", ls_PartNo As String = ""
        Dim ls_Week1 As Double = 0, ls_TotalPOQty As Double = 0, ls_MOQ As Double = 0
        Dim ls_PreviousForecast As Double = 0, ls_Forecast1 As Double = 0
        Dim ls_Forecast2 As Double = 0, ls_Forecast3 As Double = 0
        Dim ls_Variance As Double = 0, ls_VariancePercentage As Double = 0
        Dim ls_AdaData As String = ""
        Dim ls_error As String = ""
        Dim ls_FWD As String = ""
        Dim ls_OrderNo As String = ""
        Dim ls_qtybox As Double = 0
        Dim ls_CancelReffPONo As String = ""
        Dim ls_CancelReffQty As String = ""
        Dim ls_TOP As String = ""
        Dim sqlComm As New SqlCommand
        Dim a As Integer

        Session.Remove("ErrorData")
        Session.Remove("YA010Msg")

        a = e.UpdateValues.Count
        For iLoop = 0 To a - 1
            ls_Week1 = Trim(e.UpdateValues(iLoop).NewValues("Week1").ToString())
            ls_CancelReffQty = Trim(e.UpdateValues(iLoop).NewValues("CancelReffQty").ToString())
            ls_qtybox = Trim(e.UpdateValues(iLoop).NewValues("QtyBox").ToString())

            If ls_Week1 = "0" Then
                lblInfo.Text = "[ Please give a checkmark to save data ! ] "
                Session("YA010Msg") = lblInfo.Text
                Exit Sub
            End If
            If CDbl(ls_Week1) > CDbl(ls_CancelReffQty) Then
                lblInfo.Text = "[ Qty Cancel is bigger than Refrence PO Qty ! ] "
                Session("YA010Msg") = lblInfo.Text
                Exit Sub
            End If
            If (CDbl(ls_Week1) Mod CDbl(ls_qtybox)) <> 0 Then
                lblInfo.Text = "[ Qty Cancel must be multiple then Qty/Box ! ] "
                Session("YA010Msg") = lblInfo.Text
                Exit Sub
            End If
        Next

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryCancelDetail")
                Session.Remove("ErrorData")
                ls_TampungError = 0

                If grid.VisibleRowCount = 0 Then
                    Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 2, False, clsAppearance.PagerMode.ShowAllRecord, False, False)
                    Exit Sub
                End If

                If e.UpdateValues.Count = 0 Then
                    ls_MsgID = "6011"
                    Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                    Session("YA010Msg") = lblInfo.Text
                    Exit Sub
                End If

                ls_SQL = "DELETE FROM PO_Tampung_Detail_Export where PONo = '" & Trim(txtpono.Text) & "'"
                If txtOrderNo.Text <> "" Then ls_SQL = ls_SQL + " and SupplierID = '" & Replace(Trim(txtOrderNo.Text), Trim(txtpono.Text) & "-", "") & "'"

                Dim sqlComm2 As New SqlCommand
                sqlComm2 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm2.ExecuteNonQuery()
                sqlComm2.Dispose()

                'master
                Dim ls_EmergencyCls As String
                Dim ls_Commercial As String
                Dim ls_ShipCls As String

                If rdEmergency.Checked = True Then
                    ls_EmergencyCls = "E"
                ElseIf rdMonthly.Checked = True Then
                    ls_EmergencyCls = "M"
                End If

                If rdrCom1.Checked = True Then
                    ls_Commercial = "1"
                ElseIf rdrCom2.Checked = True Then
                    ls_Commercial = "0"
                End If

                If rdrShipBy2.Checked = True Then
                    ls_ShipCls = "B"
                ElseIf rdrShipBy3.Checked = True Then
                    ls_ShipCls = "A"
                End If

                'Insert dan Update
                a = e.UpdateValues.Count
                For iLoop = 0 To a - 1

                    ls_Active = (e.UpdateValues(iLoop).NewValues("AllowAccess").ToString())
                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                    ls_OrderNo = Trim(txtOrderNo.Text)
                    ls_PONo = Trim(txtpono.Text)
                    ls_FWD = Trim(cboDelLoc.Text)
                    ls_Affiliate = Trim(cboAffiliate.Text)
                    ls_Supplier = Trim(e.UpdateValues(iLoop).NewValues("SupplierID").ToString())
                    ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
                    ls_Week1 = Trim(e.UpdateValues(iLoop).NewValues("Week1").ToString())
                    ls_MOQ = Trim(e.UpdateValues(iLoop).NewValues("MOQ").ToString())
                    ls_qtybox = Trim(e.UpdateValues(iLoop).NewValues("QtyBox").ToString())
                    ls_TotalPOQty = Trim(e.UpdateValues(iLoop).NewValues("Week1").ToString())
                    ls_PreviousForecast = Trim(e.UpdateValues(iLoop).NewValues("PreviousForecast").ToString())
                    ls_Forecast1 = Trim(e.UpdateValues(iLoop).NewValues("Forecast1").ToString())
                    ls_Forecast2 = Trim(e.UpdateValues(iLoop).NewValues("Forecast2").ToString())
                    ls_Forecast3 = Trim(e.UpdateValues(iLoop).NewValues("Forecast3").ToString())
                    ls_Variance = Trim(e.UpdateValues(iLoop).NewValues("Variance").ToString())
                    ls_VariancePercentage = Trim(e.UpdateValues(iLoop).NewValues("VariancePercentage").ToString())
                    ls_AdaData = Trim(e.UpdateValues(iLoop).NewValues("AdaData").ToString())
                    ls_CancelReffPONo = Trim(e.UpdateValues(iLoop).NewValues("CancelReffPONo").ToString())
                    ls_CancelReffQty = Trim(e.UpdateValues(iLoop).NewValues("CancelReffQty").ToString())
                    ls_TOP = CDbl(ls_Week1) / CDbl(ls_qtybox)

                    Dim sqlstring As String
                    sqlstring = "SELECT * FROM PO_Detail_ExportCancel WHERE PONo ='" & Trim(ls_PONo) & "' AND OrderNo1 = '" & ls_OrderNo & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "'"
                    sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    sqlstring = "SELECT * FROM PO_Master_ExportCancel WHERE PONo ='" & Trim(ls_PONo) & "' AND OrderNo1 = '" & ls_OrderNo & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "'"
                    sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    Dim sqlRdr3 As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr3.Read Then
                        pIsUpdateMaster = True
                    Else
                        pIsUpdateMaster = False
                    End If

                    sqlRdr3.Close()
                    If ls_Active = "1" Then
                        If ls_TampungError = 0 Then
                            If pIsUpdateMaster = True Then
                                'Update
                                ls_SQL = " UPDATE dbo.PO_Master_ExportCancel " & _
                                         " SET     Period = '" & Convert.ToDateTime(dtPeriodFrom.Value).ToString("yyyy-MM-01") & "', " & vbCrLf & _
                                         "         EmergencyCls = '" & Trim(ls_EmergencyCls) & "'," & vbCrLf & _
                                         "         CommercialCls = '" & Trim(ls_Commercial) & "'," & vbCrLf & _
                                         "         ShipCls = '" & Trim(ls_ShipCls) & "'," & vbCrLf & _
                                         "         ForwarderID = '" & Trim(cboDelLoc.Text) & "'," & vbCrLf & _
                                         "         ETDVendor1 = '" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         ETDPort1 = '" & Convert.ToDateTime(dtETDPort.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         ETAPort1 = '" & Convert.ToDateTime(dtETAPort.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         ETAFactory1 = '" & Convert.ToDateTime(dtETAFactory.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         UpdateDate = GETDATE(), " & vbCrLf & _
                                         "         UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                         " WHERE PONo = '" & Trim(ls_PONo) & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND SupplierID = '" & Trim(ls_Supplier) & "'" & vbCrLf & _
                                         " AND OrderNo1 = '" & Trim(ls_OrderNo) & "'"


                                ls_MsgID = "1002"

                            ElseIf pIsUpdateMaster = False Then
                                'Insert
                                ls_SQL = "INSERT INTO PO_Master_ExportCancel( " & vbCrLf & _
                                    "PONo, AffiliateID, SupplierID, ForwarderID, Period, CommercialCls, " & vbCrLf & _
                                    "EmergencyCls, ShipCls, ErrorStatus, " & vbCrLf & _
                                    "OrderNo1, " & vbCrLf & _
                                    "ETDVendor1, " & vbCrLf & _
                                    "ETDPort1, " & vbCrLf & _
                                    "ETAPort1, " & vbCrLf & _
                                    "ETAFactory1, " & vbCrLf & _
                                    "OrderNo2, ETDVendor2, ETDPort2, ETAPort2, ETAFactory2, " & vbCrLf & _
                                    "OrderNo3, ETDVendor3, ETDPort3, ETAPort3, ETAFactory3, " & vbCrLf & _
                                    "OrderNo4, ETDVendor4, ETDPort4, ETAPort4, ETAFactory4, " & vbCrLf & _
                                    "OrderNo5, ETDVendor5, ETDPort5, ETAPort5, ETAFactory5, UploadDate, UploadUser, " & vbCrLf & _
                                    "PASISendToSupplierDate, PASISendToSupplierUser, SupplierApproveDate, SupplierApproveUser, " & vbCrLf & _
                                    "SupplierApprovePartialDate, SupplierApprovePartialUser, SupplierUnApproveDate, SupplierUnApproveUser, " & vbCrLf & _
                                    "PASIApproveDate, PASIApproveUser, EntryDate, EntryUser, UpdateDate, UpdateUser, PASISendToSupplierCls, " & vbCrLf & _
                                    "SupplierApprovalCls, ExcelCls, FinalApprovalCls, SplitReffPONo, SplitStatus) " & vbCrLf & _
                                    "SELECT PONo, AffiliateID, SupplierID, '" & Trim(cboDelLoc.Text) & "' ForwarderID, Period, '" & Trim(ls_Commercial) & "' CommercialCls, " & vbCrLf & _
                                    "'" & Trim(ls_EmergencyCls) & "' EmergencyCls, '" & Trim(ls_ShipCls) & "' ShipCls, ErrorStatus, " & vbCrLf & _
                                    "'" & txtOrderNo.Text & "' OrderNo1, " & vbCrLf & _
                                    "'" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "' ETDVendor1, " & vbCrLf & _
                                    "'" & Convert.ToDateTime(dtETDPort.Value).ToString("yyyy-MM-dd") & "' ETDPort1, " & vbCrLf & _
                                    "'" & Convert.ToDateTime(dtETAPort.Value).ToString("yyyy-MM-dd") & "' ETAPort1, " & vbCrLf & _
                                    "'" & Convert.ToDateTime(dtETAFactory.Value).ToString("yyyy-MM-dd") & "' ETAFactory1, " & vbCrLf & _
                                    "OrderNo2, ETDVendor2, ETDPort2, ETAPort2, ETAFactory2, " & vbCrLf & _
                                    "OrderNo3, ETDVendor3, ETDPort3, ETAPort3, ETAFactory3, " & vbCrLf & _
                                    "OrderNo4, ETDVendor4, ETDPort4, ETAPort4, ETAFactory4, " & vbCrLf & _
                                    "OrderNo5, ETDVendor5, ETDPort5, ETAPort5, ETAFactory5, UploadDate, UploadUser, " & vbCrLf & _
                                    "PASISendToSupplierDate, PASISendToSupplierUser, SupplierApproveDate, SupplierApproveUser, " & vbCrLf & _
                                    "SupplierApprovePartialDate, SupplierApprovePartialUser, SupplierUnApproveDate, SupplierUnApproveUser, " & vbCrLf & _
                                    "PASIApproveDate, PASIApproveUser, GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, NULL UpdateDate, NULL UpdateUser, PASISendToSupplierCls, " & vbCrLf & _
                                    "SupplierApprovalCls, ExcelCls, FinalApprovalCls, '" & ls_CancelReffPONo & "' SplitReffPONo, '" & Session("SplitStatus") & "' SplitStatus " & vbCrLf & _
                                    "FROM PO_Master_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf

                                Select Case Session("SplitStatus")
                                    Case "2", "3", "4", "5", "6"
                                        ls_SQL = ls_SQL + "UPDATE PO_Master_ExportCancel " & vbCrLf & _
                                            "SET ExcelCls = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & txtOrderNo.Text & "' " & vbCrLf

                                        Select Case Session("SplitStatus")
                                            Case "2", "3"
                                                ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                    "SET ExcelCls = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                    "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf

                                                If Session("SplitStatus") = "2" Then
                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET PASISendToSupplierDate = GETDATE(), PASISendToSupplierUser = '" & Session("UserID").ToString & "', UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf
                                                End If

                                                If Session("SplitStatus") = "3" Then
                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET SupplierApproveDate = CASE WHEN SupplierApproveDate IS NOT NULL THEN GETDATE() ELSE NULL END, " & vbCrLf & _
                                                        "SupplierApproveUser = CASE WHEN SupplierApproveUser IS NOT NULL THEN '" & Session("UserID").ToString & "' ELSE NULL END, " & vbCrLf & _
                                                        "SupplierApprovePartialDate = CASE WHEN SupplierApprovePartialDate IS NOT NULL THEN GETDATE() ELSE NULL END, " & vbCrLf & _
                                                        "SupplierApprovePartialUser = CASE WHEN SupplierApprovePartialUser IS NOT NULL THEN '" & Session("UserID").ToString & "' ELSE NULL END, " & vbCrLf & _
                                                        "SupplierUnApproveDate = CASE WHEN SupplierUnApproveDate IS NOT NULL THEN GETDATE() ELSE NULL END, " & vbCrLf & _
                                                        "SupplierUnApproveUser = CASE WHEN SupplierUnApproveUser IS NOT NULL THEN '" & Session("UserID").ToString & "' ELSE NULL END, " & vbCrLf & _
                                                        "UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf & _
                                                        "AND SupplierApproveDate IS NOT NULL " & vbCrLf
                                                End If
                                        End Select

                                        Select Case Session("SplitStatus")
                                            Case "3", "4", "5", "6"
                                                ls_SQL = ls_SQL + "INSERT INTO PO_MasterUpload_ExportCancel (" & vbCrLf & _
                                                        "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, ETDVendor1, Remarks, " & vbCrLf & _
                                                        "EntryDate, EntryUser, UpdateDate, UpdateUser) " & vbCrLf & _
                                                        "SELECT PONo, AffiliateID, SupplierID, " & vbCrLf & _
                                                        "'" & Trim(cboDelLoc.Text) & "' ForwarderID, " & vbCrLf & _
                                                        "'" & txtOrderNo.Text & "' OrderNo1, " & vbCrLf & _
                                                        "'" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "' ETDVendor1, " & vbCrLf & _
                                                        "'' Remarks, GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, UpdateDate, UpdateUser " & vbCrLf & _
                                                        "FROM PO_MasterUpload_Export " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf

                                                Select Case Session("SplitStatus")
                                                    Case "4", "5", "6"
                                                        ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                            "SET FinalApprovalCls = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                            "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                            "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf

                                                        If Session("SplitStatus") = "4" Then
                                                            ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                                "SET PASIApproveDate = GETDATE(), PASIApproveUser = '" & Session("UserID").ToString & "', UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                "AND OrderNo1 = '" & ls_CancelReffPONo & "' " & vbCrLf
                                                        End If

                                                        Select Case Session("SplitStatus")
                                                            Case "5", "6"
                                                                ls_SQL = ls_SQL + "INSERT INTO DOSupplier_Master_ExportCancel(" & vbCrLf & _
                                                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, OrderNo, DeliveryDate, PIC, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                                    "EntryDate, EntryUser, UpdateDate, UpdateUser, ExcelCls, MovingList) " & vbCrLf & _
                                                                    "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, '" & Trim(ls_OrderNo) & "' OrderNo, DeliveryDate, PIC, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                                    "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, NULL UpdateDate, NULL UpdateUser, '1' ExcelCls, '1' MovingList " & vbCrLf & _
                                                                    "FROM DOSupplier_Master_Export " & vbCrLf & _
                                                                    "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf

                                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Master_Export " & vbCrLf & _
                                                                    "SET ExcelCls = '1', UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf

                                                                If Session("SplitStatus") = "6" Then
                                                                    ls_SQL = ls_SQL + "INSERT INTO ReceiveForwarder_MasterCancel(" & vbCrLf & _
                                                                        "SuratJalanNo, AffiliateID, SupplierID, PONo, ForwarderID, OrderNo, ExcelCls, ReceiveDate, ReceiveBy, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                                        "EntryDate, EntryUser, UpdateDate, UpdateUser, MovingList, SplitReffPONo) " & vbCrLf & _
                                                                        "SELECT SuratJalanNo, AffiliateID, SupplierID, PONo, ForwarderID, '" & Trim(ls_OrderNo) & "' OrderNo, ExcelCls, ReceiveDate, ReceiveBy, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                                        "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, NULL UpdateDate, NULL UpdateUser, MovingList, SplitReffPO " & vbCrLf & _
                                                                        "FROM ReceiveForwarder_Master " & vbCrLf & _
                                                                        "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                        "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                        "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf
                                                                End If
                                                        End Select
                                                End Select
                                        End Select
                                End Select

                                ls_MsgID = "1001"
                            End If
                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                            If pIsUpdate = False Then
                                'INSERT DATA
                                ls_SQL = " 	INSERT INTO dbo.PO_Detail_ExportCancel " & vbCrLf & _
                                         " 	        (PONo, OrderNo1, ForwarderID, AffiliateID, SupplierID, PartNo, Week1, TotalPOQty, PreviousForecast, " & vbCrLf & _
                                         " 	        Forecast1, Forecast2, Forecast3, Variance, VariancePercentage, SplitReffQty, EntryDate, EntryUser) " & vbCrLf & _
                                         " 	VALUES  ( '" & Trim(ls_PONo) & "', " & vbCrLf & _
                                         " 	          '" & Trim(ls_OrderNo) & "', " & vbCrLf & _
                                         " 	          '" & Trim(cboDelLoc.Text) & "', " & vbCrLf & _
                                         " 	          '" & Trim(ls_Affiliate) & "', " & vbCrLf & _
                                         " 	          '" & Trim(ls_Supplier) & "', " & vbCrLf & _
                                         " 	          '" & Trim(ls_PartNo) & "', " & vbCrLf & _
                                         " 	          '" & ls_Week1 & "', " & vbCrLf & _
                                         " 	          '" & ls_TotalPOQty & "', " & vbCrLf & _
                                         " 	          '" & ls_PreviousForecast & "', " & vbCrLf & _
                                         " 	          '" & ls_Forecast1 & "', " & vbCrLf & _
                                         " 	          '" & ls_Forecast2 & "', " & vbCrLf & _
                                         " 	          '" & ls_Forecast3 & "', " & vbCrLf & _
                                         " 	          '" & ls_Variance & "', " & vbCrLf & _
                                         " 	          '" & ls_VariancePercentage & "', " & vbCrLf & _
                                         " 	          '" & ls_CancelReffQty & "', " & vbCrLf & _
                                         " 	          GETDATE(), " & vbCrLf & _
                                         " 	          '" & Session("UserID").ToString & "' " & vbCrLf & _
                                         " 	        ) " & vbCrLf

                                ls_SQL = ls_SQL + "UPDATE PO_Detail_Export " & vbCrLf & _
                                    "SET Week1 = Week1 - " & CDbl(ls_Week1) & ", TotalPOQty = TotalPOQty - " & CDbl(ls_TotalPOQty) & ", UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                ls_SQL = ls_SQL + "DELETE PO_Detail_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                    "AND Week1 = 0 " & vbCrLf

                                ls_SQL = ls_SQL + "DELETE PO_Master_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                    "AND NOT EXISTS(" & vbCrLf & _
                                    "   SELECT PDE.PONo FROM PO_Detail_Export PDE " & vbCrLf & _
                                    "   WHERE PDE.PONo = PO_Master_Export.PONo " & vbCrLf & _
                                    "   AND PDE.AffiliateID = PO_Master_Export.AffiliateID " & vbCrLf & _
                                    "   AND PDE.SupplierID = PO_Master_Export.SupplierID " & vbCrLf & _
                                    "   AND PDE.OrderNo1 = PO_Master_Export.OrderNo1" & vbCrLf & _
                                    ") " & vbCrLf

                                Select Case Session("SplitStatus")
                                    Case "3", "4", "5", "6"
                                        ls_SQL = ls_SQL + "INSERT INTO PO_DetailUpload_ExportCancel(" & vbCrLf & _
                                            "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, PartNo, Week1, Week1Old, TotalPOQty, TotalPOQtyOld, " & vbCrLf & _
                                            "EntryDate, EntryUser) " & vbCrLf & _
                                            "VALUES( " & vbCrLf & _
                                            "'" & Trim(ls_PONo) & "', " & vbCrLf & _
                                            "'" & Trim(ls_Affiliate) & "', " & vbCrLf & _
                                            "'" & Trim(ls_Supplier) & "', " & vbCrLf & _
                                            "'" & cboDelLoc.Text & "', " & vbCrLf & _
                                            "'" & Trim(ls_OrderNo) & "', " & vbCrLf & _
                                            "'" & Trim(ls_PartNo) & "', " & vbCrLf & _
                                            "'" & ls_Week1 & "', " & vbCrLf & _
                                            "'" & ls_Week1 & "', " & vbCrLf & _
                                            "'" & ls_TotalPOQty & "', " & vbCrLf & _
                                            "'" & ls_TotalPOQty & "', " & vbCrLf & _
                                            "GETDATE(), '" & Session("UserID").ToString & "') " & vbCrLf

                                        ls_SQL = ls_SQL + "UPDATE PO_DetailUpload_Export " & vbCrLf & _
                                            "SET Week1 = Week1 - " & CDbl(ls_Week1) & ", TotalPOQty = TotalPOQty - " & CDbl(ls_TotalPOQty) & ", UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                            "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                        ls_SQL = ls_SQL + "DELETE PO_DetailUpload_Export " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                            "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                            "AND Week1 = 0 " & vbCrLf

                                        ls_SQL = ls_SQL + "DELETE PO_MasterUpload_Export " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                            "AND NOT EXISTS(" & vbCrLf & _
                                            "   SELECT PDE.PONo FROM PO_DetailUpload_Export PDE " & vbCrLf & _
                                            "   WHERE PDE.PONo = PO_MasterUpload_Export.PONo " & vbCrLf & _
                                            "   AND PDE.AffiliateID = PO_MasterUpload_Export.AffiliateID " & vbCrLf & _
                                            "   AND PDE.SupplierID = PO_MasterUpload_Export.SupplierID " & vbCrLf & _
                                            "   AND PDE.OrderNo1 = PO_MasterUpload_Export.OrderNo1" & vbCrLf & _
                                            ") " & vbCrLf

                                        Select Case Session("SplitStatus")
                                            Case "5", "6"
                                                ls_SQL = ls_SQL + "INSERT INTO DOSupplier_Detail_ExportCancel(" & vbCrLf & _
                                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, DOQty) " & vbCrLf & _
                                                    "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, '" & Trim(ls_OrderNo) & "' OrderNo, '" & ls_Week1 & "' DOQty " & vbCrLf & _
                                                    "FROM DOSupplier_Detail_Export " & vbCrLf & _
                                                    "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Detail_Export " & vbCrLf & _
                                                    "SET DOQty = DOQty - " & CDbl(ls_Week1) & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                ls_SQL = ls_SQL + "DELETE DOSupplier_Detail_Export " & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                    "AND DOQty = 0 " & vbCrLf

                                                ls_SQL = ls_SQL + "INSERT INTO DOSupplier_DetailBox_ExportCancel(" & vbCrLf & _
                                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, BoxNo) " & vbCrLf & _
                                                    "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, '" & Trim(ls_OrderNo) & "' OrderNo, BoxNo " & vbCrLf & _
                                                    "FROM DOSupplier_DetailBox_Export " & vbCrLf & _
                                                    "WHERE SuratJalanNo + SupplierID + AffiliateID + PONo + PartNo + OrderNo + BoxNo IN ( " & vbCrLf & _
                                                    "   SELECT TOP " & ls_TOP & " A.SuratJalanNo + A.SupplierID + A.AffiliateID + A.PONo + A.PartNo + A.OrderNo + A.BoxNo " & vbCrLf & _
                                                    "   FROM DOSupplier_DetailBox_Export A " & vbCrLf & _
                                                    "   WHERE A.SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "   AND A.AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "   AND A.PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "   AND A.OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                    "   AND A.PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                    "   ORDER BY A.BoxNo DESC " & vbCrLf & _
                                                    ") "

                                                ls_SQL = ls_SQL + "DELETE DOSupplier_DetailBox_Export " & vbCrLf & _
                                                    "WHERE SuratJalanNo + SupplierID + AffiliateID + PONo + PartNo + OrderNo + BoxNo IN ( " & vbCrLf & _
                                                    "   SELECT TOP " & ls_TOP & " A.SuratJalanNo + A.SupplierID + A.AffiliateID + A.PONo + A.PartNo + A.OrderNo + A.BoxNo " & vbCrLf & _
                                                    "   FROM DOSupplier_DetailBox_Export A " & vbCrLf & _
                                                    "   WHERE A.SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "   AND A.AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "   AND A.PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "   AND A.OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                    "   AND A.PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                    "   ORDER BY A.BoxNo DESC " & vbCrLf & _
                                                    ") "

                                                If Session("SplitStatus") = "6" Then
                                                    ls_SQL = ls_SQL + "INSERT INTO ReceiveForwarder_DetailCancel(" & vbCrLf & _
                                                        "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, GoodRecQty, DefectRecQty) " & vbCrLf & _
                                                        "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, '" & Trim(ls_OrderNo) & "' OrderNo, '" & ls_Week1 & "' GoodRecQty, DefectRecQty " & vbCrLf & _
                                                        "FROM ReceiveForwarder_Detail " & vbCrLf & _
                                                        "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                        "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "UPDATE ReceiveForwarder_Detail " & vbCrLf & _
                                                        "SET GoodRecQty = GoodRecQty - " & CDbl(ls_Week1) & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                        "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "DELETE ReceiveForwarder_Detail " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                        "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                        "AND GoodRecQty = 0 " & vbCrLf

                                                    ls_SQL = ls_SQL + "INSERT INTO ReceiveForwarder_DetailBox(" & vbCrLf & _
                                                        "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, BoxNo) " & vbCrLf & _
                                                        "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, '" & Trim(ls_OrderNo) & "' OrderNo, BoxNo " & vbCrLf & _
                                                        "FROM ReceiveForwarder_Master " & vbCrLf & _
                                                        "WHERE SuratJalanNo + SupplierID + AffiliateID + PONo + PartNo + OrderNo + BoxNo IN ( " & vbCrLf & _
                                                        "   SELECT TOP " & ls_TOP & " A.SuratJalanNo + A.SupplierID + A.AffiliateID + A.PONo + A.PartNo + A.OrderNo + A.BoxNo " & vbCrLf & _
                                                        "   FROM ReceiveForwarder_Master A " & vbCrLf & _
                                                        "   WHERE A.SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "   AND A.AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "   AND A.PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "   AND A.OrderNo = '" & Trim(ls_CancelReffPONo) & "' " & vbCrLf & _
                                                        "   AND A.PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                        "   ORDER BY A.BoxNo DESC " & vbCrLf & _
                                                        ") "
                                                End If
                                        End Select
                                End Select

                                ls_MsgID = "1001"
                            Else
                                ls_SQL = " 	UPDATE dbo.PO_Detail_ExportCancel " & vbCrLf & _
                                         " 	   SET ForwarderID = '" & Trim(cboDelLoc.Text) & "' , " & vbCrLf & _
                                         " 	       Week1 = '" & ls_Week1 & "' , " & vbCrLf & _
                                         " 	       TotalPOQty = '" & ls_TotalPOQty & "' , " & vbCrLf & _
                                         " 	       PreviousForecast = '" & ls_PreviousForecast & "' , " & vbCrLf & _
                                         " 	       Forecast1 = '" & ls_Forecast1 & "' , " & vbCrLf & _
                                         " 	       Forecast2 = '" & ls_Forecast2 & "' , " & vbCrLf & _
                                         " 	       Forecast3 = '" & ls_Forecast3 & "' , " & vbCrLf & _
                                         " 	       Variance = '" & ls_Variance & "' , " & vbCrLf & _
                                         " 	       VariancePercentage = '" & ls_VariancePercentage & "' , " & vbCrLf & _
                                         " 	       UpdateDate = GETDATE(), " & vbCrLf & _
                                         " 	       UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                         " 	 WHERE PONo ='" & Trim(ls_PONo) & "' AND OrderNo1 = '" & ls_OrderNo & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                         "   UPDATE dbo.PO_DETAIL_EXPORT SET Week1 = " & ls_CancelReffQty - ls_TotalPOQty & ", TotalPOQty = " & ls_CancelReffQty - ls_TotalPOQty & vbCrLf & _
                                         " 	 WHERE PONo ='" & Trim(ls_PONo) & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                         "   AND OrderNo1 = '" & Trim(ls_CancelReffPONo) & "'"

                                ls_MsgID = "1002"
                            End If

                        ElseIf ls_Active = "0" And pIsUpdate = True And ls_AdaData = "1" Then
                            ls_SQL = "  DELETE from dbo.PO_Detail_ExportCancel " & vbCrLf & _
                                     "  WHERE PONo = '" & Trim(ls_PONo) & "'" & vbCrLf & _
                                     "  AND OrderNo1 = '" & ls_OrderNo & "' " & vbCrLf & _
                                     "  AND AffiliateID = '" & Trim(ls_Affiliate) & "' " & vbCrLf & _
                                     "  AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                     "  AND PartNo = '" & Trim(ls_PartNo) & "' "
                            ls_MsgID = "1003"

                        ElseIf ls_Active = "0" And pIsUpdate = False Then
                            lblInfo.Text = "[ Please give a checkmark to save data ! ] "
                            Session("YA010Msg") = lblInfo.Text
                            Exit Sub
                        End If

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()
                    Else
                        'delete data
                        ls_SQL = "  DELETE from dbo.PO_Detail_ExportCancel" & vbCrLf & _
                                     "  WHERE PONo = '" & Trim(ls_PONo) & "'" & vbCrLf & _
                                     "  AND OrderNo1 = '" & ls_OrderNo & "' " & vbCrLf & _
                                     "  AND AffiliateID = '" & Trim(ls_Affiliate) & "' " & vbCrLf & _
                                     "  AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                     "  AND PartNo = '" & Trim(ls_PartNo) & "' "
                        ls_MsgID = "1003"
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()

                        ls_SQL = " Delete PO_Master_ExportCancel " & vbCrLf & _
                                     "  WHERE PONo = '" & Trim(ls_PONo) & "'" & vbCrLf & _
                                     "  AND AffiliateID = '" & Trim(ls_Affiliate) & "' " & vbCrLf & _
                                     "  AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                     "  AND OrderNo1 = '" & ls_OrderNo & "' " & vbCrLf & _
                                     "  AND NOT EXISTS( " & vbCrLf & _
                                     "      SELECT * FROM PO_Detail_Export a " & vbCrLf & _
                                     "      WHERE a.PONo = PO_Master_Export.PONo " & vbCrLf & _
                                     "      AND a.AffiliateID = PO_Master_Export.AffiliateID " & vbCrLf & _
                                     "      AND a.SupplierID = PO_Master_Export.SupplierID " & vbCrLf & _
                                     "      AND a.OrderNo1 = PO_Master_Export.OrderNo1 " & vbCrLf & _
                                     "  )"

                        ls_MsgID = "1003"
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()
                    End If
                Next iLoop

                If ls_TampungError = 0 Then
                    Session("DataTersimpan") = "1"
                ElseIf ls_TampungError > 0 Then
                    ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                    Session("ErrorData") = lblInfo.Text
                    Exit Sub
                End If

                sqlTran.Commit()

            End Using

            sqlConn.Close()
        End Using

        Call ColorGrid()
        Call clsMsg.DisplayMessage(lblInfo, "1001", clsMessage.MsgType.InformationMessage)
        grid.JSProperties("cpMessage") = lblInfo.Text

        'Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)

        ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        Session("ErrorData") = lblInfo.Text
        lblInfo.Visible = True
        Session.Remove("CekData")
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click

        If btnSubMenu.Text = "BACK" And Session("GOTOStatus") <> "" Then
            Session.Remove("GOTOStatus")
            Session.Remove("LoadSupplier")
            Session.Remove("SplitStatus")
            Session.Remove("GOTOStatus")
            Response.Redirect("~/PurchaseOrderExport/POExportListCancel.aspx")
        Else
            Session.Remove("GOTOStatus")
            Session.Remove("LoadSupplier")
            Session.Remove("SplitStatus")
            Session.Remove("GOTOStatus")
            Response.Redirect("~/PurchaseOrderExport/POExportListCancel.aspx")
        End If

    End Sub

    Private Sub grid_CellEditorInitialize(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewEditorEventArgs) Handles grid.CellEditorInitialize
        If (e.Column.FieldName = "NoUrut" Or e.Column.FieldName = "PartNo" Or e.Column.FieldName = "PartName" Or e.Column.FieldName = "UOM" Or e.Column.FieldName = "MOQ" Or e.Column.FieldName = "PreviousForecast" Or e.Column.FieldName = "Variance" Or e.Column.FieldName = "VariancePercentage") And CType(sender, DevExpress.Web.ASPxGridView.ASPxGridView).IsNewRowEditing = False Then
            e.Editor.ReadOnly = True
        Else
            e.Editor.ReadOnly = False
        End If

        Call ColorGrid()
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Try
            Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False, False)
            grid.JSProperties("cpMessage") = Session("YA010Msg")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "loaddata"

                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "gridload"
                    If Session("YA010Msg") = "" Then
                        Call up_GridLoad()
                        Call ColorGrid()
                    End If

                    Session.Remove("ErrorData")
                    btnApprove.Enabled = False

                Case "exitarea"

                    Exit Sub

                Case "kosong"

                    Call up_GridLoadWhenEventChange()

                Case "savedata"

                    Call up_SaveData()

                Case "gridloadupdate"
                    'If Session("YA010Msg") = "" Then
                    Call up_GridLoadUpdate()
                    Call ColorGrid()
                    'End If

                    Call clsMsg.DisplayMessage(lblInfo, "1016", clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text

                    Session.Remove("ErrorData")
                    btnApprove.Enabled = False

            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        If Session("pCheckError") <> "1" Then
            If e.GetValue("AdaData") = "1" Then
                If CDbl(e.GetValue("Week1")) Mod CDbl(e.GetValue("QtyBox")) <> 0 Then
                    If e.DataColumn.FieldName = "Week1" Or e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Or e.DataColumn.FieldName = "UOM" Or e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "Week1" Or e.DataColumn.FieldName = "PreviousForecast" Or e.DataColumn.FieldName = "Variance" Or e.DataColumn.FieldName = "VariancePercentage" Or e.DataColumn.FieldName = "Forecast1" Or e.DataColumn.FieldName = "Forecast2" Or e.DataColumn.FieldName = "Forecast3" Or e.DataColumn.FieldName = "QtyBox" Or e.DataColumn.FieldName = "SupplierID" Then
                        e.Cell.BackColor = Color.Red
                    End If
                End If

                If CDbl(e.GetValue("Week1")) = 0 Then
                    If e.DataColumn.FieldName = "Week1" Or e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Or e.DataColumn.FieldName = "UOM" Or e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "Week1" Or e.DataColumn.FieldName = "PreviousForecast" Or e.DataColumn.FieldName = "Variance" Or e.DataColumn.FieldName = "VariancePercentage" Or e.DataColumn.FieldName = "Forecast1" Or e.DataColumn.FieldName = "Forecast2" Or e.DataColumn.FieldName = "Forecast3" Or e.DataColumn.FieldName = "QtyBox" Or e.DataColumn.FieldName = "SupplierID" Then
                        e.Cell.BackColor = Color.Red
                    End If
                End If

                If Trim(e.GetValue("ErrorStatus")) <> "" Then
                    If e.DataColumn.FieldName = "Week1" Or e.DataColumn.FieldName = "NoUrut" Or e.DataColumn.FieldName = "PartNo" Or e.DataColumn.FieldName = "PartName" Or e.DataColumn.FieldName = "UOM" Or e.DataColumn.FieldName = "MOQ" Or e.DataColumn.FieldName = "Week1" Or e.DataColumn.FieldName = "PreviousForecast" Or e.DataColumn.FieldName = "Variance" Or e.DataColumn.FieldName = "VariancePercentage" Or e.DataColumn.FieldName = "Forecast1" Or e.DataColumn.FieldName = "Forecast2" Or e.DataColumn.FieldName = "Forecast3" Or e.DataColumn.FieldName = "QtyBox" Or e.DataColumn.FieldName = "SupplierID" Then
                        e.Cell.BackColor = Color.Red
                    End If
                End If

            End If
        End If

        If e.GetValue("VariancePercentage") > 30 Then
            e.Cell.BackColor = Color.Magenta
        End If

    End Sub

    Private Sub ASPxCallback1_Callback(source As Object, e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "recoverypocancel"

                    Call uf_RecoveryCancel()

                    Select Case Session("SplitStatus")
                        Case "2", "3", "4"
                            Call uf_RecoveryCancelEmailSupplier()
                    End Select

                    Select Case Session("SplitStatus")
                        Case "4", "5", "6"
                            Call uf_RecoveryCancelEmailForwarder()
                            Call uf_RecoveryCancelEmailAffiliate()
                    End Select
            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub up_GridLoad()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_supplier As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Session("LoadSupplier") <> "" Then ls_supplier = Session("LoadSupplier")

            ls_SQL = "SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY AllowAccess DESC, PartNo, AffiliateID, SupplierID) NoUrut, * " & vbCrLf & _
                "FROM(" & vbCrLf & _
                "    Select DISTINCT " & vbCrLf & _
                "    '1' AllowAccess, " & vbCrLf & _
                "    '1' AdaData, " & vbCrLf & _
                "    RTRIM(B.PartNo)PartNo, " & vbCrLf & _
                "    RTRIM(C.PartName)PartName, " & vbCrLf & _
                "    RTRIM(ISNULL(d.Description, UPO.UOM))UOM, " & vbCrLf & _
                "    MOQ = CONVERT(NUMERIC(18,0), ISNULL(b.POMOQ,MPM.MOQ)), " & vbCrLf & _
                "    QtyBox = CONVERT(NUMERIC(18,0), ISNULL(b.POQtyBox,MPM.QtyBox)), " & vbCrLf & _
                "    Week1 = 0, " & vbCrLf & _
                "    B.Week2, " & vbCrLf & _
                "    B.Week3, " & vbCrLf & _
                "    B.Week4, " & vbCrLf & _
                "    B.Week5, " & vbCrLf & _
                "    TotalPOQty = 0, " & vbCrLf & _
                "    B.PreviousForecast, " & vbCrLf & _
                "    B.Forecast1, " & vbCrLf & _
                "    B.Forecast2, " & vbCrLf & _
                "    B.Forecast3, " & vbCrLf & _
                "    B.Variance, " & vbCrLf & _
                "    B.VariancePercentage, " & vbCrLf & _
                "    a.PONo, " & vbCrLf & _
                "    a.ShipCls, " & vbCrLf & _
                "    a.CommercialCls, " & vbCrLf & _
                "    a.ForwarderID, " & vbCrLf & _
                "    a.Period, " & vbCrLf & _
                "    RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                "    RTRIM(a.SupplierID)SupplierID, " & vbCrLf & _
                "    ErrorStatus = ISNULL(UPO.errorCls,''), " & vbCrLf & _
                "    a.OrderNo1 CancelReffPONo, " & vbCrLf & _
                "    CONVERT(NUMERIC(18,0), B.Week1) CancelReffQty " & vbCrLf & _
                "    FROM PO_Master_Export a " & vbCrLf & _
                "    INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = B.AffiliateID AND a.SupplierID = B.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                "    LEFT JOIN MS_Parts c ON c.PartNo = B.PartNo " & vbCrLf & _
                "    LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = b.PartNo AND MPM.AffiliateID = b.AffiliateID AND MPM.SupplierID= b.SupplierID " & vbCrLf & _
                "    LEFT JOIN MS_UnitCls d ON d.UnitCls = c.UnitCls " & vbCrLf & _
                "    LEFT JOIN UploadPOExport UPO ON UPO.PONo = a.Pono AND a.AffiliateID = UPO.AffiliateID AND UPO.SupplierID = a.supplierID AND UPO.ForwarderID = a.ForwarderID AND UPO.Partno = b.PartNo " & vbCrLf & _
                "    WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf

            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "    AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            If Trim(ls_supplier) <> "" Then
                ls_SQL = ls_SQL + _
                    "    AND a.SupplierID = '" & Trim(ls_supplier) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + _
                "    AND a.PONO = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                ")X "

            txtOrderNo.Text = up_CreatePOCancelNo(txtpono.Text, cboAffiliate.Text, ls_supplier)

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadUpdate()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_supplier As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Session("LoadSupplier") <> "" Then ls_supplier = Session("LoadSupplier")

            ls_SQL = "SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY AllowAccess DESC, PartNo, AffiliateID, SupplierID) NoUrut, * " & vbCrLf & _
                "FROM(" & vbCrLf & _
                "    Select DISTINCT " & vbCrLf & _
                "    '1' AllowAccess, " & vbCrLf & _
                "    '1' AdaData, " & vbCrLf & _
                "    RTRIM(B.PartNo)PartNo, " & vbCrLf & _
                "    RTRIM(C.PartName)PartName, " & vbCrLf & _
                "    RTRIM(ISNULL(d.Description, UPO.UOM))UOM, " & vbCrLf & _
                "    MOQ = CONVERT(NUMERIC(18,0), MPM.MOQ), " & vbCrLf & _
                "    QtyBox = CONVERT(NUMERIC(18,0), MPM.QtyBox), " & vbCrLf & _
                "    B.Week1, " & vbCrLf & _
                "    B.Week2, " & vbCrLf & _
                "    B.Week3, " & vbCrLf & _
                "    B.Week4, " & vbCrLf & _
                "    B.Week5, " & vbCrLf & _
                "    B.TotalPOQty, " & vbCrLf & _
                "    B.PreviousForecast, " & vbCrLf & _
                "    B.Forecast1, " & vbCrLf & _
                "    B.Forecast2, " & vbCrLf & _
                "    B.Forecast3, " & vbCrLf & _
                "    B.Variance, " & vbCrLf & _
                "    B.VariancePercentage, " & vbCrLf & _
                "    a.PONo, " & vbCrLf & _
                "    a.ShipCls, " & vbCrLf & _
                "    a.CommercialCls, " & vbCrLf & _
                "    a.ForwarderID, " & vbCrLf & _
                "    a.Period, " & vbCrLf & _
                "    RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                "    RTRIM(a.SupplierID)SupplierID, " & vbCrLf & _
                "    ErrorStatus = ISNULL(UPO.errorCls,''), " & vbCrLf & _
                "    a.SplitReffPONo CancelReffPONo, " & vbCrLf & _
                "    CONVERT(NUMERIC(18,0), B.SplitReffQty) CancelReffQty " & vbCrLf & _
                "    FROM PO_Master_ExportCancel a " & vbCrLf & _
                "    INNER JOIN PO_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = B.AffiliateID AND a.SupplierID = B.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                "    LEFT JOIN MS_Parts c ON c.PartNo = B.PartNo " & vbCrLf & _
                "    LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = b.PartNo AND MPM.AffiliateID = b.AffiliateID AND MPM.SupplierID= b.SupplierID " & vbCrLf & _
                "    LEFT JOIN MS_UnitCls d ON d.UnitCls = c.UnitCls " & vbCrLf & _
                "    LEFT JOIN UploadPOExport UPO ON UPO.PONo = a.Pono AND a.AffiliateID = UPO.AffiliateID AND UPO.SupplierID = a.supplierID AND UPO.ForwarderID = a.ForwarderID AND UPO.Partno = b.PartNo " & vbCrLf & _
                "    WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf

            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "    AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            If Trim(ls_supplier) <> "" Then
                ls_SQL = ls_SQL + _
                    "    AND a.SupplierID = '" & Trim(ls_supplier) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + _
                "    AND a.PONO = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                ")X "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False, False)
            End With
            sqlConn.Close()

        End Using
    End Sub

    Private Sub up_GridLoadWhenEventChange()
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select top 0 '' NoUrut, " & vbCrLf & _
                  " '' PartNo, '' PartName, '' UnitCls, '' MOQ, '' QtyBox, " & vbCrLf & _
                  " '' PONo, 0 POQty, '' Week1, '' Week2, '' Week3,   " & vbCrLf & _
                  " '' Week4, '' week5, '' TotalPOQty, " & vbCrLf & _
                  " '' AffiliateID, '' SupplierID " & vbCrLf

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

    Private Sub ColorGrid()
        grid.VisibleColumns(0).CellStyle.BackColor = Color.White
        grid.VisibleColumns(9).CellStyle.BackColor = Color.White
        grid.VisibleColumns(13).CellStyle.BackColor = Color.White
        grid.VisibleColumns(14).CellStyle.BackColor = Color.White
        grid.VisibleColumns(15).CellStyle.BackColor = Color.White
    End Sub

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""

        'Affiliate ID
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(AffiliateName), [Consignee Code] = Rtrim(isnull(AffiliateCode,'')) FROM MS_Affiliate " & vbCrLf
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboAffiliate
                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Affiliate Code")
                .Columns(0).Width = 100
                .Columns.Add("Affiliate Name")
                .Columns(1).Width = 240
                .TextField = "Affiliate Code"
                .Columns.Add("Consignee Code")
                .Columns(2).Width = 100
                .DataBind()
            End With

            sqlConn.Close()
        End Using

        'Forwarder ID
        ls_sql = "SELECT [Forwarder Code] = RTRIM(ForwarderID) ,[Forwarder Name] = RTRIM(ForwarderName), DEFAULTCLS FROM MS_Forwarder ORDER BY DEFAULTCLS DESC, [Forwarder Code] "
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            With cboDelLoc
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
    End Sub

    Private Sub up_SaveData()
        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Dim pIsUpdate As Boolean
            Dim ls_EmergencyCls As String
            Dim ls_Commercial As String
            Dim ls_ShipCls As String
            Dim i As Integer = 0

            If rdEmergency.Checked = True Then
                ls_EmergencyCls = "E"
            ElseIf rdMonthly.Checked = True Then
                ls_EmergencyCls = "M"
            End If

            If rdrCom1.Checked = True Then
                ls_Commercial = "1"
            ElseIf rdrCom2.Checked = True Then
                ls_Commercial = "0"
            End If

            If rdrShipBy2.Checked = True Then
                ls_ShipCls = "B"
            ElseIf rdrShipBy3.Checked = True Then
                ls_ShipCls = "A"
            End If

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryCancelDetail")


                    For i = 0 To grid.VisibleRowCount - 1
                        If Trim(grid.GetRowValues(i, "AllowAccess").ToString) = "1" Then
                            ls_SQL = "SELECT * FROM dbo.PO_MASTER_EXPORT " & vbCrLf & _
                                "WHERE PONo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                                "AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                                "AND SupplierID = '" & Trim(grid.GetRowValues(i, "SupplierID").ToString) & "' " & vbCrLf & _
                                "AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' "

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
                                ls_SQL = " UPDATE dbo.PO_Master_Export " & _
                                         " SET     Period = '" & Convert.ToDateTime(dtPeriodFrom.Value).ToString("yyyy-MM-01") & "', " & vbCrLf & _
                                         "         EmergencyCls = '" & Trim(ls_EmergencyCls) & "'," & vbCrLf & _
                                         "         CommercialCls = '" & Trim(ls_Commercial) & "'," & vbCrLf & _
                                         "         ShipCls = '" & Trim(ls_ShipCls) & "'," & vbCrLf & _
                                         "         ForwarderID = '" & Trim(cboDelLoc.Text) & "'," & vbCrLf & _
                                         "         ETDVendor1 = '" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         ETDPort1 = '" & Convert.ToDateTime(dtETDPort.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         ETAPort1 = '" & Convert.ToDateTime(dtETAPort.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         ETAFactory1 = '" & Convert.ToDateTime(dtETAFactory.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                         "         UpdateDate = GETDATE(), " & vbCrLf & _
                                         "         UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                         "         WHERE PONo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                                         "         AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                                         "         AND SupplierID = '" & Trim(grid.GetRowValues(i, "SupplierID").ToString) & "' " & vbCrLf & _
                                         "         AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' "

                                ls_MsgID = "1002"
                            ElseIf pIsUpdate = False Then

                            End If

                            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm.ExecuteNonQuery()
                            sqlComm.Dispose()

                        End If
                    Next

                    sqlTran.Commit()
                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
            ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text

        Catch ex As Exception
            Me.lblInfo.Visible = True
            Me.lblInfo.Text = ex.Message.ToString
        End Try
    End Sub

    Private Function up_CreatePOCancelNo(pPONo As String, pAffiliate As String, pSupplier As String) As String
        Dim strNewPO As String
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_sql = "SELECT DISTINCT CAST( ISNULL(COUNT (OrderNo1), 0) + 1 AS VARCHAR) PO_COUNT " & vbCrLf & _
                "FROM PO_Master_ExportCancel " & vbCrLf & _
                "WHERE PONo = '" & pPONo & "' " & vbCrLf & _
                "AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                "AND OrderNo1 <> PONo "
            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                strNewPO = pPONo + "~" + Trim(ds.Tables(0).Rows(0)("PO_COUNT"))
            Else
                strNewPO = pPONo + "~1"
            End If

            sqlConn.Close()
        End Using

        Return strNewPO
    End Function

    Private Sub uf_RecoveryCancel()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryCancelDetail")
                ls_sql = "INSERT INTO PO_Master_ExportRecoveryCancel( " & vbCrLf & _
                    "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1) " & vbCrLf & _
                    "SELECT PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1 " & vbCrLf & _
                    "FROM PO_Master_ExportCancel " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "AND NOT EXISTS( " & vbCrLf & _
                    "   SELECT * FROM PO_Master_ExportRecoveryCancel a " & vbCrLf & _
                    "   WHERE a.PONo = PO_Master_ExportCancel.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Master_ExportCancel.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Master_ExportCancel.SupplierID " & vbCrLf & _
                    "   AND a.OrderNo1 = PO_Master_ExportCancel.OrderNo1 " & vbCrLf & _
                    ")" & vbCrLf

                ls_sql = ls_sql & "INSERT INTO PO_Master_Export( " & vbCrLf & _
                    "PONo, AffiliateID, SupplierID, ForwarderID, Period, CommercialCls, EmergencyCls, ShipCls, ErrorStatus, " & vbCrLf & _
                    "OrderNo1, ETDVendor1, ETDPort1, ETAPort1, ETAFactory1, " & vbCrLf & _
                    "OrderNo2, ETDVendor2, ETDPort2, ETAPort2, ETAFactory2, " & vbCrLf & _
                    "OrderNo3, ETDVendor3, ETDPort3, ETAPort3, ETAFactory3, " & vbCrLf & _
                    "OrderNo4, ETDVendor4, ETDPort4, ETAPort4, ETAFactory4, " & vbCrLf & _
                    "OrderNo5, ETDVendor5, ETDPort5, ETAPort5, ETAFactory5, " & vbCrLf & _
                    "UploadDate, UploadUser, PASISendToSupplierDate, PASISendToSupplierUser, SupplierApproveDate, SupplierApproveUser, " & vbCrLf & _
                    "SupplierApprovePartialDate, SupplierApprovePartialUser, SupplierUnApproveDate, SupplierUnApproveUser, " & vbCrLf & _
                    "PASIApproveDate, PASIApproveUser, EntryDate, EntryUser, UpdateDate, UpdateUser, PASISendToSupplierCls, " & vbCrLf & _
                    "SupplierApprovalCls, ExcelCls, FinalApprovalCls, SplitReffPONo, SplitStatus) " & vbCrLf & _
                    "SELECT PONo, AffiliateID, SupplierID, ForwarderID, Period, CommercialCls, EmergencyCls, ShipCls, ErrorStatus, " & vbCrLf & _
                    "SplitReffPONo OrderNo1, ETDVendor1, ETDPort1, ETAPort1, ETAFactory1, " & vbCrLf & _
                    "OrderNo2, ETDVendor2, ETDPort2, ETAPort2, ETAFactory2, " & vbCrLf & _
                    "OrderNo3, ETDVendor3, ETDPort3, ETAPort3, ETAFactory3, " & vbCrLf & _
                    "OrderNo4, ETDVendor4, ETDPort4, ETAPort4, ETAFactory4, " & vbCrLf & _
                    "OrderNo5, ETDVendor5, ETDPort5, ETAPort5, ETAFactory5, " & vbCrLf & _
                    "UploadDate, UploadUser, PASISendToSupplierDate, PASISendToSupplierUser, SupplierApproveDate, SupplierApproveUser, " & vbCrLf & _
                    "SupplierApprovePartialDate, SupplierApprovePartialUser, SupplierUnApproveDate, SupplierUnApproveUser, " & vbCrLf & _
                    "PASIApproveDate, PASIApproveUser, EntryDate, EntryUser, UpdateDate, UpdateUser, PASISendToSupplierCls, " & vbCrLf & _
                    "SupplierApprovalCls, ExcelCls, FinalApprovalCls, NULL SplitReffPONo, NULL SplitStatus " & vbCrLf & _
                    "FROM PO_Master_ExportCancel " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "AND NOT EXISTS( " & vbCrLf & _
                    "   SELECT * FROM PO_Master_Export a " & vbCrLf & _
                    "   WHERE a.PONo = PO_Master_ExportCancel.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Master_ExportCancel.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Master_ExportCancel.SupplierID " & vbCrLf & _
                    "   AND a.OrderNo1 = PO_Master_ExportCancel.SplitReffPONo " & vbCrLf & _
                    ") " & vbCrLf

                ls_sql = ls_sql & "UPDATE PO_Detail_Export SET Week1 = Week1 + ISNULL(( " & vbCrLf & _
                    "   SELECT b.Week1 FROM PO_Master_ExportCancel a " & vbCrLf & _
                    "   INNER JOIN PO_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                    "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "   AND a.PONo = PO_Detail_Export.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Detail_Export.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Detail_Export.SupplierID " & vbCrLf & _
                    "   AND a.SplitReffPONo = PO_Detail_Export.OrderNo1 " & vbCrLf & _
                    "   AND b.PartNo = PO_Detail_Export.PartNo " & vbCrLf & _
                    "), 0), " & vbCrLf & _
                    "TotalPOQty = TotalPOQty + ISNULL(( " & vbCrLf & _
                    "   SELECT b.TotalPOQty FROM PO_Master_ExportCancel a " & vbCrLf & _
                    "   INNER JOIN PO_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                    "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "   AND a.PONo = PO_Detail_Export.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Detail_Export.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Detail_Export.SupplierID " & vbCrLf & _
                    "   AND a.SplitReffPONo = PO_Detail_Export.OrderNo1 " & vbCrLf & _
                    "   AND b.PartNo = PO_Detail_Export.PartNo " & vbCrLf & _
                    "), 0) " & vbCrLf & _
                    "WHERE EXISTS( " & vbCrLf & _
                    "   SELECT * FROM PO_Master_ExportCancel a " & vbCrLf & _
                    "   INNER JOIN PO_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                    "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "   AND a.PONo = PO_Detail_Export.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Detail_Export.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Detail_Export.SupplierID " & vbCrLf & _
                    "   AND a.SplitReffPONo = PO_Detail_Export.OrderNo1 " & vbCrLf & _
                    "   AND b.PartNo = PO_Detail_Export.PartNo " & vbCrLf & _
                    ") " & vbCrLf

                ls_sql = ls_sql & "INSERT INTO PO_Detail_Export( " & vbCrLf & _
                    "PONo, AffiliateID, SupplierID, ForwarderID, PartNo, OrderNo1, " & vbCrLf & _
                    "Week1, Week2, Week3, Week4, Week5, TotalPOQty, " & vbCrLf & _
                    "PreviousForecast, Forecast1, Forecast2, Forecast3, Variance, VariancePercentage, " & vbCrLf & _
                    "EntryDate, EntryUser, UpdateDate, UpdateUser, " & vbCrLf & _
                    "CloseCls, CloseDate, CloseSupplierPIC, SplitReffQty) " & vbCrLf & _
                    "SELECT b.PONo, b.AffiliateID, b.SupplierID, b.ForwarderID, b.PartNo, a.SplitReffPONo OrderNo1, " & vbCrLf & _
                    "b.Week1, b.Week2, b.Week3, b.Week4, b.Week5, b.TotalPOQty, " & vbCrLf & _
                    "b.PreviousForecast, b.Forecast1, b.Forecast2, b.Forecast3, b.Variance, b.VariancePercentage, " & vbCrLf & _
                    "b.EntryDate, b.EntryUser, b.UpdateDate, b.UpdateUser, " & vbCrLf & _
                    "b.CloseCls, b.CloseDate, b.CloseSupplierPIC, NULL SplitReffQty " & vbCrLf & _
                    "FROM PO_Master_ExportCancel a " & vbCrLf & _
                    "INNER JOIN PO_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                    "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "AND NOT EXISTS( " & vbCrLf & _
                    "   SELECT * FROM PO_Detail_Export " & vbCrLf & _
                    "   WHERE PO_Detail_Export.PONo = a.PONo " & vbCrLf & _
                    "   AND PO_Detail_Export.AffiliateID = a.AffiliateID " & vbCrLf & _
                    "   AND PO_Detail_Export.SupplierID = a.SupplierID " & vbCrLf & _
                    "   AND PO_Detail_Export.OrderNo1 = a.SplitReffPONo " & vbCrLf & _
                    "   AND PO_Detail_Export.PartNo = b.PartNo " & vbCrLf & _
                    ") " & vbCrLf

                Select Case Session("SplitStatus")
                    Case "3", "4", "5", "6"
                        ls_sql = ls_sql & "INSERT INTO PO_MasterUpload_Export( " & vbCrLf & _
                            "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, ETDVendor1, Remarks, " & vbCrLf & _
                            "EntryDate, EntryUser, UpdateDate, UpdateUser) " & vbCrLf & _
                            "SELECT b.PONo, b.AffiliateID, b.SupplierID, b.ForwarderID, a.SplitReffPONo OrderNo1, b.ETDVendor1, b.Remarks, " & vbCrLf & _
                            "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, GETDATE() UpdateDate, '" & Session("UserID").ToString & "' UpdateUser " & vbCrLf & _
                            "FROM PO_Master_ExportCancel a " & vbCrLf & _
                            "INNER JOIN PO_MasterUpload_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                            "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                            "AND NOT EXISTS( " & vbCrLf & _
                            "   SELECT * FROM PO_MasterUpload_Export " & vbCrLf & _
                            "   WHERE PO_MasterUpload_Export.PONo = a.PONo " & vbCrLf & _
                            "   AND PO_MasterUpload_Export.AffiliateID = a.AffiliateID " & vbCrLf & _
                            "   AND PO_MasterUpload_Export.SupplierID = a.SupplierID " & vbCrLf & _
                            "   AND PO_MasterUpload_Export.OrderNo1 = a.SplitReffPONo " & vbCrLf & _
                            ") " & vbCrLf

                        ls_sql = ls_sql & "UPDATE PO_DetailUpload_Export SET Week1 = Week1 + ISNULL(( " & vbCrLf & _
                            "   SELECT b.Week1 FROM PO_Master_ExportCancel a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                            "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                            "   AND a.PONo = PO_DetailUpload_Export.PONo " & vbCrLf & _
                            "   AND a.AffiliateID = PO_DetailUpload_Export.AffiliateID " & vbCrLf & _
                            "   AND a.SupplierID = PO_DetailUpload_Export.SupplierID " & vbCrLf & _
                            "   AND a.SplitReffPONo = PO_DetailUpload_Export.OrderNo1 " & vbCrLf & _
                            "   AND b.PartNo = PO_DetailUpload_Export.PartNo " & vbCrLf & _
                            "), 0), " & vbCrLf & _
                            "TotalPOQty = TotalPOQty + ISNULL(( " & vbCrLf & _
                            "   SELECT b.TotalPOQty FROM PO_Master_ExportCancel a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                            "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                            "   AND a.PONo = PO_DetailUpload_Export.PONo " & vbCrLf & _
                            "   AND a.AffiliateID = PO_DetailUpload_Export.AffiliateID " & vbCrLf & _
                            "   AND a.SupplierID = PO_DetailUpload_Export.SupplierID " & vbCrLf & _
                            "   AND a.SplitReffPONo = PO_DetailUpload_Export.OrderNo1 " & vbCrLf & _
                            "   AND b.PartNo = PO_DetailUpload_Export.PartNo " & vbCrLf & _
                            "), 0) " & vbCrLf & _
                            "WHERE EXISTS( " & vbCrLf & _
                            "   SELECT * FROM PO_Master_ExportCancel a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                            "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                            "   AND a.PONo = PO_DetailUpload_Export.PONo " & vbCrLf & _
                            "   AND a.AffiliateID = PO_DetailUpload_Export.AffiliateID " & vbCrLf & _
                            "   AND a.SupplierID = PO_DetailUpload_Export.SupplierID " & vbCrLf & _
                            "   AND a.SplitReffPONo = PO_DetailUpload_Export.OrderNo1 " & vbCrLf & _
                            "   AND b.PartNo = PO_DetailUpload_Export.PartNo " & vbCrLf & _
                            ") " & vbCrLf

                        ls_sql = ls_sql & "INSERT INTO PO_DetailUpload_Export( " & vbCrLf & _
                            "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, PartNo, Week1, Week1Old, TotalPOQty, TotalPOQtyOld, " & vbCrLf & _
                            "EntryDate, EntryUser, UpdateDate, UpdateUser) " & vbCrLf & _
                            "SELECT b.PONo, b.AffiliateID, b.SupplierID, b.ForwarderID, a.SplitReffPONo OrderNo1, b.PartNo, " & vbCrLf & _
                            "b.Week1, b.Week1, b.TotalPOQty, b.TotalPOQty, " & vbCrLf & _
                            "b.EntryDate, b.EntryUser, b.UpdateDate, b.UpdateUser " & vbCrLf & _
                            "FROM PO_Master_ExportCancel a " & vbCrLf & _
                            "INNER JOIN PO_DetailUpload_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                            "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                            "AND NOT EXISTS( " & vbCrLf & _
                            "   SELECT * FROM PO_DetailUpload_Export " & vbCrLf & _
                            "   WHERE PO_DetailUpload_Export.PONo = a.PONo " & vbCrLf & _
                            "   AND PO_DetailUpload_Export.AffiliateID = a.AffiliateID " & vbCrLf & _
                            "   AND PO_DetailUpload_Export.SupplierID = a.SupplierID " & vbCrLf & _
                            "   AND PO_DetailUpload_Export.OrderNo1 = a.SplitReffPONo " & vbCrLf & _
                            "   AND PO_DetailUpload_Export.PartNo = b.PartNo " & vbCrLf & _
                            ") " & vbCrLf

                        ls_sql = ls_sql & "UPDATE PrintLabelExport SET OrderNo = ISNULL(( " & vbCrLf & _
                            "   SELECT a.SplitReffPONo " & vbCrLf & _
                            "   FROM PO_Master_ExportCancel a " & vbCrLf & _
                            "   WHERE a.SupplierID = PrintLabelExport.SupplierID " & vbCrLf & _
                            "   AND a.AffiliateID = PrintLabelExport.AffiliateID " & vbCrLf & _
                            "   AND a.PONo = PrintLabelExport.PONo " & vbCrLf & _
                            "   AND a.OrderNo1 = PrintLabelExport.OrderNo " & vbCrLf & _
                            "), '') " & vbCrLf & _
                            "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                        Select Case Session("SplitStatus")
                            Case "5", "6"
                                ls_sql = ls_sql & "INSERT INTO DOSupplier_Master_Export( " & vbCrLf & _
                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, OrderNo, DeliveryDate, PIC, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                    "EntryDate, EntryUser, UpdateDate, UpdateUser, ExcelCls, MovingList) " & vbCrLf & _
                                    "SELECT b.SuratJalanNo, b.SupplierID, b.AffiliateID, b.PONo, a.SplitReffPONo OrderNo, b.DeliveryDate, b.PIC, b.JenisArmada, b.DriverName, b.DriverContact, b.NoPol, b.TotalBox, " & vbCrLf & _
                                    "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, GETDATE() UpdateDate, '" & Session("UserID").ToString & "' UpdateUser, b.ExcelCls, b.MovingList " & vbCrLf & _
                                    "FROM PO_Master_ExportCancel a " & vbCrLf & _
                                    "INNER JOIN DOSupplier_Master_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                    "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                    "AND NOT EXISTS( " & vbCrLf & _
                                    "   SELECT * FROM DOSupplier_Master_Export " & vbCrLf & _
                                    "   WHERE DOSupplier_Master_Export.PONo = a.PONo " & vbCrLf & _
                                    "   AND DOSupplier_Master_Export.AffiliateID = a.AffiliateID " & vbCrLf & _
                                    "   AND DOSupplier_Master_Export.SupplierID = a.SupplierID " & vbCrLf & _
                                    "   AND DOSupplier_Master_Export.OrderNo = a.SplitReffPONo " & vbCrLf & _
                                    ") " & vbCrLf

                                ls_sql = ls_sql & "UPDATE DOSupplier_Detail_Export SET DOQty = DOQty + ISNULL(( " & vbCrLf & _
                                    "   SELECT b.DOQty FROM PO_Master_ExportCancecl a " & vbCrLf & _
                                    "   INNER JOIN DOSupplier_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                    "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                    "   AND a.PONo = DOSupplier_Detail_Export.PONo " & vbCrLf & _
                                    "   AND a.AffiliateID = DOSupplier_Detail_Export.AffiliateID " & vbCrLf & _
                                    "   AND a.SupplierID = DOSupplier_Detail_Export.SupplierID " & vbCrLf & _
                                    "   AND a.SplitReffPONo = DOSupplier_Detail_Export.OrderNo " & vbCrLf & _
                                    "   AND b.PartNo = DOSupplier_Detail_Export.PartNo " & vbCrLf & _
                                    "), 0) " & vbCrLf & _
                                    "WHERE EXISTS( " & vbCrLf & _
                                    "   SELECT * FROM PO_Master_ExportCancel a " & vbCrLf & _
                                    "   INNER JOIN DOSupplier_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                    "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                    "   AND a.PONo = DOSupplier_Detail_Export.PONo " & vbCrLf & _
                                    "   AND a.AffiliateID = DOSupplier_Detail_Export.AffiliateID " & vbCrLf & _
                                    "   AND a.SupplierID = DOSupplier_Detail_Export.SupplierID " & vbCrLf & _
                                    "   AND a.SplitReffPONo = DOSupplier_Detail_Export.OrderNo " & vbCrLf & _
                                    "   AND b.PartNo = DOSupplier_Detail_Export.PartNo " & vbCrLf & _
                                    ") " & vbCrLf

                                ls_sql = ls_sql & "INSERT INTO DOSupplier_Detail_Export( " & vbCrLf & _
                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, DOQty) " & vbCrLf & _
                                    "SELECT b.SuratJalanNo, b.SupplierID, b.AffiliateID, b.PONo, b.PartNo, a.SplitReffPONo OrderNo, b.DOQty " & vbCrLf & _
                                    "FROM PO_Master_ExportCancel a " & vbCrLf & _
                                    "INNER JOIN DOSupplier_Detail_ExportCancel b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                    "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                    "AND NOT EXISTS( " & vbCrLf & _
                                    "   SELECT * FROM DOSupplier_Detail_Export " & vbCrLf & _
                                    "   WHERE DOSupplier_Detail_Export.PONo = a.PONo " & vbCrLf & _
                                    "   AND DOSupplier_Detail_Export.AffiliateID = a.AffiliateID " & vbCrLf & _
                                    "   AND DOSupplier_Detail_Export.SupplierID = a.SupplierID " & vbCrLf & _
                                    "   AND DOSupplier_Detail_Export.OrderNo = a.SplitReffPONo " & vbCrLf & _
                                    "   AND DOSupplier_Detail_Export.PartNo = b.PartNo " & vbCrLf & _
                                    ") " & vbCrLf

                                ls_sql = ls_sql & "UPDATE DOSupplier_DetailBox_Export SET OrderNo = ISNULL(( " & vbCrLf & _
                                    "   SELECT a.SplitReffPONo " & vbCrLf & _
                                    "   FROM PO_Master_ExportCancel a " & vbCrLf & _
                                    "   WHERE a.SupplierID = DOSupplier_DetailBox_Export.SupplierID " & vbCrLf & _
                                    "   AND a.AffiliateID = DOSupplier_DetailBox_Export.AffiliateID " & vbCrLf & _
                                    "   AND a.PONo = DOSupplier_DetailBox_Export.PONo " & vbCrLf & _
                                    "   AND a.OrderNo1 = DOSupplier_DetailBox_Export.OrderNo " & vbCrLf & _
                                    "), '') " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                                ls_sql = ls_sql & "DELETE DOSupplier_Detail_ExportCancel " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                                ls_sql = ls_sql & "DELETE DOSupplier_Master_ExportCancel " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf
                        End Select

                        ls_sql = ls_sql & "DELETE PO_DetailUpload_ExportCancel " & vbCrLf & _
                            "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                        ls_sql = ls_sql & "DELETE PO_MasterUpload_ExportCancel " & vbCrLf & _
                            "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf
                End Select

                ls_sql = ls_sql & "DELETE PO_Detail_ExportCancel " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                ls_sql = ls_sql & "DELETE PO_Master_ExportCancel " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub uf_RecoveryCancelEmailSupplier()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim dsEmail As New DataSet

            dsEmail = GetEmailToSupplier(cboAffiliate.Text.Trim, "PASI", Session("LoadSupplier"))

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    fromEmail = dsEmail.Tables(0).Rows(iRow)("toEmail")
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepoto")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "SUPP" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("affiliatepocc")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

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
            mailMessage.Subject = "[TRIAL] Notification For PO Cancel Recovery, Order No : " & txtOrderNo.Text.Trim

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

            Dim ls_Body As String = ""
            ls_Body = clsNotification.GetNotification("28", "", txtOrderNo.Text.Trim)

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
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub uf_RecoveryCancelEmailForwarder()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim dsEmail As New DataSet

            dsEmail = GetEmailToForwarder(cboDelLoc.Text.Trim, cboAffiliate.Text.Trim)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    If fromEmail = "" Then
                        fromEmail = dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    Else
                        fromEmail = fromEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") <> "PASI" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

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
            mailMessage.Subject = "[TRIAL] Notification For PO Cancel Recovery, Order No : " & txtOrderNo.Text.Trim

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

            Dim ls_Body As String = ""
            ls_Body = clsNotification.GetNotification("28", "", txtOrderNo.Text.Trim)

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
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub uf_RecoveryCancelEmailAffiliate()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim dsEmail As New DataSet

            dsEmail = GetEmailToAffiliate(cboDelLoc.Text.Trim, cboAffiliate.Text.Trim)

            For iRow = 0 To dsEmail.Tables(0).Rows.Count - 1
                If dsEmail.Tables(0).Rows(iRow)("flag") = "PASI" Then
                    If fromEmail = "" Then
                        fromEmail = dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    Else
                        fromEmail = fromEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanFrom")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptEmail = "" Then
                        receiptEmail = dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    Else
                        receiptEmail = receiptEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanTO")
                    End If
                End If
                If dsEmail.Tables(0).Rows(iRow)("flag") = "AFF" Then
                    If receiptCCEmail = "" Then
                        receiptCCEmail = dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    Else
                        receiptCCEmail = receiptCCEmail & ";" & dsEmail.Tables(0).Rows(iRow)("KanbanCC")
                    End If
                End If
            Next

            receiptCCEmail = Replace(receiptCCEmail, ",", ";")
            receiptEmail = Replace(receiptEmail, ",", ";")

            receiptCCEmail = Replace(receiptCCEmail, " ", "")
            receiptEmail = Replace(receiptEmail, " ", "")
            fromEmail = Replace(fromEmail, " ", "")

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
            mailMessage.Subject = "[TRIAL] Notification For PO Cancel Recovery, Order No : " & txtOrderNo.Text.Trim

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

            Dim ls_Body As String = ""
            ls_Body = clsNotification.GetNotification("28", "", txtOrderNo.Text.Trim)

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
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try
    End Sub

    Private Sub GetSettingEmail()
        Dim ls_SQL As String = ""
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = "SELECT * FROM dbo.Ms_EmailSetting_Export"
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

    Private Function GetEmailToSupplier(ByVal pAfffCode As String, ByVal pPASI As String, ByVal pSupplierID As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "--Affiliate TO-CC: " & vbCrLf & _
                    " select 'AFF' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailAffiliate_Export where AffiliateID='" & pAfffCode & "'" & vbCrLf & _
                    " union all " & vbCrLf & _
                    " --PASI TO -CC " & vbCrLf & _
                    " select 'PASI' flag,affiliatepocc,affiliatepoto,toEmail = affiliatepoto  from ms_emailPASI_Export where AffiliateID='" & Trim(pPASI) & "' " & vbCrLf & _
                    " union all " & vbCrLf & _
                    " --Supplier TO- CC " & vbCrLf & _
                    " select 'SUPP' flag,affiliatepocc,affiliatepoto,toEmail= '' from ms_emailSupplier_Export where SupplierID='" & Trim(pSupplierID) & "'"

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Function GetEmailToForwarder(ByVal pFWD As String, ByVal pAff As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select Flag = 'FWD', KanbanCC = isnull(POExportCC,''), KanbanTO = isnull(POExportTo,''), KanbanFrom ='' From ms_emailForwarder where ForwarderID = '" & Trim(pFWD) & "'" & vbCrLf & _
                     " union ALL " & vbCrLf & _
                     " select Flag = 'PASI', kanbanCC = isnull(AffiliatePOCC,'') , kanbanTo = isnull(AffiliatePOTo,''), kanbanFrom = isnull(AffiliatePOTo,'') from MS_EmailPasi_Export  " & vbCrLf & _
                     " select Flag = 'AFF', kanbanCC = isnull(AffiliatePOCC,'') , kanbanTo = isnull(AffiliatePOTo,''), kanbanFrom = isnull(AffiliatePOTo,'') from MS_EmailAffiliate_Export where AffiliateID = '" & Trim(pAff) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function

    Private Function GetEmailToAffiliate(ByVal pFWD As String, ByVal pAff As String) As DataSet
        Dim ls_SQL As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = " select Flag = 'FWD', KanbanCC = isnull(POExportCC,''), KanbanTO = isnull(POExportTo,''), KanbanFrom ='' From ms_emailForwarder where ForwarderID = '" & Trim(pFWD) & "'" & vbCrLf & _
                     " union ALL " & vbCrLf & _
                     " select Flag = 'PASI', kanbanCC = isnull(AffiliatePOCC,'') , kanbanTo = isnull(AffiliatePOTo,''), kanbanFrom = isnull(AffiliatePOTo,'') from MS_EmailPasi_Export  " & vbCrLf & _
                     " select Flag = 'AFF', kanbanCC = isnull(AffiliatePOCC,'') , kanbanTo = isnull(AffiliatePOTo,''), kanbanFrom = isnull(AffiliatePOTo,'') from MS_EmailAffiliate_Export where AffiliateID = '" & Trim(pAff) & "' " & vbCrLf

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            End If
        End Using
    End Function
#End Region

End Class