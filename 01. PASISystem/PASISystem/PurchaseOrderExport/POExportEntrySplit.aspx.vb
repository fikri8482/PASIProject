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

Public Class POExportEntrySplit

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

                Session("MenuDesc") = "SPLIT PO"
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
                            Session.Remove("LoadForwarder")
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
                            pStatus = True

                            Session("LoadSupplier") = pSupplierCode
                            Session("LoadForwarder") = pDeliveryCode

                            Call up_GridLoad()
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
                            pStatus = True

                            Session("LoadSupplier") = pSupplierCode
                            Session("LoadForwarder") = pDeliveryCode

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
                            pStatus = True

                            Session("LoadSupplier") = pSupplierCode
                            Session("LoadForwarder") = pDeliveryCode

                            Call up_GridLoad()
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
        Dim ls_SplitReffPONo As String = ""
        Dim ls_SplitReffQty As String = ""
        Dim ls_TOP As String = ""
        Dim sqlComm As New SqlCommand
        Dim a As Integer

        Session.Remove("ErrorData")
        Session.Remove("YA010Msg")

        a = e.UpdateValues.Count
        For iLoop = 0 To a - 1
            ls_Week1 = Trim(e.UpdateValues(iLoop).NewValues("Week1").ToString())
            ls_SplitReffQty = Trim(e.UpdateValues(iLoop).NewValues("SplitReffQty").ToString())
            ls_qtybox = Trim(e.UpdateValues(iLoop).NewValues("QtyBox").ToString())

            If ls_Week1 = "0" Then
                lblInfo.Text = "[ Please give a checkmark to save data ! ] "
                Session("YA010Msg") = lblInfo.Text
                Exit Sub
            End If
            If CDbl(ls_Week1) > CDbl(ls_SplitReffQty) Then
                lblInfo.Text = "[ Qty Split is bigger than Refrence PO Qty ! ] "
                Session("YA010Msg") = lblInfo.Text
                Exit Sub
            End If
            If (CDbl(ls_Week1) Mod CDbl(ls_qtybox)) <> 0 Then
                lblInfo.Text = "[ Qty Split must be multiple then Qty/Box ! ] "
                Session("YA010Msg") = lblInfo.Text
                Exit Sub
            End If
        Next

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntrySplit")
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
                    ls_SplitReffPONo = Trim(e.UpdateValues(iLoop).NewValues("SplitReffPONo").ToString())
                    ls_SplitReffQty = Trim(e.UpdateValues(iLoop).NewValues("SplitReffQty").ToString())
                    ls_TOP = CDbl(ls_Week1) / CDbl(ls_qtybox)

                    Dim sqlstring As String
                    sqlstring = "SELECT * FROM PO_Detail_Export WHERE PONo ='" & Trim(ls_PONo) & "' AND OrderNo1 = '" & ls_OrderNo & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "'"
                    sqlComm = New SqlCommand(sqlstring, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    sqlstring = "SELECT * FROM PO_Master_Export WHERE PONo ='" & Trim(ls_PONo) & "' AND OrderNo1 = '" & ls_OrderNo & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "'"
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
                                         " WHERE PONo = '" & Trim(ls_PONo) & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND SupplierID = '" & Trim(ls_Supplier) & "'" & vbCrLf & _
                                         " AND OrderNo1 = '" & Trim(ls_OrderNo) & "'"

                                ls_MsgID = "1002"

                            ElseIf pIsUpdateMaster = False Then
                                'Insert
                                ls_SQL = "INSERT INTO PO_Master_Export( " & vbCrLf & _
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
                                    "SupplierApprovalCls, ExcelCls, FinalApprovalCls, '" & ls_SplitReffPONo & "' SplitReffPONo, '" & Session("POStatus") & "' SplitStatus " & vbCrLf & _
                                    "FROM PO_Master_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf

                                Select Session("POStatus")
                                    Case "2", "3", "4", "5", "6"
                                        Select Case Session("POStatus")
                                            Case "2", "3"
                                                ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                    "SET ExcelCls = 3, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                    "AND OrderNo1 = '" & txtOrderNo.Text & "' " & vbCrLf

                                                If Session("POStatus") = "2" Then
                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET PASISendToSupplierDate = GETDATE(), PASISendToSupplierUser = '" & Session("UserID").ToString & "', UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf
                                                End If

                                                If Session("POStatus") = "3" Then
                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET SupplierApproveDate = NULL , " & vbCrLf & _
                                                        "SupplierApproveUser = NULL, " & vbCrLf & _
                                                        "SupplierApprovePartialDate = NULL, " & vbCrLf & _
                                                        "SupplierApprovePartialUser = NULL, " & vbCrLf & _
                                                        "SupplierUnApproveDate = NULL, " & vbCrLf & _
                                                        "SupplierUnApproveUser = NULL, " & vbCrLf & _
                                                        "UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & txtOrderNo.Text & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET SupplierApproveDate = NULL , " & vbCrLf & _
                                                        "SupplierApproveUser = NULL, " & vbCrLf & _
                                                        "SupplierApprovePartialDate = NULL, " & vbCrLf & _
                                                        "SupplierApprovePartialUser = NULL, " & vbCrLf & _
                                                        "SupplierUnApproveDate = NULL, " & vbCrLf & _
                                                        "SupplierUnApproveUser = NULL, " & vbCrLf & _
                                                        "UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "DELETE PO_MasterUpload_Export " & vbCrLf & _
                                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                            "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                            "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf
                                                End If
                                        End Select

                                        Select Case Session("POStatus")
                                            Case "4", "5", "6"
                                                ls_SQL = ls_SQL + "INSERT INTO PO_MasterUpload_Export (" & vbCrLf & _
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
                                                                "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf

                                                If Session("POStatus") = "4" Then
                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET FinalApprovalCls = 3, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & txtOrderNo.Text & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                        "SET PASIApproveDate = GETDATE(), PASIApproveUser = '" & Session("UserID").ToString & "', UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                        "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf

                                                    If cboDelLoc.Text.Trim <> Session("LoadForwarder").ToString.Trim Then
                                                        ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                            "SET FinalApprovalCls = 3, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                            "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                            "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf
                                                    End If
                                                End If

                                                Select Case Session("POStatus")
                                                    Case "5", "6"
                                                        ls_SQL = ls_SQL + "INSERT INTO DOSupplier_Master_Export(" & vbCrLf & _
                                                            "SuratJalanNo, SupplierID, AffiliateID, PONo, OrderNo, DeliveryDate, PIC, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                            "EntryDate, EntryUser, UpdateDate, UpdateUser, ExcelCls, MovingList, SplitReffPONo) " & vbCrLf & _
                                                            "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, '" & Trim(ls_OrderNo) & "' OrderNo, DeliveryDate, PIC, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                            "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, NULL UpdateDate, NULL UpdateUser, ExcelCls, MovingList, '" & Trim(ls_SplitReffPONo) & "' SplitReffPONo " & vbCrLf & _
                                                            "FROM DOSupplier_Master_Export " & vbCrLf & _
                                                            "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                            "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                            "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf

                                                        If Session("POStatus") = 5 Then
                                                            If cboDelLoc.Text.Trim = Session("LoadForwarder").ToString.Trim Then
                                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Master_Export " & vbCrLf & _
                                                                    "SET ExcelCls = 3, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & txtOrderNo.Text & "' " & vbCrLf
                                                            Else
                                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Master_Export " & vbCrLf & _
                                                                    "SET ExcelCls = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf

                                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Master_Export " & vbCrLf & _
                                                                    "SET ExcelCls = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & txtOrderNo.Text & "' " & vbCrLf

                                                                ls_SQL = ls_SQL + "UPDATE PO_Master_Export " & vbCrLf & _
                                                                    "SET FinalApprovalCls = 4, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                    "AND OrderNo1 = '" & txtOrderNo.Text & "' " & vbCrLf
                                                            End If
                                                        End If

                                                        If Session("POStatus") = 6 Then
                                                            ls_SQL = ls_SQL + "INSERT INTO ReceiveForwarder_Master(" & vbCrLf & _
                                                                "SuratJalanNo, AffiliateID, SupplierID, PONo, ForwarderID, OrderNo, ExcelCls, ReceiveDate, ReceiveBy, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                                "EntryDate, EntryUser, UpdateDate, UpdateUser, MovingList, SplitReffPONo) " & vbCrLf & _
                                                                "SELECT SuratJalanNo, AffiliateID, SupplierID, PONo, '" & cboDelLoc.Text.Trim & "' ForwarderID, '" & Trim(ls_OrderNo) & "' OrderNo, ExcelCls, ReceiveDate, ReceiveBy, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                                                "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, NULL UpdateDate, NULL UpdateUser, '1' MovingList, '" & Trim(ls_SplitReffPONo) & "' SplitReffPONo " & vbCrLf & _
                                                                "FROM ReceiveForwarder_Master " & vbCrLf & _
                                                                "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                                "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf

                                                            If cboDelLoc.Text.Trim <> Session("LoadForwarder").ToString.Trim Then
                                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Master_Export " & vbCrLf & _
                                                                    "SET ExcelCls = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf

                                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Master_Export " & vbCrLf & _
                                                                    "SET MovingList = 1, UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                                    "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                                                    "AND OrderNo = '" & txtOrderNo.Text & "' " & vbCrLf
                                                            End If
                                                        End If
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
                                ls_SQL = " 	INSERT INTO dbo.PO_Detail_Export " & vbCrLf & _
                                         " 	        (PONo, OrderNo1, ForwarderID, AffiliateID, SupplierID, PartNo, Week1, TotalPOQty, PreviousForecast, " & vbCrLf & _
                                         " 	        Forecast1, Forecast2, Forecast3, Variance, VariancePercentage, SplitReffQty, EntryDate, EntryUser, POMOQ, POQtyBox) " & vbCrLf & _
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
                                         " 	          '" & ls_SplitReffQty & "', " & vbCrLf & _
                                         " 	          GETDATE(), " & vbCrLf & _
                                         " 	          '" & Session("UserID").ToString & "', " & vbCrLf & _
                                         " 	          '" & uf_GetMOQ(Trim(ls_PartNo), Trim(ls_Supplier), Trim(ls_Affiliate)) & "', " & vbCrLf & _
                                         " 	          '" & uf_GetQtybox(Trim(ls_PartNo), Trim(ls_Supplier), Trim(ls_Affiliate)) & "' " & vbCrLf & _
                                         " 	        ) " & vbCrLf

                                ls_SQL = ls_SQL + "UPDATE PO_Detail_Export " & vbCrLf & _
                                    "SET Week1 = Week1 - " & CDbl(ls_Week1) & ", TotalPOQty = TotalPOQty - " & CDbl(ls_TotalPOQty) & ", UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                ls_SQL = ls_SQL + "DELETE PO_Detail_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                    "AND Week1 = 0 " & vbCrLf

                                ls_SQL = ls_SQL + "DELETE PO_Master_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                    "AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                    "AND NOT EXISTS(" & vbCrLf & _
                                    "   SELECT PDE.PONo FROM PO_Detail_Export PDE " & vbCrLf & _
                                    "   WHERE PDE.PONo = PO_Master_Export.PONo " & vbCrLf & _
                                    "   AND PDE.AffiliateID = PO_Master_Export.AffiliateID " & vbCrLf & _
                                    "   AND PDE.SupplierID = PO_Master_Export.SupplierID " & vbCrLf & _
                                    "   AND PDE.OrderNo1 = PO_Master_Export.OrderNo1 " & vbCrLf & _
                                    ") " & vbCrLf

                                If Session("POStatus") = "3" Then
                                    ls_SQL = ls_SQL + "DELETE PO_DetailUpload_Export " & vbCrLf & _
                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                        "AND SupplierID = '" & ls_Supplier & "' " & vbCrLf & _
                                        "AND OrderNo1 = '" & ls_SplitReffPONo & "' " & vbCrLf
                                End If

                                Select Case Session("POStatus")
                                    Case "4", "5", "6"
                                        ls_SQL = ls_SQL + "INSERT INTO PO_DetailUpload_Export(" & vbCrLf & _
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
                                            "SET Week1 = Week1 - " & CDbl(ls_Week1) & ", TotalPOQty = TotalPOQty - " & CDbl(ls_TotalPOQty) & ", " & vbCrLf & _
                                            "Week1Old = Week1Old - " & CDbl(ls_Week1) & ", TotalPOQtyOld = TotalPOQtyOld - " & CDbl(ls_TotalPOQty) & ", " & vbCrLf & _
                                            "UpdateDate = GETDATE(), UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                            "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                        ls_SQL = ls_SQL + "DELETE PO_DetailUpload_Export " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                            "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                            "AND Week1 = 0 " & vbCrLf

                                        ls_SQL = ls_SQL + "DELETE PO_MasterUpload_Export " & vbCrLf & _
                                            "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                            "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                            "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                            "AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                            "AND NOT EXISTS(" & vbCrLf & _
                                            "   SELECT PDE.PONo FROM PO_DetailUpload_Export PDE " & vbCrLf & _
                                            "   WHERE PDE.PONo = PO_MasterUpload_Export.PONo " & vbCrLf & _
                                            "   AND PDE.AffiliateID = PO_MasterUpload_Export.AffiliateID " & vbCrLf & _
                                            "   AND PDE.SupplierID = PO_MasterUpload_Export.SupplierID " & vbCrLf & _
                                            "   AND PDE.OrderNo1 = PO_MasterUpload_Export.OrderNo1 " & vbCrLf & _
                                            ") " & vbCrLf

                                        ls_SQL = ls_SQL + " UPDATE PrintLabelExport SET OrderNo = '" & Trim(ls_OrderNo) & "' " & vbCrLf & _
                                                      " WHERE SupplierID + AffiliateID + OrderNo + LabelNo + PartNo in ( " & vbCrLf & _
                                                      " SELECT TOP " & ls_TOP & " SupplierID + AffiliateID + OrderNo + LabelNo + PartNo from PrintLabelExport " & vbCrLf & _
                                                      " WHERE PONo = '" & txtpono.Text & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                      " ORDER BY LabelNo DESC) " & vbCrLf & _
                                                      "  " & vbCrLf

                                        Select Case Session("POStatus")
                                            Case "5", "6"
                                                ls_SQL = ls_SQL + "INSERT INTO DOSupplier_Detail_Export(" & vbCrLf & _
                                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, DOQty, POMOQ, POQtyBox) " & vbCrLf & _
                                                    "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, '" & Trim(ls_OrderNo) & "' OrderNo, '" & ls_Week1 & "' DOQty, " & vbCrLf & _
                                                    "'" & uf_GetMOQ(Trim(ls_PartNo), Trim(ls_Supplier), cboAffiliate.Text) & "', " & vbCrLf & _
                                                    "'" & uf_GetQtybox(Trim(ls_PartNo), Trim(ls_Supplier), cboAffiliate.Text) & "' " & vbCrLf & _
                                                    "FROM DOSupplier_Detail_Export " & vbCrLf & _
                                                    "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                ls_SQL = ls_SQL + "UPDATE DOSupplier_Detail_Export " & vbCrLf & _
                                                    "SET DOQty = DOQty - " & CDbl(ls_Week1) & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                ls_SQL = ls_SQL + "DELETE DOSupplier_Detail_Export " & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                    "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                    "AND DOQty = 0 " & vbCrLf

                                                ls_SQL = ls_SQL + "DELETE DOSupplier_Master_Export " & vbCrLf & _
                                                    "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                    "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                    "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                    "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                    "AND NOT EXISTS(" & vbCrLf & _
                                                    "   SELECT PDE.PONo FROM DOSupplier_Detail_Export PDE " & vbCrLf & _
                                                    "   WHERE PDE.PONo = DOSupplier_Master_Export.PONo " & vbCrLf & _
                                                    "   AND PDE.AffiliateID = DOSupplier_Master_Export.AffiliateID " & vbCrLf & _
                                                    "   AND PDE.SupplierID = DOSupplier_Master_Export.SupplierID " & vbCrLf & _
                                                    "   AND PDE.OrderNo = DOSupplier_Master_Export.OrderNo " & vbCrLf & _
                                                    ") " & vbCrLf

                                                ls_SQL = ls_SQL + " UPDATE DOSupplier_DetailBox_Export SET OrderNo = '" & Trim(ls_OrderNo) & "' " & vbCrLf & _
                                                      " WHERE SuratJalanNo + SupplierID + AffiliateID + OrderNo + BoxNo + PartNo in ( " & vbCrLf & _
                                                      " SELECT TOP " & ls_TOP & " SuratJalanNo + SupplierID + AffiliateID + OrderNo + BoxNo + PartNo from DOSupplier_DetailBox_Export " & vbCrLf & _
                                                      " WHERE PONo = '" & txtpono.Text & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                      " ORDER BY BoxNo DESC) " & vbCrLf & _
                                                      "  " & vbCrLf

                                                If Session("POStatus") = "6" Then
                                                    ls_SQL = ls_SQL + "INSERT INTO ReceiveForwarder_Detail(" & vbCrLf & _
                                                        "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, GoodRecQty, DefectRecQty) " & vbCrLf & _
                                                        "SELECT SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, '" & Trim(ls_OrderNo) & "' OrderNo, '" & ls_Week1 & "' GoodRecQty, 0 DefectRecQty " & vbCrLf & _
                                                        "FROM ReceiveForwarder_Detail " & vbCrLf & _
                                                        "WHERE SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                        "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "UPDATE ReceiveForwarder_Detail " & vbCrLf & _
                                                        "SET GoodRecQty = GoodRecQty - " & CDbl(ls_Week1) & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                        "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + " Update ReceiveForwarder_DetailBox " & vbCrLf & _
                                                        " Set OrderNo = '" & Trim(txtOrderNo.Text) & "'" & vbCrLf & _
                                                        " Where PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        " And OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                        " And SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        " And AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        " And PartNo = '" & ls_PartNo & "' " & vbCrLf

                                                    ls_SQL = ls_SQL + "DELETE ReceiveForwarder_Detail " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                        "AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                                        "AND GoodRecQty = 0 " & vbCrLf & _
                                                        "AND DefectRecQty = 0 " & vbCrLf

                                                    ls_SQL = ls_SQL + "DELETE ReceiveForwarder_Master " & vbCrLf & _
                                                        "WHERE PONo = '" & txtpono.Text & "' " & vbCrLf & _
                                                        "AND AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                                                        "AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                                        "AND OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                                                        "AND NOT EXISTS(" & vbCrLf & _
                                                        "   SELECT PDE.PONo FROM ReceiveForwarder_Detail PDE " & vbCrLf & _
                                                        "   WHERE PDE.PONo = ReceiveForwarder_Master.PONo " & vbCrLf & _
                                                        "   AND PDE.AffiliateID = ReceiveForwarder_Master.AffiliateID " & vbCrLf & _
                                                        "   AND PDE.SupplierID = ReceiveForwarder_Master.SupplierID " & vbCrLf & _
                                                        "   AND PDE.OrderNo = ReceiveForwarder_Master.OrderNo " & vbCrLf & _
                                                        ") " & vbCrLf
                                                End If
                                        End Select
                                End Select

                                ls_MsgID = "1001"
                            Else
                                ls_SQL = " 	UPDATE dbo.PO_DETAIL_EXPORT " & vbCrLf & _
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
                                         "   UPDATE dbo.PO_DETAIL_EXPORT SET Week1 = " & ls_SplitReffQty - ls_TotalPOQty & ", TotalPOQty = " & ls_SplitReffQty - ls_TotalPOQty & vbCrLf & _
                                         " 	 WHERE PONo ='" & Trim(ls_PONo) & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "' " & vbCrLf & _
                                         "   AND OrderNo1 = '" & Trim(ls_SplitReffPONo) & "'"

                                ls_MsgID = "1002"
                            End If

                        ElseIf ls_Active = "0" And pIsUpdate = True And ls_AdaData = "1" Then
                            ls_SQL = "  DELETE from dbo.PO_Detail_Export" & vbCrLf & _
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
                        ls_SQL = "  DELETE from dbo.PO_Detail_Export" & vbCrLf & _
                                     "  WHERE PONo = '" & Trim(ls_PONo) & "'" & vbCrLf & _
                                     "  AND OrderNo1 = '" & ls_OrderNo & "' " & vbCrLf & _
                                     "  AND AffiliateID = '" & Trim(ls_Affiliate) & "' " & vbCrLf & _
                                     "  AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                     "  AND PartNo = '" & Trim(ls_PartNo) & "' "
                        ls_MsgID = "1003"
                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()

                        ls_SQL = " Delete PO_Master_Export " & vbCrLf & _
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

                ''SPLIT TO ReceiveForwarder_DetailBox
                'If Session("POStatus") = "6" Then
                '    Dim ls_qtybox1 As Integer = 0
                '    Dim ls_box1 As Integer = 0
                '    Dim ls_BoxRec As Integer = 0
                '    Dim ls_label1 As String = ""
                '    Dim ls_label2 As String = ""
                '    Dim ls_BoxMin As Integer = 0
                '    Dim ls_BoxMax As Integer = 0
                '    a = e.UpdateValues.Count
                '    For iLoop = 0 To a - 1
                '        ls_PartNo = Trim(e.UpdateValues(iLoop).NewValues("PartNo").ToString())
                '        'Cek Qty/Box
                '        ls_SQL = " Select QtyPerBox = RD.GoodRecQty/SUM(RDB.Box) From ReceiveForwarder_Detail RD " & vbCrLf & _
                '                 " Left Join ReceiveForwarder_DetailBox RDB ON RD.PONo = RDB.PONo And RD.SupplierID = RDB.SupplierID And RD.PartNo = RDB.PartNo " & vbCrLf & _
                '                 " Where RD.PONo = '" & txtpono.Text & "' And RD.OrderNo = '" & Trim(ls_SplitReffPONo) & "' And RD.SupplierID = '" & Trim(ls_Supplier) & "' And RD.AffiliateID = '" & cboAffiliate.Text & "' And RD.PartNo = '" & ls_PartNo & "' And RDB.StatusDefect = 0 " & vbCrLf & _
                '                 " Group By RD.GoodRecQty "
                '        Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                '        sqlDA.SelectCommand.Transaction = sqlTran
                '        Dim ds As New DataSet
                '        sqlDA.Fill(ds)
                '        If ds.Tables(0).Rows.Count <> 0 Then
                '            ls_qtybox1 = ds.Tables(0).Rows(0).Item("QtyPerBox")
                '        End If
                '        If ls_qtybox1 > 0 Then
                '            ls_box1 = ls_Week1 / ls_qtybox1
                '        End If

                '        'SPLIT PO
                '        ls_SQL = "Select * From ReceiveForwarder_DetailBox Where PONo = '" & txtpono.Text & "' And OrderNo = '" & Trim(ls_SplitReffPONo) & "' And SupplierID = '" & Trim(ls_Supplier) & "' And AffiliateID = '" & cboAffiliate.Text & "' And PartNo = '" & ls_PartNo & "' And StatusDefect = 0"
                '        Dim sqlDA2 As New SqlDataAdapter(ls_SQL, sqlConn)
                '        sqlDA2.SelectCommand.Transaction = sqlTran
                '        Dim ds2 As New DataSet
                '        sqlDA2.Fill(ds2)
                '        If ds2.Tables(0).Rows.Count <> 0 Then
                '            For i = 0 To ds2.Tables(0).Rows.Count - 1
                '                ls_BoxRec = ds2.Tables(0).Rows(i).Item("Box")
                '                ls_label1 = Trim(ds2.Tables(0).Rows(i).Item("Label1"))
                '                ls_label2 = Trim(ds2.Tables(0).Rows(i).Item("Label2"))
                '                If ls_box1 >= ls_BoxRec Then
                '                    ls_SQL = " Update ReceiveForwarder_DetailBox " & vbCrLf & _
                '                             " Set OrderNo = '" & Trim(ls_PONo) & "'" & vbCrLf & _
                '                             " Where PONo = '" & txtpono.Text & "' " & vbCrLf & _
                '                             " And OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                '                             " And SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                '                             " And AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                '                             " And PartNo = '" & ls_PartNo & "' " & vbCrLf & _
                '                             " And Label1 = '" & ls_label1 & "' " & vbCrLf & _
                '                             " And Label2 = '" & ls_label2 & "' " & vbCrLf & _
                '                             ""
                '                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                '                    sqlComm.ExecuteNonQuery()
                '                    sqlComm.Dispose()
                '                Else
                '                    'ls_SQL = " Insert Into ReceiveForwarder_DetailBox (SuratJalanNo,SupplierID,AffiliateID,PONo,OrderNo,PartNo,Label1,Label2,Box,StatusDefect,ExcelCls)" & vbCrLf & _
                '                    '         " Values ( " & vbCrLf & _
                '                    '         " '" & Trim(ds2.Tables(0).Rows(i).Item("SuratJalanNo")) & "', " & vbCrLf & _
                '                    '         " '" & Trim(ds2.Tables(0).Rows(i).Item("AffiliateID")) & "', " & vbCrLf & _
                '                    '         " '" & Trim(ds2.Tables(0).Rows(i).Item("PONo")) & "', " & vbCrLf & _
                '                    '         " '" & Trim(ls_PONo) & "', " & vbCrLf & _
                '                    '         " )"
                '                End If
                '                ls_box1 = ls_box1 - ls_BoxRec
                '                If ls_box1 = 0 Then Exit For
                '            Next

                '            'Delete Data When Box 0
                '            ls_SQL = " Delete ReceiveForwarder_DetailBox " & vbCrLf & _
                '                     " Where PONo = '" & txtpono.Text & "' " & vbCrLf & _
                '                     " And OrderNo = '" & Trim(ls_SplitReffPONo) & "' " & vbCrLf & _
                '                     " And SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                '                     " And AffiliateID = '" & cboAffiliate.Text & "' " & vbCrLf & _
                '                     " And PartNo = '" & ls_PartNo & "' " & vbCrLf & _
                '                     " And Label1 = '" & ls_label1 & "' " & vbCrLf & _
                '                     " And Label2 = '" & ls_label2 & "' " & vbCrLf & _
                '                     " And Box = 0 " & vbCrLf & _
                '                     " And StatusDefect = 0 "
                '            sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                '            sqlComm.ExecuteNonQuery()
                '            sqlComm.Dispose()
                '        End If
                '    Next iLoop
                'End If

                If ls_TampungError = 0 Then
                    Session("DataTersimpan") = "1"
                ElseIf ls_TampungError > 0 Then
                    'ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
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

        'ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        Session("ErrorData") = lblInfo.Text
        lblInfo.Visible = True
        Session.Remove("CekData")
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
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntrySplit")


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
            'ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text

        Catch ex As Exception
            Me.lblInfo.Visible = True
            Me.lblInfo.Text = ex.Message.ToString
        End Try
    End Sub

    Private Function up_CreateSplitNo(pPONo As String, pAffiliate As String, pSupplier As String) As String
        Dim strNewPO As String
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_sql = "SELECT DISTINCT CAST( ISNULL(COUNT (OrderNo1), 0) + 1 AS VARCHAR) PO_COUNT " & vbCrLf & _
                "FROM( " & vbCrLf & _
                "   SELECT PONo, AffiliateID, OrderNo1 FROM PO_Master_Export " & vbCrLf & _
                "   UNION " & vbCrLf & _
                "   SELECT PONo, AffiliateID, OrderNo1 FROM PO_Master_ExportRecoverySplit " & vbCrLf & _
                ")PO " & vbCrLf & _
                "WHERE PONo = '" & pPONo & "' " & vbCrLf & _
                "AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                "AND OrderNo1 <> PONo "
            Dim sqlDA As New SqlDataAdapter(ls_sql, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                strNewPO = pPONo + "-" + Trim(ds.Tables(0).Rows(0)("PO_COUNT"))
            Else
                strNewPO = pPONo + "-1"
            End If

            sqlConn.Close()
        End Using

        Return strNewPO
    End Function

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click

        If btnSubMenu.Text = "BACK" And Session("GOTOStatus") <> "" Then
            Session.Remove("GOTOStatus")
            Session.Remove("LoadSupplier")
            Session.Remove("GOTOStatus")
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        Else
            Session.Remove("GOTOStatus")
            Session.Remove("LoadSupplier")
            Session.Remove("GOTOStatus")
            Response.Redirect("~/MainMenu.aspx")
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
                    'If Session("YA010Msg") = "" Then
                    Call up_GridLoad()
                    Call ColorGrid()
                    'End If

                    Call clsMsg.DisplayMessage(lblInfo, "1014", clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text

                    Session.Remove("ErrorData")

                Case "exitarea"

                    Exit Sub

                Case "kosong"

                    Call up_GridLoadWhenEventChange()

                Case "savedata"

                    Call up_SaveData()

                Case "gridloadupdate"
                    'If Session("YA010Msg") = "" Then
                    Call up_GridLoadUpdate()
                    'End If

                    Call clsMsg.DisplayMessage(lblInfo, "1014", clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text

                    Session.Remove("ErrorData")

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
                "    a.OrderNo1 SplitReffPONo, " & vbCrLf & _
                "    CONVERT(NUMERIC(18,0), B.Week1) SplitReffQty " & vbCrLf & _
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

            txtOrderNo.Text = up_CreateSplitNo(txtpono.Text, cboAffiliate.Text, ls_supplier)

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
                "    MOQ = CONVERT(NUMERIC(18,0), ISNULL(b.POMOQ,MPM.MOQ)), " & vbCrLf & _
                "    QtyBox = CONVERT(NUMERIC(18,0), ISNULL(b.POQtyBox,MPM.QtyBox)), " & vbCrLf & _
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
                "    a.SplitReffPONo, " & vbCrLf & _
                "    CONVERT(NUMERIC(18,0), B.SplitReffQty) SplitReffQty " & vbCrLf & _
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

    Private Function uf_GetMOQ(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim MOQ As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(MOQ,0) MOQ FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
            dt = uf_GetDataTable(ls_SQL, Cn)
            If dt.Rows.Count > 0 Then
                MOQ = dt.Rows(0)("MOQ")
            End If
        End Using
        Return MOQ
    End Function

    Private Function uf_GetQtybox(ByVal pPartNo As String, ByVal pSupplierID As String, ByVal pAffiliateID As String) As Integer
        Dim Qty As Integer = 0
        Dim dt As New DataTable
        Using Cn As New SqlConnection(clsGlobal.ConnectionString)
            Dim ls_SQL As String
            ls_SQL = "SELECT ISNULL(QtyBox,0) Qty FROM dbo.MS_PartMapping WHERE PartNo='" + pPartNo + "' AND SupplierID='" + pSupplierID + "' AND AffiliateID='" + pAffiliateID + "'"
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
#End Region

End Class