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

Public Class POExportEntryMonthly

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
                Call up_fillcombo()

                Session("MenuDesc") = "INPUT PO MANUAL"
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
                    param = Request.QueryString("prm").ToString
                    If param = "  'back'" Then
                        btnSubMenu.Text = "BACK"
                    Else
                        If pStatus = False Then
                            Session.Remove("LoadSupplier")
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
                                'btnSplit.Enabled = False
                            Else
                                rdMonthly.Checked = True
                                'btnSplit.Enabled = True
                            End If

                            If pShipBy = "B" Then
                                rdrShipBy2.Checked = True
                            Else
                                rdrShipBy3.Checked = True
                            End If

                            If pSplitRefPONo = "" Then
                                'btnSplit.Enabled = True
                                btnRecover.Enabled = False
                            ElseIf pSplitStatus = Session("GOTOStatus") Then
                                'btnSplit.Enabled = False
                                btnRecover.Enabled = True
                            Else
                                btnRecover.Enabled = False
                            End If

                            cboDelLoc.ReadOnly = False

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
                            Session("LoadSupplier") = pSupplierCode
                            pStatus = True
                            txtconsignee.Text = pConsignee

                            Call up_GridLoad()
                            Session("pCheckError") = "1"

                            Session("pFilter") = pFilter
                            Session.Remove("EmergencyUrl")
                            btnSubMenu.Text = "BACK"
                        End If
                    End If

                ElseIf param <> "" And Session("GOTOStatus") = "satu" Then
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
                                'btnSplit.Enabled = False
                            Else
                                rdMonthly.Checked = True
                                'btnSplit.Enabled = True
                            End If

                            If pShipBy = "B" Then
                                rdrShipBy2.Checked = True
                            Else
                                rdrShipBy3.Checked = True
                            End If

                            cboDelLoc.ReadOnly = True

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
                            Session("LoadSupplier") = pSupplierCode
                            pStatus = True

                            Call up_GridLoad()
                            Session("pCheckError") = "1"

                            Session("pFilter") = pFilter
                            Session.Remove("EmergencyUrl")
                            btnSubMenu.Text = "BACK"
                        End If
                    End If

                ElseIf param <> "" Then
                    param = Request.QueryString("prm").ToString

                    btnApprove.Enabled = False
                    If Session("GOTOStatus") = 7 Then
                        btnSplit.Enabled = False
                        btnCancel.Enabled = False
                    'ElseIf Session("GOTOStatus") = 5 Or Session("GOTOStatus") = 6 Then
                    '    btnSplit.Enabled = False
                    Else
                        btnSplit.Enabled = True
                        btnCancel.Enabled = True
                    End If

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
                                'btnSplit.Enabled = False
                            Else
                                rdMonthly.Checked = True
                                'btnSplit.Enabled = True
                            End If

                            If pShipBy = "B" Then
                                rdrShipBy2.Checked = True
                            Else
                                rdrShipBy3.Checked = True
                            End If

                            If pSplitRefPONo = "" Then
                                'btnSplit.Enabled = True
                                btnRecover.Enabled = False
                            ElseIf pSplitStatus = Session("GOTOStatus") Then
                                'btnSplit.Enabled = False
                                btnRecover.Enabled = True
                            Else
                                btnRecover.Enabled = False
                            End If

                            cboDelLoc.ReadOnly = True

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
                            Session("LoadSupplier") = pSupplierCode
                            pStatus = True

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

    Private Sub up_CheckData()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            If grid.VisibleRowCount = 0 Then Exit Sub
            '***
            Dim ls_SQL As String = "", ls_MsgID As String = ""
            Dim ls_Active As String = "", iLoop As Long = 0, iCheckLoop As Long = 0
            Dim ls_PONo As String = "", ls_Affiliate As String = "", ls_Supplier As String = "", ls_PartNo As String = ""
            Dim ls_Week1 As String = "", ls_TotalPOQty As String = "", ls_MOQ As String = ""
            Dim ls_PreviousForecast As String = "", ls_Forecast1 As String = ""
            Dim ls_Forecast2 As String = "", ls_Forecast3 As String = ""
            Dim ls_Variance As String = "", ls_VariancePercentage As String = ""
            Dim ls_AdaData As String = ""
            Dim ls_error As String = ""
            Dim ls_QtyBox As Integer = 0
            'Dim li_Col As Integer
            '***
            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
                Dim SqlComm As New SqlCommand
                Dim ls_sql1 As String = ""
                Dim iRow As Integer = 0
                Dim user As String = Trim(Session("UserID").ToString)

                For iRow = 0 To grid.VisibleRowCount - 1
                    If Trim(grid.GetRowValues(iRow, "AllowAccess").ToString()) = 1 Then
                        ls_PONo = Trim(txtOrderNo.Text)
                        ls_Affiliate = Trim(cboAffiliate.Text)
                        ls_PartNo = Trim(grid.GetRowValues(iRow, "PartNo").ToString())
                        ls_Week1 = Trim(grid.GetRowValues(iRow, "Week1").ToString())
                        ls_MOQ = Replace(Trim(grid.GetRowValues(iRow, "MOQ").ToString()), ".00", "")
                        ls_QtyBox = Trim(grid.GetRowValues(iRow, "QtyBox").ToString())
                        ls_TotalPOQty = Trim(grid.GetRowValues(iRow, "Week1").ToString())
                        ls_PreviousForecast = Trim(grid.GetRowValues(iRow, "PreviousForecast").ToString())
                        ls_Forecast1 = Trim(grid.GetRowValues(iRow, "Forecast1").ToString())
                        ls_Forecast2 = Trim(grid.GetRowValues(iRow, "Forecast2").ToString())
                        ls_Forecast3 = Trim(grid.GetRowValues(iRow, "Forecast3").ToString())
                        ls_Variance = Trim(grid.GetRowValues(iRow, "Variance").ToString())
                        ls_VariancePercentage = Trim(grid.GetRowValues(iRow, "VariancePercentage").ToString())
                        ls_AdaData = Trim(grid.GetRowValues(iRow, "AdaData").ToString())

                        'MOQ
                        If ls_TotalPOQty <> 0 Or ls_MOQ <> 0 Then
                            If (ls_TotalPOQty < ls_MOQ) Then
                                Call clsMsg.DisplayMessage(lblInfo, "Total Firm Qty must be same or bigger then MOQ !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text

                            End If
                        End If

                        'Qty Box
                        If ls_TotalPOQty <> 0 Or ls_QtyBox <> 0 Then
                            If (ls_TotalPOQty Mod ls_QtyBox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Total Firm Qty must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If
                        If ls_Forecast1 <> 0 And ls_QtyBox <> 0 Then
                            If (ls_Forecast1 Mod ls_QtyBox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Forecast must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If
                        If ls_Forecast2 <> 0 And ls_QtyBox <> 0 Then
                            If (ls_Forecast2 Mod ls_QtyBox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Forecast must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If
                        If ls_Forecast3 <> 0 And ls_QtyBox <> 0 Then
                            If (ls_Forecast3 Mod ls_QtyBox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Forecast must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If

                        If Session("ErrorData") <> "" Then
                            ls_TampungError = ls_TampungError + 1
                            Session("JumlahError") = ls_TampungError
                        End If
                    End If
                Next iRow
                sqlTran.Commit()

            End Using
            sqlConn.Close()
        End Using
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
        Dim sqlComm As New SqlCommand

        Session.Remove("ErrorData")

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If HF.Get("hfTest") = "1" Then
                ls_SQL = "Select * from PO_Master_Export where AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                         " AND PONo = '" & Trim(txtpono.Text) & "' and excelcls = '2' " & vbCrLf
                If txtOrderNo.Text <> "" Then
                    ls_SQL = ls_SQL + " AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "'"
                End If

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    Call clsMsg.DisplayMessage(lblInfo, 5003, clsMessage.MsgType.ErrorMessage)
                    Session("ErrorData") = lblInfo.Text
                    sqlConn.Close()
                    Exit Sub
                End If
            End If

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
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

                'Cek Data
                Session("UbahDataGrid") = "1"

                Dim CheckData As Integer
                CheckData = e.UpdateValues.Count
                For iCheckLoop = 0 To CheckData - 1

                    ls_Active = (e.UpdateValues(iCheckLoop).NewValues("AllowAccess").ToString())
                    If ls_Active = True Then ls_Active = "1" Else ls_Active = "0"

                    ls_OrderNo = Trim(txtOrderNo.Text)
                    ls_PONo = Trim(txtpono.Text)
                    ls_FWD = Trim(cboDelLoc.Text)
                    ls_Affiliate = Trim(cboAffiliate.Text)
                    ls_Supplier = Trim(e.UpdateValues(iCheckLoop).NewValues("SupplierID").ToString())
                    ls_PartNo = Trim(e.UpdateValues(iCheckLoop).NewValues("PartNo").ToString())
                    ls_Week1 = Trim(e.UpdateValues(iCheckLoop).NewValues("Week1").ToString())
                    ls_MOQ = Replace(Trim(e.UpdateValues(iCheckLoop).NewValues("MOQ").ToString()), ".00", "")
                    ls_qtybox = Trim(e.UpdateValues(iCheckLoop).NewValues("QtyBox").ToString())
                    ls_TotalPOQty = Trim(e.UpdateValues(iCheckLoop).NewValues("Week1").ToString())
                    ls_PreviousForecast = Trim(e.UpdateValues(iCheckLoop).NewValues("PreviousForecast").ToString())
                    ls_Forecast1 = Trim(e.UpdateValues(iCheckLoop).NewValues("Forecast1").ToString())
                    ls_Forecast2 = Trim(e.UpdateValues(iCheckLoop).NewValues("Forecast2").ToString())
                    ls_Forecast3 = Trim(e.UpdateValues(iCheckLoop).NewValues("Forecast3").ToString())
                    ls_Variance = Trim(e.UpdateValues(iCheckLoop).NewValues("Variance").ToString())
                    ls_VariancePercentage = Trim(e.UpdateValues(iCheckLoop).NewValues("VariancePercentage").ToString())
                    ls_AdaData = Trim(e.UpdateValues(iCheckLoop).NewValues("AdaData").ToString())

                    If ls_Active = "1" Then
                        'MOQ
                        If ls_TotalPOQty <> 0 Or ls_MOQ <> 0 Then
                            If (ls_TotalPOQty < ls_MOQ) Then
                                Call clsMsg.DisplayMessage(lblInfo, "Total Firm Qty must be same or bigger then MOQ !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                                'grid..BackColor = Color.Red
                            End If
                        End If

                        'Qty Box
                        If ls_TotalPOQty <> 0 Or ls_qtybox <> 0 Then
                            If (ls_TotalPOQty Mod ls_qtybox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Total Firm Qty must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                                'grid.GetRow(iCheckLoop).CellStyle.BackColor = Color.Red

                            End If
                        End If
                        If ls_Forecast1 <> 0 And ls_qtybox <> 0 Then
                            If (ls_Forecast1 Mod ls_qtybox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Forecast must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If
                        If ls_Forecast2 <> 0 And ls_qtybox <> 0 Then
                            If (ls_Forecast2 Mod ls_qtybox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Forecast must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If
                        If ls_Forecast3 <> 0 And ls_qtybox <> 0 Then
                            If (ls_Forecast3 Mod ls_qtybox) <> 0 Then
                                Call clsMsg.DisplayMessage(lblInfo, "Forecast must be same or multiple of the Qty Box !", clsMessage.MsgType.ErrorMessage)
                                Session("ErrorData") = lblInfo.Text
                            End If
                        End If

                        If Session("ErrorData") <> "" Then
                            ls_TampungError = ls_TampungError + 1
                            Session("JumlahError") = ls_TampungError
                        End If

                        If ls_TampungError > 0 Then
                            ls_SQL = " 	INSERT INTO dbo.PO_Tampung_Detail_Export " & vbCrLf & _
                                         " 	        (PONo, OrderNo1, ForwarderID ,AffiliateID ,SupplierID , PartNo, Week1, TotalPOQty, PreviousForecast, " & vbCrLf & _
                                         " 	        Forecast1 ,Forecast2 ,Forecast3 ,Variance , VariancePercentage, EntryDate ,EntryUser) " & vbCrLf & _
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
                                         " 	          GETDATE(), " & vbCrLf & _
                                         " 	          '" & Session("UserID").ToString & "' " & vbCrLf & _
                                         " 	        ) "

                            Dim sqlComm1 As New SqlCommand
                            sqlComm1 = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                            sqlComm1.ExecuteNonQuery()
                            sqlComm1.Dispose()
                        End If

                    End If

                Next iCheckLoop

                If HF.Get("hfTest") = "1" Then
                    sqlTran.Commit()
                End If

                If ls_TampungError > 0 Then
                    Exit Sub
                End If

                'End If

                If HF.Get("hfTest") = "2" Then
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
                    Dim a As Integer
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
                                    ls_SQL = " INSERT INTO dbo.PO_Master_Export " & _
                                                "(PONo, AffiliateID, SupplierID, ForwarderID, Period, CommercialCls, EmergencyCls, " & vbCrLf & _
                                                " ShipCls, OrderNo1, ETDVendor1, ETDPort1, ETAPort1, ETAFactory1, EntryDate, EntryUser)" & _
                                                " VALUES ('" & Trim(ls_PONo) & "'," & vbCrLf & _
                                                " '" & Trim(cboAffiliate.Text) & "'," & vbCrLf & _
                                                " '" & Trim(ls_Supplier) & "'," & vbCrLf & _
                                                " '" & Trim(cboDelLoc.Text) & "'," & vbCrLf & _
                                                " '" & Convert.ToDateTime(dtPeriodFrom.Value).ToString("yyyy-MM-01") & "'," & vbCrLf & _
                                                " '" & Trim(ls_Commercial) & "'," & vbCrLf & _
                                                " '" & Trim(ls_EmergencyCls) & "'," & vbCrLf & _
                                                " '" & Trim(ls_ShipCls) & "'," & vbCrLf & _
                                                " '" & Trim(ls_OrderNo) & "' ," & vbCrLf & _
                                                " '" & Convert.ToDateTime(dtETDVendor.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                                " '" & Convert.ToDateTime(dtETDPort.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                                " '" & Convert.ToDateTime(dtETAPort.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                                " '" & Convert.ToDateTime(dtETAFactory.Value).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                                " GETDATE()," & vbCrLf & _
                                                " '" & Session("UserID").ToString & "')" & vbCrLf
                                    ls_MsgID = "1001"
                                End If
                                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                                sqlComm.ExecuteNonQuery()
                                sqlComm.Dispose()

                                If pIsUpdate = False Then
                                    'INSERT DATA
                                    ls_SQL = " 	INSERT INTO dbo.PO_Detail_Export " & vbCrLf & _
                                             " 	        (PONo, OrderNo1, ForwarderID ,AffiliateID ,SupplierID , PartNo, Week1, TotalPOQty, PreviousForecast, " & vbCrLf & _
                                             " 	        Forecast1 ,Forecast2 ,Forecast3 ,Variance , VariancePercentage, EntryDate ,EntryUser) " & vbCrLf & _
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
                                             " 	          GETDATE(), " & vbCrLf & _
                                             " 	          '" & Session("UserID").ToString & "' " & vbCrLf & _
                                             " 	        ) "
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
                                             " 	 WHERE PONo ='" & Trim(ls_PONo) & "' AND OrderNo1 = '" & ls_OrderNo & "' AND AffiliateID = '" & Trim(ls_Affiliate) & "' AND SupplierID = '" & Trim(ls_Supplier) & "' AND PartNo = '" & Trim(ls_PartNo) & "'"
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
                                         "  AND AffiliateID = '" & Trim(ls_Affiliate) & "' " & vbCrLf & _
                                         "  AND SupplierID = '" & Trim(ls_Supplier) & "' " & vbCrLf & _
                                         "  AND OrderNo1 = '" & ls_OrderNo & "' " & vbCrLf & _
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

NextLoop:
                    Next iLoop

                    If ls_TampungError = 0 Then
                        Session("DataTersimpan") = "1"
                    ElseIf ls_TampungError > 0 Then
                        ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
                        Session("ErrorData") = lblInfo.Text
                        Exit Sub
                    End If

                    sqlTran.Commit()

                End If

            End Using

            sqlConn.Close()
        End Using


        Call ColorGrid()
        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
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

                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
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
                                         "         AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "'" & vbCrLf & _
                                         " UPDATE PO_Detail_Export " & vbCrLf & _
                                         " SET     ForwarderID = '" & Trim(cboDelLoc.Text) & "'," & vbCrLf & _
                                         "         WHERE PONo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                                         "         AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                                         "         AND SupplierID = '" & Trim(grid.GetRowValues(i, "SupplierID").ToString) & "' " & vbCrLf & _
                                         "         AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "'"
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

    Private Sub up_AddItem()
        Dim pIsUpdate As Boolean

        Try
            Dim ls_SQL As String = "", ls_MsgID As String = ""

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                Session.Remove("ErrorData")

                sqlConn.Open()
                Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")

                    ls_SQL = "SELECT * FROM PO_Detail_Export " & vbCrLf & _
                                "WHERE PONo = '" & Trim(txtpono.Text) & "' " & vbCrLf & _
                                "AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf & _
                                "AND SupplierID = '" & Trim(txtSupplier.Text) & "' " & vbCrLf & _
                                "AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf & _
                                "AND PartNo = '" & Trim(txtPartNo.Text) & "' "

                    Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                    Dim sqlRdr As SqlDataReader = sqlComm.ExecuteReader()

                    If sqlRdr.Read Then
                        pIsUpdate = True
                    Else
                        pIsUpdate = False
                    End If
                    sqlRdr.Close()

                    If pIsUpdate Then
                        ls_MsgID = "6018"

                        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.ErrorMessage)
                        Session("ErrorData") = lblInfo.Text

                        Exit Sub
                    Else
                        'Insert data
                        ls_SQL = " 	INSERT INTO dbo.PO_Detail_Export " & vbCrLf & _
                                 " 	        (PONo, OrderNo1, ForwarderID ,AffiliateID ,SupplierID , PartNo, Week1, TotalPOQty, PreviousForecast, " & vbCrLf & _
                                 " 	        Forecast1 ,Forecast2 ,Forecast3 ,Variance , VariancePercentage, EntryDate ,EntryUser) " & vbCrLf & _
                                 " 	VALUES  ( '" & Trim(txtpono.Text) & "', " & vbCrLf & _
                                 " 	          '" & Trim(txtOrderNo.Text) & "', " & vbCrLf & _
                                 " 	          '" & Trim(cboDelLoc.Text) & "', " & vbCrLf & _
                                 " 	          '" & Trim(cboAffiliate.Text) & "', " & vbCrLf & _
                                 " 	          '" & Trim(txtSupplier.Text) & "', " & vbCrLf & _
                                 " 	          '" & Trim(txtPartNo.Text) & "', " & vbCrLf & _
                                 " 	          '" & txtTotFirmQty.Text & "', " & vbCrLf & _
                                 " 	          '" & txtTotFirmQty.Text & "', " & vbCrLf & _
                                 " 	          '" & txtPrevForecast.Text & "', " & vbCrLf & _
                                 " 	          '" & txtForcast1.Text & "', " & vbCrLf & _
                                 " 	          '" & txtForcast2.Text & "', " & vbCrLf & _
                                 " 	          '" & txtForcast3.Text & "', " & vbCrLf & _
                                 " 	          '" & txtVariance.Text & "', " & vbCrLf & _
                                 " 	          '" & txtVariancePerc.Text & "', " & vbCrLf & _
                                 " 	          GETDATE(), " & vbCrLf & _
                                 " 	          '" & Session("UserID").ToString & "' " & vbCrLf & _
                                 " 	        ) "

                        ls_MsgID = "1002"

                        sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                        sqlComm.ExecuteNonQuery()
                        sqlComm.Dispose()
                    End If

                    sqlTran.Commit()
                End Using

                sqlConn.Close()
            End Using

            Call ColorGrid()
            Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
            Session("ErrorData") = lblInfo.Text

        Catch ex As Exception
            Me.lblInfo.Visible = True
            Me.lblInfo.Text = ex.Message.ToString
        End Try
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click

        If btnSubMenu.Text = "BACK" And Session("GOTOStatus") <> "" Then
            Session.Remove("GOTOStatus")
            Session.Remove("LoadSupplier")
            Session.Remove("GOTOStatus")
            Response.Redirect("~/PurchaseOrderExport/POExportList.aspx")
        ElseIf btnSubMenu.Text = "BACK" And Session("GOTOStatus") <> "" Then
            Session.Remove("GOTOStatus")
            Session.Remove("LoadSupplier")
            Session.Remove("GOTOStatus")
            Response.Redirect("~/PurchaseOrderExport/POExportFinalApprovalList.aspx")
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
            grid.JSProperties("cpMessage") = Session("YA010IsSubmit")

            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "loaddata"

                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                Case "recoveryposplit"

                    Call up_GridLoadRecovery()

                    grid.JSProperties("cpOrderNo") = Session("SplitReffPONo")

                    Call clsMsg.DisplayMessage(lblInfo, "1016", clsMessage.MsgType.InformationMessage)
                    grid.JSProperties("cpMessage") = lblInfo.Text

                Case "gridload"

                    'If Not IsNothing(Session("DataTersimpan")) = True Then
                    '    Call up_SaveData()
                    '    Session.Remove("DataTersimpan")
                    'ElseIf IsNothing(Session("DataTersimpan")) = True Then
                    '    Call clsMsg.DisplayMessage(lblInfo, "7012", clsMessage.MsgType.InformationMessage)
                    '    lblInfo.Text = lblInfo.Text
                    '    Exit Sub
                    'End If
                    Call up_SaveData()
                    Session.Remove("DataTersimpan")

                    Call up_GridLoad()
                    Call ColorGrid()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        grid.JSProperties("cpMessage") = lblInfo.Text
                    End If

                    ASPxCallback1.JSProperties("cpMessage") = Session("ErrorData")

                    If Session("ErrorData") <> "" Then
                        ASPxCallback1.JSProperties("cpJumlahError") = Session("JumlahError")
                        Session.Remove("ErrorData")
                        Exit Sub
                    End If
                    Session.Remove("ErrorData")

                Case "exitarea"

                    Exit Sub

                Case "loaddatacell"

                    Session.Remove("pCheckError")
                    Call up_GridLoadCekData()

                Case "kosong"

                    Call up_GridLoadWhenEventChange()

                Case "savedata"

                    Call up_SaveData()

                Case "additem"

                    Call up_AddItem()

                Case "downloadSummary"

                    Dim psERR As String = ""
                    Dim pSuppType As String = ""

                    Call up_GridLoadCekData()
                    FileName = "TemplatePOExportMonthly.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If grid.VisibleRowCount > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", "A:3", psERR)
                    End If

                Case "saveApprove"

                    Call uf_Approve()

            End Select

EndProcedure:
            Session("YA010IsSubmit") = ""


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
                "    Week1 = CONVERT(NUMERIC(18,0), B.Week1), " & vbCrLf & _
                "    B.Week2, " & vbCrLf & _
                "    B.Week3, " & vbCrLf & _
                "    B.Week4, " & vbCrLf & _
                "    B.Week5, " & vbCrLf & _
                "    TotalPOQty = CONVERT(NUMERIC(18,0), B.Week1), " & vbCrLf & _
                "    PreviousForecast = CASE WHEN a.EmergencyCls = 'E' then 0 else ISNULL(PrevQty.Forecast1,0) end, " & vbCrLf & _
                "    B.Forecast1, " & vbCrLf & _
                "    B.Forecast2, " & vbCrLf & _
                "    B.Forecast3, " & vbCrLf & _
                "    Variance = CASE WHEN a.EmergencyCls = 'E' then 0 else CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE B.Week1 - PrevQty.Forecast1 END END, " & vbCrLf & _
                "    VariancePercentage = CASE WHEN a.EmergencyCls = 'E' then 0 else CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE ((B.Week1 - PrevQty.Forecast1) / PrevQty.Forecast1) * 100 END END, " & vbCrLf & _
                "    a.PONo, " & vbCrLf & _
                "    a.ShipCls, " & vbCrLf & _
                "    a.CommercialCls, " & vbCrLf & _
                "    a.ForwarderID, " & vbCrLf & _
                "    a.Period, " & vbCrLf & _
                "    RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                "    RTRIM(a.SupplierID)SupplierID, " & vbCrLf & _
                "    ErrorStatus = ISNULL(UPO.errorCls,'') " & vbCrLf & _
                "    FROM PO_Master_Export a " & vbCrLf & _
                "    INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = B.AffiliateID AND a.SupplierID = B.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                "    LEFT JOIN ( " & vbCrLf & _
                "       SELECT Forecast1, PartNo, a.AffiliateID, a.PONo, a.OrderNo1 FROM PO_Detail_Export a " & vbCrLf & _
                "       INNER JOIN PO_Master_Export b ON a.PONo = b.PONo and a.OrderNo1 = b.OrderNo1 and a.AffiliateID = b.AffiliateID  and a.SupplierID = b.SupplierID " & vbCrLf & _
                "       WHERE Period = '" & DateAdd(DateInterval.Month, -1, dtPeriodFrom.Value) & "' and a.PONo = a.PONo and b.EmergencyCls <> 'E' and a.OrderNo1 = b.OrderNo1 and Forecast1 > 0" & vbCrLf & _
                "    )PrevQty ON PrevQty.PartNo = b.PartNo and PrevQty.AffiliateID = b.AffiliateID --and PrevQty.PONo = b.PONo and PrevQty.OrderNo1 = b.OrderNo1" & vbCrLf & _
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

    Private Sub up_GridLoadRecovery()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_supplier As String = ""
        Dim ls_splitreffpono As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            If Session("LoadSupplier") <> "" Then ls_supplier = Session("LoadSupplier")
            If Session("SplitReffPONo") <> "" Then ls_splitreffpono = Session("SplitReffPONo")

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
                "    Week1 = CONVERT(NUMERIC(18,0), B.Week1), " & vbCrLf & _
                "    B.Week2, " & vbCrLf & _
                "    B.Week3, " & vbCrLf & _
                "    B.Week4, " & vbCrLf & _
                "    B.Week5, " & vbCrLf & _
                "    TotalPOQty = CONVERT(NUMERIC(18,0), B.Week1), " & vbCrLf & _
                "    PreviousForecast = ISNULL(PrevQty.Forecast1,0), " & vbCrLf & _
                "    B.Forecast1, " & vbCrLf & _
                "    B.Forecast2, " & vbCrLf & _
                "    B.Forecast3, " & vbCrLf & _
                "    Variance = CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE B.Week1 - PrevQty.Forecast1 END, " & vbCrLf & _
                "    VariancePercentage = CASE WHEN ISNULL(PrevQty.Forecast1,0) = 0 THEN 0 ELSE ((B.Week1 - PrevQty.Forecast1) / PrevQty.Forecast1) * 100 END, " & vbCrLf & _
                "    a.PONo, " & vbCrLf & _
                "    a.ShipCls, " & vbCrLf & _
                "    a.CommercialCls, " & vbCrLf & _
                "    a.ForwarderID, " & vbCrLf & _
                "    a.Period, " & vbCrLf & _
                "    RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                "    RTRIM(a.SupplierID)SupplierID, " & vbCrLf & _
                "    ErrorStatus = ISNULL(UPO.errorCls,'') " & vbCrLf & _
                "    FROM PO_Master_Export a " & vbCrLf & _
                "    INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = B.AffiliateID AND a.SupplierID = B.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
                "    LEFT JOIN ( " & vbCrLf & _
                "       SELECT Forecast1, PartNo, a.AffiliateID, a.PONo, a.OrderNo1 FROM PO_Detail_Export a " & vbCrLf & _
                "       INNER JOIN PO_Master_Export b ON a.PONo = b.PONo and a.OrderNo1 = b.OrderNo1 and a.AffiliateID = b.AffiliateID  and a.SupplierID = b.SupplierID " & vbCrLf & _
                "       WHERE Period = '" & DateAdd(DateInterval.Month, -1, dtPeriodFrom.Value) & "'" & vbCrLf & _
                "    )PrevQty ON PrevQty.PartNo = b.PartNo and PrevQty.AffiliateID = b.AffiliateID --and a.PONo = b.PONo and a.OrderNo1 = b.OrderNo1" & vbCrLf & _
                "    LEFT JOIN MS_Parts c ON c.PartNo = B.PartNo " & vbCrLf & _
                "    LEFT JOIN MS_PartMapping MPM ON MPM.PartNo = b.PartNo AND MPM.AffiliateID = b.AffiliateID AND MPM.SupplierID= b.SupplierID " & vbCrLf & _
                "    LEFT JOIN MS_UnitCls d ON d.UnitCls = c.UnitCls " & vbCrLf & _
                "    LEFT JOIN UploadPOExport UPO ON UPO.PONo = a.Pono AND a.AffiliateID = UPO.AffiliateID AND UPO.SupplierID = a.supplierID AND UPO.ForwarderID = a.ForwarderID AND UPO.Partno = b.PartNo " & vbCrLf & _
                "    WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf

            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "    AND a.OrderNo1 = '" & ls_splitreffpono & "' " & vbCrLf
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

    Private Sub up_ItemLoad()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_supplier As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            Session.Remove("ErrorData")

            sqlConn.Open()

            If Session("LoadSupplier") <> "" Then ls_supplier = Session("LoadSupplier")

            ls_SQL = "SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY AllowAccess DESC, PartNo, AffiliateID, SupplierID) NoUrut, * " & vbCrLf & _
                "FROM(" & vbCrLf & _
                "    SELECT " & vbCrLf & _
                "    '0' AllowAccess, " & vbCrLf & _
                "    '0' AdaData, " & vbCrLf & _
                "    RTRIM(a.PartNo)PartNo, " & vbCrLf & _
                "    RTRIM(b.PartName)PartName, " & vbCrLf & _
                "    RTRIM(c.Description)UOM, " & vbCrLf & _
                "    MOQ = CONVERT(NUMERIC(18,0), a.MOQ), " & vbCrLf & _
                "    QtyBox = CONVERT(NUMERIC(18,0), a.QtyBox), " & vbCrLf & _
                "    Week1 = CONVERT(NUMERIC(18,0), '0'), " & vbCrLf & _
                "    '0' Week2, " & vbCrLf & _
                "    '0' Week3, " & vbCrLf & _
                "    '0' Week4, " & vbCrLf & _
                "    '0' Week5, " & vbCrLf & _
                "    TotalPOQty = CONVERT(NUMERIC(18,0),'0'), " & vbCrLf & _
                "    PreviousForecast = ISNULL((SELECT CONVERT(NUMERIC(18,0),qty) qty FROM MS_Forecast MF WHERE MF.PartNo = a.PartNo AND a.AffiliateID = MF.AffiliateID AND YEAR(Period) = Year(DATEADD(MONTH,-1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) AND MONTH(Period) = MONTH(DATEADD(MONTH,-1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                "    Forecast1 = ISNULL((SELECT qty FROM MS_Forecast MF WHERE MF.PartNo = a.PartNo AND a.AffiliateID = MF.AffiliateID AND YEAR(Period) = Year(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "' )) AND MONTH(Period) = MONTH(DATEADD(MONTH,1,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                "    Forecast2 = ISNULL((SELECT qty FROM MS_Forecast MF WHERE MF.PartNo = a.PartNo AND a.AffiliateID = MF.AffiliateID AND YEAR(Period) = Year(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) AND MONTH(Period) = MONTH(DATEADD(MONTH,2,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                "    Forecast3 = ISNULL((SELECT qty FROM MS_Forecast MF WHERE MF.PartNo = a.PartNo AND a.AffiliateID = MF.AffiliateID AND YEAR(Period) = Year(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "')) AND MONTH(Period) = MONTH(DATEADD(MONTH,3,'" & Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "'))),0), " & vbCrLf & _
                "    '0' Variance, " & vbCrLf & _
                "    '0' VariancePercentage, " & vbCrLf & _
                "    '' PONo, " & vbCrLf & _
                "    '' ShipCls, " & vbCrLf & _
                "    '' CommercialCls, " & vbCrLf & _
                "    '' ForwarderID, " & vbCrLf & _
                "    '' Period, " & vbCrLf & _
                "    RTRIM(a.AffiliateID)AffiliateID, " & vbCrLf & _
                "    RTRIM(a.SupplierID)SupplierID, ErrorStatus = '' " & vbCrLf & _
                "    FROM MS_PartMapping a " & vbCrLf & _
                "    INNER JOIN MS_Parts b ON a.PartNo = b.PartNo " & vbCrLf & _
                "    LEFT JOIN MS_UnitCls c ON c.UnitCls = b.UnitCls " & vbCrLf & _
                "    WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf

            If Trim(ls_supplier) <> "" Then
                ls_SQL = ls_SQL + _
                    "    AND a.SupplierID = '" & Trim(ls_supplier) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + _
                "    AND a.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf & _
                "    --AND NOT EXISTS( " & vbCrLf & _
                "    --    SELECT * FROM PO_Detail_Export X " & vbCrLf & _
                "    --    WHERE X.pono = '" & Trim(txtpono.Text) & "' AND X.PartNo = a.PartNo " & vbCrLf & _
                "    --) " & vbCrLf & _
                ")X "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                txtPartNo.Text = ds.Tables(0).Rows(0)("PartNo") & ""
                txtPartName.Text = ds.Tables(0).Rows(0)("PartName") & ""
                txtUOM.Text = ds.Tables(0).Rows(0)("UOM") & ""
                txtMOQ.Text = ds.Tables(0).Rows(0)("MOQ") & ""
                txtQtyBox.Text = ds.Tables(0).Rows(0)("QtyBox") & ""
                txtTotFirmQty.Text = ds.Tables(0).Rows(0)("TotalPOQty") & ""
                txtPrevForecast.Text = ds.Tables(0).Rows(0)("PreviousForecast") & ""
                txtVariance.Text = ds.Tables(0).Rows(0)("Variance") & ""
                txtVariancePerc.Text = ds.Tables(0).Rows(0)("VariancePercentage") & ""
                txtForcast1.Text = ds.Tables(0).Rows(0)("Forecast1") & ""
                txtForcast2.Text = ds.Tables(0).Rows(0)("Forecast2") & ""
                txtForcast3.Text = ds.Tables(0).Rows(0)("Forecast3") & ""
                txtSupplier.Text = ds.Tables(0).Rows(0)("SupplierID") & ""
            Else
                txtPartName.Text = ""
                txtUOM.Text = ""
                txtMOQ.Text = ""
                txtQtyBox.Text = ""
                txtTotFirmQty.Text = ""
                txtPrevForecast.Text = ""
                txtVariance.Text = ""
                txtVariancePerc.Text = ""
                txtForcast1.Text = ""
                txtForcast2.Text = ""
                txtForcast3.Text = ""
                txtSupplier.Text = ""

                If txtPartNo.Text <> "" Then
                    Call clsMsg.DisplayMessage(lblInfo, "[6011] Part No. doesn't exists !", clsMessage.MsgType.ErrorMessage)
                    Session("ErrorData") = lblInfo.Text
                End If
            End If

            ASPxCallback1.JSProperties("cpPartNo") = txtPartNo.Text
            ASPxCallback1.JSProperties("cpPartName") = txtPartName.Text
            ASPxCallback1.JSProperties("cpUOM") = txtUOM.Text
            ASPxCallback1.JSProperties("cpMOQ") = txtMOQ.Text
            ASPxCallback1.JSProperties("cpQtyBox") = txtQtyBox.Text
            ASPxCallback1.JSProperties("cpTotalPOQty") = txtTotFirmQty.Text
            ASPxCallback1.JSProperties("cpPreviousForecast") = txtPrevForecast.Text
            ASPxCallback1.JSProperties("cpVariance") = txtVariance.Text
            ASPxCallback1.JSProperties("cpVariancePercentage") = txtVariancePerc.Text
            ASPxCallback1.JSProperties("cpForecast1") = txtForcast1.Text
            ASPxCallback1.JSProperties("cpForecast2") = txtForcast2.Text
            ASPxCallback1.JSProperties("cpForecast3") = txtForcast3.Text
            ASPxCallback1.JSProperties("cpSupplierID") = txtSupplier.Text

            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_SendDataToSupplier()
        Dim ls_SQL As String = ""
        Dim ls_MsgID As String = ""
        Dim ls_supplier As String

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
                If txtOrderNo.Text = "" Then
                    ls_SQL = " 	UPDATE dbo.PO_Master_Export " & vbCrLf & _
                                            " 	   SET ExcelCls = '1' , " & vbCrLf & _
                                            " 	       PASISendToSupplierDate = GETDATE(), " & vbCrLf & _
                                            " 	       PASISendToSupplierUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                                            " 	       UpdateDate = GETDATE(), " & vbCrLf & _
                                            " 	       UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                            " 	 WHERE PONO = '" & Trim(txtpono.Text) & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND ForwarderID = '" & Trim(cboDelLoc.Text) & "'"
                    ls_MsgID = "1008"
                Else
                    If Session("LoadSupplier") <> "" Then ls_supplier = Session("LoadSupplier")
                    ls_SQL = " 	UPDATE dbo.PO_Master_Export " & vbCrLf & _
                                            " 	   SET ExcelCls = '1' , " & vbCrLf & _
                                            " 	       PASISendToSupplierDate = GETDATE(), " & vbCrLf & _
                                            " 	       PASISendToSupplierUser = '" & Session("UserID").ToString & "', " & vbCrLf & _
                                            " 	       UpdateDate = GETDATE(), " & vbCrLf & _
                                            " 	       UpdateUser = '" & Session("UserID").ToString & "' " & vbCrLf & _
                                            " 	 WHERE PONO = '" & Trim(txtpono.Text) & "' AND OrderNo1 = '" & Trim(txtOrderNo.Text) & "' AND AffiliateID = '" & Trim(cboAffiliate.Text) & "' AND SupplierID = '" & Trim(ls_supplier) & "' AND ForwarderID = '" & Trim(cboDelLoc.Text) & "'"
                    ls_MsgID = "1008"
                End If

                Dim sqlComm As New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm = New SqlCommand(ls_SQL, sqlConn, sqlTran)
                sqlComm.ExecuteNonQuery()
                sqlComm.Dispose()
                sqlTran.Commit()

            End Using
            sqlConn.Close()
        End Using

        Call clsMsg.DisplayMessage(lblInfo, ls_MsgID, clsMessage.MsgType.InformationMessage)
        ASPxCallback1.JSProperties("cpMessage") = lblInfo.Text
        Session("ErrorData") = lblInfo.Text
    End Sub

    Private Sub up_GridLoadCekData()
        Dim ls_SQL As String = ""
        Dim pWhere As String = ""
        Dim ls_supplier As String = ""

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            If Session("LoadSupplier") <> "" Then ls_supplier = Session("LoadSupplier")
            If ls_supplier = "" Then
                If txtOrderNo.Text = "" Then ls_supplier = "" Else ls_supplier = Replace(Trim(txtOrderNo.Text), Trim(txtpono.Text) & "-", "")
            End If

            ls_SQL = "   SELECT ROW_NUMBER() over (order by allowAccess desc, PartNo, AffiliateID, SupplierID) NoUrut, *   " & vbCrLf & _
                  "   FROM   " & vbCrLf & _
                  "   (   " & vbCrLf & _
                  "   SELECT    " & vbCrLf & _
                  "   	'0' AllowAccess,   " & vbCrLf & _
                  "   	'0' AdaData,   " & vbCrLf & _
                  "   	RTRIM(a.PartNo)PartNo,     " & vbCrLf & _
                  "   	RTRIM(b.PartName)PartName,     " & vbCrLf & _
                  "   	RTRIM(c.Description)UOM,     " & vbCrLf & _
                  "   	RTRIM(isnull(a.MOQ,0))MOQ,  QtyBox = convert(numeric(18,0),isnull(a.QtyBox,0)),      " & vbCrLf & _
                  "   	'0' Week1,      	'0' TotalPOQty,      " & vbCrLf

            ls_SQL = ls_SQL + "   	'0' PreviousForecast,    " & vbCrLf & _
                              "   	'0' Forecast1,    " & vbCrLf & _
                              "   	'0' Forecast2 ,    " & vbCrLf & _
                              "   	'0' Forecast3 ,    " & vbCrLf & _
                              "   	'0' Variance,      " & vbCrLf & _
                              "   	'0' VariancePercentage,      " & vbCrLf & _
                              "   	'' PONo, a.AffiliateID, a.SupplierID, ErrorStatus = ''  " & vbCrLf & _
                              "   FROM MS_PartMapping a     " & vbCrLf & _
                              "   	INNER join MS_Parts b on a.PartNo = b.PartNo     " & vbCrLf & _
                              "   	LEFT join MS_UnitCls c on c.UnitCls = b.UnitCls      		 " & vbCrLf & _
                              "   	WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "'  " & vbCrLf

            'SupplierID
            If Trim(ls_supplier) <> "" Then
                ls_SQL = ls_SQL + _
                    "          AND a.SupplierID = '" & Trim(ls_supplier) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + _
                              " AND NOT EXISTS    " & vbCrLf & _
                  "   		(    " & vbCrLf & _
                  "   		SELECT * FROM  PO_Detail_Export X    " & vbCrLf & _
                  "  		                WHERE X.PartNo = a.PartNo " & vbCrLf & _
                  "  		                and X.PONo = '" & Trim(txtpono.Text) & "'  " & vbCrLf & _
                  "   		)   " & vbCrLf & _
                  "  " & vbCrLf & _
                  "   UNION ALL  " & vbCrLf & _
                  "   SELECT    " & vbCrLf & _
                  "   	'1' AllowAccess,     "

            ls_SQL = ls_SQL + "   	CASE WHEN a.PartNo <> '' Then '1' ELSE '0' End AdaData,   	RTRIM(a.PartNo)PartNo,     " & vbCrLf & _
                              "   	RTRIM(C.PartName)PartName,     " & vbCrLf & _
                              "   	RTRIM(d.Description)UOM,     " & vbCrLf & _
                              "   	RTRIM(isnull(a.POMOQ,e.MOQ))MOQ,    QtyBox = convert(numeric(18,0),isnull(a.POQtyBox,e.QtyBox)),    " & vbCrLf & _
                              "   	a.Week1,    " & vbCrLf & _
                              "   	TotalPOQty = (a.Week1),      " & vbCrLf & _
                              "   	PreviousForecast = ISNULL(tde.PreviousForecast,a.PreviousForecast), " & vbCrLf & _
                              "   	Forecast1 = ISNULL(tde.Forecast1,a.Forecast1), " & vbCrLf & _
                              "   	Forecast2 = ISNULL(tde.Forecast2,a.Forecast2), " & vbCrLf & _
                              "   	Forecast3 = ISNULL(tde.Forecast3,a.Forecast3),  " & vbCrLf & _
                              "   	a.Variance,       	a.VariancePercentage,      "

            ls_SQL = ls_SQL + "   	a.PONo, a.AffiliateID, a.SupplierID, ErrorStatus = Isnull(UPO.errorCls,'')       " & vbCrLf & _
                              "   FROM PO_Detail_Export a  " & vbCrLf & _
                              "   LEFT JOIN PO_Tampung_Detail_Export tde  " & vbCrLf & _
                              "   on a.PONo = tde.PONo AND a.AffiliateID = tde.AffiliateID AND a.SupplierID = tde.SupplierID AND a.PartNo = tde.PartNo  " & vbCrLf & _
                              "     LEFT JOIN UploadPOExport UPO ON UPO.PONo = a.Pono and a.AffiliateID = UPO.AffiliateID" & vbCrLf & _
                              "     AND UPO.SupplierID = a.supplierID and UPO.ForwarderID = a.ForwarderID " & vbCrLf & _
                              "     and UPO.Partno = a.PartNo " & vbCrLf & _
                              "   	LEFT join MS_Parts c on c.PartNo = a.PartNo     " & vbCrLf & _
                              "   	LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls     " & vbCrLf & _
                              "   	LEFT join MS_PartMapping e on e.PartNo = a.PartNo and e.AffiliateID = a.AffiliateID and e.SupplierID = a.SupplierID     " & vbCrLf & _
                              "   WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf

            'SupplierID
            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "          AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            If Trim(ls_supplier) <> "" Then
                ls_SQL = ls_SQL + _
                    "          AND a.SupplierID = '" & Trim(ls_supplier) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "           AND a.PONo = '" & Trim(txtpono.Text) & "'  " & vbCrLf & _
                              "           --AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf & _
                              " UNION ALL  " & vbCrLf & _
                              "   SELECT    " & vbCrLf & _
                              "   	'1' AllowAccess,     " & vbCrLf & _
                              "   	CASE WHEN a.PartNo <> '' Then '1' ELSE '0' End AdaData,   	RTRIM(a.PartNo)PartNo,     " & vbCrLf & _
                              "   	RTRIM(C.PartName)PartName,     " & vbCrLf & _
                              "   	RTRIM(d.Description)UOM,     " & vbCrLf & _
                              "   	RTRIM(isnull(e.MOQ,0))MOQ,  QtyBox = convert(numeric(18,0),isnull(e.QtyBox,0)),      " & vbCrLf & _
                              "   	a.Week1,    " & vbCrLf & _
                              "   	TotalPOQty = (a.Week1),      " & vbCrLf & _
                              "   	PreviousForecast = ISNULL(a.PreviousForecast,a.PreviousForecast), " & vbCrLf & _
                              "   	Forecast1 = ISNULL(a.Forecast1,a.Forecast1), "

            ls_SQL = ls_SQL + "   	Forecast2 = ISNULL(a.Forecast2,a.Forecast2), " & vbCrLf & _
                              "   	Forecast3 = ISNULL(a.Forecast3,a.Forecast3),  " & vbCrLf & _
                              "   	a.Variance,       	a.VariancePercentage,      " & vbCrLf & _
                              "   	a.PONo, a.AffiliateID, a.SupplierID, ErrorStatus = ''      " & vbCrLf & _
                              "   FROM PO_Tampung_Detail_Export a  " & vbCrLf & _
                              "   	LEFT join MS_Parts c on c.PartNo = a.PartNo     " & vbCrLf & _
                              "   	LEFT join MS_UnitCls d on d.UnitCls = c.UnitCls " & vbCrLf & _
                              "   	LEFT join MS_PartMapping e on e.PartNo = a.PartNo and e.AffiliateID = a.AffiliateID and e.SupplierID = a.SupplierID " & vbCrLf & _
                              "   WHERE a.AffiliateID = '" & Trim(cboAffiliate.Text) & "' " & vbCrLf
            'SupplierID
            If Trim(txtOrderNo.Text) <> "" Then
                ls_SQL = ls_SQL + _
                    "          AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf
            End If

            If Trim(ls_supplier) <> "" Then
                ls_SQL = ls_SQL + _
                    "          AND a.SupplierID = '" & Trim(ls_supplier) & "' " & vbCrLf
            End If

            ls_SQL = ls_SQL + "           AND a.PONo = '" & Trim(txtpono.Text) & "'  " & vbCrLf & _
                              "           --AND a.OrderNo1 = '" & Trim(txtOrderNo.Text) & "' " & vbCrLf & _
                              "  	)  " & vbCrLf & _
                              "  X    "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
                Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowAllRecord, False)
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
        grid.VisibleColumns(7).CellStyle.BackColor = Color.White
        grid.VisibleColumns(11).CellStyle.BackColor = Color.White
        grid.VisibleColumns(12).CellStyle.BackColor = Color.White
        grid.VisibleColumns(13).CellStyle.BackColor = Color.White
    End Sub

    Private Sub uf_Approve()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
                ls_sql = " Update PO_Master_Export set AffiliateApproveDate = getdate(), AffiliateApproveUser = '" & Session("UserID") & "'" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & rdEmergency.Text & "' and SupplierID = '" & Session("LoadSupplier") & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub uf_UnApprove()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
                ls_sql = " Update PO_Master_Export set AffiliateApproveDate = NULL, AffiliateApproveUser = NULL" & vbCrLf & _
                            " WHERE AffiliateID = '" & Session("AffiliateID") & "' and PONo = '" & rdEmergency.Text & "' and SupplierID = '" & Session("LoadSupplier") & "'" & vbCrLf

                Dim SqlComm As New SqlCommand(ls_sql, sqlConn, sqlTran)
                x = SqlComm.ExecuteNonQuery()

                SqlComm.Dispose()
                sqlTran.Commit()
            End Using
            sqlConn.Close()
        End Using
    End Sub

    Private Sub up_fillcombo()
        Dim ls_sql As String

        ls_sql = ""

        'Affiliate ID
        ls_sql = "SELECT [Affiliate Code] = RTRIM(AffiliateID) ,[Affiliate Name] = RTRIM(AffiliateName), [Consignee Code] = Rtrim(isnull(AffiliateCode,'')) FROM MS_Affiliate  where isnull(overseascls, '0') = '1'" & vbCrLf
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
                '.TextField = "Consignee Code"
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

    Private Sub uf_RecoverySplit()
        Dim ls_sql As String
        Dim x As Integer

        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            Using sqlTran As SqlTransaction = sqlConn.BeginTransaction("POExportEntryMonthly")
                ls_sql = "INSERT INTO PO_Master_ExportRecoverySplit( " & vbCrLf & _
                    "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1) " & vbCrLf & _
                    "SELECT PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1 " & vbCrLf & _
                    "FROM PO_Master_Export " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "AND NOT EXISTS( " & vbCrLf & _
                    "   SELECT * FROM PO_Master_ExportRecoverySplit a " & vbCrLf & _
                    "   WHERE a.PONo = PO_Master_Export.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Master_Export.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Master_Export.SupplierID " & vbCrLf & _
                    "   AND a.OrderNo1 = PO_Master_Export.OrderNo1 " & vbCrLf & _
                    ")"

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
                    "FROM PO_Master_Export " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                    "AND NOT EXISTS( " & vbCrLf & _
                    "   SELECT * FROM PO_Master_Export a " & vbCrLf & _
                    "   WHERE a.PONo = PO_Master_Export.PONo " & vbCrLf & _
                    "   AND a.AffiliateID = PO_Master_Export.AffiliateID " & vbCrLf & _
                    "   AND a.SupplierID = PO_Master_Export.SupplierID " & vbCrLf & _
                    "   AND a.OrderNo1 = PO_Master_Export.SplitReffPONo " & vbCrLf & _
                    ") " & vbCrLf

                ls_sql = ls_sql & "UPDATE PO_Detail_Export SET Week1 = Week1 + ISNULL(( " & vbCrLf & _
                    "   SELECT b.Week1 FROM PO_Master_Export a " & vbCrLf & _
                    "   INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                    "   SELECT b.TotalPOQty FROM PO_Master_Export a " & vbCrLf & _
                    "   INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                    "   SELECT * FROM PO_Master_Export a " & vbCrLf & _
                    "   INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                    "FROM PO_Master_Export a " & vbCrLf & _
                    "INNER JOIN PO_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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

                Select Case Session("POStatus")
                    Case "3", "4", "5", "6"
                        ls_sql = ls_sql & "INSERT INTO PO_MasterUpload_Export( " & vbCrLf & _
                            "PONo, AffiliateID, SupplierID, ForwarderID, OrderNo1, ETDVendor1, Remarks, " & vbCrLf & _
                            "EntryDate, EntryUser, UpdateDate, UpdateUser) " & vbCrLf & _
                            "SELECT b.PONo, b.AffiliateID, b.SupplierID, b.ForwarderID, a.SplitReffPONo OrderNo1, b.ETDVendor1, b.Remarks, " & vbCrLf & _
                            "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, GETDATE() UpdateDate, '" & Session("UserID").ToString & "' UpdateUser " & vbCrLf & _
                            "FROM PO_Master_Export a " & vbCrLf & _
                            "INNER JOIN PO_MasterUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "   SELECT b.Week1 FROM PO_Master_Export a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "Week1Old = Week1Old + ISNULL(( " & vbCrLf & _
                            "   SELECT b.Week1Old FROM PO_Master_Export a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "   SELECT b.TotalPOQty FROM PO_Master_Export a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "TotalPOQtyOld = TotalPOQtyOld + ISNULL(( " & vbCrLf & _
                            "   SELECT b.TotalPOQtyOld FROM PO_Master_Export a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "   SELECT * FROM PO_Master_Export a " & vbCrLf & _
                            "   INNER JOIN PO_DetailUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "FROM PO_Master_Export a " & vbCrLf & _
                            "INNER JOIN PO_DetailUpload_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo1 " & vbCrLf & _
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
                            "   FROM PO_Master_Export a " & vbCrLf & _
                            "   WHERE a.SupplierID = PrintLabelExport.SupplierID " & vbCrLf & _
                            "   AND a.AffiliateID = PrintLabelExport.AffiliateID " & vbCrLf & _
                            "   AND a.PONo = PrintLabelExport.PONo " & vbCrLf & _
                            "   AND a.OrderNo1 = PrintLabelExport.OrderNo " & vbCrLf & _
                            "), '') " & vbCrLf & _
                            "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                        Select Case Session("POStatus")
                            Case "5", "6"
                                ls_sql = ls_sql & "INSERT INTO DOSupplier_Master_Export( " & vbCrLf & _
                                    "SuratJalanNo, SupplierID, AffiliateID, PONo, OrderNo, DeliveryDate, PIC, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                    "EntryDate, EntryUser, UpdateDate, UpdateUser, ExcelCls, MovingList) " & vbCrLf & _
                                    "SELECT b.SuratJalanNo, b.SupplierID, b.AffiliateID, b.PONo, a.SplitReffPONo OrderNo, b.DeliveryDate, b.PIC, b.JenisArmada, b.DriverName, b.DriverContact, b.NoPol, b.TotalBox, " & vbCrLf & _
                                    "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, GETDATE() UpdateDate, '" & Session("UserID").ToString & "' UpdateUser, b.ExcelCls, b.MovingList " & vbCrLf & _
                                    "FROM PO_Master_Export a " & vbCrLf & _
                                    "INNER JOIN DOSupplier_Master_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
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
                                    "   SELECT b.DOQty FROM PO_Master_Export a " & vbCrLf & _
                                    "   INNER JOIN DOSupplier_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
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
                                    "   SELECT * FROM PO_Master_Export a " & vbCrLf & _
                                    "   INNER JOIN DOSupplier_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
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
                                    "FROM PO_Master_Export a " & vbCrLf & _
                                    "INNER JOIN DOSupplier_Detail_Export b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
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
                                    "   FROM PO_Master_Export a " & vbCrLf & _
                                    "   WHERE a.SupplierID = DOSupplier_DetailBox_Export.SupplierID " & vbCrLf & _
                                    "   AND a.AffiliateID = DOSupplier_DetailBox_Export.AffiliateID " & vbCrLf & _
                                    "   AND a.PONo = DOSupplier_DetailBox_Export.PONo " & vbCrLf & _
                                    "   AND a.OrderNo1 = DOSupplier_DetailBox_Export.OrderNo " & vbCrLf & _
                                    "), '') " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                                ls_sql = ls_sql & "DELETE DOSupplier_Detail_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                                ls_sql = ls_sql & "DELETE DOSupplier_Master_Export " & vbCrLf & _
                                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                    "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                                If Session("POStatus") = "6" Then
                                    ls_sql = ls_sql & "INSERT INTO ReceiveForwarder_Master( " & vbCrLf & _
                                        "SuratJalanNo, AffiliateID, SupplierID, PONo, ForwarderID, OrderNo, " & vbCrLf & _
                                        "ExcelCls, ReceiveDate, ReceiveBy, JenisArmada, DriverName, DriverContact, NoPol, TotalBox, " & vbCrLf & _
                                        "EntryDate, EntryUser, UpdateDate, UpdateUser, MovingList, SplitReffPONo) " & vbCrLf & _
                                        "SELECT b.SuratJalanNo, b.AffiliateID, b.SupplierID, b.PONo, b.ForwarderID, a.SplitReffPONo OrderNo, " & vbCrLf & _
                                        "b.ExcelCls, b.ReceiveDate, b.ReceiveBy, b.JenisArmada, b.DriverName, b.DriverContact, b.NoPol, b.TotalBox, " & vbCrLf & _
                                        "GETDATE() EntryDate, '" & Session("UserID").ToString & "' EntryUser, GETDATE() UpdateDate, '" & Session("UserID").ToString & "' UpdateUser, b.MovingList, b.SplitReffPONo " & vbCrLf & _
                                        "FROM PO_Master_Export a " & vbCrLf & _
                                        "INNER JOIN ReceiveForwarder_Master b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                        "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                        "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                        "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                        "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                        "AND NOT EXISTS( " & vbCrLf & _
                                        "   SELECT * FROM ReceiveForwarder_Master " & vbCrLf & _
                                        "   WHERE ReceiveForwarder_Master.PONo = a.PONo " & vbCrLf & _
                                        "   AND ReceiveForwarder_Master.AffiliateID = a.AffiliateID " & vbCrLf & _
                                        "   AND ReceiveForwarder_Master.SupplierID = a.SupplierID " & vbCrLf & _
                                        "   AND ReceiveForwarder_Master.OrderNo = a.SplitReffPONo " & vbCrLf & _
                                        ") " & vbCrLf

                                    ls_sql = ls_sql & "UPDATE ReceiveForwarder_Detail SET GoodRecQty = GoodRecQty + ISNULL(( " & vbCrLf & _
                                        "   SELECT b.GoodRecQty FROM PO_Master_Export a " & vbCrLf & _
                                        "   INNER JOIN ReceiveForwarder_Detail b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                        "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                        "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                        "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                        "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                        "   AND a.PONo = ReceiveForwarder_Detail.PONo " & vbCrLf & _
                                        "   AND a.AffiliateID = ReceiveForwarder_Detail.AffiliateID " & vbCrLf & _
                                        "   AND a.SupplierID = ReceiveForwarder_Detail.SupplierID " & vbCrLf & _
                                        "   AND a.SplitReffPONo = ReceiveForwarder_Detail.OrderNo " & vbCrLf & _
                                        "   AND b.PartNo = ReceiveForwarder_Detail.PartNo " & vbCrLf & _
                                        "), 0) " & vbCrLf & _
                                        "WHERE EXISTS( " & vbCrLf & _
                                        "   SELECT * FROM PO_Master_Export a " & vbCrLf & _
                                        "   INNER JOIN ReceiveForwarder_Detail b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                        "   WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                        "   AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                        "   AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                        "   AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                        "   AND a.PONo = ReceiveForwarder_Detail.PONo " & vbCrLf & _
                                        "   AND a.AffiliateID = ReceiveForwarder_Detail.AffiliateID " & vbCrLf & _
                                        "   AND a.SupplierID = ReceiveForwarder_Detail.SupplierID " & vbCrLf & _
                                        "   AND a.SplitReffPONo = ReceiveForwarder_Detail.OrderNo " & vbCrLf & _
                                        "   AND b.PartNo = ReceiveForwarder_Detail.PartNo " & vbCrLf & _
                                        ") " & vbCrLf

                                    ls_sql = ls_sql & "INSERT INTO ReceiveForwarder_Detail( " & vbCrLf & _
                                        "SuratJalanNo, SupplierID, AffiliateID, PONo, PartNo, OrderNo, GoodRecQty, DefectRecQty) " & vbCrLf & _
                                        "SELECT b.SuratJalanNo, b.SupplierID, b.AffiliateID, b.PONo, b.PartNo, a.SplitReffPONo OrderNo, b.GoodRecQty, 0 DefectRecQty " & vbCrLf & _
                                        "FROM PO_Master_Export a " & vbCrLf & _
                                        "INNER JOIN ReceiveForwarder_Detail b ON a.PONo = b.PONo AND a.AffiliateID = b.AffiliateID AND a.SupplierID = b.SupplierID AND a.OrderNo1 = b.OrderNo " & vbCrLf & _
                                        "WHERE a.PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                        "AND a.AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                        "AND a.SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                        "AND a.OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf & _
                                        "AND NOT EXISTS( " & vbCrLf & _
                                        "   SELECT * FROM ReceiveForwarder_Detail " & vbCrLf & _
                                        "   WHERE ReceiveForwarder_Detail.PONo = a.PONo " & vbCrLf & _
                                        "   AND ReceiveForwarder_Detail.AffiliateID = a.AffiliateID " & vbCrLf & _
                                        "   AND ReceiveForwarder_Detail.SupplierID = a.SupplierID " & vbCrLf & _
                                        "   AND ReceiveForwarder_Detail.OrderNo = a.SplitReffPONo " & vbCrLf & _
                                        "   AND ReceiveForwarder_Detail.PartNo = b.PartNo " & vbCrLf & _
                                        ") " & vbCrLf

                                    ls_sql = ls_sql & "DELETE ReceiveForwarder_Detail " & vbCrLf & _
                                        "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                        "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                        "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                        "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                                    ls_sql = ls_sql & "DELETE ReceiveForwarder_Master " & vbCrLf & _
                                        "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                                        "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                                        "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                                        "AND OrderNo = '" & txtOrderNo.Text.Trim & "' " & vbCrLf
                                End If
                        End Select

                        ls_sql = ls_sql & "DELETE PO_DetailUpload_Export " & vbCrLf & _
                            "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                        ls_sql = ls_sql & "DELETE PO_MasterUpload_Export " & vbCrLf & _
                            "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                            "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                            "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                            "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf
                End Select

                ls_sql = ls_sql & "DELETE PO_Detail_Export " & vbCrLf & _
                    "WHERE PONo = '" & txtpono.Text.Trim & "' " & vbCrLf & _
                    "AND AffiliateID = '" & cboAffiliate.Text.Trim & "' " & vbCrLf & _
                    "AND SupplierID = '" & Session("LoadSupplier") & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & txtOrderNo.Text.Trim & "' " & vbCrLf

                ls_sql = ls_sql & "DELETE PO_Master_Export " & vbCrLf & _
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

    Private Sub uf_RecoverySplitEmailSupplier()
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim dsEmail As New DataSet
            dsEmail = GetEmailToSupplier(Session("AffiliateID"), "PASI", Session("LoadSupplier"))

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
            mailMessage.Subject = "[TRIAL] Notification For PO Split Recovery, Order No : " & txtpono.Text.Trim & " Split (" & txtOrderNo.Text.Trim & ")"

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
            ls_Body = clsNotification.GetNotification("26", "", txtpono.Text.Trim & " Split (" & txtOrderNo.Text.Trim & ")")

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

    Private Sub uf_RecoverySplitEmailForwarder(ByVal pPONo As String, ByVal pAffiliate As String, ByVal pSupplier As String, ByVal pOrderNo As String)
        Try
            Dim receiptEmail As String = ""
            Dim receiptCCEmail As String = ""
            Dim fromEmail As String = ""

            Dim ls_SQL As String = ""
            Dim ls_Forwarder As String = ""
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()

                ls_SQL = "SELECT ForwarderID FROM PO_Master_Export " & vbCrLf & _
                    "WHERE PONo = '" & pPONo & "' " & vbCrLf & _
                    "AND AffiliateID = '" & pAffiliate & "' " & vbCrLf & _
                    "AND SupplierID = '" & pSupplier & "' " & vbCrLf & _
                    "AND OrderNo1 = '" & pOrderNo & "' "

                Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    ls_Forwarder = ds.Tables(0).Rows(0)("ForwarderID")
                End If
            End Using

            Dim dsEmail As New DataSet
            dsEmail = GetEmailToForwarder(cboDelLoc.Text.Trim, Session("AffiliateID"))

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
            mailMessage.Subject = "[TRIAL] Notification For PO Split Recovery, Order No : " & txtpono.Text.Trim & " Split (" & txtOrderNo.Text.Trim & ")"

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
            ls_Body = clsNotification.GetNotification("26", "", txtpono.Text.Trim & " Split (" & txtOrderNo.Text.Trim & ")")

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
#End Region

#Region "Download"

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                             ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim NewFileName As String = Server.MapPath("~\PurchaseOrderExport\" & FileName)
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
                'Header
                .Cells(1, 3).Value = "M"
                .Cells(2, 3).Value = Format(dtPeriodFrom.Value, "yyyy-MM")
                .Cells(3, 3).Value = Trim(cboAffiliate.Text)
                .Cells(4, 3).Value = Trim(txtconsignee.Text)
                .Cells(5, 3).Value = IIf(rdrCom1.Value = "1", "YES", "NO")
                .Cells(6, 3).Value = IIf(rdrShipBy2.Value = "1", "B", "A")
                .Cells(9, 4).Value = Trim(txtpono.Text)
                .Cells(10, 4).Value = Format(dtETDVendor.Value, "yyyy-MM-dd")
                .Cells(11, 4).Value = Format(dtETDPort.Value, "yyyy-MM-dd")
                .Cells(12, 4).Value = Format(dtETAPort.Value, "yyyy-MM-dd")
                .Cells(13, 4).Value = Format(dtETAFactory.Value, "yyyy-MM-dd")
                icol = 19
                For irow = 0 To grid.VisibleRowCount - 1
                    .Cells(icol, 1).Value = irow + 1
                    .Cells(icol, 2).Value = Trim(grid.GetRowValues(irow, "PartNo"))
                    .Cells(icol, 3).Value = Trim(grid.GetRowValues(irow, "UOM"))
                    .Cells(icol, 4).Value = Trim(grid.GetRowValues(irow, "MOQ"))
                    .Cells(icol, 5).Value = Trim(grid.GetRowValues(irow, "Week1"))
                    .Cells(icol, 6).Value = 0
                    .Cells(icol, 7).Value = 0
                    .Cells(icol, 8).Value = 0
                    .Cells(icol, 9).Value = 0
                    .Cells(icol, 10).Value = Trim(grid.GetRowValues(irow, "PreviousForecast"))
                    .Cells(icol, 11).Value = Trim(grid.GetRowValues(irow, "Forecast1"))
                    .Cells(icol, 12).Value = Trim(grid.GetRowValues(irow, "Forecast2"))
                    .Cells(icol, 13).Value = Trim(grid.GetRowValues(irow, "Forecast3"))
                    .Cells(icol, 14).Value = Trim(grid.GetRowValues(irow, "ErrorStatus"))
                    icol = icol + 1
                Next

                Dim rgAll As ExcelRange = .Cells("A19:N" & icol - 1)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback(fi.Name)

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

#End Region

    Private Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Session("buttonSubMenu") = "Direct"
        Response.Redirect("~/PurchaseOrderExport/POUploadExport.aspx")
    End Sub

    Private Sub ASPxCallback1_Callback(ByVal source As Object, ByVal e As DevExpress.Web.ASPxCallback.CallbackEventArgs) Handles ASPxCallback1.Callback
        Dim pAction As String = Split(e.Parameter, "|")(0)
        Try
            Select Case pAction
                Case "loaditem"
                    ASPxCallback1.JSProperties("cpLossFocus") = "ON"
                    Call up_ItemLoad()
                    lblInfo.Text = ""
                Case "recoveryposplit"

                    Call uf_RecoverySplit()

                    Select Case Session("GOTOStatus")
                        Case "2", "3", "4"
                            Call uf_RecoverySplitEmailSupplier()
                    End Select

                    Select Case Session("GOTOStatus")
                        Case "4", "5", "6"
                            Call uf_RecoverySplitEmailForwarder(txtpono.Text.Trim, cboAffiliate.Text.Trim, Session("LoadSupplier"), txtOrderNo.Text.Trim)

                            If txtOrderNo.Text.Trim <> Session("SplitReffPONo") Then
                                Call uf_RecoverySplitEmailForwarder(txtpono.Text.Trim, cboAffiliate.Text.Trim, Session("LoadSupplier"), Session("SplitReffPONo"))
                            End If
                    End Select
            End Select
        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        ASPxCallback1.JSProperties("cpMessage") = Session("ErrorData")
        If Session("ErrorData") <> "" Then
            ASPxCallback1.JSProperties("cpJumlahError") = Session("JumlahError")
            If Session("UbahDataGrid") <> "1" Then
                Session.Remove("ErrorData")
            End If
        ElseIf HF.Get("hfTest") = 1 Then
            Call up_GridLoad()
            Call up_CheckData()

            ASPxCallback1.JSProperties("cpMessage") = Session("ErrorData")
            Session.Remove("ErrorData")
        ElseIf HF.Get("hfTest") = 2 Then
            Call up_SaveData()
            Call up_SendDataToSupplier()
            ASPxCallback1.JSProperties("cpMessage") = Session("ErrorData")
            Session.Remove("ErrorData")
        ElseIf Session("DataTersimpan") = "1" Then
            ASPxCallback1.JSProperties("cpMessage") = Session("ErrorData")
            Session.Remove("DataTersimpan")
            Session.Remove("ErrorData")
        End If
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

    Private Sub btnSplit_Click(sender As Object, e As System.EventArgs) Handles btnSplit.Click
        pSupplierCode = Session("LoadSupplier")

        Dim sRedirect As String
        sRedirect = "~/PurchaseOrderExport/POExportEntrySplit.aspx?prm=" & _
            Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "|" & _
            Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "|"

        If rdrCom1.Checked = True Then
            sRedirect = sRedirect + "1|"
        ElseIf rdrCom2.Checked = True Then
            sRedirect = sRedirect + "0|"
        End If

        If rdEmergency.Checked = True Then
            sRedirect = sRedirect + "E|"
        ElseIf rdMonthly.Checked = True Then
            sRedirect = sRedirect + "M|"
        End If

        If rdrShipBy2.Checked = True Then
            sRedirect = sRedirect + "B|"
        ElseIf rdrShipBy3.Checked = True Then
            sRedirect = sRedirect + "A|"
        End If

        sRedirect = sRedirect + _
            cboAffiliate.Text & "|" & _
            txtAffiliate.Text & "|" & _
            pSupplierCode & "|" & _
            pSupplierName & "|" & _
            txtOrderNo.Text & "|" & _
            Format(dtETDVendor.Value, "yyyy-MM-dd") & "|" & _
            Format(dtETDPort.Value, "yyyy-MM-dd") & "|" & _
            Format(dtETAPort.Value, "yyyy-MM-dd") & "|" & _
            Format(dtETAFactory.Value, "yyyy-MM-dd") & "|" & _
            cboDelLoc.Text & "|" & _
            txtDelLoc.Text & "|" & _
            txtpono.Text & "|" & _
            txtconsignee.Text
        Response.Redirect(sRedirect)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        pSupplierCode = Session("LoadSupplier")

        Dim sRedirect As String
        sRedirect = "~/PurchaseOrderExport/POExportEntryCancel.aspx?prm=" & _
            Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "|" & _
            Format(dtPeriodFrom.Value, "yyyy-MM-dd") & "|"

        If rdrCom1.Checked = True Then
            sRedirect = sRedirect + "1|"
        ElseIf rdrCom2.Checked = True Then
            sRedirect = sRedirect + "0|"
        End If

        If rdEmergency.Checked = True Then
            sRedirect = sRedirect + "E|"
        ElseIf rdMonthly.Checked = True Then
            sRedirect = sRedirect + "M|"
        End If

        If rdrShipBy2.Checked = True Then
            sRedirect = sRedirect + "B|"
        ElseIf rdrShipBy3.Checked = True Then
            sRedirect = sRedirect + "A|"
        End If

        sRedirect = sRedirect + _
            cboAffiliate.Text & "|" & _
            txtAffiliate.Text & "|" & _
            pSupplierCode & "|" & _
            pSupplierName & "|" & _
            txtOrderNo.Text & "|" & _
            Format(dtETDVendor.Value, "yyyy-MM-dd") & "|" & _
            Format(dtETDPort.Value, "yyyy-MM-dd") & "|" & _
            Format(dtETAPort.Value, "yyyy-MM-dd") & "|" & _
            Format(dtETAFactory.Value, "yyyy-MM-dd") & "|" & _
            cboDelLoc.Text & "|" & _
            txtDelLoc.Text & "|" & _
            txtpono.Text & "|" & _
            txtconsignee.Text
        Response.Redirect(sRedirect)
    End Sub
End Class