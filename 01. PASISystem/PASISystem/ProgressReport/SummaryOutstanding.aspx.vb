Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO


Public Class SummaryOutstanding
    Inherits System.Web.UI.Page

#Region "Declaration"
    Dim clsGlobal As New clsGlobal
    Dim clsMsg As New clsMessage
    Dim clsAppearance As New clsAppearance
    Dim ls_SQL As String = ""

    Dim ls_AllowUpdate As Boolean = False
    Dim ls_AllowDelete As Boolean = False
    Dim menuID As String = "G01"

    Const colNo As Byte = 1
    Const colPeriod As Byte = 2
    Const colPONo As Byte = 3
    Const colAffiliateCode As Byte = 4
    Const colSupplierCode As Byte = 5
    Const colPOKanban As Byte = 6
    Const colKanbanNo As Byte = 7
    Const colSupplierPlanDelDate As Byte = 8
    Const colPartNo As Byte = 9
    Const colPartName As Byte = 10
    Const colQtyPO As Byte = 11
    Const colSupplierDelDate As Byte = 12
    Const colSupplierSJNo As Byte = 13
    Const colSupplierDeliveryQty As Byte = 14
    Const colPASIRecDate As Byte = 15
    Const colPASIReceivingQty As Byte = 16
    Const colInvoiceNoFromSupplier As Byte = 17
    Const colInvoiceDateFromSupplier As Byte = 18
    Const colInvoiceFromSupplierCurr As Byte = 19
    Const colInvoiceFromSupplierAmount As Byte = 20
    Const colPASIDelDate As Byte = 21
    Const colPASISJNo As Byte = 22
    Const colPASIDeliveryQty As Byte = 23
    Const colAffiliateRecDate As Byte = 24
    Const colAffiliateReceivingQty As Byte = 25
    Const colInvoiceNoToAffiliate As Byte = 26
    Const colInvoiceDateToAffiliate As Byte = 27
    Const colInvoiceToAffiliateCurr As Byte = 28
    Const colInvoiceToAffiliateAmount As Byte = 29
    Const colCount As Byte = 30

    Dim FilePath As String = ""
    Dim FileName As String = ""
    Dim FileExt As String = ""
    Dim Ext As String = ""
    Dim FolderPath As String = ""
#End Region

#Region "Procedures"
    Private Sub up_Initialize()
        Dim script As String = _
            "if (cboAffiliateCode.GetItemCount() > 1) { " & vbCrLf & _
            "   txtAffiliateName.SetText('==ALL=='); " & vbCrLf & _
            "   cboAffiliateCode.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriodFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtPOPeriodTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierPlanDelDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierPlanDelDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierDelDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtSupplierDelDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtPASIRecDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtPASIRecDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtPASIDelDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtPASIDelDateTo.SetDate(PeriodTo); " & vbCrLf & _
            "dtAffiliateRecDateFrom.SetDate(PeriodTo); " & vbCrLf & _
            "dtAffiliateRecDateTo.SetDate(PeriodTo); " & vbCrLf & _
            " " & vbCrLf & _
            "txtPONo.SetText(''); " & vbCrLf & _
            "txtSupplierSJNo.SetText(''); " & vbCrLf & _
            "txtPASISJNo.SetText(''); " & vbCrLf & _
            "txtSupplierInvNo.SetText(''); " & vbCrLf & _
            "txtPASIInvNo.SetText(''); " & vbCrLf & _
            " " & vbCrLf & _
            "chkSupplierPlanDelDate.SetValue(false); " & vbCrLf & _
            "chkSupplierDelDate.SetValue(false); " & vbCrLf & _
            "chkPASIRecDate.SetValue(false); " & vbCrLf & _
            "chkPASIDelDate.SetValue(false); " & vbCrLf & _
            "chkAffiliateRecDate.SetValue(false); " & vbCrLf & _
            " " & vbCrLf & _
            "rdrPRAll.SetValue(true); " & vbCrLf & _
            "rdrPRComplete.SetValue(false); " & vbCrLf & _
            "rdrPRRemaining.SetValue(false); " & vbCrLf & _
            "rdrPRDiff.SetValue(false); " & vbCrLf & _
            "rdrPDAll.SetValue(true); " & vbCrLf & _
            "rdrPDComplete.SetValue(false); " & vbCrLf & _
            "rdrPDRemaining.SetValue(false); " & vbCrLf & _            
            "rdrARAll.SetValue(true); " & vbCrLf & _
            "rdrARComplete.SetValue(false); " & vbCrLf & _
            "rdrARRemaining.SetValue(false); " & vbCrLf & _
            "rdrARDiff.SetValue(false); " & vbCrLf & _
            "rdrPIAll.SetValue(true); " & vbCrLf & _
            "rdrPIComplete.SetValue(false); " & vbCrLf & _
            "rdrPIRemaining.SetValue(false); " & vbCrLf & _
            "rdrSIAll.SetValue(true); " & vbCrLf & _
            "rdrSIComplete.SetValue(false); " & vbCrLf & _
            "rdrSIRemaining.SetValue(false); " & vbCrLf & _
            " " & vbCrLf & _
            " " & vbCrLf & _
            "if (cboPart.GetItemCount() > 1) { " & vbCrLf & _
            "   txtPartName.SetText('==ALL=='); " & vbCrLf & _
            "   cboPart.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(chkSupplierPlanDelDate, chkSupplierPlanDelDate.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim ls_filter As String = ""
            ls_SQL = ""

            Dim ls_End As String = ""
            ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

            'AFFILIATE CODE
            If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            End If

            'AFFILIATE PO PERIOD
            If chkPOPeriod.Checked = True Then
                ls_filter = ls_filter + _
                              "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyy-MM-01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End) & "' " & vbCrLf
            End If

            'SUPPLIER PLAN DELIVERY DATE
            If chkSupplierPlanDelDate.Checked = True Then
                ls_filter = ls_filter + _
                              "                      AND CONVERT(date,KM.KanbanDate) BETWEEN '" & Format(dtSupplierPlanDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierPlanDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'SUPPLIER DELIVERY DATE
            If chkSupplierDelDate.Checked = True Then
                ls_filter = ls_filter + _
                              "                      AND CONVERT(date,SDM.DeliveryDate) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'PASI RECEIVE DATE
            If chkPASIRecDate.Checked = True Then
                ls_filter = ls_filter + _
                              "                      AND CONVERT(date,PRM.ReceiveDate) BETWEEN '" & Format(dtPASIRecDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtPASIRecDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'PASI DELIVERY DATE
            If chkPASIDelDate.Checked = True Then
                ls_filter = ls_filter + _
                              "                      AND CONVERT(date,PDM.DeliveryDate) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'AFFILIATE RECEIVE DATE
            If chkAffiliateRecDate.Checked = True Then
                ls_filter = ls_filter + _
                              "                      AND CONVERT(date,RAM.ReceiveDate) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
            End If

            'PART CODE
            If Trim(cboPart.Text) <> "==ALL==" And Trim(cboPart.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND POD.PartNo = '" & Trim(cboPart.Text) & "' " & vbCrLf
            End If

            'PONO
            If Trim(txtPONo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(POM.PONo,'') LIKE '%" & Trim(txtPONo.Text) & "%' " & vbCrLf
            End If

            'SUPPLIER SJ NO
            If Trim(txtSupplierSJNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(SDD.SuratJalanNo,'') LIKE '%" & Trim(txtSupplierSJNo.Text) & "%'" & vbCrLf
            End If

            'PASI SJ NO
            If Trim(txtPASISJNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(PDD.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
            End If

            'SUPPLIER INV NO
            If Trim(txtSupplierInvNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(ISD.InvoiceNo,'') LIKE '%" & Trim(txtSupplierInvNo.Text) & "%'" & vbCrLf
            End If

            'PASI INV NO
            If Trim(txtPASIInvNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(IPD.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
            End If
            'PO PROGRESS

            'PASI RECEIVE
            If rdrPRComplete.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0) " & vbCrLf
            ElseIf rdrPRRemaining.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(SDD.DOQty,0) > ISNULL(PRD.GoodRecQty,0)  " & vbCrLf
            ElseIf rdrPRDiff.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND (ISNULL(SDD.DOQty,0) < ISNULL(PRD.GoodRecQty,0) OR ISNULL(PRD.DefectRecQty,0) > 0) " & vbCrLf
            End If

            'PASI DELIVERY
            If rdrPDComplete.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) = ISNULL(PDD.DOQty,0) " & vbCrLf & _
                              "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
            ElseIf rdrPDComplete.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) > (PDD.DOQty) " & vbCrLf & _
                              "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
            End If

            'AFFILIATE RECEIVE
            If rdrARComplete.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                              "                            END) = RAD.RecQty AND " & vbCrLf & _
                              "                           (CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.DefectQty,0) " & vbCrLf & _
                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.DefectRecQty,0) " & vbCrLf & _
                              "                            END) = 0 " & vbCrLf & _
                              "                          ) " & vbCrLf
            ElseIf rdrARRemaining.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                              "                            END) > RAD.RecQty " & vbCrLf & _
                              "                          ) " & vbCrLf
            ElseIf rdrARDiff.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                              "                            END) < RAD.RecQty OR " & vbCrLf & _
                              "                           ISNULL(RAD.DefectQty,0) > 0  " & vbCrLf & _
                              "                          )" & vbCrLf
            End If

            'SUPPLIER INVOICE
            If rdrSIComplete.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
                              "                            END) = ISD.InvQty " & vbCrLf & _
                              "                          ) " & vbCrLf
            ElseIf rdrSIRemaining.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
                              "                            END) > ISD.InvQty " & vbCrLf & _
                              "                          ) " & vbCrLf
            End If

            'PASI INVOICE
            If rdrPIComplete.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(RAD.RecQty,0) = ISNULL(IPD.DOQty,0) " & vbCrLf
            ElseIf rdrPIRemaining.Value = True Then
                ls_filter = ls_filter + _
                              "                      AND ISNULL(RAD.RecQty,0) > ISNULL(IPD.DOQty,0) " & vbCrLf
            End If

            ls_SQL = " SELECT DISTINCT ColNo = '', * FROM " & vbCrLf & _
                  " ( " & vbCrLf & _
                  " 	SELECT  " & vbCrLf & _
                  " 		POM.Period " & vbCrLf & _
                  " 		,POM.PONo " & vbCrLf & _
                  " 		,POM.AffiliateID " & vbCrLf & _
                  " 		,POM.SupplierID " & vbCrLf & _
                  " 		,POKanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '1' THEN 'YES' ELSE 'NO' END " & vbCrLf & _
                  " 		,POM.EntryDate " & vbCrLf & _
                  " 		,POM.PASISendAffiliateDate " & vbCrLf & _
                  " 		,POD.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " 		,MP.PartName " & vbCrLf & _
                              " 		,QtyPO = ISNULL(POD.POQty,0) " & vbCrLf & _
                              " 		,KD.KanbanNo " & vbCrLf & _
                              " 		,KD.KanbanQty " & vbCrLf & _
                              " 		,ETDSupp = ABC.ETDSupplier " & vbCrLf & _
                              " 		,ETAAff = KM.KanbanDate " & vbCrLf & _
                              " 		,SupplierDeliveryDate = SDM.DeliveryDate " & vbCrLf & _
                              " 		,SupplierSuratJalanNo = SDM.SuratJalanNo " & vbCrLf & _
                              " 		,SupplierDeliveryQty = SDD.DOQty " & vbCrLf & _
                              " 		,RemainingQtyPOPASI = ISNULL(KD.KanbanQty,0) - " & vbCrLf & _
                              " 		                      ISNULL( " & vbCrLf & _
                              " 		                        (select SUM(DOQty) from DOSupplier_Detail ABC " & vbCrLf & _
                              " 		                         INNER JOIN DOSupplier_Master EFG ON ABC.SuratJalanNo = EFG.SuratJalanNo AND ABC.SupplierID = EFG.SupplierID AND ABC.AffiliateID = EFG.AffiliateID " & vbCrLf & _
                              " 		                         WHERE ABC.SupplierID = SDD.SupplierID and ABC.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                              " 		                         and ABC.KanbanNo = SDD.KanbanNo and ABC.PartNo = SDD.PartNo and ABC.PONo = SDD.PONo and EFG.SuratJalanNo <= SDM.SuratJalanNo),0) " & vbCrLf & _
                              " 		,PASIReceiveDate = PRM.ReceiveDate " & vbCrLf

            ls_SQL = ls_SQL + " 		,PASIReceivingQty = PRD.GoodRecQty " & vbCrLf & _
                              " 		,InvoiceNoFromSupplier = ISM.InvoiceNo " & vbCrLf & _
                              " 		,InvoiceDateFromSupplier = ISM.InvoiceDate " & vbCrLf & _
                              " 		,InvoiceFromSupplierCurr = MCS.Description " & vbCrLf & _
                              " 		,InvoiceFromSupplierAmount = ISNULL(ISD.InvAmount,0) " & vbCrLf & _
                              " 		,PASIDeliveryDate = PDM.DeliveryDate " & vbCrLf & _
                              " 		,PASISuratJalanNo = IPM.SuratJalanNo " & vbCrLf & _
                              " 		,PASIDeliveryQty = IPD.DOQty " & vbCrLf & _
                              " 		,AffiliateReceiveDate = RAM.ReceiveDate " & vbCrLf & _
                              " 		,AffiliateReceivingQty = RAD.RecQty " & vbCrLf & _
                              " 		,InvoiceNoToAffiliate = IPM.InvoiceNo " & vbCrLf

            ls_SQL = ls_SQL + " 		,InvoiceDateToAffiliate = IPM.DeliveryDate " & vbCrLf & _
                              " 		,InvoiceToAffiliateCurr = 'IDR' " & vbCrLf & _
                              " 		,InvoiceToAffiliateAmount = ISNULL(IPD.DOQty,0) * ISNULL(PDD.Price,0) " & vbCrLf & _
                              " 	FROM PO_Master POM " & vbCrLf & _
                              " 	LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 							AND POM.PoNo = POD.PONo " & vbCrLf & _
                              " 							AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
                              " 								AND KD.PoNo = POD.PONo " & vbCrLf & _
                              " 								AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
                              " 								AND KD.PartNo = POD.PartNo " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                              " 								AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                              " 								AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                              " 								AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                              " 	LEFT JOIN DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                              " 									AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
                              " 									AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
                              " 									AND KD.PartNo = SDD.PartNo " & vbCrLf & _
                              " 									AND KD.PoNo = SDD.PoNo " & vbCrLf & _
                              " 	LEFT JOIN DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                              " 									AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf

            ls_SQL = ls_SQL + " 									AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                              " 									AND SDD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                              " 									AND SDD.SupplierID = PRD.SupplierID " & vbCrLf & _
                              " 									AND SDD.PartNo = PRD.PartNo " & vbCrLf & _
                              " 									AND SDD.PONo = PRD.PONo								 " & vbCrLf & _
                              " 									AND SDD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                              " 	LEFT JOIN ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                              " 									AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                              " 									AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
                              " 	LEFT JOIN InvoiceSupplier_Detail ISD ON ISD.AffiliateID = PRD.AffiliateID " & vbCrLf

            ls_SQL = ls_SQL + " 										AND ISD.SupplierID = PRD.SupplierID " & vbCrLf & _
                              " 										AND ISD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                              " 										AND ISD.PONo = PRD.PONo " & vbCrLf & _
                              " 										AND ISD.PartNo = PRD.PartNo " & vbCrLf & _
                              " 										AND ISD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                              " 	LEFT JOIN InvoiceSupplier_Master ISM ON ISM.InvoiceNo = ISD.InvoiceNo " & vbCrLf & _
                              "   										AND ISM.AffiliateID = ISD.AffiliateID " & vbCrLf & _
                              "   										AND ISM.SupplierID = ISD.SupplierID " & vbCrLf & _
                              "   										AND ISM.suratJalanno = ISD.SuratJalanNo " & vbCrLf & _
                              " 	LEFT JOIN DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID " & vbCrLf & _
                              " 								AND PRD.KanbanNo = PDD.KanbanNo " & vbCrLf

            ls_SQL = ls_SQL + " 								AND PRD.SupplierID = PDD.SupplierID " & vbCrLf & _
                              " 								AND PRD.PartNo = PDD.PartNo " & vbCrLf & _
                              " 								AND PRD.PONo = PDD.PONo " & vbCrLf & _
                              " 								AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf & _
                              " 	LEFT JOIN DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                              " 								AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
                              " 	LEFT JOIN ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                              " 									AND PDD.KanbanNo = RAD.KanbanNo " & vbCrLf & _
                              " 									AND PDD.SupplierID = RAD.SupplierID " & vbCrLf & _
                              " 									AND PDD.PartNo = RAD.PartNo " & vbCrLf & _
                              " 									AND PDD.PONo = RAD.PONo " & vbCrLf

            ls_SQL = ls_SQL + " 									AND PDD.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                              " 	LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                              " 									AND RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                              " 	LEFT JOIN PLPASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
                              " 									AND PDD.KanbanNo = IPD.KanbanNo								 " & vbCrLf & _
                              " 									AND PDD.PartNo = IPD.PartNo " & vbCrLf & _
                              " 									AND PDD.PONo = IPD.PONo " & vbCrLf & _
                              " 									AND PDD.SuratJalanNo = IPD.SuratJalanNo " & vbCrLf & _
                              " 	LEFT JOIN PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID " & vbCrLf & _
                              " 									--AND IPD.InvoiceNo = IPM.InvoiceNo " & vbCrLf & _
                              " 									AND IPD.SuratJalanNo = IPM.SuratJalanNo " & vbCrLf

            ls_SQL = ls_SQL + " 	LEFT JOIN (  " & vbCrLf & _
                              "  				SELECT * FROM MS_ETD_PASI a  " & vbCrLf & _
                              "  				INNER JOIN MS_ETD_Supplier_PASI b on a.ETDPASI =  b.ETAPASI  " & vbCrLf & _
                              "  				)ABC ON POM.SupplierID = ABC.SupplierID and POM.AffiliateID = ABC.AffiliateID AND KM.KanbanDate = ABC.ETAAffiliate  " & vbCrLf & _
                              " 	LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                              " 	LEFT JOIN MS_CurrCls MCS ON MCS.CurrCls = ISD.InvCurrCls " & vbCrLf & _
                              " 	LEFT JOIN MS_Price MSP ON MSP.AffiliateID = IPD.AffiliateID and MSP.PartNo = IPD.PartNo and (IPM.DeliveryDate between MSP.StartDate and MSP.EndDate)  " & vbCrLf & _
                              " 	WHERE KD.KanbanQty > 0 " & vbCrLf

            ls_SQL = ls_SQL + ls_filter & vbCrLf

            ls_SQL = ls_SQL + " )XYZ " & vbCrLf & _
                              "  "
            Dim cmd As New SqlCommand(ls_SQL, sqlConn)
            cmd.CommandTimeout = 300
            Dim sqlDA As New SqlDataAdapter
            sqlDA.SelectCommand = cmd
            Dim ds As New DataSet
            sqlDA.SelectCommand.CommandTimeout = 300
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Function GetSummaryOutStanding_Old() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                Dim ls_End As String = ""
                ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

                'AFFILIATE CODE
                If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                End If
                'AFFILIATE PO PERIOD
                If chkPOPeriod.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyyMM01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyyMM" & ls_End) & "' " & vbCrLf
                End If
                'SUPPLIER PLAN DELIVERY DATE
                If chkSupplierPlanDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,KM.KanbanDate) BETWEEN '" & Format(dtSupplierPlanDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtSupplierPlanDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If
                'SUPPLIER DELIVERY DATE
                If chkSupplierDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,SDM.DeliveryDate) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If
                'PASI RECEIVE DATE
                If chkPASIRecDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,PRM.ReceiveDate) BETWEEN '" & Format(dtPASIRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If
                'PASI DELIVERY DATE
                If chkPASIDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,PDM.DeliveryDate) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If
                'AFFILIATE RECEIVE DATE
                If chkAffiliateRecDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,RAM.ReceiveDate) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
                End If

                'PART CODE
                If Trim(cboPart.Text) <> "==ALL==" And Trim(cboPart.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND POD.PartNo = '" & Trim(cboPart.Text) & "' " & vbCrLf
                End If
                'PONO
                If Trim(txtPONo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(POM.PONo,'') LIKE '%" & Trim(txtPONo.Text) & "%' " & vbCrLf
                End If
                'SUPPLIER SJ NO
                If Trim(txtSupplierSJNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(SDD.SuratJalanNo,'') LIKE '%" & Trim(txtSupplierSJNo.Text) & "%'" & vbCrLf
                End If
                'PASI SJ NO
                If Trim(txtPASISJNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(PDD.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
                End If
                'SUPPLIER INV NO
                If Trim(txtSupplierInvNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(ISD.InvoiceNo,'') LIKE '%" & Trim(txtSupplierInvNo.Text) & "%'" & vbCrLf
                End If
                'PASI INV NO
                If Trim(txtPASIInvNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(IPD.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
                End If
                'PO PROGRESS

                'PASI RECEIVE
                If rdrPRComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0) " & vbCrLf
                ElseIf rdrPRRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(SDD.DOQty,0) > ISNULL(PRD.GoodRecQty,0)  " & vbCrLf
                ElseIf rdrPRDiff.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND (ISNULL(SDD.DOQty,0) < ISNULL(PRD.GoodRecQty,0) OR ISNULL(PRD.DefectRecQty,0) > 0) " & vbCrLf
                End If

                'PASI DELIVERY
                If rdrPDComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) = ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
                ElseIf rdrPDComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) > (PDD.DOQty) " & vbCrLf & _
                                  "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
                    'ElseIf rdPRDDiff.Value = True Then
                    '    ls_SQL = ls_SQL + _
                    '                  "                      AND (ISNULL(PRD.DefectRecQty,0) > 0 " & vbCrLf & _
                    '                  "                         AND ISNULL(PDD.DOQty,0) > 0) " & vbCrLf
                End If

                'AFFILIATE RECEIVE
                If rdrARComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                            END) = RAD.RecQty AND " & vbCrLf & _
                                  "                           (CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.DefectQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.DefectRecQty,0) " & vbCrLf & _
                                  "                            END) = 0 " & vbCrLf & _
                                  "                          ) " & vbCrLf
                ElseIf rdrARRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                            END) > RAD.RecQty " & vbCrLf & _
                                  "                          ) " & vbCrLf
                ElseIf rdrARDiff.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                            END) < RAD.RecQty OR " & vbCrLf & _
                                  "                           ISNULL(RAD.DefectQty,0) > 0  " & vbCrLf & _
                                  "                          )" & vbCrLf
                End If

                'SUPPLIER INVOICE
                If rdrSIComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
                                  "                            END) = ISD.InvQty " & vbCrLf & _
                                  "                          ) " & vbCrLf
                ElseIf rdrSIRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
                                  "                            END) > ISD.InvQty " & vbCrLf & _
                                  "                          ) " & vbCrLf
                End If

                'PASI INVOICE
                If rdrPIComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(RAD.RecQty,0) = ISNULL(IPD.DOQty,0) " & vbCrLf
                ElseIf rdrPIRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(RAD.RecQty,0) > ISNULL(IPD.DOQty,0) " & vbCrLf
                End If

                ls_sql = " SELECT DISTINCT * FROM " & vbCrLf & _
                      " ( " & vbCrLf & _
                      " 	SELECT  " & vbCrLf & _
                      " 		POM.Period " & vbCrLf & _
                      " 		,POM.PONo " & vbCrLf & _
                      " 		,POM.AffiliateID " & vbCrLf & _
                      " 		,POM.SupplierID " & vbCrLf & _
                      " 		,POKanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '1' THEN 'YES' ELSE 'NO' END " & vbCrLf & _
                      " 		,POM.EntryDate " & vbCrLf & _
                      " 		,POM.PASISendAffiliateDate " & vbCrLf & _
                      " 		,POD.PartNo " & vbCrLf

                ls_sql = ls_sql + " 		,MP.PartName " & vbCrLf & _
                                  " 		,QtyPO = ISNULL(POD.POQty,0) " & vbCrLf & _
                                  " 		,KD.KanbanNo " & vbCrLf & _
                                  " 		,KD.KanbanQty " & vbCrLf & _
                                  " 		,ETDSupp = ABC.ETDSupplier " & vbCrLf & _
                                  " 		,ETAAff = KM.KanbanDate " & vbCrLf & _
                                  " 		,SupplierDeliveryDate = SDM.DeliveryDate " & vbCrLf & _
                                  " 		,SupplierSuratJalanNo = SDM.SuratJalanNo " & vbCrLf & _
                                  " 		,SupplierDeliveryQty = SDD.DOQty " & vbCrLf & _
                                  " 		,RemainingQtyPOPASI = ISNULL(KD.KanbanQty,0) - " & vbCrLf & _
                                  " 		                      ISNULL( " & vbCrLf & _
                                  " 		                        (select SUM(DOQty) from DOSupplier_Detail ABC " & vbCrLf & _
                                  " 		                         WHERE ABC.SupplierID = SDD.SupplierID and ABC.AffiliateID = SDD.AffiliateID" & vbCrLf & _
                                  " 		                         and ABC.KanbanNo = SDD.KanbanNo and ABC.PartNo = SDD.PartNo and ABC.PONo = SDD.PONo and ABC.SuratJalanNo = SDD.SuratJalanNo),0) " & vbCrLf & _
                                  " 		,PASIReceiveDate = PRM.ReceiveDate " & vbCrLf

                ls_sql = ls_sql + " 		,PASIReceivingQty = PRD.GoodRecQty " & vbCrLf & _
                                  " 		,InvoiceNoFromSupplier = ISM.InvoiceNo " & vbCrLf & _
                                  " 		,InvoiceDateFromSupplier = ISM.InvoiceDate " & vbCrLf & _
                                  " 		,InvoiceFromSupplierCurr = MCS.Description " & vbCrLf & _
                                  " 		,InvoiceFromSupplierAmount = ISNULL(ISD.InvAmount,0) " & vbCrLf & _
                                  " 		,PASIDeliveryDate = PDM.DeliveryDate " & vbCrLf & _
                                  " 		,PASISuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
                                  " 		,PASIDeliveryQty = PDD.DOQty " & vbCrLf & _
                                  " 		,AffiliateReceiveDate = RAM.ReceiveDate " & vbCrLf & _
                                  " 		,AffiliateReceivingQty = RAD.RecQty " & vbCrLf & _
                                  " 		,InvoiceNoToAffiliate = IPM.InvoiceNo " & vbCrLf

                ls_sql = ls_sql + " 		,InvoiceDateToAffiliate = IPM.DeliveryDate " & vbCrLf & _
                                  " 		,InvoiceToAffiliateCurr = 'IDR' " & vbCrLf & _
                                  " 		,InvoiceToAffiliateAmount = ISNULL(IPD.DOQty,0) * ISNULL(PDD.Price,0) " & vbCrLf & _
                                  " 	FROM PO_Master POM " & vbCrLf & _
                                  " 	LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                                  " 							AND POM.PoNo = POD.PONo " & vbCrLf & _
                                  " 							AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
                                  " 								AND KD.PoNo = POD.PONo " & vbCrLf & _
                                  " 								AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
                                  " 								AND KD.PartNo = POD.PartNo " & vbCrLf

                ls_sql = ls_sql + " 	LEFT JOIN Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                                  " 								AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                                  " 								AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                                  " 								AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                                  " 	LEFT JOIN DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                                  " 									AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
                                  " 									AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
                                  " 									AND KD.PartNo = SDD.PartNo " & vbCrLf & _
                                  " 									AND KD.PoNo = SDD.PoNo " & vbCrLf & _
                                  " 	LEFT JOIN DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                                  " 									AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf

                ls_sql = ls_sql + " 									AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                                  " 									AND SDD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                                  " 									AND SDD.SupplierID = PRD.SupplierID " & vbCrLf & _
                                  " 									AND SDD.PartNo = PRD.PartNo " & vbCrLf & _
                                  " 									AND SDD.PONo = PRD.PONo								 " & vbCrLf & _
                                  " 									AND SDD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                                  " 									AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                                  " 									AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN InvoiceSupplier_Detail ISD ON ISD.AffiliateID = PRD.AffiliateID " & vbCrLf

                ls_sql = ls_sql + " 										AND ISD.SupplierID = PRD.SupplierID " & vbCrLf & _
                                  " 										AND ISD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                                  " 										AND ISD.PONo = PRD.PONo " & vbCrLf & _
                                  " 										AND ISD.PartNo = PRD.PartNo " & vbCrLf & _
                                  " 										AND ISD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                                  " 	LEFT JOIN InvoiceSupplier_Master ISM ON ISM.InvoiceNo = ISD.InvoiceNo " & vbCrLf & _
                                  "   										AND ISM.AffiliateID = ISD.AffiliateID " & vbCrLf & _
                                  "   										AND ISM.SupplierID = ISD.SupplierID " & vbCrLf & _
                                  "   										AND ISM.suratJalanno = ISD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID " & vbCrLf & _
                                  " 								AND PRD.KanbanNo = PDD.KanbanNo " & vbCrLf

                ls_sql = ls_sql + " 								AND PRD.SupplierID = PDD.SupplierID " & vbCrLf & _
                                  " 								AND PRD.PartNo = PDD.PartNo " & vbCrLf & _
                                  " 								AND PRD.PONo = PDD.PONo " & vbCrLf & _
                                  " 								AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf & _
                                  " 	LEFT JOIN DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                                  " 								AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                                  " 									AND PDD.KanbanNo = RAD.KanbanNo " & vbCrLf & _
                                  " 									AND PDD.SupplierID = RAD.SupplierID " & vbCrLf & _
                                  " 									AND PDD.PartNo = RAD.PartNo " & vbCrLf & _
                                  " 									AND PDD.PONo = RAD.PONo " & vbCrLf

                ls_sql = ls_sql + " 									AND PDD.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                                  " 									AND RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                                  " 	LEFT JOIN PLPASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
                                  " 									AND PDD.KanbanNo = IPD.KanbanNo								 " & vbCrLf & _
                                  " 									AND PDD.PartNo = IPD.PartNo " & vbCrLf & _
                                  " 									AND PDD.PONo = IPD.PONo " & vbCrLf & _
                                  " 									AND PDD.SuratJalanNo = IPD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID " & vbCrLf & _
                                  " 									--AND IPD.InvoiceNo = IPM.InvoiceNo " & vbCrLf & _
                                  " 									AND IPD.SuratJalanNo = IPM.SuratJalanNo " & vbCrLf

                ls_sql = ls_sql + " 	LEFT JOIN (  " & vbCrLf & _
                                  "  				SELECT * FROM MS_ETD_PASI a  " & vbCrLf & _
                                  "  				INNER JOIN MS_ETD_Supplier_PASI b on a.ETDPASI =  b.ETAPASI  " & vbCrLf & _
                                  "  				)ABC ON POM.SupplierID = ABC.SupplierID and POM.AffiliateID = ABC.AffiliateID AND KM.KanbanDate = ABC.ETAAffiliate  " & vbCrLf & _
                                  " 	LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                                  " 	LEFT JOIN MS_CurrCls MCS ON MCS.CurrCls = ISD.InvCurrCls " & vbCrLf & _
                                  " 	LEFT JOIN MS_Price MSP ON MSP.AffiliateID = IPD.AffiliateID and MSP.PartNo = IPD.PartNo and (IPM.DeliveryDate between MSP.StartDate and MSP.EndDate)  " & vbCrLf & _
                                  " 	WHERE KD.KanbanQty > 0 " & vbCrLf

                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " )XYZ " & vbCrLf & _
                                  "  "

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 300
                da.Fill(dt)

                Return dt
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                Dim ls_End As String = ""
                ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

                'AFFILIATE CODE
                If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                End If

                'AFFILIATE PO PERIOD
                If chkPOPeriod.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyy-MM-01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyy-MM-" & ls_End) & "' " & vbCrLf
                End If

                'SUPPLIER PLAN DELIVERY DATE
                If chkSupplierPlanDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,KM.KanbanDate) BETWEEN '" & Format(dtSupplierPlanDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierPlanDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                'SUPPLIER DELIVERY DATE
                If chkSupplierDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,SDM.DeliveryDate) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                'PASI RECEIVE DATE
                If chkPASIRecDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,PRM.ReceiveDate) BETWEEN '" & Format(dtPASIRecDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtPASIRecDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                'PASI DELIVERY DATE
                If chkPASIDelDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,PDM.DeliveryDate) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                'AFFILIATE RECEIVE DATE
                If chkAffiliateRecDate.Checked = True Then
                    ls_filter = ls_filter + _
                                  "                      AND CONVERT(date,RAM.ReceiveDate) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyy-MM-dd") & "' " & vbCrLf
                End If

                'PART CODE
                If Trim(cboPart.Text) <> "==ALL==" And Trim(cboPart.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND POD.PartNo = '" & Trim(cboPart.Text) & "' " & vbCrLf
                End If

                'PONO
                If Trim(txtPONo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(POM.PONo,'') LIKE '%" & Trim(txtPONo.Text) & "%' " & vbCrLf
                End If

                'SUPPLIER SJ NO
                If Trim(txtSupplierSJNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(SDD.SuratJalanNo,'') LIKE '%" & Trim(txtSupplierSJNo.Text) & "%'" & vbCrLf
                End If

                'PASI SJ NO
                If Trim(txtPASISJNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(PDD.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
                End If

                'SUPPLIER INV NO
                If Trim(txtSupplierInvNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(ISD.InvoiceNo,'') LIKE '%" & Trim(txtSupplierInvNo.Text) & "%'" & vbCrLf
                End If

                'PASI INV NO
                If Trim(txtPASIInvNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(IPD.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
                End If
                'PO PROGRESS

                'PASI RECEIVE
                If rdrPRComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0) " & vbCrLf
                ElseIf rdrPRRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(SDD.DOQty,0) > ISNULL(PRD.GoodRecQty,0)  " & vbCrLf
                ElseIf rdrPRDiff.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND (ISNULL(SDD.DOQty,0) < ISNULL(PRD.GoodRecQty,0) OR ISNULL(PRD.DefectRecQty,0) > 0) " & vbCrLf
                End If

                'PASI DELIVERY
                If rdrPDComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) = ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
                ElseIf rdrPDComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) > (PDD.DOQty) " & vbCrLf & _
                                  "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
                End If

                'AFFILIATE RECEIVE
                If rdrARComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                            END) = RAD.RecQty AND " & vbCrLf & _
                                  "                           (CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.DefectQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.DefectRecQty,0) " & vbCrLf & _
                                  "                            END) = 0 " & vbCrLf & _
                                  "                          ) " & vbCrLf
                ElseIf rdrARRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                            END) > RAD.RecQty " & vbCrLf & _
                                  "                          ) " & vbCrLf
                ElseIf rdrARDiff.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
                                  "                            END) < RAD.RecQty OR " & vbCrLf & _
                                  "                           ISNULL(RAD.DefectQty,0) > 0  " & vbCrLf & _
                                  "                          )" & vbCrLf
                End If

                'SUPPLIER INVOICE
                If rdrSIComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
                                  "                            END) = ISD.InvQty " & vbCrLf & _
                                  "                          ) " & vbCrLf
                ElseIf rdrSIRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
                                  "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
                                  "                            END) > ISD.InvQty " & vbCrLf & _
                                  "                          ) " & vbCrLf
                End If

                'PASI INVOICE
                If rdrPIComplete.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(RAD.RecQty,0) = ISNULL(IPD.DOQty,0) " & vbCrLf
                ElseIf rdrPIRemaining.Value = True Then
                    ls_filter = ls_filter + _
                                  "                      AND ISNULL(RAD.RecQty,0) > ISNULL(IPD.DOQty,0) " & vbCrLf
                End If

                ls_sql = " SELECT DISTINCT * FROM " & vbCrLf & _
                      " ( " & vbCrLf & _
                      " 	SELECT  " & vbCrLf & _
                      " 		POM.Period " & vbCrLf & _
                      " 		,RTRIM(POM.PONo) PONo " & vbCrLf & _
                      " 		,RTRIM(POM.AffiliateID) AffiliateID " & vbCrLf & _
                      " 		,RTRIM(POM.SupplierID) SupplierID " & vbCrLf & _
                      " 		,POKanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '1' THEN 'YES' ELSE 'NO' END " & vbCrLf & _
                      " 		,POM.EntryDate " & vbCrLf & _
                      " 		,POM.PASISendAffiliateDate " & vbCrLf & _
                      " 		,RTRIM(POD.PartNo) PartNo " & vbCrLf

                ls_sql = ls_sql + " 		,RTRIM(MP.PartName) PartName " & vbCrLf & _
                                  " 		,QtyPO = ISNULL(POD.POQty,0) " & vbCrLf & _
                                  " 		,RTRIM(KD.KanbanNo) KanbanNo " & vbCrLf & _
                                  " 		,KD.KanbanQty " & vbCrLf & _
                                  " 		,ETDSupp = ABC.ETDSupplier " & vbCrLf & _
                                  " 		,ETAAff = KM.KanbanDate " & vbCrLf & _
                                  " 		,SupplierDeliveryDate = SDM.DeliveryDate " & vbCrLf & _
                                  " 		,SupplierSuratJalanNo = RTRIM(SDM.SuratJalanNo) " & vbCrLf & _
                                  " 		,SupplierDeliveryQty = SDD.DOQty " & vbCrLf & _
                                  " 		,RemainingQtyPOPASI = ISNULL(KD.KanbanQty,0) - " & vbCrLf & _
                                  " 		                      ISNULL( " & vbCrLf & _
                                  " 		                        (select SUM(DOQty) from DOSupplier_Detail ABC " & vbCrLf & _
                                  " 		                         INNER JOIN DOSupplier_Master EFG ON ABC.SuratJalanNo = EFG.SuratJalanNo AND ABC.SupplierID = EFG.SupplierID AND ABC.AffiliateID = EFG.AffiliateID " & vbCrLf & _
                                  " 		                         WHERE ABC.SupplierID = SDD.SupplierID and ABC.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                                  " 		                         and ABC.KanbanNo = SDD.KanbanNo and ABC.PartNo = SDD.PartNo and ABC.PONo = SDD.PONo AND EFG.SuratJalanNo <= SDM.SuratJalanNo),0) " & vbCrLf & _
                                  " 		,PASIReceiveDate = PRM.ReceiveDate " & vbCrLf

                ls_sql = ls_sql + " 		,PASIReceivingQty = PRD.GoodRecQty " & vbCrLf & _
                                  " 		,InvoiceNoFromSupplier = ISM.InvoiceNo " & vbCrLf & _
                                  " 		,InvoiceDateFromSupplier = ISM.InvoiceDate " & vbCrLf & _
                                  " 		,InvoiceFromSupplierCurr = RTRIM(MCS.Description) " & vbCrLf & _
                                  " 		,InvoiceFromSupplierAmount = ISNULL(ISD.InvAmount,0) " & vbCrLf & _
                                  " 		,PASIDeliveryDate = PDM.DeliveryDate " & vbCrLf & _
                                  " 		,PASISuratJalanNo = RTRIM(IPM.SuratJalanNo) " & vbCrLf & _
                                  " 		,PASIDeliveryQty = IPD.DOQty " & vbCrLf & _
                                  " 		,AffiliateReceiveDate = RAM.ReceiveDate " & vbCrLf & _
                                  " 		,AffiliateReceivingQty = RAD.RecQty " & vbCrLf & _
                                  " 		,InvoiceNoToAffiliate = RTRIM(IPM.InvoiceNo) " & vbCrLf

                ls_sql = ls_sql + " 		,InvoiceDateToAffiliate = IPM.DeliveryDate " & vbCrLf & _
                                  " 		,InvoiceToAffiliateCurr = 'IDR' " & vbCrLf & _
                                  " 		,InvoiceToAffiliateAmount = ISNULL(IPD.DOQty,0) * ISNULL(PDD.Price,0) " & vbCrLf & _
                                  " 	FROM PO_Master POM " & vbCrLf & _
                                  " 	LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
                                  " 							AND POM.PoNo = POD.PONo " & vbCrLf & _
                                  " 							AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
                                  " 								AND KD.PoNo = POD.PONo " & vbCrLf & _
                                  " 								AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
                                  " 								AND KD.PartNo = POD.PartNo " & vbCrLf

                ls_sql = ls_sql + " 	LEFT JOIN Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
                                  " 								AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
                                  " 								AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
                                  " 								AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
                                  " 	LEFT JOIN DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                                  " 									AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
                                  " 									AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
                                  " 									AND KD.PartNo = SDD.PartNo " & vbCrLf & _
                                  " 									AND KD.PoNo = SDD.PoNo " & vbCrLf & _
                                  " 	LEFT JOIN DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf & _
                                  " 									AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf

                ls_sql = ls_sql + " 									AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                                  " 									AND SDD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                                  " 									AND SDD.SupplierID = PRD.SupplierID " & vbCrLf & _
                                  " 									AND SDD.PartNo = PRD.PartNo " & vbCrLf & _
                                  " 									AND SDD.PONo = PRD.PONo								 " & vbCrLf & _
                                  " 									AND SDD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID " & vbCrLf & _
                                  " 									AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                                  " 									AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
                                  " 	LEFT JOIN InvoiceSupplier_Detail ISD ON ISD.AffiliateID = PRD.AffiliateID " & vbCrLf

                ls_sql = ls_sql + " 										AND ISD.SupplierID = PRD.SupplierID " & vbCrLf & _
                                  " 										AND ISD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
                                  " 										AND ISD.PONo = PRD.PONo " & vbCrLf & _
                                  " 										AND ISD.PartNo = PRD.PartNo " & vbCrLf & _
                                  " 										AND ISD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
                                  " 	LEFT JOIN InvoiceSupplier_Master ISM ON ISM.InvoiceNo = ISD.InvoiceNo " & vbCrLf & _
                                  "   										AND ISM.AffiliateID = ISD.AffiliateID " & vbCrLf & _
                                  "   										AND ISM.SupplierID = ISD.SupplierID " & vbCrLf & _
                                  "   										AND ISM.suratJalanno = ISD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID " & vbCrLf & _
                                  " 								AND PRD.KanbanNo = PDD.KanbanNo " & vbCrLf

                ls_sql = ls_sql + " 								AND PRD.SupplierID = PDD.SupplierID " & vbCrLf & _
                                  " 								AND PRD.PartNo = PDD.PartNo " & vbCrLf & _
                                  " 								AND PRD.PONo = PDD.PONo " & vbCrLf & _
                                  " 								AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf & _
                                  " 	LEFT JOIN DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
                                  " 								AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                                  " 									AND PDD.KanbanNo = RAD.KanbanNo " & vbCrLf & _
                                  " 									AND PDD.SupplierID = RAD.SupplierID " & vbCrLf & _
                                  " 									AND PDD.PartNo = RAD.PartNo " & vbCrLf & _
                                  " 									AND PDD.PONo = RAD.PONo " & vbCrLf

                ls_sql = ls_sql + " 									AND PDD.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
                                  " 									AND RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
                                  " 	LEFT JOIN PLPASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
                                  " 									AND PDD.KanbanNo = IPD.KanbanNo								 " & vbCrLf & _
                                  " 									AND PDD.PartNo = IPD.PartNo " & vbCrLf & _
                                  " 									AND PDD.PONo = IPD.PONo " & vbCrLf & _
                                  " 									AND PDD.SuratJalanNo = IPD.SuratJalanNo " & vbCrLf & _
                                  " 	LEFT JOIN PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID " & vbCrLf & _
                                  " 									--AND IPD.InvoiceNo = IPM.InvoiceNo " & vbCrLf & _
                                  " 									AND IPD.SuratJalanNo = IPM.SuratJalanNo " & vbCrLf

                ls_sql = ls_sql + " 	LEFT JOIN (  " & vbCrLf & _
                                  "  				SELECT * FROM MS_ETD_PASI a  " & vbCrLf & _
                                  "  				INNER JOIN MS_ETD_Supplier_PASI b on a.ETDPASI =  b.ETAPASI  " & vbCrLf & _
                                  "  				)ABC ON POM.SupplierID = ABC.SupplierID and POM.AffiliateID = ABC.AffiliateID AND KM.KanbanDate = ABC.ETAAffiliate  " & vbCrLf & _
                                  " 	LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
                                  " 	LEFT JOIN MS_CurrCls MCS ON MCS.CurrCls = ISD.InvCurrCls " & vbCrLf & _
                                  " 	LEFT JOIN MS_Price MSP ON MSP.AffiliateID = IPD.AffiliateID and MSP.PartNo = IPD.PartNo and (IPM.DeliveryDate between MSP.StartDate and MSP.EndDate)  " & vbCrLf & _
                                  " 	WHERE KD.KanbanQty > 0 " & vbCrLf

                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " )XYZ " & vbCrLf & _
                                  "  "

                Dim Cmd As New SqlCommand(ls_sql, cn)
                Dim da As New SqlDataAdapter(Cmd)
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 300
                da.Fill(dt)

                Return dt
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub up_GridLoadWhenEventChange()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()

            ls_SQL = "SELECT TOP 0 " & vbCrLf & _
                     " 	     ColNo = 0,  " & vbCrLf & _
                     "       Period = '', AffiliateCode = '', AffiliateName = '', PONo = '', SupplierCode = '', SupplierName = '', POKanban = '',   " & vbCrLf & _
                     " 	     KanbanNo = '', SupplierPlanDeliveryDate = '', SupplierDeliveryDate = '', SupplierSJNo = '', PASIReceiveDate = '', PASIDeliveryDate = '',  " & vbCrLf & _
                     " 	     PASISJNo = '', AffiliateReceiveDate = '', PartNo = '', PartName = '', UOM = '', SupplierDeliveryQty = '', PASIReceivingQty = '', PASIDeliveryQty = '', AffiliateReceivingQty = '', InvoiceFromSupplierQty = '',  " & vbCrLf & _
                     " 	     InvoiceToAffiliateQty = '', InvoiceNoFromSupplier = '', InvoiceDataFromSupplier = '', InvoiceFromSupplierCurr = '', " & vbCrLf & _
                     "       InvoiceFromSupplierAmount = '', InvoiceNoToAffiliate = '', InvoiceDateToAffiliate = '', InvoiceToAffiliateCurr = '', InvoiceToAffiliateAmount = '', " & vbCrLf & _
                     " 	     SortPONo = '', SortKanbanNo = '', SortHeader = 0, PODelivery = '' "

            Dim sqlDA As New SqlDataAdapter(ls_SQL, sqlConn)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            With grid
                .DataSource = ds.Tables(0)
                .DataBind()
            End With
        End Using
    End Sub

    Private Sub up_FillCombo()
        Dim sqlDA As New SqlDataAdapter()
        Dim ds As New DataSet

        'Combo Affiliate
        With cboAffiliateCode
            ls_SQL = "--SELECT AffiliateID = '==ALL==', AffiliateName = '==ALL=='" & vbCrLf & _
                     " --UNION ALL " & vbCrLf & _
                     "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate Where isnull(overseascls, '0') = '0'"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("AffiliateID")
                .Columns(0).Width = 90
                .Columns.Add("AffiliateName")
                .Columns(1).Width = 240

                .TextField = "AffiliateID"
                .DataBind()
            End Using
        End With


        'Combo Parts
        With cboPart
            ls_SQL = "SELECT PartNo = '==ALL==', PartName = '==ALL=='" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     "SELECT PartNo = RTRIM(PartNo), PartName = RTRIM(PartName) FROM dbo.MS_Parts"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("PartNo")
                .Columns(0).Width = 90
                .Columns.Add("PartName")
                .Columns(1).Width = 240

                .TextField = "PartNo"
                .DataBind()
            End Using
        End With
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

    Private Sub epplusExportExcel(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Summary Outstanding " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
            With ws
                .Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 34).AutoFitColumns()
                .Cells(8, 1, pData.Rows.Count + 7, 34).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

                .Cells(8, 1, pData.Rows.Count + 7, 1).Style.Numberformat.Format = "mmm-yy"
                .Cells(8, 6, pData.Rows.Count + 7, 7).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 13, pData.Rows.Count + 7, 15).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 19, pData.Rows.Count + 7, 19).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 22, pData.Rows.Count + 7, 22).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 25, pData.Rows.Count + 7, 25).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 28, pData.Rows.Count + 7, 28).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 31, pData.Rows.Count + 7, 31).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 10, pData.Rows.Count + 7, 10).Style.Numberformat.Format = "#,##0"
                .Cells(8, 12, pData.Rows.Count + 7, 12).Style.Numberformat.Format = "#,##0"
                .Cells(8, 17, pData.Rows.Count + 7, 18).Style.Numberformat.Format = "#,##0"
                .Cells(8, 20, pData.Rows.Count + 7, 20).Style.Numberformat.Format = "#,##0"
                .Cells(8, 24, pData.Rows.Count + 7, 24).Style.Numberformat.Format = "#,##0"
                .Cells(8, 27, pData.Rows.Count + 7, 27).Style.Numberformat.Format = "#,##0"
                .Cells(8, 29, pData.Rows.Count + 7, 29).Style.Numberformat.Format = "#,##0"
                .Cells(8, 33, pData.Rows.Count + 7, 33).Style.Numberformat.Format = "#,##0"

                'Dim irow As Integer = 0
                'Dim irowtmp1 As Integer = 0
                'Dim irowtmp2 As Integer = 0
                'Dim sKey1 As String = ""
                'Dim sKey2 As String = ""

                'For irow = 8 To pData.Rows.Count + 7
                '    If irow = 8 Then
                '        sKey1 = Trim(.Cells(irow, 2).Text) & Trim(.Cells(irow, 3).Text) & Trim(.Cells(irow, 4).Text) & Trim(.Cells(irow, 8).Text)
                '        sKey2 = Trim(.Cells(irow, 2).Text) & Trim(.Cells(irow, 3).Text) & Trim(.Cells(irow, 4).Text) & Trim(.Cells(irow, 8).Text) & Trim(.Cells(irow, 16).Text)
                '        irowtmp1 = irow
                '        irowtmp2 = irow
                '    End If

                '    If Trim(sKey1) <> Trim(.Cells(irow, 2).Text) & Trim(.Cells(irow, 3).Text) & Trim(.Cells(irow, 4).Text) & Trim(.Cells(irow, 8).Text) Then
                '        .Cells(irowtmp1, 1, irow - 1, 1).Merge = True
                '        .Cells(irowtmp1, 2, irow - 1, 2).Merge = True
                '        .Cells(irowtmp1, 3, irow - 1, 3).Merge = True
                '        .Cells(irowtmp1, 4, irow - 1, 4).Merge = True
                '        .Cells(irowtmp1, 5, irow - 1, 5).Merge = True
                '        .Cells(irowtmp1, 6, irow - 1, 6).Merge = True
                '        .Cells(irowtmp1, 7, irow - 1, 7).Merge = True
                '        .Cells(irowtmp1, 8, irow - 1, 8).Merge = True
                '        .Cells(irowtmp1, 9, irow - 1, 9).Merge = True
                '        .Cells(irowtmp1, 10, irow - 1, 10).Merge = True
                '        .Cells(irowtmp1, 11, irow - 1, 11).Merge = True
                '        .Cells(irowtmp1, 12, irow - 1, 12).Merge = True
                '        .Cells(irowtmp1, 13, irow - 1, 13).Merge = True
                '        .Cells(irowtmp1, 14, irow - 1, 14).Merge = True
                '        .Cells(irowtmp1, 15, irow - 1, 15).Merge = True

                '        sKey1 = Trim(.Cells(irow, 2).Text) & Trim(.Cells(irow, 3).Text) & Trim(.Cells(irow, 4).Text) & Trim(.Cells(irow, 8).Text)
                '        irowtmp1 = irow
                '    End If

                '    If Trim(sKey2) <> Trim(.Cells(irow, 2).Text) & Trim(.Cells(irow, 3).Text) & Trim(.Cells(irow, 4).Text) & Trim(.Cells(irow, 8).Text) & Trim(.Cells(irow, 16).Text) Then
                '        .Cells(irowtmp2, 16, irow - 1, 16).Merge = True
                '        .Cells(irowtmp2, 17, irow - 1, 17).Merge = True
                '        .Cells(irowtmp2, 18, irow - 1, 18).Merge = True
                '        .Cells(irowtmp2, 19, irow - 1, 19).Merge = True
                '        .Cells(irowtmp2, 20, irow - 1, 20).Merge = True
                '        .Cells(irowtmp2, 21, irow - 1, 21).Merge = True
                '        .Cells(irowtmp2, 22, irow - 1, 22).Merge = True
                '        .Cells(irowtmp2, 23, irow - 1, 23).Merge = True
                '        .Cells(irowtmp2, 24, irow - 1, 24).Merge = True

                '        sKey2 = Trim(.Cells(irow, 2).Text) & Trim(.Cells(irow, 3).Text) & Trim(.Cells(irow, 4).Text) & Trim(.Cells(irow, 8).Text) & Trim(.Cells(irow, 16).Text)
                '        irowtmp2 = irow
                '    End If

                '    If irow = pData.Rows.Count + 7 Then
                '        If irow <> irowtmp1 Then
                '            .Cells(irowtmp1, 1, irow, 1).Merge = True
                '            .Cells(irowtmp1, 2, irow, 2).Merge = True
                '            .Cells(irowtmp1, 3, irow, 3).Merge = True
                '            .Cells(irowtmp1, 4, irow, 4).Merge = True
                '            .Cells(irowtmp1, 5, irow, 5).Merge = True
                '            .Cells(irowtmp1, 6, irow, 6).Merge = True
                '            .Cells(irowtmp1, 7, irow, 7).Merge = True
                '            .Cells(irowtmp1, 8, irow, 8).Merge = True
                '            .Cells(irowtmp1, 9, irow, 9).Merge = True
                '            .Cells(irowtmp1, 10, irow, 10).Merge = True
                '            .Cells(irowtmp1, 11, irow, 11).Merge = True
                '            .Cells(irowtmp1, 12, irow, 12).Merge = True
                '            .Cells(irowtmp1, 13, irow, 13).Merge = True
                '            .Cells(irowtmp1, 14, irow, 14).Merge = True
                '            .Cells(irowtmp1, 15, irow, 15).Merge = True
                '        End If

                '        If irow <> irowtmp2 Then
                '            .Cells(irowtmp2, 16, irow, 16).Merge = True
                '            .Cells(irowtmp2, 17, irow, 17).Merge = True
                '            .Cells(irowtmp2, 18, irow, 18).Merge = True
                '            .Cells(irowtmp2, 19, irow, 19).Merge = True
                '            .Cells(irowtmp2, 20, irow, 20).Merge = True
                '            .Cells(irowtmp2, 21, irow, 21).Merge = True
                '            .Cells(irowtmp2, 22, irow, 22).Merge = True
                '            .Cells(irowtmp2, 23, irow, 23).Merge = True
                '            .Cells(irowtmp2, 24, irow, 24).Merge = True
                '        End If
                '    End If
                'Next

                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 33)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub epplusExportExcelOLD(ByVal pFilename As String, ByVal pSheetName As String,
                              ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

        Try
            Dim tempFile As String = "Summary Outstanding " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                .Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                For irow = 0 To pData.Rows.Count - 1
                    For icol = 1 To pData.Columns.Count
                        .Cells(irow + rowstart, icol).Value = pData.Rows(irow)(icol - 1)
                        If icol = 7 Or icol = 8 Or icol = 14 Or icol = 15 Or icol = 16 Or icol = 20 Or icol = 23 Or icol = 26 Or icol = 29 Then
                            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "dd-mmm-yy"
                        End If
                        If icol = 11 Or icol = 13 Or icol = 18 Or icol = 19 Or icol = 21 Or icol = 28 Or icol = 30 Or icol = 25 Or icol = 34 Then
                            .Cells(irow + rowstart, icol).Style.Numberformat.Format = "#,##0"
                        End If
                    Next
                Next

                Dim rgAll As ExcelRange = .Cells(8, 1, irow + 8, 34)
                EpPlusDrawAllBorders(rgAll)

            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub
#End Region

#Region "Form Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If (Not IsPostBack) AndAlso (Not IsCallback) Then
                Call up_FillCombo()
                Call up_GridLoadWhenEventChange()
                Call up_Initialize()
            End If

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
        End Try

        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowPager)
    End Sub

    Private Sub btnSubMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubMenu.Click
        Session.Remove("G01Msg")
        Response.Redirect("~/MainMenu.aspx")
    End Sub

    Private Sub grid_CustomCallback(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewCustomCallbackEventArgs) Handles grid.CustomCallback
        Call clsAppearance.setAppearanceControlsDevEx13(Me.Page, clsAppearance.ShowHorizontalScrollMode.Visible, True, False, False, 6, False, clsAppearance.PagerMode.ShowPager)

        Try
            Dim pAction As String = Split(e.Parameters, "|")(0)
            Select Case pAction
                Case "load"
                    Call up_GridLoad()

                    If grid.VisibleRowCount = 0 Then
                        Call clsMsg.DisplayMessage(lblInfo, "2001", clsMessage.MsgType.InformationMessage)
                        Session("G01Msg") = lblInfo.Text
                    Else
                        grid.PageIndex = 0
                    End If
                Case "clear"
                    Call up_GridLoadWhenEventChange()

                Case "excel"
                    Dim psERR As String = ""
                    Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplateSummaryOutstanding.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:8", psERR)
                    End If
            End Select

        Catch ex As Exception
            Call clsMsg.DisplayMessage(lblInfo, Err.Number.ToString, clsMessage.MsgType.ErrorMessageFromVS, ex.Message.ToString())
            Session("G01Msg") = lblInfo.Text
        End Try

        If (Not IsNothing(Session("G01Msg"))) Then grid.JSProperties("cpMessage") = Session("G01Msg") : Session.Remove("G01Msg")

    End Sub

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As DevExpress.Web.ASPxGridView.ASPxGridViewTableDataCellEventArgs) Handles grid.HtmlDataCellPrepared
        e.Cell.Attributes.Add("onclick", "event.cancelBubble = true")
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

End Class