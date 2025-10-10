Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel


Public Class ForecastInquiryMonthly
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
            "   cboAffiliateCode.SetValue('" & Trim(Session("AffiliateID")) & "'); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "if (cboSupplierCode.GetItemCount() > 1) { " & vbCrLf & _
            "   cboSupplierCode.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            " " & vbCrLf & _
            "if (cboRevision.GetItemCount() > 1) { " & vbCrLf & _
            "   cboRevision.SetValue('==ALL=='); " & vbCrLf & _
            "} " & vbCrLf & _
            "var PeriodTo = new Date(); " & vbCrLf & _
            "dtPOPeriod.SetDate(PeriodTo); " & vbCrLf & _
            "lblInfo.SetText(''); "

        ScriptManager.RegisterStartupScript(dtPOPeriod, dtPOPeriod.GetType(), "Initialize", script, True)
    End Sub

    Private Sub up_GridLoad()
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            Dim ls_filter As String = ""
            ls_SQL = ""

            'grid.VisibleColumns(8).Caption = "Forecast Quantity " & Format(dtPOPeriod.Value, "MMM-yyyy")
            'grid.VisibleColumns(9).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 1, dtPOPeriod.Value), "MMM-yyyy")
            'grid.VisibleColumns(10).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 2, dtPOPeriod.Value), "MMM-yyyy")
            'grid.VisibleColumns(11).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 3, dtPOPeriod.Value), "MMM-yyyy")
            grid.VisibleColumns(8).Caption = "Forecast Quantity Jul-" & Year(dtPOPeriod.Value)
            grid.VisibleColumns(9).Caption = "Forecast Quantity Aug-" & Year(dtPOPeriod.Value)
            grid.VisibleColumns(10).Caption = "Forecast Quantity Sep-" & Year(dtPOPeriod.Value)
            grid.VisibleColumns(11).Caption = "Forecast Quantity Oct-" & Year(dtPOPeriod.Value)
            grid.VisibleColumns(12).Caption = "Forecast Quantity Nov-" & Year(dtPOPeriod.Value)
            grid.VisibleColumns(13).Caption = "Forecast Quantity Dec-" & Year(dtPOPeriod.Value)
            grid.VisibleColumns(14).Caption = "Forecast Quantity Jan-" & Year(dtPOPeriod.Value) + 1
            grid.VisibleColumns(15).Caption = "Forecast Quantity Feb-" & Year(dtPOPeriod.Value) + 1
            grid.VisibleColumns(16).Caption = "Forecast Quantity Mar-" & Year(dtPOPeriod.Value) + 1
            grid.VisibleColumns(17).Caption = "Forecast Quantity Apr-" & Year(dtPOPeriod.Value) + 1
            grid.VisibleColumns(18).Caption = "Forecast Quantity May-" & Year(dtPOPeriod.Value) + 1
            grid.VisibleColumns(19).Caption = "Forecast Quantity Jun-" & Year(dtPOPeriod.Value) + 1

            'AFFILIATE CODE
            If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND FM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            End If

            'SUPPLIER CODE
            If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND MPM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            End If

            'SUPPLIER CODE
            If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND FM.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
            End If

            'REVISION
            If Trim(cboRevision.Text) <> "==ALL==" And Trim(cboRevision.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND FM.Rev = '" & Trim(cboRevision.Text) & "' " & vbCrLf
            End If

            ls_SQL = "  Select row_number() over (order by FM.Year, FM.Rev, FM.AffiliateID, FM.PartNo asc) as no,FM.Year, FM.Rev, FM.AffiliateID, MPM.SupplierID, FM.PartNo, MP.PartName, MP.Project, MPQ = MPM.MOQ " & vbCrLf & _
                     "  ,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec " & vbCrLf & _
                     "  From ForecastMonthly FM " & vbCrLf & _
                     "  Left Join MS_PartMapping MPM ON FM.PartNo = MPM.PartNo And FM.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                     "  Left Join MS_Parts MP ON FM.PartNo = MP.PartNo " & vbCrLf & _
                     "  Where FM.Year = '" & Year(dtPOPeriod.Value) & "' " & vbCrLf & _
                     "   "

            ls_SQL = ls_SQL + ls_filter & vbCrLf

            ls_SQL = ls_SQL + " Order By FM.Year,FM.AffiliateID,FM.PartNo,FM.Rev " & vbCrLf & _
                              " " & vbCrLf

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

    'Private Function GetSummaryOutStanding_Old() As DataTable
    '    Dim ls_sql As String = ""
    '    Dim ls_filter As String = ""

    '    Try
    '        Dim clsGlobal As New clsGlobal
    '        Using cn As New SqlConnection(clsGlobal.ConnectionString)
    '            cn.Open()
    '            Dim sql As String = ""

    '            Dim ls_End As String = ""
    '            ls_End = Right("0" & Day(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Format(dtPOPeriodTo.Value, "yyyy-MM-01"))))), 2)

    '            'AFFILIATE CODE
    '            If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND POM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
    '            End If
    '            'AFFILIATE PO PERIOD
    '            If chkPOPeriod.Checked = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND CONVERT(date,POM.Period) BETWEEN '" & Format(dtPOPeriodFrom.Value, "yyyyMM01") & "' AND '" & Format(dtPOPeriodTo.Value, "yyyyMM" & ls_End) & "' " & vbCrLf
    '            End If
    '            'SUPPLIER PLAN DELIVERY DATE
    '            If chkSupplierPlanDelDate.Checked = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND CONVERT(date,KM.KanbanDate) BETWEEN '" & Format(dtSupplierPlanDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtSupplierPlanDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
    '            End If
    '            'SUPPLIER DELIVERY DATE
    '            If chkSupplierDelDate.Checked = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND CONVERT(date,SDM.DeliveryDate) BETWEEN '" & Format(dtSupplierDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtSupplierDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
    '            End If
    '            'PASI RECEIVE DATE
    '            If chkPASIRecDate.Checked = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND CONVERT(date,PRM.ReceiveDate) BETWEEN '" & Format(dtPASIRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
    '            End If
    '            'PASI DELIVERY DATE
    '            If chkPASIDelDate.Checked = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND CONVERT(date,PDM.DeliveryDate) BETWEEN '" & Format(dtPASIDelDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtPASIDelDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
    '            End If
    '            'AFFILIATE RECEIVE DATE
    '            If chkAffiliateRecDate.Checked = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND CONVERT(date,RAM.ReceiveDate) BETWEEN '" & Format(dtAffiliateRecDateFrom.Value, "yyyyMMdd") & "' AND '" & Format(dtAffiliateRecDateTo.Value, "yyyyMMdd") & "' " & vbCrLf
    '            End If

    '            'PART CODE
    '            If Trim(cboPart.Text) <> "==ALL==" And Trim(cboPart.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND POD.PartNo = '" & Trim(cboPart.Text) & "' " & vbCrLf
    '            End If
    '            'PONO
    '            If Trim(txtPONo.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(POM.PONo,'') LIKE '%" & Trim(txtPONo.Text) & "%' " & vbCrLf
    '            End If
    '            'SUPPLIER SJ NO
    '            If Trim(txtSupplierSJNo.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(SDD.SuratJalanNo,'') LIKE '%" & Trim(txtSupplierSJNo.Text) & "%'" & vbCrLf
    '            End If
    '            'PASI SJ NO
    '            If Trim(txtPASISJNo.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(PDD.SuratJalanNo,'') LIKE '%" & Trim(txtPASISJNo.Text) & "%'" & vbCrLf
    '            End If
    '            'SUPPLIER INV NO
    '            If Trim(txtSupplierInvNo.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(ISD.InvoiceNo,'') LIKE '%" & Trim(txtSupplierInvNo.Text) & "%'" & vbCrLf
    '            End If
    '            'PASI INV NO
    '            If Trim(txtPASIInvNo.Text) <> "" Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(IPD.InvoiceNo,'') LIKE '%" & Trim(txtPASIInvNo.Text) & "%'" & vbCrLf
    '            End If
    '            'PO PROGRESS

    '            'PASI RECEIVE
    '            If rdrPRComplete.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0) " & vbCrLf
    '            ElseIf rdrPRRemaining.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(SDD.DOQty,0) > ISNULL(PRD.GoodRecQty,0)  " & vbCrLf
    '            ElseIf rdrPRDiff.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND (ISNULL(SDD.DOQty,0) < ISNULL(PRD.GoodRecQty,0) OR ISNULL(PRD.DefectRecQty,0) > 0) " & vbCrLf
    '            End If

    '            'PASI DELIVERY
    '            If rdrPDComplete.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) = ISNULL(PDD.DOQty,0) " & vbCrLf & _
    '                              "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
    '            ElseIf rdrPDComplete.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((ISNULL(PRD.GoodRecQty,0) + ISNULL(PRD.DefectRecQty,0)) > (PDD.DOQty) " & vbCrLf & _
    '                              "                         AND ISNULL(SDD.DOQty,0) = ISNULL(PRD.GoodRecQty,0)) " & vbCrLf
    '                'ElseIf rdPRDDiff.Value = True Then
    '                '    ls_SQL = ls_SQL + _
    '                '                  "                      AND (ISNULL(PRD.DefectRecQty,0) > 0 " & vbCrLf & _
    '                '                  "                         AND ISNULL(PDD.DOQty,0) > 0) " & vbCrLf
    '            End If

    '            'AFFILIATE RECEIVE
    '            If rdrARComplete.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
    '                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
    '                              "                            END) = RAD.RecQty AND " & vbCrLf & _
    '                              "                           (CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.DefectQty,0) " & vbCrLf & _
    '                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.DefectRecQty,0) " & vbCrLf & _
    '                              "                            END) = 0 " & vbCrLf & _
    '                              "                          ) " & vbCrLf
    '            ElseIf rdrARRemaining.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
    '                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
    '                              "                            END) > RAD.RecQty " & vbCrLf & _
    '                              "                          ) " & vbCrLf
    '            ElseIf rdrARDiff.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(SDD.DOQty,0) " & vbCrLf & _
    '                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PDD.DOQty,0) " & vbCrLf & _
    '                              "                            END) < RAD.RecQty OR " & vbCrLf & _
    '                              "                           ISNULL(RAD.DefectQty,0) > 0  " & vbCrLf & _
    '                              "                          )" & vbCrLf
    '            End If

    '            'SUPPLIER INVOICE
    '            If rdrSIComplete.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
    '                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
    '                              "                            END) = ISD.InvQty " & vbCrLf & _
    '                              "                          ) " & vbCrLf
    '            ElseIf rdrSIRemaining.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ((CASE WHEN POM.DeliveryByPASICls = '0' THEN ISNULL(RAD.RecQty,0) " & vbCrLf & _
    '                              "                                 WHEN POM.DeliveryByPASICls = '1' THEN ISNULL(PRD.GoodRecQty,0) " & vbCrLf & _
    '                              "                            END) > ISD.InvQty " & vbCrLf & _
    '                              "                          ) " & vbCrLf
    '            End If

    '            'PASI INVOICE
    '            If rdrPIComplete.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(RAD.RecQty,0) = ISNULL(IPD.DOQty,0) " & vbCrLf
    '            ElseIf rdrPIRemaining.Value = True Then
    '                ls_filter = ls_filter + _
    '                              "                      AND ISNULL(RAD.RecQty,0) > ISNULL(IPD.DOQty,0) " & vbCrLf
    '            End If

    '            ls_sql = " SELECT DISTINCT * FROM " & vbCrLf & _
    '                  " ( " & vbCrLf & _
    '                  " 	SELECT  " & vbCrLf & _
    '                  " 		POM.Period " & vbCrLf & _
    '                  " 		,POM.PONo " & vbCrLf & _
    '                  " 		,POM.AffiliateID " & vbCrLf & _
    '                  " 		,POM.SupplierID " & vbCrLf & _
    '                  " 		,POKanban = CASE WHEN ISNULL(POD.KanbanCls, '0') = '1' THEN 'YES' ELSE 'NO' END " & vbCrLf & _
    '                  " 		,POM.EntryDate " & vbCrLf & _
    '                  " 		,POM.PASISendAffiliateDate " & vbCrLf & _
    '                  " 		,POD.PartNo " & vbCrLf

    '            ls_sql = ls_sql + " 		,MP.PartName " & vbCrLf & _
    '                              " 		,QtyPO = ISNULL(POD.POQty,0) " & vbCrLf & _
    '                              " 		,KD.KanbanNo " & vbCrLf & _
    '                              " 		,KD.KanbanQty " & vbCrLf & _
    '                              " 		,ETDSupp = ABC.ETDSupplier " & vbCrLf & _
    '                              " 		,ETAAff = KM.KanbanDate " & vbCrLf & _
    '                              " 		,SupplierDeliveryDate = SDM.DeliveryDate " & vbCrLf & _
    '                              " 		,SupplierSuratJalanNo = SDM.SuratJalanNo " & vbCrLf & _
    '                              " 		,SupplierDeliveryQty = SDD.DOQty " & vbCrLf & _
    '                              " 		,RemainingQtyPOPASI = ISNULL(KD.KanbanQty,0) - " & vbCrLf & _
    '                              " 		                      ISNULL( " & vbCrLf & _
    '                              " 		                        (select SUM(DOQty) from DOSupplier_Detail ABC " & vbCrLf & _
    '                              " 		                         WHERE ABC.SupplierID = SDD.SupplierID and ABC.AffiliateID = SDD.AffiliateID" & vbCrLf & _
    '                              " 		                         and ABC.KanbanNo = SDD.KanbanNo and ABC.PartNo = SDD.PartNo and ABC.PONo = SDD.PONo and ABC.SuratJalanNo = SDD.SuratJalanNo),0) " & vbCrLf & _
    '                              " 		,PASIReceiveDate = PRM.ReceiveDate " & vbCrLf

    '            ls_sql = ls_sql + " 		,PASIReceivingQty = PRD.GoodRecQty " & vbCrLf & _
    '                              " 		,InvoiceNoFromSupplier = ISM.InvoiceNo " & vbCrLf & _
    '                              " 		,InvoiceDateFromSupplier = ISM.InvoiceDate " & vbCrLf & _
    '                              " 		,InvoiceFromSupplierCurr = MCS.Description " & vbCrLf & _
    '                              " 		,InvoiceFromSupplierAmount = ISNULL(ISD.InvAmount,0) " & vbCrLf & _
    '                              " 		,PASIDeliveryDate = PDM.DeliveryDate " & vbCrLf & _
    '                              " 		,PASISuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
    '                              " 		,PASIDeliveryQty = PDD.DOQty " & vbCrLf & _
    '                              " 		,AffiliateReceiveDate = RAM.ReceiveDate " & vbCrLf & _
    '                              " 		,AffiliateReceivingQty = RAD.RecQty " & vbCrLf & _
    '                              " 		,InvoiceNoToAffiliate = IPM.InvoiceNo " & vbCrLf

    '            ls_sql = ls_sql + " 		,InvoiceDateToAffiliate = IPM.DeliveryDate " & vbCrLf & _
    '                              " 		,InvoiceToAffiliateCurr = 'IDR' " & vbCrLf & _
    '                              " 		,InvoiceToAffiliateAmount = ISNULL(IPD.DOQty,0) * ISNULL(MSP.Price,0) " & vbCrLf & _
    '                              " 	FROM PO_Master POM " & vbCrLf & _
    '                              " 	LEFT JOIN PO_Detail POD ON POM.AffiliateID = POD.AffiliateID " & vbCrLf & _
    '                              " 							AND POM.PoNo = POD.PONo " & vbCrLf & _
    '                              " 							AND POM.SupplierID = POD.SupplierID " & vbCrLf & _
    '                              " 	LEFT JOIN Kanban_Detail KD ON KD.AffiliateID = POD.AffiliateID " & vbCrLf & _
    '                              " 								AND KD.PoNo = POD.PONo " & vbCrLf & _
    '                              " 								AND KD.SupplierID = POD.SupplierID " & vbCrLf & _
    '                              " 								AND KD.PartNo = POD.PartNo " & vbCrLf

    '            ls_sql = ls_sql + " 	LEFT JOIN Kanban_Master KM ON KD.AffiliateID = KM.AffiliateID " & vbCrLf & _
    '                              " 								AND KD.KanbanNo = KM.KanbanNo " & vbCrLf & _
    '                              " 								AND KD.SupplierID = KM.SupplierID " & vbCrLf & _
    '                              " 								AND KD.DeliveryLocationCode = KM.DeliveryLocationCode " & vbCrLf & _
    '                              " 	LEFT JOIN DOSupplier_Detail SDD ON KD.AffiliateID = SDD.AffiliateID " & vbCrLf & _
    '                              " 									AND KD.KanbanNo = SDD.KanbanNo " & vbCrLf & _
    '                              " 									AND KD.SupplierID = SDD.SupplierID " & vbCrLf & _
    '                              " 									AND KD.PartNo = SDD.PartNo " & vbCrLf & _
    '                              " 									AND KD.PoNo = SDD.PoNo " & vbCrLf & _
    '                              " 	LEFT JOIN DOSupplier_Master SDM ON SDM.AffiliateID = SDD.AffiliateID " & vbCrLf & _
    '                              " 									AND SDM.SuratJalanNo = SDD.SuratJalanNo " & vbCrLf

    '            ls_sql = ls_sql + " 									AND SDM.SupplierID = SDD.SupplierID " & vbCrLf & _
    '                              " 	LEFT JOIN ReceivePASI_Detail PRD ON SDD.AffiliateID = PRD.AffiliateID " & vbCrLf & _
    '                              " 									AND SDD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
    '                              " 									AND SDD.SupplierID = PRD.SupplierID " & vbCrLf & _
    '                              " 									AND SDD.PartNo = PRD.PartNo " & vbCrLf & _
    '                              " 									AND SDD.PONo = PRD.PONo								 " & vbCrLf & _
    '                              " 									AND SDD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
    '                              " 	LEFT JOIN ReceivePASI_Master PRM ON PRM.AffiliateID = PRD.AffiliateID " & vbCrLf & _
    '                              " 									AND PRM.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
    '                              " 									AND PRM.SupplierID = PRD.SupplierID " & vbCrLf & _
    '                              " 	LEFT JOIN InvoiceSupplier_Detail ISD ON ISD.AffiliateID = PRD.AffiliateID " & vbCrLf

    '            ls_sql = ls_sql + " 										AND ISD.SupplierID = PRD.SupplierID " & vbCrLf & _
    '                              " 										AND ISD.SuratJalanNo = PRD.SuratJalanNo " & vbCrLf & _
    '                              " 										AND ISD.PONo = PRD.PONo " & vbCrLf & _
    '                              " 										AND ISD.PartNo = PRD.PartNo " & vbCrLf & _
    '                              " 										AND ISD.KanbanNo = PRD.KanbanNo " & vbCrLf & _
    '                              " 	LEFT JOIN InvoiceSupplier_Master ISM ON ISM.InvoiceNo = ISD.InvoiceNo " & vbCrLf & _
    '                              "   										AND ISM.AffiliateID = ISD.AffiliateID " & vbCrLf & _
    '                              "   										AND ISM.SupplierID = ISD.SupplierID " & vbCrLf & _
    '                              "   										AND ISM.suratJalanno = ISD.SuratJalanNo " & vbCrLf & _
    '                              " 	LEFT JOIN DOPASI_Detail PDD ON PRD.AffiliateID = PDD.AffiliateID " & vbCrLf & _
    '                              " 								AND PRD.KanbanNo = PDD.KanbanNo " & vbCrLf

    '            ls_sql = ls_sql + " 								AND PRD.SupplierID = PDD.SupplierID " & vbCrLf & _
    '                              " 								AND PRD.PartNo = PDD.PartNo " & vbCrLf & _
    '                              " 								AND PRD.PONo = PDD.PONo " & vbCrLf & _
    '                              " 								AND PRD.SuratJalanNo = PDD.SuratJalanNoSupplier " & vbCrLf & _
    '                              " 	LEFT JOIN DOPASI_Master PDM ON PDD.AffiliateID = PDM.AffiliateID " & vbCrLf & _
    '                              " 								AND PDD.SuratJalanNo = PDM.SuratJalanNo " & vbCrLf & _
    '                              " 	LEFT JOIN ReceiveAffiliate_Detail RAD ON PDD.AffiliateID = RAD.AffiliateID " & vbCrLf & _
    '                              " 									AND PDD.KanbanNo = RAD.KanbanNo " & vbCrLf & _
    '                              " 									AND PDD.SupplierID = RAD.SupplierID " & vbCrLf & _
    '                              " 									AND PDD.PartNo = RAD.PartNo " & vbCrLf & _
    '                              " 									AND PDD.PONo = RAD.PONo " & vbCrLf

    '            ls_sql = ls_sql + " 									AND PDD.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
    '                              " 	LEFT JOIN ReceiveAffiliate_Master RAM ON RAM.SuratJalanNo = RAD.SuratJalanNo " & vbCrLf & _
    '                              " 									AND RAM.AffiliateID = RAD.AffiliateID " & vbCrLf & _
    '                              " 	LEFT JOIN PLPASI_Detail IPD ON PDD.AffiliateID = IPD.AffiliateID   " & vbCrLf & _
    '                              " 									AND PDD.KanbanNo = IPD.KanbanNo								 " & vbCrLf & _
    '                              " 									AND PDD.PartNo = IPD.PartNo " & vbCrLf & _
    '                              " 									AND PDD.PONo = IPD.PONo " & vbCrLf & _
    '                              " 									AND PDD.SuratJalanNo = IPD.SuratJalanNo " & vbCrLf & _
    '                              " 	LEFT JOIN PLPASI_Master IPM ON IPD.AffiliateID = IPM.AffiliateID " & vbCrLf & _
    '                              " 									--AND IPD.InvoiceNo = IPM.InvoiceNo " & vbCrLf & _
    '                              " 									AND IPD.SuratJalanNo = IPM.SuratJalanNo " & vbCrLf

    '            ls_sql = ls_sql + " 	LEFT JOIN (  " & vbCrLf & _
    '                              "  				SELECT * FROM MS_ETD_PASI a  " & vbCrLf & _
    '                              "  				INNER JOIN MS_ETD_Supplier_PASI b on a.ETDPASI =  b.ETAPASI  " & vbCrLf & _
    '                              "  				)ABC ON POM.SupplierID = ABC.SupplierID and POM.AffiliateID = ABC.AffiliateID AND KM.KanbanDate = ABC.ETAAffiliate  " & vbCrLf & _
    '                              " 	LEFT JOIN MS_Parts MP ON MP.PartNo = POD.PartNo " & vbCrLf & _
    '                              " 	LEFT JOIN MS_CurrCls MCS ON MCS.CurrCls = ISD.InvCurrCls " & vbCrLf & _
    '                              " 	LEFT JOIN MS_Price MSP ON MSP.AffiliateID = IPD.AffiliateID and MSP.PartNo = IPD.PartNo and (IPM.DeliveryDate between MSP.StartDate and MSP.EndDate)  " & vbCrLf & _
    '                              " 	WHERE KD.KanbanQty > 0 " & vbCrLf

    '            ls_sql = ls_sql + ls_filter & vbCrLf

    '            ls_sql = ls_sql + " )XYZ " & vbCrLf & _
    '                              "  "

    '            Dim Cmd As New SqlCommand(ls_sql, cn)
    '            Dim da As New SqlDataAdapter(Cmd)
    '            Dim dt As New DataTable
    '            da.SelectCommand.CommandTimeout = 300
    '            da.Fill(dt)

    '            Return dt
    '        End Using
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function

    Private Function GetSummaryOutStanding() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                'AFFILIATE CODE
                If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND MPM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FM.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
                End If

                'REVISION
                If Trim(cboRevision.Text) <> "==ALL==" And Trim(cboRevision.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FM.Rev = '" & Trim(cboRevision.Text) & "' " & vbCrLf
                End If

                ls_sql = "  Select /*row_number() over (order by FM.Year, FM.Rev, FM.AffiliateID, FM.PartNo asc) as no,*/FM.Year, FM.Rev, FM.AffiliateID, MPM.SupplierID, FM.PartNo, MP.PartName, MP.Project, MPQ = MPM.MOQ " & vbCrLf & _
                         "  ,Jul,Aug,Sep,Oct,Nov,Dec,Jan,Feb,Mar,Apr,May,Jun " & vbCrLf & _
                         "  From ForecastMonthly FM " & vbCrLf & _
                         "  Left Join MS_PartMapping MPM ON FM.PartNo = MPM.PartNo And FM.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                         "  Left Join MS_Parts MP ON FM.PartNo = MP.PartNo " & vbCrLf & _
                         "  Where FM.Year = '" & Year(dtPOPeriod.Value) & "' " & vbCrLf & _
                         "  --AND FM.Rev = '" & Trim(cboRevision.Text) & "'  "



                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " Order By FM.Year,FM.AffiliateID,FM.PartNo,FM.Rev " & vbCrLf & _
                                  " " & vbCrLf


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

    Private Function GetSummaryOutStanding2() As DataTable
        Dim ls_sql As String = ""
        Dim ls_filter As String = ""

        Try
            Dim clsGlobal As New clsGlobal
            Using cn As New SqlConnection(clsGlobal.ConnectionString)
                cn.Open()
                Dim sql As String = ""

                'AFFILIATE CODE
                If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FM.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND MPM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FM.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
                End If

                'REVISION
                If Trim(cboRevision.Text) <> "==ALL==" And Trim(cboRevision.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FM.Rev = '" & Trim(cboRevision.Text) & "' " & vbCrLf
                End If

                ls_sql = "  Select /*row_number() over (order by FM.Year, FM.Rev, FM.AffiliateID, FM.PartNo asc) as no,*/FM.Year, FM.Rev, FM.AffiliateID, MPM.SupplierID, FM.PartNo, MP.PartName, MP.Project, MPQ = MPM.MOQ " & vbCrLf & _
                         "  ,Jul,Aug,Sep,Oct,Nov,Dec,Jan,Feb,Mar,Apr,May,Jun " & vbCrLf & _
                         "  --,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12 " & vbCrLf & _
                         "  From ForecastMonthly FM " & vbCrLf & _
                         "  Left Join MS_PartMapping MPM ON FM.PartNo = MPM.PartNo And FM.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                         "  Left Join MS_Parts MP ON FM.PartNo = MP.PartNo " & vbCrLf & _
                         "  Where FM.Year = '" & Year(dtPOPeriod.Value) & "' " & vbCrLf & _
                         "  "



                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " Order By FM.Year,FM.AffiliateID,FM.PartNo,FM.Rev " & vbCrLf & _
                                  " " & vbCrLf


                'Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
                'Dim ds As New DataSet
                'sqlDA.Fill(ds)
                'Return ds
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

            ls_SQL = " SELECT  Top 0 " & vbCrLf & _
                  "  	 Period = '' " & vbCrLf & _
                  "  	,PONo = '' " & vbCrLf & _
                  "  	,AffiliateID = '' " & vbCrLf & _
                  "  	,SupplierID = '' " & vbCrLf & _
                  "  	,POKanban = '' " & vbCrLf & _
                  "  	,PASISendAffiliateDate = '' " & vbCrLf & _
                  "  	,PartNo = '' " & vbCrLf & _
                  "  	,PartName = '' " & vbCrLf & _
                  "  	,QtyPO = '' " & vbCrLf & _
                  " 	,QtyBox = '' "

            ls_SQL = ls_SQL + " 	,BoxPallet = '' " & vbCrLf & _
                              " 	,VolumePallet = '' " & vbCrLf & _
                              "  	,ETDSupp = '' " & vbCrLf & _
                              "  	,ETAAff = '' " & vbCrLf & _
                              "  	,SupplierDeliveryDate = '' " & vbCrLf & _
                              "  	,SupplierSuratJalanNo = '' " & vbCrLf & _
                              "  	,SupplierDeliveryQty = '' " & vbCrLf & _
                              " 	,PASIReceiveDate = '' " & vbCrLf & _
                              "  	,PASIReceivingQty = '' " & vbCrLf & _
                              " 	,Remaining = '' " & vbCrLf & _
                              " 	,StatusPO = '' "


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
                     "SELECT AffiliateID = RTRIM(AffiliateID), AffiliateName = RTRIM(AffiliateName) FROM dbo.MS_Affiliate Where overseascls = '0'"
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

        'Combo Affiliate
        With cboSupplierCode
            ls_SQL = "--SELECT SupplierID = '==ALL==', SupplierName = '==ALL=='" & vbCrLf & _
                     " --UNION ALL " & vbCrLf & _
                     "SELECT SupplierID = RTRIM(SupplierID), SupplierName = RTRIM(SupplierName) FROM dbo.MS_supplier Where overseas = '0'"
            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("SupplierID")
                .Columns(0).Width = 90
                .Columns.Add("SupplierName")
                .Columns(1).Width = 240

                .TextField = "SupplierID"
                .DataBind()
            End Using
        End With

        'Combo Affiliate
        With cboRevision
            ls_SQL = " SELECT Rev = 0" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = 1" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = 2" & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     " SELECT Rev = 3" & vbCrLf

            Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
                sqlConn.Open()
                sqlDA = New SqlDataAdapter(ls_SQL, sqlConn)
                ds = New DataSet
                sqlDA.Fill(ds)

                .Items.Clear()
                .Columns.Clear()
                .DataSource = ds.Tables(0)
                .Columns.Add("Rev")
                .Columns(0).Width = 90
                '.Columns.Add("SupplierName")
                '.Columns(1).Width = 240

                .TextField = "Rev"
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
            Dim tempFile As String = "TemplateForecastReportMonthly " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
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
                .Cells(3, 4).Value = ": " & Format(dtPOPeriod.Value, "yyyy")
                .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text)
                .Cells(5, 4).Value = ": " & Trim(cboSupplierCode.Text)
                .Cells(6, 4).Value = ": " & Trim(cboRevision.Text)

                .Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(10, 1, pData.Rows.Count + 9, 20).AutoFitColumns()
                .Cells(10, 1, pData.Rows.Count + 9, 20).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

                '.Cells(10, 1, pData.Rows.Count + 9, 1).Style.Numberformat.Format = "yyyy"
                .Cells(8, 9).Value = "Forecast Quantity Jul-" & Year(dtPOPeriod.Value)
                .Cells(8, 10).Value = "Forecast Quantity Aug-" & Year(dtPOPeriod.Value)
                .Cells(8, 11).Value = "Forecast Quantity Sep-" & Year(dtPOPeriod.Value)
                .Cells(8, 12).Value = "Forecast Quantity Oct-" & Year(dtPOPeriod.Value)
                .Cells(8, 13).Value = "Forecast Quantity Nov-" & Year(dtPOPeriod.Value)
                .Cells(8, 14).Value = "Forecast Quantity Dec-" & Year(dtPOPeriod.Value)
                .Cells(8, 15).Value = "Forecast Quantity Jan-" & Year(dtPOPeriod.Value) + 1
                .Cells(8, 16).Value = "Forecast Quantity Feb-" & Year(dtPOPeriod.Value) + 1
                .Cells(8, 17).Value = "Forecast Quantity Mar-" & Year(dtPOPeriod.Value) + 1
                .Cells(8, 18).Value = "Forecast Quantity Apr-" & Year(dtPOPeriod.Value) + 1
                .Cells(8, 19).Value = "Forecast Quantity May-" & Year(dtPOPeriod.Value) + 1
                .Cells(8, 20).Value = "Forecast Quantity Jun-" & Year(dtPOPeriod.Value) + 1

                For x = 9 To 20
                    .Cells(10, x, pData.Rows.Count + 9, x).Style.Numberformat.Format = "#,##0"
                Next


                Dim rgAll As ExcelRange = .Cells(10, 1, pData.Rows.Count + 9, 20)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    'Private Sub epplusExportExcelNew(ByVal pFilename As String, ByVal pSheetName As String,
    '                          ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

    '    Try
    '        Dim tempFile As String = "TemplateForecastReportMonthly " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
    '        Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
    '        If (System.IO.File.Exists(pFilename)) Then
    '            System.IO.File.Copy(pFilename, NewFileName, True)
    '        End If

    '        Dim rowstart As String = Split(pCellStart, ":")(1)
    '        Dim Coltart As String = Split(pCellStart, ":")(0)
    '        'Dim fi As New FileInfo(NewFileName)

    '        'Dim exl As New ExcelPackage(fi)
    '        'Dim ws As ExcelWorksheet
    '        Dim ExcelBook As excel.Workbook
    '        Dim ExcelSheet As excel.Worksheet

    '        Dim xlApp = New excel.Application
    '        Dim ls_file As String = NewFileName
    '        '
    '        ExcelBook = xlApp.Workbooks.Open(ls_file)
    '        ExcelSheet = CType(ExcelBook.Worksheets(pSheetName), excel.Worksheet)

    '        'ws = exl.Workbook.Worksheets(pSheetName)
    '        With ExcelSheet
    '            .Cells(3, 4).Value = ": " & Year(dtPOPeriod.Value)
    '            .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text)
    '            .Cells(5, 4).Value = ": " & Trim(cboSupplierCode.Text)
    '            .Cells(6, 4).Value = ": " & Trim(cboRevision.Text)

    '            ExcelSheet.Range("I8").Value = "Forecast Quantity Jul-" & Year(dtPOPeriod.Value)
    '            ExcelSheet.Range("J8").Value = "Forecast Quantity Aug-" & Year(dtPOPeriod.Value)
    '            ExcelSheet.Range("K8").Value = "Forecast Quantity Sep-" & Year(dtPOPeriod.Value)
    '            ExcelSheet.Range("L8").Value = "Forecast Quantity Oct-" & Year(dtPOPeriod.Value)
    '            ExcelSheet.Range("M8").Value = "Forecast Quantity Nov-" & Year(dtPOPeriod.Value)
    '            ExcelSheet.Range("N8").Value = "Forecast Quantity Dec-" & Year(dtPOPeriod.Value)
    '            ExcelSheet.Range("O8").Value = "Forecast Quantity Jan-" & Year(dtPOPeriod.Value) + 1
    '            ExcelSheet.Range("P8").Value = "Forecast Quantity Feb-" & Year(dtPOPeriod.Value) + 1
    '            ExcelSheet.Range("Q8").Value = "Forecast Quantity Mar-" & Year(dtPOPeriod.Value) + 1
    '            ExcelSheet.Range("R8").Value = "Forecast Quantity Apr-" & Year(dtPOPeriod.Value) + 1
    '            ExcelSheet.Range("S8").Value = "Forecast Quantity May-" & Year(dtPOPeriod.Value) + 1
    '            ExcelSheet.Range("T8").Value = "Forecast Quantity Jun-" & Year(dtPOPeriod.Value) + 1

    '            '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
    '            Dim ds As New DataSet
    '            ds = GetSummaryOutStanding2()
    '            If ds.Tables(0).Rows.Count > 0 Then
    '                For i = 0 To ds.Tables(0).Rows.Count - 1
    '                    ExcelSheet.Range("A" & i + 10).Value = ds.Tables(0).Rows(i)("Year")
    '                    ExcelSheet.Range("B" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Rev"))
    '                    ExcelSheet.Range("C" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
    '                    ExcelSheet.Range("D" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
    '                    ExcelSheet.Range("E" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))
    '                    ExcelSheet.Range("F" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartName"))
    '                    ExcelSheet.Range("G" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Project"))
    '                    ExcelSheet.Range("H" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("MPQ"))
    '                    ExcelSheet.Range("I" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Jan"))
    '                    ExcelSheet.Range("J" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Feb"))
    '                    ExcelSheet.Range("K" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Mar"))
    '                    ExcelSheet.Range("L" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Apr"))
    '                    ExcelSheet.Range("M" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("May"))
    '                    ExcelSheet.Range("N" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Jun"))
    '                    ExcelSheet.Range("O" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Jul"))
    '                    ExcelSheet.Range("P" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Aug"))
    '                    ExcelSheet.Range("Q" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Sep"))
    '                    ExcelSheet.Range("R" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Oct"))
    '                    ExcelSheet.Range("S" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Nov"))
    '                    ExcelSheet.Range("T" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Dec"))

    '                    ExcelSheet.Range("H" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("I" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("J" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("K" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("L" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("M" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("N" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("O" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("P" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("Q" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("R" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("S" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("T" & i + 10).NumberFormat = "#,##0"

    '                    'Looping Column
    '                    For z = 0 To 11
    '                        'Cek Cls
    '                        If Trim(ds.Tables(0).Rows(i)("C" & z + 1)) = 1 Then
    '                            'Cek rev
    '                            If .Cells(10 + i, 2).Value = 1 Then
    '                                'ExcelSheet.Range(10 + i, 13 + z).Interior.Color = Color.Yellow
    '                                ExcelSheet.Range(uf_NoChar(8 + z) & 10 + i).Interior.Color = Color.Yellow
    '                            ElseIf .Cells(10 + i, 2).Value = 2 Then
    '                                ExcelSheet.Range(uf_NoChar(8 + z) & 10 + i).Interior.Color = Color.Orange
    '                            ElseIf .Cells(10 + i, 2).Value = 2 Then
    '                                ExcelSheet.Range(uf_NoChar(8 + z) & 10 + i).Interior.Color = Color.Green
    '                            End If
    '                        End If
    '                    Next
    '                Next
    '            End If

    '            ExcelSheet.Range("M10: T" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    '            ExcelSheet.Range("A10: T" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: T" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: T" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: T" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: T" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: T" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

    '        End With

    '        'exl.Save()
    '        ExcelBook.Save()

    '        DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

    '        'exl = Nothing
    '        xlApp.Workbooks.Close()
    '        xlApp.Quit()
    '        xlApp = Nothing
    '    Catch ex As Exception
    '        pErr = ex.Message
    '    End Try

    'End Sub

    Private Function uf_NoChar(ByVal iNo As Integer)
        Dim ls_char As String = ""

        If iNo <= 25 Then
            ls_char = Chr(65 + iNo)
        Else
            ls_char = Chr(65 + (iNo - 26))
        End If

        uf_NoChar = ls_char
    End Function

    Private Function uf_ColorCls(ByVal pYear As Integer, ByVal pAffiliate As String, ByVal pRev As Integer, ByVal pPartNo As String, ByVal pTgl As Integer) As Boolean
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = "Select C" & pTgl & " From ForecastMonthly Where Year = '" & pYear & "' And AffiliateID = '" & pAffiliate & "' And Rev = '" & pRev & "' And PartNo = '" & pPartNo & "' "
            Dim sqlCmd As New SqlCommand(ls_SQL, sqlConn)
            Dim sqlDA As New SqlDataAdapter(sqlCmd)
            Dim ds As New DataSet
            sqlDA.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("C" & pTgl & "") = "1" Then
                    Return True
                End If
            Else
                Return False
            End If
        End Using
    End Function

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
                    Dim dtProd As DataTable = GetSummaryOutStanding2()
                    FileName = "TemplateForecastReportMonthly.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:10", psERR)
                        'Call epplusExportExcelNew(FilePath, "Sheet1", dtProd, "A:10", psERR)
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
        With e.DataColumn
            If .FieldName = "Jan" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 1) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Feb" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 2) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Mar" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 3) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Apr" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 4) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "May" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 5) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Jun" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 6) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Jul" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 7) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Aug" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 8) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Sep" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 9) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Oct" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 10) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Nov" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 11) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "Dec" Then
                If uf_ColorCls(e.GetValue("Year"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 12) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub grid_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.PageIndexChanged
        Call up_GridLoad()
    End Sub
#End Region

End Class