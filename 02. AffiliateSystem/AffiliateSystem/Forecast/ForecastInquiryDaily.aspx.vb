Imports System.Data
Imports System.Data.SqlClient
Imports DevExpress
Imports DevExpress.Web.ASPxGridView
Imports System.Transactions
Imports System.Drawing

Imports OfficeOpenXml
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel


Public Class ForecastInquiryDaily
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
            " " & vbCrLf & _
            " " & vbCrLf & _
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

            grid.VisibleColumns(8).Caption = "Forecast Quantity " & Format(dtPOPeriod.Value, "MMM-yyyy")
            grid.VisibleColumns(9).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 1, dtPOPeriod.Value), "MMM-yyyy")
            grid.VisibleColumns(10).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 2, dtPOPeriod.Value), "MMM-yyyy")
            grid.VisibleColumns(11).Caption = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 3, dtPOPeriod.Value), "MMM-yyyy")

            'AFFILIATE CODE
            If Trim(cboAffiliateCode.Text) <> "==ALL==" And Trim(cboAffiliateCode.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND FD.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
            End If

            'SUPPLIER CODE
            If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND MPM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
            End If

            'SUPPLIER CODE
            If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND FD.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
            End If

            'REVISION
            If Trim(cboRevision.Text) <> "==ALL==" And Trim(cboRevision.Text) <> "" Then
                ls_filter = ls_filter + _
                              "                      AND FD.Rev = '" & Trim(cboRevision.Text) & "' " & vbCrLf
            End If

            ls_SQL = " Select row_number() over (order by FD.Period, FD.Rev, FD.AffiliateID, FD.PartNo asc) as no,FD.Period, FD.Rev, FD.AffiliateID, MPM.SupplierID, FD.PartNo, MP.PartName, MP.Project, MPQ = MPM.MOQ, FD.ForecastQty1, FD.ForecastQty2, FD.ForecastQty3, FD.ForecastQty4 " & vbCrLf & _
                  " ,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf & _
                  " From ForecastDaily FD " & vbCrLf & _
                  " Left Join MS_PartMapping MPM ON FD.PartNo = MPM.PartNo And FD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                  " Left Join MS_Parts MP ON FD.PartNo = MP.PartNo " & vbCrLf & _
                  " Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf & _
                  "  "


            ls_SQL = ls_SQL + ls_filter & vbCrLf

            ls_SQL = ls_SQL + " Order By FD.Period,FD.AffiliateID,FD.PartNo,FD.Rev " & vbCrLf & _
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

    Private Function uf_ColorCls(ByVal pPeriod As Date, ByVal pAffiliate As String, ByVal pRev As Integer, ByVal pPartNo As String, ByVal pTgl As Integer) As Boolean
        Using sqlConn As New SqlConnection(clsGlobal.ConnectionString)
            sqlConn.Open()
            ls_SQL = ""

            ls_SQL = "Select C" & pTgl & " From ForecastDaily Where Period = '" & pPeriod & "' And AffiliateID = '" & pAffiliate & "' And Rev = '" & pRev & "' And PartNo = '" & pPartNo & "' "
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
                                  "                      AND FD.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND MPM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FD.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
                End If

                'REVISION
                If Trim(cboRevision.Text) <> "==ALL==" And Trim(cboRevision.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FD.Rev = '" & Trim(cboRevision.Text) & "' " & vbCrLf
                End If

                ls_sql = " Select FD.Period, FD.Rev, FD.AffiliateID, MPM.SupplierID, FD.PartNo, MP.PartName, MP.Project, MPQ = MPM.MOQ, FD.ForecastQty1, FD.ForecastQty2, FD.ForecastQty3, FD.ForecastQty4 " & vbCrLf & _
                      " ,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf & _
                      " ,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,C17,C18,C19,C20,C21,C22,C23,C24,C25,C26,C27,C28,C29,C30,C31 " & vbCrLf & _
                      " From ForecastDaily FD " & vbCrLf & _
                      " Left Join MS_PartMapping MPM ON FD.PartNo = MPM.PartNo And FD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                      " Left Join MS_Parts MP ON FD.PartNo = MP.PartNo " & vbCrLf & _
                      " Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf & _
                      "  "


                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " Order By FD.Period,FD.AffiliateID,FD.PartNo,FD.Rev " & vbCrLf & _
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

    Private Function GetSummaryOutStanding2() As DataSet
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
                                  "                      AND FD.AffiliateID = '" & Trim(cboAffiliateCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(cboSupplierCode.Text) <> "==ALL==" And Trim(cboSupplierCode.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND MPM.SupplierID = '" & Trim(cboSupplierCode.Text) & "' " & vbCrLf
                End If

                'SUPPLIER CODE
                If Trim(txtPartNo.Text) <> "==ALL==" And Trim(txtPartNo.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FD.PartNo = '" & Trim(txtPartNo.Text) & "' " & vbCrLf
                End If

                'REVISION
                If Trim(cboRevision.Text) <> "==ALL==" And Trim(cboRevision.Text) <> "" Then
                    ls_filter = ls_filter + _
                                  "                      AND FD.Rev = '" & Trim(cboRevision.Text) & "' " & vbCrLf
                End If

                ls_sql = " Select Period=ISNULL(FD.Period,''), Rev=ISNULL(FD.Rev,''), AffiliateID=ISNULL(FD.AffiliateID,''), SupplierID=ISNULL(MPM.SupplierID,''), FD.PartNo, PartName=ISNULL(MP.PartName,''), Project=ISNULL(MP.Project,'') , MPQ = ISNULL(MPM.MOQ,0), " & vbCrLf & _
                      " ForecastQty1 = ISNULL(FD.ForecastQty1,0), ForecastQty2= ISNULL(FD.ForecastQty2,0), ForecastQty3=ISNULL(FD.ForecastQty3,0), ForecastQty4=ISNULL(FD.ForecastQty4,0) " & vbCrLf & _
                      " ,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22,F23,F24,F25,F26,F27,F28,F29,F30,F31 " & vbCrLf & _
                      " ,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,C17,C18,C19,C20,C21,C22,C23,C24,C25,C26,C27,C28,C29,C30,C31 " & vbCrLf & _
                      " From ForecastDaily FD " & vbCrLf & _
                      " Left Join MS_PartMapping MPM ON FD.PartNo = MPM.PartNo And FD.AffiliateID = MPM.AffiliateID " & vbCrLf & _
                      " Left Join MS_Parts MP ON FD.PartNo = MP.PartNo " & vbCrLf & _
                      " Where Period = '" & Format(dtPOPeriod.Value, "yyyy-MM") & "-01' " & vbCrLf & _
                      "  "


                ls_sql = ls_sql + ls_filter & vbCrLf

                ls_sql = ls_sql + " Order By FD.Period,FD.AffiliateID,FD.PartNo,FD.Rev " & vbCrLf & _
                                  " " & vbCrLf


                Dim sqlDA As New SqlDataAdapter(ls_sql, cn)
                Dim ds As New DataSet
                sqlDA.Fill(ds)
                Return ds
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function uf_NoChar(ByVal iNo As Integer)
        Dim ls_char As String = ""

        If iNo <= 25 Then
            ls_char = Chr(65 + iNo)
        Else
            ls_char = Chr(65 + (iNo - 26))
        End If

        uf_NoChar = ls_char
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
            Dim tempFile As String = "TemplateSummaryOutstandingDeliverySupplier " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
                '.Cells(3, 4).Value = ": " & Format(dtPOPeriodFrom.Value, "MMM yyyy") & " - " & Format(dtPOPeriodTo.Value, "MMM yyyy")
                '.Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text) & " / " & Trim(txtAffiliateName.Text)

                .Cells("A8").LoadFromDataTable(DirectCast(pData, DataTable), False)
                .Cells(8, 1, pData.Rows.Count + 7, 34).AutoFitColumns()
                .Cells(8, 1, pData.Rows.Count + 7, 34).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center

                .Cells(8, 1, pData.Rows.Count + 7, 1).Style.Numberformat.Format = "mmm-yy"
                .Cells(8, 6, pData.Rows.Count + 7, 6).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 13, pData.Rows.Count + 7, 13).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 14, pData.Rows.Count + 7, 14).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 15, pData.Rows.Count + 7, 15).Style.Numberformat.Format = "dd-mmm-yy"
                .Cells(8, 18, pData.Rows.Count + 7, 18).Style.Numberformat.Format = "dd-mmm-yy"

                .Cells(8, 9, pData.Rows.Count + 7, 9).Style.Numberformat.Format = "#,##0"
                .Cells(8, 10, pData.Rows.Count + 7, 10).Style.Numberformat.Format = "#,##0"
                .Cells(8, 11, pData.Rows.Count + 7, 11).Style.Numberformat.Format = "#,##0"
                .Cells(8, 12, pData.Rows.Count + 7, 12).Style.Numberformat.Format = "#,##0"
                .Cells(8, 17, pData.Rows.Count + 7, 17).Style.Numberformat.Format = "#,##0"
                .Cells(8, 19, pData.Rows.Count + 7, 19).Style.Numberformat.Format = "#,##0"
                .Cells(8, 20, pData.Rows.Count + 7, 20).Style.Numberformat.Format = "#,##0"


                Dim rgAll As ExcelRange = .Cells(8, 1, pData.Rows.Count + 7, 21)
                EpPlusDrawAllBorders(rgAll)
            End With

            exl.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\ProgressReport\Import\" & tempFile & "")

            exl = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    Private Sub epplusExportExcelNew(ByVal pFilename As String, ByVal pSheetName As String,
                                ByVal pCellStart As String, Optional ByRef pErr As String = "")
        'ByVal pData As DataTable
        Try
            Dim tempFile As String = "TemplateForecastReportDaily " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
            Dim NewFileName As String = Server.MapPath("~\Forecast\Import\" & tempFile & "")
            If (System.IO.File.Exists(pFilename)) Then
                System.IO.File.Copy(pFilename, NewFileName, True)
            End If

            Dim rowstart As String = Split(pCellStart, ":")(1)
            Dim Coltart As String = Split(pCellStart, ":")(0)
            Dim fi As New FileInfo(NewFileName)

            Dim exl As New ExcelPackage(fi)
            Dim ws As ExcelWorksheet
            'Dim ExcelBook As excel.Workbook
            'Dim ExcelSheet As excel.Worksheet

            'Dim xlApp = New excel.Application
            'Dim ls_file As String = NewFileName
            ''
            'ExcelBook = xlApp.Workbooks.Open(ls_file)
            'ExcelSheet = CType(ExcelBook.Worksheets(pSheetName), excel.Worksheet)

            ws = exl.Workbook.Worksheets(pSheetName)
            'With ExcelSheet
            With ws
                .Cells(3, 4).Value = ": " & Format(dtPOPeriod.Value, "MMM yyyy")
                .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text)
                .Cells(5, 4).Value = ": " & Trim(cboSupplierCode.Text)
                .Cells(6, 4).Value = ": " & Trim(cboRevision.Text)

                .Cells(8, 9).Value = "Forecast Quantity " & Format(dtPOPeriod.Value, "MMM yyyy")
                .Cells(8, 10).Value = Format(DateAdd(DateInterval.Month, 1, dtPOPeriod.Value), "MMM yyyy")
                .Cells(8, 11).Value = Format(DateAdd(DateInterval.Month, 2, dtPOPeriod.Value), "MMM yyyy")
                .Cells(8, 12).Value = Format(DateAdd(DateInterval.Month, 3, dtPOPeriod.Value), "MMM yyyy")
                'ExcelSheet.Range("I8").Value = "Forecast Quantity " & Format(dtPOPeriod.Value, "MMM yyyy")
                'ExcelSheet.Range("J8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 1, dtPOPeriod.Value), "MMM yyyy")
                'ExcelSheet.Range("K8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 2, dtPOPeriod.Value), "MMM yyyy")
                'ExcelSheet.Range("L8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 3, dtPOPeriod.Value), "MMM yyyy")

                '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
                Dim ds As New DataSet
                ds = GetSummaryOutStanding2()
                If ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        .Cells("A" & i + 10).Value = ds.Tables(0).Rows(i)("Period")
                        .Cells("A" & i + 10).Style.Numberformat.Format = "MMM-yy"
                        .Cells("B" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Rev"))
                        .Cells("C" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
                        .Cells("D" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
                        .Cells("E" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))
                        .Cells("F" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartName"))
                        .Cells("G" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Project"))
                        .Cells("H" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("MPQ"))
                        .Cells("I" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty1"))
                        .Cells("J" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty2"))
                        .Cells("K" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty3"))
                        .Cells("L" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty4"))
                        .Cells("M" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F1"))
                        .Cells("N" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F2"))
                        .Cells("O" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F3"))
                        .Cells("P" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F4"))
                        .Cells("Q" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F5"))
                        .Cells("R" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F6"))
                        .Cells("S" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F7"))
                        .Cells("T" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F8"))
                        .Cells("U" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F9"))
                        .Cells("V" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F10"))
                        .Cells("W" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F11"))
                        .Cells("X" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F12"))
                        .Cells("Y" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F13"))
                        .Cells("Z" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F14"))
                        .Cells("AA" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F15"))
                        .Cells("AB" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F16"))
                        .Cells("AC" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F17"))
                        .Cells("AD" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F18"))
                        .Cells("AE" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F19"))
                        .Cells("AF" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F20"))
                        .Cells("AG" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F21"))
                        .Cells("AH" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F22"))
                        .Cells("AI" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F23"))
                        .Cells("AJ" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F24"))
                        .Cells("AK" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F25"))
                        .Cells("AL" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F26"))
                        .Cells("AM" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F27"))
                        .Cells("AN" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F28"))
                        .Cells("AO" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F29"))
                        .Cells("AP" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F30"))
                        .Cells("AQ" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F31"))

                        .Cells("H" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("I" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("J" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("K" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("L" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("M" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("N" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("O" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("P" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("Q" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("R" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("S" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("T" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("U" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("V" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("W" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("X" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("Y" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("Z" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AA" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AB" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AC" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AD" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AE" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AF" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AG" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AH" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AI" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AJ" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AK" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AL" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AM" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AN" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AO" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AP" & i + 10).Style.Numberformat.Format = "#,##0"
                        .Cells("AQ" & i + 10).Style.Numberformat.Format = "#,##0"

                        ''Looping Column
                        'For z = 0 To 30
                        '    'Cek Cls
                        '    If Trim(ds.Tables(0).Rows(i)("C" & z + 1)) = 1 Then
                        '        'Cek rev
                        '        If .Cells(10 + i, 2).Value = 1 Then
                        '            'ExcelSheet.Range(10 + i, 13 + z).Interior.Color = Color.Yellow
                        '            .Cells(10 + i, 12 + z).Style.Fill.BackgroundColor.SetColor(Color.Yellow)
                        '        ElseIf .Cells(10 + i, 2).Value = 2 Then
                        '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Orange
                        '        ElseIf .Cells(10 + i, 2).Value = 2 Then
                        '            ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Green
                        '        End If
                        '    End If
                        'Next
                    Next
                End If

                Dim rgAll As ExcelRange = .Cells(10, 1, ds.Tables(0).Rows.Count + 9, 43)
                EpPlusDrawAllBorders(rgAll)

                '.Cells(10, 1, pData.Rows.Count + 9, 43).AutoFitColumns()
                '.Range("A10:AQ" & pData.Rows.Count).AutoFit()

                'ExcelSheet.Range("M10: AQ" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

                'ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                'ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                'ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                'ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                'ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                'ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            End With

            exl.Save()
            'ExcelBook.Save()

            DevExpress.Web.ASPxClasses.ASPxWebControl.RedirectOnCallback("~\Forecast\Import\" & tempFile & "")

            exl = Nothing
            'xlApp.Workbooks.Close()
            'xlApp.Quit()
            'xlApp = Nothing
        Catch ex As Exception
            pErr = ex.Message
        End Try

    End Sub

    'Private Sub epplusExportExcelNew(ByVal pFilename As String, ByVal pSheetName As String,
    '                          ByVal pData As DataTable, ByVal pCellStart As String, Optional ByRef pErr As String = "")

    '    Try
    '        Dim tempFile As String = "TemplateForecastReportDaily " & Format(Now, "yyyyMMdd hhmmss") & ".xlsx"
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
    '            .Cells(3, 4).Value = ": " & Format(dtPOPeriod.Value, "MMM yyyy")
    '            .Cells(4, 4).Value = ": " & Trim(cboAffiliateCode.Text)
    '            .Cells(5, 4).Value = ": " & Trim(cboSupplierCode.Text)
    '            .Cells(6, 4).Value = ": " & Trim(cboRevision.Text)

    '            ExcelSheet.Range("I8").Value = "Forecast Quantity " & Format(dtPOPeriod.Value, "MMM yyyy")
    '            ExcelSheet.Range("J8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 1, dtPOPeriod.Value), "MMM yyyy")
    '            ExcelSheet.Range("K8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 2, dtPOPeriod.Value), "MMM yyyy")
    '            ExcelSheet.Range("L8").Value = "Forecast Quantity " & Format(DateAdd(DateInterval.Month, 3, dtPOPeriod.Value), "MMM yyyy")

    '            '.Cells("A10").LoadFromDataTable(DirectCast(pData, DataTable), False)
    '            Dim ds As New DataSet
    '            ds = GetSummaryOutStanding2()
    '            If ds.Tables(0).Rows.Count > 0 Then
    '                For i = 0 To ds.Tables(0).Rows.Count - 1
    '                    ExcelSheet.Range("A" & i + 10).Value = ds.Tables(0).Rows(i)("Period")
    '                    ExcelSheet.Range("A" & i + 10 & ": A" & i + 10).NumberFormat = "MMM-yy"
    '                    ExcelSheet.Range("B" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Rev"))
    '                    ExcelSheet.Range("C" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("AffiliateID"))
    '                    ExcelSheet.Range("D" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("SupplierID"))
    '                    ExcelSheet.Range("E" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartNo"))
    '                    ExcelSheet.Range("F" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("PartName"))
    '                    ExcelSheet.Range("G" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("Project"))
    '                    ExcelSheet.Range("H" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("MPQ"))
    '                    ExcelSheet.Range("I" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty1"))
    '                    ExcelSheet.Range("J" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty2"))
    '                    ExcelSheet.Range("K" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty3"))
    '                    ExcelSheet.Range("L" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("ForecastQty4"))
    '                    ExcelSheet.Range("M" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F1"))
    '                    ExcelSheet.Range("N" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F2"))
    '                    ExcelSheet.Range("O" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F3"))
    '                    ExcelSheet.Range("P" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F4"))
    '                    ExcelSheet.Range("Q" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F5"))
    '                    ExcelSheet.Range("R" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F6"))
    '                    ExcelSheet.Range("S" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F7"))
    '                    ExcelSheet.Range("T" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F8"))
    '                    ExcelSheet.Range("U" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F9"))
    '                    ExcelSheet.Range("V" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F10"))
    '                    ExcelSheet.Range("W" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F11"))
    '                    ExcelSheet.Range("X" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F12"))
    '                    ExcelSheet.Range("Y" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F13"))
    '                    ExcelSheet.Range("Z" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F14"))
    '                    ExcelSheet.Range("AA" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F15"))
    '                    ExcelSheet.Range("AB" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F16"))
    '                    ExcelSheet.Range("AC" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F17"))
    '                    ExcelSheet.Range("AD" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F18"))
    '                    ExcelSheet.Range("AE" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F19"))
    '                    ExcelSheet.Range("AF" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F20"))
    '                    ExcelSheet.Range("AG" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F21"))
    '                    ExcelSheet.Range("AH" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F22"))
    '                    ExcelSheet.Range("AI" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F23"))
    '                    ExcelSheet.Range("AJ" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F24"))
    '                    ExcelSheet.Range("AK" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F25"))
    '                    ExcelSheet.Range("AL" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F26"))
    '                    ExcelSheet.Range("AM" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F27"))
    '                    ExcelSheet.Range("AN" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F28"))
    '                    ExcelSheet.Range("AO" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F29"))
    '                    ExcelSheet.Range("AP" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F30"))
    '                    ExcelSheet.Range("AQ" & i + 10).Value = Trim(ds.Tables(0).Rows(i)("F31"))

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
    '                    ExcelSheet.Range("U" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("V" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("W" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("X" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("Y" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("Z" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AA" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AB" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AC" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AD" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AE" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AF" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AG" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AH" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AI" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AJ" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AK" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AL" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AM" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AN" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AO" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AP" & i + 10).NumberFormat = "#,##0"
    '                    ExcelSheet.Range("AQ" & i + 10).NumberFormat = "#,##0"

    '                    'Looping Column
    '                    For z = 0 To 30
    '                        'Cek Cls
    '                        If Trim(ds.Tables(0).Rows(i)("C" & z + 1)) = 1 Then
    '                            'Cek rev
    '                            If .Cells(10 + i, 2).Value = 1 Then
    '                                'ExcelSheet.Range(10 + i, 13 + z).Interior.Color = Color.Yellow
    '                                ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Yellow
    '                            ElseIf .Cells(10 + i, 2).Value = 2 Then
    '                                ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Orange
    '                            ElseIf .Cells(10 + i, 2).Value = 2 Then
    '                                ExcelSheet.Range(uf_NoChar(12 + z) & 10 + i).Interior.Color = Color.Green
    '                            End If
    '                        End If
    '                    Next
    '                Next
    '            End If

    '            '.Cells(10, 1, pData.Rows.Count + 9, 43).AutoFitColumns()
    '            '.Range("A10:AQ" & pData.Rows.Count).AutoFit()

    '            ExcelSheet.Range("M10: AQ" & ds.Tables(0).Rows.Count + 8).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    '            ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '            ExcelSheet.Range("A10: AQ" & ds.Tables(0).Rows.Count + 9).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

    '            ''Dim exl2 As excel.Application
    '            ''exl2.Workbooks.Add()
    '            ''Looping Row
    '            'For i = 0 To pData.Rows.Count
    '            '    'Looping Column
    '            '    For z = 0 To 30
    '            '        'Cek Cls
    '            '        If .Cells(10 + i, 44 + z).Value = 1 Then
    '            '            'Cek rev
    '            '            If .Cells(10 + i, 2).Value = 1 Then
    '            '                .Range(10 + i, 13 + z).Interior.Color = Color.Yellow
    '            '            ElseIf .Cells(10 + i, 2).Value = 2 Then
    '            '                .Range(10 + i, 13 + z).Interior.Color = Color.Orange
    '            '            ElseIf .Cells(10 + i, 2).Value = 2 Then
    '            '                .Range(10 + i, 13 + z).Interior.Color = Color.Green
    '            '            End If
    '            '        End If
    '            '    Next
    '            'Next

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
                    'Dim dtProd As DataTable = GetSummaryOutStanding()
                    FileName = "TemplateForecastReportDaily.xlsx"
                    FilePath = Server.MapPath("~\Template\" & FileName)
                    If grid.VisibleRowCount > 0 Then
                        'If dtProd.Rows.Count > 0 Then
                        Call epplusExportExcelNew(FilePath, "Sheet1", "A:10", psERR)
                    End If
                    
                    'Dim psERR As String = ""
                    'Dim dtProd As DataTable = GetSummaryOutStanding()
                    'FileName = "TemplateSummaryOutstandingDeliverySupplier.xlsx"
                    'FilePath = Server.MapPath("~\Template\" & FileName)
                    'If dtProd.Rows.Count > 0 Then
                    '    Call epplusExportExcel(FilePath, "Sheet1", dtProd, "A:8", psERR)
                    'End If
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
            If .FieldName = "F1" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 1) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F2" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 2) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F3" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 3) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F4" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 4) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F5" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 5) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F6" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 6) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F7" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 7) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F8" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 8) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F9" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 9) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F10" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 10) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F11" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 11) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F12" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 12) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F13" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 13) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F14" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 14) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F15" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 15) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F16" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 16) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F17" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 17) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F18" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 18) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F19" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 19) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F20" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 20) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F21" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 21) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F22" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 22) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F23" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 23) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F24" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 24) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F25" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 25) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F26" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 26) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F27" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 27) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F28" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 28) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F29" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 29) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F30" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 30) = True Then
                    If e.GetValue("Rev") = "1" Then
                        e.Cell.BackColor = Color.Yellow
                    ElseIf e.GetValue("Rev") = "2" Then
                        e.Cell.BackColor = Color.Orange
                    ElseIf e.GetValue("Rev") = "3" Then
                        e.Cell.BackColor = Color.Green
                    End If
                End If
            End If
            If .FieldName = "F31" Then
                If uf_ColorCls(e.GetValue("Period"), e.GetValue("AffiliateID"), e.GetValue("Rev"), e.GetValue("PartNo"), 31) = True Then
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